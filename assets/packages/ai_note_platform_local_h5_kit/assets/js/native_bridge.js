// native_bridge.js
// NativeBridge 初始化：向 WebView 注入的 NativeBridge 对象添加 sendMessage 方法
// 由 native 侧在 onPageStarted 时注入

console.log('NativeBridge: Initializing...');

// This script can be injected multiple times during one page lifecycle.
// Preserve bridge state so pending callbacks are not lost on reinjection.
window._nativeBridgeCallbacks = window._nativeBridgeCallbacks || {};
window._nativeBridgeReady = window._nativeBridgeReady || false;
window._nativeBridgeMessageQueue = window._nativeBridgeMessageQueue || [];
window._nativeBridgeCallbackCounter = window._nativeBridgeCallbackCounter || 0;

function _generateCallbackId() {
  window._nativeBridgeCallbackCounter++;
  return Date.now().toString() + '_' + window._nativeBridgeCallbackCounter;
}

function _sendToNative(message) {
  var messageStr = typeof message === 'string' ? message : JSON.stringify(message);
  
  if (window.webkit && window.webkit.messageHandlers && window.webkit.messageHandlers.NativeBridge) {
    try {
      window.webkit.messageHandlers.NativeBridge.postMessage(messageStr);
      return true;
    } catch (e) {
      console.error('NativeBridge: Failed to send via webkit.messageHandlers:', e);
      return false;
    }
  }
  
  if (window.NativeBridge && typeof window.NativeBridge.postMessage === 'function') {
    try {
      window.NativeBridge.postMessage(messageStr);
      return true;
    } catch (e) {
      console.error('NativeBridge: Failed to send via NativeBridge.postMessage:', e);
      return false;
    }
  }
  
  console.error('NativeBridge: No available channel to send message');
  return false;
}


function _installContentScrollFastPath() {
  var bridge = window.NativeBridge;
  if (!bridge || typeof bridge.sendMessage !== 'function' || bridge.__detailTitleScrollFastPath) {
    return;
  }
  var send = bridge.sendMessage.bind(bridge);
  bridge.__detailTitleScrollFastPath = true;
  bridge.sendMessage = function (message) {
    if (message && message.type === 'content_scroll') {
      var sent = _sendToNative({
        type: 'message',
        data: message
      });
      if (!sent) {
        return Promise.reject(new Error('Failed to send message to native'));
      }
      return Promise.resolve({ success: true });
    }
    return send(message);
  };
}

function _processQueue() {
  while (window._nativeBridgeMessageQueue.length > 0 && window._nativeBridgeReady) {
    var message = window._nativeBridgeMessageQueue.shift();
    _sendToNative(message);
  }
}

function _waitForBridgeReady(timeoutMs) {
  var timeout = typeof timeoutMs === 'number' ? timeoutMs : 10000;
  var startedAt = Date.now();

  return new Promise(function(resolve, reject) {
    function checkReady() {
      if (_ensureBridge()) {
        resolve(window.NativeBridge || {});
        return;
      }

      if (Date.now() - startedAt >= timeout) {
        reject(new Error('Native bridge is not ready'));
        return;
      }

      setTimeout(checkReady, 100);
    }

    checkReady();
  });
}

function _ensureBridge() {
  var hasWebKit = window.webkit && window.webkit.messageHandlers && window.webkit.messageHandlers.NativeBridge;
  var hasNativeBridge = window.NativeBridge && typeof window.NativeBridge.postMessage === 'function';
  
  if (hasWebKit || hasNativeBridge) {
    if (!window.NativeBridge) {
      window.NativeBridge = {};
    }
    
    if (typeof window.NativeBridge.postMessage !== 'function') {
      window.NativeBridge.postMessage = function(message) {
        _sendToNative(message);
      };
    }
    
    if (typeof window.NativeBridge.sendMessage !== 'function') {
      window.NativeBridge.sendMessage = function(message) {
        var callbackId = _generateCallbackId();
        return new Promise(function(resolve, reject) {
          window._nativeBridgeCallbacks[callbackId] = { resolve: resolve, reject: reject };

          var sent = _sendToNative({
            type: 'message',
            data: message,
            callbackId: callbackId
          });
          
          if (!sent) {
            delete window._nativeBridgeCallbacks[callbackId];
            reject(new Error('Failed to send message to native'));
          }
        });
      };
    }
    _installContentScrollFastPath();
    
    window._nativeBridgeReady = true;
    console.log('NativeBridge: Bridge ready, channel:', hasWebKit ? 'webkit.messageHandlers' : 'NativeBridge.postMessage');
    _processQueue();
    return true;
  }
  
  return false;
}

if (!_ensureBridge()) {
  var _retries = 0;
  var _maxRetries = 50;
  var _interval = setInterval(function() {
    _retries++;
    if (_ensureBridge()) {
      clearInterval(_interval);
    } else if (_retries >= _maxRetries) {
      clearInterval(_interval);
      console.error('NativeBridge: Failed to initialize after ' + _maxRetries + ' retries');
    }
  }, 200);
}

window._handleNativeBridgeCallback = function(callbackId, success, response, error) {
  console.log('NativeBridge: Handling callback:', callbackId, 'success:', success);
  
  var callback = window._nativeBridgeCallbacks[callbackId];
  var pending = window.__LOCAL_NATIVE_BRIDGE_PENDING__ || {};
  if (!callback && pending[callbackId]) {
    callback = pending[callbackId];
  }
  if (callback) {
    if (callback.timeoutId) {
      clearTimeout(callback.timeoutId);
    }
    delete window._nativeBridgeCallbacks[callbackId];
    if (pending[callbackId]) {
      delete pending[callbackId];
    }
    
    if (success) {
      callback.resolve(response);
    } else {
      callback.reject(new Error(error || 'Unknown error'));
    }
  } else {
    console.warn('NativeBridge: No callback found for ID:', callbackId);
  }
};

function _normalizeAppMessage(handlerName, message) {
  var payload = message;
  if (typeof handlerName === 'object' && message === undefined) {
    payload = handlerName;
  } else if (
    handlerName === 'nativeChannel' ||
    handlerName === 'NativeBridge' ||
    handlerName === 'sendMessage'
  ) {
    payload = message;
  } else if (handlerName === 'onMessage') {
    try {
      payload = typeof message === 'string' ? JSON.parse(message) : message;
    } catch (e) {
      payload = message;
    }
  } else if (handlerName) {
    payload = {
      type: handlerName,
      data: message
    };
  }

  if (payload && typeof payload === 'object' && payload.event && !payload.type) {
    payload = {
      type: payload.event,
      data: payload.data
    };
  }
  return payload;
}

// Backward compatibility for older H5 bundles that still call window.sendAppMessage.
// Keep a private native sender because some H5 bundles replace window.sendAppMessage
// during app bootstrap and then call flutter_inappwebview.callHandler('onMessage', ...).
window.__APP_NATIVE_SEND_APP_MESSAGE__ = function(message) {
  return _waitForBridgeReady().then(function(bridge) {
    if (bridge && typeof bridge.sendMessage === 'function') {
      return bridge.sendMessage(message);
    }

    if (bridge && typeof bridge.postMessage === 'function') {
      bridge.postMessage(JSON.stringify({
        type: 'message',
        data: message
      }));
      return true;
    }

    throw new Error('Native bridge has no compatible send method');
  });
};

window.sendAppMessage = window.__APP_NATIVE_SEND_APP_MESSAGE__;

// Backward compatibility for H5 bundles that call flutter_inappwebview.callHandler.
window.flutter_inappwebview = window.flutter_inappwebview || {};
window.flutter_inappwebview.callHandler = function(handlerName, message) {
  return window.__APP_NATIVE_SEND_APP_MESSAGE__(
    _normalizeAppMessage(handlerName, message)
  );
};

console.log('NativeBridge: Initialization complete');
console.log(
  'NativeBridge: postMessage available:',
  !!window.NativeBridge &&
    typeof window.NativeBridge.postMessage === 'function'
);
console.log(
  'NativeBridge: sendMessage available:',
  !!window.NativeBridge &&
    typeof window.NativeBridge.sendMessage === 'function'
);
console.log('NativeBridge: sendAppMessage available:', typeof window.sendAppMessage === 'function');
