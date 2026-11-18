(function () {
  if (window.__AINOTE_LOCAL_H5_RUNTIME_INSTALLED__) {
    return;
  }
  window.__AINOTE_LOCAL_H5_RUNTIME_INSTALLED__ = true;

  window.__LOCAL_NATIVE_BRIDGE_PENDING__ =
    window.__LOCAL_NATIVE_BRIDGE_PENDING__ || {};

  window._handleNativeBridgeCallback =
    window._handleNativeBridgeCallback ||
    function (callbackId, success, response, error) {
      var pending = window.__LOCAL_NATIVE_BRIDGE_PENDING__ || {};
      var task = pending[callbackId];
      if (!task) {
        return;
      }
      delete pending[callbackId];
      if (success) {
        task.resolve(response);
        return;
      }
      task.reject(new Error(error || 'NativeBridge callback failed'));
    };

  (function () {
    var nativeBridge = window.NativeBridge || {};
    if (typeof nativeBridge.sendMessage !== 'function') {
      nativeBridge.sendMessage = function (message) {
        return new Promise(function (resolve, reject) {
          var callbackId =
            'native_' +
            Date.now() +
            '_' +
            Math.random().toString(36).slice(2);
          window.__LOCAL_NATIVE_BRIDGE_PENDING__[callbackId] = {
            resolve: resolve,
            reject: reject
          };
          var payload = JSON.stringify({
            type: 'message',
            data: message,
            callbackId: callbackId
          });
          if (typeof nativeBridge.postMessage === 'function') {
            nativeBridge.postMessage(payload);
            return;
          }
          if (
            window.webkit &&
            window.webkit.messageHandlers &&
            window.webkit.messageHandlers.NativeBridge &&
            typeof window.webkit.messageHandlers.NativeBridge.postMessage ===
              'function'
          ) {
            window.webkit.messageHandlers.NativeBridge.postMessage(payload);
            return;
          }
          delete window.__LOCAL_NATIVE_BRIDGE_PENDING__[callbackId];
          reject(new Error('NativeBridge not available'));
        });
      };
    }
    window.NativeBridge = nativeBridge;
  })();

  (function () {
    var clipboardBridge = {
      writeText: function (text) {
        var payload = JSON.stringify({
          type: 'message',
          data: {
            type: 'copy_text',
            data: { text: text == null ? '' : String(text) }
          }
        });
        if (
          window.NativeBridge &&
          typeof window.NativeBridge.postMessage === 'function'
        ) {
          window.NativeBridge.postMessage(payload);
          return Promise.resolve();
        }
        if (
          window.webkit &&
          window.webkit.messageHandlers &&
          window.webkit.messageHandlers.NativeBridge &&
          typeof window.webkit.messageHandlers.NativeBridge.postMessage ===
            'function'
        ) {
          window.webkit.messageHandlers.NativeBridge.postMessage(payload);
          return Promise.resolve();
        }
        return Promise.reject(new Error('NativeBridge not available'));
      }
    };
    try {
      if (!navigator.clipboard) {
        Object.defineProperty(navigator, 'clipboard', {
          value: clipboardBridge,
          configurable: true
        });
      } else {
        navigator.clipboard.writeText = clipboardBridge.writeText;
      }
    } catch (_) {
      try {
        navigator.clipboard = clipboardBridge;
      } catch (_) {}
    }
  })();

  window.handleNativeMessage =
    window.handleNativeMessage || (async function () { return null; });

  window.wx = window.wx || {
    config: function () {},
    ready: function (callback) {
      if (typeof callback === 'function') {
        callback();
      }
    },
    error: function () {},
    updateAppMessageShareData: function () {},
    updateTimelineShareData: function () {},
    onMenuShareAppMessage: function () {},
    onMenuShareTimeline: function () {},
    hideOptionMenu: function () {},
    showOptionMenu: function () {}
  };
})();
