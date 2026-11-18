// xhr_interceptor.js
// 拦截所有 XMLHttpRequest 请求，自动注入 Note-Access header
// {{NOTE_ACCESS_TOKEN}} 由 native 侧注入前替换为实际值

(function() {
  var originalOpen = XMLHttpRequest.prototype.open;
  var originalSetRequestHeader = XMLHttpRequest.prototype.setRequestHeader;
  var noteAccessToken = '{{NOTE_ACCESS_TOKEN}}';

  XMLHttpRequest.prototype.open = function(method, url) {
    console.log('[XHR Interceptor] Intercepting XHR request:', method, url);

    // 标记这个请求，以便在 setRequestHeader 时识别
    this._noteAccessToken = noteAccessToken;
    this._noteAccessTokenAdded = false;

    return originalOpen.apply(this, arguments);
  };

  XMLHttpRequest.prototype.setRequestHeader = function(name, value) {
    // 在第一次调用 setRequestHeader 时添加 Note-Access 头
    if (this._noteAccessToken && !this._noteAccessTokenAdded) {
      console.log('[XHR Interceptor] Adding Note-Access header to XHR request');
      originalSetRequestHeader.call(this, 'Note-Access', this._noteAccessToken);
      console.log('[XHR Interceptor] Added Note-Access header:', this._noteAccessToken);
      this._noteAccessTokenAdded = true;
    }

    return originalSetRequestHeader.call(this, name, value);
  };

  // 如果请求没有调用 setRequestHeader，我们也要添加头
  var originalSend = XMLHttpRequest.prototype.send;
  XMLHttpRequest.prototype.send = function() {
    if (this._noteAccessToken && !this._noteAccessTokenAdded) {
      console.log('[XHR Interceptor] Adding Note-Access header during send');
      originalSetRequestHeader.call(this, 'Note-Access', this._noteAccessToken);
      console.log('[XHR Interceptor] Added Note-Access header:', this._noteAccessToken);
      this._noteAccessTokenAdded = true;
    }

    return originalSend.apply(this, arguments);
  };

  console.log('[XHR Interceptor] XHR interceptor installed successfully');
})();
