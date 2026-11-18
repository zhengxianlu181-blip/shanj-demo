// fetch_interceptor.js
// 拦截所有 fetch 请求，自动注入 Note-Access header
// {{NOTE_ACCESS_TOKEN}} 由 native 侧注入前替换为实际值

(function() {
  var originalFetch = window.fetch;
  var noteAccessToken = '{{NOTE_ACCESS_TOKEN}}';

  window.fetch = function() {
    var args = Array.prototype.slice.call(arguments);
    console.log('[Fetch Interceptor] Intercepting fetch request:', args[0]);

    var options = args[1] || {};
    if (!options.headers) {
      options.headers = {};
    }

    // 添加 Note-Access 头
    if (noteAccessToken) {
      options.headers['Note-Access'] = noteAccessToken;
      console.log('[Fetch Interceptor] Added Note-Access header:', options.headers['Note-Access']);
    }

    // 记录最终的请求头
    console.log('[Fetch Interceptor] Final headers:', JSON.stringify(options.headers));

    // 调用原始 fetch
    return originalFetch.apply(this, [args[0], options]);
  };

  console.log('[Fetch Interceptor] Fetch interceptor installed successfully');
})();
