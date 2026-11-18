(function () {
  if (window.__consoleInterceptorInstalled) return;
  window.__consoleInterceptorInstalled = true;

  var levels = ['log', 'info', 'warn', 'error', 'debug'];
  levels.forEach(function (level) {
    var original = console[level];
    console[level] = function () {
      original.apply(console, arguments);
      try {
        var msg = Array.prototype.map
          .call(arguments, function (a) {
            if (typeof a === 'object') {
              try {
                return JSON.stringify(a);
              } catch (e) {
                return String(a);
              }
            }
            return String(a);
          })
          .join(' ');
        window.NativeBridge.postMessage(
          JSON.stringify({
            type: 'log',
            data: { level: level, message: msg, source: 'console' },
          })
        );
      } catch (e) {
        // NativeBridge not ready yet — silently drop
      }
    };
  });
})();
