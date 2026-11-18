(function () {
  if (window.__AINOTE_IOS_VIEWPORT_SCROLL_LOCK_INSTALLED__) {
    return;
  }

  var userAgent = window.navigator && window.navigator.userAgent
    ? window.navigator.userAgent
    : '';
  var platform = window.navigator && window.navigator.platform
    ? window.navigator.platform
    : '';
  var isIos = /iP(ad|hone|od)/i.test(userAgent) ||
    (platform === 'MacIntel' && window.navigator && window.navigator.maxTouchPoints > 1);

  if (!isIos) {
    return;
  }

  window.__AINOTE_IOS_VIEWPORT_SCROLL_LOCK_INSTALLED__ = true;

  try {
    var style = document.createElement('style');
    style.setAttribute('data-ainote-ios-viewport-scroll-lock', 'true');
    style.textContent = [
      'html, body {',
      '  overscroll-behavior: none;',
      '}',
      'body {',
      '  overflow-x: hidden;',
      '}'
    ].join('\n');
    (document.head || document.documentElement).appendChild(style);
  } catch (_) {}

  var startY = 0;
  var activeScrollable = null;

  function isScrollableElement(node) {
    if (!node || node.nodeType !== 1) {
      return false;
    }
    var style = window.getComputedStyle(node);
    if (!style) {
      return false;
    }
    var overflowY = style.overflowY;
    if (overflowY !== 'auto' && overflowY !== 'scroll' && overflowY !== 'overlay') {
      return false;
    }
    return node.scrollHeight > node.clientHeight + 1;
  }

  function findScrollable(node) {
    var current = node;
    while (current && current !== document.body && current !== document.documentElement) {
      if (isScrollableElement(current)) {
        return current;
      }
      current = current.parentNode;
    }
    return null;
  }

  document.addEventListener('touchstart', function (event) {
    if (!event.touches || event.touches.length !== 1) {
      activeScrollable = null;
      return;
    }
    startY = event.touches[0].clientY;
    activeScrollable = findScrollable(event.target);
  }, { passive: true });

  document.addEventListener('touchmove', function (event) {
    if (!event.touches || event.touches.length !== 1) {
      return;
    }

    var currentY = event.touches[0].clientY;
    var deltaY = currentY - startY;
    var scrollable = activeScrollable || findScrollable(event.target);

    if (!scrollable) {
      event.preventDefault();
      return;
    }

    var scrollTop = scrollable.scrollTop;
    var maxScrollTop = Math.max(0, scrollable.scrollHeight - scrollable.clientHeight);
    var pullingDownAtTop = deltaY > 0 && scrollTop <= 0;
    var pullingUpAtBottom = deltaY < 0 && scrollTop >= maxScrollTop;

    if (pullingDownAtTop || pullingUpAtBottom) {
      event.preventDefault();
    }
  }, { passive: false });
})();
