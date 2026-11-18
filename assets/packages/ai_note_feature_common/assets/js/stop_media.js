// stop_media.js
// 停止所有音频和视频播放，用于页面销毁前清理

// 停止所有音频和视频元素
var mediaElements = document.querySelectorAll('audio, video');
mediaElements.forEach(function(element) {
  if (!element.paused) {
    element.pause();
    element.currentTime = 0;
  }
});

// 如果有使用 Web Audio API，也尝试停止
if (window.AudioContext || window.webkitAudioContext) {
  // 尝试获取所有 AudioContext 实例
  if (window.audioContexts) {
    window.audioContexts.forEach(function(context) {
      if (context.state !== 'closed') {
        context.close();
      }
    });
  }
}

// 尝试停止任何可能的 iframe 中的媒体
var iframes = document.querySelectorAll('iframe');
iframes.forEach(function(iframe) {
  try {
    var iframeDoc = iframe.contentDocument || iframe.contentWindow.document;
    var iframeMediaElements = iframeDoc.querySelectorAll('audio, video');
    iframeMediaElements.forEach(function(element) {
      if (!element.paused) {
        element.pause();
        element.currentTime = 0;
      }
    });
  } catch (e) {
    // 跨域 iframe 可能无法访问，忽略错误
    console.log('Could not access iframe content:', e);
  }
});
