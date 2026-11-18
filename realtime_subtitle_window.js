(function () {
  let subtitleWindow = null;
  let originalText = '';
  let translationText = '';
  let duration = '00:00:00';
  let originalSpeakerRanges = [];
  let translationSpeakerRanges = [];
  let showTranslation = false;
  let lastTranscriptHeight = 0;
  let lastRenderedRanges = '';
  let captionScale = 1;
  let compactViewportHeight = 136;
  let settingsBaseHeight = 136;
  let settingsOriginalTop = null;
  let settingsOriginalLeft = null;
  let settingsBottom = null;
  const fontSizes = [16, 18, 20, 22, 24, 30, 36, 48, 64];
  const textColors = ['#956fe7', '#4da9ef', '#37d9f1', '#a3e044', '#ffda01',
    '#ff7802', '#ff0000', '#000000', '#ffffff'];
  const backgroundColors = ['#521b78', '#19439c', '#1d7992', '#1c7231', '#ad730e',
    '#b95514', '#be191d', '#333333', '#000000'];
  const settings = {lines: 2, fontSize: 18, textColor: '#ffffff',
    backgroundColor: '#000000', backgroundOpacity: 100};
  const settingKinds = ['lines', 'fontSize', 'textColor', 'backgroundColor'];
  const settingPanelSizes = {
    lines: {width: 74, height: 106},
    fontSize: {width: 60, height: 310},
    textColor: {width: 70, height: 310},
    backgroundColor: {width: 108, height: 310},
  };
  let openSettingsKind = null;

  function backgroundCss() {
    const hex = settings.backgroundColor.slice(1);
    const channels = [0, 2, 4].map((offset) => parseInt(hex.slice(offset, offset + 2), 16));
    return `rgba(${channels.join(', ')}, ${settings.backgroundOpacity / 100})`;
  }

  function alignTranscriptToLine(transcript, content, desiredOffset) {
    if (!subtitleWindow || !transcript || !content) return;
    const doc = subtitleWindow.document;
    const lineHeight = Math.max(1, settings.fontSize * captionScale * 1.5);
    const fullscreen = Boolean(doc.fullscreenElement) ||
      doc.body.classList.contains('fullscreen-fallback');
    const lineOrigin = fullscreen ? Math.max(0, subtitleWindow.innerHeight * 0.07) : 0;
    content.style.paddingBottom = `${lineHeight}px`;
    const maximumOffset = Math.max(0, transcript.scrollHeight - transcript.clientHeight);
    const textOffset = Math.max(0, maximumOffset - lineHeight);
    const offset = desiredOffset === undefined ? transcript.scrollTop : desiredOffset;
    const clamped = Math.min(textOffset, Math.max(0, offset));
    const aligned = clamped === 0 ? 0 : lineOrigin +
      Math.ceil(Math.max(0, clamped - lineOrigin) / lineHeight) * lineHeight;
    transcript.scrollTop = Math.min(maximumOffset, aligned);
  }

  function tryMoveTo(x, y) {
    try {
      if (typeof subtitleWindow?.moveTo !== 'function') return false;
      subtitleWindow.moveTo(x, y);
      return true;
    } catch (_) { return false; }
  }

  function tryResizeTo(width, height) {
    try {
      if (typeof subtitleWindow?.resizeTo !== 'function') return false;
      subtitleWindow.resizeTo(width, height);
      return true;
    } catch (_) { return false; }
  }

  function resizeAndMove(width, height, left, top) {
    if (!subtitleWindow || subtitleWindow.closed) return false;
    const previousWidth = subtitleWindow.outerWidth || subtitleWindow.innerWidth;
    const previousHeight = subtitleWindow.outerHeight || subtitleWindow.innerHeight;
    const previousLeft = Number(subtitleWindow.screenX);
    const previousTop = Number(subtitleWindow.screenY);
    if (!tryResizeTo(width, height)) return false;
    if (Number.isFinite(left) && Number.isFinite(top) && !tryMoveTo(left, top)) {
      tryResizeTo(previousWidth, previousHeight);
      if (Number.isFinite(previousLeft) && Number.isFinite(previousTop)) {
        tryMoveTo(previousLeft, previousTop);
      }
      return false;
    }
    return true;
  }

  function compactWindowHeight() {
    const availableScreen = subtitleWindow?.screen || screen;
    const size = compactSize();
    return Math.min(availableScreen.availHeight, Math.max(size.height,
      Math.ceil(settings.lines * settings.fontSize * 1.5 + 16 + 44)));
  }

  function applySettings(resizeWindow = true) {
    if (!subtitleWindow || subtitleWindow.closed) return;
    const doc = subtitleWindow.document;
    const style = doc.documentElement.style;
    style.setProperty('--subtitle-font-size', `${settings.fontSize}px`);
    style.setProperty('--subtitle-text-color', settings.textColor);
    style.setProperty('--subtitle-background', backgroundCss());
    style.setProperty('--subtitle-compact-height', `${settings.lines * settings.fontSize * 1.5 + 16}px`);
    for (const [key, values] of Object.entries({lines: [1, 2, 3], fontSize: fontSizes,
      textColor: textColors, backgroundColor: backgroundColors})) {
      for (const value of values) {
        doc.getElementById(`subtitle-option-${key}-${value}`).classList.toggle(
          'selected', settings[key] === value);
      }
    }
    doc.getElementById('subtitle-setting-backgroundOpacity').value = String(settings.backgroundOpacity);
    doc.getElementById('subtitle-opacity-value').textContent = `${settings.backgroundOpacity}%`;
    doc.getElementById('subtitle-settings-lines').textContent = `${settings.lines}行`;
    doc.getElementById('subtitle-settings-fontSize').textContent = String(settings.fontSize);
    alignTranscriptToLine(doc.getElementById('subtitle-text'),
      doc.getElementById('subtitle-content'));
    if (resizeWindow && !doc.body.classList.contains('expanded') && !doc.fullscreenElement) {
      const availableScreen = subtitleWindow.screen || screen;
      const size = compactSize();
      const height = compactWindowHeight();
      compactViewportHeight = height;
      if (doc.body.classList.contains('settings-open')) {
        settingsBaseHeight = height;
        syncSettingsWindow();
      } else {
        const screenLeft = Number(subtitleWindow.screenX);
        const screenTop = Number(availableScreen.availTop ?? 0);
        if (Number.isFinite(screenLeft)) {
          tryMoveTo(screenLeft, Math.max(screenTop,
            screenTop + availableScreen.availHeight - height - 60));
        }
        tryResizeTo(size.width, height);
      }
    }
  }

  function beginSettingsPlacement() {
    if (settingsOriginalTop !== null) return;
    const screenX = Number(subtitleWindow.screenX);
    const screenY = Number(subtitleWindow.screenY);
    const outerHeight = subtitleWindow.outerHeight || subtitleWindow.innerHeight
      || compactViewportHeight;
    settingsBaseHeight = outerHeight;
    settingsOriginalLeft = Number.isFinite(screenX) ? screenX : null;
    settingsOriginalTop = Number.isFinite(screenY) ? screenY : null;
    settingsBottom = settingsOriginalTop === null ? null : settingsOriginalTop + outerHeight;
  }

  function syncSettingsWindow(kind = openSettingsKind) {
    if (!subtitleWindow || subtitleWindow.closed) return false;
    const reserved = kind ? settingPanelSizes[kind].height + 8 : 0;
    const width = subtitleWindow.outerWidth || subtitleWindow.innerWidth;
    const height = settingsBaseHeight + reserved;
    const top = settingsBottom !== null ? settingsBottom - height : settingsOriginalTop;
    const adjusted = resizeAndMove(width, height, settingsOriginalLeft,
      top === null ? null : Math.max(0, top));
    if (!adjusted) return false;
    subtitleWindow.document.documentElement.style.setProperty(
      '--settings-reserved', `${reserved}px`);
    return true;
  }

  function endSettingsPlacement() {
    settingsOriginalTop = null;
    settingsOriginalLeft = null;
    settingsBottom = null;
  }

  function compactSize() {
    const availableScreen = subtitleWindow?.screen || screen;
    return {
      width: Math.min(availableScreen.availWidth, Math.max(480, Math.round(availableScreen.availWidth * 0.7))),
      height: Math.min(availableScreen.availHeight, Math.max(136, Math.round(availableScreen.availHeight * 0.12))),
    };
  }

  function render() {
    if (!subtitleWindow || subtitleWindow.closed) return;
    const doc = subtitleWindow.document;
    const transcript = doc.getElementById('subtitle-text');
    const content = doc.getElementById('subtitle-content');
    if (!transcript || !content) return;
    const text = showTranslation && translationText ? translationText : originalText;
    const ranges = showTranslation && translationText ? translationSpeakerRanges : originalSpeakerRanges;
    const rangesKey = JSON.stringify(ranges);
    if (content.textContent !== text || lastRenderedRanges !== rangesKey) {
      const previousOffset = transcript.scrollTop;
      content.replaceChildren();
      let offset = 0;
      for (const [start, length] of ranges) {
        if (start < offset || length <= 0 || start + length > text.length) continue;
        content.appendChild(doc.createTextNode(text.slice(offset, start)));
        const label = doc.createElement('span');
        label.className = 'subtitle-speaker';
        label.textContent = text.slice(start, start + length);
        content.appendChild(label);
        offset = start + length;
      }
      content.appendChild(doc.createTextNode(text.slice(offset)));
      const height = transcript.scrollHeight;
      const maximumOffset = Math.max(0, height - transcript.clientHeight);
      let nextOffset = 0;
      if (maximumOffset <= 1) {
        nextOffset = 0;
      } else if (height > lastTranscriptHeight + 1) {
        nextOffset = maximumOffset;
      } else {
        nextOffset = Math.min(previousOffset, maximumOffset);
      }
      alignTranscriptToLine(transcript, content, nextOffset);
      lastTranscriptHeight = height;
      lastRenderedRanges = rangesKey;
    }
    doc.getElementById('subtitle-duration').textContent = duration;
    doc.getElementById('subtitle-translation').disabled = !translationText;
  }

  function updateCaptionScale() {
    const outer = subtitleWindow?.outerWidth;
    const inner = subtitleWindow?.innerWidth;
    const topLevel = !subtitleWindow?.top || subtitleWindow.top === subtitleWindow;
    const raw = topLevel && outer > 0 && inner > 0 ? inner / outer : 1;
    if (!Number.isFinite(raw)) return false;
    const next = raw > 0.9 && raw < 1.1 ? 1 : Math.max(0.25, Math.min(4, raw));
    if (captionScale === next) return false;
    captionScale = next;
    subtitleWindow.document.documentElement.style.setProperty('--subtitle-scale', String(next));
    return true;
  }

  function parseRanges(encoded) {
    try {
      const ranges = JSON.parse(encoded);
      return Array.isArray(ranges) ? ranges : [];
    } catch (_) { return []; }
  }

  window.realtimeSubtitleOpen = async function (original, translation, time, _speakerName = '', _translatedSpeakerName = '', originalRanges = '[]', translatedRanges = '[]') {
    if (!window.documentPictureInPicture?.requestWindow) return false;
    originalText = original;
    translationText = translation;
    duration = time;
    originalSpeakerRanges = parseRanges(originalRanges);
    translationSpeakerRanges = parseRanges(translatedRanges);
    if (subtitleWindow && !subtitleWindow.closed) {
      subtitleWindow.focus();
      render();
      return true;
    }
    const size = compactSize();
    const height = compactWindowHeight();
    const screenLeft = screen.availLeft ?? 0;
    const screenTop = screen.availTop ?? 0;
    subtitleWindow = await window.documentPictureInPicture.requestWindow({
      width: size.width,
      height,
      left: Math.round(screenLeft + (screen.availWidth - size.width) / 2),
      top: Math.round(screenTop + screen.availHeight - height - 60),
      preferInitialWindowPlacement: false,
    });
    compactViewportHeight = subtitleWindow.innerHeight;
    lastTranscriptHeight = 0;
    lastRenderedRanges = '';
    captionScale = 1;
    const doc = subtitleWindow.document;
    doc.title = '实时字幕';
    const style = doc.createElement('style');
    style.textContent = `
      * { box-sizing: border-box; }
      html, body { margin: 0; height: 100%; color: #fff;
        font-family: system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif; }
      html { --subtitle-scale: 1; --subtitle-font-size: 18px; --subtitle-text-color: #fff;
        --subtitle-background: #000; --subtitle-compact-height: 70px; }
      body { display: flex; flex-direction: column; overflow: hidden; border-radius: 10px;
        background: var(--subtitle-background); }
      body.settings-open { padding-top: var(--settings-reserved, 0px); }
      #subtitle-toolbar { height: 44px; flex: none; display: grid; align-items: center;
        grid-template-columns: minmax(0, 1fr) auto minmax(0, 1fr); position: relative; z-index: 2;
        padding: 0 16px; gap: 8px; background: #171717;
        border-bottom: 1px solid #444; }
      body.unfocused #subtitle-toolbar, :fullscreen #subtitle-toolbar,
      body.fullscreen-fallback #subtitle-toolbar { display: none; }
      #subtitle-duration { grid-column: 1; display: flex; align-items: center; gap: 6px; color: #fff;
        position: relative; font-size: 12px; font-weight: 500; font-variant-numeric: tabular-nums; }
      #subtitle-duration::before { content: ''; width: 12px; height: 12px;
        border: 1.5px solid #fff; border-radius: 50%; opacity: .9; }
      #subtitle-duration::after { content: ''; position: absolute; left: 6px; top: 11px;
        width: 1px; height: 4px; background: #fff; transform-origin: 50% 0; transform: rotate(-18deg); }
      #subtitle-toolbar-center { grid-column: 2; display: flex; align-items: center; justify-content: center;
        gap: 12px; min-width: 0; }
      #subtitle-modes { display: flex; gap: 0;
        border: 1px solid #565656; border-radius: 4px; overflow: hidden; }
      #subtitle-close { grid-column: 3; justify-self: end; margin-left: 0; }
      #subtitle-display-settings { display: flex; gap: 4px; }
      #subtitle-display-settings button { height: 29px; font-size: 12px;
        border: 1px solid #565656; border-radius: 4px; }
      #subtitle-settings-lines { width: 44px; }
      #subtitle-settings-fontSize { width: 38px; }
      #subtitle-settings-textColor { width: 48px; }
      #subtitle-settings-backgroundColor { width: 38px; }
      .subtitle-settings-panel { display: none; position: absolute; z-index: 3;
        bottom: calc(100% + 4px); left: 0; padding: 0; color: #191919;
        background: rgba(0, 0, 0, .84); border: 1px solid #55585e; border-radius: 2px; }
      .subtitle-settings-panel.open { display: block; }
      .subtitle-settings-panel.list { background: rgba(0, 0, 0, .84); }
      .subtitle-settings-panel.list .subtitle-option { padding: 0 10px; color: #fff;
        background: rgba(0, 0, 0, .35); border: 0;
        border-bottom: 1px solid #6c6e72; }
      .subtitle-settings-panel.list .subtitle-option:last-child { border-bottom: 0; }
      .subtitle-settings-panel.list .subtitle-option:hover,
      .subtitle-settings-panel.list .subtitle-option.selected { color: #fff;
        background: rgba(35, 128, 245, .62); border: 1px solid #2380f5; }
      .subtitle-settings-list { height: 100%; overflow: hidden; }
      .subtitle-option { width: 100%; height: 34px; text-align: left; padding: 0 12px;
        color: #fff; background: rgba(0, 0, 0, .35); border: 2px solid #55585e;
        font-size: 18px; font-weight: 400; }
      .subtitle-option:hover, .subtitle-option.selected { border: 2px solid #2380f5;
        background: rgba(35, 128, 245, .62); color: #fff; }
      .subtitle-swatch { display: block; width: 66px; height: 34px;
        background: var(--swatch-color); border: 2px solid #55585e; }
      .subtitle-swatch:hover, .subtitle-swatch.selected { border-color: #2380f5; }
      #subtitle-background-options { position: relative; height: 100%; }
      #subtitle-opacity-control { position: absolute; left: 70px; top: 4px;
        display: flex; align-items: center; flex-direction: column; gap: 8px;
        color: #fff; font-size: 10px; }
      #subtitle-setting-backgroundOpacity { writing-mode: vertical-lr; direction: rtl;
        appearance: none; -webkit-appearance: none; width: 32px; height: 260px;
        background: transparent; cursor: pointer; }
      #subtitle-setting-backgroundOpacity::-webkit-slider-runnable-track {
        width: 8px; height: 260px; background: #989aa0; border: 0; border-radius: 0; }
      #subtitle-setting-backgroundOpacity::-webkit-slider-thumb {
        appearance: none; -webkit-appearance: none; width: 10px; height: 10px;
        margin-left: -1px; border: 0; border-radius: 50%; background: #6d717b; }
      #subtitle-setting-backgroundOpacity::-moz-range-track {
        width: 8px; height: 260px; background: #989aa0; border: 0; }
      #subtitle-setting-backgroundOpacity::-moz-range-thumb {
        width: 10px; height: 10px; border: 0; border-radius: 50%; background: #6d717b; }
      #subtitle-opacity-value { font-variant-numeric: tabular-nums; }
      #subtitle-opacity-title { font-weight: 400; }
      button { appearance: none; -webkit-appearance: none; box-shadow: none;
        color: #fff; background: #242424; border: 0;
        padding: 0; cursor: pointer; font-family: system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
        font-size: 13px; font-weight: 500; line-height: 1.2; }
      button:focus-visible { outline: 2px solid #fff; outline-offset: -2px; }
      #subtitle-modes button { width: 46px; height: 29px; border-right: 1px solid #565656; }
      #subtitle-modes button:last-child { border-right: 0; }
      button:hover:not(:disabled) { background: #393939; }
      .subtitle-settings-panel button:focus-visible { outline: 2px solid #2380f5;
        outline-offset: -2px; }
      .subtitle-settings-panel .subtitle-swatch:hover,
      .subtitle-settings-panel .subtitle-swatch.selected { background: var(--swatch-color); }
      button:disabled { opacity: .38; cursor: default; }
      button.active { color: #fff; background: #147ff2; }
      #subtitle-close { min-width: 104px; height: 30px; padding: 0 8px;
        display: inline-flex; align-items: center; justify-content: center; gap: 8px;
        border-radius: 4px; background: transparent; }
      #subtitle-close:hover { background: #393939; }
      .eye-off { position: relative; display: inline-block; width: 16px; height: 10px;
        border: 1.5px solid #fff; border-radius: 80% 20%; transform: rotate(-45deg); }
      .eye-off::after { content: ''; position: absolute; left: -3px; top: 3px;
        width: 21px; height: 1.5px; background: #fff; transform: rotate(45deg); }
      #subtitle-text { flex: 1; min-height: 0;
        max-height: calc(var(--subtitle-compact-height) * var(--subtitle-scale));
        margin: auto 0; display: flex; flex-direction: column;
        overflow: auto; overflow-anchor: none;
        clip-path: inset(calc(8px * var(--subtitle-scale)) 0 0 0);
        background: transparent; border-radius: 10px;
        padding: calc(8px * var(--subtitle-scale)) calc(28px * var(--subtitle-scale));
        color: var(--subtitle-text-color);
        font-size: calc(var(--subtitle-font-size) * var(--subtitle-scale)); font-weight: 500;
        line-height: 1.5;
        white-space: pre-wrap; overflow-wrap: anywhere; letter-spacing: 0; }
      #subtitle-content { flex: none; min-width: 0; }
      .subtitle-speaker { color: #a8c6ff; font-weight: 600; }
      body.expanded #subtitle-text { max-height: none; }
      :fullscreen #subtitle-text, body.fullscreen-fallback #subtitle-text {
        flex: 1; height: auto; max-height: none; min-height: 44.8px;
        margin: 0; overflow: auto; clip-path: none;
        padding: 7vh 9vw; }
      @media (max-width: 640px) {
        #subtitle-duration { display: none; }
        #subtitle-modes { margin-right: 0; }
      }
      @media (max-width: 600px) {
        #subtitle-toolbar { padding: 0 8px; gap: 4px; }
        #subtitle-toolbar-center { gap: 6px; }
        #subtitle-display-settings { gap: 2px; }
        #subtitle-modes button { width: 40px; }
        #subtitle-settings-lines { width: 38px; }
        #subtitle-settings-fontSize { width: 32px; }
        #subtitle-settings-textColor { width: 42px; }
        #subtitle-settings-backgroundColor { width: 32px; }
        #subtitle-close { min-width: 88px; padding: 0 4px; }
      }
    `;
    doc.head.appendChild(style);
    doc.body.innerHTML = `
      <div id="subtitle-toolbar">
        <span id="subtitle-duration">00:00:00</span>
        <span id="subtitle-toolbar-center">
          <span id="subtitle-modes">
            <button id="subtitle-compact" class="active" title="窗口模式">窗口</button>
            <button id="subtitle-fullscreen" title="全屏">全屏</button>
            <button id="subtitle-translation" title="切换翻译">翻译</button>
          </span>
          <span id="subtitle-display-settings">
            <button id="subtitle-settings-lines" title="字幕行数" aria-expanded="false">2行</button>
            <button id="subtitle-settings-fontSize" title="字幕字号" aria-expanded="false">18</button>
            <button id="subtitle-settings-textColor" title="文字颜色" aria-expanded="false">文字色</button>
            <button id="subtitle-settings-backgroundColor" title="背景颜色" aria-expanded="false">背景</button>
          </span>
        </span>
        <div id="subtitle-settings-panel-lines" class="subtitle-settings-panel list">
          <div class="subtitle-settings-list">${[1, 2, 3].map((count) => `<button class="subtitle-option" id="subtitle-option-lines-${count}">${count} 行</button>`).join('')}</div>
        </div>
        <div id="subtitle-settings-panel-fontSize" class="subtitle-settings-panel list">
          <div class="subtitle-settings-list">${fontSizes.map((size) => `<button class="subtitle-option" id="subtitle-option-fontSize-${size}">${size}</button>`).join('')}</div>
        </div>
        <div id="subtitle-settings-panel-textColor" class="subtitle-settings-panel">
          <div class="subtitle-settings-list">${textColors.map((color) => `<button class="subtitle-swatch" id="subtitle-option-textColor-${color}" title="${color.toUpperCase()}" aria-label="文字颜色 ${color.toUpperCase()}" style="--swatch-color:${color}"></button>`).join('')}</div>
        </div>
        <div id="subtitle-settings-panel-backgroundColor" class="subtitle-settings-panel">
          <div id="subtitle-background-options" class="subtitle-settings-list">${backgroundColors.map((color) => `<button class="subtitle-swatch" id="subtitle-option-backgroundColor-${color}" title="${color.toUpperCase()}" aria-label="背景颜色 ${color.toUpperCase()}" style="--swatch-color:${color}"></button>`).join('')}<div id="subtitle-opacity-control"><output id="subtitle-opacity-value">100%</output><input id="subtitle-setting-backgroundOpacity" type="range" min="1" max="100" step="1" value="100" aria-label="背景透明度"><span id="subtitle-opacity-title">透明度</span></div></div>
        </div>
        <button id="subtitle-close" title="关闭字幕"><span class="eye-off" aria-hidden="true"></span>隐藏字幕</button>
      </div>
      <div id="subtitle-text"><div id="subtitle-content"></div></div>`;
    updateCaptionScale();
    doc.body.classList.add('unfocused');
    const setSettingsOpen = (kind) => {
      const previous = openSettingsKind;
      const wasOpen = previous !== null;
      const next = previous === kind ? null : kind;
      if (next !== null && !wasOpen) beginSettingsPlacement();
      if (!syncSettingsWindow(next)) {
        if (!wasOpen) endSettingsPlacement();
        return false;
      }
      openSettingsKind = next;
      for (const key of settingKinds) {
        const open = key === next;
        const panel = doc.getElementById(`subtitle-settings-panel-${key}`);
        const button = doc.getElementById(`subtitle-settings-${key}`);
        panel.classList.toggle('open', open);
        button.setAttribute('aria-expanded', String(open));
        if (open) {
          const size = settingPanelSizes[key];
          panel.style.width = `${size.width}px`;
          panel.style.height = `${size.height}px`;
          panel.style.left = `${Math.max(4, Math.min(button.offsetLeft || 0,
            subtitleWindow.innerWidth - size.width - 4))}px`;
        }
      }
      doc.body.classList.toggle('settings-open', next !== null);
      if (next === null) {
        endSettingsPlacement();
      }
      return true;
    };
    doc.body.addEventListener('mouseenter', () => {
      doc.body.classList.remove('unfocused');
    });
    doc.body.addEventListener('mouseleave', () => {
      doc.body.classList.add('unfocused');
      if (openSettingsKind) setSettingsOpen(openSettingsKind);
    });
    doc.addEventListener('click', (event) => {
      if (!openSettingsKind) return;
      const target = event.target;
      if (target && typeof target.closest === 'function' &&
          (target.closest('.subtitle-settings-panel') ||
           target.closest('#subtitle-display-settings'))) return;
      setSettingsOpen(openSettingsKind);
    }, true);
    subtitleWindow.addEventListener('blur', () => {
      if (openSettingsKind) setSettingsOpen(openSettingsKind);
    });
    for (const key of settingKinds) {
      doc.getElementById(`subtitle-settings-${key}`).onclick = () => setSettingsOpen(key);
      const values = key === 'lines' ? [1, 2, 3] : key === 'fontSize' ? fontSizes
        : key === 'textColor' ? textColors : backgroundColors;
      for (const value of values) {
        doc.getElementById(`subtitle-option-${key}-${value}`).onclick = () => {
          settings[key] = value;
          applySettings();
        };
      }
    }
    doc.getElementById('subtitle-setting-backgroundOpacity').oninput = (event) => {
      settings.backgroundOpacity = Math.max(1, Math.min(100, Number(event.target.value) || 100));
      applySettings();
    };
    applySettings(false);
    doc.getElementById('subtitle-close').onclick = () => subtitleWindow.close();
    doc.getElementById('subtitle-translation').onclick = () => {
      showTranslation = !showTranslation;
      doc.getElementById('subtitle-translation').classList.toggle('active', showTranslation);
      render();
    };
    doc.getElementById('subtitle-fullscreen').onclick = async () => {
      if (openSettingsKind) setSettingsOpen(openSettingsKind);
      try { await doc.documentElement.requestFullscreen(); } catch (_) {
        tryResizeTo(screen.availWidth, screen.availHeight);
        doc.body.classList.add('fullscreen-fallback');
      }
      doc.getElementById('subtitle-fullscreen').classList.add('active');
      doc.getElementById('subtitle-compact').classList.remove('active');
    };
    doc.getElementById('subtitle-compact').onclick = () => {
      if (openSettingsKind) setSettingsOpen(openSettingsKind);
      doc.body.classList.remove('fullscreen-fallback');
      doc.body.classList.remove('expanded');
      const size = compactSize();
      const height = compactWindowHeight();
      const availableScreen = subtitleWindow.screen || screen;
      const screenLeft = Number(subtitleWindow.screenX);
      const screenTop = Number(availableScreen.availTop ?? 0);
      compactViewportHeight = height;
      const compacted = resizeAndMove(size.width, height,
        Number.isFinite(screenLeft) ? screenLeft : null,
        Math.max(screenTop, screenTop + availableScreen.availHeight - height - 60));
      if (doc.fullscreenElement) doc.exitFullscreen().catch(() => {});
      doc.getElementById('subtitle-compact').classList.add('active');
      doc.getElementById('subtitle-fullscreen').classList.remove('active');
      return compacted;
    };
    subtitleWindow.addEventListener('pagehide', () => {
      subtitleWindow = null;
      lastTranscriptHeight = 0;
      lastRenderedRanges = '';
      openSettingsKind = null;
      endSettingsPlacement();
    });
    subtitleWindow.addEventListener('resize', () => {
      if (updateCaptionScale()) {
        lastRenderedRanges = '';
        render();
      }
      alignTranscriptToLine(doc.getElementById('subtitle-text'),
        doc.getElementById('subtitle-content'));
      if (doc.body.classList.contains('settings-open')) return;
      if (subtitleWindow.innerHeight > compactViewportHeight + 24) {
        doc.body.classList.add('expanded');
      } else {
        doc.body.classList.remove('expanded');
      }
    });
    doc.addEventListener('fullscreenchange', () => {
      const fullscreen = Boolean(doc.fullscreenElement);
      doc.getElementById('subtitle-fullscreen').classList.toggle('active', fullscreen);
      doc.getElementById('subtitle-compact').classList.toggle('active', !fullscreen);
    });
    doc.addEventListener('keydown', (event) => {
      if (event.key !== 'Escape') return;
      if (event.key === 'Escape' && openSettingsKind) {
        setSettingsOpen(openSettingsKind);
        return;
      }
      if (doc.fullscreenElement) {
        doc.exitFullscreen().catch(() => {});
      } else if (doc.body.classList.contains('fullscreen-fallback') &&
          !doc.getElementById('subtitle-compact').onclick()) {
        subtitleWindow.close();
      }
    });
    render();
    return true;
  };

  window.realtimeSubtitleUpdate = function (original, translation, time, _speakerName = '', _translatedSpeakerName = '', originalRanges = '[]', translatedRanges = '[]') {
    if (!subtitleWindow || subtitleWindow.closed) return false;
    originalText = original;
    translationText = translation;
    duration = time;
    originalSpeakerRanges = parseRanges(originalRanges);
    translationSpeakerRanges = parseRanges(translatedRanges);
    render();
    return true;
  };

  window.realtimeSubtitleClose = function () {
    if (subtitleWindow && !subtitleWindow.closed) subtitleWindow.close();
    subtitleWindow = null;
    lastTranscriptHeight = 0;
    lastRenderedRanges = '';
  };
})();
