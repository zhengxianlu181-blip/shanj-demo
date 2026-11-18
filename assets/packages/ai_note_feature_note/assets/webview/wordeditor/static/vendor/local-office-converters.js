(function () {
    const EMU_PER_PX = 9525;
    const PX_PER_INCH = 96;

    function requireGlobal(name, value) {
        if (!value) throw new Error(`缺少本地浏览器依赖: ${name}`);
        return value;
    }

    function safeFilename(filename, fallback) {
        return String(filename || fallback || 'document')
            .replace(/[\\/:*?"<>|]+/g, '_')
            .slice(0, 180) || fallback;
    }

    function dataUrlToUint8Array(dataUrl) {
        const base64 = String(dataUrl || '').split(',', 2)[1] || '';
        const binary = atob(base64);
        const bytes = new Uint8Array(binary.length);
        for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
        return bytes;
    }

    function uint8ArrayToDataUrl(bytes, contentType) {
        let binary = '';
        const chunkSize = 0x8000;
        for (let i = 0; i < bytes.length; i += chunkSize) {
            binary += String.fromCharCode.apply(null, bytes.subarray(i, i + chunkSize));
        }
        return `data:${contentType || 'application/octet-stream'};base64,${btoa(binary)}`;
    }

    function emuToPx(value) {
        return Math.round((Number(value || 0) / EMU_PER_PX) * 100) / 100;
    }

    function pxToIn(value) {
        return Math.max(0, Number(value || 0) / PX_PER_INCH);
    }

    function getLocalName(node) {
        return node && (node.localName || String(node.nodeName || '').split(':').pop());
    }

    function childrenByLocalName(node, localName) {
        return Array.from(node ? node.childNodes : []).filter(child => getLocalName(child) === localName);
    }

    function firstDescendant(node, localName) {
        if (!node) return null;
        if (getLocalName(node) === localName) return node;
        for (const child of Array.from(node.childNodes || [])) {
            const found = firstDescendant(child, localName);
            if (found) return found;
        }
        return null;
    }

    function descendantsByLocalName(node, localName, results) {
        results = results || [];
        if (!node) return results;
        if (getLocalName(node) === localName) results.push(node);
        Array.from(node.childNodes || []).forEach(child => descendantsByLocalName(child, localName, results));
        return results;
    }

    function parseXml(xml) {
        const doc = new DOMParser().parseFromString(xml, 'application/xml');
        const parserError = doc.getElementsByTagName('parsererror')[0];
        if (parserError) throw new Error('Office XML 解析失败');
        return doc;
    }

    function attr(node, name, fallback) {
        if (!node) return fallback || '';
        return node.getAttribute(name) || node.getAttribute(`r:${name}`) || fallback || '';
    }

    function normalizeZipPath(path) {
        const parts = [];
        String(path || '').split('/').forEach(part => {
            if (!part || part === '.') return;
            if (part === '..') parts.pop();
            else parts.push(part);
        });
        return parts.join('/');
    }

    function resolveTarget(baseDir, target) {
        if (!target) return '';
        if (target[0] === '/') return normalizeZipPath(target.slice(1));
        return normalizeZipPath(`${baseDir}/${target}`);
    }

    function contentTypeForPath(path) {
        const ext = String(path || '').split('.').pop().toLowerCase();
        if (ext === 'jpg' || ext === 'jpeg') return 'image/jpeg';
        if (ext === 'png') return 'image/png';
        if (ext === 'gif') return 'image/gif';
        if (ext === 'webp') return 'image/webp';
        if (ext === 'svg') return 'image/svg+xml';
        return 'application/octet-stream';
    }

    async function readZipText(zip, path) {
        const file = zip.file(path);
        return file ? file.async('string') : '';
    }

    function parseRelationships(xml, baseDir) {
        if (!xml) return {};
        const doc = parseXml(xml);
        const map = {};
        descendantsByLocalName(doc, 'Relationship').forEach(rel => {
            const id = rel.getAttribute('Id');
            if (!id) return;
            map[id] = {
                type: rel.getAttribute('Type') || '',
                target: resolveTarget(baseDir, rel.getAttribute('Target') || '')
            };
        });
        return map;
    }

    function getShapeBox(shape) {
        const xfrm = firstDescendant(shape, 'xfrm');
        const off = firstDescendant(xfrm, 'off');
        const ext = firstDescendant(xfrm, 'ext');
        return {
            left: emuToPx(attr(off, 'x', 0)),
            top: emuToPx(attr(off, 'y', 0)),
            width: emuToPx(attr(ext, 'cx', 120 * EMU_PER_PX)),
            height: emuToPx(attr(ext, 'cy', 40 * EMU_PER_PX))
        };
    }

    function parseTextStyle(shape) {
        const runPr = firstDescendant(shape, 'rPr') || firstDescendant(shape, 'defRPr');
        const colorNode = firstDescendant(runPr, 'srgbClr');
        const fontNode = firstDescendant(runPr, 'latin') || firstDescendant(runPr, 'ea');
        const size = Number(attr(runPr, 'sz', 2400));
        return {
            fontSize: size ? Math.round(size / 100) : 24,
            fontFamily: attr(fontNode, 'typeface', 'Arial') || 'Arial',
            color: colorNode ? `#${attr(colorNode, 'val', '111111')}` : '#111111',
            bold: attr(runPr, 'b', '') === '1',
            italic: attr(runPr, 'i', '') === '1'
        };
    }

    function parseShapeText(shape) {
        return descendantsByLocalName(shape, 't').map(node => node.textContent || '').join('\n').trim();
    }

    async function inspectPptxFile(file) {
        if (String(file.name || '').toLowerCase().endsWith('.ppt')) {
            throw new Error('本地模式暂不支持旧版 .ppt，请先另存为 .pptx 后打开。');
        }
        const JSZipLib = requireGlobal('JSZip', window.JSZip);
        const zip = await JSZipLib.loadAsync(await file.arrayBuffer());
        const presentationXml = await readZipText(zip, 'ppt/presentation.xml');
        if (!presentationXml) throw new Error('没有找到 PPTX 主文档');

        const presentationDoc = parseXml(presentationXml);
        const sizeNode = firstDescendant(presentationDoc, 'sldSz');
        const slideWidth = emuToPx(attr(sizeNode, 'cx', 12192000));
        const slideHeight = emuToPx(attr(sizeNode, 'cy', 6858000));
        const presRels = parseRelationships(
            await readZipText(zip, 'ppt/_rels/presentation.xml.rels'),
            'ppt'
        );

        const slideRefs = descendantsByLocalName(presentationDoc, 'sldId')
            .map(node => presRels[attr(node, 'id', '')] && presRels[attr(node, 'id', '')].target)
            .filter(Boolean);
        const slidePaths = slideRefs.length ? slideRefs : Object.keys(zip.files)
            .filter(name => /^ppt\/slides\/slide\d+\.xml$/.test(name))
            .sort((a, b) => Number(a.match(/\d+/)[0]) - Number(b.match(/\d+/)[0]));

        const slides = [];
        for (let slideIndex = 0; slideIndex < slidePaths.length; slideIndex++) {
            const slidePath = slidePaths[slideIndex];
            const slideXml = await readZipText(zip, slidePath);
            if (!slideXml) continue;
            const slideDoc = parseXml(slideXml);
            const relPath = slidePath.replace('/slides/', '/slides/_rels/') + '.rels';
            const slideBase = slidePath.split('/').slice(0, -1).join('/');
            const rels = parseRelationships(await readZipText(zip, relPath), slideBase);
            const elements = [];

            descendantsByLocalName(slideDoc, 'sp').forEach((shape, zIndex) => {
                const text = parseShapeText(shape);
                if (!text) return;
                elements.push(Object.assign({
                    id: `s${slideIndex + 1}-t${zIndex + 1}`,
                    type: 'text',
                    text,
                    zIndex,
                    style: parseTextStyle(shape)
                }, getShapeBox(shape)));
            });

            descendantsByLocalName(slideDoc, 'pic').forEach((picture, picIndex) => {
                const blip = firstDescendant(picture, 'blip');
                const rel = rels[attr(blip, 'embed', '')];
                if (!rel || !zip.file(rel.target)) return;
                const bytesPromise = zip.file(rel.target).async('uint8array');
                elements.push(Object.assign({
                    id: `s${slideIndex + 1}-i${picIndex + 1}`,
                    type: 'image',
                    name: rel.target.split('/').pop(),
                    contentType: contentTypeForPath(rel.target),
                    zIndex: 100 + picIndex,
                    _bytesPromise: bytesPromise
                }, getShapeBox(picture)));
            });

            for (const element of elements) {
                if (element._bytesPromise) {
                    const bytes = await element._bytesPromise;
                    element.src = uint8ArrayToDataUrl(bytes, element.contentType);
                    delete element._bytesPromise;
                }
            }

            slides.push({
                id: String(slideIndex + 1),
                title: `幻灯片 ${slideIndex + 1}`,
                width: slideWidth,
                height: slideHeight,
                background: '#ffffff',
                elements: elements.sort((a, b) => (a.zIndex || 0) - (b.zIndex || 0))
            });
        }

        return {
            filename: safeFilename(file.name, 'presentation.pptx'),
            slideWidth,
            slideHeight,
            slides
        };
    }

    const PPT_DECK_DB = 'local-office-ppt-decks';
    const PPT_DECK_STORE = 'decks';
    const memoryDeckStore = new Map();

    function openDeckDatabase() {
        return new Promise((resolve, reject) => {
            if (!window.indexedDB) {
                resolve(null);
                return;
            }
            const request = indexedDB.open(PPT_DECK_DB, 1);
            request.onupgradeneeded = () => {
                request.result.createObjectStore(PPT_DECK_STORE);
            };
            request.onsuccess = () => resolve(request.result);
            request.onerror = () => reject(request.error || new Error('IndexedDB 打开失败'));
        });
    }

    async function storePptDeck(deck) {
        const key = `deck-${Date.now()}-${Math.random().toString(36).slice(2)}`;
        memoryDeckStore.set(key, deck);
        try {
            const db = await openDeckDatabase();
            if (!db) return key;
            await new Promise((resolve, reject) => {
                const tx = db.transaction(PPT_DECK_STORE, 'readwrite');
                tx.objectStore(PPT_DECK_STORE).put(deck, key);
                tx.oncomplete = resolve;
                tx.onerror = () => reject(tx.error || new Error('PPT 缓存写入失败'));
            });
            db.close();
        } catch (error) {
            console.warn('PPT 缓存写入 IndexedDB 失败，使用内存缓存:', error);
        }
        return key;
    }

    async function takeStoredPptDeck(key) {
        if (!key) return null;
        if (memoryDeckStore.has(key)) {
            const deck = memoryDeckStore.get(key);
            memoryDeckStore.delete(key);
            return deck;
        }
        try {
            const db = await openDeckDatabase();
            if (!db) return null;
            const deck = await new Promise((resolve, reject) => {
                const tx = db.transaction(PPT_DECK_STORE, 'readwrite');
                const store = tx.objectStore(PPT_DECK_STORE);
                const request = store.get(key);
                request.onsuccess = () => {
                    const value = request.result || null;
                    store.delete(key);
                    resolve(value);
                };
                request.onerror = () => reject(request.error || new Error('PPT 缓存读取失败'));
            });
            db.close();
            return deck;
        } catch (error) {
            console.warn('PPT 缓存读取 IndexedDB 失败:', error);
            return null;
        }
    }

    async function clearStoredPptDecks() {
        memoryDeckStore.clear();
        try {
            const db = await openDeckDatabase();
            if (!db) return;
            await new Promise((resolve, reject) => {
                const tx = db.transaction(PPT_DECK_STORE, 'readwrite');
                tx.objectStore(PPT_DECK_STORE).clear();
                tx.oncomplete = resolve;
                tx.onerror = () => reject(tx.error || new Error('PPT 缓存清理失败'));
            });
            db.close();
        } catch (error) {
            console.warn('PPT 缓存清理 IndexedDB 失败:', error);
        }
    }

    function configurePdfWorker() {
        const pdfjs = requireGlobal('PDF.js', window.pdfjsLib);
        if (pdfjs.GlobalWorkerOptions) {
            pdfjs.GlobalWorkerOptions.workerSrc = 'static/vendor/pdf.worker.min.js';
        }
        return pdfjs;
    }

    function multiplyMatrix(a, b) {
        return [
            a[0] * b[0] + a[2] * b[1],
            a[1] * b[0] + a[3] * b[1],
            a[0] * b[2] + a[2] * b[3],
            a[1] * b[2] + a[3] * b[3],
            a[0] * b[4] + a[2] * b[5] + a[4],
            a[1] * b[4] + a[3] * b[5] + a[5]
        ];
    }

    function applyMatrixToPoint(matrix, x, y) {
        return {
            x: matrix[0] * x + matrix[2] * y + matrix[4],
            y: matrix[1] * x + matrix[3] * y + matrix[5]
        };
    }

    function getMatrixBoundingBox(matrix, width, height) {
        const points = [
            applyMatrixToPoint(matrix, 0, 0),
            applyMatrixToPoint(matrix, width, 0),
            applyMatrixToPoint(matrix, 0, height),
            applyMatrixToPoint(matrix, width, height)
        ];
        const xs = points.map(point => point.x);
        const ys = points.map(point => point.y);
        return {
            left: Math.min.apply(null, xs),
            top: Math.min.apply(null, ys),
            width: Math.max.apply(null, xs) - Math.min.apply(null, xs),
            height: Math.max.apply(null, ys) - Math.min.apply(null, ys),
            angle: Math.atan2(matrix[1], matrix[0]) * 180 / Math.PI
        };
    }

    function rgbToCss(values) {
        if (!values) return '';
        const channels = Array.from(values).slice(0, 3).map(value => Number(value) || 0);
        const normalizedInput = channels.every(value => value >= 0 && value <= 1);
        const parts = channels.map(value => {
            const normalized = Number(value);
            const channel = normalizedInput ? normalized * 255 : normalized;
            return Math.max(0, Math.min(255, Math.round(channel)));
        });
        return `rgb(${parts[0]}, ${parts[1]}, ${parts[2]})`;
    }

    function grayToCss(value) {
        const channel = Math.max(0, Math.min(255, Math.round(Number(value || 0) * 255)));
        return `rgb(${channel}, ${channel}, ${channel})`;
    }

    function cmykToCss(values) {
        const c = Number(values && values[0] || 0);
        const m = Number(values && values[1] || 0);
        const y = Number(values && values[2] || 0);
        const k = Number(values && values[3] || 0);
        return rgbToCss([
            1 - Math.min(1, c + k),
            1 - Math.min(1, m + k),
            1 - Math.min(1, y + k)
        ]);
    }

    function colorWithAlpha(color, alpha) {
        const opacity = clampOpacity(alpha, 1);
        const value = String(color || '').trim();
        if (/gradient\(/i.test(value)) {
            if (opacity >= 0.999) return value;
            return value
                .replace(/#([0-9a-f]{6})/gi, (_, raw) => {
                    return `rgba(${parseInt(raw.slice(0, 2), 16)}, ${parseInt(raw.slice(2, 4), 16)}, ${parseInt(raw.slice(4, 6), 16)}, ${opacity})`;
                })
                .replace(/rgb\(\s*([\d.]+)\s*,\s*([\d.]+)\s*,\s*([\d.]+)\s*\)/gi, (_, r, g, b) => {
                    return `rgba(${Math.round(Number(r))}, ${Math.round(Number(g))}, ${Math.round(Number(b))}, ${opacity})`;
                });
        }
        const rgbMatch = value.match(/^rgb\(\s*([\d.]+)\s*,\s*([\d.]+)\s*,\s*([\d.]+)\s*\)$/i);
        if (rgbMatch) {
            return `rgba(${Math.round(Number(rgbMatch[1]))}, ${Math.round(Number(rgbMatch[2]))}, ${Math.round(Number(rgbMatch[3]))}, ${opacity})`;
        }
        const hexMatch = value.match(/^#([0-9a-f]{6})$/i);
        if (hexMatch) {
            const raw = hexMatch[1];
            return `rgba(${parseInt(raw.slice(0, 2), 16)}, ${parseInt(raw.slice(2, 4), 16)}, ${parseInt(raw.slice(4, 6), 16)}, ${opacity})`;
        }
        return opacity >= 0.999 ? value : value || 'transparent';
    }

    function pdfPatternToCss(pattern) {
        if (!Array.isArray(pattern)) return '';
        const type = pattern[0];
        const subtype = pattern[1];
        const stops = Array.isArray(pattern[3]) ? pattern[3] : [];
        if (type !== 'RadialAxial' || !stops.length) return '';
        const cssStops = stops.map(stop => {
            const offset = Math.max(0, Math.min(100, Number(stop[0] || 0) * 100));
            return `${stop[1] || '#000000'} ${offset.toFixed(2)}%`;
        }).join(', ');
        if (subtype === 'radial') return `radial-gradient(circle, ${cssStops})`;
        const start = Array.isArray(pattern[4]) ? pattern[4] : [0, 1];
        const end = Array.isArray(pattern[5]) ? pattern[5] : [0, 0];
        const dx = Number(end[0] || 0) - Number(start[0] || 0);
        const dy = Number(end[1] || 0) - Number(start[1] || 0);
        const angle = (Math.atan2(dx, dy) * 180 / Math.PI + 180) % 360;
        return `linear-gradient(${angle.toFixed(2)}deg, ${cssStops})`;
    }

    function parsePdfTilingImagePattern(args) {
        if (!Array.isArray(args) || args[0] !== 'TilingPattern') return null;
        const operatorList = args[2];
        const argsArray = operatorList && Array.isArray(operatorList.argsArray) ? operatorList.argsArray : [];
        let imageId = '';
        for (const entry of argsArray) {
            if (Array.isArray(entry) && typeof entry[0] === 'string' && /^img_/.test(entry[0])) {
                imageId = entry[0];
            }
        }
        if (!imageId) return null;
        return {
            imageId,
            matrix: Array.isArray(args[3]) ? args[3].slice() : null,
            bbox: Array.isArray(args[4]) ? args[4].slice() : null,
            width: Number(args[5]) || 0,
            height: Number(args[6]) || 0
        };
    }

    function readPdfColorFromArgs(fn, args, ops) {
        if (fn === ops.setFillRGBColor || fn === ops.setStrokeRGBColor) return rgbToCss(args || []);
        if (fn === ops.setFillGray || fn === ops.setStrokeGray) return grayToCss(args && args[0]);
        if (fn === ops.setFillCMYKColor || fn === ops.setStrokeCMYKColor) return cmykToCss(args || []);
        return '';
    }

    function isPdfFillColorOp(fn, ops) {
        return fn === ops.setFillRGBColor || fn === ops.setFillGray || fn === ops.setFillCMYKColor ||
            fn === ops.setFillColor || fn === ops.setFillColorN;
    }

    function isPdfStrokeColorOp(fn, ops) {
        return fn === ops.setStrokeRGBColor || fn === ops.setStrokeGray || fn === ops.setStrokeCMYKColor ||
            fn === ops.setStrokeColor || fn === ops.setStrokeColorN;
    }

    function clampOpacity(value, fallback) {
        const next = Number(value);
        if (!Number.isFinite(next)) return fallback;
        return Math.max(0, Math.min(1, next));
    }

    function applyPdfGState(args, state) {
        const entries = Array.isArray(args && args[0]) ? args[0] : args;
        if (!Array.isArray(entries)) return;
        entries.forEach(entry => {
            if (!Array.isArray(entry) || entry.length < 2) return;
            const key = entry[0];
            const value = entry[1];
            if (key === 'LW') state.lineWidth = Math.max(0, Number(value) || state.lineWidth || 1);
            if (key === 'ca') state.fillAlpha = clampOpacity(value, state.fillAlpha);
            if (key === 'CA') state.strokeAlpha = clampOpacity(value, state.strokeAlpha);
        });
    }

    function getPdfLineWidthPx(lineWidth, scale) {
        return Math.max(0.5, (Number(lineWidth) || 1) * scale);
    }

    function getPdfTransformedBox(viewport, currentTransform, rect) {
        const rectMatrix = multiplyMatrix(viewport.transform, multiplyMatrix(currentTransform, [
            rect.width, 0, 0, rect.height, rect.x, rect.y
        ]));
        return getMatrixBoundingBox(rectMatrix, 1, 1);
    }

    function getPdfPathCoordBoundingBox(coords) {
        const xs = [];
        const ys = [];
        for (let index = 0; index + 1 < coords.length; index += 2) {
            xs.push(Number(coords[index]) || 0);
            ys.push(Number(coords[index + 1]) || 0);
        }
        if (!xs.length || !ys.length) return null;
        return {
            x: Math.min.apply(null, xs),
            y: Math.min.apply(null, ys),
            width: Math.max.apply(null, xs) - Math.min.apply(null, xs),
            height: Math.max.apply(null, ys) - Math.min.apply(null, ys)
        };
    }

    function parsePdfPathShape(pathOps, coords, ops) {
        if (!Array.isArray(pathOps) || !Array.isArray(coords)) return null;
        if (pathOps.length === 1 && pathOps[0] === ops.rectangle && coords.length >= 4) {
            return {
                type: 'rect',
                x: coords[0],
                y: coords[1],
                width: coords[2],
                height: coords[3]
            };
        }
        if (pathOps.length === 2 && pathOps[0] === ops.moveTo && pathOps[1] === ops.lineTo && coords.length >= 4) {
            return {
                type: 'line',
                x1: coords[0],
                y1: coords[1],
                x2: coords[2],
                y2: coords[3]
            };
        }
        if (pathOps.length >= 4 &&
            pathOps[0] === ops.moveTo &&
            pathOps[pathOps.length - 1] === ops.closePath &&
            pathOps.every(op => op === ops.moveTo || op === ops.lineTo || op === ops.closePath) &&
            coords.length >= 6) {
            const box = getPdfPathCoordBoundingBox(coords);
            if (box && box.width > 0 && box.height > 0) {
                return Object.assign({ type: 'rect' }, box);
            }
        }
        if (pathOps.length === 6 &&
            pathOps[0] === ops.moveTo &&
            pathOps[1] === ops.curveTo &&
            pathOps[2] === ops.curveTo &&
            pathOps[3] === ops.curveTo &&
            pathOps[4] === ops.curveTo &&
            pathOps[5] === ops.closePath &&
            coords.length >= 26) {
            const box = getPdfPathCoordBoundingBox(coords);
            if (box && box.width > 0 && box.height > 0) {
                return Object.assign({ type: 'ellipse' }, box);
            }
        }
        if (pathOps.length === 6 &&
            pathOps[0] === ops.moveTo &&
            pathOps[1] === ops.lineTo &&
            pathOps[2] === ops.curveTo &&
            pathOps[3] === ops.lineTo &&
            pathOps[4] === ops.lineTo &&
            pathOps[5] === ops.closePath &&
            coords.length >= 14) {
            const box = getPdfPathCoordBoundingBox(coords);
            if (box && box.width > 0 && box.height > 0) {
                return Object.assign({
                    type: 'roundRect',
                    borderTopRightRadius: Math.min(box.width, box.height) * 0.22
                }, box);
            }
        }
        if (pathOps.includes(ops.curveTo) && pathOps.includes(ops.closePath) && coords.length >= 8) {
            const box = getPdfPathCoordBoundingBox(coords);
            if (box && box.width > 0 && box.height > 0) {
                return Object.assign({
                    type: 'roundRect',
                    borderTopRightRadius: Math.min(box.width, box.height) * 0.22
                }, box);
            }
        }
        return null;
    }

    function getPdfLineBox(viewport, currentTransform, line, scale, lineWidth) {
        const matrix = multiplyMatrix(viewport.transform, currentTransform);
        const start = applyMatrixToPoint(matrix, line.x1, line.y1);
        const end = applyMatrixToPoint(matrix, line.x2, line.y2);
        const width = Math.max(1, Math.hypot(end.x - start.x, end.y - start.y) * scale);
        const strokeWidth = getPdfLineWidthPx(lineWidth, scale);
        return {
            type: 'line',
            left: start.x * scale,
            top: start.y * scale - strokeWidth / 2,
            width,
            height: strokeWidth,
            angle: Math.atan2(end.y - start.y, end.x - start.x) * 180 / Math.PI
        };
    }

    function makePdfShapeFromPath(pathShape, paint, state, viewport, scale, zIndex) {
        if (!pathShape) return null;
        if (pathShape.type === 'rect' || pathShape.type === 'ellipse' || pathShape.type === 'roundRect') {
            const box = getPdfTransformedBox(viewport, state.currentTransform, pathShape);
            return {
                type: pathShape.type,
                left: box.left * scale,
                top: box.top * scale,
                width: Math.max(0.5, box.width * scale),
                height: Math.max(0.5, box.height * scale),
                angle: box.angle,
                fillColor: paint.fill ? state.fillColor : 'transparent',
                fillAlpha: paint.fill ? state.fillAlpha : 0,
                strokeColor: paint.stroke ? state.strokeColor : 'transparent',
                strokeAlpha: paint.stroke ? state.strokeAlpha : 0,
                strokeWidth: paint.stroke ? getPdfLineWidthPx(state.lineWidth, scale) : 0,
                borderTopRightRadius: (pathShape.borderTopRightRadius || 0) * scale,
                zIndex
            };
        }
        if (pathShape.type === 'line' && paint.stroke) {
            return Object.assign(getPdfLineBox(viewport, state.currentTransform, pathShape, scale, state.lineWidth), {
                fillColor: state.strokeColor,
                fillAlpha: state.strokeAlpha,
                strokeColor: 'transparent',
                strokeAlpha: 0,
                strokeWidth: 0,
                zIndex
            });
        }
        return null;
    }

    function isFullPageOpaqueFill(shape, viewport, scale) {
        if (!shape || shape.type !== 'rect' || shape.fillAlpha < 0.98 || shape.strokeWidth > 0) return false;
        const pageWidth = viewport.width * scale;
        const pageHeight = viewport.height * scale;
        return shape.width >= pageWidth * 0.82 && shape.height >= pageHeight * 0.82 &&
            shape.left <= pageWidth * 0.1 && shape.top <= pageHeight * 0.1;
    }

    function imageDataToDataUrl(imageData) {
        if (!imageData) return '';
        const canvas = document.createElement('canvas');
        const width = imageData.width || (imageData.bitmap && imageData.bitmap.width);
        const height = imageData.height || (imageData.bitmap && imageData.bitmap.height);
        if (!width || !height) return '';
        const maxSide = 1600;
        const resizeScale = Math.min(1, maxSide / Math.max(width, height));
        canvas.width = Math.max(1, Math.round(width * resizeScale));
        canvas.height = Math.max(1, Math.round(height * resizeScale));
        const context = canvas.getContext('2d');
        if (imageData.bitmap) {
            context.drawImage(imageData.bitmap, 0, 0, canvas.width, canvas.height);
        } else if (imageData.data) {
            const data = imageData.data instanceof Uint8ClampedArray
                ? imageData.data
                : new Uint8ClampedArray(imageData.data);
            if (resizeScale < 1) {
                const sourceCanvas = document.createElement('canvas');
                sourceCanvas.width = width;
                sourceCanvas.height = height;
                sourceCanvas.getContext('2d').putImageData(new ImageData(data, width, height), 0, 0);
                context.drawImage(sourceCanvas, 0, 0, canvas.width, canvas.height);
            } else {
                context.putImageData(new ImageData(data, width, height), 0, 0);
            }
        } else {
            return '';
        }
        return canvas.toDataURL('image/png');
    }

    function withTimeout(promise, timeoutMs, fallbackValue) {
        return new Promise(resolve => {
            let settled = false;
            const timer = setTimeout(() => {
                if (settled) return;
                settled = true;
                resolve(fallbackValue);
            }, timeoutMs);
            Promise.resolve(promise)
                .then(value => {
                    if (settled) return;
                    settled = true;
                    clearTimeout(timer);
                    resolve(value);
                })
                .catch(() => {
                    if (settled) return;
                    settled = true;
                    clearTimeout(timer);
                    resolve(fallbackValue);
                });
        });
    }

    function getPdfObject(objectStore, name) {
        const candidates = [];
        if (typeof name === 'string') {
            const imageIndex = name.indexOf('img_');
            if (imageIndex > 0) candidates.push(name.slice(imageIndex));
        }
        candidates.push(name);
        const uniqueCandidates = candidates.filter((candidate, index) => candidate && candidates.indexOf(candidate) === index);
        for (const candidate of uniqueCandidates) {
            try {
                const value = objectStore.get(candidate);
                if (value) {
                    return Promise.resolve(value);
                }
            } catch (_) {
                // Some PDF.js stores throw before an object has resolved.
            }
        }
        return new Promise(resolve => {
            let settled = false;
            let pending = uniqueCandidates.length;
            const finish = value => {
                if (settled) return;
                if (value) {
                    settled = true;
                    resolve(value);
                    return;
                }
                pending -= 1;
                if (pending <= 0) {
                    settled = true;
                    resolve(null);
                }
            };
            try {
                uniqueCandidates.forEach(candidate => {
                    try {
                        objectStore.get(candidate, finish);
                    } catch (_) {
                        finish(null);
                    }
                });
            } catch (_) {
                finish(null);
            }
            setTimeout(() => finish(null), 3000);
        });
    }

    async function getPdfPageObject(page, name) {
        const stores = [page && page.objs, page && page.commonObjs].filter(Boolean);
        if (!stores.length) return null;
        return new Promise(resolve => {
            let settled = false;
            let pending = stores.length;
            const finish = value => {
                if (settled) return;
                if (value) {
                    settled = true;
                    resolve(value);
                    return;
                }
                pending -= 1;
                if (pending <= 0) {
                    settled = true;
                    resolve(null);
                }
            };
            stores.forEach(store => {
                getPdfObject(store, name).then(finish).catch(() => finish(null));
            });
        });
    }

    async function extractPdfPageLayersInternal(page, viewport, scale, options) {
        options = options || {};
        const debugPdfLayers = Array.isArray(options.debugPdfLayers) ? options.debugPdfLayers : null;
        const pdfjs = requireGlobal('PDF.js', window.pdfjsLib);
        const ops = pdfjs.OPS || {};
        const operatorList = await page.getOperatorList();
        const graphicsStack = [];
        const graphicsState = {
            currentTransform: [1, 0, 0, 1, 0, 0],
            fillColor: 'rgb(0, 0, 0)',
            fillImagePattern: null,
            strokeColor: 'rgb(0, 0, 0)',
            fillAlpha: 1,
            strokeAlpha: 1,
            lineWidth: 1
        };
        let pendingPath = null;
        const textColors = [];
        const shapes = [];
        const images = [];
        const imageTasks = [];
        let backgroundColor = '#ffffff';
        let zIndex = 0;

        for (let index = 0; index < operatorList.fnArray.length; index++) {
            const fn = operatorList.fnArray[index];
            const args = operatorList.argsArray[index] || [];

            if (fn === ops.save) {
                graphicsStack.push({
                    currentTransform: graphicsState.currentTransform.slice(),
                    fillColor: graphicsState.fillColor,
                    fillImagePattern: graphicsState.fillImagePattern,
                    strokeColor: graphicsState.strokeColor,
                    fillAlpha: graphicsState.fillAlpha,
                    strokeAlpha: graphicsState.strokeAlpha,
                    lineWidth: graphicsState.lineWidth,
                    pendingPath
                });
                continue;
            }
            if (fn === ops.restore) {
                const state = graphicsStack.pop();
                if (state) {
                    graphicsState.currentTransform = state.currentTransform;
                    graphicsState.fillColor = state.fillColor;
                    graphicsState.fillImagePattern = state.fillImagePattern;
                    graphicsState.strokeColor = state.strokeColor;
                    graphicsState.fillAlpha = state.fillAlpha;
                    graphicsState.strokeAlpha = state.strokeAlpha;
                    graphicsState.lineWidth = state.lineWidth;
                    pendingPath = state.pendingPath;
                }
                continue;
            }
            if (fn === ops.transform) {
                graphicsState.currentTransform = multiplyMatrix(graphicsState.currentTransform, args);
                continue;
            }
            if (fn === ops.setLineWidth) {
                graphicsState.lineWidth = Math.max(0, Number(args[0]) || 0);
                continue;
            }
            if (fn === ops.setGState) {
                applyPdfGState(args, graphicsState);
                continue;
            }

            const nextColor = readPdfColorFromArgs(fn, args, ops);
            if (nextColor) {
                if (isPdfStrokeColorOp(fn, ops)) {
                    graphicsState.strokeColor = nextColor;
                } else if (isPdfFillColorOp(fn, ops)) {
                    graphicsState.fillColor = nextColor;
                    graphicsState.fillImagePattern = null;
                }
                continue;
            }
            if (fn === ops.setFillColorN) {
                let fillPattern = '';
                let imagePattern = null;
                if (args && args[0] === 'Shading' && args[1]) {
                    fillPattern = pdfPatternToCss(await withTimeout(getPdfPageObject(page, args[1]), 1200, null));
                } else if (args && args[0] === 'TilingPattern') {
                    imagePattern = parsePdfTilingImagePattern(args);
                }
                graphicsState.fillImagePattern = imagePattern;
                graphicsState.fillColor = fillPattern || (imagePattern ? 'transparent' : 'transparent');
                if (debugPdfLayers) debugPdfLayers.push({ event: 'setFillColorN', args0: args && args[0], imagePattern, fillPattern });
                continue;
            }

            if (fn === ops.constructPath) {
                pendingPath = parsePdfPathShape(args[0] || [], args[1] || [], ops);
                if (debugPdfLayers && args && args[0] && args[0].includes && args[0].includes(ops.curveTo)) {
                    debugPdfLayers.push({ event: 'constructPath', ops: Array.from(args[0] || []), coordLength: (args[1] || []).length, pendingPath });
                }
                continue;
            }

            if ((fn === ops.fill || fn === ops.eoFill || fn === ops.fillStroke || fn === ops.eoFillStroke ||
                fn === ops.stroke || fn === ops.closeStroke || fn === ops.closeFillStroke || fn === ops.closeEOFillStroke) && pendingPath) {
                const paint = {
                    fill: fn === ops.fill || fn === ops.eoFill || fn === ops.fillStroke || fn === ops.eoFillStroke ||
                        fn === ops.closeFillStroke || fn === ops.closeEOFillStroke,
                    stroke: fn === ops.stroke || fn === ops.closeStroke || fn === ops.fillStroke || fn === ops.eoFillStroke ||
                        fn === ops.closeFillStroke || fn === ops.closeEOFillStroke
                };
                const nextZIndex = zIndex++;
                const shape = makePdfShapeFromPath(pendingPath, paint, graphicsState, viewport, scale, nextZIndex);
                if (debugPdfLayers && graphicsState.fillImagePattern) {
                    debugPdfLayers.push({ event: 'paintPatternPath', paint, pendingPath, shape, fillImagePattern: graphicsState.fillImagePattern });
                }
                if (shape && paint.fill && graphicsState.fillImagePattern && graphicsState.fillImagePattern.imageId) {
                    const patternBox = {
                        id: `${options.imageIdPrefix || 'pdf-pattern'}-${images.length + 1}`,
                        left: shape.left,
                        top: shape.top,
                        width: shape.width,
                        height: shape.height,
                        angle: shape.angle || 0,
                        zIndex: nextZIndex,
                        borderTopRightRadius: shape.borderTopRightRadius || 0
                    };
                    const imageData = await withTimeout(getPdfPageObject(page, graphicsState.fillImagePattern.imageId), 5000, null);
                    const src = imageDataToDataUrl(imageData);
                    images.push(src ? Object.assign({ src }, patternBox) : Object.assign({ failed: true }, patternBox));
                    pendingPath = null;
                    continue;
                }
                if (isFullPageOpaqueFill(shape, viewport, scale)) {
                    backgroundColor = shape.fillColor || backgroundColor;
                } else if (shape && shape.width >= 0.5 && shape.height >= 0.5) {
                    if (debugPdfLayers) debugPdfLayers.push({ event: 'shape', shape: Object.assign({}, shape) });
                    shapes.push(shape);
                }
                pendingPath = null;
                continue;
            }

            if (fn === ops.endPath) {
                pendingPath = null;
                continue;
            }

            if (fn === ops.showText || fn === ops.showSpacedText || fn === ops.nextLineShowText || fn === ops.nextLineSetSpacingShowText) {
                textColors.push(graphicsState.fillColor || 'rgb(0, 0, 0)');
                continue;
            }

            if (fn === ops.paintImageXObject || fn === ops.paintJpegXObject || fn === ops.paintInlineImageXObject) {
                const imageMatrix = multiplyMatrix(viewport.transform, graphicsState.currentTransform);
                const box = getMatrixBoundingBox(imageMatrix, 1, 1);
                const imageBox = {
                    id: `${options.imageIdPrefix || 'pdf-image'}-${images.length + 1}`,
                    left: box.left * scale,
                    top: box.top * scale,
                    width: Math.max(1, box.width * scale),
                    height: Math.max(1, box.height * scale),
                    angle: box.angle,
                    zIndex: zIndex++
                };
                const loadImage = async () => {
                    try {
                        let imageData = null;
                        if (fn === ops.paintInlineImageXObject) {
                            imageData = args[0];
                        } else if (args[0]) {
                            const timeoutMs = options.deferImages ? 8000 : 3500;
                            imageData = await withTimeout(getPdfPageObject(page, args[0]), timeoutMs, null);
                        }
                        const src = imageDataToDataUrl(imageData);
                        if (debugPdfLayers && (!src || !imageData)) {
                            debugPdfLayers.push({
                                event: 'imageLoad',
                                imageId: args[0],
                                hasImageData: !!imageData,
                                hasBitmap: !!(imageData && imageData.bitmap),
                                hasData: !!(imageData && imageData.data),
                                width: imageData && (imageData.width || (imageData.bitmap && imageData.bitmap.width)),
                                height: imageData && (imageData.height || (imageData.bitmap && imageData.bitmap.height)),
                                srcLength: src.length
                            });
                        }
                        return src ? Object.assign({ src }, imageBox) : Object.assign({ failed: true }, imageBox);
                    } catch (_) {
                        return Object.assign({ failed: true }, imageBox);
                    }
                };
                if (options.deferImages) {
                    images.push(Object.assign({ placeholder: true }, imageBox));
                    imageTasks.push(loadImage());
                } else {
                    const loadedImage = await loadImage();
                    if (loadedImage) images.push(loadedImage);
                }
                continue;
            }
        }

        return { backgroundColor, textColors, shapes, images, imageTasks };
    }

    async function extractPdfPageLayers(page, viewport, scale, options) {
        return withTimeout(
            extractPdfPageLayersInternal(page, viewport, scale, options),
            options && options.deferImages ? 8000 : 15000,
            { backgroundColor: '#ffffff', textColors: [], shapes: [], images: [], imageTasks: [] }
        );
    }

    async function renderPdfPageToDataUrl(page, maxWidthPx) {
        const viewport = page.getViewport({ scale: 1 });
        const scale = Math.min(2, Math.max(1, (maxWidthPx || 900) / viewport.width));
        const scaledViewport = page.getViewport({ scale });
        const canvas = document.createElement('canvas');
        canvas.width = Math.ceil(scaledViewport.width);
        canvas.height = Math.ceil(scaledViewport.height);
        const context = canvas.getContext('2d', { alpha: false });
        await page.render({ canvasContext: context, viewport: scaledViewport }).promise;
        return {
            dataUrl: canvas.toDataURL('image/png'),
            width: Math.round(Math.min(620, scaledViewport.width)),
            height: Math.round(Math.min(620, scaledViewport.width) * scaledViewport.height / scaledViewport.width)
        };
    }

    function escapeHtml(value) {
        return String(value || '')
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;')
            .replace(/'/g, '&#39;');
    }

    function renderPdfImageBox(image) {
        const radius = Number(image.borderTopRightRadius || 0);
        const radiusStyle = radius > 0 ? `border-top-right-radius:${radius.toFixed(2)}px;overflow:hidden;` : '';
        const style = `left:${image.left.toFixed(2)}px;top:${image.top.toFixed(2)}px;width:${image.width.toFixed(2)}px;height:${image.height.toFixed(2)}px;transform:rotate(${image.angle.toFixed(2)}deg);z-index:${image.zIndex};${radiusStyle}`;
        const idAttr = image.id ? ` data-pdf-image-id="${escapeHtml(image.id)}"` : '';
        if (image.failed) {
            return `<div class="pdf-image-box pdf-image-missing"${idAttr} style="${style}"></div>`;
        }
        if (image.placeholder) {
            return `<div class="pdf-image-box pdf-image-placeholder"${idAttr} style="${style}"></div>`;
        }
        return `<img class="pdf-image-box"${idAttr} src="${image.src}" alt="" style="${style}">`;
    }

    function renderPdfShapeBox(shape) {
        if (!shape) return '';
        const background = shape.fillAlpha > 0
            ? colorWithAlpha(shape.fillColor || 'transparent', shape.fillAlpha)
            : 'transparent';
        const border = shape.strokeWidth > 0 && shape.strokeAlpha > 0
            ? `${shape.strokeWidth.toFixed(2)}px solid ${colorWithAlpha(shape.strokeColor || 'transparent', shape.strokeAlpha)}`
            : '0 solid transparent';
        const className = shape.type === 'line'
            ? 'pdf-shape-box pdf-shape-line'
            : shape.type === 'ellipse'
                ? 'pdf-shape-box pdf-shape-ellipse'
                : 'pdf-shape-box pdf-shape-rect';
        const radius = Number(shape.borderTopRightRadius || 0);
        const style = [
            `left:${shape.left.toFixed(2)}px`,
            `top:${shape.top.toFixed(2)}px`,
            `width:${shape.width.toFixed(2)}px`,
            `height:${shape.height.toFixed(2)}px`,
            `background:${escapeHtml(background)}`,
            `border:${escapeHtml(border)}`,
            shape.type === 'ellipse' ? 'border-radius:50%' : '',
            shape.type === 'roundRect' && radius > 0 ? `border-top-right-radius:${radius.toFixed(2)}px` : '',
            `transform:rotate(${(shape.angle || 0).toFixed(2)}deg)`,
            `z-index:${shape.zIndex || 0}`
        ].filter(Boolean).join(';');
        return `<div class="${className}" data-pdf-shape="true" style="${style};"></div>`;
    }

    async function createPdfPageRenderSampler(page, viewport, finalScale) {
        try {
            const sampleScale = Math.min(0.35, Math.max(0.12, 220 / Math.max(1, viewport.width)));
            const sampleViewport = page.getViewport({ scale: sampleScale });
            const canvas = document.createElement('canvas');
            canvas.width = Math.max(1, Math.ceil(sampleViewport.width));
            canvas.height = Math.max(1, Math.ceil(sampleViewport.height));
            const context = canvas.getContext('2d', { alpha: false });
            await page.render({ canvasContext: context, viewport: sampleViewport }).promise;
            const imageData = context.getImageData(0, 0, canvas.width, canvas.height);
            const ratio = sampleScale / Math.max(0.01, finalScale || 1);
            return rect => getAverageRenderedColor(imageData, canvas.width, canvas.height, rect, ratio);
        } catch (_) {
            return null;
        }
    }

    function getAverageRenderedColor(imageData, canvasWidth, canvasHeight, rect, ratio) {
        if (!imageData || !rect || !ratio) return null;
        const left = Math.max(0, Math.floor((Number(rect.left) || 0) * ratio));
        const top = Math.max(0, Math.floor((Number(rect.top) || 0) * ratio));
        const right = Math.min(canvasWidth, Math.ceil(((Number(rect.left) || 0) + (Number(rect.width) || 0)) * ratio));
        const bottom = Math.min(canvasHeight, Math.ceil(((Number(rect.top) || 0) + (Number(rect.height) || 0)) * ratio));
        if (right <= left || bottom <= top) return null;
        const stepX = Math.max(1, Math.floor((right - left) / 12));
        const stepY = Math.max(1, Math.floor((bottom - top) / 12));
        let totalR = 0;
        let totalG = 0;
        let totalB = 0;
        let count = 0;
        for (let y = top; y < bottom; y += stepY) {
            for (let x = left; x < right; x += stepX) {
                const index = (y * canvasWidth + x) * 4;
                const r = imageData.data[index];
                const g = imageData.data[index + 1];
                const b = imageData.data[index + 2];
                totalR += r;
                totalG += g;
                totalB += b;
                count += 1;
            }
        }
        if (!count) return null;
        return getColorMetrics(totalR / count, totalG / count, totalB / count);
    }

    function isLikelyPdfMaskShape(shape, images, pageViewport, renderedSampler, siblingShapes) {
        if (!shape || !Array.isArray(images) || !images.length) return false;
        const pageWidth = pageViewport && pageViewport.width ? pageViewport.width : 0;
        const pageHeight = pageViewport && pageViewport.height ? pageViewport.height : 0;
        if (!pageWidth || !pageHeight) return false;
        if (shape.strokeWidth > 0 || shape.fillAlpha < 0.95) return false;
        const fillMetrics = getPdfFillColorMetrics(shape.fillColor);
        if (!fillMetrics) return false;
        const shapeAreaRatio = ((Number(shape.width) || 0) * (Number(shape.height) || 0)) / Math.max(1, pageWidth * pageHeight);
        if (/gradient\(/i.test(String(shape.fillColor || '')) &&
            fillMetrics.luminance >= 0.72 &&
            hasSameBoundsWhiteFillUnderlay(shape, siblingShapes)) {
            return true;
        }
        if (typeof renderedSampler === 'function') {
            const renderedMetrics = renderedSampler(shape);
            if (renderedMetrics) {
                const colorDistance = getColorDistance(fillMetrics, renderedMetrics);
                if (fillMetrics.darkNeutral && renderedMetrics.luminance >= 0.58) return true;
                if (fillMetrics.saturation >= 0.12 &&
                    fillMetrics.luminance >= 0.72 &&
                    shapeAreaRatio >= 0.025 &&
                    colorDistance >= 0.16 &&
                    renderedMetrics.saturation <= fillMetrics.saturation * 0.72) {
                    return true;
                }
            }
        }
        return fillMetrics.darkNeutral && images.some(image => getPdfRectOverlapRatio(shape, image) >= 0.48);
    }

    function hasSameBoundsWhiteFillUnderlay(shape, siblingShapes) {
        if (!shape || !Array.isArray(siblingShapes)) return false;
        return siblingShapes.some(other => {
            if (!other || other === shape || (other.zIndex || 0) >= (shape.zIndex || 0)) return false;
            if (!isWhiteSolidPdfShape(other)) return false;
            return getPdfRectOverlapRatio(shape, other) >= 0.96 && getPdfRectOverlapRatio(other, shape) >= 0.96;
        });
    }

    function isWhiteSolidPdfShape(shape) {
        if (!shape || shape.strokeWidth > 0 || shape.fillAlpha < 0.95) return false;
        const metrics = getPdfFillColorMetrics(shape.fillColor);
        return !!(metrics && metrics.luminance >= 0.96 && metrics.saturation <= 0.04);
    }

    function isDarkNeutralGradient(value) {
        const metrics = getPdfFillColorMetrics(value);
        return !!(metrics && metrics.darkNeutral);
    }

    function getPdfFillColorMetrics(value) {
        const text = String(value || '');
        const colors = text.match(/#[0-9a-f]{6}/gi) || [];
        if (/linear-gradient/i.test(text) && colors.length >= 2) {
            const samples = colors.map(color => getColorMetrics(
                parseInt(color.slice(1, 3), 16),
                parseInt(color.slice(3, 5), 16),
                parseInt(color.slice(5, 7), 16)
            ));
            const average = samples.reduce((memo, sample) => {
                memo.r += sample.r;
                memo.g += sample.g;
                memo.b += sample.b;
                memo.luminance += sample.luminance;
                memo.saturation += sample.saturation;
                return memo;
            }, { r: 0, g: 0, b: 0, luminance: 0, saturation: 0 });
            Object.keys(average).forEach(key => { average[key] /= samples.length; });
            average.darkNeutral = samples.every(sample => sample.luminance <= 0.38 && sample.saturation <= 0.12);
            return average;
        }
        const rgbMatch = text.match(/rgba?\(\s*([\d.]+)\s*,\s*([\d.]+)\s*,\s*([\d.]+)/i);
        if (!rgbMatch) return null;
        const metrics = getColorMetrics(Number(rgbMatch[1]), Number(rgbMatch[2]), Number(rgbMatch[3]));
        metrics.darkNeutral = metrics.luminance <= 0.38 && metrics.saturation <= 0.12;
        return metrics;
    }

    function getColorMetrics(r, g, b) {
        const red = Math.max(0, Math.min(255, Number(r) || 0));
        const green = Math.max(0, Math.min(255, Number(g) || 0));
        const blue = Math.max(0, Math.min(255, Number(b) || 0));
        const maxChannel = Math.max(red, green, blue);
        const minChannel = Math.min(red, green, blue);
        return {
            r: red,
            g: green,
            b: blue,
            luminance: (0.2126 * red + 0.7152 * green + 0.0722 * blue) / 255,
            saturation: maxChannel === 0 ? 0 : (maxChannel - minChannel) / maxChannel
        };
    }

    function getColorDistance(a, b) {
        if (!a || !b) return 0;
        const dr = (a.r - b.r) / 255;
        const dg = (a.g - b.g) / 255;
        const db = (a.b - b.b) / 255;
        return Math.sqrt((dr * dr + dg * dg + db * db) / 3);
    }

    function getPdfRectOverlapRatio(a, b) {
        if (!a || !b) return 0;
        const left = Math.max(Number(a.left) || 0, Number(b.left) || 0);
        const top = Math.max(Number(a.top) || 0, Number(b.top) || 0);
        const right = Math.min((Number(a.left) || 0) + (Number(a.width) || 0), (Number(b.left) || 0) + (Number(b.width) || 0));
        const bottom = Math.min((Number(a.top) || 0) + (Number(a.height) || 0), (Number(b.top) || 0) + (Number(b.height) || 0));
        const overlap = Math.max(0, right - left) * Math.max(0, bottom - top);
        const shapeArea = Math.max(1, (Number(a.width) || 0) * (Number(a.height) || 0));
        return overlap / shapeArea;
    }

    function renderPdfLoadedImage(image) {
        if (!image || (!image.src && !image.failed)) return null;
        return {
            id: image.id,
            html: renderPdfImageBox(image),
            src: image.src || '',
            failed: !!image.failed
        };
    }

    function getPdfTextItemStyle(item, textStyle, viewport, scale, color) {
        const transform = pdfjsLib.Util.transform(viewport.transform, item.transform);
        const fontHeight = Math.hypot(transform[2], transform[3]) || (item.height || 12) * scale;
        const left = transform[4];
        const ascent = Number.isFinite(textStyle && textStyle.ascent) ? textStyle.ascent : 0.8;
        const top = transform[5] - fontHeight * ascent;
        const width = Math.max(1, (item.width || String(item.str || '').length * fontHeight * 0.5) * scale);
        const fontFamily = textStyle && textStyle.fontFamily ? textStyle.fontFamily : 'Arial, sans-serif';
        return {
            left,
            top,
            width,
            height: fontHeight * 1.25,
            fontSize: fontHeight,
            fontFamily,
            color: color || '#111111'
        };
    }

    function normalizePdfFontFamily(fontFamily) {
        return String(fontFamily || 'Arial, sans-serif').replace(/\s+/g, ' ').trim();
    }

    function shouldInsertPdfTextSpace(prev, next) {
        if (!prev || !next) return false;
        if (/\s$/.test(prev.text) || /^\s/.test(next.text)) return false;
        if (/^[，。！？；：、,.!?;:%）)\]}》〉」』]/.test(next.text)) return false;
        if (/[（([{《〈「『]$/.test(prev.text)) return false;
        const prevAscii = /[A-Za-z0-9]$/.test(prev.text);
        const nextAscii = /^[A-Za-z0-9]/.test(next.text);
        return prevAscii && nextAscii;
    }

    function canMergePdfTextBoxes(prev, next) {
        if (!prev || !next) return false;
        const baselineTolerance = Math.max(2.5, Math.min(prev.fontSize, next.fontSize) * 0.35);
        const sameLine = Math.abs(prev.baseline - next.baseline) <= baselineTolerance;
        if (!sameLine) return false;
        const sizeClose = Math.abs(prev.fontSize - next.fontSize) <= Math.max(1.2, prev.fontSize * 0.16);
        const sameFont = normalizePdfFontFamily(prev.fontFamily) === normalizePdfFontFamily(next.fontFamily);
        const gap = next.left - prev.right;
        const maxGap = Math.max(3, Math.min(prev.fontSize, next.fontSize) * 0.9);
        return sizeClose && sameFont && gap >= -2 && gap <= maxGap;
    }

    function canMergePdfTextLineIntoParagraph(paragraph, nextLine) {
        if (!paragraph || !nextLine || !paragraph.lastLine) return false;
        const prevLine = paragraph.lastLine;
        const sizeClose = Math.abs(prevLine.fontSize - nextLine.fontSize) <= Math.max(1.2, prevLine.fontSize * 0.14);
        const sameFont = normalizePdfFontFamily(prevLine.fontFamily) === normalizePdfFontFamily(nextLine.fontFamily);
        if (!sizeClose || !sameFont) return false;

        const leftTolerance = Math.max(8, Math.min(prevLine.fontSize, nextLine.fontSize) * 1.1);
        const alignedLeft = Math.abs(nextLine.left - paragraph.left) <= leftTolerance;
        const mildIndent = nextLine.left > paragraph.left && nextLine.left - paragraph.left <= Math.max(18, nextLine.fontSize * 1.8);
        if (!alignedLeft && !mildIndent) return false;

        const baselineGap = nextLine.baseline - prevLine.baseline;
        const minGap = Math.max(4, Math.min(prevLine.fontSize, nextLine.fontSize) * 0.75);
        const maxGap = Math.max(10, Math.max(prevLine.fontSize, nextLine.fontSize) * 1.95);
        if (baselineGap < minGap || baselineGap > maxGap) return false;

        const paragraphWidth = Math.max(paragraph.width, prevLine.width, nextLine.width);
        const lineLooksRelated = nextLine.width >= paragraphWidth * 0.22 || /[，。！？；：、,.!?;:]$/.test(prevLine.text);
        return lineLooksRelated;
    }

    function appendPdfTextSegments(target, source, spacer) {
        target.segments = Array.isArray(target.segments)
            ? target.segments
            : [{ text: target.text || '', color: target.color || '#111111' }];
        const sourceSegments = Array.isArray(source && source.segments) && source.segments.length
            ? source.segments
            : [{ text: source && source.text || '', color: source && source.color || target.color || '#111111' }];
        if (spacer) {
            target.segments.push({ text: spacer, color: target.color || '#111111' });
        }
        sourceSegments.forEach(segment => {
            if (segment && segment.text) target.segments.push({ text: segment.text, color: segment.color || target.color || '#111111' });
        });
    }

    function renderPdfTextBoxContent(box) {
        const segments = Array.isArray(box && box.segments) ? box.segments.filter(segment => segment && segment.text) : [];
        if (!segments.length) return escapeHtml(box && box.text || '');

        const baseColor = String(box.color || '#111111').toLowerCase();
        const hasInlineStyle = segments.some(segment => String(segment.color || baseColor).toLowerCase() !== baseColor);
        if (!hasInlineStyle) {
            return escapeHtml(segments.map(segment => segment.text || '').join(''));
        }

        return segments.map(segment => {
            const text = escapeHtml(segment.text || '');
            const color = String(segment.color || box.color || '#111111');
            if (color.toLowerCase() === baseColor) return text;
            return `<span style="color:${escapeHtml(color)};">${text}</span>`;
        }).join('');
    }

    function mergePdfTextLinesIntoParagraphs(lineBoxes) {
        const paragraphs = [];
        lineBoxes.forEach(line => {
            let prev = null;
            for (let index = paragraphs.length - 1; index >= 0; index -= 1) {
                const candidate = paragraphs[index];
                if (canMergePdfTextLineIntoParagraph(candidate, line)) {
                    prev = candidate;
                    break;
                }
            }
            if (canMergePdfTextLineIntoParagraph(prev, line)) {
                const spacer = shouldInsertPdfTextSpace(prev, line) ? ' ' : '';
                prev.text += spacer + line.text;
                appendPdfTextSegments(prev, line, spacer);
                prev.left = Math.min(prev.left, line.left);
                prev.top = Math.min(prev.top, line.top);
                prev.right = Math.max(prev.right, line.right);
                prev.width = Math.max(1, prev.right - prev.left);
                prev.height = Math.max(prev.height, line.top + line.height - prev.top);
                prev.fontSize = Math.max(prev.fontSize, line.fontSize);
                prev.lineCount += 1;
                prev.lineHeight = (line.baseline - prev.lastLine.baseline) / Math.max(1, prev.fontSize);
                prev.lastLine = line;
            } else {
                paragraphs.push(Object.assign({ lineCount: 1, lineHeight: 1.25, lastLine: line }, line));
            }
        });
        return paragraphs.map(paragraph => {
            delete paragraph.lastLine;
            paragraph.isParagraph = paragraph.lineCount > 1;
            if (paragraph.isParagraph) {
                paragraph.height = Math.max(paragraph.height, paragraph.fontSize * paragraph.lineHeight * paragraph.lineCount);
            }
            return paragraph;
        });
    }

    function mergePdfTextItems(items, styleMap, viewport, scaledViewport, scale, textColors) {
        const rawBoxes = items
            .filter(item => item.str && item.str.trim())
            .map((item, index) => {
                const box = getPdfTextItemStyle(item, styleMap[item.fontName], scaledViewport, scale, textColors[index]);
                return Object.assign(box, {
                    text: item.str,
                    right: box.left + box.width,
                    baseline: box.top + box.fontSize * 0.8
                });
            })
            .sort((a, b) => Math.abs(a.baseline - b.baseline) > Math.max(2, Math.min(a.fontSize, b.fontSize) * 0.35)
                ? a.baseline - b.baseline
                : a.left - b.left);

        const merged = [];
        rawBoxes.forEach(box => {
            const prev = merged[merged.length - 1];
            if (canMergePdfTextBoxes(prev, box)) {
                const spacer = shouldInsertPdfTextSpace(prev, box) ? ' ' : '';
                prev.text += spacer + box.text;
                appendPdfTextSegments(prev, box, spacer);
                prev.right = Math.max(prev.right, box.right);
                prev.top = Math.min(prev.top, box.top);
                prev.height = Math.max(prev.height, box.top + box.height - prev.top);
                prev.width = Math.max(1, prev.right - prev.left);
                prev.fontSize = Math.max(prev.fontSize, box.fontSize);
                prev.baseline = Math.max(prev.baseline, box.baseline);
            } else {
                merged.push(Object.assign({
                    segments: [{ text: box.text, color: box.color || '#111111' }]
                }, box));
            }
        });

        return mergePdfTextLinesIntoParagraphs(merged).map(box => {
            const maxWidth = Math.max(1, viewport.width * scale - box.left);
            const minReadableWidth = box.isParagraph ? box.width : box.text.length * box.fontSize * 0.35;
            const selectionSlack = Math.max(4, Math.min(14, box.fontSize * (box.isParagraph ? 0.45 : 0.7)));
            box.width = Math.min(Math.max(box.width, minReadableWidth) + selectionSlack, maxWidth);
            box.height = Math.max(box.height, box.fontSize * 1.25);
            return box;
        });
    }

    function renderPdfTextBox(box) {
        const style = [
            `left:${box.left.toFixed(2)}px`,
            `top:${box.top.toFixed(2)}px`,
            `width:${box.width.toFixed(2)}px`,
            `min-height:${box.height.toFixed(2)}px`,
            `font-size:${box.fontSize.toFixed(2)}px`,
            `font-family:${escapeHtml(box.fontFamily)}`,
            `color:${escapeHtml(box.color)}`,
            'background-color:rgba(255,255,255,0)',
            'border:0 solid transparent',
            box.isParagraph ? `line-height:${Math.max(1, Math.min(2.4, box.lineHeight || 1.35)).toFixed(3)}` : 'line-height:1',
            box.isParagraph ? 'white-space:normal' : 'white-space:pre'
        ].join(';');
        const paragraphAttr = box.isParagraph ? ' data-pdf-paragraph="true"' : '';
        return `<div class="editor-text-box pdf-text-box" contenteditable="true" data-text-box="true" data-pdf-text-box="true"${paragraphAttr} style="${style};">${renderPdfTextBoxContent(box)}</div>`;
    }

    async function pdfFileToEditorHtml(file, options) {
        options = options || {};
        const pdfjs = configurePdfWorker();
        const pdf = await pdfjs.getDocument({ data: await file.arrayBuffer() }).promise;
        const pages = [];
        const deferredImagePromises = [];

        for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber++) {
            const page = await pdf.getPage(pageNumber);
            const viewport = page.getViewport({ scale: 1 });
            const scale = Math.min(1.6, Math.max(0.8, 760 / viewport.width));
            const scaledViewport = page.getViewport({ scale });
            const imageIdPrefix = `pdf-p${pageNumber}-img`;
            const layers = await extractPdfPageLayers(page, viewport, scale, {
                deferImages: !!options.deferImages,
                imageIdPrefix,
                debugPdfLayers: options.debugPdfLayers
            });
            const textContent = await page.getTextContent();
            const styleMap = textContent.styles || {};
            const textBoxes = mergePdfTextItems(textContent.items, styleMap, viewport, scaledViewport, scale, layers.textColors || [])
                .map(renderPdfTextBox)
                .join('');
            const renderedSampler = await createPdfPageRenderSampler(page, viewport, scale);
            const visibleShapes = (layers.shapes || []).filter(shape => !isLikelyPdfMaskShape(shape, layers.images || [], scaledViewport, renderedSampler, layers.shapes || []));
            const objectBoxes = []
                .concat(visibleShapes.map(shape => ({ zIndex: shape.zIndex || 0, html: renderPdfShapeBox(shape) })))
                .concat((layers.images || []).map(image => ({ zIndex: image.zIndex || 0, html: renderPdfImageBox(image) })))
                .sort((a, b) => a.zIndex - b.zIndex)
                .map(item => item.html)
                .join('');

            const pageHtml = `<div class="pdf-page-layout" data-pdf-page="${pageNumber}" style="width:${scaledViewport.width.toFixed(2)}px;height:${scaledViewport.height.toFixed(2)}px;background:${escapeHtml(layers.backgroundColor || '#ffffff')};">
                <div class="pdf-object-layer pdf-shape-layer pdf-image-layer">${objectBoxes}</div>
                <div class="pdf-text-layer">${textBoxes}</div>
            </div><p><br></p>`;
            pages.push(pageHtml);
            if (typeof options.onPageHtml === 'function') {
                await options.onPageHtml(pageHtml, pageNumber, pdf.numPages);
            }
            if (typeof options.onPageImage === 'function' && layers.imageTasks && layers.imageTasks.length) {
                layers.imageTasks.forEach(task => {
                    const imagePromise = Promise.resolve(task)
                        .then(renderPdfLoadedImage)
                        .then(loaded => {
                            if (loaded) options.onPageImage(loaded, pageNumber, pdf.numPages);
                            return loaded;
                        });
                    deferredImagePromises.push(imagePromise);
                });
            }
            await new Promise(resolve => setTimeout(resolve, 0));
        }

        if (options.waitForDeferredImages && deferredImagePromises.length) {
            await Promise.allSettled(deferredImagePromises);
        }

        return pages.join('');
    }

    async function pdfFileToDocxFile(file) {
        const pdfjs = configurePdfWorker();
        const docxLib = requireGlobal('docx.js', window.docx);
        const pdf = await pdfjs.getDocument({ data: await file.arrayBuffer() }).promise;
        const children = [];

        for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber++) {
            const page = await pdf.getPage(pageNumber);
            const textContent = await page.getTextContent();
            const lines = [];
            let currentY = null;
            let currentLine = [];
            textContent.items.forEach(item => {
                const y = item.transform && item.transform.length > 5 ? Math.round(item.transform[5]) : 0;
                if (currentY !== null && Math.abs(y - currentY) > 5 && currentLine.length) {
                    lines.push(currentLine.join(' ').replace(/\s+/g, ' ').trim());
                    currentLine = [];
                }
                currentY = y;
                if (item.str) currentLine.push(item.str);
            });
            if (currentLine.length) lines.push(currentLine.join(' ').replace(/\s+/g, ' ').trim());

            if (pageNumber > 1) {
                children.push(new docxLib.Paragraph({ children: [new docxLib.PageBreak()] }));
            }
            children.push(new docxLib.Paragraph({
                heading: docxLib.HeadingLevel.HEADING_1,
                children: [new docxLib.TextRun({ text: `第 ${pageNumber} 页` })]
            }));
            const writtenLines = lines.filter(Boolean);
            if (writtenLines.length) {
                writtenLines.forEach(line => children.push(new docxLib.Paragraph({
                    children: [new docxLib.TextRun({ text: line })]
                })));
            } else {
                children.push(new docxLib.Paragraph({
                    children: [new docxLib.TextRun({ text: '（本页未提取到可编辑文本，可能是扫描版 PDF。）', italics: true })]
                }));
            }

            const pageImage = await renderPdfPageToDataUrl(page, 900);
            children.push(new docxLib.Paragraph({
                children: [new docxLib.ImageRun({
                    data: dataUrlToUint8Array(pageImage.dataUrl),
                    transformation: { width: pageImage.width, height: pageImage.height }
                })]
            }));
        }

        const document = new docxLib.Document({
            sections: [{
                properties: {},
                children
            }]
        });
        const blob = await docxLib.Packer.toBlob(document);
        const filename = safeFilename(String(file.name || 'converted.pdf').replace(/\.[^.]+$/, '.docx'), 'converted.docx');
        return new File([blob], filename, {
            type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        });
    }

    function getPptxConstructor() {
        return window.pptxgenjs || window.pptxgen || window.PptxGenJS;
    }

    async function exportDeckToPptxBlob(deck) {
        const PptxGen = requireGlobal('PptxGenJS', getPptxConstructor());
        const pptx = new PptxGen();
        const firstSlide = deck.slides && deck.slides[0] ? deck.slides[0] : {};
        const slideWidth = Number(firstSlide.width || deck.slideWidth || 1280);
        const slideHeight = Number(firstSlide.height || deck.slideHeight || 720);
        pptx.defineLayout({ name: 'LOCAL_EDITOR', width: pxToIn(slideWidth), height: pxToIn(slideHeight) });
        pptx.layout = 'LOCAL_EDITOR';
        pptx.author = 'Local Word Editor';
        pptx.subject = 'Edited presentation';
        pptx.title = deck.filename || 'presentation.pptx';

        (deck.slides || []).forEach(slideData => {
            const slide = pptx.addSlide();
            const bg = String(slideData.background || '#ffffff').replace('#', '');
            if (/^[0-9a-fA-F]{6}$/.test(bg)) slide.background = { color: bg.toUpperCase() };
            (slideData.elements || [])
                .slice()
                .sort((a, b) => (a.zIndex || 0) - (b.zIndex || 0))
                .forEach(element => {
                    const box = {
                        x: pxToIn(element.left),
                        y: pxToIn(element.top),
                        w: pxToIn(element.width || 120),
                        h: pxToIn(element.height || 40)
                    };
                    if (element.type === 'image' && element.src) {
                        slide.addImage(Object.assign({ data: element.src }, box));
                    } else if (element.type === 'text') {
                        const style = element.style || {};
                        slide.addText(String(element.text || ''), Object.assign(box, {
                            fontFace: style.fontFamily || 'Arial',
                            fontSize: Number(style.fontSize || 24),
                            color: String(style.color || '#111111').replace('#', '').toUpperCase(),
                            bold: !!style.bold,
                            italic: !!style.italic,
                            fit: 'shrink'
                        }));
                    }
                });
        });

        return pptx.write({ outputType: 'blob' });
    }

    window.LocalOfficeConverters = {
        pdfFileToDocxFile,
        pdfFileToEditorHtml,
        inspectPptxFile,
        storePptDeck,
        takeStoredPptDeck,
        clearStoredPptDecks,
        exportDeckToPptxBlob
    };
})();
