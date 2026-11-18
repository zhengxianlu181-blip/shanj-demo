importScripts('lamejs.min.js');

const encoder = new lamejs.Mp3Encoder(1, 16000, 64);
const chunks = [];
let totalBytes = 0;
let sampleCount = 0;
let pendingByte = null;

function append(chunk) {
  if (!chunk || chunk.length === 0) return new Uint8Array(0);
  const copy = new Uint8Array(chunk);
  chunks.push(copy);
  totalBytes += copy.length;
  return copy;
}

function encode(bytes) {
  const length = bytes.length + (pendingByte === null ? 0 : 1);
  const pcm = new Uint8Array(length);
  if (pendingByte !== null) pcm[0] = pendingByte;
  pcm.set(bytes, length - bytes.length);
  const count = Math.floor(length / 2);
  pendingByte = length % 2 === 0 ? null : pcm[length - 1];
  if (count === 0) return new Uint8Array(0);
  const samples = new Int16Array(count);
  for (let index = 0; index < count; index += 1) {
    const value = pcm[index * 2] | (pcm[index * 2 + 1] << 8);
    samples[index] = value < 0x8000 ? value : value - 0x10000;
  }
  sampleCount += count;
  const encoded = encoder.encodeBuffer(samples);
  return append(encoded);
}

self.onmessage = (event) => {
  try {
    const message = event.data;
    if (message.type === 'pcm') {
      self.postMessage({
        type: 'encoded',
        sequence: message.sequence,
        bytes: encode(message.bytes),
      });
      return;
    }
    if (message.type !== 'finish') return;
    if (pendingByte !== null) throw new Error('Incomplete PCM16 sample');
    if (sampleCount === 0) throw new Error('No audio samples captured');
    const finishStart = performance.now();
    const flushStart = performance.now();
    append(encoder.flush());
    const flushMs = performance.now() - flushStart;
    const concatStart = performance.now();
    const output = new Uint8Array(totalBytes);
    let offset = 0;
    for (const chunk of chunks) {
      output.set(chunk, offset);
      offset += chunk.length;
    }
    const concatMs = performance.now() - concatStart;
    self.postMessage(
      {
        type: 'finished',
        bytes: output,
        finishMs: performance.now() - finishStart,
        flushMs,
        concatMs,
        chunkCount: chunks.length,
        totalBytes,
      },
      [output.buffer],
    );
  } catch (error) {
    self.postMessage({ type: 'error', message: String(error) });
  }
};

self.postMessage({ type: 'ready' });
