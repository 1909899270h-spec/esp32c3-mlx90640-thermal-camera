const srcWidth = 32;
const srcHeight = 24;
const frameSize = srcWidth * srcHeight;

const canvas = document.getElementById('thermalCanvas');
const ctx = canvas.getContext('2d');
ctx.imageSmoothingEnabled = true;

const importButton = document.getElementById('importDataButton');
const detectButton = document.getElementById('detectButton');
const stopButton = document.getElementById('stopButton');
const downloadButton = document.getElementById('downloadButton');
const clearButton = document.getElementById('clearButton');
const fullscreenButton = document.getElementById('fullscreenButton');
const fpsEl = document.getElementById('fps');
const minTempEl = document.getElementById('minTemp');
const maxTempEl = document.getElementById('maxTemp');
const scaleMinEl = document.getElementById('scaleMin');
const scaleMaxEl = document.getElementById('scaleMax');
const scaleMidEl = document.getElementById('scaleMid');
const colorBarEl = document.getElementById('colorBar');
const frameNoEl = document.getElementById('frameNo');
const recordCountEl = document.getElementById('recordCount');
const sourceLabel = document.getElementById('sourceLabel');
const colormapSelect = document.getElementById('colormap');
const smoothingSelect = document.getElementById('smoothingMode');

let streaming = false;
let activeMode = 'idle';
let frameCount = 0;
let fpsFrameCounter = 0;
let fpsTime = performance.now();
let lastData = null;
let timerId = null;
let records = [];

const simState = {
  frame: 0,
  previous: null,
  fixedOffsets: Array.from({ length: frameSize }, (_, i) => {
    const x = i % srcWidth;
    const y = Math.floor(i / srcWidth);
    return 0.16 * Math.sin(x * 0.47 + y * 0.19) + 0.10 * Math.cos(x * 0.13 - y * 0.41);
  }),
  noise: Array.from({ length: frameSize }, () => 0)
};

function colorStops(name) {
  const maps = {
    instrument: [[8,18,82],[28,68,180],[35,132,220],[40,150,150],[48,178,96],[232,210,44],[224,80,44],[145,30,54],[245,238,215]],
    hot: [[0,0,0],[80,0,0],[180,20,0],[255,130,0],[255,230,80],[255,255,255]],
    jet: [[0,0,120],[0,80,255],[0,220,255],[220,255,0],[255,110,0],[170,0,0]],
    turbo: [[48,18,59],[40,95,196],[50,185,230],[213,226,62],[255,143,24],[180,20,30]],
    gray: [[0,0,0],[255,255,255]]
  };
  return maps[name] || maps.instrument;
}

function colorbarGradient(name) {
  const stops = colorStops(name);
  const parts = stops.map((color, index) => {
    const pct = (index / Math.max(1, stops.length - 1)) * 100;
    return `rgb(${color[0]},${color[1]},${color[2]}) ${pct.toFixed(1)}%`;
  });
  return `linear-gradient(to top, ${parts.join(',')})`;
}

function updateColorbar(min, max) {
  colorBarEl.style.background = colorbarGradient(colormapSelect.value);
  if (Number.isFinite(min) && Number.isFinite(max)) {
    scaleMinEl.textContent = `${min.toFixed(1)} °C`;
    scaleMidEl.textContent = `${((min + max) / 2).toFixed(1)} °C`;
    scaleMaxEl.textContent = `${max.toFixed(1)} °C`;
  } else {
    scaleMinEl.textContent = '-- °C';
    scaleMidEl.textContent = '-- °C';
    scaleMaxEl.textContent = '-- °C';
  }
}

function interpolateColor(stops, t) {
  t = Math.max(0, Math.min(1, t));
  const scaled = t * (stops.length - 1);
  const i = Math.min(stops.length - 2, Math.floor(scaled));
  const f = scaled - i;
  const a = stops[i];
  const b = stops[i + 1];
  return [
    Math.round(a[0] + (b[0] - a[0]) * f),
    Math.round(a[1] + (b[1] - a[1]) * f),
    Math.round(a[2] + (b[2] - a[2]) * f)
  ];
}

function smoothingPasses() {
  if (smoothingSelect.value === 'off') return 0;
  if (smoothingSelect.value === 'high') return 3;
  return 2;
}

function smoothGrid(data, passes) {
  if (passes <= 0) return Float32Array.from(data);
  let current = Float32Array.from(data);
  let next = new Float32Array(current.length);

  for (let pass = 0; pass < passes; pass++) {
    for (let y = 0; y < srcHeight; y++) {
      for (let x = 0; x < srcWidth; x++) {
        let sum = 0;
        let weightSum = 0;
        for (let dy = -1; dy <= 1; dy++) {
          for (let dx = -1; dx <= 1; dx++) {
            const sx = Math.max(0, Math.min(srcWidth - 1, x + dx));
            const sy = Math.max(0, Math.min(srcHeight - 1, y + dy));
            const weight = dx === 0 && dy === 0 ? 4 : (dx === 0 || dy === 0 ? 2 : 1);
            sum += current[sy * srcWidth + sx] * weight;
            weightSum += weight;
          }
        }
        next[y * srcWidth + x] = sum / weightSum;
      }
    }
    const temp = current;
    current = next;
    next = temp;
  }
  return current;
}

function sampleBilinear(data, gx, gy) {
  const x0 = Math.floor(gx);
  const y0 = Math.floor(gy);
  const x1 = Math.min(srcWidth - 1, x0 + 1);
  const y1 = Math.min(srcHeight - 1, y0 + 1);
  const fx = gx - x0;
  const fy = gy - y0;
  const a = data[y0 * srcWidth + x0];
  const b = data[y0 * srcWidth + x1];
  const c = data[y1 * srcWidth + x0];
  const d = data[y1 * srcWidth + x1];
  const top = a + (b - a) * fx;
  const bottom = c + (d - c) * fx;
  return top + (bottom - top) * fy;
}

function rotatedGaussian(x, y, cx, cy, sx, sy, angle, amp) {
  const dx = x - cx;
  const dy = y - cy;
  const ca = Math.cos(angle);
  const sa = Math.sin(angle);
  const rx = ca * dx + sa * dy;
  const ry = -sa * dx + ca * dy;
  return amp * Math.exp(-0.5 * ((rx * rx) / (sx * sx) + (ry * ry) / (sy * sy)));
}

function makeSimulatedFrame() {
  const t = simState.frame * 0.035;
  const cx = 10.5 + 7.2 * Math.sin(t * 0.72) + 1.5 * Math.sin(t * 1.7);
  const cy = 11.4 + 4.1 * Math.cos(t * 0.57);
  const warmCx = cx + 3.2 * Math.cos(t * 0.33);
  const warmCy = cy + 1.8 * Math.sin(t * 0.41);
  const angle = 0.55 * Math.sin(t * 0.42);
  const out = new Array(frameSize);

  for (let y = 0; y < srcHeight; y++) {
    for (let x = 0; x < srcWidth; x++) {
      const i = y * srcWidth + x;
      simState.noise[i] = simState.noise[i] * 0.82 + (Math.random() - 0.5) * 0.045;

      const background = 22.6
        + 0.35 * Math.sin((x + simState.frame * 0.03) * 0.22)
        + 0.28 * Math.cos((y - simState.frame * 0.02) * 0.31)
        + simState.fixedOffsets[i]
        + simState.noise[i];

      const main = rotatedGaussian(x, y, cx, cy, 4.6, 3.0, angle, 10.1);
      const core = rotatedGaussian(x, y, cx - 1.0, cy - 0.4, 2.0, 1.35, angle, 3.7);
      const shoulder = rotatedGaussian(x, y, warmCx, warmCy, 5.1, 2.2, -0.38, 2.4);
      const notch = rotatedGaussian(x, y, cx + 1.7, cy + 0.7, 1.7, 1.2, angle, -1.25);
      const trail = rotatedGaussian(x, y, cx - 5.0, cy + 0.9, 6.8, 1.6, 0.12, 1.5);
      out[i] = background + main + core + shoulder + notch + trail;
    }
  }

  if (simState.previous) {
    for (let i = 0; i < frameSize; i++) {
      out[i] = simState.previous[i] * 0.76 + out[i] * 0.24;
    }
  }
  simState.previous = out.slice();
  simState.frame += 1;
  return out.map(v => Number(v.toFixed(2)));
}

function drawThermalFrame(data) {
  if (!Array.isArray(data) || data.length !== frameSize) return;

  lastData = data;
  const displayData = smoothGrid(data, smoothingPasses());
  let min = Infinity;
  let max = -Infinity;
  let maxIndex = 0;

  for (let i = 0; i < displayData.length; i++) {
    const value = displayData[i];
    if (value < min) min = value;
    if (value > max) {
      max = value;
      maxIndex = i;
    }
  }

  const range = Math.max(0.1, max - min);
  const stops = colorStops(colormapSelect.value);
  const imageData = ctx.createImageData(canvas.width, canvas.height);
  const pixels = imageData.data;
  const xScale = (srcWidth - 1) / Math.max(1, canvas.width - 1);
  const yScale = (srcHeight - 1) / Math.max(1, canvas.height - 1);

  for (let y = 0; y < canvas.height; y++) {
    const gy = y * yScale;
    for (let x = 0; x < canvas.width; x++) {
      const gx = x * xScale;
      const temp = sampleBilinear(displayData, gx, gy);
      const [r, g, b] = interpolateColor(stops, (temp - min) / range);
      const offset = (y * canvas.width + x) * 4;
      pixels[offset] = r;
      pixels[offset + 1] = g;
      pixels[offset + 2] = b;
      pixels[offset + 3] = 255;
    }
  }
  ctx.putImageData(imageData, 0, 0);

  const hotX = maxIndex % srcWidth;
  const hotY = Math.floor(maxIndex / srcWidth);
  const markerX = (hotX / (srcWidth - 1)) * canvas.width;
  const markerY = (hotY / (srcHeight - 1)) * canvas.height;
  const markerSize = Math.max(18, Math.min(canvas.width, canvas.height) * 0.06);
  ctx.save();
  ctx.strokeStyle = 'rgba(255,255,255,0.82)';
  ctx.lineWidth = 2;
  ctx.strokeRect(markerX - markerSize / 2, markerY - markerSize / 2, markerSize, markerSize);
  ctx.beginPath();
  ctx.moveTo(markerX - markerSize * 0.7, markerY);
  ctx.lineTo(markerX + markerSize * 0.7, markerY);
  ctx.moveTo(markerX, markerY - markerSize * 0.7);
  ctx.lineTo(markerX, markerY + markerSize * 0.7);
  ctx.stroke();
  ctx.restore();

  minTempEl.textContent = min.toFixed(1);
  maxTempEl.textContent = max.toFixed(1);
  updateColorbar(min, max);
  frameNoEl.textContent = String(frameCount);
  recordCountEl.textContent = String(records.length);
}

function recordFrame(data) {
  records.push({
    time: new Date().toISOString(),
    frame: frameCount,
    source: activeMode === 'detect' ? 'hardware' : 'import',
    data: data.slice()
  });
  if (records.length > 1200) records.shift();
}

function updateFps() {
  const now = performance.now();
  if (now - fpsTime >= 1000) {
    fpsEl.textContent = String(fpsFrameCounter);
    fpsFrameCounter = 0;
    fpsTime = now;
  }
}

async function fetchHardwareFrame() {
  const response = await fetch('/thermal-data', { cache: 'no-store' });
  if (!response.ok) throw new Error(`HTTP ${response.status}`);
  return await response.json();
}

async function nextFrame() {
  if (!streaming) return;
  try {
    const data = activeMode === 'detect' ? await fetchHardwareFrame() : makeSimulatedFrame();
    frameCount += 1;
    fpsFrameCounter += 1;
    recordFrame(data);
    drawThermalFrame(data);
    updateFps();
    timerId = window.setTimeout(nextFrame, activeMode === 'detect' ? 100 : 80);
  } catch (err) {
    console.error(err);
    stopStream('传感器未连接或数据读取失败');
  }
}

function prepareRun(mode) {
  streaming = true;
  activeMode = mode;
  frameCount = 0;
  fpsFrameCounter = 0;
  fpsTime = performance.now();
  importButton.disabled = true;
  detectButton.disabled = true;
  stopButton.disabled = false;
  sourceLabel.textContent = mode === 'detect' ? '真实检测中' : '导入数据中';
  nextFrame();
}

function startImport() {
  if (streaming) return;
  prepareRun('import');
}

function startDetect() {
  if (streaming) return;
  prepareRun('detect');
}

function stopStream(label = '已停止') {
  streaming = false;
  activeMode = 'idle';
  if (timerId) {
    window.clearTimeout(timerId);
    timerId = null;
  }
  importButton.disabled = false;
  detectButton.disabled = false;
  stopButton.disabled = true;
  sourceLabel.textContent = label;
}

function downloadCsv() {
  if (!records.length) {
    sourceLabel.textContent = '暂无可下载数据';
    return;
  }
  const header = ['timestamp', 'frame', 'source'];
  for (let i = 0; i < frameSize; i++) header.push(`p${String(i).padStart(3, '0')}`);
  const lines = [header.join(',')];
  for (const row of records) {
    lines.push([row.time, row.frame, row.source, ...row.data.map(v => Number(v).toFixed(2))].join(','));
  }
  const blob = new Blob([lines.join('\n')], { type: 'text/csv;charset=utf-8' });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = `mlx90640_${Date.now()}.csv`;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

function clearRecords() {
  records = [];
  frameCount = 0;
  fpsFrameCounter = 0;
  fpsEl.textContent = '0';
  frameNoEl.textContent = '0';
  recordCountEl.textContent = '0';
  sourceLabel.textContent = streaming ? (activeMode === 'detect' ? '真实检测中' : '导入数据中') : '记录已清空';
}

async function toggleFullscreen() {
  const target = document.querySelector('.app') || document.documentElement;
  try {
    if (!document.fullscreenElement) {
      await target.requestFullscreen();
    } else {
      await document.exitFullscreen();
    }
  } catch (err) {
    console.warn(err);
  }
}

function updateFullscreenButton() {
  fullscreenButton.textContent = document.fullscreenElement ? '退出全屏' : '全屏';
}

importButton.addEventListener('click', startImport);
detectButton.addEventListener('click', startDetect);
stopButton.addEventListener('click', () => stopStream());
downloadButton.addEventListener('click', downloadCsv);
clearButton.addEventListener('click', clearRecords);
fullscreenButton.addEventListener('click', toggleFullscreen);
document.addEventListener('fullscreenchange', updateFullscreenButton);
colormapSelect.addEventListener('change', () => {
  updateColorbar(Number(minTempEl.textContent), Number(maxTempEl.textContent));
  if (lastData) drawThermalFrame(lastData);
});
smoothingSelect.addEventListener('change', () => {
  if (lastData) drawThermalFrame(lastData);
});

updateColorbar(NaN, NaN);
