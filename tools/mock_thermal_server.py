#!/usr/bin/env python3
"""Local MLX90640 mock data server for presentation/demo use.

It serves a small thermal-imaging webpage and a /thermal-data endpoint that
returns 768 floating-point temperatures in the same 32x24 shape used by
MLX90640.
"""

from __future__ import annotations

import argparse
import csv
import io
import json
import math
import random
import socket
import threading
import time
import urllib.error
import urllib.request
import webbrowser
from datetime import datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from typing import List
from urllib.parse import parse_qs, urlparse


WIDTH = 32
HEIGHT = 24
FRAME_SIZE = WIDTH * HEIGHT
DEFAULT_DEVICE_BASE_URL = "http://192.168.4.1"


HTML_PAGE = r"""<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>MLX90640 热成像相机</title>
  <style>
    html, body {
      margin: 0;
      width: 100%;
      height: 100%;
      overflow: hidden;
      background: #050505;
      color: #f5f5f5;
      font-family: Arial, "Microsoft YaHei", sans-serif;
    }

    .app {
      height: 100vh;
      display: grid;
      grid-template-rows: auto 1fr auto;
      background: #050505;
    }

    .toolbar {
      display: flex;
      align-items: center;
      gap: 10px;
      padding: 10px 14px;
      background: #101010;
      border-bottom: 1px solid #262626;
      flex-wrap: wrap;
    }

    .title {
      font-size: 17px;
      font-weight: 700;
      margin-right: 10px;
      white-space: nowrap;
    }

    button, select, input {
      height: 34px;
      border: 1px solid #2e6f3f;
      background: #17823a;
      color: #fff;
      border-radius: 4px;
      padding: 0 12px;
      font-size: 14px;
      cursor: pointer;
    }

    input {
      width: 156px;
      background: #171717;
      border-color: #3a3a3a;
      color: #e8e8e8;
      cursor: text;
      box-sizing: border-box;
    }

    button.secondary {
      background: #252525;
      border-color: #4a4a4a;
    }

    button:disabled {
      opacity: 0.45;
      cursor: default;
    }

    .metric {
      color: #d7d7d7;
      font-size: 14px;
      white-space: nowrap;
    }

    .stage {
      min-height: 0;
      display: flex;
      align-items: center;
      justify-content: center;
      gap: 14px;
      padding: 10px;
      box-sizing: border-box;
      background: #000;
    }

    canvas {
      width: min(calc(100vw - 96px), calc((100vh - 112px) * 1.3333));
      height: min(calc(100vh - 112px), 75vw);
      background: #000;
      image-rendering: auto;
      display: block;
    }

    .color-scale {
      height: min(calc(100vh - 132px), 75vw);
      min-height: 260px;
      display: grid;
      grid-template-rows: auto 1fr auto auto;
      align-items: center;
      gap: 8px;
      color: #f2f2f2;
      font-size: 13px;
      text-align: center;
    }

    .color-bar {
      width: 26px;
      height: 100%;
      border: 1px solid #444;
      background: linear-gradient(to top, rgb(8,18,82), rgb(28,68,180), rgb(35,132,220), rgb(48,178,96), rgb(232,210,44), rgb(224,80,44), rgb(245,238,215));
    }

    .footer {
      display: flex;
      justify-content: center;
      gap: 32px;
      padding: 9px 12px;
      background: #101010;
      border-top: 1px solid #262626;
      color: #dcdcdc;
      font-size: 14px;
    }

    @media (max-width: 720px) {
      .toolbar { gap: 7px; padding: 8px; }
      .title { width: 100%; }
      button, select, input { height: 32px; padding: 0 9px; }
      input { width: 145px; }
      .footer { gap: 14px; font-size: 13px; flex-wrap: wrap; }
      .stage { gap: 8px; padding: 8px; }
      canvas { width: calc(100vw - 70px); }
      .color-scale { min-height: 220px; font-size: 12px; }
      .color-bar { width: 22px; }
    }
  </style>
</head>
<body>
  <main class="app">
    <section class="toolbar">
      <div class="title">MLX90640 热成像相机</div>
      <button id="importDataButton">导入数据</button>
      <button id="detectButton">开始检测</button>
      <button id="stopButton" class="secondary" disabled>停止采集</button>
      <button id="downloadButton" class="secondary">下载CSV数据</button>
      <button id="clearButton" class="secondary">清空记录</button>
      <button id="fullscreenButton" class="secondary">全屏</button>
      <input id="deviceAddress" value="http://192.168.4.1" title="ESP32设备地址">
      <select id="colormap">
        <option value="instrument" selected>热像仪参考</option>
        <option value="jet">彩虹图</option>
        <option value="hot">热力图</option>
        <option value="turbo">涡轮图</option>
        <option value="gray">灰度图</option>
      </select>
      <select id="smoothingMode">
        <option value="medium" selected>平滑中</option>
        <option value="high">强平滑</option>
        <option value="off">关闭平滑</option>
      </select>
      <span class="metric">FPS: <span id="fps">0</span></span>
      <span class="metric">采集状态: <span id="sourceLabel">等待采集</span></span>
    </section>

    <section class="stage">
      <canvas id="thermalCanvas" width="640" height="480"></canvas>
      <aside class="color-scale">
        <span id="scaleMax">-- °C</span>
        <div id="colorBar" class="color-bar"></div>
        <span id="scaleMid">-- °C</span>
        <span id="scaleMin">-- °C</span>
      </aside>
    </section>

    <section class="footer">
      <span>最低温度: <span id="minTemp">--</span> °C</span>
      <span>最高温度: <span id="maxTemp">--</span> °C</span>
      <span>帧序号: <span id="frameNo">0</span></span>
      <span>已记录: <span id="recordCount">0</span> 帧</span>
    </section>
  </main>

  <script>
    const srcWidth = 32;
    const srcHeight = 24;
    const canvas = document.getElementById('thermalCanvas');
    const ctx = canvas.getContext('2d');
    ctx.imageSmoothingEnabled = true;

    const importButton = document.getElementById('importDataButton');
    const detectButton = document.getElementById('detectButton');
    const stopButton = document.getElementById('stopButton');
    const downloadButton = document.getElementById('downloadButton');
    const clearButton = document.getElementById('clearButton');
    const fullscreenButton = document.getElementById('fullscreenButton');
    const deviceAddressInput = document.getElementById('deviceAddress');
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

    function drawThermalFrame(data) {
      if (!Array.isArray(data) || data.length !== srcWidth * srcHeight) return;

      lastData = data;
      const displayData = smoothGrid(data, smoothingPasses());
      let min = Infinity;
      let max = -Infinity;
      for (const value of displayData) {
        if (value < min) min = value;
        if (value > max) max = value;
      }
      const range = Math.max(0.1, max - min);
      const stops = colorStops(colormapSelect.value);
      const imageData = ctx.createImageData(canvas.width, canvas.height);
      const pixels = imageData.data;
      const xScale = (srcWidth - 1) / Math.max(1, canvas.width - 1);
      const yScale = (srcHeight - 1) / Math.max(1, canvas.height - 1);

      ctx.clearRect(0, 0, canvas.width, canvas.height);
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

      minTempEl.textContent = min.toFixed(1);
      maxTempEl.textContent = max.toFixed(1);
      updateColorbar(min, max);
      frameNoEl.textContent = String(frameCount);
      recordCountEl.textContent = String(frameCount);
    }

    async function fetchFrame() {
      if (!streaming) return;
      try {
        const endpoint = activeMode === 'detect'
          ? `/hardware-data?base=${encodeURIComponent(deviceAddressInput.value.trim())}`
          : '/thermal-data';
        const response = await fetch(endpoint, { cache: 'no-store' });
        if (!response.ok) throw new Error(`HTTP ${response.status}`);
        const data = await response.json();
        frameCount += 1;
        fpsFrameCounter += 1;
        drawThermalFrame(data);

        const now = performance.now();
        if (now - fpsTime >= 1000) {
          fpsEl.textContent = String(fpsFrameCounter);
          fpsFrameCounter = 0;
          fpsTime = now;
        }
      } catch (err) {
        sourceLabel.textContent = activeMode === 'detect' ? '设备未连接' : '数据读取失败';
        streaming = false;
        activeMode = 'idle';
        importButton.disabled = false;
        detectButton.disabled = false;
        stopButton.disabled = true;
        console.error(err);
      }
      if (streaming) window.setTimeout(fetchFrame, 80);
    }

    async function startImport() {
      await fetch('/control/start-import', { cache: 'no-store' });
      streaming = true;
      activeMode = 'import';
      frameCount = 0;
      fpsFrameCounter = 0;
      fpsTime = performance.now();
      recordCountEl.textContent = '0';
      importButton.disabled = true;
      detectButton.disabled = true;
      stopButton.disabled = false;
      sourceLabel.textContent = '导入采集中';
      fetchFrame();
    }

    async function startDetect() {
      await fetch('/control/start-detect', { cache: 'no-store' });
      streaming = true;
      activeMode = 'detect';
      frameCount = 0;
      fpsFrameCounter = 0;
      fpsTime = performance.now();
      recordCountEl.textContent = '0';
      importButton.disabled = true;
      detectButton.disabled = true;
      stopButton.disabled = false;
      sourceLabel.textContent = '实时检测中';
      fetchFrame();
    }

    async function stopImport() {
      streaming = false;
      activeMode = 'idle';
      await fetch('/control/stop', { cache: 'no-store' });
      importButton.disabled = false;
      detectButton.disabled = false;
      stopButton.disabled = true;
      sourceLabel.textContent = '已停止';
    }

    function downloadCsv() {
      window.location.href = `/download/csv?t=${Date.now()}`;
    }

    async function clearRecords() {
      await fetch('/control/clear', { cache: 'no-store' });
      frameCount = 0;
      fpsFrameCounter = 0;
      fpsEl.textContent = '0';
      frameNoEl.textContent = '0';
      recordCountEl.textContent = '0';
      sourceLabel.textContent = streaming ? (activeMode === 'detect' ? '实时检测中' : '导入采集中') : '记录已清空';
    }

    async function toggleFullscreen() {
      const target = document.querySelector('.app') || document.documentElement;
      if (!document.fullscreenElement) {
        await target.requestFullscreen();
      } else {
        await document.exitFullscreen();
      }
    }

    function updateFullscreenButton() {
      fullscreenButton.textContent = document.fullscreenElement ? '退出全屏' : '全屏';
    }

    importButton.addEventListener('click', startImport);
    detectButton.addEventListener('click', startDetect);
    stopButton.addEventListener('click', stopImport);
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
  </script>
</body>
</html>
"""


class MockThermalState:
    def __init__(self) -> None:
        self.running = False
        self.mode = "idle"
        self.frame_index = 0
        self.previous_frame: List[float] | None = None
        self.last_frame_time = 0.0
        self.records: list[dict] = []
        self.lock = threading.Lock()
        self.rng = random.Random()
        self.fixed_offsets = self._make_fixed_offsets()
        self.noise_state = [0.0 for _ in range(FRAME_SIZE)]

    def start(self, mode: str = "import") -> None:
        with self.lock:
            self.running = True
            self.mode = mode
            self.frame_index = 0
            self.previous_frame = None
            self.last_frame_time = 0.0
            self.records.clear()

    def stop(self) -> None:
        with self.lock:
            self.running = False
            self.mode = "idle"

    def clear_records(self) -> None:
        with self.lock:
            self.records.clear()
            self.frame_index = 0

    def status(self) -> dict:
        with self.lock:
            return {
                "running": self.running,
                "mode": self.mode,
                "frame_index": self.frame_index,
                "record_count": len(self.records),
                "width": WIDTH,
                "height": HEIGHT,
            }

    def record_values(self, values: List[float]) -> List[float]:
        if len(values) != FRAME_SIZE:
            raise ValueError(f"Expected {FRAME_SIZE} points, got {len(values)}")

        rounded = [round(float(value), 1) for value in values]
        frame_min = min(rounded)
        frame_max = max(rounded)
        frame_avg = sum(rounded) / len(rounded)

        with self.lock:
            if self.running:
                self.frame_index += 1
            index = self.frame_index
            if self.running and self.mode == "detect":
                self.records.append(
                    {
                        "timestamp": datetime.now().isoformat(timespec="milliseconds"),
                        "frame_index": index,
                        "min_c": frame_min,
                        "max_c": frame_max,
                        "avg_c": frame_avg,
                        "values": rounded,
                    }
                )
        return rounded

    def csv_bytes(self) -> bytes:
        with self.lock:
            records = list(self.records)

        output = io.StringIO(newline="")
        writer = csv.writer(output)
        writer.writerow(
            ["timestamp", "frame_index", "min_c", "max_c", "avg_c"]
            + [f"p{i:03d}" for i in range(FRAME_SIZE)]
        )
        for record in records:
            writer.writerow(
                [
                    record["timestamp"],
                    record["frame_index"],
                    f"{record['min_c']:.1f}",
                    f"{record['max_c']:.1f}",
                    f"{record['avg_c']:.2f}",
                ]
                + [f"{value:.1f}" for value in record["values"]]
            )
        return output.getvalue().encode("utf-8-sig")

    @staticmethod
    def _soft_box(x: float, y: float, cx: float, cy: float, half_w: float, half_h: float, edge: float) -> float:
        dx = max(abs(x - cx) - half_w, 0.0)
        dy = max(abs(y - cy) - half_h, 0.0)
        dist = math.hypot(dx, dy)
        if dist <= 0.0:
            return 1.0
        return math.exp(-(dist * dist) / (2.0 * edge * edge))

    @staticmethod
    def _gaussian(x: float, y: float, cx: float, cy: float, sx: float, sy: float) -> float:
        return math.exp(-(((x - cx) ** 2) / (2.0 * sx * sx) + ((y - cy) ** 2) / (2.0 * sy * sy)))

    @staticmethod
    def _rotated_gaussian(
        x: float,
        y: float,
        cx: float,
        cy: float,
        sx: float,
        sy: float,
        theta: float,
    ) -> float:
        dx = x - cx
        dy = y - cy
        cos_t = math.cos(theta)
        sin_t = math.sin(theta)
        xr = dx * cos_t + dy * sin_t
        yr = -dx * sin_t + dy * cos_t
        return math.exp(-((xr * xr) / (2.0 * sx * sx) + (yr * yr) / (2.0 * sy * sy)))

    def _make_fixed_offsets(self) -> List[float]:
        row_bias = [self.rng.gauss(0.0, 0.035) for _ in range(HEIGHT)]
        col_bias = [self.rng.gauss(0.0, 0.025) for _ in range(WIDTH)]
        offsets: List[float] = []
        for y in range(HEIGHT):
            for x in range(WIDTH):
                checker = 0.018 if (x + y) % 2 == 0 else -0.018
                offsets.append(row_bias[y] + col_bias[x] + checker + self.rng.gauss(0.0, 0.028))
        return offsets

    def frame(self) -> List[float]:
        with self.lock:
            if self.running:
                self.frame_index += 1
            index = self.frame_index
            should_record = self.running and self.mode == "import"
            previous = self.previous_frame
            last_frame_time = self.last_frame_time

        t = time.time()
        hot_x = 15.8 + 6.2 * math.sin(t * 0.082) + 0.65 * math.sin(t * 0.29 + 0.7)
        hot_y = 11.6 + 3.6 * math.sin(t * 0.067 + 1.1) + 0.45 * math.cos(t * 0.24)
        theta = 0.28 * math.sin(t * 0.046 + 0.8) + 0.14 * math.sin(t * 0.13)
        heat_gain = 1.0 + 0.055 * math.sin(t * 0.055) + 0.025 * math.sin(t * 0.31 + 2.0)
        ambient_drift = 0.16 * math.sin(t * 0.020) + 0.06 * math.sin(t * 0.073 + 1.5)

        target: List[float] = []
        for y in range(HEIGHT):
            for x in range(WIDTH):
                xf = x + 0.5
                yf = y + 0.5
                idx = y * WIDTH + x

                ambient = 23.15 + ambient_drift + 0.017 * x - 0.015 * y
                lens_gradient = 0.18 * math.sin((xf + t * 0.035) / 5.3) + 0.15 * math.cos((yf - t * 0.028) / 4.2)
                cool_corner = -0.70 * self._gaussian(xf, yf, 27.5, 18.2, 5.8, 4.6)
                row_readout = 0.045 * math.sin(y * 1.73 + t * 0.52)
                col_readout = 0.030 * math.sin(x * 2.11 - t * 0.37)

                asymmetry = 1.0 + 0.055 * math.sin(0.75 * xf + 0.42 * yf + t * 0.11)
                hot_body = 7.4 * heat_gain * self._rotated_gaussian(xf, yf, hot_x, hot_y, 4.8, 3.0, theta) * asymmetry
                hot_lobe = 2.7 * self._rotated_gaussian(xf, yf, hot_x - 2.2, hot_y + 0.8, 2.2, 1.55, theta + 0.45)
                core = 2.4 * self._rotated_gaussian(xf, yf, hot_x - 1.25, hot_y - 0.4, 1.45, 1.05, theta - 0.2)
                soft_edge = 1.4 * self._rotated_gaussian(xf, yf, hot_x + 2.4, hot_y + 1.3, 3.9, 2.6, theta + 0.25)
                notch = -1.05 * self._rotated_gaussian(xf, yf, hot_x + 1.3, hot_y - 1.2, 1.4, 1.0, theta + 0.7)
                trailing_warmth = 0.95 * self._rotated_gaussian(
                    xf,
                    yf,
                    hot_x + 4.1 + 0.4 * math.sin(t * 0.15),
                    hot_y + 1.7,
                    4.1,
                    2.6,
                    theta + 0.35,
                )

                noise = 0.82 * self.noise_state[idx] + self.rng.gauss(0.0, 0.042)
                if self.rng.random() < 0.004:
                    noise += self.rng.uniform(-0.20, 0.20)
                self.noise_state[idx] = max(-0.22, min(0.22, noise))

                target.append(
                    ambient
                    + lens_gradient
                    + cool_corner
                    + row_readout
                    + col_readout
                    + self.fixed_offsets[idx]
                    + hot_body
                    + hot_lobe
                    + core
                    + soft_edge
                    + notch
                    + trailing_warmth
                    + self.noise_state[idx]
                )

        if previous is None or len(previous) != FRAME_SIZE:
            values = target
        else:
            elapsed = max(0.001, t - last_frame_time) if last_frame_time else 0.08
            alpha = min(0.42, max(0.10, elapsed * 2.4))
            max_delta = min(0.17, max(0.035, elapsed * 1.15))
            values = []
            for old, new in zip(previous, target):
                proposed = old + (new - old) * alpha
                delta = max(-max_delta, min(max_delta, proposed - old))
                values.append(old + delta)

        rounded = [round(value, 1) for value in values]
        if should_record:
            frame_min = min(rounded)
            frame_max = max(rounded)
            frame_avg = sum(rounded) / len(rounded)
            record = {
                "timestamp": datetime.now().isoformat(timespec="milliseconds"),
                "frame_index": index,
                "min_c": frame_min,
                "max_c": frame_max,
                "avg_c": frame_avg,
                "values": rounded,
            }
        else:
            record = None

        with self.lock:
            self.previous_frame = values
            self.last_frame_time = t
            if record is not None:
                self.records.append(record)
        return rounded


STATE = MockThermalState()


def normalize_device_data_url(base: str | None) -> str:
    base = (base or DEFAULT_DEVICE_BASE_URL).strip()
    if not base:
        base = DEFAULT_DEVICE_BASE_URL
    if not base.startswith(("http://", "https://")):
        base = "http://" + base

    parsed = urlparse(base)
    if parsed.path and parsed.path != "/":
        return base
    return base.rstrip("/") + "/thermal-data"


def read_hardware_frame(base: str | None) -> List[float]:
    url = normalize_device_data_url(base)
    request = urllib.request.Request(url, headers={"Cache-Control": "no-store"})
    with urllib.request.urlopen(request, timeout=1.5) as response:
        payload = response.read().decode("utf-8")

    data = json.loads(payload)
    if not isinstance(data, list) or len(data) != FRAME_SIZE:
        raise ValueError(f"Invalid thermal frame from device: expected {FRAME_SIZE} values")
    return [float(value) for value in data]


class MockThermalHandler(BaseHTTPRequestHandler):
    server_version = "MockThermalHTTP/1.0"

    def log_message(self, fmt: str, *args: object) -> None:
        print("[%s] %s" % (self.log_date_time_string(), fmt % args))

    def _send(self, status: int, body: bytes, content_type: str) -> None:
        self.send_response(status)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store, no-cache, must-revalidate, max-age=0")
        self.send_header("Access-Control-Allow-Origin", "*")
        self.end_headers()
        self.wfile.write(body)

    def _send_json(self, value: object) -> None:
        body = json.dumps(value, ensure_ascii=False, separators=(",", ":")).encode("utf-8")
        self._send(200, body, "application/json; charset=utf-8")

    def _send_download(self, body: bytes, filename: str, content_type: str) -> None:
        self.send_response(200)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.send_header("Content-Disposition", f'attachment; filename="{filename}"')
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self) -> None:
        parsed_url = urlparse(self.path)
        path = parsed_url.path
        query = parse_qs(parsed_url.query)
        if path in ("/", "/index.html"):
            self._send(200, HTML_PAGE.encode("utf-8"), "text/html; charset=utf-8")
        elif path == "/thermal-data":
            self._send_json(STATE.frame())
        elif path == "/hardware-data":
            try:
                base = query.get("base", [DEFAULT_DEVICE_BASE_URL])[0]
                values = read_hardware_frame(base)
                self._send_json(STATE.record_values(values))
            except (OSError, urllib.error.URLError, TimeoutError, ValueError, json.JSONDecodeError) as exc:
                self._send(502, json.dumps({"ok": False, "error": str(exc)}, ensure_ascii=False).encode("utf-8"), "application/json; charset=utf-8")
        elif path in ("/control/start", "/control/start-import"):
            STATE.start("import")
            self._send_json({"ok": True, "running": True, "mode": "import"})
        elif path == "/control/start-detect":
            STATE.start("detect")
            self._send_json({"ok": True, "running": True, "mode": "detect"})
        elif path == "/control/stop":
            STATE.stop()
            self._send_json({"ok": True, "running": False})
        elif path == "/control/clear":
            STATE.clear_records()
            self._send_json({"ok": True, "record_count": 0})
        elif path == "/download/csv":
            stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            self._send_download(
                STATE.csv_bytes(),
                f"mlx90640_thermal_frames_{stamp}.csv",
                "text/csv; charset=utf-8",
            )
        elif path == "/status":
            self._send_json(STATE.status())
        else:
            self._send(404, b"Not found", "text/plain; charset=utf-8")


def find_port(host: str, requested_port: int) -> int:
    for port in range(requested_port, requested_port + 50):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
            sock.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
            try:
                sock.bind((host, port))
            except OSError:
                continue
            return port
    raise RuntimeError(f"No free port found from {requested_port} to {requested_port + 49}")


def local_ip_address() -> str:
    with socket.socket(socket.AF_INET, socket.SOCK_DGRAM) as sock:
        try:
            sock.connect(("8.8.8.8", 80))
            return sock.getsockname()[0]
        except OSError:
            return "127.0.0.1"


def main() -> None:
    parser = argparse.ArgumentParser(description="Run a local mock MLX90640 thermal-data webpage.")
    parser.add_argument("--host", default="127.0.0.1", help="Bind host, default: 127.0.0.1")
    parser.add_argument("--port", type=int, default=8000, help="Preferred port, default: 8000")
    parser.add_argument("--no-browser", action="store_true", help="Do not open a browser automatically")
    args = parser.parse_args()

    port = find_port(args.host, args.port)
    server = ThreadingHTTPServer((args.host, port), MockThermalHandler)
    browser_host = "127.0.0.1" if args.host == "0.0.0.0" else args.host
    url = f"http://{browser_host}:{port}/"
    print(f"Mock thermal server running at {url}")
    if args.host == "0.0.0.0":
        print(f"LAN URL for another device: http://{local_ip_address()}:{port}/")
    print("Open the page and click the '导入数据' button to start simulated frames.")

    if not args.no_browser:
        webbrowser.open(url)

    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nStopping mock thermal server.")
    finally:
        server.server_close()


if __name__ == "__main__":
    main()
