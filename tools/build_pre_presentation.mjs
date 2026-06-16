import fs from "node:fs/promises";
import path from "node:path";
import { pathToFileURL } from "node:url";

if (process.env.USERPROFILE) {
  process.env.HOME = process.env.USERPROFILE;
}

const ROOT = "H:\\ircam_pio_c3";
const WORKSPACE = path.join(ROOT, "outputs", "manual-20260605-ircam", "presentations", "ircam-pre");
const PREVIEW_DIR = path.join(WORKSPACE, "preview");
const LAYOUT_DIR = path.join(WORKSPACE, "layout", "final");
const OUT_DIR = path.join(ROOT, "deliverables");
const FINAL_PPTX = path.join(OUT_DIR, "热成像相机项目完整汇报_PRE正式版.pptx");
const FINAL_ASCII = path.join(OUT_DIR, "ircam_pre_presentation.pptx");
const ASSET_DIR = path.join(ROOT, "outputs", "manual-20260605-ircam", "presentations", "ircam-full", "assets");

const skillUtils = path.join(
  process.env.USERPROFILE || "C:\\Users\\追梦的路口",
  ".codex",
  "plugins",
  "cache",
  "openai-primary-runtime",
  "presentations",
  "26.601.10930",
  "skills",
  "presentations",
  "scripts",
  "artifact_tool_utils.mjs",
);

const {
  ensureArtifactToolWorkspace,
  importArtifactTool,
  createSlideContext,
  saveBlobToFile,
  padSlideNumber,
} = await import(pathToFileURL(skillUtils).href);

const C = {
  bg: "#FBFAF7",
  ink: "#202124",
  muted: "#5F6368",
  line: "#D8D3C8",
  blue: "#365F7D",
  blue2: "#2F4F66",
  green: "#5B7F64",
  orange: "#B98244",
  red: "#A33A2F",
  dark: "#263238",
  panel: "#FFFFFF",
  paleBlue: "#EEF3F5",
  paleGreen: "#EEF5EF",
  paleOrange: "#F7EFE3",
  paleRed: "#F8EAE7",
  codeBg: "#263238",
  codeText: "#F3F4F1",
};

const SIZE = { width: 1280, height: 720 };

function line(ctx, fill = "#00000000", width = 0) {
  return ctx.line(fill, width);
}

function addShape(ctx, slide, x, y, w, h, fill, stroke = "#00000000", width = 0) {
  return ctx.addShape(slide, { x, y, w, h, fill, line: line(ctx, stroke, width) });
}

function text(ctx, slide, x, y, w, h, value, opts = {}) {
  return ctx.addText(slide, {
    x, y, w, h,
    text: value,
    fontSize: opts.size ?? 18,
    color: opts.color ?? C.ink,
    bold: opts.bold ?? false,
    align: opts.align ?? "left",
    valign: opts.valign ?? "top",
    typeface: opts.face ?? "Microsoft YaHei",
    fill: opts.fill ?? "#00000000",
    line: opts.line ?? line(ctx),
    insets: opts.insets ?? { left: 0, right: 0, top: 0, bottom: 0 },
  });
}

function slideBase(presentation, ctx, no, kicker, title, subtitle = "") {
  const slide = presentation.slides.add();
  addShape(ctx, slide, 0, 0, ctx.W, ctx.H, C.bg);
  addShape(ctx, slide, 64, 118, 1080, 1.2, C.line);
  text(ctx, slide, 64, 34, 380, 20, kicker, { size: 11, color: C.green, bold: true, valign: "middle" });
  text(ctx, slide, 64, 60, 1060, 50, title, { size: 30, color: C.ink, bold: true, face: "Microsoft YaHei" });
  if (subtitle) text(ctx, slide, 64, 126, 1050, 26, subtitle, { size: 14, color: C.muted });
  text(ctx, slide, 1100, 666, 90, 20, String(no).padStart(2, "0"), { size: 10, color: C.muted, align: "right" });
  return slide;
}

function panel(ctx, slide, x, y, w, h, fill = C.panel, stroke = C.line) {
  addShape(ctx, slide, x, y, w, h, fill, stroke, 1);
}

function bullets(ctx, slide, x, y, w, items, opts = {}) {
  const gap = opts.gap ?? 40;
  const size = opts.size ?? 17;
  const dot = opts.dot ?? C.green;
  const itemH = opts.itemH ?? 32;
  items.forEach((item, i) => {
    const top = y + i * gap;
    addShape(ctx, slide, x, top + 9, 8, 8, dot);
    text(ctx, slide, x + 22, top, w, itemH, item, { size, color: opts.color ?? C.ink });
  });
}

function smallBullets(ctx, slide, x, y, w, items, opts = {}) {
  bullets(ctx, slide, x, y, w, items, { gap: opts.gap ?? 30, size: opts.size ?? 14, itemH: opts.itemH ?? 25, dot: opts.dot ?? C.orange });
}

function card(ctx, slide, x, y, w, h, title, body, fill = C.panel, accent = C.green) {
  panel(ctx, slide, x, y, w, h, fill, C.line);
  addShape(ctx, slide, x + 20, y + 52, w - 40, 1.1, accent);
  text(ctx, slide, x + 22, y + 18, w - 44, 28, title, { size: 18, bold: true, color: C.dark });
  text(ctx, slide, x + 22, y + 66, w - 44, h - 82, body, { size: 14, color: C.muted });
}

function metric(ctx, slide, x, y, w, h, value, label, fill = C.paleBlue) {
  panel(ctx, slide, x, y, w, h, "#FFFFFF", C.line);
  text(ctx, slide, x + 16, y + 14, w - 32, 30, value, { size: 23, bold: true, color: C.dark, valign: "middle" });
  text(ctx, slide, x + 16, y + 50, w - 32, h - 58, label, { size: 12, color: C.muted });
}

function table(ctx, slide, x, y, widths, rows, opts = {}) {
  const rowH = opts.rowH ?? 42;
  rows.forEach((row, r) => {
    let cx = x;
    row.forEach((cell, c) => {
      panel(ctx, slide, cx, y + r * rowH, widths[c], rowH, r === 0 ? "#ECE7DD" : C.panel, r === 0 ? "#CFC7B8" : C.line);
      text(ctx, slide, cx + 10, y + r * rowH + 6, widths[c] - 20, rowH - 12, cell, {
        size: r === 0 ? (opts.headerSize ?? 14) : (opts.size ?? 13),
        bold: r === 0,
        color: r === 0 ? C.dark : C.ink,
        valign: "middle",
      });
      cx += widths[c];
    });
  });
}

function code(ctx, slide, x, y, w, h, value, title = "") {
  panel(ctx, slide, x, y, w, h, C.codeBg, "#3F4A4F");
  if (title) text(ctx, slide, x + 18, y + 14, w - 36, 22, title, { size: 13, bold: true, color: "#D7C6A5", face: "Consolas" });
  text(ctx, slide, x + 18, y + (title ? 44 : 18), w - 36, h - (title ? 56 : 30), value, {
    size: 14,
    color: C.codeText,
    face: "Consolas",
  });
}

function arrow(ctx, slide, x, y, w, color = C.green) {
  addShape(ctx, slide, x, y + 9, w, 2.5, color);
  addShape(ctx, slide, x + w - 8, y + 3, 13, 14, color);
}

async function image(ctx, slide, name, x, y, w, h, fit = "cover") {
  const imagePath = path.join(ASSET_DIR, name);
  try {
    await ctx.addImage(slide, { path: imagePath, x, y, w, h, fit, alt: name });
  } catch {
    panel(ctx, slide, x, y, w, h, C.paleRed, "#F3B6B1");
    text(ctx, slide, x + 18, y + 18, w - 36, h - 36, `图片缺失：${name}`, { size: 16, color: C.red });
  }
}

function section(presentation, ctx, no, title, subtitle, chips = []) {
  const slide = presentation.slides.add();
  addShape(ctx, slide, 0, 0, ctx.W, ctx.H, C.bg);
  addShape(ctx, slide, 64, 106, 1080, 1.2, C.line);
  text(ctx, slide, 72, 70, 180, 26, `第 ${chips[0] ?? ""} 部分`, { size: 17, color: C.green, bold: true });
  text(ctx, slide, 72, 160, 900, 70, title, { size: 40, color: C.ink, bold: true });
  text(ctx, slide, 76, 254, 860, 62, subtitle, { size: 20, color: C.muted });
  chips.slice(1).forEach((chip, i) => metric(ctx, slide, 80 + i * 235, 430, 205, 88, chip.title, chip.body));
  text(ctx, slide, 1100, 666, 90, 20, String(no).padStart(2, "0"), { size: 10, color: C.muted, align: "right" });
  return slide;
}

function threeCards(presentation, ctx, no, kicker, title, subtitle, cards) {
  const slide = slideBase(presentation, ctx, no, kicker, title, subtitle);
  const w = 345;
  cards.forEach((c, i) => card(ctx, slide, 82 + i * 385, 170, w, 390, c.title, c.body, c.fill, c.accent));
  if (cards.footer) text(ctx, slide, 90, 600, 1040, 32, cards.footer, { size: 16, color: C.muted, bold: true });
  return slide;
}

function twoColumns(presentation, ctx, no, kicker, title, subtitle, left, right, footer = "") {
  const slide = slideBase(presentation, ctx, no, kicker, title, subtitle);
  panel(ctx, slide, 70, 160, 520, 430, left.fill ?? C.panel, left.stroke ?? C.line);
  panel(ctx, slide, 650, 160, 520, 430, right.fill ?? C.panel, right.stroke ?? C.line);
  text(ctx, slide, 108, 194, 440, 30, left.title, { size: 23, bold: true, color: C.dark });
  bullets(ctx, slide, 112, 254, 410, left.items, { gap: left.gap ?? 47, size: left.size ?? 17, itemH: left.itemH ?? 36, dot: left.dot ?? C.green });
  text(ctx, slide, 688, 194, 440, 30, right.title, { size: 23, bold: true, color: C.dark });
  bullets(ctx, slide, 692, 254, 410, right.items, { gap: right.gap ?? 47, size: right.size ?? 17, itemH: right.itemH ?? 36, dot: right.dot ?? C.orange });
  if (footer) text(ctx, slide, 92, 620, 1040, 32, footer, { size: 15, color: C.red, bold: true, align: "center" });
  return slide;
}

function node(ctx, slide, x, y, w, h, title, body, fill = C.paleBlue) {
  panel(ctx, slide, x, y, w, h, fill, "#D2CDC2");
  text(ctx, slide, x + 14, y + 14, w - 28, 28, title, { size: 17, bold: true, color: C.dark, align: "center" });
  text(ctx, slide, x + 14, y + 48, w - 28, h - 56, body, { size: 12, color: C.muted, align: "center" });
}

async function buildDeck() {
  await fs.mkdir(OUT_DIR, { recursive: true });
  await fs.mkdir(PREVIEW_DIR, { recursive: true });
  await fs.mkdir(LAYOUT_DIR, { recursive: true });
  await ensureArtifactToolWorkspace(WORKSPACE);
  const artifact = await importArtifactTool(WORKSPACE);
  const { Presentation, PresentationFile } = artifact;
  const presentation = Presentation.create({ slideSize: SIZE });

  let no = 1;
  const ctxFor = () => createSlideContext(artifact, {
    slideSize: SIZE,
    slideNumber: no,
    outputDir: OUT_DIR,
    assetDir: ASSET_DIR,
    workspaceDir: WORKSPACE,
    titleFont: "Microsoft YaHei",
    bodyFont: "Microsoft YaHei",
    monoFont: "Consolas",
  });

  {
    const ctx = ctxFor();
    const slide = presentation.slides.add();
    addShape(ctx, slide, 0, 0, ctx.W, ctx.H, C.bg);
    addShape(ctx, slide, 62, 92, 1060, 1.2, C.line);
    text(ctx, slide, 64, 54, 360, 28, "微机接口与技术大作业", { size: 18, color: C.green, bold: true });
    text(ctx, slide, 64, 128, 640, 122, "基于 ESP32-C3 与 MLX90640 的热成像相机", { size: 42, color: C.ink, bold: true });
    text(ctx, slide, 66, 286, 560, 76, "围绕器件选择、硬件接线、环境配置、烧录流程和网页验证展开，形成从系统设计到阶段性结果的完整汇报。", { size: 18, color: C.muted });
    panel(ctx, slide, 715, 128, 380, 145, "#FFFFFF", C.line);
    text(ctx, slide, 742, 160, 320, 30, "项目当前状态", { size: 23, bold: true, color: C.dark });
    text(ctx, slide, 742, 202, 315, 44, "软件与网页链路已经跑通，硬件端等待 MLX90640 数据稳定进入。", { size: 15, color: C.muted });
    panel(ctx, slide, 66, 410, 1030, 165, "#FFFFFF", C.line);
    text(ctx, slide, 96, 440, 250, 30, "阶段性核心结论", { size: 22, bold: true, color: C.dark });
    bullets(ctx, slide, 365, 438, 660, [
      "软件链路已跑通：PlatformIO 环境、固件编译、固件烧录、SPIFFS 网页资源烧录均已完成。",
      "网页链路已验证：ESP32-C3 能创建热点，浏览器可访问 192.168.4.1 并打开控制界面。",
      "硬件链路待收尾：MLX90640 为裸传感器，仍需补齐 I2C 上拉、电源滤波与接触可靠性检查。",
      "后续只要传感器数据连通，就可以补上真实热图作为最终展示结果。",
    ], { size: 14, gap: 32, itemH: 26 });
    metric(ctx, slide, 90, 600, 160, 70, "COM5", "CH343 串口已确认");
    metric(ctx, slide, 275, 600, 160, 70, "IO04/05", "I2C 引脚已同步");
    metric(ctx, slide, 460, 600, 160, 70, "37页", "详细汇报内容");
    text(ctx, slide, 960, 666, 120, 20, "2026.06", { size: 11, color: C.muted, align: "right" });
  }
  no++;

  threeCards(presentation, ctxFor(), no++, "PROJECT MOTIVATION", "为什么选择热成像相机作为微机接口项目", "这个题目能同时覆盖传感器接口、嵌入式软件、网页服务和现场调试，比较适合做成可演示的大作业。",
    [
      { title: "课程匹配度高", body: "项目需要完成 GPIO/I2C 总线、外设供电、串口下载、存储分区、网络服务等内容，能体现“微机接口”的核心训练。", fill: C.paleGreen, accent: C.green },
      { title: "展示效果直观", body: "热成像不是单纯打印数据，而是把 32×24 温度矩阵转成网页热图。手、热水杯等热源进入视野后，画面会产生明显颜色变化。", fill: C.paleBlue, accent: C.blue },
      { title: "工程问题真实", body: "裸 MLX90640 对接线、上拉、电源稳定性更敏感，调试过程能说明为什么硬件细节会影响软件结果。", fill: C.paleOrange, accent: C.orange },
    ]);

  threeCards(presentation, ctxFor(), no++, "PROJECT SCOPE", "项目目标拆解：不是只让传感器亮，而是形成完整系统", "最终目标是让 ESP32-C3 独立采集热数据、独立提供网页，并能被手机/电脑访问。",
    [
      { title: "硬件目标", body: "ESP32-C3 与 MLX90640 通过 I2C 连接；传感器使用 3.3V；SDA/SCL 加上拉；VDD/GND 加滤波电容；所有裸脚避免短路。", fill: C.panel, accent: C.green },
      { title: "软件目标", body: "完成 I2C 初始化、MLX90640 参数读取、帧数据采集、温度换算、HTTP 接口返回、SPIFFS 静态网页加载。", fill: C.panel, accent: C.blue },
      { title: "展示目标", body: "手机或电脑连接 ESP32 热点，在浏览器打开热图界面；能解释当前黑屏/--°C 是传感器数据未连通，不是网页失败。", fill: C.panel, accent: C.orange },
    ]);

  {
    const ctx = ctxFor();
    const slide = slideBase(presentation, ctx, no++, "SYSTEM ARCHITECTURE", "系统总体结构：传感器数据到网页热图的闭环", "每一层都有清晰职责：采集、处理、服务、展示。");
    node(ctx, slide, 70, 230, 190, 130, "MLX90640", "32×24 红外阵列\nI2C 地址 0x33\n输出原始帧数据", C.paleOrange);
    arrow(ctx, slide, 280, 285, 75);
    node(ctx, slide, 370, 230, 190, 130, "ESP32-C3", "I2C 采集\n温度换算\nWi-Fi SoftAP", C.paleGreen);
    arrow(ctx, slide, 580, 285, 75);
    node(ctx, slide, 670, 230, 190, 130, "SPIFFS", "存放 HTML/JS\n网页资源随固件\n写入 flash 分区", C.paleBlue);
    arrow(ctx, slide, 880, 285, 75);
    node(ctx, slide, 970, 230, 190, 130, "浏览器", "访问 192.168.4.1\n请求 /thermal-data\n绘制热图", C.panel);
    panel(ctx, slide, 100, 455, 1010, 120, C.panel, C.line);
    text(ctx, slide, 130, 482, 960, 30, "当前验证位置", { size: 21, bold: true, color: C.dark });
    text(ctx, slide, 130, 522, 960, 38, "网页已能打开，说明 SoftAP、HTTP 服务和 SPIFFS 资源链路已基本成立；下一步重点是让 MLX90640 能稳定被 I2C 识别并返回温度数组。", { size: 17, color: C.muted });
  }

  section(presentation, ctxFor(), no++, "一、器件与硬件接口", "这一部分解释为什么裸 MLX90640 不能像模块一样直接乱插，以及 ESP32-C3 上哪些脚位承担关键功能。", [
    "1",
    { title: "MLX90640", body: "裸传感器与 I2C" },
    { title: "ESP32-C3", body: "Wi-Fi 与控制核心" },
    { title: "面包板", body: "临时无焊连接方案" },
  ]);

  twoColumns(presentation, ctxFor(), no++, "MLX90640 BASICS", "MLX90640 的作用：把红外辐射转成 32×24 温度矩阵", "它不是普通摄像头，而是热红外阵列传感器，每个像素对应一个温度测量点。",
    {
      title: "核心理解",
      items: [
        "32×24 = 768 个热像素，分辨率不高，但足够展示人体、热水杯、电子器件发热。",
        "传感器内部有 EEPROM 校准参数，程序需要先读取参数，再把原始帧换算为摄氏温度。",
        "I2C 地址通常为 0x33，本项目代码也按 0x33 进行访问。",
        "采集频率越高，对 I2C 总线、电源稳定性和程序处理能力要求越高。",
      ],
      gap: 52,
      itemH: 42,
    },
    {
      title: "为什么适合大作业",
      items: [
        "能体现总线通信：SDA/SCL、上拉电阻、地址访问、数据帧读取。",
        "能体现数据处理：原始数据不是最终温度，需要进行参数补偿与矩阵转换。",
        "能体现人机交互：温度矩阵可以在网页端变成彩色热图。",
        "能体现调试过程：硬件未稳定时，网页能开但温度为空，这是典型系统分层问题。",
      ],
      gap: 52,
      itemH: 42,
      fill: C.paleBlue,
    },
  );

  {
    const ctx = ctxFor();
    const slide = slideBase(presentation, ctx, no++, "MLX90640 PINOUT", "裸 MLX90640 引脚方向：先认方向，再谈接线", "以下方向按“传感器底部金属面朝上，小凸起在上方”来说明；翻面后左右会反。");
    panel(ctx, slide, 85, 165, 445, 385, C.panel, C.line);
    text(ctx, slide, 155, 198, 300, 30, "底面朝上，小凸起在上方", { size: 24, bold: true, color: C.dark, align: "center" });
    addShape(ctx, slide, 292, 228, 32, 22, C.orange);
    addShape(ctx, slide, 238, 255, 145, 145, C.codeBg);
    [["SCL", 180, 255, "左上"], ["SDA", 392, 255, "右上"], ["GND", 180, 360, "左下"], ["VDD", 392, 360, "右下"]].forEach(([pin, x, y, pos]) => {
      panel(ctx, slide, x, y, 62, 38, C.paleBlue, "#C9D8EA");
      text(ctx, slide, x, y + 8, 62, 20, pin, { size: 15, bold: true, color: C.dark, align: "center" });
      text(ctx, slide, x, y + 48, 62, 20, pos, { size: 12, color: C.muted, align: "center" });
    });
    text(ctx, slide, 128, 470, 360, 48, "重要：这是从传感器底部看的方向；一旦翻到正面或转动角度，左右关系会变化，接线前必须重新确认。", { size: 16, color: C.red, bold: true });
    table(ctx, slide, 590, 165, [180, 180, 250], [
      ["MLX90640", "ESP32-C3", "说明"],
      ["VDD", "3.3V", "只能接 3.3V，不接 +5V"],
      ["GND", "GND", "公共地，所有信号必须共地"],
      ["SDA", "IO04 / GPIO4", "I2C 数据线，代码已同步"],
      ["SCL", "IO05 / GPIO5", "I2C 时钟线，代码已同步"],
    ], { rowH: 56, size: 14 });
    text(ctx, slide, 625, 505, 540, 44, "代码匹配：SDA = GPIO4，SCL = GPIO5，MLX90640 地址 = 0x33。硬件接线和代码必须保持一致。", { size: 18, bold: true, color: C.dark });
  }

  {
    const ctx = ctxFor();
    const slide = slideBase(presentation, ctx, no++, "ESP32-C3 BOARD", "ESP32-C3 开发板识别：引脚丝印决定接线位置", "实物板背面丝印已经确认，使用 IO04/IO05 作为 I2C 引脚。");
    await image(ctx, slide, "esp32_back.jpg", 70, 150, 470, 360, "cover");
    panel(ctx, slide, 590, 150, 560, 360, C.panel, C.line);
    table(ctx, slide, 610, 175, [150, 160, 210], [
      ["板上丝印", "软件 GPIO", "本项目用途"],
      ["IO04", "GPIO4", "连接 MLX90640 SDA"],
      ["IO05", "GPIO5", "连接 MLX90640 SCL"],
      ["3.3V", "3.3V", "传感器 VDD"],
      ["GND", "GND", "公共地"],
      ["USB-C", "CH343 串口", "烧录与串口日志"],
    ], { rowH: 48, size: 13 });
    text(ctx, slide, 86, 540, 990, 38, "接线依据来自实物开发板丝印与程序 GPIO 定义，硬件连接和软件配置保持一致。", { size: 17, color: C.red, bold: true, align: "center" });
  }

  twoColumns(presentation, ctxFor(), no++, "BREADBOARD LOGIC", "面包板连接逻辑：同一排是同一个电气节点，不是串联", "理解这一点后，公对母线、跳线、电阻、电容就都能按节点连接。",
    {
      title: "如何理解“同一排”",
      items: [
        "面包板中间区域通常每一横排的 5 个孔内部相通，相当于同一个接线点。",
        "传感器母头线插到某一排后，ESP32 的对应引脚也插到同一排，就表示两者已经连接。",
        "电阻和电容不是接到“线的中间”，而是跨接在两个节点之间。",
        "电源轨红线/蓝线通常是纵向相通，但中间断开处要用跳线桥接才会连续。",
      ],
      size: 16,
      gap: 49,
      itemH: 40,
    },
    {
      title: "本项目对应节点",
      items: [
        "SDA 节点：MLX90640 SDA、ESP32 IO04、5k 上拉到 3.3V。",
        "SCL 节点：MLX90640 SCL、ESP32 IO05、5k 上拉到 3.3V。",
        "VDD 节点：ESP32 3.3V、MLX90640 VDD、10uF 电容正端。",
        "GND 节点：ESP32 GND、MLX90640 GND、10uF 电容负端。",
      ],
      fill: C.paleOrange,
      size: 16,
      gap: 49,
      itemH: 40,
    },
    "关键结论：跳线只是把两个节点连起来；真正要检查的是每根线两端所在的“电气节点”是否正确。");

  twoColumns(presentation, ctxFor(), no++, "PULLUP & CAPACITOR", "为什么需要电阻和电容：解决总线电平与供电稳定", "裸传感器没有模块板自带的外围元件，因此需要在面包板上补齐基础条件。",
    {
      title: "5k 上拉电阻",
      items: [
        "I2C 的 SDA/SCL 本质上需要上拉到高电平，器件通过拉低线来发送 0。",
        "没有上拉时，总线高电平可能不稳定，表现为找不到地址、初始化失败、数据异常。",
        "你手头的 5kΩ 可以使用，效果接近常见 4.7kΩ 上拉。",
        "连接方式：SDA 到 3.3V 一个 5kΩ，SCL 到 3.3V 一个 5kΩ。",
      ],
      gap: 50,
      itemH: 40,
      fill: C.paleGreen,
    },
    {
      title: "10uF 滤波电容",
      items: [
        "电容跨接在 VDD 与 GND 之间，不参与数据传输，主要用于缓冲电源波动。",
        "如果是电解电容，正极接 VDD/3.3V，负极接 GND，极性不能反。",
        "10uF 可以先用；若后续有 0.1uF，可并联在传感器旁边抑制高频噪声。",
        "电容越靠近传感器供电脚越好，线太长会降低滤波效果。",
      ],
      gap: 50,
      itemH: 40,
      fill: C.paleBlue,
    });

  {
    const ctx = ctxFor();
    const slide = slideBase(presentation, ctx, no++, "CURRENT HARDWARE", "当前硬件状态：主连接已做，稳定性元件暂未完全补齐", "网页链路已验证，传感器数据链路仍需完成硬件稳定性检查。");
    await image(ctx, slide, "wiring_current.jpg", 70, 150, 520, 370, "cover");
    panel(ctx, slide, 650, 150, 500, 370, C.panel, C.line);
    text(ctx, slide, 690, 184, 420, 28, "已经完成", { size: 23, bold: true, color: C.dark });
    bullets(ctx, slide, 694, 245, 380, [
      "ESP32-C3 开发板可通过 USB-C 连接电脑。",
      "COM5 已确认为 CH343 串口，软件烧录链路已验证。",
      "SDA/SCL/VDD/GND 主连线已经按方案尝试。",
      "网页端已能访问，说明网络和文件系统正常。",
    ], { gap: 46, size: 16, itemH: 36 });
    text(ctx, slide, 690, 450, 420, 28, "尚未完全验证", { size: 20, bold: true, color: C.red });
    text(ctx, slide, 690, 485, 420, 46, "传感器裸脚接触可靠性、两个 5kΩ 上拉、10uF 电容，以及 /thermal-data 是否能返回真实温度数组。", { size: 15, color: C.muted });
  }

  section(presentation, ctxFor(), no++, "二、软件环境与程序结构", "这一部分说明 VS Code、PlatformIO、ESP-IDF、分区表、网页资源之间的关系，避免只说“我烧录好了”。", [
    "2",
    { title: "VS Code", body: "编辑与入口" },
    { title: "PlatformIO", body: "工具链管理" },
    { title: "ESP-IDF", body: "底层框架" },
  ]);

  twoColumns(presentation, ctxFor(), no++, "PLATFORMIO ROLE", "PlatformIO 的作用：把环境、编译、烧录统一起来", "PlatformIO 不是芯片，也不是底层框架，而是工程管理和工具链入口。",
    {
      title: "为什么不用手动配置",
      items: [
        "自动下载/管理 ESP32-C3 交叉编译工具链和 ESP-IDF 框架。",
        "根据 platformio.ini 选择芯片、框架、分区表和文件系统。",
        "统一提供 Build、Upload、Monitor、Upload Filesystem 等命令。",
        "减少 PATH、Python、编译器版本和烧录参数不一致带来的问题。",
      ],
      gap: 50,
      itemH: 40,
      fill: C.panel,
    },
    {
      title: "在本项目中的具体作用",
      items: [
        "编译 C/ESP-IDF 主程序，生成 firmware.bin。",
        "调用 esptool.py 通过 COM5 把程序写入 ESP32-C3。",
        "把 data 目录下网页脚本打包为 spiffs.bin 并写入 flash。",
        "提供串口监视，便于观察启动日志和传感器初始化错误。",
      ],
      gap: 50,
      itemH: 40,
      fill: C.paleBlue,
    });

  {
    const ctx = ctxFor();
    const slide = slideBase(presentation, ctx, no++, "ENVIRONMENT CONFIG", "PlatformIO 配置：一份 platformio.ini 定义整个构建环境", "本项目的开发环境、芯片目标、文件系统和烧录端口由 platformio.ini 统一描述。");
    table(ctx, slide, 70, 150, [220, 355, 460], [
      ["配置项", "当前配置", "意义"],
      ["platform", "espressif32@6.10.0", "指定 ESP32 平台包版本，保证工具链和框架一致"],
      ["board", "esp32-c3-devkitm-1", "以通用 ESP32-C3 开发板作为编译目标"],
      ["framework", "espidf", "使用 ESP-IDF，适合 HTTP、Wi-Fi、SPIFFS 等系统级能力"],
      ["filesystem", "spiffs", "把网页脚本作为文件系统资源写入 flash"],
      ["partitions.csv", "1MB app + 1MB SPIFFS", "为程序和网页资源分别预留空间"],
      ["upload_port", "COM5", "Windows 设备管理器确认的 CH343 串口"],
    ], { rowH: 50, size: 13 });
    code(ctx, slide, 160, 520, 860, 116, "[env:esp32-c3-devkitm-1]\nplatform = espressif32@6.10.0\nboard = esp32-c3-devkitm-1\nframework = espidf", "platformio.ini 核心片段");
  }

  {
    const ctx = ctxFor();
    const slide = slideBase(presentation, ctx, no++, "PROJECT FILES", "工程文件结构：每个目录在系统中都有明确职责", "本项目由主程序、网页资源、传感器驱动、构建配置和分区表共同组成。");
    table(ctx, slide, 80, 155, [230, 365, 455], [
      ["路径/文件", "内容", "项目作用"],
      ["src/main.c", "主程序逻辑", "Wi-Fi、HTTP、I2C、MLX90640 采集、接口返回"],
      ["data/web_script.js", "网页前端脚本", "请求温度数据，绘制热图，处理按钮和参数"],
      ["lib/MLX90640", "传感器驱动库", "读取 EEPROM 参数、计算温度矩阵"],
      ["platformio.ini", "构建配置", "指定平台、框架、板型、文件系统"],
      ["partitions.csv", "Flash 分区表", "定义 app 与 SPIFFS 的烧录位置和大小"],
      ["sdkconfig.*", "ESP-IDF 配置", "4MB flash、组件选项和系统参数"],
    ], { rowH: 56, size: 13 });
  }

  {
    const ctx = ctxFor();
    const slide = slideBase(presentation, ctx, no++, "CODE PARAMETERS", "程序关键参数：硬件接线必须与代码定义一致", "IO04/IO05 的选择来自程序中的 I2C 引脚定义，传感器地址采用 MLX90640 默认地址 0x33。");
    code(ctx, slide, 70, 150, 520, 320, `#define I2C_MASTER_SDA_IO GPIO_NUM_4\n#define I2C_MASTER_SCL_IO GPIO_NUM_5\n#define I2C_MASTER_FREQ_HZ 400000\n#define MLX90640_I2C_ADDR 0x33`, "I2C 与传感器地址");
    panel(ctx, slide, 650, 150, 500, 320, C.panel, C.line);
    text(ctx, slide, 690, 184, 420, 30, "这些参数决定了什么", { size: 23, bold: true, color: C.dark });
    bullets(ctx, slide, 694, 245, 390, [
      "SDA/SCL 决定 ESP32 从哪两个引脚访问传感器。",
      "400kHz 是 I2C 快速模式，适合读取 32×24 阵列帧数据。",
      "0x33 是 MLX90640 默认 I2C 地址，若地址读不到，网页端就不会有温度数组。",
      "硬件换脚后必须改代码，否则电气连接正确也无法通信。",
    ], { gap: 48, size: 16, itemH: 39 });
    text(ctx, slide, 120, 535, 980, 40, "调试原则：先确认 VDD/GND，再确认 SDA/SCL，再确认代码引脚，最后看串口日志和 /thermal-data 返回。", { size: 18, bold: true, color: C.red, align: "center" });
  }

  {
    const ctx = ctxFor();
    const slide = slideBase(presentation, ctx, no++, "DATA PIPELINE", "软件数据流：从 I2C 帧读取到网页热图刷新", "系统运行过程按初始化、采集、温度换算、HTTP 返回和浏览器渲染依次完成。");
    node(ctx, slide, 70, 230, 170, 120, "初始化", "I2C 总线\nMLX 参数\nWi-Fi AP", C.paleGreen);
    arrow(ctx, slide, 260, 280, 62);
    node(ctx, slide, 340, 230, 170, 120, "采集帧", "读取原始帧\n校验状态\n异常重试", C.paleOrange);
    arrow(ctx, slide, 530, 280, 62);
    node(ctx, slide, 610, 230, 170, 120, "温度换算", "EEPROM 参数\n像素补偿\nTa/To 计算", C.paleBlue);
    arrow(ctx, slide, 800, 280, 62);
    node(ctx, slide, 880, 230, 170, 120, "HTTP 返回", "JSON 数组\nmin/max\n状态字段", C.panel);
    arrow(ctx, slide, 1070, 280, 62);
    node(ctx, slide, 70, 430, 1080, 96, "浏览器渲染", "网页脚本周期性请求 /thermal-data，把 768 个温度值映射成颜色块；插值模式只影响视觉效果，不改变原始传感器分辨率。", C.paleGreen);
    text(ctx, slide, 120, 578, 1000, 34, "当前卡点位于最左侧传感器数据入口：网页层已打开，但 MLX90640 数据尚未稳定进入。", { size: 18, color: C.red, bold: true, align: "center" });
  }

  section(presentation, ctxFor(), no++, "三、烧录与运行过程", "这一部分把 COM5、BOOT 下载模式、主固件、SPIFFS 两次烧录和网页访问讲清楚。", [
    "3",
    { title: "COM5", body: "USB 串口链路" },
    { title: "Firmware", body: "主程序" },
    { title: "SPIFFS", body: "网页资源" },
  ]);

  twoColumns(presentation, ctxFor(), no++, "SERIAL PORT", "COM5 是烧录通道，不是网页网络端口", "设备管理器确认 USB-Enhanced-SERIAL CH343(COM5)，说明电脑已经识别开发板串口芯片。",
    {
      title: "COM5 的作用",
      items: [
        "电脑通过 USB-C 线连接 ESP32-C3 开发板。",
        "PlatformIO 使用 COM5 调用 esptool.py 下载程序。",
        "串口监视器也通过 COM5 查看启动日志、报错和调试信息。",
        "COM5 是有线调试通道，拔掉 USB 后电脑就无法通过它继续烧录。",
      ],
      gap: 48,
      itemH: 39,
    },
    {
      title: "HTTP 80 的作用",
      items: [
        "ESP32-C3 自己创建 Wi-Fi 热点，手机/电脑连接热点。",
        "浏览器访问 192.168.4.1，实际访问的是 ESP32 上的 HTTP 服务。",
        "HTTP 默认使用 80 端口，这是无线网页服务端口。",
        "COM5 与 HTTP 80 完全不是一回事：一个是 USB 串口，一个是网络服务。",
      ],
      gap: 48,
      itemH: 39,
      fill: C.paleBlue,
    });

  {
    const ctx = ctxFor();
    const slide = slideBase(presentation, ctx, no++, "UPLOAD 1", "第一次烧录：主固件 firmware 写入 app 分区", "主固件包含 ESP32-C3 的核心运行逻辑，负责硬件初始化、网络服务和传感器采集。");
    code(ctx, slide, 80, 160, 520, 250, `pio run -t upload --upload-port COM5\n\n# 写入内容：\n# bootloader.bin\n# partition-table.bin\n# firmware.bin\n\n# 典型写入地址：\n# app 分区 0x10000`, "固件烧录命令");
    panel(ctx, slide, 660, 160, 470, 250, C.panel, C.line);
    text(ctx, slide, 700, 194, 390, 30, "固件烧录操作要点", { size: 22, bold: true, color: C.dark });
    bullets(ctx, slide, 704, 250, 360, [
      "确认 COM5 仍然存在。",
      "若自动进入下载失败，按住 BOOT 再点 RESET。",
      "看到写入百分比和校验完成后，说明主程序烧录成功。",
      "固件烧完通常会自动复位，因此第二次烧录前可能要重新进入 BOOT。",
    ], { gap: 38, size: 15, itemH: 31 });
    text(ctx, slide, 120, 515, 980, 42, "如果只烧录主固件、不烧录 SPIFFS，ESP32 可以运行程序，但网页脚本等资源可能缺失。", { size: 18, color: C.red, bold: true, align: "center" });
  }

  {
    const ctx = ctxFor();
    const slide = slideBase(presentation, ctx, no++, "UPLOAD 2", "第二次烧录：SPIFFS 网页资源写入 storage 分区", "本项目网页不是从电脑服务器加载，而是存放在 ESP32 的 flash 文件系统里。");
    code(ctx, slide, 80, 160, 520, 250, `pio run -t uploadfs --upload-port COM5\n\n# 写入内容：\n# spiffs.bin\n# 包含 data/web_script.js 等网页资源\n\n# 典型写入地址：\n# storage/SPIFFS 分区 0x110000`, "SPIFFS 烧录命令");
    panel(ctx, slide, 660, 160, 470, 250, C.paleOrange, "#F3D3A1");
    text(ctx, slide, 700, 194, 390, 30, "为什么要单独烧录网页资源", { size: 22, bold: true, color: C.dark });
    bullets(ctx, slide, 704, 250, 360, [
      "主程序和网页资源属于不同 flash 分区。",
      "网页脚本更新时，可以只更新 SPIFFS。",
      "浏览器能打开页面，说明 SPIFFS 和 HTTP 静态资源服务已经有效。",
      "若 SPIFFS 缺失，可能出现页面空白、脚本 404 或按钮无响应。",
    ], { gap: 38, size: 15, itemH: 31, dot: C.orange });
    text(ctx, slide, 120, 515, 980, 42, "这就是本项目存在“两次烧录”的原因：一次给 ESP32 写程序，一次给网页端写资源。", { size: 18, color: C.red, bold: true, align: "center" });
  }

  twoColumns(presentation, ctxFor(), no++, "BOOT MODE", "BOOT/RESET 操作：解决 ESP32-C3 无法进入下载模式的问题", "如果 PlatformIO 连接失败，不一定是代码错，可能只是芯片没有进入 ROM 下载模式。",
    {
      title: "推荐操作步骤",
      items: [
        "先连接 USB-C，并在设备管理器确认 COM5。",
        "点击 PlatformIO Upload 或执行 pio upload 命令。",
        "若提示连接超时，按住 BOOT，轻按 RESET，再松开 BOOT。",
        "等待终端出现 writing / verifying / hash verified 等信息。",
      ],
      gap: 48,
      itemH: 39,
      fill: C.panel,
    },
    {
      title: "常见失败原因",
      items: [
        "线只有充电功能，没有数据线功能。",
        "COM 口选错，仍指向蓝牙虚拟串口而不是 CH343。",
        "固件烧录后自动复位，第二次 uploadfs 时需要重新进入下载模式。",
        "串口监视器占用了 COM5，导致烧录端口无法打开。",
      ],
      gap: 48,
      itemH: 39,
      fill: C.paleRed,
      dot: C.red,
    });

  section(presentation, ctxFor(), no++, "四、网络端与网页端", "这一部分解释手机为什么能看到网页，以及 /thermal-data 和按钮控制的作用。", [
    "4",
    { title: "SoftAP", body: "ESP32 自建热点" },
    { title: "HTTP", body: "浏览器访问" },
    { title: "Web UI", body: "热图绘制" },
  ]);

  {
    const ctx = ctxFor();
    const slide = slideBase(presentation, ctx, no++, "SOFTAP MODEL", "ESP32 网络端：本项目使用 SoftAP 热点模式", "不依赖宿舍/教室路由器，ESP32-C3 自己就是一个小型无线接入点。");
    node(ctx, slide, 95, 245, 210, 120, "ESP32-C3", "创建 Wi-Fi 热点\nSSID: IRCAM-XXXX\n默认地址 192.168.4.1", C.paleGreen);
    arrow(ctx, slide, 335, 295, 110);
    node(ctx, slide, 480, 245, 210, 120, "手机/电脑", "连接这个热点\n获得同一网段 IP\n打开浏览器", C.paleBlue);
    arrow(ctx, slide, 720, 295, 110);
    node(ctx, slide, 865, 245, 210, 120, "网页服务", "HTTP 80\n/ 首页\n/thermal-data 数据", C.paleOrange);
    panel(ctx, slide, 110, 460, 950, 94, C.panel, C.line);
    text(ctx, slide, 135, 486, 900, 38, "端口关系：COM5 属于 USB 串口通道，负责烧录与日志；HTTP 80 属于 TCP/IP 网络端口，负责浏览器访问。", { size: 17, color: C.dark, bold: true, align: "center" });
  }

  {
    const ctx = ctxFor();
    const slide = slideBase(presentation, ctx, no++, "HTTP ENDPOINTS", "网页端数据接口：页面、脚本、温度数组各有路径", "浏览器看到的热图来自周期性请求，而不是一次性截图。");
    table(ctx, slide, 80, 160, [220, 380, 450], [
      ["HTTP 路径", "返回内容", "本项目用途"],
      ["/", "主页面 HTML", "浏览器入口，加载控制界面"],
      ["/web_script.js", "网页脚本", "处理按钮、请求数据、绘制热图"],
      ["/thermal-data", "温度 JSON 数组", "返回 32×24 温度矩阵和 min/max"],
      ["/settings 或相关接口", "参数更新结果", "更新发射率、温度范围、采集状态等"],
    ], { rowH: 64, size: 14 });
    text(ctx, slide, 130, 550, 980, 42, "当前页面能打开但 min/max 为 --°C，说明页面和脚本加载成功，但传感器数据接口还没有拿到有效温度。", { size: 18, color: C.red, bold: true, align: "center" });
  }

  {
    const ctx = ctxFor();
    const slide = slideBase(presentation, ctx, no++, "WEB UI RESULT", "当前网页结果：能访问页面，等待传感器数据进入", "这张截图可以作为阶段性成果，同时也解释下一步为什么要回到硬件排查。");
    await image(ctx, slide, "webpage.jpg", 90, 150, 335, 500, "contain");
    panel(ctx, slide, 500, 150, 610, 500, C.panel, C.line);
    text(ctx, slide, 540, 186, 530, 30, "从截图能判断出的结论", { size: 24, bold: true, color: C.dark });
    bullets(ctx, slide, 544, 250, 480, [
      "浏览器已经进入 ESP32-C3 提供的页面，说明 Wi-Fi AP 和 HTTP 服务可用。",
      "页面上的按钮、温度范围、发射率输入框能显示，说明 SPIFFS 网页资源已经成功写入。",
      "热图区为黑色、最低/最高温度为 --°C，说明还没有收到有效温度矩阵。",
      "因此当前问题应优先查 MLX90640 供电、SDA/SCL、上拉电阻、裸脚接触，而不是重新写网页。",
    ], { gap: 58, size: 16, itemH: 46 });
  }

  {
    const ctx = ctxFor();
    const slide = slideBase(presentation, ctx, no++, "FRONTEND CONTROLS", "网页控件含义：不是装饰按钮，而是调试入口", "这些控件可以帮助展示传感器状态和图像显示效果。");
    table(ctx, slide, 80, 155, [230, 300, 520], [
      ["控件/字段", "作用", "展示时可以怎么讲"],
      ["开始/停止采集", "控制 ESP32 是否持续读取传感器", "用于说明采集任务可以被网页端控制"],
      ["插值模式", "把 32×24 原始矩阵放大显示", "插值让视觉更平滑，但原始数据仍是 32×24"],
      ["温度范围", "手动设置颜色映射上下限", "可以突出热源，避免自动范围导致对比度不稳定"],
      ["发射率", "影响温度换算参数", "真实测温会受到材料发射率影响，本项目默认 0.95"],
      ["min/max", "显示当前帧最低/最高温", "传感器连通后这里会从 --°C 变成具体数值"],
    ], { rowH: 58, size: 13 });
  }

  section(presentation, ctxFor(), no++, "五、单片机平台选择与课程思考", "这一部分用于把项目放回“微机原理/接口技术”的课程框架，说明为什么选择 ESP32-C3。", [
    "5",
    { title: "C51", body: "传统入门" },
    { title: "STM32", body: "工业控制" },
    { title: "ESP32", body: "联网嵌入式" },
  ]);

  {
    const ctx = ctxFor();
    const slide = slideBase(presentation, ctx, no++, "MCU COMPARISON", "C51、STM32、ESP32 的区别：从传统控制到联网嵌入式", "这一页单独用于回答老师可能问的“为什么不用 51 或 STM32”。");
    table(ctx, slide, 55, 145, [130, 230, 250, 250, 255], [
      ["维度", "C51 单片机", "STM32 单片机", "ESP32 / ESP32-C3", "本项目选择原因"],
      ["核心定位", "经典 8 位单片机", "ARM Cortex-M 微控制器", "带 Wi-Fi/BLE 的 SoC", "需要无线热点和网页服务"],
      ["性能/资源", "频率低，RAM/Flash 较小", "性能跨度大，外设丰富", "频率高，片上资源适合网络应用", "可同时处理采集和 Web"],
      ["联网能力", "通常需要外接模块", "通常需要外接 Wi-Fi/以太网模块", "内置 Wi-Fi/BLE", "直接自建 AP 热点"],
      ["开发方式", "Keil C51 常见", "Keil/STM32CubeIDE/PlatformIO", "ESP-IDF/Arduino/PlatformIO", "PlatformIO + ESP-IDF"],
      ["典型场景", "基础控制、教学入门", "工业控制、外设驱动、实时控制", "物联网、无线传感、网页交互", "热成像网页展示更合适"],
    ], { rowH: 64, size: 12, headerSize: 13 });
  }

  threeCards(presentation, ctxFor(), no++, "WHY ESP32-C3", "为什么本项目选 ESP32-C3：不是因为它最强，而是它刚好合适", "项目目标包含传感器采集与浏览器展示，ESP32-C3 的内置 Wi-Fi 能显著减少外围硬件。",
    [
      { title: "硬件更少", body: "不需要再额外接 Wi-Fi 模块；USB-C 开发板带串口芯片，烧录调试链路也更简单。", fill: C.paleGreen, accent: C.green },
      { title: "功能闭环", body: "一个芯片同时完成 I2C 采集、温度计算、HTTP 服务和无线热点，适合做成独立演示设备。", fill: C.paleBlue, accent: C.blue },
      { title: "课程价值", body: "既能讲传统接口问题，如 I2C、上拉、供电，也能讲现代嵌入式系统，如文件系统和网络服务。", fill: C.paleOrange, accent: C.orange },
    ]);

  section(presentation, ctxFor(), no++, "六、当前结果、问题定位与后续计划", "本部分汇总已完成内容、未完成硬件点、最终验证方法和后续工作安排。", [
    "6",
    { title: "已完成", body: "软件与网页链路" },
    { title: "待完成", body: "MLX90640 数据链路" },
    { title: "最终展示", body: "实时热图" },
  ]);

  {
    const ctx = ctxFor();
    const slide = slideBase(presentation, ctx, no++, "CURRENT STATUS", "当前完成度：软件主链路完成，硬件数据入口待验证", "用状态表说明进展，避免只说“差不多好了”。");
    table(ctx, slide, 80, 155, [230, 220, 600], [
      ["模块", "状态", "证据/说明"],
      ["PlatformIO 环境", "完成", "VS Code 中可编译；工具链和 ESP-IDF 已安装到项目环境"],
      ["主固件烧录", "完成", "COM5 写入并校验通过；固件在 app 分区运行"],
      ["SPIFFS 烧录", "完成", "网页脚本写入 storage/SPIFFS 分区；页面可以访问"],
      ["网页访问", "完成", "手机/电脑连接 ESP32 热点后可打开 192.168.4.1"],
      ["MLX90640 数据", "待验证", "裸传感器接线、上拉、电容和接触稳定性仍需收尾"],
    ], { rowH: 66, size: 13 });
  }

  twoColumns(presentation, ctxFor(), no++, "HARDWARE CHECKLIST", "下一次接线后优先检查什么：先电源，再总线，再数据", "硬件复测按照上电前检查和上电后检查两部分进行。",
    {
      title: "上电前检查",
      items: [
        "确认 MLX90640 VDD 只接 3.3V，绝不接 +5V。",
        "确认 GND 共地：ESP32 GND 与传感器 GND 必须连通。",
        "确认裸传感器四个引脚没有相互碰到或被母头线挤压短路。",
        "确认 10uF 电容极性正确，正极 VDD，负极 GND。",
      ],
      gap: 48,
      itemH: 39,
      fill: C.paleGreen,
    },
    {
      title: "上电后检查",
      items: [
        "SDA 接 IO04，SCL 接 IO05，不能反接。",
        "两个 5kΩ 上拉分别从 SDA/SCL 接到 3.3V。",
        "打开串口监视器，看是否找到 MLX90640 地址 0x33。",
        "访问 /thermal-data，看是否返回 768 个温度值。",
      ],
      gap: 48,
      itemH: 39,
      fill: C.paleBlue,
    });

  {
    const ctx = ctxFor();
    const slide = slideBase(presentation, ctx, no++, "FINAL VALIDATION", "最终验证标准：热源移动时，网页热图和 min/max 同步变化", "最终结果不只看网页能打开，而是看传感器数据是否能驱动画面变化。");
    node(ctx, slide, 90, 210, 230, 135, "1. 串口确认", "启动日志无 I2C 初始化失败\n能识别 MLX90640", C.paleGreen);
    arrow(ctx, slide, 350, 268, 100);
    node(ctx, slide, 480, 210, 230, 135, "2. 数据确认", "/thermal-data 返回 JSON\n包含 768 个温度点", C.paleBlue);
    arrow(ctx, slide, 740, 268, 100);
    node(ctx, slide, 870, 210, 230, 135, "3. 画面确认", "手靠近/热水杯靠近\n热图颜色局部变化", C.paleOrange);
    panel(ctx, slide, 120, 460, 960, 90, C.panel, C.line);
    text(ctx, slide, 150, 490, 900, 34, "最终可展示现象", { size: 22, bold: true, color: C.dark, align: "center" });
    text(ctx, slide, 150, 525, 900, 32, "热源进入视野后最高温度上升，热图出现明显高温区域；移走热源后颜色逐渐恢复。", { size: 17, color: C.muted, align: "center" });
  }

  twoColumns(presentation, ctxFor(), no++, "TROUBLESHOOTING", "常见问题定位：网页链路和传感器链路分开判断", "调试过程按照分层思路定位问题，先判断网页服务，再回到传感器通信入口。",
    {
      title: "网页能开但无温度",
      items: [
        "优先判断网页层正常，问题多半在 MLX90640 入口。",
        "检查 VDD/GND、SDA/SCL、5k 上拉、电容极性。",
        "看串口日志是否提示 I2C 读失败或 EEPROM 参数读取失败。",
        "用 /thermal-data 判断接口是否返回有效数组。",
      ],
      gap: 48,
      itemH: 39,
    },
    {
      title: "烧录或连接失败",
      items: [
        "确认数据线不是纯充电线，设备管理器能看到 COM5。",
        "串口监视器关闭后再烧录，避免 COM5 被占用。",
        "主固件烧完后，uploadfs 失败时重新进入 BOOT 下载模式。",
        "若改了引脚，必须重新编译并烧录主固件。",
      ],
      gap: 48,
      itemH: 39,
      fill: C.paleOrange,
    });

  {
    const ctx = ctxFor();
    const slide = slideBase(presentation, ctx, no++, "PROJECT TRACE", "项目实施脉络：从目标到实现，再到验证", "项目过程按照目标定义、硬件连接、软件实现、阶段结果和后续验证展开。");
    table(ctx, slide, 80, 150, [190, 430, 430], [
      ["环节", "项目内容", "支撑材料"],
      ["目标定义", "基于 ESP32-C3 和 MLX90640 构建网页热成像相机", "项目目标页与总体架构页"],
      ["硬件连接", "裸 MLX90640 使用 3.3V 供电，I2C 总线需要上拉和稳定接触", "引脚页、接线表、面包板页"],
      ["软件实现", "PlatformIO 管理 ESP-IDF 环境，完成编译、烧录和文件系统上传", "配置页、文件结构页、烧录页"],
      ["阶段结果", "网页已能打开，网络与 SPIFFS 链路成功；热图等待传感器数据连通", "网页截图页、当前状态页"],
      ["后续验证", "补齐上拉和电容后，验证 /thermal-data 与实时热图变化", "检查清单与最终验证页"],
    ], { rowH: 70, size: 13 });
  }

  threeCards(presentation, ctxFor(), no++, "FINAL SUMMARY", "阶段总结：软件地基已完成，下一步集中打通 MLX90640 数据入口", "当前项目已形成可运行的软件框架，后续重点集中在裸传感器数据链路的稳定性验证。",
    [
      { title: "已经证明", body: "ESP32-C3 工程能编译、能烧录、能创建热点、能提供网页。当前软件框架具备继续接入热数据的基础。", fill: C.paleGreen, accent: C.green },
      { title: "还要证明", body: "MLX90640 裸件接线稳定后，I2C 能读取 EEPROM 和帧数据，/thermal-data 能返回温度矩阵。", fill: C.paleBlue, accent: C.blue },
      { title: "最终展示", body: "把热源放到传感器前，网页热图出现明显颜色变化，同时 min/max 温度同步变化。", fill: C.paleOrange, accent: C.orange },
    ]);

  const slides = [];
  for (let i = 0; i < presentation.slides.count; i += 1) {
    slides.push(presentation.slides.getItem(i));
  }

  const previewPaths = [];
  for (let i = 0; i < slides.length; i += 1) {
    const preview = await presentation.export({ slide: slides[i], format: "png", scale: 1 });
    const previewPath = path.join(PREVIEW_DIR, `slide-${padSlideNumber(i + 1)}.png`);
    await saveBlobToFile(preview, previewPath);
    previewPaths.push(previewPath);
    try {
      const layout = await presentation.export({ slide: slides[i], format: "layout" });
      await fs.writeFile(path.join(LAYOUT_DIR, `slide-${padSlideNumber(i + 1)}.layout.json`), await layout.text(), "utf8");
    } catch {
      // Layout export is only used for QA; PPTX export remains the deliverable.
    }
  }

  const pptx = await PresentationFile.exportPptx(presentation);
  await pptx.save(FINAL_PPTX);
  await fs.copyFile(FINAL_PPTX, FINAL_ASCII);

  const manifest = {
    output: FINAL_PPTX,
    asciiOutput: FINAL_ASCII,
    outputBytes: (await fs.stat(FINAL_PPTX)).size,
    slideCount: presentation.slides.count,
    previewDir: PREVIEW_DIR,
    previewPaths,
  };
  await fs.writeFile(path.join(WORKSPACE, "artifact-build-manifest.json"), JSON.stringify(manifest, null, 2), "utf8");
  console.log(JSON.stringify(manifest, null, 2));
}

await buildDeck();
