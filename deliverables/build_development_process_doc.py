"""Generate the project development-process DOCX.

This deliverable is maintained in the name of rdfrdream.  It intentionally
keeps the report concise, visual, and engineering-oriented: repository files,
hardware evidence, data flow, web visualization, and debugging logic.
"""

from __future__ import annotations

import math
import random
import shutil
from pathlib import Path

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Inches, Pt, RGBColor

try:
    from PIL import Image, ImageDraw, ImageFilter, ImageFont
except ImportError as exc:  # pragma: no cover - generation environment check
    raise SystemExit("Pillow is required to build the visual assets.") from exc


REPO_ROOT = Path(__file__).resolve().parents[1]
OUT_DIR = Path(__file__).resolve().parent
ASSET_DIR = OUT_DIR / "assets"
DOCX_PATH = OUT_DIR / "基于ESP32-C3与MLX90640的热成像相机系统开发流程.docx"

SOURCE_IMAGE_DIR = Path(r"H:\using for study\2026年春季\微机接口与技术\微机原理\图片与视频")
LEGACY_ASSET_DIR = Path(r"H:\ircam_pio_c3\outputs\manual-20260605-ircam\presentations\ircam-full\assets")

FONT_NAME = "微软雅黑"
MONO_FONT = "Consolas"
ACCENT = RGBColor(31, 78, 121)
ACCENT_DARK = RGBColor(15, 44, 76)
MUTED = RGBColor(92, 102, 112)
BODY = RGBColor(38, 45, 54)
LIGHT_FILL = "EEF4FA"
CALLOUT_FILL = "F5F8FB"


def font_path(bold: bool = False) -> str | None:
    candidates = [
        Path(r"C:\Windows\Fonts\msyhbd.ttc" if bold else r"C:\Windows\Fonts\msyh.ttc"),
        Path(r"C:\Windows\Fonts\simhei.ttf"),
        Path(r"C:\Windows\Fonts\arial.ttf"),
    ]
    for candidate in candidates:
        if candidate.exists():
            return str(candidate)
    return None


def pil_font(size: int, bold: bool = False) -> ImageFont.FreeTypeFont | ImageFont.ImageFont:
    path = font_path(bold)
    if path:
        return ImageFont.truetype(path, size=size)
    return ImageFont.load_default()


def fit_text(draw: ImageDraw.ImageDraw, text: str, max_width: int, start_size: int, bold: bool = False):
    size = start_size
    while size >= 16:
        font = pil_font(size, bold)
        if draw.textbbox((0, 0), text, font=font)[2] <= max_width:
            return font
        size -= 1
    return pil_font(16, bold)


def rounded_box(draw: ImageDraw.ImageDraw, xy, radius: int, fill: str, outline: str, width: int = 2):
    draw.rounded_rectangle(xy, radius=radius, fill=fill, outline=outline, width=width)


def draw_centered_multiline(draw: ImageDraw.ImageDraw, box, text: str, font, fill: str):
    x1, y1, x2, y2 = box
    lines = text.split("\n")
    line_boxes = [draw.textbbox((0, 0), line, font=font) for line in lines]
    total_h = sum(bb[3] - bb[1] for bb in line_boxes) + (len(lines) - 1) * 8
    y = y1 + (y2 - y1 - total_h) / 2
    for line, bb in zip(lines, line_boxes):
        w = bb[2] - bb[0]
        draw.text((x1 + (x2 - x1 - w) / 2, y), line, font=font, fill=fill)
        y += (bb[3] - bb[1]) + 8


def draw_arrow(draw: ImageDraw.ImageDraw, start: tuple[int, int], end: tuple[int, int], color: str = "#2E74B5"):
    draw.line([start, end], fill=color, width=6)
    angle = math.atan2(end[1] - start[1], end[0] - start[0])
    head = 20
    points = [
        end,
        (int(end[0] - head * math.cos(angle - math.pi / 7)), int(end[1] - head * math.sin(angle - math.pi / 7))),
        (int(end[0] - head * math.cos(angle + math.pi / 7)), int(end[1] - head * math.sin(angle + math.pi / 7))),
    ]
    draw.polygon(points, fill=color)


def create_flow_diagram(path: Path):
    w, h = 1900, 560
    img = Image.new("RGB", (w, h), "white")
    draw = ImageDraw.Draw(img)
    title_font = pil_font(42, True)
    box_font = pil_font(28, True)
    sub_font = pil_font(22)

    draw.text((70, 42), "系统总体数据链路", font=title_font, fill="#102A43")
    draw.text(
        (70, 98),
        "从红外阵列采样到浏览器热图显示，ESP32-C3 同时承担主控、无线接入和小型 Web Server。",
        font=sub_font,
        fill="#52616B",
    )

    labels = [
        "MLX90640\n热红外阵列",
        "I2C 总线\nSDA / SCL",
        "ESP32-C3\n采集与处理",
        "WiFi / AP\n无线接入",
        "HTTP Server\nJSON 接口",
        "浏览器热图\nCanvas 显示",
        "CSV\n数据记录",
    ]
    fills = ["#EAF3FF", "#F0F8FF", "#EAF7EF", "#FFF7E8", "#F5F2FF", "#EEF8FB", "#F3F5F7"]
    outlines = ["#2E74B5", "#2E74B5", "#198754", "#D9822B", "#6F42C1", "#087990", "#6C757D"]
    box_w, box_h = 230, 140
    gap = 32
    x = 70
    y = 245
    for i, label in enumerate(labels):
        box = (x, y, x + box_w, y + box_h)
        rounded_box(draw, box, 22, fills[i], outlines[i], 4)
        draw_centered_multiline(draw, box, label, box_font, "#102A43")
        if i < len(labels) - 1:
            draw_arrow(draw, (x + box_w + 5, y + box_h // 2), (x + box_w + gap - 5, y + box_h // 2))
        x += box_w + gap

    footer_font = pil_font(22)
    draw.text(
        (70, 455),
        "核心接口：GPIO4/SDA、GPIO5/SCL、3.3V、GND；网页访问：AP 模式默认 192.168.4.1。",
        font=footer_font,
        fill="#52616B",
    )
    img.save(path, quality=92)


def thermal_color(v: float) -> tuple[int, int, int]:
    stops = [
        (0.00, (15, 34, 97)),
        (0.16, (31, 91, 199)),
        (0.32, (38, 168, 218)),
        (0.48, (52, 185, 111)),
        (0.62, (233, 218, 45)),
        (0.75, (238, 128, 38)),
        (0.88, (196, 45, 49)),
        (1.00, (246, 235, 215)),
    ]
    v = max(0.0, min(1.0, v))
    for (p0, c0), (p1, c1) in zip(stops, stops[1:]):
        if p0 <= v <= p1:
            t = (v - p0) / (p1 - p0)
            return tuple(int(c0[i] + (c1[i] - c0[i]) * t) for i in range(3))
    return stops[-1][1]


def create_web_debug_image(path: Path):
    random.seed(90640)
    cols, rows = 32, 24
    temps = []
    for y in range(rows):
        row = []
        for x in range(cols):
            base = 23.5 + 0.4 * math.sin(x / 5) + 0.25 * math.cos(y / 4)
            hot = 7.8 * math.exp(-(((x - 12.5) ** 2) / 48 + ((y - 11.5) ** 2) / 28))
            warm_tail = 3.2 * math.exp(-(((x - 18.5) ** 2) / 70 + ((y - 13.5) ** 2) / 44))
            row.append(base + hot + warm_tail + random.uniform(-0.08, 0.08))
        temps.append(row)

    min_t = min(min(row) for row in temps)
    max_t = max(max(row) for row in temps)
    small = Image.new("RGB", (cols, rows))
    for y, row in enumerate(temps):
        for x, temp in enumerate(row):
            small.putpixel((x, y), thermal_color((temp - min_t) / (max_t - min_t)))
    heat = small.resize((980, 735), Image.Resampling.BICUBIC).filter(ImageFilter.GaussianBlur(radius=0.7))

    img = Image.new("RGB", (1600, 930), "#050505")
    draw = ImageDraw.Draw(img)
    title_font = pil_font(28, True)
    button_font = pil_font(20, True)
    text_font = pil_font(22)
    small_font = pil_font(18)

    draw.rectangle((0, 0, 1600, 72), fill="#111111", outline="#343434")
    draw.text((28, 22), "MLX90640 热成像相机", font=title_font, fill="white")
    controls = [
        ("开始检测", "#198754"),
        ("导入数据", "#198754"),
        ("下载CSV数据", "#0D6EFD"),
        ("清空记录", "#495057"),
        ("全屏", "#495057"),
    ]
    x = 360
    for text, fill in controls:
        bw = 120 if len(text) <= 4 else 160
        draw.rounded_rectangle((x, 17, x + bw, 56), radius=7, fill=fill, outline="#5C636A")
        tw = draw.textbbox((0, 0), text, font=button_font)[2]
        draw.text((x + (bw - tw) / 2, 26), text, font=button_font, fill="white")
        x += bw + 18
    draw.text((1240, 24), "FPS: 10", font=text_font, fill="#F8F9FA")

    draw.rectangle((315, 102, 1295, 837), fill="#0A0A0A")
    img.paste(heat, (315, 102))
    draw.rectangle((315, 102, 1295, 837), outline="#1F2937", width=2)

    bar_x, bar_y, bar_w, bar_h = 1325, 145, 46, 620
    for i in range(bar_h):
        color = thermal_color(1 - i / (bar_h - 1))
        draw.line((bar_x, bar_y + i, bar_x + bar_w, bar_y + i), fill=color)
    draw.rectangle((bar_x, bar_y, bar_x + bar_w, bar_y + bar_h), outline="#DEE2E6", width=1)
    draw.text((bar_x, bar_y - 36), f"{max_t:.1f} °C", font=text_font, fill="white")
    draw.text((bar_x, bar_y + bar_h + 12), f"{min_t:.1f} °C", font=text_font, fill="white")
    draw.text((bar_x, bar_y + bar_h + 48), "色温条", font=small_font, fill="#ADB5BD")

    draw.rectangle((0, 858, 1600, 930), fill="#111111", outline="#343434")
    footer = f"最低温度: {min_t:.1f} °C      最高温度: {max_t:.1f} °C      帧序号: 249      已记录: 249 帧"
    fw = draw.textbbox((0, 0), footer, font=text_font)[2]
    draw.text(((1600 - fw) / 2, 884), footer, font=text_font, fill="#F8F9FA")

    img.save(path, quality=92)


def create_or_copy_asset(source: Path | None, target_name: str, max_size=(1600, 1100)) -> Path:
    ASSET_DIR.mkdir(parents=True, exist_ok=True)
    target = ASSET_DIR / target_name
    if source and source.exists():
        try:
            with Image.open(source) as im:
                im = im.convert("RGB")
                im.thumbnail(max_size, Image.Resampling.LANCZOS)
                im.save(target, quality=88)
        except Exception:
            shutil.copyfile(source, target)
        return target
    if target.exists():
        return target
    placeholder = Image.new("RGB", (1200, 700), "#F4F6F9")
    draw = ImageDraw.Draw(placeholder)
    draw.text((60, 310), target_name, font=pil_font(42, True), fill="#52616B")
    placeholder.save(target, quality=88)
    return target


def prepare_assets() -> dict[str, Path]:
    ASSET_DIR.mkdir(parents=True, exist_ok=True)
    assets = {
        "esp32_front": create_or_copy_asset(SOURCE_IMAGE_DIR / "ESP32-C3正面.jpg", "esp32_front.jpg"),
        "esp32_back": create_or_copy_asset(SOURCE_IMAGE_DIR / "ESP32背面-接口信息.jpg", "esp32_back.jpg"),
        "mlx90640": create_or_copy_asset(SOURCE_IMAGE_DIR / "MLX90640红外摄像头.jpg", "mlx90640_sensor.jpg"),
        "wiring": create_or_copy_asset(SOURCE_IMAGE_DIR / "连线操作.jpg", "wiring.jpg"),
        "wifi_diagram": create_or_copy_asset(SOURCE_IMAGE_DIR / "ESP32 WiFi 热点与网页服务.png", "wifi_diagram.png"),
        "i2c_diagram": create_or_copy_asset(SOURCE_IMAGE_DIR / "I2C 与 MLX90640 引脚配置.png", "i2c_diagram.png"),
        "dev_env": create_or_copy_asset(SOURCE_IMAGE_DIR / "开发环境配置.png", "dev_environment.png"),
        "code_init": create_or_copy_asset(SOURCE_IMAGE_DIR / "MLX90640 初始化.png", "code_mlx90640_init.png"),
        "code_task": create_or_copy_asset(SOURCE_IMAGE_DIR / "温度数据采集任务.png", "code_capture_task.png"),
        "code_api": create_or_copy_asset(SOURCE_IMAGE_DIR / "网页温度数据接口.png", "code_thermal_api.png"),
        "web_waiting": create_or_copy_asset(LEGACY_ASSET_DIR / "webpage.jpg", "web_waiting.jpg"),
    }
    flow = ASSET_DIR / "system_flow.png"
    web_debug = ASSET_DIR / "web_heatmap_debug.jpg"
    create_flow_diagram(flow)
    create_web_debug_image(web_debug)
    assets["flow"] = flow
    assets["web_debug"] = web_debug
    return assets


def set_run_font(run, size: float | None = None, bold: bool | None = None, color: RGBColor | None = None, name: str = FONT_NAME):
    run.font.name = name
    run._element.rPr.rFonts.set(qn("w:eastAsia"), name)
    run._element.rPr.rFonts.set(qn("w:ascii"), name)
    run._element.rPr.rFonts.set(qn("w:hAnsi"), name)
    if size is not None:
        run.font.size = Pt(size)
    if bold is not None:
        run.bold = bold
    if color is not None:
        run.font.color.rgb = color


def shade_cell(cell, fill: str):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), fill)
    tc_pr.append(shd)


def set_cell_text(cell, text: str, bold: bool = False, size: int = 10.5, color: RGBColor = BODY):
    cell.text = ""
    p = cell.paragraphs[0]
    p.paragraph_format.space_after = Pt(0)
    run = p.add_run(text)
    set_run_font(run, size=size, bold=bold, color=color)
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER


def set_table_borders(table, color: str = "D7DEE8", size: str = "8"):
    tbl = table._tbl
    tbl_pr = tbl.tblPr
    borders = tbl_pr.first_child_found_in("w:tblBorders")
    if borders is None:
        borders = OxmlElement("w:tblBorders")
        tbl_pr.append(borders)
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        tag = "w:" + edge
        element = borders.find(qn(tag))
        if element is None:
            element = OxmlElement(tag)
            borders.append(element)
        element.set(qn("w:val"), "single")
        element.set(qn("w:sz"), size)
        element.set(qn("w:space"), "0")
        element.set(qn("w:color"), color)


def paragraph_border_bottom(paragraph, color: str = "2E74B5", size: str = "12"):
    p_pr = paragraph._p.get_or_add_pPr()
    p_bdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), size)
    bottom.set(qn("w:space"), "6")
    bottom.set(qn("w:color"), color)
    p_bdr.append(bottom)
    p_pr.append(p_bdr)


def setup_document() -> Document:
    doc = Document()
    section = doc.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(1.8)
    section.bottom_margin = Cm(1.7)
    section.left_margin = Cm(1.8)
    section.right_margin = Cm(1.8)
    section.header_distance = Cm(0.8)
    section.footer_distance = Cm(0.8)

    styles = doc.styles
    normal = styles["Normal"]
    normal.font.name = FONT_NAME
    normal._element.rPr.rFonts.set(qn("w:eastAsia"), FONT_NAME)
    normal.font.size = Pt(10.5)
    normal.font.color.rgb = BODY
    normal.paragraph_format.space_after = Pt(6)
    normal.paragraph_format.line_spacing = 1.15

    for style_name, size, color in [
        ("Heading 1", 18, ACCENT),
        ("Heading 2", 14, ACCENT),
        ("Heading 3", 12, ACCENT_DARK),
    ]:
        style = styles[style_name]
        style.font.name = FONT_NAME
        style._element.rPr.rFonts.set(qn("w:eastAsia"), FONT_NAME)
        style.font.size = Pt(size)
        style.font.bold = True
        style.font.color.rgb = color
        style.paragraph_format.space_before = Pt(10)
        style.paragraph_format.space_after = Pt(6)

    footer = section.footer.paragraphs[0]
    footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = footer.add_run("微机接口与技术期末大作业 | ESP32-C3 + MLX90640 热成像相机系统")
    set_run_font(run, size=9, color=MUTED)

    props = doc.core_properties
    props.author = "rdfrdream"
    props.title = "基于 ESP32-C3 与 MLX90640 的热成像相机系统开发流程说明"
    props.subject = "微机接口与技术期末大作业"
    return doc


def add_title_page(doc: Document):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(18)
    p.paragraph_format.space_after = Pt(6)
    r = p.add_run("基于 ESP32-C3 与 MLX90640 的\n热成像相机系统开发流程说明")
    set_run_font(r, size=25, bold=True, color=ACCENT_DARK)
    paragraph_border_bottom(p, "2E74B5", "16")

    subtitle = doc.add_paragraph()
    subtitle.paragraph_format.space_after = Pt(18)
    r = subtitle.add_run("嵌入式采集 · WiFi Web 服务 · 热成像可视化 · 数据记录")
    set_run_font(r, size=13, color=MUTED)

    add_callout(
        doc,
        "项目定位",
        "本项目围绕 MLX90640 热红外阵列传感器与 ESP32-C3 开发板，完成从 I2C 采集、无线接入、HTTP 数据接口到浏览器热图显示的完整软件链路。当前报告用于记录开发流程、仓库文件、硬件连接方式和调试验证方法。",
    )

    facts = [
        ("主控平台", "ESP32-C3，集成 WiFi/BLE，可承担主控与 Web Server"),
        ("热成像器件", "MLX90640，32 x 24 红外阵列，I2C 地址 0x33"),
        ("开发环境", "VS Code + PlatformIO + ESP-IDF"),
        ("仓库地址", "https://github.com/1909899270h-spec/esp32c3-mlx90640-thermal-camera"),
    ]
    for label, value in facts:
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Cm(0.2)
        p.paragraph_format.space_after = Pt(5)
        r = p.add_run(label + "：")
        set_run_font(r, size=11.5, bold=True, color=ACCENT)
        r = p.add_run(value)
        set_run_font(r, size=11.5, color=BODY)

    doc.add_page_break()


def add_callout(doc: Document, title: str, text: str):
    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(table, "D7DEE8", "8")
    cell = table.cell(0, 0)
    shade_cell(cell, CALLOUT_FILL)
    p = cell.paragraphs[0]
    p.paragraph_format.space_after = Pt(4)
    r = p.add_run(title)
    set_run_font(r, size=12, bold=True, color=ACCENT_DARK)
    p2 = cell.add_paragraph()
    p2.paragraph_format.space_after = Pt(0)
    r = p2.add_run(text)
    set_run_font(r, size=10.5, color=BODY)
    doc.add_paragraph().paragraph_format.space_after = Pt(2)


def add_heading(doc: Document, text: str, level: int = 1):
    p = doc.add_heading(text, level=level)
    # Word heading styles keep with the next paragraph by default. For this
    # image-heavy report that can create blank pages, so rdfrdream disables it.
    p.paragraph_format.keep_with_next = False
    return p


def add_para(doc: Document, text: str, size: float = 10.5, color: RGBColor = BODY, bold: bool = False):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(6)
    r = p.add_run(text)
    set_run_font(r, size=size, bold=bold, color=color)
    return p


def add_bullets(doc: Document, items: list[str]):
    for item in items:
        p = doc.add_paragraph(style="List Bullet")
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.left_indent = Cm(0.55)
        p.paragraph_format.first_line_indent = Cm(-0.25)
        r = p.add_run(item)
        set_run_font(r, size=10.5, color=BODY)


def add_numbered(doc: Document, items: list[str]):
    for item in items:
        p = doc.add_paragraph(style="List Number")
        p.paragraph_format.space_after = Pt(4)
        r = p.add_run(item)
        set_run_font(r, size=10.5, color=BODY)


def add_figure(doc: Document, path: Path, caption: str, width: float = 6.4):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(2)
    p.add_run().add_picture(str(path), width=Inches(width))
    cp = doc.add_paragraph()
    cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cp.paragraph_format.space_after = Pt(8)
    r = cp.add_run(caption)
    set_run_font(r, size=9.5, color=MUTED)


def add_image_grid(doc: Document, items: list[tuple[Path, str]], width: float = 3.05):
    table = doc.add_table(rows=0, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(table, "FFFFFF", "0")
    for i in range(0, len(items), 2):
        row = table.add_row()
        for j in range(2):
            cell = row.cells[j]
            cell.vertical_alignment = WD_ALIGN_VERTICAL.TOP
            if i + j >= len(items):
                continue
            img_path, caption = items[i + j]
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.add_run().add_picture(str(img_path), width=Inches(width))
            cp = cell.add_paragraph()
            cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = cp.add_run(caption)
            set_run_font(r, size=9, color=MUTED)
    doc.add_paragraph().paragraph_format.space_after = Pt(4)


def add_code_block(doc: Document, text: str):
    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(table, "D7DEE8", "6")
    cell = table.cell(0, 0)
    shade_cell(cell, "F6F8FA")
    p = cell.paragraphs[0]
    p.paragraph_format.space_after = Pt(0)
    for line_no, line in enumerate(text.splitlines()):
        if line_no:
            p.add_run().add_break()
        r = p.add_run(line)
        set_run_font(r, size=9.2, color=BODY, name=MONO_FONT)
    doc.add_paragraph().paragraph_format.space_after = Pt(3)


def add_wiring_table(doc: Document):
    rows = [
        ("MLX90640 SDA", "ESP32-C3 GPIO4", "I2C 数据线", "5kΩ 上拉到 3.3V"),
        ("MLX90640 SCL", "ESP32-C3 GPIO5", "I2C 时钟线", "5kΩ 上拉到 3.3V"),
        ("MLX90640 VDD", "ESP32-C3 3.3V", "传感器供电", "建议 10uF 去耦电容"),
        ("MLX90640 GND", "ESP32-C3 GND", "公共地", "必须与主控共地"),
    ]
    table = doc.add_table(rows=1, cols=4)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    set_table_borders(table)
    headers = ("传感器端", "ESP32-C3 端", "作用", "备注")
    widths = [1.55, 1.55, 1.35, 2.0]
    for idx, header in enumerate(headers):
        cell = table.cell(0, idx)
        cell.width = Inches(widths[idx])
        shade_cell(cell, LIGHT_FILL)
        set_cell_text(cell, header, bold=True, size=10, color=ACCENT_DARK)
    for row_data in rows:
        row = table.add_row()
        for idx, value in enumerate(row_data):
            cell = row.cells[idx]
            cell.width = Inches(widths[idx])
            set_cell_text(cell, value, size=9.5, color=BODY)
    doc.add_paragraph().paragraph_format.space_after = Pt(4)


def section_system_flow(doc: Document, assets: dict[str, Path]):
    add_heading(doc, "1. 系统总体链路", 1)
    add_para(
        doc,
        "系统的核心不是单独读取一个传感器，而是把“采集、网络、网页、记录”连成一条可演示的数据链路。ESP32-C3 在其中同时承担主控、无线通信和轻量级服务器角色。",
    )
    add_figure(doc, assets["flow"], "图 1  从 MLX90640 到浏览器热图的完整数据链路", width=6.9)
    add_bullets(
        doc,
        [
            "传感器侧：MLX90640 输出 32 x 24 阵列温度数据，通过 I2C 与主控通信。",
            "主控侧：ESP32-C3 完成采样调度、温度计算、数据缓存和 HTTP 接口返回。",
            "显示侧：浏览器网页通过 JSON 接口刷新 Canvas 热图，并支持色温条与 CSV 数据记录。",
        ],
    )
    doc.add_page_break()


def section_hardware(doc: Document, assets: dict[str, Path]):
    add_heading(doc, "2. 硬件组成与连接", 1)
    add_para(
        doc,
        "硬件搭建采用面包板与杜邦线完成，便于在调试阶段反复调整接线。系统最小硬件包括 ESP32-C3 开发板、MLX90640 红外阵列传感器、I2C 上拉电阻和供电去耦电容。",
    )
    add_image_grid(
        doc,
        [
            (assets["esp32_front"], "ESP32-C3 开发板正面"),
            (assets["mlx90640"], "MLX90640 裸传感器"),
            (assets["esp32_back"], "开发板背面引脚标识"),
            (assets["wiring"], "面包板接线过程"),
        ],
    )
    doc.add_page_break()
    add_heading(doc, "核心接线", 2)
    add_wiring_table(doc)
    add_callout(
        doc,
        "电阻与电容的作用",
        "SDA/SCL 是开漏/开集电极风格的总线信号，需要通过上拉电阻保持高电平；10uF 电容用于降低供电波动，减少传感器读数异常和 I2C 通信不稳定。",
    )
    add_figure(doc, assets["i2c_diagram"], "图 2  I2C 与 MLX90640 引脚配置说明", width=5.6)
    doc.add_page_break()


def section_dev_env_repo(doc: Document, assets: dict[str, Path]):
    add_heading(doc, "3. 开发环境与仓库结构", 1)
    add_para(
        doc,
        "开发链路以 VS Code + PlatformIO 为主。PlatformIO 负责工程依赖、ESP-IDF 框架、编译、串口烧录、SPIFFS 文件系统上传和串口监视，适合把嵌入式代码与网页资源放在同一个工程中管理。",
    )
    add_figure(doc, assets["dev_env"], "图 3  VS Code + PlatformIO 的工程配置与烧录入口", width=5.7)
    add_heading(doc, "仓库关键文件", 2)
    repo_tree = """esp32c3-mlx90640-thermal-camera/
├─ platformio.ini                  # PlatformIO 工程、板卡、烧录与文件系统配置
├─ src/main.c                      # ESP32-C3 主程序：I2C、WiFi AP、HTTP Server、接口逻辑
├─ data/web_script.js              # 前端脚本：热图绘制、色温条、按钮交互、CSV 导出
├─ lib/MLX90640/                   # MLX90640 API 与参数计算代码
├─ tools/mock_thermal_server.py    # 无硬件调试用的本地网页与数据生成服务
├─ partitions.csv                  # Flash 分区，支持 SPIFFS 网页资源
├─ README.md                       # 仓库说明、接线、烧录、演示入口
└─ deliverables/                   # 答辩 PPT、开发流程 DOCX/PDF 与生成脚本"""
    add_code_block(doc, repo_tree)
    add_bullets(
        doc,
        [
            "固件端网页的 HTML/CSS 由 src/main.c 中的字符串返回，前端交互逻辑主要放在 data/web_script.js。",
            "调试端 tools/mock_thermal_server.py 可在没有传感器稳定数据时验证热图、按钮、CSV 下载等网页功能。",
            "README 与 deliverables 记录了硬件连接、烧录流程、文档和演示材料，便于后续开源维护。",
        ],
    )


def section_software(doc: Document, assets: dict[str, Path]):
    add_heading(doc, "4. 软件架构与关键代码", 1)
    add_para(
        doc,
        "软件部分按“采集任务 - 数据缓存 - HTTP 接口 - 网页渲染”的方式组织。这样可以把传感器采样与网页请求解耦，浏览器访问数据时读取最近一帧，而不是每次请求都阻塞等待传感器。",
    )
    add_heading(doc, "关键代码截图", 2)
    add_figure(doc, assets["code_init"], "图 4  MLX90640 初始化：I2C 参数、传感器参数读取与初始化状态判断", width=5.7)
    add_figure(doc, assets["code_task"], "图 5  温度采集任务：持续读取帧数据、计算 To 数组并更新共享缓存", width=5.9)
    add_figure(doc, assets["code_api"], "图 6  网页数据接口：将最新温度数组编码为 JSON 并返回给浏览器", width=5.9)
    doc.add_page_break()


def section_web(doc: Document, assets: dict[str, Path]):
    add_heading(doc, "5. 网页热图与局域网访问", 1)
    add_para(
        doc,
        "网页端由 HTML、CSS 与 JavaScript 构成：HTML 定义按钮和画布，CSS 控制深色界面与控制栏样式，JavaScript 负责定时请求数据、Canvas 热图绘制、色温条显示、开始检测、导入数据、下载 CSV 和全屏显示。",
    )
    add_figure(doc, assets["web_debug"], "图 7  网页热图调试效果：用于验证 Canvas、色温条、按钮交互和数据记录逻辑", width=6.7)
    add_heading(doc, "为什么只能在局域网访问", 2)
    add_bullets(
        doc,
        [
            "AP 模式下，ESP32-C3 自己建立 WiFi 热点，手机或电脑连接该热点后访问 192.168.4.1。",
            "STA 模式下，ESP32-C3 连接路由器，浏览器访问的是路由器分配给 ESP32 的局域网 IP。",
            "这些地址属于局域网私有地址，不是公网地址；手机切换到其他网络后无法直接访问，这是网络结构原因，不是网页代码错误。",
        ],
    )
    doc.add_page_break()


def section_debug(doc: Document):
    add_heading(doc, "6. 调试与问题定位思路", 1)
    add_para(
        doc,
        "调试时不能只看网页是否出现热图，需要分层判断问题发生在供电、总线、数据计算还是网页显示。下面的三层判断可以覆盖大部分异常情况。",
    )
    add_heading(doc, "第一层：硬件层", 2)
    add_bullets(
        doc,
        [
            "确认 ESP32-C3 Type-C 供电正常，电脑设备管理器能识别 COM 口。",
            "确认 MLX90640 的 VDD 接 3.3V、GND 共地，杜邦线没有松动或短接。",
            "确认 SDA/SCL 上拉电阻和 10uF 去耦电容接入合理。",
        ],
    )
    add_heading(doc, "第二层：总线层", 2)
    add_bullets(
        doc,
        [
            "通过 I2C scan 或串口日志判断是否能扫描到 MLX90640 地址 0x33。",
            "如果扫描不到，优先检查 GPIO4/GPIO5 是否接反、上拉是否有效、传感器引脚方向是否一致。",
        ],
    )
    add_heading(doc, "第三层：数据层与网页层", 2)
    add_bullets(
        doc,
        [
            "观察网页 min/max/avg 是否随热源变化，若始终为空或固定值，说明采集链路仍未稳定。",
            "若数据正常但图像不动，检查 /thermal-data JSON、Canvas 绘制和前端刷新逻辑。",
            "CSV 下载只记录已接收到的数据帧，不应伪造未经验证的真实采样结果。",
        ],
    )
    doc.add_page_break()


def section_status_future(doc: Document, assets: dict[str, Path]):
    add_heading(doc, "7. 当前完成情况与后续优化", 1)
    add_callout(
        doc,
        "当前状态",
        "目前已完成 ESP32-C3 Web 服务、网页热图显示、数据接口、CSV 导出逻辑、调试数据链路和硬件接线方案；MLX90640 实时采集链路正在进行稳定性验证。",
    )
    add_heading(doc, "已完成内容", 2)
    add_bullets(
        doc,
        [
            "PlatformIO 工程、ESP32-C3 固件编译、串口烧录与 SPIFFS 文件系统上传流程已经建立。",
            "网页端支持热图可视化、色温条、开始检测、导入调试数据、全屏显示与 CSV 下载。",
            "GitHub 仓库已整理代码、README、答辩 PPT、开发流程文档和生成脚本。",
        ],
    )
    add_heading(doc, "后续优化方向", 2)
    add_bullets(
        doc,
        [
            "硬件优化：PCB、传感器固定、外壳结构、供电去耦与接口防松。",
            "算法优化：异常点滤波、帧间平滑、热点追踪、发射率与环境温度修正。",
            "功能拓展：温度极值触发 LED 或报警器，加入阈值报警和实验记录。",
            "网络拓展：通过上位服务器、云端转发或内网穿透，使数据不只局限于局域网查看。",
            "识别拓展：结合红外目标识别算法，进一步发展为自动识别或设备异常检测装置。",
            "校准系统：加入已知温度参考源或两点标定流程，对偏移和增益进行软件补偿；该方案可行，但需要稳定参考温度和传感器固定结构配合。",
        ],
    )


def build_docx():
    assets = prepare_assets()
    doc = setup_document()
    add_title_page(doc)
    section_system_flow(doc, assets)
    section_hardware(doc, assets)
    section_dev_env_repo(doc, assets)
    section_software(doc, assets)
    section_web(doc, assets)
    section_debug(doc)
    section_status_future(doc, assets)
    doc.save(DOCX_PATH)
    print(DOCX_PATH)


if __name__ == "__main__":
    build_docx()
