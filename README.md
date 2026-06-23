# 基于 ESP32-C3 与 MLX90640 的热成像相机系统设计

维护者：**rdfrdream**

本项目用于“微机接口与技术 / 微机原理”课程期末大作业。系统以 ESP32-C3 为主控，连接 MLX90640 32 x 24 热红外阵列传感器，通过 I2C 读取温度矩阵，并由 ESP32-C3 建立 WiFi 热点与 HTTP Web Server，使手机或电脑浏览器能够查看热图、色温条、实时温度数据和 CSV 记录。

## 1. 项目功能

- ESP32-C3 通过 I2C 与 MLX90640 通信，目标地址为 `0x33`。
- ESP32-C3 可建立 WiFi AP 热点，浏览器访问 `http://192.168.4.1` 查看网页。
- 网页端支持热图可视化、色温条、开始检测、导入演示数据、全屏展示和 CSV 下载。
- 提供本地 Python 演示服务器，在硬件采集链路不稳定时用于展示网页交互和数据记录流程。
- 提供答辩 PPT、开发流程 PDF/DOCX、成果提交报告 PDF/DOCX。

## 2. 系统主线

```text
MLX90640 热红外阵列传感器
        ↓ I2C 总线（SDA / SCL）
ESP32-C3 主控
        ↓ WiFi AP / HTTP Web Server
浏览器网页热图显示
        ↓
CSV 温度数据记录与导出
```

## 3. 硬件准备与接线

### 3.1 硬件清单

- ESP32-C3 开发板
- MLX90640 热红外阵列传感器
- 面包板与杜邦线
- 5kΩ 上拉电阻 2 个
- 10uF 电容 1 个
- Type-C 数据线

### 3.2 接线表

| MLX90640 引脚 | ESP32-C3 引脚 | 说明 |
| --- | --- | --- |
| VDD | 3.3V | 传感器供电，避免接 5V |
| GND | GND | 必须共地 |
| SDA | GPIO4 | I2C 数据线 |
| SCL | GPIO5 | I2C 时钟线 |

建议将 SDA 与 SCL 分别通过约 `5kΩ` 电阻上拉到 `3.3V`，并在 VDD 与 GND 之间并联 `10uF` 电容，用于改善供电稳定性。裸 MLX90640 引脚较细，面包板和杜邦线连接容易松动，后续可以使用转接板、焊接或 PCB 固定方案提高稳定性。

## 4. 开发环境

- VS Code
- PlatformIO
- ESP-IDF framework
- Python 3（用于本地演示服务器）
- Windows 设备管理器中可识别 ESP32-C3 串口，例如 `COM5`

建议将工程放在不含空格和特殊字符的路径下，例如：

```powershell
H:\esp32c3-mlx90640-thermal-camera
```

PlatformIO 主要配置位于 `platformio.ini`：

```ini
[env:esp32-c3-devkitm-1]
platform = espressif32@6.9.0
board = esp32-c3-devkitm-1
framework = espidf
monitor_speed = 115200
upload_speed = 921600
board_build.filesystem = spiffs
```

## 5. 真实硬件运行方式

### 5.1 编译固件

在项目根目录执行：

```powershell
pio run -e esp32-c3-devkitm-1
```

### 5.2 烧录主程序

```powershell
pio run -e esp32-c3-devkitm-1 -t upload
```

若烧录停在 `Connecting...`，可按住开发板 `BOOT/FLASH` 键，再轻按 `RST/RESET` 键，出现 `Writing at...` 后松开 `BOOT/FLASH`。

### 5.3 上传网页文件系统

网页脚本位于 `data/web_script.js`，需要上传到 SPIFFS：

```powershell
pio run -e esp32-c3-devkitm-1 -t uploadfs
```

### 5.4 打开串口监视

```powershell
pio device monitor -e esp32-c3-devkitm-1
```

串口波特率为 `115200`。若串口被占用，需要先关闭其他 Monitor、串口助手或 WPS/VS Code 中正在占用 COM 口的任务。

### 5.5 浏览器访问

主程序运行后，ESP32-C3 会建立 WiFi 热点，名称通常类似：

```text
IRCAM-XXXX
```

手机或电脑连接该热点后，在浏览器访问：

```text
http://192.168.4.1
```

网页中点击：

- `开始检测`：从 ESP32-C3 的 `/thermal-data` 接口读取真实传感器数据。
- `导入数据`：使用网页端演示数据，用于无硬件或硬件不稳定时展示热图效果。
- `下载CSV数据`：导出当前网页记录的温度数据。
- `全屏`：用于录制演示视频或课堂展示。

## 6. 本地演示运行方式

当 MLX90640 接线不稳定、暂时无法获得连续真实数据时，可以运行本地 Python 演示服务。该方式用于验证网页端热图、色温条、按钮交互和 CSV 导出逻辑，不代表真实传感器测量结果。

### 6.1 命令行启动

```powershell
python tools/mock_thermal_server.py
```

浏览器打开：

```text
http://127.0.0.1:8000
```

### 6.2 脚本启动

也可以双击：

```text
Run_Mock_Thermal_Web.cmd
```

停止本地演示服务可双击：

```text
Stop_Mock_Thermal_Web.cmd
```

## 7. 调试接口

| 地址 | 作用 |
| --- | --- |
| `/` | 热成像网页 |
| `/web_script.js` | 前端脚本文件 |
| `/thermal-data` | 返回 32 x 24 共 768 个温度点 |
| `/status` | 查看初始化状态、帧计数和 I2C 状态 |
| `/i2c-scan` | 扫描 I2C 总线，MLX90640 应识别为 `0x33` |

推荐排查顺序：

1. 硬件层：确认 ESP32-C3 上电、电脑识别 COM 口、3.3V/GND 接线正确。
2. 总线层：访问 `/i2c-scan` 或查看串口日志，确认是否扫描到 `0x33`。
3. 数据层：观察网页最低温度、最高温度和热图是否随热源变化。
4. 网页层：若数据正常但热图不动，检查 `/thermal-data` JSON 与 Canvas 刷新逻辑。

## 8. 目录结构

```text
esp32c3-mlx90640-thermal-camera/
├─ platformio.ini                  # PlatformIO 工程配置
├─ src/main.c                      # ESP32-C3 主程序
├─ data/web_script.js              # 网页前端脚本
├─ lib/MLX90640/                   # MLX90640 驱动与算法代码
├─ tools/mock_thermal_server.py    # 本地演示服务器
├─ partitions.csv                  # Flash 分区配置
├─ README.md                       # 项目说明与运行方式
├─ deliverables/                   # 提交报告、PPT、开发流程文档
└─ LICENSE / NOTICE.md             # 开源许可与来源说明
```

## 9. 当前完成情况

目前已完成 ESP32-C3 Web 服务、网页热图显示、数据接口、CSV 导出逻辑、调试数据链路和硬件接线方案。MLX90640 实时采集链路已经完成接线方案和软件接口，但由于当前采用裸传感器、面包板和杜邦线临时连接，实地采集稳定性仍需继续验证。

在最终展示中，可以采用“真实硬件搭建照片 + 前期实验视频 + 网页端演示数据”的方式说明系统流程；后续通过转接板、焊接固定、PCB、供电去耦和传感器校准进一步提高实物稳定性。

## 10. 成果文件

`deliverables/` 目录中保留了提交与展示材料：

| 文件 | 说明 |
| --- | --- |
| `基于ESP32-C3与MLX90640的热成像相机系统项目报告.pdf` | 完整项目报告 PDF |
| `基于ESP32-C3与MLX90640的热成像相机系统项目报告.docx` | 完整项目报告 Word 源文件 |
| `基于ESP32-C3与MLX90640的热成像相机系统设计_答辩PPT.pptx` | 答辩 PPT |
| `基于ESP32-C3与MLX90640的热成像相机系统开发流程.pdf` | 开发流程说明 PDF |
| `基于ESP32-C3与MLX90640的热成像相机系统开发流程.docx` | 开发流程说明 Word |

## 11. 开源说明

本项目使用 MIT License。项目整理、文档与课程展示说明由 **rdfrdream** 维护。部分 MLX90640 驱动与基础代码参考开源项目和芯片示例库，详见 `NOTICE.md`。
