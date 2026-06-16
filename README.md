# 基于 ESP32-C3 与 MLX90640 的热成像相机系统设计

维护者：**rdfrdream**

本项目是“微机接口与技术 / 微机原理”课程大作业的软件工程整理版。系统以 ESP32-C3 为主控，连接 MLX90640 32x24 热红外阵列传感器，通过 I2C 采集温度矩阵，并由 ESP32-C3 建立 WiFi 热点与 HTTP Web Server，使手机或电脑浏览器能够查看热成像画面、色温条和温度数据。

## 项目主线

```text
MLX90640 热红外阵列
        ↓ I2C
ESP32-C3 主控
        ↓ WiFi AP / HTTP Server
浏览器网页热图
        ↓
CSV 数据记录与导出
```

## 当前完成情况

- 已完成 ESP32-C3 PlatformIO 工程整理。
- 已完成 MLX90640 I2C 接口代码、初始化流程和温度帧读取逻辑。
- 已完成 ESP32-C3 WiFi AP 模式与 HTTP Web Server。
- 已完成网页端 Canvas 热图、色温条、平滑显示、开始检测、导入数据和 CSV 导出逻辑。
- 已完成本地 Python 演示服务器，用于无传感器或硬件不稳定时展示热图交互。
- MLX90640 实时采集链路仍需要结合实物接线继续做稳定性验证。

## 硬件连接

| MLX90640 引脚 | ESP32-C3 引脚 | 说明 |
| --- | --- | --- |
| VDD | 3.3V | 传感器供电，避免接 5V |
| GND | GND | 共地 |
| SDA | GPIO4 | I2C 数据线 |
| SCL | GPIO5 | I2C 时钟线 |

建议：

- SDA 与 SCL 分别通过约 5kΩ 电阻上拉到 3.3V。
- VDD 与 GND 之间可并联 10uF 电容，用于供电去耦。
- 裸 MLX90640 传感器引脚较细，面包板连接容易松动，建议后续使用转接板或焊接固定。

## 软件架构

| 模块 | 作用 |
| --- | --- |
| `src/main.c` | ESP32-C3 主程序，负责 I2C、MLX90640、WiFi AP、HTTP 接口 |
| `data/web_script.js` | 烧录到 SPIFFS 的网页脚本，负责热图绘制、色温条、CSV 导出 |
| `lib/MLX90640` | MLX90640 参数解析与温度计算库 |
| `tools/mock_thermal_server.py` | 本地演示网页服务器，可生成平滑仿真热源数据 |
| `platformio.ini` | PlatformIO 工程配置 |
| `partitions.csv` | ESP32 Flash 分区配置，包含 SPIFFS |

## 开发环境

- VS Code
- PlatformIO
- ESP-IDF framework
- ESP32-C3 DevKit
- Python 3，用于本地仿真演示

建议将工程放在不含空格的路径下，例如：

```text
H:\esp32c3-mlx90640-thermal-camera
```

PlatformIO / ESP-IDF 在包含空格的路径中可能出现 CMake 配置失败。

PlatformIO 配置位于：

```ini
[env:esp32-c3-devkitm-1]
platform = espressif32@6.9.0
board = esp32-c3-devkitm-1
framework = espidf
monitor_speed = 115200
upload_speed = 921600
board_build.filesystem = spiffs
```

## 编译与烧录

进入项目根目录：

```powershell
cd esp32c3-mlx90640-thermal-camera
```

编译主程序：

```powershell
pio run -e esp32-c3-devkitm-1
```

烧录主程序：

```powershell
pio run -e esp32-c3-devkitm-1 -t upload
```

上传网页文件系统：

```powershell
pio run -e esp32-c3-devkitm-1 -t uploadfs
```

打开串口监视：

```powershell
pio device monitor -e esp32-c3-devkitm-1
```

如果烧录时卡在 `Connecting...`，按住 ESP32-C3 的 `BOOT/FLASH`，再轻按 `RST/RESET`，出现 `Writing at...` 后松开 `BOOT/FLASH`。

## 使用方式

1. ESP32-C3 烧录主程序和 SPIFFS 文件系统。
2. 上电后，ESP32-C3 建立热点，名称类似 `IRCAM-XXXX`。
3. 手机或电脑连接该热点。
4. 浏览器访问：

```text
http://192.168.4.1
```

网页功能：

- `导入数据`：使用浏览器端仿真热源，演示热图、色温条和 CSV 导出。
- `开始检测`：访问 ESP32 的 `/thermal-data` 接口，读取真实 MLX90640 数据。
- `下载CSV数据`：导出当前网页记录的温度帧。
- `全屏`：用于课堂展示或视频录制。

## 调试接口

| 地址 | 作用 |
| --- | --- |
| `/` | 热成像网页 |
| `/web_script.js` | 前端脚本文件 |
| `/thermal-data` | 返回 768 个温度点，32x24 |
| `/status` | 返回传感器初始化、帧计数和 I2C 状态 |
| `/i2c-scan` | 扫描 I2C 总线，MLX90640 应识别为 `0x33` |

调试建议：

1. 硬件层：确认 ESP32-C3 上电、COM 口识别、3.3V/GND 正确。
2. 总线层：访问 `/i2c-scan` 或查看串口日志，确认是否扫描到 `0x33`。
3. 数据层：查看网页最低温度、最高温度和热图是否随热源变化。

## 本地仿真演示

如果传感器暂时未稳定连接，可以运行本地演示服务器：

```powershell
python tools/mock_thermal_server.py
```

也可以双击：

```text
Run_Mock_Thermal_Web.cmd
```

浏览器打开：

```text
http://127.0.0.1:8000
```

本地仿真只用于展示网页交互和数据记录流程，不代表真实 MLX90640 测量结果。

## 开源说明

本项目使用 MIT License。项目整理、文档与课程展示说明由 **rdfrdream** 维护。部分 MLX90640 驱动与基础代码参考开源项目和芯片示例库，详见 `NOTICE.md`。

## GitHub 发布参考

如果本机已安装并登录 GitHub CLI，可以在项目根目录执行：

```powershell
gh repo create rdfrdream/esp32c3-mlx90640-thermal-camera --public --source . --remote origin --push
```

如果已经在 GitHub 网页端创建了空仓库，也可以执行：

```powershell
git remote add origin https://github.com/rdfrdream/esp32c3-mlx90640-thermal-camera.git
git push -u origin main
```
