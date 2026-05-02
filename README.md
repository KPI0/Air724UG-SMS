# Air724UG-SMS
# 📩 短信监听系统（LUAT Modem 自动识别版）

一款基于 **Python + Tkinter** 的 **4G / LTE 模组短信接收与预警显示程序**，  
专为 **LUAT（合宙）系列 Modem** 多串口环境设计，支持 **自动识别 Modem 串口、自动重连、无人值守、离线运行**。

适用于：  
- 预警短信接收  
- 值班室 / 监控室大屏展示  
- 工业上位机  
- 无人值守终端
- 内网部署  

---
### [USB 驱动安装](https://docs.openluat.com/air724ug/common/usb_drv/)
---

## 📦 依赖安装

```bash
pip install pyserial pystray pillow pyttsx3
```

## 🚀 源码运行
```bash
py sms.pyw
```
## 🏗 打包为 EXE

安装 PyInstaller：
```bash
pip install pyinstaller
```
然后在项目目录下执行以下命令：
```bash
pyinstaller ^
  --noconfirm ^
  --onefile ^
  --windowed ^
  --name "sms" ^
  --icon=icon.ico ^
  --add-data "icon.ico;." ^
  sms.pyw
```
生成的 exe 文件位于：
```bash
dist/sms.exe
```
## 📁 项目结构
```
Air724UG-SMS/
│
├── sms.pyw    # 主程序入口
├── icon.ico   # 应用图标
├── config.ini # 软件配置文件（首次运行自动创建）
├── sms_logs/  # 短信日志存储目录
└── tts/       # 语音播报缓存目录
```
## ✨ 主要特性

### 🔌 串口与设备
-  **自动识别 LUAT Modem 串口**
  - 只连接 `LUAT USB Device X Modem`
  - 自动忽略 AT / Diag / MOS / NPI 等非业务串口
-  **串口掉线自动重连**
  - USB 拔插
  - 设备重启
  - COM 号变化
-  **手动模式**
  - 可手动锁定指定 COM 口

---

### 🔒 安全与离线
-  **信息保密设计**
  - 本软件所有数据（短信内容、日志、配置）均保存在本地
  - 不上传服务器、不经过第三方平台
  - 适用于涉密环境或内网部署
-  **完全离线运行**
  - 软件可在无网络环境下正常运行
  - 无需联网即可完成短信接收、显示、播报等全部功能

---

### 📩 短信接收与显示
-  **每天 0 点自动清空显示窗口**
-  **自定义短信字体颜色字号**
-  **关键词过滤**

---

### 📢 提醒与日志
-  **自动创建日志目录**
-  **短信日志按天保存**
-  **短信日志定时清理**
-  **短信播报提示语音**
-  **语音播报自定义**
-  **语音播报开关**
-  **短信弹窗开关**

---

### 🖥 GUI 体验
-  **实时串口连接状态**
-  **在线检测新版本**
-  **最小化到托盘**
-  **串口调试窗口**
-  **桌面快捷方式**
-  **开机自启**
-  **程序多开**

---

## 🖼 界面展示
![](https://github.com/KPI0/Air724UG-SMS/blob/main/png/1.png)   
![](https://github.com/KPI0/Air724UG-SMS/blob/main/png/2.png)   
![](https://github.com/KPI0/Air724UG-SMS/blob/main/png/3.png)   
![](https://github.com/KPI0/Air724UG-SMS/blob/main/png/4.png)   
![](https://github.com/KPI0/Air724UG-SMS/blob/main/png/5.png)   
![](https://github.com/KPI0/Air724UG-SMS/blob/main/png/6.png)   
