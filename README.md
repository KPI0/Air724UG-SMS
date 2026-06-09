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
pip install -r requirements.txt
```

## 🚀 源码运行
```bash
py sms\sms.pyw
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
  --paths "sms" ^
  --hidden-import "pyttsx3.drivers.sapi5" ^
  --hidden-import "pythoncom" ^
  --hidden-import "pywintypes" ^
  --hidden-import "win32com" ^
  --hidden-import "win32com.client" ^
  --hidden-import "sms_app" ^
  --hidden-import "sms_core" ^
  --hidden-import "sms_ui" ^
  --collect-submodules "win32com" ^
  --collect-submodules "websockets" ^
  --collect-submodules "sms_app" ^
  --collect-submodules "sms_core" ^
  --collect-submodules "sms_ui" ^
  --clean ^
  sms\sms.pyw
```
生成的 exe 文件位于：
```bash
dist/sms.exe
```
## 📁 项目结构
```
Air724UG-SMS/
│
├── sms.pyw             # 原始整合版/安全备份版，Release 不打包此文件
├── sms/                # 重构版源码目录，GitHub Actions 和手动打包均使用此目录
│   ├── sms.pyw         # 重构版主程序入口
│   ├── sms_app/        # 启动装配层，负责 main/bootstrap 与跨层绑定
│   ├── sms_core/       # 串口、短信、云控、推送、配置等核心逻辑
│   └── sms_ui/         # Tkinter 窗口、菜单、设置页、托盘等 UI 逻辑
├── requirements.txt    # Python 依赖
├── icon.ico   # 应用图标
├── config.ini # 软件配置文件（首次运行自动创建）
├── sms_logs/  # 短信日志存储目录（首次运行自动创建）
└── tts/       # 语音播报缓存目录（首次运行自动创建）
```
## ✨ 主要特性

### 🔌 串口与设备
- **自动识别 LUAT Modem 串口**
  - 只连接 `LUAT USB Device X Modem`
  - 自动忽略 AT / Diag / MOS / NPI 等非业务串口
- **串口掉线自动重连**
  - USB 拔插
  - 设备重启
  - COM 号变化
- **手动模式**
  - 可手动锁定指定 COM 口
- **实时感知与监控**
  - 状态栏实时显示模组温度与 4G 信号强度 (RSRP)

---

### 🔒 安全、离线与防骚扰
- **信息保密设计**
  - 本软件所有数据（短信内容、日志、配置）均保存在本地
  - 不上传服务器、不经过第三方平台
  - 适用于涉密环境或内网部署
- **完全离线运行**
  - 软件可在无网络环境下正常运行
  - 无需联网即可完成短信接收、显示、播报等全部功能
- **来电防骚扰**
  - 支持“白名单/黑名单/关闭”模式自由切换
  - 底层静默拦截，拒接骚扰电话，杜绝弹窗

---

### 📞 语音通话管理
- **来电互动弹窗**
  - 实时置顶显示来电号码，支持一键接听、挂断或忽略
- **主动呼叫支持**
  - 快捷拨打外部电话，支持国际前缀
 
---

### 📩 短信接收与发送
- **PDU 模式发送短信**
  - 支持中英文长短信发送，完美兼容国际前缀号码
- **精准接收与过滤**
  - 严格匹配短信回调标识，支持关键词过滤（命中后才显示 / 弹窗 / 播报）
  - 支持自定义短信字体颜色与字号
- **自动清屏**
  - 每天 0 点自动清空主显示窗口

---

### 📢 提醒与日志
- **多维度强提醒**
  - 支持短信弹窗提醒、语音播报提示（可自定义播报文本）
- **日志自动化管理**
  - 自动创建日志目录，短信日志按天、分端口记录保存
  - 支持静默自动清理过期日志

---

### 🖥 GUI 体验
- **专业级串口调试窗口**
  - 支持原始数据旁路监控、筛选与交互
  - 内置超全快捷 AT 指令集，支持一键发送测试
- **SIM 卡高级设置**
  - 图形化交互，支持一键输入 PIN/PUK 解锁、开启/关闭 PIN 码锁、修改本机号码
- **系统级便捷集成**
  - 支持在线检测更新（内置 GitHub 代理加速）
  - 最小化到系统托盘、支持程序多开、开机自启、一键创建桌面快捷方式

---

## 🖼 界面展示
![](https://github.com/KPI0/Air724UG-SMS/blob/main/png/1.png)   
![](https://github.com/KPI0/Air724UG-SMS/blob/main/png/2.png)   
![](https://github.com/KPI0/Air724UG-SMS/blob/main/png/3.png)   
![](https://github.com/KPI0/Air724UG-SMS/blob/main/png/4.png)   
![](https://github.com/KPI0/Air724UG-SMS/blob/main/png/5.png)   
![](https://github.com/KPI0/Air724UG-SMS/blob/main/png/6.png)   
![](https://github.com/KPI0/Air724UG-SMS/blob/main/png/7.png)   
![](https://github.com/KPI0/Air724UG-SMS/blob/main/png/8.png)   
![](https://github.com/KPI0/Air724UG-SMS/blob/main/png/9.png) 
