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

在项目根目录执行：

```powershell
py -m pip install -r requirements.txt
```

## 🚀 源码运行

在项目根目录执行：

```powershell
py sms\sms.pyw
```

源码运行至少需要保留：

```text
sms/
requirements.txt
icon.ico
```

首次运行会在项目根目录自动生成：

```text
config.ini
sms_logs/
tts/
autostart_instances.json
```

其中 `config.ini` 可能包含云控密钥、第三方推送 Token、SMTP 密码等个人配置，`sms_logs/` 可能包含短信、来电、设备 IMEI 等运行日志，`autostart_instances.json` 用于记录最近一次运行的实例数量，以便开机自启恢复多开实例。这些都是本地运行态文件，请勿提交到公开仓库，也不要打进公开发布包。

## ✅ 回归测试

仓库已包含 `tests/` 回归测试。源码修改或打包发布前建议执行：

```powershell
$repoRoot = (Resolve-Path ".").Path
$smsSrc = (Resolve-Path "sms").Path
$env:PYTHONPATH = "$repoRoot;$smsSrc"
py -m unittest discover -s tests -q
```

GitHub Actions 会在打包 EXE 之前自动执行同样的测试；测试失败时不会继续发布，避免生成明显异常的 Release。

## 🔖 发布版本号

桌面客户端版本号只在 `sms/sms_app/version.py` 中维护。发布稳定版前先更新其中的
`APP_VERSION`，然后创建完全一致的 `vX.Y.Z` Tag。例如 `APP_VERSION = "3.8.1"`
只能使用 `v3.8.1` 发布。

GitHub Actions 会在安装依赖和打包前执行 `tools/check_release_version.py`。Tag 格式
不正确、包含预发布后缀，或与客户端版本不一致时，发布任务会立即失败。

## 🏗 打包为 EXE

先安装完整依赖：

```powershell
py -m pip install -r requirements.txt
```

然后在项目目录下执行以下命令：

```powershell
py -m PyInstaller `
  --noconfirm `
  --onefile `
  --windowed `
  --name "sms" `
  --icon "icon.ico" `
  --add-data "icon.ico;." `
  --paths "sms" `
  --hidden-import "pyttsx3.drivers.sapi5" `
  --hidden-import "pythoncom" `
  --hidden-import "pywintypes" `
  --hidden-import "win32com" `
  --hidden-import "win32com.client" `
  --hidden-import "sms_app" `
  --hidden-import "sms_core" `
  --hidden-import "sms_ui" `
  --collect-submodules "win32com" `
  --collect-submodules "websockets" `
  --collect-data "certifi" `
  --collect-submodules "sms_app" `
  --collect-submodules "sms_core" `
  --collect-submodules "sms_ui" `
  --clean `
  sms\sms.pyw
```

生成的 exe 文件位于：

```text
dist\sms.exe
```

GitHub Actions 会把 `sms.exe` 压缩为 `sms-vX.Y.Z-win64.zip`，并同时发布对应的
`sms-vX.Y.Z-win64.zip.sha256` 校验文件。下载后建议先核对 SHA-256，再解压运行。
不要把个人使用过的 `config.ini`、`sms_logs/`、`tts/` 或 `autostart_instances.json` 打进公开发布包；程序首次启动会在 exe 同级自动创建所需配置和运行目录。

## 📁 项目结构
```
Air724UG-SMS/
│
├── sms.pyw             # 已停止维护的历史单文件整合版，仅供留档；不再分析、修复、测试或打包
├── sms/                # 重构版源码目录，GitHub Actions 和手动打包均使用此目录
│   ├── sms.pyw         # 重构版主程序入口
│   ├── sms_app/        # 启动装配层，负责 main/bootstrap 与跨层绑定
│   ├── sms_core/       # 串口、短信、云控、推送、配置等核心逻辑
│   └── sms_ui/         # Tkinter 窗口、菜单、设置页、托盘等 UI 逻辑
├── tests/              # 回归测试，Actions 打包前会自动运行
├── tools/              # 测试与维护辅助工具，例如短信串口日志 replay
├── requirements.txt    # Python 依赖
├── icon.ico   # 应用图标
├── config.ini # 软件配置文件（首次运行自动创建）
├── sms_logs/  # 短信日志存储目录（首次运行自动创建）
├── tts/       # 语音播报缓存目录（首次运行自动创建）
└── autostart_instances.json # 多实例开机自启运行态文件（按需生成）
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
- **默认本地运行**
  - 未启用云控和第三方推送时，短信内容、运行日志和配置保存在本地，短信接收、显示、播报、来电处理等核心功能可离线运行
  - 本软件未经过涉密或信息安全产品认证；用于敏感环境前，请根据实际网络、主机和数据管理要求独立评估
- **外部连接与数据范围**
  - 启用云控后，软件会向配置的 WebSocket 服务发送设备身份、短信/来电事件和串口日志；“主动公开设备”只控制设备是否出现在公开列表中，不会停止这些数据上传
  - 公网或其他不可信网络应使用 `wss://`；`ws://` 不加密，会明文传输设备控制密码和业务数据
  - 云端敏感权限默认全部关闭；在“设置 → 云端控制 → 安全设置”中可全部开启、全部关闭或分别控制发送短信、拨打电话、PIN/PUK、修改本机号码或 SN、基站定位、USSD、呼叫转移/限制、信息中心号码、删除设备数据和重置/关闭设备。设置点击后立即保存并生效；`AT+RESET` 不属于“重置或关闭设备”权限
  - 启用第三方推送后，短信或来电数据会按用户配置发送至钉钉、飞书、企业微信、Pushover、Gotify、SMTP 等外部服务
  - 手动检测更新默认会通过第三方 GitHub API/下载代理查询并打开 Release 下载链接，程序不会自动下载或安装更新；下载后应核对 Release 提供的 `.sha256` 文件
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
  - 支持手动在线检测更新（默认使用可在“代理设置”中修改或清空的第三方 GitHub 代理）
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
