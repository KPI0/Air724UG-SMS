# Air724UG-SMS
# 📩 短信接收系统（LUAT Modem 自动识别版）

一款基于 **Python + Tkinter** 的 **4G / LTE 模组短信接收与预警显示程序**，  
专为 **LUAT（合宙）系列 Modem** 多串口环境设计，支持 **自动识别 Modem 串口、自动重连、无人值守运行**。

适用于：  
- 预警短信接收  
- 值班室 / 监控室大屏展示  
- 工业上位机  
- 无人值守终端  

---

## ✨ 主要特性

### 🔌 串口与设备
- ✅ **自动识别 LUAT Modem 串口**
  - 只连接 `LUAT USB Device X Modem`
  - 自动忽略 AT / Diag / MOS / NPI 等非业务串口
- 🔁 **串口掉线自动重连**
  - USB 拔插
  - 设备重启
  - COM 号变化
- 🔒 **手动模式**
  - 可手动锁定指定 COM 口（用于特殊调试场景）

---

### 📩 短信接收与显示
- 🔴 **短信红色大号字体显示**
- 🔍 **关键词过滤**
- 📅 **每天 0 点自动清空显示窗口**

---

### 🔊 提醒与日志
- 🔔 系统提示音
- 🗣 语音播报（TTS，自动生成语音文件）
- 📝 **短信日志按天保存**
- 📂 自动创建日志目录

---

### 🖥 GUI 体验
- 🟢 左下角 **实时串口连接状态**
- ⚙ 串口设置弹窗 **始终居中**
- ℹ “关于”窗口 **居中显示**
- 🧭 清晰的 auto / manual 模式提示说明

---

## 🖼 界面展示
![](https://github.com/KPI0/Air724UG-SMS/blob/main/png/1.png)   
![](https://github.com/KPI0/Air724UG-SMS/blob/main/png/2.png)   
![](https://github.com/KPI0/Air724UG-SMS/blob/main/png/3.png)   
![](https://github.com/KPI0/Air724UG-SMS/blob/main/png/4.png)   

