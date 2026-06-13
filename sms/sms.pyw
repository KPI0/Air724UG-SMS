# ================================================================
#  短信监听系统
#
#  功能简介：
#  ------------------------------------------------
#  本软件用于监听 LUAT 模组（如 Air724UG）串口输出，
#  实时识别指定短信回调日志，并进行如下处理：
#
#  1. 自动识别 LUAT Modem 串口（支持 Auto / Manual 模式）
#  2. 严格匹配短信回调标识 [handler_sms.smsCallback]
#  3. 支持关键词过滤（命中后才显示 / 弹窗 / 播报）
#  4. 支持语音播报（可自定义播报内容）
#  5. 支持短信弹窗提醒，电话弹窗提醒
#  6. 支持串口调试窗口（原始数据旁路）
#  7. 支持日志分端口记录与自动清理
#  8. 支持单实例运行（可选允许多开）
#  9. 支持开机自启与桌面快捷方式创建
# 10. 支持在线检测更新（支持代理）
#
#  作者：ChatGPT、Gemini、Codex、Claude、KPI0
#  GitHub：https://github.com/KPI0/Air724UG-SMS
# 
#  Air724UG-SMS
#
#  Refactored entrypoint: startup wiring lives in sms_app.bootstrap.main.
# ================================================================

from sms_app.bootstrap import main


if __name__ == "__main__":
    main()
