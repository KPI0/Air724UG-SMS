def format_connected_status(port):
    port_text = str(port or "").strip()
    return f"🟢 已连接：{port_text}" if port_text else "🟢 已连接"


def format_connecting_status(port):
    port_text = str(port or "").strip()
    return f"🟡 连接中：{port_text}" if port_text else "🟡 连接中"
