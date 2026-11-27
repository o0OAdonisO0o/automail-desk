# -*- coding: utf-8 -*-
"""Automail Desk 启动器。

- 面板未运行：后台拉起 server.py（Windows 下不弹控制台），等端口就绪后打开浏览器
- 面板已在运行：仅打开浏览器，不会产生第二个实例
- 启动失败：给出日志路径提示（Windows 弹窗，其他平台打印到 stderr）

用法：python launcher.py [--port 8765] [--no-browser]
"""
import json
import os
import socket
import subprocess
import sys
import time
import webbrowser

BASE = os.path.dirname(os.path.abspath(__file__))


def load_port():
    port = os.environ.get("AUTOMAIL_PORT")
    if port:
        return int(port)
    try:
        with open(os.path.join(BASE, "config.json"), "r", encoding="utf-8-sig") as f:
            return int(json.load(f).get("port") or 8765)
    except Exception:
        return 8765


def listening(port):
    s = socket.socket()
    s.settimeout(0.6)
    try:
        return s.connect_ex(("127.0.0.1", port)) == 0
    finally:
        s.close()


def python_exe():
    """优先用无控制台的解释器；找不到则退回当前解释器。"""
    if os.name == "nt":
        cand = os.path.join(os.path.dirname(sys.executable), "pythonw.exe")
        if os.path.exists(cand):
            return cand
    return sys.executable


def fail(msg, log_path):
    text = msg + "\n\n日志：" + log_path + "\n可改用 start.bat（Windows）或 run.sh（macOS/Linux）前台启动排查。"
    if os.name == "nt":
        try:
            import ctypes

            ctypes.windll.user32.MessageBoxW(0, text, "Automail Desk", 0x10)
            return
        except Exception:
            pass
    print(text, file=sys.stderr)


def main():
    argv = sys.argv[1:]
    port = load_port()
    if "--port" in argv:
        port = int(argv[argv.index("--port") + 1])
    no_browser = "--no-browser" in argv
    url = "http://127.0.0.1:%d" % port

    logdir = os.path.join(BASE, "logs")
    os.makedirs(logdir, exist_ok=True)
    log_path = os.path.join(logdir, "stdout.log")

    if not listening(port):
        out = open(log_path, "ab", buffering=0)
        kwargs = dict(cwd=BASE, stdin=subprocess.DEVNULL, stdout=out, stderr=out, close_fds=True)
        if os.name == "nt":
            # DETACHED_PROCESS | CREATE_NEW_PROCESS_GROUP：脱离当前控制台，关闭窗口不终止服务
            kwargs["creationflags"] = 0x00000008 | 0x00000200
        else:
            kwargs["start_new_session"] = True
        env = dict(os.environ, AUTOMAIL_PORT=str(port))
        subprocess.Popen([python_exe(), os.path.join(BASE, "server.py")], env=env, **kwargs)
        for _ in range(120):
            if listening(port):
                break
            time.sleep(0.25)

    if not listening(port):
        fail("面板未能启动，端口 %d 无响应。" % port, log_path)
        return 1

    if not no_browser:
        webbrowser.open(url)
    print("Automail Desk: " + url)
    return 0


if __name__ == "__main__":
    sys.exit(main())
