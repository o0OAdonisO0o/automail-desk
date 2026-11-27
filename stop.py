# -*- coding: utf-8 -*-
"""停止面板服务：优先调用 /api/shutdown（进程自行退出），失败则按端口结束进程。

用法：python stop.py [--port 8765]
"""
import json
import os
import signal
import socket
import subprocess
import sys
import time
import urllib.request

BASE = os.path.dirname(os.path.abspath(__file__))


def load_port():
    if "--port" in sys.argv:
        return int(sys.argv[sys.argv.index("--port") + 1])
    if os.environ.get("AUTOMAIL_PORT"):
        return int(os.environ["AUTOMAIL_PORT"])
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


def by_endpoint(port):
    try:
        req = urllib.request.Request(
            "http://127.0.0.1:%d/api/shutdown" % port, data=b"{}",
            headers={"Content-Type": "application/json"}, method="POST")
        urllib.request.urlopen(req, timeout=5).read()
        return True
    except Exception:
        return False


def by_pid(port):
    pids = set()
    if os.name == "nt":
        out = subprocess.run(["netstat", "-ano"], capture_output=True).stdout.decode("gbk", "ignore")
        for line in out.splitlines():
            if ":%d" % port in line and "LISTENING" in line:
                pids.add(line.split()[-1])
    else:
        for cmd in (["lsof", "-ti", "tcp:%d" % port], ["fuser", "%d/tcp" % port]):
            try:
                out = subprocess.run(cmd, capture_output=True).stdout.decode()
                pids.update(out.split())
                if pids:
                    break
            except Exception:
                continue
    for p in pids:
        try:
            os.kill(int(p), signal.SIGTERM)
        except Exception as e:
            print("kill %s failed: %s" % (p, e))
    return len(pids)


def main():
    port = load_port()
    if not listening(port):
        print("no listener on %d" % port)
        return 0
    by_endpoint(port)
    for _ in range(20):
        if not listening(port):
            print("stopped")
            return 0
        time.sleep(0.25)
    n = by_pid(port)
    time.sleep(0.5)
    print("force stopped %d process(es)" % n if n else "stop failed")
    return 0 if not listening(port) else 1


if __name__ == "__main__":
    sys.exit(main())
