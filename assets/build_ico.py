# -*- coding: utf-8 -*-
"""把多尺寸 PNG 打包成 Windows .ico（PNG 压缩条目，Vista 及以上支持）。"""
import struct, os

HERE = os.path.dirname(os.path.abspath(__file__))
SIZES = [256, 64, 48, 32, 16]
OUT = os.path.join(os.path.dirname(HERE), "automail.ico")

blobs = []
for s in SIZES:
    p = os.path.join(HERE, f"icon{s}.png")
    with open(p, "rb") as f:
        d = f.read()
    w, h = struct.unpack(">II", d[16:24])
    assert w == s and h == s, (p, w, h)
    blobs.append((s, d))

header = struct.pack("<HHH", 0, 1, len(blobs))
offset = 6 + 16 * len(blobs)
entries, data = b"", b""
for s, d in blobs:
    dim = 0 if s >= 256 else s
    entries += struct.pack("<BBBBHHII", dim, dim, 0, 0, 1, 32, len(d), offset)
    data += d
    offset += len(d)

with open(OUT, "wb") as f:
    f.write(header + entries + data)
print("ICO ->", OUT, os.path.getsize(OUT), "bytes", [s for s, _ in blobs])
