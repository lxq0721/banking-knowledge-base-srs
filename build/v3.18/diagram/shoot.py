# -*- coding: utf-8 -*-
"""按 SVG 实际尺寸截图并裁齐（Edge 无头截图的可视高度比 window-size 少约 70px）。"""
import os, subprocess, sys
from PIL import Image

EDGE = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
D = os.path.dirname(os.path.abspath(__file__))
PAD = 140
SCALE = 2


def shoot(stem, w, h):
    html = os.path.join(D, stem + ".html")
    raw = os.path.join(D, stem + "_raw.png")
    out = os.path.join(D, stem + ".png")
    url = "file:///" + html.replace("\\", "/")
    subprocess.run([EDGE, "--headless=new", "--disable-gpu", "--hide-scrollbars",
                    "--force-device-scale-factor=%d" % SCALE,
                    "--screenshot=" + raw,
                    "--window-size=%d,%d" % (w, h + PAD), url],
                   check=True, capture_output=True)
    im = Image.open(raw)
    im.crop((0, 0, w * SCALE, h * SCALE)).save(out)
    os.remove(raw)
    print(stem, im.size, "->", (w * SCALE, h * SCALE))


if __name__ == "__main__":
    shoot(sys.argv[1], int(sys.argv[2]), int(sys.argv[3]))
