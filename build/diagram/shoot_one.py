# -*- coding: utf-8 -*-
"""单张 SVG 按实际尺寸截图，便于逐图检查排版。"""
import io, os, subprocess, sys
from PIL import Image

EDGE = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
D = os.path.dirname(os.path.abspath(__file__))
SCALE = 2
SIZES = {"_r0.svg": (1500, 692), "_r1.svg": (1500, 852), "_r2.svg": (1500, 702),
         "_r3.svg": (1500, 1096), "_r4.svg": (1500, 704),
         "_r5.svg": (1500, 706), "_r6.svg": (1500, 890)}


def shoot(stem):
    w, h = SIZES[stem]
    svg = io.open(os.path.join(D, stem), encoding="utf-8").read()
    html = os.path.join(D, stem.replace(".svg", ".html"))
    io.open(html, "w", encoding="utf-8", newline="\n").write(
        '<!DOCTYPE html><html><head><meta charset="utf-8">'
        '<style>html,body{margin:0;padding:0;background:#fff}svg{display:block}</style>'
        '</head><body>' + svg + '</body></html>')
    raw = os.path.join(D, stem.replace(".svg", "_raw.png"))
    out = os.path.join(D, stem.replace(".svg", ".png"))
    subprocess.run([EDGE, "--headless=new", "--disable-gpu", "--hide-scrollbars",
                    "--force-device-scale-factor=%d" % SCALE, "--screenshot=" + raw,
                    "--window-size=%d,%d" % (w, h + 60),
                    "file:///" + html.replace("\\", "/")], check=True, capture_output=True)
    im = Image.open(raw)
    im.crop((0, 0, w * SCALE, h * SCALE)).save(out)
    os.remove(raw)
    print(stem, "->", out, im.size)


if __name__ == "__main__":
    # SVG 是中间产物，不随仓库提交：每次截图前按当前脚本重新生成
    import gen_recall
    for name, fn in (("_r0.svg", gen_recall.fig0), ("_r1.svg", gen_recall.fig1),
                     ("_r2.svg", gen_recall.fig3), ("_r3.svg", gen_recall.fig2),
                     ("_r4.svg", gen_recall.fig4), ("_r5.svg", gen_recall.fig5),
                     ("_r6.svg", gen_recall.fig6)):
        io.open(os.path.join(D, name), "w", encoding="utf-8", newline="\n").write(fn())
    for s in (sys.argv[1:] or SIZES.keys()):
        shoot(s)
