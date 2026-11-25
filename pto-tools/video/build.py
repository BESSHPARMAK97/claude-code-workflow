"""Сборка видео: кадры (HTML -> PNG), озвучка RHVoice, сегменты ffmpeg, склейка, субтитры."""
import base64, html, io, json, os, re, subprocess, sys
from PIL import Image

from pathlib import Path
V = str(Path(__file__).resolve().parent / "work") + "/"
A, S, OUT = V + "assets/", V + "slides/", V + "out/"
os.makedirs(S, exist_ok=True); os.makedirs(OUT, exist_ok=True)
sys.path.insert(0, str(Path(__file__).resolve().parent))
from scenes import SCENES

VOICE = "aleksandr-hq"
W, H = 1920, 1080

# составной кадр: акт + приложение рядом
# составной кадр: п. 3 акта («согласно приложению») над началом листа приложения
a = Image.open(A + "kit-4.png").crop((90, 668, 1000, 800))
c = Image.open(A + "kit-5.png").crop((60, 36, 1015, 515))
pair = Image.new("RGB", (max(a.width, c.width), a.height + c.height + 50), "white")
pair.paste(a, (0, 0))
from PIL import ImageDraw
ImageDraw.Draw(pair).line((0, a.height + 25, pair.width, a.height + 25), fill=(232, 89, 12), width=4)
pair.paste(c, (0, a.height + 50)); pair.save(A + "kit_pair.png")


def img_uri(name, crop=None):
    im = Image.open(A + name).convert("RGB")
    if crop:
        im = im.crop(crop)
    buf = io.BytesIO(); im.save(buf, "PNG")
    return "data:image/png;base64," + base64.b64encode(buf.getvalue()).decode()


CSS = """
* { box-sizing: border-box; margin: 0; }
body { width: 1920px; height: 1080px; overflow: hidden; background: #eef1f6; color: #1d2430;
       font-family: "Noto Sans", "DejaVu Sans", "Liberation Sans", sans-serif; }
.top { position: absolute; left: 80px; top: 54px; right: 80px; }
.step { display: inline-block; background: #2f5597; color: #fff; font-size: 26px; font-weight: 700; padding: 6px 18px; border-radius: 8px; }
h1 { font-size: 56px; margin-top: 16px; letter-spacing: -0.5px; }
.shot { position: absolute; background: #fff; border-radius: 14px; box-shadow: 0 10px 40px rgba(20,40,80,.18); overflow: hidden; display: flex; align-items: center; justify-content: center; }
.shot img { max-width: 100%; max-height: 100%; display: block; }
ul { list-style: none; padding: 0; }
li { font-size: 32px; line-height: 1.3; margin-bottom: 24px; padding-left: 44px; position: relative; }
li:before { content: ""; position: absolute; left: 0; top: 12px; width: 20px; height: 20px; border-radius: 50%; background: #e8590c; }
.side .shot { left: 80px; top: 230px; width: 1120px; height: 780px; padding: 10px; }
.side ul { position: absolute; left: 1260px; top: 250px; width: 580px; }
.wide .shot { left: 80px; top: 220px; width: 1760px; height: 560px; padding: 10px; }
.wide ul { position: absolute; left: 80px; top: 820px; right: 80px; display: flex; flex-wrap: wrap; gap: 8px 60px; }
.wide li { font-size: 30px; margin-bottom: 8px; }
.card ul { position: absolute; left: 160px; top: 290px; right: 160px; }
.card li { font-size: 44px; margin-bottom: 36px; padding-left: 60px; }
.card li:before { width: 26px; height: 26px; top: 18px; }
.title { display: flex; flex-direction: column; justify-content: center; align-items: center; height: 100%; text-align: center;
         background: linear-gradient(135deg, #1f3f7a, #2f5597 55%, #3d6bbd); color: #fff; }
.title h1 { font-size: 84px; margin: 0 120px; }
.title p { font-size: 36px; margin-top: 30px; opacity: .9; }
.bar { position: absolute; left: 0; bottom: 0; height: 10px; background: #e8590c; }
"""


def slide_html(i, sc):
    prog = f'<div class="bar" style="width:{(i + 1) / len(SCENES) * 100:.1f}%"></div>'
    if sc["layout"] == "title":
        body = f'<div class="title"><h1>{html.escape(sc["title"])}</h1><p>{html.escape(sc["sub"])}</p></div>{prog}'
    else:
        top = f'<div class="top"><span class="step">{html.escape(sc["step"])}</span><h1>{html.escape(sc["title"])}</h1></div>'
        shot = f'<div class="shot"><img src="{img_uri(sc["img"], sc.get("crop"))}"></div>' if sc.get("img") else ""
        lis = "".join(f"<li>{html.escape(b)}</li>" for b in sc.get("bullets", []))
        body = f'<div class="{sc["layout"]}" style="position:absolute;inset:0">{top}{shot}<ul>{lis}</ul>{prog}</div>'
    return f'<!doctype html><html><head><meta charset="utf-8"><style>{CSS}</style></head><body>{body}</body></html>'


for i, sc in enumerate(SCENES):
    open(S + f"s{i:02d}.html", "w", encoding="utf-8").write(slide_html(i, sc))
subprocess.run(["node", str(Path(__file__).resolve().parent / "slideshots.js"), S], check=True)

# озвучка + сегменты
segs, srt, t = [], [], 0.0
PAD_BEFORE, PAD_AFTER = 0.5, 0.9


def dur(p):
    return float(subprocess.check_output(["ffprobe", "-v", "error", "-show_entries", "format=duration", "-of", "csv=p=0", p]).decode())


def ts(x):
    h, r = divmod(x, 3600); m, s = divmod(r, 60)
    return f"{int(h):02d}:{int(m):02d}:{int(s):02d},{int(round((s - int(s)) * 1000)):03d}"


for i, sc in enumerate(SCENES):
    wav = OUT + f"a{i:02d}.wav"
    subprocess.run(["RHVoice-test", "-p", VOICE, "-o", wav], input=sc["say"].encode(), check=True, capture_output=True)
    d = dur(wav)
    total = PAD_BEFORE + d + PAD_AFTER
    seg = OUT + f"v{i:02d}.mp4"
    fade_out = max(total - 0.35, 0)
    subprocess.run(["ffmpeg", "-y", "-v", "error", "-loop", "1", "-framerate", "30", "-i", S + f"s{i:02d}.png", "-i", wav,
                    "-filter_complex",
                    f"[0:v]scale={W}:{H},format=yuv420p,fade=t=in:st=0:d=0.35,fade=t=out:st={fade_out:.2f}:d=0.35[v];"
                    f"[1:a]adelay={int(PAD_BEFORE * 1000)}|{int(PAD_BEFORE * 1000)},apad,atrim=0:{total:.3f},aresample=44100[a]",
                    "-map", "[v]", "-map", "[a]", "-t", f"{total:.3f}", "-c:v", "libx264", "-preset", "medium", "-crf", "20",
                    "-tune", "stillimage", "-r", "30", "-c:a", "aac", "-b:a", "128k", "-ac", "1", seg], check=True)
    segs.append(seg)
    # субтитры: по предложениям, время пропорционально длине
    sents = [s.strip() for s in re.split(r"(?<=[.!?])\s+", sc["sub_text"] if "sub_text" in sc else sc["say"]) if s.strip()]
    chars = sum(len(s) for s in sents)
    cur = t + PAD_BEFORE
    for s in sents:
        sd = d * len(s) / chars
        srt.append((cur, cur + sd, s)); cur += sd
    t += total
    print(f"сцена {i + 1:2d}: {total:5.1f} с")

with open(OUT + "list.txt", "w") as f:
    for s in segs:
        f.write(f"file '{s}'\n")
with open(OUT + "subs.srt", "w", encoding="utf-8") as f:
    for k, (a_, b_, s) in enumerate(srt, 1):
        f.write(f"{k}\n{ts(a_)} --> {ts(b_)}\n{s}\n\n")
final = str(Path(__file__).resolve().parent.parent / "dist" / "Видео_инструкция_АОСР.mp4")
subprocess.run(["ffmpeg", "-y", "-v", "error", "-f", "concat", "-safe", "0", "-i", OUT + "list.txt", "-i", OUT + "subs.srt",
                "-map", "0:v", "-map", "0:a", "-map", "1", "-c:v", "copy", "-c:a", "copy", "-c:s", "mov_text",
                "-metadata:s:s:0", "language=rus", "-metadata:s:a:0", "language=rus", "-movflags", "+faststart", final], check=True)
print("итого:", round(t, 1), "с ->", final)
