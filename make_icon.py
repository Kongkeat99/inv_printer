"""Generates app.ico (a simple printer-style icon) so the .exe has an icon. Run by build.bat."""
from PIL import Image, ImageDraw
S = 256
im = Image.new("RGBA", (S, S), (0, 0, 0, 0))
d = ImageDraw.Draw(im)
d.rounded_rectangle((8, 8, S - 8, S - 8), 44, fill=(31, 78, 121))          # background
d.rectangle((72, 40, 184, 110), fill="white")                               # paper in
d.rounded_rectangle((36, 100, 220, 176), 14, fill=(230, 235, 240))          # printer body
d.rectangle((72, 150, 184, 220), fill="white", outline=(31, 78, 121), width=4)  # paper out
for y in (170, 186, 202):
    d.line((88, y, 168, y), fill=(31, 78, 121), width=4)
d.ellipse((188, 118, 206, 136), fill=(60, 180, 90))
im.save("app.ico", sizes=[(256, 256), (128, 128), (64, 64), (48, 48), (32, 32), (16, 16)])
print("app.ico created")
