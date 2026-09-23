from pathlib import Path

from PIL import Image, ImageDraw, ImageFont


ROOT = Path(__file__).resolve().parents[1]
SOURCE = ROOT / "output" / "qa-evidence"
TARGET = SOURCE / "framed"
TARGET.mkdir(parents=True, exist_ok=True)

FORM_CASES = {"01", "03", "04", "05", "06", "08"}
CROP_BOTTOM = {"11": 930, "12": 970}


def font(size: int, bold: bool = False) -> ImageFont.FreeTypeFont:
    name = "arialbd.ttf" if bold else "arial.ttf"
    return ImageFont.truetype(str(Path("C:/Windows/Fonts") / name), size)


def url_for(path: Path) -> str:
    prefix = path.name[:2]
    if prefix in FORM_CASES:
        return "http://localhost:8080/bookings/new"
    return "http://localhost:8080/bookings"


for source in sorted(SOURCE.glob("*.png")):
    image = Image.open(source).convert("RGB")
    prefix = source.name[:2]
    if prefix in CROP_BOTTOM and image.height > CROP_BOTTOM[prefix]:
        image = image.crop((0, 0, image.width, CROP_BOTTOM[prefix]))
    width, height = image.size
    chrome_h = 104
    canvas = Image.new("RGB", (width, height + chrome_h), "#f4f6f8")
    draw = ImageDraw.Draw(canvas)

    draw.rectangle((0, 0, width, 38), fill="#dfe3e8")
    for x, color in [(20, "#ff5f57"), (43, "#febc2e"), (66, "#28c840")]:
        draw.ellipse((x - 7, 12, x + 7, 26), fill=color)
    draw.rounded_rectangle((95, 7, min(width - 16, 390), 36), radius=8, fill="#ffffff")
    draw.text((112, 13), "Booking Management - QA Evidence", fill="#28323c", font=font(15, True))

    draw.rectangle((0, 38, width, chrome_h), fill="#ffffff")
    draw.text((20, 60), "<", fill="#52606d", font=font(24, True))
    draw.text((52, 60), ">", fill="#9aa5b1", font=font(24, True))
    draw.text((86, 60), "↻", fill="#52606d", font=font(22, True))
    draw.rounded_rectangle((122, 50, width - 20, 92), radius=17, fill="#edf1f5", outline="#d1d7de")
    draw.text((143, 61), "🔒", fill="#52606d", font=font(16))
    draw.text((172, 62), url_for(source), fill="#1f2933", font=font(17))

    canvas.paste(image, (0, chrome_h))
    canvas.save(TARGET / source.name, quality=95)

print(f"Created {len(list(TARGET.glob('*.png')))} framed screenshots in {TARGET}")
