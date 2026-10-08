"""Create uncommitted visual proposals for the brand and Modes icons."""

from pathlib import Path

from PIL import Image, ImageDraw, ImageFont


ROOT = Path(__file__).resolve().parents[1]
OUTPUT = ROOT / "design" / "icon-proposals"
SIZE = 512
FONT = Path(r"C:\Windows\Fonts\arialbd.ttf")


def base_icon() -> Image.Image:
    image = Image.new("RGBA", (SIZE, SIZE), (0, 0, 0, 0))
    gradient = Image.new("RGBA", image.size)
    draw = ImageDraw.Draw(gradient)
    for y in range(SIZE):
        t = y / (SIZE - 1)
        color = (
            int(16 + (2 - 16) * t),
            int(137 + (62 - 137) * t),
            int(231 + (151 - 231) * t),
            255,
        )
        draw.line((0, y, SIZE, y), fill=color)
    mask = Image.new("L", image.size, 0)
    ImageDraw.Draw(mask).rounded_rectangle((20, 20, 492, 492), radius=118, fill=255)
    image.paste(gradient, (0, 0), mask)
    shine = Image.new("RGBA", image.size, (0, 0, 0, 0))
    ImageDraw.Draw(shine).ellipse((-50, -250, 562, 270), fill=(255, 255, 255, 42))
    image.alpha_composite(shine)
    return image


def brand_proposal() -> Image.Image:
    image = base_icon()
    draw = ImageDraw.Draw(image)
    # Red screwdriver and silver wrench repeat the visual language of app.ico.
    draw.line((95, 408, 177, 108), fill=(126, 18, 25, 255), width=45)
    draw.line((100, 405, 182, 110), fill=(244, 55, 46, 255), width=30)
    draw.rounded_rectangle((78, 340, 125, 465), radius=18, fill=(189, 23, 30, 255))
    draw.line((390, 430, 433, 122), fill=(135, 156, 177, 255), width=54)
    draw.line((394, 425, 437, 120), fill=(221, 232, 241, 255), width=35)
    draw.ellipse((386, 72, 482, 168), fill=(221, 232, 241, 255), outline=(105, 132, 158, 255), width=8)
    draw.polygon(((423, 87), (480, 63), (469, 129), (427, 143)), fill=(6, 101, 190, 255))

    shadow = [(141, 76), (346, 76), (417, 147), (417, 431), (141, 431)]
    draw.polygon([(x+10, y+13) for x, y in shadow], fill=(0, 36, 91, 85))
    page = [(141, 76), (346, 76), (417, 147), (417, 431), (141, 431)]
    draw.polygon(page, fill=(249, 252, 255, 255), outline=(172, 194, 216, 255), width=9)
    draw.polygon(((346, 76), (346, 147), (417, 147)), fill=(210, 226, 241, 255))
    draw.line((346, 76, 346, 147, 417, 147), fill=(153, 181, 207, 255), width=8)

    font = ImageFont.truetype(str(FONT), 80)
    text = "АКТ"
    box = draw.textbbox((0, 0), text, font=font)
    draw.text(((279-(box[2]-box[0]))/2+38, 137), text, font=font, fill=(9, 91, 188, 255))
    for y, length in ((238, 205), (278, 195), (318, 150)):
        draw.rounded_rectangle((181, y, 181+length, y+13), radius=6, fill=(144, 169, 196, 255))
    draw.line((181, 366, 220, 405, 312, 328), fill=(5, 80, 177, 255), width=22, joint="curve")
    draw.ellipse((333, 341, 389, 397), outline=(8, 98, 190, 255), width=8)
    draw.ellipse((346, 354, 376, 384), outline=(8, 98, 190, 255), width=6)
    return image


def modes_proposal() -> Image.Image:
    image = base_icon()
    draw = ImageDraw.Draw(image)
    # Four raised modules communicate optional application modes.
    tiles = (
        (100, 100, (74, 205, 255, 255)),
        (274, 100, (92, 231, 169, 255)),
        (100, 274, (255, 195, 62, 255)),
        (274, 274, (151, 120, 255, 255)),
    )
    for x, y, accent in tiles:
        draw.rounded_rectangle((x+8, y+12, x+145, y+149), radius=31, fill=(0, 35, 88, 90))
        draw.rounded_rectangle((x, y, x+145, y+145), radius=31, fill=(240, 248, 255, 255), outline=(151, 187, 219, 255), width=7)
        draw.rounded_rectangle((x+29, y+29, x+116, y+116), radius=22, fill=accent)
    # Distinct simple symbols remain readable after reduction to 40 px.
    draw.ellipse((146, 146, 199, 199), outline=(255, 255, 255, 255), width=13)
    draw.ellipse((164, 164, 181, 181), fill=(255, 255, 255, 255))
    draw.line((316, 144, 381, 144), fill=(255, 255, 255, 255), width=13)
    draw.line((316, 173, 365, 173), fill=(255, 255, 255, 255), width=13)
    draw.line((316, 202, 376, 202), fill=(255, 255, 255, 255), width=13)
    draw.rectangle((143, 314, 201, 381), outline=(255, 255, 255, 255), width=12)
    draw.rectangle((160, 296, 218, 363), outline=(255, 255, 255, 255), width=12)
    draw.ellipse((312, 312, 381, 381), outline=(255, 255, 255, 255), width=12)
    draw.line((346, 324, 346, 347, 366, 359), fill=(255, 255, 255, 255), width=12)
    return image


def save_versions(name: str, image: Image.Image, final_size: int) -> None:
    OUTPUT.mkdir(parents=True, exist_ok=True)
    image.resize((256, 256), Image.Resampling.LANCZOS).save(OUTPUT / f"{name}-preview.png")
    image.resize((final_size, final_size), Image.Resampling.LANCZOS).save(OUTPUT / f"{name}-candidate.png")


if __name__ == "__main__":
    save_versions("brand-v2", brand_proposal(), 64)
    save_versions("modes-v2", modes_proposal(), 40)
