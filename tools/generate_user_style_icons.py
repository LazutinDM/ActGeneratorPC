"""Generate glossy PNG icons that match the original ActGenerator artwork."""

from pathlib import Path

from PIL import Image, ImageDraw


ROOT = Path(__file__).resolve().parents[1]
ICONS = ROOT / "Data" / "Icons"
SCALE = 8


def rounded_gradient(size: int, radius: int) -> Image.Image:
    canvas = size * SCALE
    image = Image.new("RGBA", (canvas, canvas), (0, 0, 0, 0))
    gradient = Image.new("RGBA", image.size)
    draw = ImageDraw.Draw(gradient)
    top = (18, 132, 230, 255)
    bottom = (3, 73, 158, 255)
    for y in range(canvas):
        t = y / max(1, canvas - 1)
        color = tuple(int(top[i] + (bottom[i] - top[i]) * t) for i in range(4))
        draw.line((0, y, canvas, y), fill=color)
    mask = Image.new("L", image.size, 0)
    ImageDraw.Draw(mask).rounded_rectangle(
        (SCALE, SCALE, canvas - SCALE, canvas - SCALE),
        radius=radius * SCALE,
        fill=255,
    )
    image.paste(gradient, (0, 0), mask)
    shine = Image.new("RGBA", image.size, (0, 0, 0, 0))
    ImageDraw.Draw(shine).ellipse(
        (-canvas // 5, -canvas // 2, canvas + canvas // 5, canvas // 2),
        fill=(255, 255, 255, 34),
    )
    image.alpha_composite(shine)
    return image


def make_brand() -> Image.Image:
    size = 64
    image = rounded_gradient(size, 15)
    draw = ImageDraw.Draw(image)
    s = SCALE
    # Tool silhouettes behind the document echo the original application icon.
    draw.line((13*s, 49*s, 25*s, 20*s), fill=(231, 67, 54, 255), width=5*s)
    draw.line((15*s, 48*s, 27*s, 19*s), fill=(255, 144, 62, 255), width=2*s)
    draw.line((48*s, 51*s, 54*s, 22*s), fill=(205, 218, 231, 255), width=5*s)
    draw.ellipse((49*s, 16*s, 59*s, 27*s), fill=(205, 218, 231, 255))
    draw.ellipse((52*s, 17*s, 59*s, 23*s), fill=(10, 89, 173, 255))

    page = [(20*s, 11*s), (42*s, 11*s), (51*s, 20*s), (51*s, 54*s), (20*s, 54*s)]
    draw.polygon(page, fill=(250, 252, 255, 255), outline=(173, 196, 220, 255), width=2*s)
    draw.polygon([(42*s, 11*s), (42*s, 20*s), (51*s, 20*s)], fill=(213, 228, 242, 255))
    for y, width in ((27, 22), (33, 20), (39, 17)):
        draw.rounded_rectangle((25*s, y*s, (25+width)*s, (y+2)*s), radius=s, fill=(126, 158, 193, 255))
    draw.line((26*s, 47*s, 31*s, 51*s, 43*s, 41*s), fill=(5, 79, 173, 255), width=3*s, joint="curve")
    return image.resize((64, 64), Image.Resampling.LANCZOS)


def make_modes() -> Image.Image:
    size = 40
    image = rounded_gradient(size, 10)
    draw = ImageDraw.Draw(image)
    s = SCALE
    tracks = ((10, 13, 25, (80, 220, 255, 255)),
              (10, 20, 16, (255, 205, 67, 255)),
              (10, 27, 22, (75, 225, 151, 255)))
    for left, y, knob, color in tracks:
        draw.rounded_rectangle(
            (left*s, (y-1)*s, 31*s, (y+1)*s),
            radius=s,
            fill=(236, 247, 255, 235),
        )
        draw.ellipse(((knob-3)*s, (y-3)*s, (knob+3)*s, (y+3)*s), fill=(2, 55, 123, 170))
        draw.ellipse(((knob-2)*s, (y-2)*s, (knob+2)*s, (y+2)*s), fill=color)
    return image.resize((40, 40), Image.Resampling.LANCZOS)


if __name__ == "__main__":
    ICONS.mkdir(parents=True, exist_ok=True)
    make_brand().save(ICONS / "brand.png")
    make_modes().save(ICONS / "modes.png")
