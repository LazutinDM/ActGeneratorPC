"""Create a transparent toolbar mark from the original application icon."""

from collections import deque
from pathlib import Path

from PIL import Image


ROOT = Path(__file__).resolve().parents[1]
SOURCE = ROOT / "Data" / "Icons" / "app.ico"
TARGET = ROOT / "Data" / "Icons" / "brand.png"


def is_outer_white(pixel: tuple[int, int, int, int]) -> bool:
    red, green, blue, _alpha = pixel
    return min(red, green, blue) >= 226 and max(red, green, blue) - min(red, green, blue) <= 22


def extract() -> Image.Image:
    image = Image.open(SOURCE).convert("RGBA")
    width, height = image.size
    pixels = image.load()
    outside: set[tuple[int, int]] = set()
    queue: deque[tuple[int, int]] = deque()

    for x in range(width):
        queue.append((x, 0))
        queue.append((x, height - 1))
    for y in range(height):
        queue.append((0, y))
        queue.append((width - 1, y))

    while queue:
        x, y = queue.popleft()
        if (x, y) in outside or not is_outer_white(pixels[x, y]):
            continue
        outside.add((x, y))
        if x > 0:
            queue.append((x - 1, y))
        if x + 1 < width:
            queue.append((x + 1, y))
        if y > 0:
            queue.append((x, y - 1))
        if y + 1 < height:
            queue.append((x, y + 1))

    for x, y in outside:
        red, green, blue, _alpha = pixels[x, y]
        whiteness = min(red, green, blue)
        alpha = max(0, min(255, (245 - whiteness) * 14))
        pixels[x, y] = (red, green, blue, alpha)

    alpha = image.getchannel("A")
    box = alpha.getbbox()
    if box is None:
        raise ValueError("The application icon became fully transparent.")
    cropped = image.crop(box)
    side = max(cropped.size) + 8
    square = Image.new("RGBA", (side, side), (0, 0, 0, 0))
    square.alpha_composite(
        cropped,
        ((side - cropped.width) // 2, (side - cropped.height) // 2),
    )
    return square.resize((96, 96), Image.Resampling.LANCZOS)


if __name__ == "__main__":
    extract().save(TARGET)
