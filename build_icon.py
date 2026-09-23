"""Создать чёткие PNG/ICO-иконки из простого векторного знака.

Это вспомогательный скрипт; для запуска main.py пакет Pillow не нужен.
"""

from pathlib import Path

from PIL import Image, ImageDraw


ROOT = Path(__file__).resolve().parent
SCALE = 16


def point(x: float, y: float) -> tuple[float, float]:
    return x * SCALE, y * SCALE


def bezier(start, control_1, control_2, end, steps=32):
    result = []
    for index in range(steps + 1):
        t = index / steps
        u = 1 - t
        result.append((
            u**3 * start[0] + 3 * u**2 * t * control_1[0]
            + 3 * u * t**2 * control_2[0] + t**3 * end[0],
            u**3 * start[1] + 3 * u**2 * t * control_1[1]
            + 3 * u * t**2 * control_2[1] + t**3 * end[1],
        ))
    return result


image = Image.new("RGBA", (64 * SCALE, 64 * SCALE), (0, 0, 0, 0))
draw = ImageDraw.Draw(image)
draw.rounded_rectangle((*point(2, 2), *point(62, 62)), radius=14 * SCALE,
                       fill="#E4EEE5")
draw.rounded_rectangle((*point(12, 33), *point(23, 52)), radius=3 * SCALE,
                       fill="#27443B")
draw.rounded_rectangle((*point(27, 20), *point(38, 52)), radius=3 * SCALE,
                       fill="#27443B")
drop = []
segments = (
    ((49, 21), (45, 27), (40, 35), (40, 42)),
    ((40, 42), (40, 50), (44, 54), (49, 54)),
    ((49, 54), (54, 54), (58, 50), (58, 42)),
    ((58, 42), (58, 35), (53, 27), (49, 21)),
)
for segment in segments:
    drop.extend(point(*xy) for xy in bezier(*segment))
draw.polygon(drop, fill="#B77A25")

image.save(ROOT / "app_icon.png")
for size in (16, 32, 40, 48, 256):
    image.resize((size, size), Image.Resampling.LANCZOS).save(
        ROOT / f"app_icon_{size}.png"
    )
image.save(ROOT / "app_icon.ico", format="ICO", sizes=[
    (size, size) for size in (16, 24, 32, 40, 48, 64, 128, 256)
])
