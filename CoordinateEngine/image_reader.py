from PIL import Image
from CoordinateEngine.models.field_def import ContentBounds


def analyze(png_path):
    img = Image.open(png_path).convert("RGBA")
    w, h = img.size
    pixels = img.load()
    min_x, min_y = w, h
    max_x, max_y = 0, 0
    found = False
    step = max(1, min(w, h) // 500)

    for y in range(0, h, step):
        for x in range(0, w, step):
            r, g, b, a = pixels[x, y]
            if a > 128 and (r < 240 or g < 240 or b < 240):
                if x < min_x: min_x = x
                if y < min_y: min_y = y
                if x > max_x: max_x = x
                if y > max_y: max_y = y
                found = True

    if not found:
        min_x, min_y = int(w * 0.05), int(h * 0.05)
        max_x, max_y = int(w * 0.95), int(h * 0.95)

    min_x = max(0, min_x - 1)
    min_y = max(0, min_y - 1)
    max_x = min(w - 1, max_x + 1)
    max_y = min(h - 1, max_y + 1)

    return ContentBounds(
        x0=min_x, y0=min_y, x1=max_x, y1=max_y,
        width_pt=max_x - min_x, height_pt=max_y - min_y,
        page_width_pt=w, page_height_pt=h,
        method="png"
    )
