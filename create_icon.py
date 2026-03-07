"""Generate a professional PDF Tool icon as .ico file."""
from PIL import Image, ImageDraw, ImageFont
import os

SIZES = [16, 32, 48, 64, 128, 256]


def create_icon_image(size):
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)

    margin = size * 0.1
    w = size - 2 * margin
    h = size - 2 * margin

    # Page shape (white with blue border)
    page_x = margin
    page_y = margin
    page_w = w * 0.85
    page_h = h

    # Dog-ear triangle size
    ear = page_w * 0.25

    # Page outline points (with dog-ear top-right)
    page_points = [
        (page_x, page_y),
        (page_x + page_w - ear, page_y),
        (page_x + page_w, page_y + ear),
        (page_x + page_w, page_y + page_h),
        (page_x, page_y + page_h),
    ]

    # Shadow
    shadow_offset = size * 0.03
    shadow_points = [(x + shadow_offset, y + shadow_offset) for x, y in page_points]
    draw.polygon(shadow_points, fill=(0, 0, 0, 40))

    # Page fill
    draw.polygon(page_points, fill=(255, 255, 255, 255), outline=(25, 118, 210, 255))

    # Dog-ear fold
    ear_points = [
        (page_x + page_w - ear, page_y),
        (page_x + page_w - ear, page_y + ear),
        (page_x + page_w, page_y + ear),
    ]
    draw.polygon(ear_points, fill=(220, 230, 245, 255), outline=(25, 118, 210, 255))

    # "PDF" text on the page
    try:
        font_size = max(8, int(size * 0.22))
        font = ImageFont.truetype("/System/Library/Fonts/Helvetica.ttc", font_size)
    except (OSError, IOError):
        try:
            font = ImageFont.truetype("arial.ttf", font_size)
        except (OSError, IOError):
            font = ImageFont.load_default(size=font_size)

    text = "PDF"
    bbox = draw.textbbox((0, 0), text, font=font)
    tw = bbox[2] - bbox[0]
    th = bbox[3] - bbox[1]
    tx = page_x + (page_w - tw) / 2
    ty = page_y + page_h * 0.35

    draw.text((tx, ty), text, fill=(25, 118, 210, 255), font=font)

    # Compress arrows (down arrow to smaller) - bottom right
    arrow_cx = page_x + page_w + w * 0.08
    arrow_cy = page_y + page_h * 0.55
    arrow_size = size * 0.12

    # Green circle background
    circle_r = size * 0.18
    draw.ellipse(
        [arrow_cx - circle_r, arrow_cy - circle_r,
         arrow_cx + circle_r, arrow_cy + circle_r],
        fill=(76, 175, 80, 255)
    )

    # Down arrow in white
    draw.polygon([
        (arrow_cx, arrow_cy + arrow_size),
        (arrow_cx - arrow_size * 0.8, arrow_cy - arrow_size * 0.3),
        (arrow_cx + arrow_size * 0.8, arrow_cy - arrow_size * 0.3),
    ], fill=(255, 255, 255, 255))

    # Lines on page (fake text)
    line_y_start = page_y + page_h * 0.2
    line_margin = page_w * 0.15
    for i in range(3):
        ly = line_y_start + i * (size * 0.08)
        line_w = page_w * (0.6 if i < 2 else 0.4)
        draw.rectangle(
            [page_x + line_margin, ly,
             page_x + line_margin + line_w, ly + size * 0.025],
            fill=(200, 200, 200, 255)
        )

    return img


def main():
    images = [create_icon_image(s) for s in SIZES]
    ico_path = os.path.join(os.path.dirname(__file__), "app_icon.ico")
    images[-1].save(ico_path, format="ICO", sizes=[(s, s) for s in SIZES],
                    append_images=images[:-1])
    print(f"Icon saved to {ico_path}")


if __name__ == "__main__":
    main()
