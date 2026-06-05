from __future__ import annotations

import csv
import json
import math
import random
import shutil
import zipfile
from dataclasses import dataclass
from pathlib import Path

from PIL import Image, ImageDraw, ImageFilter, ImageFont


ROOT = Path(__file__).resolve().parent
OUT = ROOT / "synthetic_bottle_dataset"
IMG_SIZE = 320
RNG = random.Random(2306034)


@dataclass(frozen=True)
class BottleClass:
    label: str
    stem: str
    nominal_height: int
    nominal_width: int
    color: tuple[int, int, int]


CLASSES = [
    BottleClass("500ml", "500ml", 132, 48, (64, 160, 210)),
    BottleClass("1L", "1L", 170, 58, (86, 174, 104)),
    BottleClass("1.5L", "1_5L", 210, 68, (236, 168, 66)),
]


def load_font(size: int, bold: bool = False) -> ImageFont.FreeTypeFont | ImageFont.ImageFont:
    candidates = [
        Path("C:/Windows/Fonts/arialbd.ttf" if bold else "C:/Windows/Fonts/arial.ttf"),
        Path("C:/Windows/Fonts/calibrib.ttf" if bold else "C:/Windows/Fonts/calibri.ttf"),
    ]
    for path in candidates:
        if path.exists():
            return ImageFont.truetype(str(path), size)
    return ImageFont.load_default()


FONT_SMALL = load_font(14)
FONT_LABEL = load_font(28, bold=True)


def rounded_rect(draw: ImageDraw.ImageDraw, box: tuple[int, int, int, int], radius: int, fill, outline=None, width: int = 1) -> None:
    draw.rounded_rectangle(box, radius=radius, fill=fill, outline=outline, width=width)


def make_background() -> Image.Image:
    base = Image.new("RGB", (IMG_SIZE, IMG_SIZE), (235, 238, 232))
    draw = ImageDraw.Draw(base)

    # Soft tabletop/conveyor-like bands.
    top = RNG.randint(216, 236)
    draw.rectangle([0, 0, IMG_SIZE, IMG_SIZE], fill=(top, top + RNG.randint(0, 8), top - RNG.randint(0, 12)))
    horizon = RNG.randint(98, 142)
    draw.rectangle([0, horizon, IMG_SIZE, IMG_SIZE], fill=(199, 205, 196))
    for y in range(horizon + 18, IMG_SIZE, 36):
        shade = 185 + RNG.randint(-8, 8)
        draw.line([(0, y), (IMG_SIZE, y + RNG.randint(-3, 3))], fill=(shade, shade + 4, shade - 2), width=2)

    # Add subtle recycling-bin hints without introducing extra object labels.
    for _ in range(RNG.randint(2, 5)):
        x = RNG.randint(0, IMG_SIZE - 70)
        y = RNG.randint(12, IMG_SIZE - 65)
        w = RNG.randint(36, 72)
        h = RNG.randint(18, 36)
        col = RNG.choice([(190, 210, 188), (205, 216, 220), (214, 210, 190)])
        draw.rectangle([x, y, x + w, y + h], fill=col, outline=None)

    # Sensor/camera timestamp style overlay to make samples less identical.
    stamp = f"EDGE-{RNG.randint(1000, 9999)}"
    draw.text((10, IMG_SIZE - 22), stamp, font=FONT_SMALL, fill=(100, 105, 100))

    # Fine noise.
    px = base.load()
    for _ in range(IMG_SIZE * IMG_SIZE // 18):
        x = RNG.randrange(IMG_SIZE)
        y = RNG.randrange(IMG_SIZE)
        r, g, b = px[x, y]
        delta = RNG.randint(-9, 9)
        px[x, y] = (max(0, min(255, r + delta)), max(0, min(255, g + delta)), max(0, min(255, b + delta)))

    return base.filter(ImageFilter.GaussianBlur(radius=0.25))


def fit_text(draw: ImageDraw.ImageDraw, text: str, max_width: int) -> ImageFont.ImageFont:
    for size in range(30, 13, -1):
        font = load_font(size, bold=True)
        box = draw.textbbox((0, 0), text, font=font)
        if box[2] - box[0] <= max_width:
            return font
    return load_font(14, bold=True)


def draw_bottle(base: Image.Image, cls: BottleClass) -> dict[str, int]:
    draw = ImageDraw.Draw(base, "RGBA")
    h = max(90, int(RNG.gauss(cls.nominal_height, 8)))
    w = max(38, int(RNG.gauss(cls.nominal_width, 4)))

    x = RNG.randint(35, IMG_SIZE - w - 35)
    y = RNG.randint(38, IMG_SIZE - h - 24)
    cap_h = max(12, h // 13)
    neck_h = max(18, h // 7)
    shoulder_h = max(15, h // 10)
    body_top = y + cap_h + neck_h
    body_bottom = y + h
    neck_w = max(18, w // 3)
    neck_x = x + (w - neck_w) // 2

    shadow_offset = RNG.randint(6, 12)
    draw.ellipse(
        [x - 8, body_bottom - 5, x + w + 14, body_bottom + shadow_offset],
        fill=(55, 60, 55, 42),
    )

    cap_color = tuple(max(0, min(255, c + RNG.randint(-18, 18))) for c in cls.color)
    rounded_rect(
        draw,
        (neck_x - 2, y, neck_x + neck_w + 2, y + cap_h),
        radius=4,
        fill=cap_color + (245,),
        outline=(55, 65, 65, 120),
        width=1,
    )
    rounded_rect(
        draw,
        (neck_x, y + cap_h - 1, neck_x + neck_w, body_top + 2),
        radius=6,
        fill=(224, 244, 250, 170),
        outline=(80, 120, 140, 120),
        width=1,
    )

    shoulder = [
        (x + w // 2 - neck_w // 2, body_top - shoulder_h),
        (x + w // 2 + neck_w // 2, body_top - shoulder_h),
        (x + w - 2, body_top + 6),
        (x + w, body_bottom - 8),
        (x + w - 8, body_bottom),
        (x + 8, body_bottom),
        (x, body_bottom - 8),
        (x + 2, body_top + 6),
    ]
    draw.polygon(shoulder, fill=(218, 242, 250, 132), outline=(70, 120, 150, 125))

    # Plastic grooves.
    for gy in [body_top + h * 0.18, body_top + h * 0.32, body_top + h * 0.72]:
        yy = int(gy + RNG.randint(-2, 2))
        draw.arc([x + 4, yy - 4, x + w - 4, yy + 8], 0, 180, fill=(90, 130, 145, 92), width=2)

    # Highlight and side tint.
    draw.line([(x + w // 4, body_top + 8), (x + w // 5, body_bottom - 14)], fill=(255, 255, 255, 145), width=3)
    draw.line([(x + w - 9, body_top + 8), (x + w - 8, body_bottom - 13)], fill=(70, 115, 140, 70), width=3)

    label_h = max(30, int(h * 0.19))
    label_y = int(body_top + (body_bottom - body_top - label_h) * RNG.uniform(0.42, 0.55))
    label_margin = max(4, w // 10)
    label_box = (x + label_margin, label_y, x + w - label_margin, label_y + label_h)
    rounded_rect(
        draw,
        label_box,
        radius=5,
        fill=cls.color + (232,),
        outline=(45, 45, 45, 150),
        width=1,
    )
    text_font = fit_text(draw, cls.label, max(20, label_box[2] - label_box[0] - 8))
    text_box = draw.textbbox((0, 0), cls.label, font=text_font)
    tx = label_box[0] + ((label_box[2] - label_box[0]) - (text_box[2] - text_box[0])) // 2
    ty = label_box[1] + ((label_box[3] - label_box[1]) - (text_box[3] - text_box[1])) // 2 - 1
    draw.text((tx + 1, ty + 1), cls.label, font=text_font, fill=(0, 0, 0, 110))
    draw.text((tx, ty), cls.label, font=text_font, fill=(255, 255, 255, 255))

    # Mild sensor compression/blur variation.
    if RNG.random() < 0.32:
        base.paste(base.filter(ImageFilter.GaussianBlur(radius=0.35)))

    return {
        "x": max(0, x),
        "y": max(0, y),
        "width": min(IMG_SIZE - x, w),
        "height": min(IMG_SIZE - y, h),
    }


def build_split(split: str, count_per_class: int) -> list[dict]:
    split_dir = OUT / split
    split_dir.mkdir(parents=True, exist_ok=True)
    records: list[dict] = []

    for cls in CLASSES:
        for i in range(count_per_class):
            image = make_background()
            bbox = draw_bottle(image, cls)
            if RNG.random() < 0.42:
                image = image.filter(ImageFilter.UnsharpMask(radius=1.1, percent=105, threshold=4))
            filename = f"bottle_{cls.stem}_{split}_{i:03d}.jpg"
            image.save(split_dir / filename, quality=88, optimize=True)
            records.append(
                {
                    "path": filename,
                    "category": split,
                    "label": {"type": "label", "label": cls.label},
                    "metadata": {
                        "source": "synthetic prototype generated for PTIT sensor-network final essay",
                        "size_label": cls.label,
                    },
                    "boundingBoxes": [
                        {
                            "label": cls.label,
                            "x": bbox["x"],
                            "y": bbox["y"],
                            "width": bbox["width"],
                            "height": bbox["height"],
                        }
                    ],
                }
            )

    with (split_dir / "bounding_boxes.labels").open("w", encoding="utf-8") as f:
        json.dump({"version": 1, "files": records}, f, indent=2, ensure_ascii=False)
    return records


def write_metrics(all_records: list[dict]) -> None:
    rows = []
    summary: dict[str, list[float]] = {cls.label: [] for cls in CLASSES}
    for rec in all_records:
        bb = rec["boundingBoxes"][0]
        ratio = (bb["width"] * bb["height"]) / (IMG_SIZE * IMG_SIZE)
        summary[bb["label"]].append(ratio)
        rows.append(
            {
                "file": rec["path"],
                "category": rec["category"],
                "label": bb["label"],
                "bbox_width": bb["width"],
                "bbox_height": bb["height"],
                "bbox_area_ratio": f"{ratio:.5f}",
            }
        )
    with (OUT / "bbox_area_metrics.csv").open("w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=list(rows[0].keys()))
        writer.writeheader()
        writer.writerows(rows)

    with (OUT / "bbox_area_summary.csv").open("w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=["label", "count", "mean_bbox_area_ratio", "median_bbox_area_ratio"])
        writer.writeheader()
        for label, values in summary.items():
            values = sorted(values)
            mean = sum(values) / len(values)
            mid = len(values) // 2
            median = values[mid] if len(values) % 2 else (values[mid - 1] + values[mid]) / 2
            writer.writerow(
                {
                    "label": label,
                    "count": len(values),
                    "mean_bbox_area_ratio": f"{mean:.5f}",
                    "median_bbox_area_ratio": f"{median:.5f}",
                }
            )


def make_contact_sheet(records: list[dict]) -> None:
    picks = []
    for cls in CLASSES:
        picks.extend([r for r in records if r["label"]["label"] == cls.label][:4])
    thumb = 160
    sheet = Image.new("RGB", (thumb * 4, thumb * 3 + 30), (250, 250, 248))
    draw = ImageDraw.Draw(sheet)
    for idx, rec in enumerate(picks):
        split = rec["category"]
        img = Image.open(OUT / split / rec["path"]).convert("RGB")
        bb = rec["boundingBoxes"][0]
        d = ImageDraw.Draw(img)
        d.rectangle([bb["x"], bb["y"], bb["x"] + bb["width"], bb["y"] + bb["height"]], outline=(230, 40, 40), width=3)
        img.thumbnail((thumb, thumb))
        x = (idx % 4) * thumb
        y = (idx // 4) * thumb + 30
        sheet.paste(img, (x, y))
        draw.text((x + 6, y + 6), rec["label"]["label"], font=FONT_SMALL, fill=(20, 20, 20))
    draw.text((10, 8), "Synthetic bottle object-detection dataset preview", font=FONT_SMALL, fill=(40, 40, 40))
    sheet.save(OUT / "dataset_contact_sheet.jpg", quality=90)


def zip_dataset() -> None:
    zip_path = ROOT / "synthetic_bottle_dataset_edge_impulse.zip"
    if zip_path.exists():
        zip_path.unlink()
    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for path in sorted(OUT.rglob("*")):
            if path.is_file():
                zf.write(path, path.relative_to(OUT))


def main() -> None:
    if OUT.exists():
        shutil.rmtree(OUT)
    OUT.mkdir(parents=True, exist_ok=True)
    training = build_split("training", 30)
    testing = build_split("testing", 10)
    all_records = training + testing
    write_metrics(all_records)
    make_contact_sheet(all_records)
    zip_dataset()
    print(f"Generated {len(training)} training and {len(testing)} testing images in {OUT}")
    print(f"Zip: {ROOT / 'synthetic_bottle_dataset_edge_impulse.zip'}")


if __name__ == "__main__":
    main()
