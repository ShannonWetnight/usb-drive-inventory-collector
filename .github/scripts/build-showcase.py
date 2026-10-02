"""Build the README GIF from native-size PNGs; requires Pillow."""
from pathlib import Path
from PIL import Image, ImageDraw, ImageFont

FOLDER = Path(__file__).resolve().parents[1] / "showcase"
SLIDES = [
    ("recorded_drives", "Recorded Drives - dark theme"),
    ("recorded_drives2", "Recorded Drives - light theme"),
    ("manual_entry", "Manual Drive Entry"),
    ("review_entry", "Review before saving"),
    ("duplicate_serial", "Duplicate serial protection"),
    ("edit_entry", "Edit a recorded drive"),
    ("remove_entry", "Confirm entry removal"),
    ("workbook_setup", "Workbook columns and save locations"),
    ("activity", "Activity history"),
    ("terminal", "Embedded Terminal tab (before enabling)"),
    ("paused_scanning", "Pause and resume scanning"),
    ("version_info", "Version Information and license details"),
]


def main():
    images = [Image.open(FOLDER / f"showcase_{name}.png").convert("RGB") for name, _ in SLIDES]
    width = max(image.width for image in images)
    height = max(image.height for image in images) + 48
    font = ImageFont.load_default(size=18)
    frames = []
    for index, (image, (_, label)) in enumerate(zip(images, SLIDES), 1):
        frame = Image.new("RGB", (width, height), (24, 28, 34))
        frame.paste(image, ((width - image.width) // 2, (height - 48 - image.height) // 2))
        draw = ImageDraw.Draw(frame)
        draw.text((18, height - 35), label, fill=(235, 239, 245), font=font)
        draw.text((width - 90, height - 35), f"{index}/{len(images)}", fill=(235, 239, 245), font=font)
        # The original PNGs remain unchanged. GIF palettes use at most 256 colors.
        frames.append(frame.quantize(colors=256, method=Image.Quantize.MEDIANCUT))
    target = FOLDER / "showcase_functions.gif"
    frames[0].save(target, save_all=True, append_images=frames[1:], duration=6000, loop=0, disposal=2, optimize=False)
    with Image.open(target) as gif:
        assert gif.n_frames == len(SLIDES)
        for index in range(gif.n_frames):
            gif.seek(index)
            assert gif.info["duration"] == 6000
    print(f"Saved {target.name}: {len(frames)} frames, 6 seconds each, {width}x{height}, {target.stat().st_size:,} bytes")


if __name__ == "__main__":
    main()
