"""Build the README GIF from native-size PNGs; requires Pillow."""
from pathlib import Path
from PIL import Image, ImageDraw, ImageFont

FOLDER = Path(__file__).resolve().parents[1] / "showcase"
FUNCTIONS = [
    ("recorded_drives", "Recorded Drives"),
    ("manual_entry", "Manual Drive Entry"),
    ("review_entry", "Review Before Saving"),
    ("duplicate_serial", "Duplicate Serial Protection"),
    ("edit_entry", "Edit a Recorded Drive"),
    ("remove_entry", "Confirm Entry Removal"),
    ("workbook_setup", "Workbook Columns and Save Locations"),
    ("activity", "Activity History"),
    ("terminal", "Embedded Terminal Tab (Before Enabling)"),
    ("paused_scanning", "Pause and Resume Scanning"),
    ("version_info", "Version Information and License Details"),
]
SLIDES = [
    (f"{name}_{mode}_mode", f"{label} - {theme} Theme")
    for name, label in FUNCTIONS
    for mode, theme in [("light", "Light"), ("dark", "Dark")]
]
FRAME_DURATION_MS = 4500



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
    frames[0].save(target, save_all=True, append_images=frames[1:], duration=FRAME_DURATION_MS, loop=0, disposal=2, optimize=False)
    with Image.open(target) as gif:
        assert gif.n_frames == len(SLIDES)
        for index in range(gif.n_frames):
            gif.seek(index)
            assert gif.info["duration"] == FRAME_DURATION_MS
    print(f"Saved {target.name}: {len(frames)} frames, 4.5 seconds each, {width}x{height}, {target.stat().st_size:,} bytes")


if __name__ == "__main__":
    main()
