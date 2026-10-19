from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

from document_excel.infrastructure.excel import create_demo_template, legacy_profile, lermontov_profile

ROOT = Path(__file__).resolve().parents[1]
RESOURCES = ROOT / "src/document_excel/resources"


def main() -> None:
    image = Image.new("RGBA", (256, 256), "#153252")
    draw = ImageDraw.Draw(image)
    draw.rounded_rectangle((52, 30, 202, 225), radius=17, fill="white")
    draw.rounded_rectangle((32, 122, 195, 215), radius=14, fill="#21704e")
    font = ImageFont.truetype(str(RESOURCES / "fonts/NotoSans-Bold.ttf"), 59)
    draw.text((53, 122), "XLS", font=font, fill="white")
    for y in [63, 82, 101]:
        draw.rounded_rectangle((78, y, 178, y + 7), radius=3, fill="#95abc4")
    image.save(RESOURCES / "app.ico", sizes=[(16, 16), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)])
    templates = RESOURCES / "templates"
    templates.mkdir(exist_ok=True)
    target = templates / "DEMO-not-a-contract.xlsx"
    if not target.exists():
        create_demo_template(target)
    for mode in ("combined", "separate"):
        (templates / f"legacy-{mode}.json").write_text(
            legacy_profile(mode).model_dump_json(indent=2), encoding="utf-8"
        )
    (templates / "lermontov-combined.json").write_text(
        lermontov_profile().model_dump_json(indent=2), encoding="utf-8"
    )


if __name__ == "__main__":
    main()
