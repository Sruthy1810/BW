from pathlib import Path

from openpyxl.drawing.image import Image


LOGO_PATH = Path(__file__).resolve().parent / "Assets" / "image.png"
LOGO_WIDTH = 260



def ensure_statement_logo(ws):
    if "summary" not in ws.title.lower():
        ws._images.clear()
        return

    if ws._images:
        return

    if not LOGO_PATH.is_file():
        raise FileNotFoundError(f"Statement logo not found: {LOGO_PATH}")

    logo = Image(str(LOGO_PATH))
    logo.height = round(logo.height * LOGO_WIDTH / logo.width)
    logo.width = LOGO_WIDTH

