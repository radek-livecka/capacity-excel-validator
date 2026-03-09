"""
create_assets.py — Generuje PNG assety loga PARAMA Software.
Spusť jednou: python graphics/create_assets.py

Nevyžaduje cairosvg — symbol je překreslen přes Pillow na základě
přesných souřadnic z 05-parama-symbol.svg.
"""

from PIL import Image, ImageDraw
import os

COLOR_BG   = "#152338"   # tmavá navy z SVG
COLOR_TEAL = "#00C2A8"
COLOR_BLUE = "#0FA3FF"

# Uzly trojúhelníku v souřadnicovém systému 0–100
NODE_TL = (32, 36)   # vlevo nahoře — teal
NODE_TR = (68, 36)   # vpravo nahoře — blue
NODE_B  = (50, 64)   # dole — teal


def make_symbol(size: int) -> Image.Image:
    """Vykreslí symbol 2× (SSAA) pro anti-aliasing a zmenší na `size`."""
    sc = size * 2  # supersampling velikost

    def px(x, y):
        """Přepočet 0-100 souřadnic na pixely."""
        return (x * sc / 100, y * sc / 100)

    img = Image.new("RGB", (sc, sc), COLOR_BG)
    d = ImageDraw.Draw(img)

    lw = max(2, round(3 * sc / 100))      # tloušťka čar (škálovaná)
    r_inner = max(4, round(5 * sc / 100)) # poloměr plných koulí

    # Spojnice
    d.line([px(*NODE_TL), px(*NODE_B)],  fill=COLOR_TEAL, width=lw)
    d.line([px(*NODE_TR), px(*NODE_B)],  fill=COLOR_TEAL, width=lw)
    d.line([px(*NODE_TL), px(*NODE_TR)], fill=COLOR_BLUE,  width=lw)

    # Plné koule na uzlech
    for (x, y, col) in [
        (*NODE_TL, COLOR_TEAL),
        (*NODE_TR, COLOR_BLUE),
        (*NODE_B,  COLOR_TEAL),
    ]:
        cx, cy = px(x, y)
        d.ellipse([cx - r_inner, cy - r_inner,
                   cx + r_inner, cy + r_inner], fill=col)

    return img.resize((size, size), Image.LANCZOS)


if __name__ == "__main__":
    out_dir = os.path.dirname(os.path.abspath(__file__))

    for size, name in [(64, "parama-symbol.png"), (32, "parama-icon.png")]:
        path = os.path.join(out_dir, name)
        make_symbol(size).save(path, "PNG")
        print(f"  Uloženo: {path}")

    print("Hotovo.")
