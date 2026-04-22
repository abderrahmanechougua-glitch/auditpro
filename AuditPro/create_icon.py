#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Génère AuditPro.ico (Windows) dans resources/
Couleurs : violet #4B286D + texte blanc
Usage : python create_icon.py
"""
from pathlib import Path
from PIL import Image, ImageDraw, ImageFont


def make_icon(size: int) -> Image.Image:
    """Crée une image carré AuditPro à la taille donnée."""
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)

    # Fond violet foncé avec coins arrondis
    margin = max(2, size // 16)
    radius = size // 5
    x0, y0, x1, y1 = margin, margin, size - margin, size - margin

    # Dessin du fond arrondi (simulation par cercles + rectangles)
    primary = (75, 40, 109, 255)    # #4B286D
    accent  = (232, 51, 109, 255)   # #E8336D

    draw.rounded_rectangle([x0, y0, x1, y1], radius=radius, fill=primary)

    # Barre d'accent en bas
    bar_h = max(3, size // 8)
    draw.rounded_rectangle(
        [x0, y1 - bar_h, x1, y1],
        radius=radius // 2,
        fill=accent
    )

    # Lettre "A" centrée
    font_size = int(size * 0.52)
    try:
        font = ImageFont.truetype("arialbd.ttf", font_size)
    except OSError:
        try:
            font = ImageFont.truetype("C:/Windows/Fonts/arialbd.ttf", font_size)
        except OSError:
            font = ImageFont.load_default()

    text = "A"
    bbox = draw.textbbox((0, 0), text, font=font)
    tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
    tx = (size - tw) // 2 - bbox[0]
    ty = (size - th) // 2 - bbox[1] - size // 14
    draw.text((tx, ty), text, fill=(255, 255, 255, 255), font=font)

    return img


def main():
    out_dir = Path(__file__).parent / "resources"
    out_dir.mkdir(exist_ok=True)
    out_ico = out_dir / "AuditPro.ico"

    # Générer les tailles requises pour .ico Windows
    sizes = [16, 24, 32, 48, 64, 128, 256]
    images = [make_icon(s) for s in sizes]

    # Sauvegarder en .ico multi-résolution
    images[0].save(
        str(out_ico),
        format="ICO",
        sizes=[(s, s) for s in sizes],
        append_images=images[1:],
    )
    print(f"[OK] Icone generee : {out_ico}")

    # Aussi un PNG 256x256 de référence
    out_png = out_dir / "AuditPro_256.png"
    make_icon(256).save(str(out_png), format="PNG")
    print(f"[OK] PNG genere    : {out_png}")


if __name__ == "__main__":
    main()
