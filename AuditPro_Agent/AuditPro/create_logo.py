"""
Générateur de logo AuditPro — "AC" style Wall Street Journal
Logo violet minimaliste avec typographie serif et cadre.
"""
import matplotlib.pyplot as plt
import matplotlib.patches as patches
from matplotlib.text import TextPath
from matplotlib.patches import PathPatch
from matplotlib.font_manager import FontProperties
import numpy as np
from pathlib import Path


def create_ac_logo(output_path=None, size=1024):
    """
    Crée le logo "AC" violet style WSJ.

    Args:
        output_path: Chemin de sortie (défaut: resources/icons/logo.png)
        size: Taille du logo en pixels (carré)
    """
    if output_path is None:
        # Créer le dossier resources/icons s'il n'existe pas
        icons_dir = Path(__file__).parent / "resources" / "icons"
        icons_dir.mkdir(parents=True, exist_ok=True)
        output_path = icons_dir / "logo.png"

    # Générer un rendu haute résolution pour une image plus nette
    dpi = 300
    fig, ax = plt.subplots(figsize=(size / dpi, size / dpi), dpi=dpi)
    ax.set_xlim(0, size)
    ax.set_ylim(0, size)
    ax.axis('off')

    violet = '#6B21A8'
    fig.patch.set_alpha(0)

    margin = size * 0.1
    content_size = size - 2 * margin

    frame_width = max(2, int(size * 0.02))
    frame = patches.Rectangle(
        (margin - frame_width, margin - frame_width),
        content_size + 2 * frame_width,
        content_size + 2 * frame_width,
        linewidth=frame_width,
        edgecolor=violet,
        facecolor='none',
        alpha=0.35
    )
    ax.add_patch(frame)

    text_x = size * 0.5
    text_y = size * 0.5

    try:
        font = FontProperties(family=['Georgia', 'Times New Roman', 'serif'], size=content_size * 0.55, weight='bold')
    except Exception:
        font = FontProperties(family='serif', size=content_size * 0.55, weight='bold')

    text = ax.text(
        text_x, text_y, 'AC',
        ha='center', va='center',
        fontproperties=font,
        color=violet,
        weight='bold'
    )

    plt.savefig(
        output_path,
        dpi=dpi,
        bbox_inches='tight',
        transparent=True,
        pad_inches=0.05,
        facecolor='none'
    )
    plt.close()

    print(f"Logo généré: {output_path}")
    return output_path


def create_favicon(output_path=None, size=128):
    """Crée une version favicon (petite) du logo."""
    icons_dir = Path(__file__).parent / "resources" / "icons"
    icons_dir.mkdir(parents=True, exist_ok=True)
    if output_path is None:
        output_path = icons_dir / "favicon.png"

    # Créer d'abord un logo de haute qualité puis réduire
    temp_logo = icons_dir / "favicon_source.png"
    create_ac_logo(temp_logo, size * 4)

    try:
        from PIL import Image
        image = Image.open(temp_logo)
        image = image.convert('RGBA')
        image = image.resize((size, size), Image.LANCZOS)
        image.save(output_path, optimize=True)
        temp_logo.unlink()
    except ImportError:
        create_ac_logo(output_path, size)

    print(f"Favicon généré: {output_path}")
    return output_path


if __name__ == "__main__":
    # Générer le logo principal
    logo_path = create_ac_logo()

    # Générer le favicon
    favicon_path = create_favicon()

    print("Logos générés avec succès!")
    print(f"Logo principal: {logo_path}")
    print(f"Favicon: {favicon_path}")