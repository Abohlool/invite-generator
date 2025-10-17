#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Batch invitation generator

⚙️ Requires: 
    LibreOffice installed and available in PATH
📄 Usage:
    python generate_invites.py

This script:
1. Reads guest names from a text file.
2. Replaces a placeholder in a PowerPoint template.
3. Exports each customized slide as a JPG.
4. Compresses the final images into a ZIP file.
"""

from pptx import Presentation
from pptx.util import Pt
from pathlib import Path
import subprocess


TEMPLATE_FILE = "invite.pptx"
NAMES_FILE = "names.txt"
OUTPUT_DIR = Path("./output-images")
ZIP_NAME = "invitations.zip"

PLACEHOLDER = "{}"

FONT_NAME = "2 Davat"
FONT_SIZE = Pt(36)



def load_names(file_path: str) -> list[str]:
    """Return list of guest names from `file_path`."""
    with open(file_path, "r", encoding="utf-8") as f:
        return [name.strip() for name in f.read().splitlines() if name.strip()]


def replace_placeholder(slide, name: str):
    """Replace placeholders in slide text with the guest name."""
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        
        frame = shape.text_frame
        for p in frame.paragraphs:
            for r in p.runs:
                if PLACEHOLDER in r.text:
                    r.text = r.text.replace(PLACEHOLDER, name)
                    
                    r.font.name = FONT_NAME
                    r.font.size = FONT_SIZE


def save_and_convert(pptx_path: Path, output_path: Path):
    """Convert PPTX file to JPG using LibreOffice and remove the PPTX file."""
    cmd = [
        "libreoffice",
        "--headless",
        "--convert-to", "jpg",
        str(pptx_path),
        "--outdir", str(output_path),
    ]
    subprocess.run(cmd, check=True)
    pptx_path.unlink(missing_ok=True)


def create_invites(names: list[str]):
    """Generate invitation slides and export them as images."""
    OUTPUT_DIR.mkdir(exist_ok=True)
    
    for name in names:
        prs = Presentation(TEMPLATE_FILE)
        slide = prs.slides[0]
        
        replace_placeholder(slide, name)
        
        output_pptx = OUTPUT_DIR / f"دعوت‌نامه-{name.replace(' ', '-')}.pptx"
        prs.save(output_pptx)
        
        save_and_convert(output_pptx, OUTPUT_DIR)


def compress_files(file_path: Path):
    """Compress generated images into a ZIP archive."""
    if file_path.exists():
        subprocess.run(["zip", "-r", ZIP_NAME, str(file_path)], check=True)
    else:
        print(f"⚠️ Directory not found: {file_path}")


if __name__ == "__main__":
    names = load_names(NAMES_FILE)
    create_invites(names)
    compress_files(OUTPUT_DIR)
