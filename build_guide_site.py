#!/usr/bin/env python3
"""
Build "The audit room" into a single self-contained index.html.

Self-contained because the Artifact CSP blocks external hosts, and because a
one-file site drops straight onto any static host.

    python3 build_guide_site.py

Same pipeline as build_walkthrough.py: reads guide.template.html and inlines
every asset as a data URI, because the Artifact CSP blocks external hosts.
Edit the template, then rerun this — never edit the built file.
"""

import base64
import io
import re
from pathlib import Path

from PIL import Image

HERE = Path(__file__).parent
ASSETS = HERE / "assets"
TEMPLATE = HERE / "guide.template.html"
OUT = HERE / "index.html"

# Palette 01 warm cream. The illustrations are painted on paper, so flattening
# their alpha onto anything else puts a grey halo around the deckled edge.
PAPER = (251, 247, 239)


def data_uri(name, width, fmt="JPEG", quality=82, flatten=True, alpha_floor=0):
    im = Image.open(ASSETS / name)

    if alpha_floor:
        r, g, b, a = im.convert("RGBA").split()
        a = a.point(lambda v: 0 if v < alpha_floor else round((v - alpha_floor) * 255 / (255 - alpha_floor)))
        im = Image.merge("RGBA", (r, g, b, a))

    if flatten and im.mode in ("RGBA", "LA"):
        bg = Image.new("RGB", im.size, PAPER)
        bg.paste(im, mask=im.split()[-1])
        im = bg
    elif fmt == "JPEG":
        im = im.convert("RGB")

    im = im.resize((width, round(im.height * width / im.width)), Image.LANCZOS)

    buf = io.BytesIO()
    if fmt == "PNG":
        im.save(buf, "PNG", optimize=True)
        mime = "image/png"
    else:
        im.save(buf, "JPEG", quality=quality, optimize=True, progressive=True)
        mime = "image/jpeg"

    raw = buf.getvalue()
    print(f"  {name:28s} -> {im.size[0]}x{im.size[1]}  {len(raw)/1024:7.1f} KB")
    return f"data:{mime};base64," + base64.b64encode(raw).decode("ascii")


def rise_mark():
    """The Rise wordmark, inlined and recoloured to follow the theme."""
    svg = (ASSETS / "rise-wordmark.svg").read_text(encoding="utf-8")
    svg = re.sub(r'\s(width|height)="[^"]*"', "", svg, count=2)
    svg = re.sub(r'fill="#[0-9A-Fa-f]{6}"', 'fill="currentColor"', svg)
    svg = svg.replace("<svg", '<svg role="img" aria-label="Rise"', 1)
    return svg.strip()


def hook_url():
    """The deployed Apps Script /exec URL, or empty until it exists.

    Kept in a file rather than the template so switching the drop panel on is a
    one-line change with no HTML editing.
    """
    f = HERE / "hook-url.txt"
    url = f.read_text(encoding="utf-8").strip() if f.exists() else ""
    print(f"  intake hook: {url or 'not set, drop panel stays off'}")
    return url


def main():
    print("assets:")
    subs = {
        "{{EMBLEM}}": data_uri("emblem.png", 260, fmt="PNG", flatten=False, alpha_floor=96),
        "{{MURAL}}": data_uri("title-mural.png", 1600, quality=82),
        "{{LOUNGE}}": data_uri("mural-lounge.png", 2200, quality=80),
        "{{RISEMARK}}": rise_mark(),
        "{{HOOK}}": hook_url(),
    }

    html = TEMPLATE.read_text(encoding="utf-8")
    for token, value in subs.items():
        if token not in html:
            raise SystemExit(f"template is missing {token}")
        html = html.replace(token, value)

    leftover = re.findall(r"\{\{[A-Z]+\}\}", html)
    if leftover:
        raise SystemExit(f"unsubstituted tokens: {leftover}")

    OUT.write_text(html, encoding="utf-8")
    print(f"\nwrote {OUT.name}  {OUT.stat().st_size/1024/1024:.2f} MB")


if __name__ == "__main__":
    main()
