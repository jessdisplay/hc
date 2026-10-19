"""Read each audited document's own version block into versions.json.

Harmony's documents end with a small table: 'Approved By: ... Version: ...
Approval Date: ... Next Scheduled Review: ...'. The 16 Oct calendar took review
dates from it but not the version, approver or approval date. This reads those
three straight from the files, so the checklist shows what each document says
about itself. Files with no block are left out; nothing is guessed.

    python3 -I read_versions.py SOURCE_DIR LISTING.json [LISTING.json ...]

SOURCE_DIR holds the downloaded Drive folders; each LISTING is `rclone lsjson -R`
of one folder, saved as <name>.json with its files under SOURCE_DIR/<name>/.
"""
import json
import re
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
FIELD = {
    "approved_by": r"Approved\s*By\s*:\s*(.+?)\s*(?=Version|Approval|Next|$)",
    "version": r"Version\s*(?:No\.?|Number)?\s*:?\s*([0-9]+(?:\.[0-9]+)?)\b",
    "approval_date": r"Approval\s*Date\s*:\s*([A-Za-z]+\s+\d{4}|\d{1,2}/\d{1,2}/\d{2,4})",
}


def lines(path):
    z = zipfile.ZipFile(path)
    out = []
    for name in z.namelist():
        if re.match(r"word/(document|footer\d*|header\d*)\.xml$", name):
            for p in ET.fromstring(z.read(name)).iter(W + "p"):
                t = "".join(x.text or "" for x in p.iter(W + "t")).strip()
                if t:
                    out.append(t)
    return out


def version_block(text_lines):
    """The fields from the first line or adjacent pair that carries an 'Approved By'."""
    for i, line in enumerate(text_lines):
        if re.search(r"Approved\s*By\s*:", line, re.I):
            window = " ".join(text_lines[i:i + 4])
            found = {}
            for key, pattern in FIELD.items():
                m = re.search(pattern, window, re.I)
                if m:
                    found[key] = m.group(1).strip()
            if found:
                return found
    return {}


def main(root, listings):
    root = Path(root)
    out, missing = {}, []
    for listing in listings:
        folder = root / Path(listing).stem
        for f in json.load(open(listing)):
            path = folder / f["Path"]
            try:
                if path.suffix.lower() != ".docx" or not path.exists():
                    continue
            except OSError:  # a name longer than the filesystem allows never downloaded
                missing.append(f["Path"][:80])
                continue
            block = version_block(lines(path))
            if block:
                out[f["ID"]] = block
            else:
                missing.append(f["Path"])
    dest = Path(__file__).with_name("versions.json")
    dest.write_text(json.dumps(out, indent=1, sort_keys=True))
    print(f"{len(out)} documents carry a version block; {len(missing)} do not -> {dest}")


if __name__ == "__main__":
    if len(sys.argv) < 3:
        raise SystemExit(__doc__)
    main(sys.argv[1], sys.argv[2:])
