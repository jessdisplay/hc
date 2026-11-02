# Harmony Care team hub, website and audit room

- `index.html` (the root, hc.riseos.care): opens the compliance tool in `tool/`, the one place for everything. The team hub that lived here until 2 Nov 2026 is in git history.
- `site/`: Harmony Care's public website and the intake form. `site/checklist/` now opens the tool's Full checklist.
- `audit-room/`: the Harmony Care x Rise guide for the CPG provisional
  certification audit, 18 to 19 August 2026. Built, see below.

## How to build

```bash
python3 build_guide_site.py
```

That reads `guide.template.html`, inlines every asset from `assets/` as a data
URI, and writes `audit-room/index.html`. Two reasons for one self-contained file: the
Artifact CSP blocks external hosts, and a single file drops onto any static
host without a build step.

**Edit `guide.template.html`, then rerun the script.** Never hand-edit
`audit-room/index.html`: the next build overwrites it.

## Brand

Palette 01 and the two brand faces are inherited from the Harmony Care brand
guide. Do not introduce colours outside the token set at the top of the
template. The illustrations are painted on paper, so they always sit on a cream
plate, in light and dark themes alike.

## Contents

| File | What it is |
|---|---|
| `guide.template.html` | The page. Edit this one. |
| `build_guide_site.py` | Inlines assets, writes `audit-room/index.html`. |
| `assets/` | Brand artwork and both marks. |
| `audit-room/index.html` | Built output. Generated, not edited. |
| `index.html` | The team hub. Edit this one directly. |
| `site/checklist/` | Every audited document and form by Practice Standard. Data from `tools/checklist/`. |
