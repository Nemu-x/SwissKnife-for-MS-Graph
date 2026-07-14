#!/usr/bin/env python3
"""Generates the SwissKnife for MS Graph release landing page (GitHub Pages).

Run by .github/workflows/pages.yml on every deploy. Pulls recent releases via
`gh` (GH_TOKEN provided by the workflow) so the page always shows the latest
version, download buttons, and release notes without manual edits. Output is
self-contained: inline CSS, no external assets.

Usage: python3 packaging/site/generate.py --out site
"""

import argparse
import html
import json
import re
import subprocess
from datetime import datetime, timezone

REPO = "Nemu-x/SwissKnife-for-MS-Graph"
PAGES = "https://nemu-x.github.io/SwissKnife-for-MS-Graph"
DL = f"https://github.com/{REPO}/releases/latest/download"

# Download buttons for the latest release (stable asset names, version-free).
ASSETS = [
    ("Windows", "Installer (.exe)", "SwissKnifeGraph-windows-amd64-installer.exe"),
    ("Windows", "Portable (.exe)", "SwissKnifeGraph-windows-amd64.exe"),
    ("macOS", "Universal (.zip)", "SwissKnifeGraph-macos-universal.zip"),
    ("Linux", "Tarball (.tar.gz)", "SwissKnifeGraph-linux-amd64.tar.gz"),
    ("Linux", "Debian (.deb)", "SwissKnifeGraph-linux-amd64.deb"),
    ("Linux", "RPM (.rpm)", "SwissKnifeGraph-linux-amd64.rpm"),
]


def gh_releases(limit=5):
    try:
        out = subprocess.run(
            ["gh", "release", "list", "--repo", REPO, "--limit", str(limit),
             "--json", "tagName,publishedAt,name,isPrerelease"],
            capture_output=True, text=True, check=True,
            encoding="utf-8", errors="replace").stdout
        rels = [r for r in json.loads(out) if not r.get("isPrerelease")]
    except Exception:
        return []
    for r in rels:
        try:
            body = subprocess.run(
                ["gh", "release", "view", r["tagName"], "--repo", REPO, "--json", "body"],
                capture_output=True, text=True, check=True,
                encoding="utf-8", errors="replace").stdout
            r["body"] = json.loads(body).get("body", "")
        except Exception:
            r["body"] = ""
    return rels


def md_lite(text):
    """Tiny, safe markdown subset for release notes: headings, bullets, links, code."""
    out, in_list = [], False
    for raw in text.splitlines():
        line = html.escape(raw.rstrip())
        line = re.sub(r"`([^`]+)`", r"<code>\1</code>", line)
        line = re.sub(r"\*\*([^*]+)\*\*", r"<strong>\1</strong>", line)
        line = re.sub(r"\[([^\]]+)\]\((https?://[^)]+)\)", r'<a href="\2">\1</a>', line)
        line = re.sub(r'(?<!href=")(https?://github\.com/\S+)', r'<a href="\1">\1</a>', line)
        if re.match(r"^\s*[-*] ", line):
            if not in_list:
                out.append("<ul>"); in_list = True
            out.append("<li>" + re.sub(r"^\s*[-*] ", "", line) + "</li>")
            continue
        if in_list:
            out.append("</ul>"); in_list = False
        if re.match(r"^#{1,6}\s", line):
            out.append("<h4>" + re.sub(r"^#{1,6}\s", "", line) + "</h4>")
        elif line.strip():
            out.append("<p>" + line + "</p>")
    if in_list:
        out.append("</ul>")
    return "\n".join(out)


def fmt_date(iso):
    try:
        return datetime.fromisoformat(iso.replace("Z", "+00:00")).strftime("%b %d, %Y")
    except Exception:
        return iso


def render(rels):
    latest = rels[0] if rels else None
    version = latest["tagName"] if latest else "—"

    buttons = "\n".join(
        f'<a class="dl" href="{DL}/{fname}"><span class="os">{os}</span>{label}</a>'
        for os, label, fname in ASSETS
    )

    notes = []
    for r in rels:
        notes.append(f"""
        <details class="rel" {'open' if r is latest else ''}>
          <summary><span class="tag">{html.escape(r['tagName'])}</span>
            <span class="relname">{html.escape(r.get('name') or '')}</span>
            <time>{fmt_date(r.get('publishedAt',''))}</time></summary>
          <div class="body">{md_lite(r.get('body','') or 'No release notes.')}</div>
        </details>""")
    notes_html = "\n".join(notes) if notes else "<p class='muted'>No releases published yet.</p>"

    year = datetime.now(timezone.utc).year
    return f"""<!doctype html>
<html lang="en"><head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>SwissKnife for MS Graph — Downloads</title>
<style>
  :root {{
    --bg:#0f1117; --elev:#171a23; --elev2:#1e222d; --border:#2a2f3a;
    --text:#e6e8ee; --dim:#9aa2b1; --accent:#6366f1; --accent2:#818cf8; --ok:#22c55e;
  }}
  * {{ box-sizing:border-box; }}
  body {{ margin:0; background:var(--bg); color:var(--text);
    font-family:'Inter',system-ui,-apple-system,'Segoe UI',Roboto,sans-serif; line-height:1.6; }}
  a {{ color:var(--accent2); text-decoration:none; }}
  .wrap {{ max-width:900px; margin:0 auto; padding:2.5rem 1.25rem 4rem; }}
  header {{ text-align:center; padding:3rem 0 1.5rem; }}
  .logo {{ font-size:3.2rem; }}
  h1 {{ font-size:2.1rem; margin:.5rem 0 .25rem; }}
  .tagline {{ color:var(--dim); font-size:1.05rem; }}
  .ver {{ display:inline-block; margin-top:1rem; padding:.35rem .9rem; border-radius:999px;
    background:linear-gradient(135deg,var(--accent),var(--accent2)); color:#fff; font-weight:600; font-size:.9rem; }}
  .grid {{ display:grid; grid-template-columns:repeat(auto-fit,minmax(220px,1fr)); gap:.75rem; margin:2rem 0; }}
  .dl {{ display:flex; flex-direction:column; gap:.15rem; padding:1rem 1.2rem; border-radius:14px;
    background:var(--elev); border:1px solid var(--border); color:var(--text); transition:.15s; }}
  .dl:hover {{ border-color:var(--accent); background:var(--elev2); transform:translateY(-2px); }}
  .dl .os {{ font-size:.72rem; letter-spacing:.08em; text-transform:uppercase; color:var(--accent2); font-weight:700; }}
  section h2 {{ font-size:1.3rem; border-bottom:1px solid var(--border); padding-bottom:.5rem; margin-top:2.5rem; }}
  .rel {{ background:var(--elev); border:1px solid var(--border); border-radius:12px; margin:.75rem 0; padding:.5rem 1rem; }}
  .rel summary {{ cursor:pointer; display:flex; align-items:center; gap:.6rem; font-weight:600; list-style:none; }}
  .rel summary::-webkit-details-marker {{ display:none; }}
  .tag {{ background:var(--elev2); color:var(--accent2); padding:.1rem .5rem; border-radius:6px; font-size:.85rem; font-family:ui-monospace,monospace; }}
  .relname {{ color:var(--dim); font-weight:400; }}
  .rel time {{ margin-left:auto; color:var(--dim); font-size:.85rem; }}
  .body {{ padding:.5rem 0 .25rem; color:var(--dim); }}
  .body h4 {{ color:var(--text); margin:.8rem 0 .3rem; }}
  .body code {{ background:var(--elev2); padding:.1rem .35rem; border-radius:5px; font-size:.85em; }}
  .install {{ background:var(--elev); border:1px solid var(--border); border-radius:12px; padding:1rem 1.25rem; }}
  .install code {{ display:block; background:#0b0d12; padding:.6rem .8rem; border-radius:8px; overflow-x:auto;
    font-family:ui-monospace,monospace; font-size:.85rem; color:var(--ok); margin:.4rem 0; }}
  footer {{ text-align:center; color:var(--dim); margin-top:3rem; font-size:.85rem; }}
</style></head>
<body><div class="wrap">
  <header>
    <div class="logo">🗡️</div>
    <h1>SwissKnife for MS Graph</h1>
    <p class="tagline">A clean, fast, cross-platform Microsoft Graph desktop client for IT admins.</p>
    <div class="ver">Latest: {html.escape(version)}</div>
  </header>

  <section>
    <h2>Download</h2>
    <div class="grid">{buttons}</div>
  </section>

  <section>
    <h2>Install on Arch (AUR)</h2>
    <div class="install">
      <code>yay -S swissknife-graph-bin</code>
      <p style="color:var(--dim);margin:.3rem 0 0">Updates arrive automatically through your AUR helper.</p>
    </div>
  </section>

  <section>
    <h2>Release notes</h2>
    {notes_html}
  </section>

  <footer>
    <p><a href="https://github.com/{REPO}">GitHub</a> · MIT License · © {year} Nemu-x</p>
  </footer>
</div></body></html>"""


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default="site")
    args = ap.parse_args()
    import os
    os.makedirs(args.out, exist_ok=True)
    with open(os.path.join(args.out, "index.html"), "w", encoding="utf-8") as f:
        f.write(render(gh_releases()))
    print("wrote", os.path.join(args.out, "index.html"))


if __name__ == "__main__":
    main()
