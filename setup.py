"""
setup.py  —  Inview URL Finder
================================
Single entry-point for first-time setup AND running the tool.

Usage
-----
  python setup.py          # install deps + run everything
  python setup.py --run    # skip install check, just run
  python setup.py --install-only
"""

import sys
import subprocess
import importlib
import os
import threading
import time
import textwrap
from pathlib import Path

# ── Minimum Python version ───────────────────────────────────────────────────
if sys.version_info < (3, 10):
    print("❌  Python 3.10+ required. You have", sys.version)
    sys.exit(1)

REQUIRED = [
    "selenium",
    "pandas",
    "openpyxl",
    "rich",
]

DASHBOARD_PORT = 8000
DASHBOARD_FILE = "dashboard.html"
SCRAPER_FILE   = "inview_url_finder.py"


# ─────────────────────────────────────────────────────────────────────────────
# 1. DEPENDENCY INSTALLER
# ─────────────────────────────────────────────────────────────────────────────

def check_and_install():
    missing = []
    for pkg in REQUIRED:
        try:
            importlib.import_module(pkg if pkg != "openpyxl" else "openpyxl")
        except ImportError:
            missing.append(pkg)

    if not missing:
        print("✅  All dependencies already installed.")
        return

    print(f"📦  Installing missing packages: {', '.join(missing)}")
    subprocess.check_call([
        sys.executable, "-m", "pip", "install", "--quiet", *missing
    ])
    print("✅  Dependencies installed.\n")


# ─────────────────────────────────────────────────────────────────────────────
# 2. WRITE dashboard.html (embedded so repo stays single-folder)
# ─────────────────────────────────────────────────────────────────────────────

DASHBOARD_HTML = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width,initial-scale=1.0"/>
<title>Inview URL Finder — Dashboard</title>
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;500;600&family=IBM+Plex+Sans:wght@300;400;500;600&display=swap');
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
:root{
  --bg:#0c0e13;--surface:#13161e;--border:#1e2230;
  --accent:#00e5a0;--accent2:#0099ff;--danger:#ff4d6d;--warn:#ffb830;
  --text:#d4dbe8;--muted:#5a6278;
  --mono:'IBM Plex Mono',monospace;--sans:'IBM Plex Sans',sans-serif;
}
html,body{height:100%;background:var(--bg);color:var(--text);font-family:var(--sans);font-size:14px;line-height:1.6}
.shell{display:grid;grid-template-rows:56px 1fr;height:100vh;overflow:hidden}
header{display:flex;align-items:center;gap:16px;padding:0 28px;border-bottom:1px solid var(--border);background:var(--surface)}
header .logo{font-family:var(--mono);font-size:13px;font-weight:600;color:var(--accent);letter-spacing:.08em;text-transform:uppercase}
header .sep{color:var(--border);font-size:20px}
header .subtitle{font-size:12px;color:var(--muted);font-family:var(--mono)}
.pulse{margin-left:auto;display:flex;align-items:center;gap:8px;font-size:11px;font-family:var(--mono);color:var(--muted)}
.pulse-dot{width:7px;height:7px;border-radius:50%;background:var(--accent);animation:blink 1.4s ease-in-out infinite}
.pulse-dot.idle{background:var(--muted);animation:none}
@keyframes blink{0%,100%{opacity:1}50%{opacity:.2}}
.body{display:grid;grid-template-columns:280px 1fr;overflow:hidden}
aside{border-right:1px solid var(--border);padding:24px 20px;display:flex;flex-direction:column;gap:20px;overflow-y:auto}
.stat-card{background:var(--surface);border:1px solid var(--border);border-radius:8px;padding:16px 18px;position:relative;overflow:hidden}
.stat-card::before{content:'';position:absolute;top:0;left:0;width:3px;height:100%;border-radius:8px 0 0 8px}
.stat-card.green::before{background:var(--accent)}
.stat-card.blue::before{background:var(--accent2)}
.stat-card.red::before{background:var(--danger)}
.stat-card.warn::before{background:var(--warn)}
.stat-label{font-size:10px;font-family:var(--mono);text-transform:uppercase;letter-spacing:.1em;color:var(--muted);margin-bottom:6px}
.stat-value{font-size:32px;font-family:var(--mono);font-weight:600;line-height:1}
.stat-card.green .stat-value{color:var(--accent)}
.stat-card.blue  .stat-value{color:var(--accent2)}
.stat-card.red   .stat-value{color:var(--danger)}
.stat-card.warn  .stat-value{color:var(--warn)}
.stat-sub{font-size:11px;color:var(--muted);margin-top:4px;font-family:var(--mono)}
.prog-section{margin-top:4px}
.prog-label{display:flex;justify-content:space-between;font-size:11px;font-family:var(--mono);color:var(--muted);margin-bottom:8px}
.prog-track{height:6px;background:var(--border);border-radius:99px;overflow:hidden}
.prog-fill{height:100%;border-radius:99px;background:linear-gradient(90deg,var(--accent2),var(--accent));transition:width .6s ease;width:0%}
.current-box{background:var(--surface);border:1px solid var(--border);border-radius:8px;padding:14px 16px}
.current-box .clabel{font-size:10px;font-family:var(--mono);text-transform:uppercase;letter-spacing:.1em;color:var(--muted);margin-bottom:6px}
.current-box .ctitle{font-size:12px;color:var(--text);font-family:var(--mono);word-break:break-word;line-height:1.5}
.timing-box{font-size:11px;font-family:var(--mono);color:var(--muted);display:flex;flex-direction:column;gap:4px}
.timing-box span{display:flex;justify-content:space-between}
.timing-box b{color:var(--text)}
main{display:flex;flex-direction:column;overflow:hidden}
.tabs{display:flex;gap:0;border-bottom:1px solid var(--border);padding:0 24px;background:var(--surface)}
.tab{padding:14px 20px;font-size:12px;font-family:var(--mono);letter-spacing:.05em;cursor:pointer;border-bottom:2px solid transparent;color:var(--muted);transition:all .2s;display:flex;align-items:center;gap:8px}
.tab:hover{color:var(--text)}
.tab.active{color:var(--text);border-bottom-color:var(--accent)}
.tab .badge{font-size:10px;padding:1px 7px;border-radius:99px;font-weight:600}
.tab.active .badge-found{background:rgba(0,229,160,.15);color:var(--accent)}
.tab.active .badge-nf{background:rgba(255,77,109,.15);color:var(--danger)}
.tab .badge-found,.tab .badge-nf{background:var(--border);color:var(--muted)}
.panel{display:none;flex-direction:column;overflow:hidden;flex:1}
.panel.active{display:flex}
.search-bar{padding:14px 24px;border-bottom:1px solid var(--border)}
.search-bar input{width:100%;background:var(--surface);border:1px solid var(--border);border-radius:6px;padding:8px 14px;font-family:var(--mono);font-size:12px;color:var(--text);outline:none;transition:border-color .2s}
.search-bar input:focus{border-color:var(--accent2)}
.search-bar input::placeholder{color:var(--muted)}
.table-wrap{flex:1;overflow-y:auto;padding:0 24px 24px}
table{width:100%;border-collapse:collapse;font-size:12px;font-family:var(--mono)}
thead th{position:sticky;top:0;background:var(--bg);padding:12px 10px;text-align:left;font-size:10px;text-transform:uppercase;letter-spacing:.1em;color:var(--muted);border-bottom:1px solid var(--border);z-index:1}
tbody tr{border-bottom:1px solid var(--border);transition:background .15s}
tbody tr:hover{background:var(--surface)}
tbody td{padding:10px;vertical-align:top;color:var(--text);line-height:1.5}
td.row-num{color:var(--muted);width:60px}
td.title-cell{max-width:340px;word-break:break-word}
td.url-cell{max-width:400px;word-break:break-all}
td.url-cell a{color:var(--accent2);text-decoration:none}
td.url-cell a:hover{text-decoration:underline}
.empty{padding:60px 24px;text-align:center;color:var(--muted);font-family:var(--mono);font-size:12px}
::-webkit-scrollbar{width:6px}
::-webkit-scrollbar-track{background:transparent}
::-webkit-scrollbar-thumb{background:var(--border);border-radius:3px}
::-webkit-scrollbar-thumb:hover{background:var(--muted)}
</style>
</head>
<body>
<div class="shell">
  <header>
    <span class="logo">Inview Finder</span>
    <span class="sep">|</span>
    <span class="subtitle">URL Discovery Dashboard</span>
    <div class="pulse" id="pulse">
      <div class="pulse-dot idle" id="pulseDot"></div>
      <span id="pulseLabel">waiting for scraper to start…</span>
    </div>
  </header>
  <div class="body">
    <aside>
      <div class="stat-card blue"><div class="stat-label">Total Rows</div><div class="stat-value" id="statTotal">—</div></div>
      <div class="stat-card green"><div class="stat-label">Found</div><div class="stat-value" id="statFound">—</div><div class="stat-sub" id="statFoundPct">—</div></div>
      <div class="stat-card red"><div class="stat-label">Not Found</div><div class="stat-value" id="statNF">—</div><div class="stat-sub" id="statNFPct">—</div></div>
      <div class="stat-card warn"><div class="stat-label">Pending</div><div class="stat-value" id="statPending">—</div></div>
      <div class="prog-section">
        <div class="prog-label"><span>Progress</span><span id="progPct">0%</span></div>
        <div class="prog-track"><div class="prog-fill" id="progFill"></div></div>
      </div>
      <div class="current-box">
        <div class="clabel">Currently processing</div>
        <div class="ctitle" id="currentTitle">—</div>
      </div>
      <div class="timing-box">
        <span><span>Started</span><b id="tStarted">—</b></span>
        <span><span>Updated</span><b id="tUpdated">—</b></span>
      </div>
    </aside>
    <main>
      <div class="tabs">
        <div class="tab active" onclick="switchTab('found')" id="tabFound">Found <span class="badge badge-found" id="badgeFound">0</span></div>
        <div class="tab" onclick="switchTab('nf')" id="tabNF">Not Found <span class="badge badge-nf" id="badgeNF">0</span></div>
      </div>
      <div class="panel active" id="panelFound">
        <div class="search-bar"><input type="text" id="searchFound" placeholder="Filter by title or URL…" oninput="renderFound()"/></div>
        <div class="table-wrap">
          <table><thead><tr><th>Row</th><th>Title</th><th>URL</th></tr></thead><tbody id="bodyFound"></tbody></table>
          <div class="empty" id="emptyFound" style="display:none">No found rows yet.</div>
        </div>
      </div>
      <div class="panel" id="panelNF">
        <div class="search-bar"><input type="text" id="searchNF" placeholder="Filter by title…" oninput="renderNF()"/></div>
        <div class="table-wrap">
          <table><thead><tr><th>Row</th><th>Title</th></tr></thead><tbody id="bodyNF"></tbody></table>
          <div class="empty" id="emptyNF" style="display:none">No not-found rows yet.</div>
        </div>
      </div>
    </main>
  </div>
</div>
<script>
let state={found_rows:[],not_found_rows:[],total:0,found_count:0,not_found_count:0,pending_count:0};
function fmt(iso){if(!iso||iso==='—')return'—';try{return new Date(iso).toLocaleTimeString()}catch{return iso}}
function pct(n,t){if(!t)return'0.0%';return(n/t*100).toFixed(1)+'%'}
function switchTab(n){
  document.getElementById('tabFound').classList.toggle('active',n==='found');
  document.getElementById('tabNF').classList.toggle('active',n==='nf');
  document.getElementById('panelFound').classList.toggle('active',n==='found');
  document.getElementById('panelNF').classList.toggle('active',n==='nf');
}
function esc(s){return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;')}
function renderFound(){
  const q=document.getElementById('searchFound').value.toLowerCase();
  const rows=state.found_rows.filter(r=>!q||r.title.toLowerCase().includes(q)||r.url.toLowerCase().includes(q));
  document.getElementById('bodyFound').innerHTML=rows.map(r=>`<tr><td class="row-num">${r.row}</td><td class="title-cell">${esc(r.title)}</td><td class="url-cell"><a href="${esc(r.url)}" target="_blank">${esc(r.url)}</a></td></tr>`).join('');
  document.getElementById('emptyFound').style.display=rows.length?'none':'block';
}
function renderNF(){
  const q=document.getElementById('searchNF').value.toLowerCase();
  const rows=state.not_found_rows.filter(r=>!q||r.title.toLowerCase().includes(q));
  document.getElementById('bodyNF').innerHTML=rows.map(r=>`<tr><td class="row-num">${r.row}</td><td class="title-cell">${esc(r.title)}</td></tr>`).join('');
  document.getElementById('emptyNF').style.display=rows.length?'none':'block';
}
function applyState(data){
  state=data;
  const{total,found_count,not_found_count,pending_count}=data;
  const done=found_count+not_found_count;
  const p=total?done/total*100:0;
  document.getElementById('statTotal').textContent=total??'—';
  document.getElementById('statFound').textContent=found_count??'—';
  document.getElementById('statFoundPct').textContent=pct(found_count,total);
  document.getElementById('statNF').textContent=not_found_count??'—';
  document.getElementById('statNFPct').textContent=pct(not_found_count,total);
  document.getElementById('statPending').textContent=pending_count??'—';
  document.getElementById('progFill').style.width=p.toFixed(1)+'%';
  document.getElementById('progPct').textContent=p.toFixed(1)+'%';
  document.getElementById('currentTitle').textContent=data.current_title??'—';
  document.getElementById('tStarted').textContent=fmt(data.started_at);
  document.getElementById('tUpdated').textContent=fmt(data.updated_at);
  document.getElementById('badgeFound').textContent=found_count;
  document.getElementById('badgeNF').textContent=not_found_count;
  renderFound();renderNF();
}
async function poll(){
  try{
    const r=await fetch('progress.json?_='+Date.now());
    if(!r.ok)throw new Error();
    applyState(await r.json());
    document.getElementById('pulseDot').classList.remove('idle');
    document.getElementById('pulseLabel').textContent='live · updates every 2s';
  }catch{
    document.getElementById('pulseDot').classList.add('idle');
    document.getElementById('pulseLabel').textContent='waiting for scraper to start…';
  }
}
poll();setInterval(poll,2000);
</script>
</body>
</html>"""


# ─────────────────────────────────────────────────────────────────────────────
# 3. WRITE inview_url_finder.py (embedded)
# ─────────────────────────────────────────────────────────────────────────────

SCRAPER_PY = r'''import time
import re
import json
import logging
import pandas as pd
from pathlib import Path
from urllib.parse import urlparse
from dataclasses import dataclass, field
from typing import Optional
from datetime import datetime

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    StaleElementReferenceException, WebDriverException
)

try:
    from rich.progress import (
        Progress, SpinnerColumn, BarColumn, TextColumn,
        TimeElapsedColumn, TimeRemainingColumn, MofNCompleteColumn,
    )
    from rich.console import Console
    from rich.logging import RichHandler
    RICH = True
except ImportError:
    RICH = False

# ---------------- CONFIG ----------------
BASE_URL      = "https://www.inview.nl/zoeken"
INPUT_FILE    = "T_C.xlsx"
OUTPUT_FILE   = "output_TC.xlsx"
PROGRESS_JSON = "progress.json"

SEARCH_WAIT            = 3
REDIRECT_WAIT          = 3
BACK_WAIT              = 2
URL_STABILIZE_POLLS    = 10
URL_STABILIZE_INTERVAL = 0.5
SLUG_WAIT_TIMEOUT      = 10.0
SLUG_WAIT_INTERVAL     = 0.5
MAX_RETRIES            = 2

if RICH:
    console = Console()
    logging.basicConfig(
        level=logging.INFO,
        format="%(message)s",
        handlers=[RichHandler(console=console, show_path=False, markup=True)],
    )
else:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
    )
log = logging.getLogger(__name__)


def slugify(text: str) -> str:
    text = text.lower()
    text = re.sub(r"[^a-z0-9\s-]", " ", text)
    text = re.sub(r"\s+", "-", text)
    text = re.sub(r"-+", "-", text)
    return text.strip("-")

def get_slug_from_url(url: str) -> str:
    return urlparse(url).path.strip("/").split("/")[-1]

def slugs_match(expected: str, actual: str) -> bool:
    shorter = min(expected, actual, key=len)
    if len(shorter) < 8:
        return expected == actual
    return expected in actual or actual in expected

def load_or_create_output(input_path: str, output_path: str) -> pd.DataFrame:
    out = Path(output_path)
    if out.exists():
        log.info("Output file found — resuming from %s", output_path)
        df = pd.read_excel(output_path)
        if "found_url" not in df.columns:
            df["found_url"] = ""
    else:
        log.info("No output file — starting fresh from %s", input_path)
        df = pd.read_excel(input_path)
        df["found_url"] = ""
    df["found_url"] = df["found_url"].astype(object)
    return df

def already_done(value) -> bool:
    val = str(value).strip().lower()
    return pd.notna(value) and val not in ("", "nan", "not found")

def save_progress(df: pd.DataFrame, output_path: str) -> None:
    df.to_excel(output_path, index=False)

def write_progress_json(df, title_col, current_row, current_title, started_at):
    total = len(df)
    found = [
        {"row": int(i), "title": str(r[title_col]), "url": str(r["found_url"])}
        for i, r in df.iterrows() if already_done(r["found_url"])
    ]
    not_found = [
        {"row": int(i), "title": str(r[title_col])}
        for i, r in df.iterrows()
        if str(r["found_url"]).strip().lower() == "not found"
    ]
    pending = total - len(found) - len(not_found)
    Path(PROGRESS_JSON).write_text(json.dumps({
        "total": total,
        "found_count": len(found),
        "not_found_count": len(not_found),
        "pending_count": pending,
        "current_row": current_row,
        "current_title": current_title,
        "started_at": started_at,
        "updated_at": datetime.now().isoformat(timespec="seconds"),
        "found_rows": found,
        "not_found_rows": not_found,
    }, ensure_ascii=False, indent=2))


@dataclass
class InviewScraper:
    driver: webdriver.Chrome
    wait: WebDriverWait = field(init=False)

    def __post_init__(self):
        self.wait = WebDriverWait(self.driver, 30)

    def go_to_search(self):
        self.driver.get(BASE_URL)

    def search(self, title: str):
        self.go_to_search()
        box = self.wait.until(EC.presence_of_element_located(
            (By.CSS_SELECTOR, "input[data-testid='search-bar-input-field']")
        ))
        box.clear()
        box.send_keys(title)
        box.send_keys(Keys.ENTER)
        time.sleep(SEARCH_WAIT)

    def get_commentaar_links(self) -> list:
        clusters = self.driver.find_elements(
            By.CSS_SELECTOR, "[data-e2e-cluster-name='Commentaar']"
        )
        return clusters[0].find_elements(By.TAG_NAME, "a") if clusters else []

    def _wait_for_slug_url(self) -> str:
        deadline = time.time() + SLUG_WAIT_TIMEOUT
        last_url = self.driver.current_url
        while time.time() < deadline:
            time.sleep(SLUG_WAIT_INTERVAL)
            current = self.driver.current_url
            if current != last_url:
                last_url = current
                continue
            slug = get_slug_from_url(current)
            if not re.fullmatch(r"[a-z0-9]{20,}", slug):
                log.info("  ⏳ Slug resolved: %s", slug)
                return current
            last_url = current
        log.warning("  ⏰ Timed out waiting for slug — returning: %s", last_url)
        return last_url

    def back_to_results(self):
        self.driver.back()
        time.sleep(BACK_WAIT)

    def _snapshot_hrefs(self, links: list) -> list[str]:
        hrefs = []
        for link in links:
            try:
                href = link.get_attribute("href")
                if href:
                    hrefs.append(href)
            except StaleElementReferenceException:
                pass
        return hrefs

    def _check_href(self, href: str, expected_slug: str) -> Optional[str]:
        try:
            self.driver.get(href)
            time.sleep(REDIRECT_WAIT)
            final_url = self._wait_for_slug_url()
            slug = get_slug_from_url(final_url)
            log.info("  → checking slug/id: %s", slug)

            if re.fullmatch(r"[a-z0-9]{20,}", slug):
                source_slug = get_slug_from_url(href)
                if slugs_match(expected_slug, source_slug):
                    log.info("  ✅ Accepted (ID-based, href matched): %s", final_url)
                    return final_url
                log.info("  ⚠️  ID-based but href slug mismatch — skipping")
                self.back_to_results()
                return None

            if slugs_match(expected_slug, slug):
                log.info("  ✅ Match: %s", final_url)
                return final_url

            self.back_to_results()
            return None

        except WebDriverException as e:
            log.error("  WebDriver error on %s: %s", href, e)
            self.back_to_results()
        return None

    def find_url_for_title(self, title: str) -> str:
        expected_slug = slugify(title)
        for attempt in range(1, MAX_RETRIES + 1):
            log.info("Searching (attempt %d): %s", attempt, title)
            self.search(title)
            links = self.get_commentaar_links()
            if not links:
                log.warning("No Commentaar cluster (attempt %d)", attempt)
                continue
            for href in self._snapshot_hrefs(links):
                result = self._check_href(href, expected_slug)
                if result is not None:
                    return result
            log.warning("No match on attempt %d for: %s", attempt, title)
        return "not found"


def main():
    df = load_or_create_output(INPUT_FILE, OUTPUT_FILE)
    title_col = df.columns[1]
    total     = len(df)
    pending   = sum(1 for v in df["found_url"] if not already_done(v))
    completed = total - pending
    log.info("Rows: %d total, %d already done, %d to process", total, completed, pending)

    started_at = datetime.now().isoformat(timespec="seconds")
    write_progress_json(df, title_col, -1, "—", started_at)

    driver  = webdriver.Chrome()
    scraper = InviewScraper(driver=driver)

    def run(progress=None, task=None):
        scraper.go_to_search()
        input("👉 Login manually, then press ENTER...")
        for i, row in df.iterrows():
            if already_done(row["found_url"]):
                if progress:
                    progress.advance(task)
                continue
            title = str(row[title_col])
            write_progress_json(df, title_col, int(i), title, started_at)
            result = scraper.find_url_for_title(title)
            df.at[i, "found_url"] = result
            save_progress(df, OUTPUT_FILE)
            write_progress_json(df, title_col, int(i), title, started_at)
            log.info("Row %d saved → %s", i, result)
            if progress:
                progress.advance(task)

    try:
        if RICH:
            with Progress(
                SpinnerColumn(),
                TextColumn("[bold cyan]{task.description}"),
                BarColumn(bar_width=40),
                MofNCompleteColumn(),
                TextColumn("[green]{task.percentage:>5.1f}%"),
                TimeElapsedColumn(),
                TimeRemainingColumn(),
                console=console,
                transient=False,
            ) as progress:
                task = progress.add_task("Processing rows…", total=total, completed=completed)
                run(progress, task)
        else:
            run()
    except KeyboardInterrupt:
        log.warning("Interrupted — progress saved up to last completed row.")
    finally:
        save_progress(df, OUTPUT_FILE)
        write_progress_json(df, title_col, -1, "Done", started_at)
        driver.quit()
        log.info("🚀 DONE")

if __name__ == "__main__":
    main()
'''


# ─────────────────────────────────────────────────────────────────────────────
# 4. DASHBOARD SERVER (runs in background thread)
# ─────────────────────────────────────────────────────────────────────────────

def start_dashboard_server(port: int):
    import http.server
    import socketserver

    class QuietHandler(http.server.SimpleHTTPRequestHandler):
        def log_message(self, *args): pass  # suppress request logs

    def serve():
        with socketserver.TCPServer(("", port), QuietHandler) as httpd:
            httpd.serve_forever()

    t = threading.Thread(target=serve, daemon=True)
    t.start()


def open_browser(port: int):
    import webbrowser
    time.sleep(1.2)
    webbrowser.open(f"http://localhost:{port}/{DASHBOARD_FILE}")


# ─────────────────────────────────────────────────────────────────────────────
# 5. MAIN
# ─────────────────────────────────────────────────────────────────────────────

def main():
    import argparse
    parser = argparse.ArgumentParser(description="Inview URL Finder — setup & run")
    parser.add_argument("--run",          action="store_true", help="Skip install check")
    parser.add_argument("--install-only", action="store_true", help="Only install deps")
    args = parser.parse_args()

    # ── Always write the latest embedded files ────────────────────────────────
    Path(DASHBOARD_FILE).write_text(DASHBOARD_HTML)
    Path(SCRAPER_FILE).write_text(SCRAPER_PY, encoding="utf-8")
    print(f"✅  Wrote {DASHBOARD_FILE} and {SCRAPER_FILE}")

    if not args.run:
        check_and_install()

    if args.install_only:
        print("✅  Install complete. Run  python setup.py  to start.")
        return

    # Check input file exists
    if not Path("T_C.xlsx").exists():
        print("\n❌  T_C.xlsx not found in the current directory.")
        print("    Place your input Excel file here and re-run.\n")
        sys.exit(1)

    # ── Start dashboard HTTP server ───────────────────────────────────────────
    start_dashboard_server(DASHBOARD_PORT)
    print(f"\n🌐  Dashboard → http://localhost:{DASHBOARD_PORT}/{DASHBOARD_FILE}")
    threading.Thread(target=open_browser, args=(DASHBOARD_PORT,), daemon=True).start()

    print("🚀  Starting scraper…\n")
    time.sleep(0.5)

    # ── Run the scraper in-process ────────────────────────────────────────────
    import importlib.util, types

    spec   = importlib.util.spec_from_file_location("inview_url_finder", SCRAPER_FILE)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    module.main()


if __name__ == "__main__":
    main()
