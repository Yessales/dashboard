"""
YES SALES INC. — WMS Map Auto-Deploy Script
============================================
매주 실행 → xlsx 읽기 → HTML 생성 → GitHub push → Pages 자동 배포

Setup:
  1. config.json 에 token 저장 (최초 1회)
  2. python update_wms.py 실행
  3. Task Scheduler 등록 (run_wms.bat 사용)
"""

import os, sys, json, base64, re, glob
from pathlib import Path
from collections import defaultdict
from datetime import datetime

# ── Config ──────────────────────────────────────────────────────────
REPO        = "Yessales/dashboard"
DEPLOY_PATH = "wms/index.html"
BRANCH      = "main"
WMS_FOLDER  = r"C:\Users\ajasu\OneDrive - Yes Sales Inc\7.Project\01. WMS\WMS Inv Balance"
SCRIPT_DIR  = Path(__file__).parent
CONFIG_FILE = SCRIPT_DIR / "config.json"

def get_token():
    """config.json 에서 token 읽기, 없으면 입력 받아서 저장"""
    if CONFIG_FILE.exists():
        with open(CONFIG_FILE) as f:
            cfg = json.load(f)
        if cfg.get('github_token'):
            return cfg['github_token']
    # First time setup
    print("\n[Setup] GitHub Personal Access Token을 입력하세요:")
    token = input("  Token (ghp_...): ").strip()
    with open(CONFIG_FILE, 'w') as f:
        json.dump({'github_token': token}, f)
    print(f"  Saved to {CONFIG_FILE}")
    return token

# ── Dependencies ─────────────────────────────────────────────────────
try:
    import openpyxl
except ImportError:
    os.system(f"{sys.executable} -m pip install openpyxl")
    import openpyxl

try:
    import requests
except ImportError:
    os.system(f"{sys.executable} -m pip install requests")
    import requests

# ── Step 1: Find xlsx files ──────────────────────────────────────────
def find_wms_files(folder):
    files = sorted(glob.glob(os.path.join(folder, "*.xlsx")))
    if len(files) < 2:
        raise FileNotFoundError(
            f"Need 2 xlsx files in:\n  {folder}\n"
            f"Found {len(files)}: {[os.path.basename(f) for f in files]}"
        )
    print(f"  PICO:  {os.path.basename(files[0])}")
    print(f"  BOYLE: {os.path.basename(files[1])}")
    return files[0], files[1]

# ── Step 2: Parse inventory ──────────────────────────────────────────
def parse_inventory(xlsx_path):
    wb   = openpyxl.load_workbook(xlsx_path, read_only=True, data_only=True)
    ws   = wb.active
    rows = list(ws.iter_rows(values_only=True))

    header_row = next(i for i,r in enumerate(rows) if r and 'Area Name' in str(r))
    headers    = rows[header_row]

    def col(name):
        return next(i for i,h in enumerate(headers) if h and name in str(h))

    ai = col('Area')
    li = col('Location')
    qi = col('Qty')
    si = col('Item No')
    ni = col('Item Name')

    loc_qty     = defaultdict(int)
    default_qty = 0
    sku_data    = defaultdict(lambda: {'n': '', 'l': [], 'd': 0})

    for row in rows[header_row + 1:]:
        if not row or row[ai] is None:
            continue
        area = str(row[ai]).strip()
        loc  = str(row[li]).strip() if row[li] else ''
        sku  = str(row[si]).strip() if row[si] else ''
        name = str(row[ni]).strip() if row[ni] else ''
        try:
            qty = int(row[qi] or 0)
        except:
            qty = 0

        if area == 'Total' or not sku:
            continue

        if sku:
            sku_data[sku]['n'] = name

        if area == 'Default Area' or loc == 'Default Location':
            default_qty += qty
            if sku:
                sku_data[sku]['d'] += qty
        else:
            loc_qty[loc] += qty
            if sku and loc:
                sku_data[sku]['l'].append({'loc': loc, 'qty': qty})

    wb.close()
    located = sum(loc_qty.values())

    return {
        'L':   dict(loc_qty),
        'D':   default_qty,
        'LT':  located,
        'GT':  located + default_qty,
        'SKU': {k: {'n': v['n'], 'l': v['l'], 'd': v['d']} for k, v in sku_data.items()}
    }

# ── Step 3: Generate HTML ────────────────────────────────────────────
def generate_html(pico, boyle, updated_at):
    template = SCRIPT_DIR / "wms_template.html"
    if not template.exists():
        raise FileNotFoundError(
            f"Template not found: {template}\n"
            f"wms_template.html 을 스크립트와 같은 폴더에 놓으세요."
        )

    with open(template, 'r', encoding='utf-8') as f:
        html = f.read()

    # Replace INV block
    inv_json = json.dumps(
        {'PICO':  {'L': pico['L'],  'D': pico['D'],  'LT': pico['LT'],  'GT': pico['GT']},
         'BOYLE': {'L': boyle['L'], 'D': boyle['D'], 'LT': boyle['LT'], 'GT': boyle['GT']}},
        separators=(',', ':')
    )
    sku_json = json.dumps(
        {'PICO': pico['SKU'], 'BOYLE': boyle['SKU']},
        separators=(',', ':')
    )

    html = re.sub(r'const INV = \{.*?\};', f'const INV = {inv_json};', html, flags=re.DOTALL)
    html = re.sub(r'const SKU=\{.*?\};',   f'const SKU={sku_json};',   html, flags=re.DOTALL)

    # Update "Last updated" timestamp in title
    html = re.sub(
        r'YES SALES INC\. — Warehouse Floor Map.*?(?=<)',
        f'YES SALES INC. — Warehouse Floor Map  |  Updated: {updated_at}',
        html
    )

    return html

# ── Step 4: Push to GitHub ───────────────────────────────────────────
def push_to_github(html_content, token):
    headers_h = {
        'Authorization': f'token {token}',
        'Accept': 'application/vnd.github.v3+json'
    }
    url = f'https://api.github.com/repos/{REPO}/contents/{DEPLOY_PATH}'

    r   = requests.get(url, headers=headers_h)
    sha = r.json().get('sha') if r.status_code == 200 else None

    content_b64 = base64.b64encode(html_content.encode('utf-8')).decode('utf-8')
    payload = {
        'message': f'WMS auto-update {datetime.now().strftime("%Y-%m-%d %H:%M")}',
        'content': content_b64,
        'branch':  BRANCH
    }
    if sha:
        payload['sha'] = sha

    r = requests.put(url, headers=headers_h, json=payload)
    if r.status_code in (200, 201):
        print(f"  ✅ https://yessales.github.io/dashboard/wms/")
        return True
    else:
        print(f"  ❌ GitHub {r.status_code}: {r.json().get('message')}")
        return False

# ── Main ─────────────────────────────────────────────────────────────
def main():
    print(f"\n{'='*55}")
    print(f"  YES SALES WMS Auto-Deploy — {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    print(f"{'='*55}")

    token = get_token()

    print("\n[1/4] Finding WMS files...")
    pico_path, boyle_path = find_wms_files(WMS_FOLDER)

    print("\n[2/4] Parsing PICO...")
    pico = parse_inventory(pico_path)
    print(f"  Located:{pico['LT']:>10,}  Default:{pico['D']:>10,}  Total:{pico['GT']:>10,}")
    print(f"  SKUs: {len(pico['SKU'])}  Locations: {len(pico['L'])}")

    print("\n[3/4] Parsing BOYLE...")
    boyle = parse_inventory(boyle_path)
    print(f"  Located:{boyle['LT']:>10,}  Default:{boyle['D']:>10,}  Total:{boyle['GT']:>10,}")
    print(f"  SKUs: {len(boyle['SKU'])}  Locations: {len(boyle['L'])}")

    print("\n[3/4] Generating HTML...")
    updated_at = datetime.now().strftime('%Y-%m-%d')
    html = generate_html(pico, boyle, updated_at)
    print(f"  Size: {len(html)//1024} KB")

    print("\n[4/4] Pushing to GitHub...")
    success = push_to_github(html, token)

    if success:
        print(f"\n✅ Complete! Live at:")
        print(f"   https://yessales.github.io/dashboard/wms/\n")
    else:
        print("\n❌ Deploy failed. Check token and repo settings.\n")
        sys.exit(1)

if __name__ == '__main__':
    main()
