"""
refresh.py — ITO Staffing Dashboard · Full Auto-Refresh Pipeline
=================================================================
USAGE
-----
  First-time setup:
      python refresh.py --setup

  Refresh dashboard only (no push):
      python refresh.py

  Refresh + push to GitHub (everyone sees update):
      python refresh.py --push

  Refresh + push with a custom commit message:
      python refresh.py --push --message "Updated Wave 2 staffing"

WHAT IT DOES
------------
  1. Reads the latest Excel file (Consolidated App inventory sheet)
  2. Filters Panama = N (column O)
  3. Recalculates FTEs, wave labels, staff-ready-by dates, days left
  4. Injects fresh data into index.html (the dashboard)
  5. Updates the refresh timestamp in the header
  6. [--push] Commits and pushes to GitHub → everyone's URL updates instantly

REQUIREMENTS
------------
  pip install pandas openpyxl gitpython
"""

import json
import re
import sys
import os
import argparse
import subprocess
from datetime import timedelta, date

# ══════════════════════════════════════════════════════════════════════════════
#  CONFIGURATION — Edit these to match your setup
# ══════════════════════════════════════════════════════════════════════════════

# Excel source file — update this whenever the filename changes
EXCEL_FILE   = "BAFO ITO Apps Inventory & Transition timelines_June1st - KC_20260617.xlsx"

# Sheet name (must match exactly — case sensitive)
SHEET_NAME   = "Consolidated App inventory"

# Header row (0-indexed; row 4 in Excel = index 3)
HEADER_ROW   = 3

# Dashboard HTML file
DASHBOARD_HTML = "index.html"

# Panama filter — Column O (index 14), value = "No"
PANAMA_COL    = "Panama (Y/N)"
PANAMA_FILTER = "no"

# Onboarding buffer before wave date
ONBOARDING_BUFFER_WEEKS = 4

# ── Column indices (0-based) — confirmed from local file ─────────────────────
COL_TOWER        = 1    # B  — Tower
COL_APP_NAME     = 2    # C  — Application Name
COL_PANAMA       = 14   # O  — Panama (Y/N)
COL_FTE          = 21   # V  — ABP FTEs for Transition
COL_REGION       = 23   # X  — Region
COL_WAVE         = 25   # Z  — Transition Waves
COL_CATEGORY     = 27   # AB — Staffing approach
COL_SKILLS       = 28   # AC — Skills
COL_RESOURCES    = 29   # AD — Resources
COL_IDENTIFIED   = 30   # AE — # Identified
COL_TO_STAFF     = 31   # AF — # To Staff
COL_ALREADY_SUP  = 32   # AG — # Already supported

# Wave label map — add new waves here if needed
WAVE_LABELS = {
    "2026-10-01": "Wave 1 (Oct 2026)",
    "2027-01-01": "Wave 2 (Jan 2027)",
    "2027-04-01": "Wave 3 (Apr 2027)",
    "2027-07-01": "Wave 4 (Jul 2027)",
    "2028-01-01": "Wave 5 (Jan 2028)",
}

# GitHub settings — saved automatically by --setup
GITHUB_REPO_URL  = ""
GITHUB_BRANCH    = "main"
GITHUB_PAGES_URL = ""
CONFIG_FILE      = ".dashboard_config"

# ══════════════════════════════════════════════════════════════════════════════


def sep(char="─", width=62): print(char * width)
def header(t): print(); sep("═"); print(f"  {t}"); sep("═")
def step(n, tot, msg): print(f"\n  [{n}/{tot}] {msg}")
def ok(msg):   print(f"      ✅  {msg}")
def warn(msg): print(f"      ⚠   {msg}")
def err(msg):  print(f"      ❌  {msg}"); sys.exit(1)


# ── CONFIG ────────────────────────────────────────────────────────────────────

def load_config():
    global GITHUB_REPO_URL, GITHUB_BRANCH, GITHUB_PAGES_URL
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE) as f:
            c = json.load(f)
        GITHUB_REPO_URL  = c.get("GITHUB_REPO_URL",  GITHUB_REPO_URL)
        GITHUB_BRANCH    = c.get("GITHUB_BRANCH",    GITHUB_BRANCH)
        GITHUB_PAGES_URL = c.get("GITHUB_PAGES_URL", GITHUB_PAGES_URL)


def save_config():
    with open(CONFIG_FILE, "w") as f:
        json.dump({
            "GITHUB_REPO_URL":  GITHUB_REPO_URL,
            "GITHUB_BRANCH":    GITHUB_BRANCH,
            "GITHUB_PAGES_URL": GITHUB_PAGES_URL,
        }, f, indent=2)
    ok(f"Config saved to {CONFIG_FILE}")


# ── SETUP WIZARD ──────────────────────────────────────────────────────────────

def run_setup():
    global GITHUB_REPO_URL, GITHUB_BRANCH, GITHUB_PAGES_URL
    header("ITO Staffing Dashboard — First-Time Setup Wizard")
    print("  Configures git + GitHub. Run once, then use --push every time.\n")

    # Check git
    try:
        subprocess.run(["git", "--version"], capture_output=True, check=True)
        ok("git is installed")
    except Exception:
        err("git is not installed. Download from https://git-scm.com then re-run --setup")

    # Install python packages
    for pkg, import_name in [("pandas", "pandas"), ("openpyxl", "openpyxl"), ("gitpython", "git")]:
        try:
            __import__(import_name)
            ok(f"{pkg} is installed")
        except ImportError:
            print(f"      Installing {pkg}...")
            subprocess.run([sys.executable, "-m", "pip", "install", pkg], check=True)
            ok(f"{pkg} installed")

    print("""
  ─────────────────────────────────────────────────────────
  CREATE YOUR GITHUB REPO (if not done yet)
  ─────────────────────────────────────────────────────────
  1. Go to:  https://github.com/new
  2. Name:   ito-dashboard
  3. Set to: Public
  4. Click:  Create repository
  5. Copy the HTTPS URL  (e.g. https://github.com/user/repo.git)
  ─────────────────────────────────────────────────────────
""")
    repo_url = input("  Paste your GitHub repo HTTPS URL: ").strip()
    if not repo_url.endswith(".git"):
        repo_url = repo_url.rstrip("/") + ".git"
    GITHUB_REPO_URL = repo_url

    branch = input("  Branch name [Enter for 'main']: ").strip() or "main"
    GITHUB_BRANCH = branch

    # Derive Pages URL
    pages_url = ""
    m = re.match(r"https://github\.com/([^/]+)/([^/.]+)", repo_url)
    if m:
        pages_url = f"https://{m.group(1)}.github.io/{m.group(2)}/"
    print(f"\n  Detected GitHub Pages URL: {pages_url}")
    if input("  Correct? [Y/n]: ").strip().lower() == "n":
        pages_url = input("  Enter GitHub Pages URL: ").strip()
    GITHUB_PAGES_URL = pages_url

    save_config()

    # Init git repo
    cwd = os.getcwd()
    if not os.path.exists(os.path.join(cwd, ".git")):
        subprocess.run(["git", "init"], cwd=cwd, check=True)
        subprocess.run(["git", "checkout", "-b", branch], cwd=cwd)
        ok("git repository initialised")
    else:
        ok("git repository already exists")

    # Set remote
    result = subprocess.run(["git", "remote"], capture_output=True, text=True, cwd=cwd)
    if "origin" not in result.stdout:
        subprocess.run(["git", "remote", "add", "origin", repo_url], cwd=cwd, check=True)
        ok(f"Remote set → {repo_url}")
    else:
        subprocess.run(["git", "remote", "set-url", "origin", repo_url], cwd=cwd)
        ok(f"Remote updated → {repo_url}")

    # .gitignore
    gi = os.path.join(cwd, ".gitignore")
    if not os.path.exists(gi):
        with open(gi, "w") as f:
            f.write("*.pyc\n__pycache__/\n.DS_Store\n*.tmp\n.env\n")
        ok(".gitignore created")

    # README
    readme = os.path.join(cwd, "README.md")
    if not os.path.exists(readme):
        user_repo = f"{m.group(1)}/{m.group(2)}" if m else "user/repo"
        with open(readme, "w") as f:
            f.write(f"# ITO Staffing Dashboard\n\nLive: {pages_url}\n\n"
                    f"## Refresh\n\n```bash\npython refresh.py --push\n```\n")
        ok("README.md created")

    user_repo = f"{m.group(1)}/{m.group(2)}" if m else "user/repo"
    print(f"""
  ═══════════════════════════════════════════════════════════
  ✅  SETUP COMPLETE
  ═══════════════════════════════════════════════════════════

  NEXT STEPS:

  1. Do your first push:
         python refresh.py --push

  2. Enable GitHub Pages (one-time, in the browser):
     → https://github.com/{user_repo}/settings/pages
     → Source: Deploy from branch → {branch} → / (root) → Save

  3. Share this URL with your team:
     → {pages_url}

  4. Every future update:
         python refresh.py --push
  ═══════════════════════════════════════════════════════════
""")


# ── DATA REFRESH ──────────────────────────────────────────────────────────────

def refresh_dashboard():
    try:
        import pandas as pd
    except ImportError:
        err("pandas not installed. Run: pip install pandas openpyxl")

    today = date.today()

    def fmt_date(d):
        try:
            return d.strftime("%-d %b %Y") if sys.platform != "win32" else d.strftime("%#d %b %Y")
        except Exception:
            return str(d)

    # ── 1. Load Excel ──────────────────────────────────────────────────────────
    step(1, 5, f"Reading Excel: {EXCEL_FILE}")
    if not os.path.exists(EXCEL_FILE):
        err(f"Excel not found: {EXCEL_FILE}\n"
            "  Update EXCEL_FILE at the top of refresh.py")
    try:
        df = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_NAME, header=HEADER_ROW)
        ok(f"{len(df)} rows loaded from '{SHEET_NAME}'")
    except Exception as e:
        err(f"Cannot read Excel: {e}")

    # ── 2. Filter Panama = N ───────────────────────────────────────────────────
    step(2, 5, "Filtering apps (Panama = N)")

    # Parse wave column BEFORE filtering to avoid dtype conflict on copied slice
    wave_col_name = df.columns[COL_WAVE]
    df[wave_col_name] = pd.to_datetime(df[wave_col_name], errors="coerce")

    df_valid = df[
        df[PANAMA_COL].astype(str).str.strip().str.lower() == PANAMA_FILTER
    ].copy()

    if len(df_valid) == 0:
        err(f"No rows matched Panama=N filter.\n"
            f"  Check that PANAMA_COL='{PANAMA_COL}' exists and contains 'No' values.")
    ok(f"{len(df_valid)} apps matched (Panama = N)")

    # ── 3. Process data ────────────────────────────────────────────────────────
    step(3, 5, "Building dashboard payload")

    def safe_float(val):
        """Safely convert any value to a rounded integer, default 0."""
        try:
            f = float(val)
            return 0 if pd.isna(f) else round(f)
        except (ValueError, TypeError):
            return 0

    def wave_label(ts):
        if pd.isna(ts): return "Unassigned"
        return WAVE_LABELS.get(ts.strftime("%Y-%m-%d"), ts.strftime("%b %Y"))

    def ready_by(ts):
        if pd.isna(ts): return "Unassigned"
        return fmt_date((ts - timedelta(weeks=ONBOARDING_BUFFER_WEEKS)))

    def days_left(ts):
        if pd.isna(ts): return None
        return ((ts - timedelta(weeks=ONBOARDING_BUFFER_WEEKS)).date() - today).days

    def clean(val):
        s = str(val).strip()
        return "" if s in ("nan", "None", "") else s

    rows = []
    for _, r in df_valid.iterrows():
        rows.append({
            "app":               clean(r.iloc[COL_APP_NAME]),
            "tower":             clean(r.iloc[COL_TOWER]),
            "region":            clean(r.iloc[COL_REGION]),
            "skill":             clean(r.iloc[COL_SKILLS]),
            "category":          clean(r.iloc[COL_CATEGORY]),
            "resources":         clean(r.iloc[COL_RESOURCES]),
            "ftes":              safe_float(r.iloc[COL_FTE]),
            "identified":        safe_float(r.iloc[COL_IDENTIFIED]),
            "to_staff":          safe_float(r.iloc[COL_TO_STAFF]),
            "reclaim":           0,
            "already_supported": safe_float(r.iloc[COL_ALREADY_SUP]),
            "wave":              wave_label(r.iloc[COL_WAVE]),
            "ready_by":          ready_by(r.iloc[COL_WAVE]),
            "days_left":         days_left(r.iloc[COL_WAVE]),
        })

    n       = len(rows)
    ftes_t  = sum(r["ftes"]              for r in rows)
    id_t    = sum(r["identified"]        for r in rows)
    ts_t    = sum(r["to_staff"]          for r in rows)
    as_t    = sum(r["already_supported"] for r in rows)
    rdate   = fmt_date(pd.Timestamp(today))

    ok(f"{n} apps | FTEs: {ftes_t} | Identified: {id_t} | "
       f"To Staff: {ts_t} | Already Supp: {as_t}")

    # ── 4. Inject into HTML ────────────────────────────────────────────────────
    step(4, 5, f"Injecting data into {DASHBOARD_HTML}")
    if not os.path.exists(DASHBOARD_HTML):
        err(f"{DASHBOARD_HTML} not found in current folder.\n"
            "  Make sure index.html is in the same folder as refresh.py")

    with open(DASHBOARD_HTML, "r", encoding="utf-8") as f:
        html = f.read()

    new_json = json.dumps(rows, ensure_ascii=False)
    new_html, subs = re.subn(
        r"(const RAW\s*=\s*)\[.*?\](\s*;)",
        r"\g<1>" + new_json + r"\g<2>",
        html, count=1, flags=re.DOTALL
    )
    if subs == 0:
        err("Could not find 'const RAW = [...]' in the HTML.\n"
            "  The dashboard template may be corrupted.")

    new_html = re.sub(
        r"\d+ apps\s*&nbsp;·&nbsp;\s*Refreshed[^<]*",
        f"{n} apps &nbsp;·&nbsp; Refreshed {rdate} &nbsp;·&nbsp; All waves assigned",
        new_html, count=1
    )

    with open(DASHBOARD_HTML, "w", encoding="utf-8") as f:
        f.write(new_html)
    ok(f"Saved → {os.path.abspath(DASHBOARD_HTML)}")

    return {"n": n, "ftes": ftes_t, "id": id_t, "ts": ts_t,
            "as_": as_t, "date": rdate}


# ── GITHUB PUSH ───────────────────────────────────────────────────────────────

def push_to_github(msg=None, stats=None):
    try:
        import git as gitlib
    except ImportError:
        err("gitpython not installed. Run: pip install gitpython\n"
            "  Or run: python refresh.py --setup")

    step(5, 5, "Pushing to GitHub")

    if not GITHUB_REPO_URL:
        err("GitHub not configured. Run: python refresh.py --setup")

    cwd = os.getcwd()
    if not os.path.exists(os.path.join(cwd, ".git")):
        err("Not a git repository. Run: python refresh.py --setup")

    try:
        repo = gitlib.Repo(cwd)
    except Exception as e:
        err(f"Cannot open git repo: {e}")

    # Ensure remote exists
    try:
        origin = repo.remote("origin")
    except Exception:
        origin = repo.create_remote("origin", GITHUB_REPO_URL)
        ok(f"Remote 'origin' set → {GITHUB_REPO_URL}")

    # Stage files
    to_stage = [DASHBOARD_HTML, "refresh.py", ".gitignore", CONFIG_FILE, "README.md"]
    for f in to_stage:
        if os.path.exists(f):
            repo.index.add([f])

    # Commit message
    if not msg and stats:
        msg = (f"Dashboard refresh — {stats['date']} | "
               f"{stats['n']} apps | FTEs {stats['ftes']} | To Staff {stats['ts']}")
    elif not msg:
        from datetime import datetime
        msg = f"Dashboard refresh — {datetime.now().strftime('%Y-%m-%d %H:%M')}"

    try:
        repo.index.commit(msg)
        ok(f"Committed: \"{msg}\"")
    except gitlib.exc.GitCommandError as e:
        if "nothing to commit" in str(e).lower():
            warn("Nothing new to commit")
        else:
            err(f"Commit error: {e}")

    # Push
    try:
        print(f"      Pushing to {GITHUB_REPO_URL} [{GITHUB_BRANCH}]...")
        origin.push(refspec=f"HEAD:{GITHUB_BRANCH}", set_upstream=True)
        ok("Pushed successfully!")
    except Exception as e:
        print()
        warn(f"Push failed: {e}")
        print("""
      ── HOW TO FIX AUTHENTICATION ──────────────────────────────

      Option A — Personal Access Token (easiest):
        1. https://github.com/settings/tokens/new
           → Check "repo" scope → Generate token → Copy it
        2. Run once:
             git remote set-url origin https://TOKEN@github.com/user/repo.git
        3. Re-run: python refresh.py --push

      Option B — SSH key:
        1. ssh-keygen -t ed25519 -C "your@email.com"
        2. Add public key at: https://github.com/settings/keys
        3. git remote set-url origin git@github.com:user/repo.git
        4. Re-run: python refresh.py --push
      ────────────────────────────────────────────────────────────
""")
        sys.exit(1)


# ── MAIN ──────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        prog="refresh.py",
        description="ITO Staffing Dashboard — Auto-Refresh Pipeline",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python refresh.py                              Refresh dashboard only
  python refresh.py --push                       Refresh + push to GitHub
  python refresh.py --push --message "Wave 1"   Custom commit message
  python refresh.py --setup                      First-time GitHub setup
        """
    )
    parser.add_argument("--push",    action="store_true", help="Push to GitHub after refresh")
    parser.add_argument("--setup",   action="store_true", help="Run first-time setup wizard")
    parser.add_argument("--message", "-m", default=None,  help="Custom git commit message")
    args = parser.parse_args()

    load_config()

    if args.setup:
        run_setup()
        return

    header("ITO Staffing Dashboard — Auto-Refresh Pipeline")

    if args.push and not GITHUB_REPO_URL:
        warn("GitHub not configured yet. Starting setup wizard...\n")
        run_setup()
        load_config()

    stats = refresh_dashboard()

    if args.push:
        push_to_github(msg=args.message, stats=stats)

    # Final summary
    print()
    sep("═")
    print("  ✅  ALL DONE")
    sep("═")
    print(f"""
  Dashboard : {os.path.abspath(DASHBOARD_HTML)}
  Apps      : {stats['n']}
  FTEs Req  : {stats['ftes']}
  Identified: {stats['id']}
  To Staff  : {stats['ts']}
  Alr. Supp : {stats['as_']}
  Refreshed : {stats['date']}""")

    if args.push and GITHUB_PAGES_URL:
        print(f"""
  🌐  Live at : {GITHUB_PAGES_URL}
  Note        : GitHub Pages updates in ~60 seconds after push
""")
    elif not args.push:
        print("""
  To share with your team, push to GitHub:
      python refresh.py --push
""")
    sep("═")


if __name__ == "__main__":
    main()
