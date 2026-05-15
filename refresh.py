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
  1. Reads the latest Excel file
  2. Filters Panama=N AND Column X=Yes (191 apps)
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

EXCEL_FILE               = "Post_added_new_Apps_BAFO_ITO_Apps_Inventory_Updated.xlsx"
SHEET_NAME               = "Consolidated App inventory"
HEADER_ROW               = 3
DASHBOARD_HTML           = "index.html"
PANAMA_COL               = "Panama (Y/N)"
PANAMA_FILTER            = "no"
STAFFING_ANALYSIS_COL_INDEX = 23
STAFFING_ANALYSIS_FILTER = "yes"
ONBOARDING_BUFFER_WEEKS  = 4
GITHUB_REPO_URL          = ""
GITHUB_BRANCH            = "main"
GITHUB_PAGES_URL         = ""
CONFIG_FILE              = ".dashboard_config"

WAVE_LABELS = {
    "2026-10-01": "Wave 1 (Oct 2026)",
    "2027-01-01": "Wave 2 (Jan 2027)",
    "2027-04-01": "Wave 3 (Apr 2027)",
    "2027-07-01": "Wave 4 (Jul 2027)",
    "2028-01-01": "Wave 5 (Jan 2028)",
}

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

    # Repo URL
    print("""
  ─────────────────────────────────────────────────────────
  CREATE YOUR GITHUB REPO (if not done yet)
  ─────────────────────────────────────────────────────────
  1. Go to:  https://github.com/new
  2. Name:   ito-staffing-dashboard
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
        # Update URL in case it changed
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
        with open(readme, "w") as f:
            f.write(f"# ITO Staffing Dashboard\n\n"
                    f"Live dashboard: {pages_url}\n\n"
                    f"## Refresh\n\n"
                    f"```bash\npython refresh.py --push\n```\n")
        ok("README.md created")

    # Derive user/repo for instructions
    user_repo = ""
    if m:
        user_repo = f"{m.group(1)}/{m.group(2)}"

    print(f"""
  ═══════════════════════════════════════════════════════════
  ✅  SETUP COMPLETE
  ═══════════════════════════════════════════════════════════

  NEXT STEPS:

  1. Make sure index.html exists in this folder
     (rename ITO_Staffing_Dashboard.html → index.html)

  2. Do your first push:
         python refresh.py --push

  3. Enable GitHub Pages (one-time, in the browser):
     → https://github.com/{user_repo}/settings/pages
     → Source: Deploy from branch → {branch} → / (root) → Save

  4. Share this URL with your team:
     → {pages_url}

  5. Every future update:
         python refresh.py --push
     That's it. Everyone sees the new data within ~60 seconds.
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

    step(1, 5, f"Reading Excel: {EXCEL_FILE}")
    if not os.path.exists(EXCEL_FILE):
        err(f"Excel not found: {EXCEL_FILE}\n"
            "  Update EXCEL_FILE at the top of refresh.py")
    try:
        df = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_NAME, header=HEADER_ROW)
        ok(f"{len(df)} rows loaded from '{SHEET_NAME}'")
    except Exception as e:
        err(f"Cannot read Excel: {e}")

    step(2, 5, "Filtering apps (Panama=N AND Column X=Yes)")
    col_x = df.columns[STAFFING_ANALYSIS_COL_INDEX]
    df_v = df[
        (df[PANAMA_COL].astype(str).str.strip().str.lower() == PANAMA_FILTER) &
        (df[col_x].astype(str).str.strip().str.lower() == STAFFING_ANALYSIS_FILTER)
    ].copy()
    if len(df_v) == 0:
        err("No rows matched filter. Check column indices and values.")
    ok(f"{len(df_v)} apps matched")

    step(3, 5, "Building dashboard payload")
    df_v["Transition Waves"] = pd.to_datetime(df_v["Transition Waves"], errors="coerce")
    for col in ["# Identified", "# To Staff", "ABP FTEs for Transition"]:
        df_v[col] = pd.to_numeric(df_v[col], errors="coerce").fillna(0)

    def wave_label(ts):
        if pd.isna(ts): return "Unassigned"
        return WAVE_LABELS.get(ts.strftime("%Y-%m-%d"), ts.strftime("%b %Y"))

    def ready_by(ts):
        if pd.isna(ts): return "Unassigned"
        return fmt_date((ts - timedelta(weeks=ONBOARDING_BUFFER_WEEKS)))

    def days_left(ts):
        if pd.isna(ts): return None
        return ((ts - timedelta(weeks=ONBOARDING_BUFFER_WEEKS)).date() - today).days

    rows = []
    for _, r in df_v.iterrows():
        ftes = round(float(r["# Identified"]) + float(r["# To Staff"]), 1)
        rows.append({
            "app":        str(r["Application Name "]).strip(),
            "tower":      str(r["Tower"]).strip()          if pd.notna(r["Tower"])          else "",
            "region":     str(r["Region"]).strip()         if pd.notna(r["Region"])         else "",
            "skill":      str(r["Skill category"]).strip() if pd.notna(r["Skill category"]) else "",
            "ftes":       ftes,
            "identified": round(float(r["# Identified"]), 1),
            "to_staff":   round(float(r["# To Staff"]),   1),
            "wave":       wave_label(r["Transition Waves"]),
            "ready_by":   ready_by(r["Transition Waves"]),
            "days_left":  days_left(r["Transition Waves"]),
            "kc_remarks": str(r["KC Remarks"]).strip() if pd.notna(r.get("KC Remarks")) else "",
        })

    n      = len(rows)
    ftes_t = round(sum(r["ftes"]       for r in rows))
    id_t   = round(sum(r["identified"] for r in rows))
    ts_t   = round(sum(r["to_staff"]   for r in rows))
    rdate  = fmt_date(pd.Timestamp(today))
    ok(f"{n} apps | FTEs Required: {ftes_t} | Identified: {id_t} | To Staff: {ts_t}")

    step(4, 5, f"Injecting into {DASHBOARD_HTML}")
    if not os.path.exists(DASHBOARD_HTML):
        err(f"{DASHBOARD_HTML} not found.\n"
            "  Rename ITO_Staffing_Dashboard.html → index.html")

    with open(DASHBOARD_HTML, "r", encoding="utf-8") as f:
        html = f.read()

    new_json = json.dumps(rows, ensure_ascii=False)
    new_html, subs = re.subn(
        r"(const RAW\s*=\s*)\[.*?\](\s*;)",
        r"\g<1>" + new_json + r"\g<2>",
        html, count=1, flags=re.DOTALL
    )
    if subs == 0:
        err("Could not find 'const RAW = [...]' in HTML. Template may be corrupted.")

    new_html = re.sub(
        r"\d+ apps\s*&nbsp;·&nbsp;\s*Refreshed[^<]*",
        f"{n} apps &nbsp;·&nbsp; Refreshed {rdate} &nbsp;·&nbsp; All waves assigned",
        new_html, count=1
    )

    with open(DASHBOARD_HTML, "w", encoding="utf-8") as f:
        f.write(new_html)
    ok(f"Saved → {os.path.abspath(DASHBOARD_HTML)}")

    return {"n": n, "ftes": ftes_t, "id": id_t, "ts": ts_t, "date": rdate}


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

    # Stage all relevant files
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
