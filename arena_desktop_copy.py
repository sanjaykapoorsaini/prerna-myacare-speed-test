import pandas as pd
import time
import os
import re
import json
from datetime import datetime
from playwright.sync_api import sync_playwright

# =========================================================
# CONFIG
# =========================================================

INPUT_FILE = "InputFile.xlsx"
MASTER_FILE = "PageSpeed_History.xlsx"
strategy   = "desktop"

# =========================================================
# FOLDER / FILE SETUP
# =========================================================

current_date = datetime.now().strftime("%Y-%m-%d")
folder_name  = current_date
os.makedirs(folder_name, exist_ok=True)

PDF_FOLDER = os.path.join(folder_name, f"pdf_reports_{strategy}")
os.makedirs(PDF_FOLDER, exist_ok=True)

timestamp   = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
OUTPUT_FILE = os.path.join(folder_name, f"{strategy}_results_{timestamp}.xlsx")

# =========================================================
# HELPERS
# =========================================================

def sanitize_filename(url):
    url = re.sub(r'https?://', '', url)
    url = re.sub(r'[^a-zA-Z0-9]', '_', url)
    return url[:150]


def _empty_metrics():
    return {
        "Performance"   : "N/A",
        "Accessibility" : "N/A",
        "Best Practices": "N/A",
        "SEO"           : "N/A",
        "FCP"           : "N/A",
        "LCP"           : "N/A",
        "TBT"           : "N/A",
        "CLS"           : "N/A",
        "Speed Index"   : "N/A",
    }


def dismiss_overlays(page):
    try:
        page.keyboard.press("Escape")
    except Exception:
        pass

    for selector in [
        "button:has-text('Ok, Got it.')",
        "button:has-text('OK')",
        "button:has-text('Accept All')",
        "button:has-text('Accept')",
    ]:
        try:
            loc = page.locator(selector)
            if loc.count() > 0:
                loc.first.click(timeout=2000)
                page.wait_for_timeout(500)
                break
        except Exception:
            pass

def update_historical_data(final_results, strategy):
    """
    Appends current run data into Historical_Data sheet.
    """

    run_date = datetime.now().strftime("%Y-%m-%d")

    rows = []

    for row in final_results:

        # LIVE ENTRY
        rows.append({
            "Run Date": run_date,
            "Strategy": strategy,
            "URL Type": "Live",
            "URL": row["Live URL"],
            "Performance": row["Live Performance"],
            "Accessibility": row["Live Accessibility"],
            "Best Practices": row["Live Best Practices"],
            "SEO": row["Live SEO"],
            "FCP": row["Live FCP"],
            "LCP": row["Live LCP"],
            "TBT": row["Live TBT"],
            "CLS": row["Live CLS"],
            "Speed Index": row["Live Speed Index"],
        })

        # REPLICA ENTRY
        rows.append({
            "Run Date": run_date,
            "Strategy": strategy,
            "URL Type": "Replica",
            "URL": row["Replica URL"],
            "Performance": row["Replica Performance"],
            "Accessibility": row["Replica Accessibility"],
            "Best Practices": row["Replica Best Practices"],
            "SEO": row["Replica SEO"],
            "FCP": row["Replica FCP"],
            "LCP": row["Replica LCP"],
            "TBT": row["Replica TBT"],
            "CLS": row["Replica CLS"],
            "Speed Index": row["Replica Speed Index"],
        })

    new_df = pd.DataFrame(rows)

    if os.path.exists(MASTER_FILE):
        with pd.ExcelWriter(MASTER_FILE, engine="openpyxl", mode="a", if_sheet_exists="overlay") as writer:
            try:
                old_df = pd.read_excel(MASTER_FILE, sheet_name="Historical_Data")
                combined = pd.concat([old_df, new_df], ignore_index=True)
            except:
                combined = new_df

            combined.to_excel(writer, sheet_name="Historical_Data", index=False)
    else:
        with pd.ExcelWriter(MASTER_FILE, engine="openpyxl") as writer:
            new_df.to_excel(writer, sheet_name="Historical_Data", index=False)

    print("✅ Historical_Data updated")

def save_snapshot_sheet(final_results, strategy):
    """
    Saves current run snapshot.
    """

    run_date = datetime.now().strftime("%Y-%m-%d")

    snapshot_df = pd.DataFrame(final_results)
    snapshot_df.insert(0, "Run Date", run_date)

    sheet_name = f"{strategy.capitalize()}_Results"

    if os.path.exists(MASTER_FILE):
        with pd.ExcelWriter(MASTER_FILE, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
            snapshot_df.to_excel(writer, sheet_name=sheet_name, index=False)
    else:
        with pd.ExcelWriter(MASTER_FILE, engine="openpyxl") as writer:
            snapshot_df.to_excel(writer, sheet_name=sheet_name, index=False)

    print(f"✅ {sheet_name} sheet updated")

def generate_url_trends():
    """
    Builds last 5 run history per URL.
    """

    if not os.path.exists(MASTER_FILE):
        return

    df = pd.read_excel(MASTER_FILE, sheet_name="Historical_Data")

    df["Run Date"] = pd.to_datetime(df["Run Date"])

    # Keep only last 5 runs per URL per strategy
    trend_df = (
        df.sort_values("Run Date", ascending=False)
          .groupby(["URL", "Strategy", "URL Type"])
          .head(5)
          .sort_values(["URL", "Strategy", "Run Date"])
    )

    with pd.ExcelWriter(MASTER_FILE, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        trend_df.to_excel(writer, sheet_name="URL_Trends", index=False)

    print("✅ URL_Trends sheet generated")

def generate_summary_dashboard():
    """
    Creates summary analytics sheet.
    """

    if not os.path.exists(MASTER_FILE):
        return

    df = pd.read_excel(MASTER_FILE, sheet_name="Historical_Data")

    df["Run Date"] = pd.to_datetime(df["Run Date"])
    df["Performance"] = pd.to_numeric(df["Performance"], errors="coerce")

    latest_date = df["Run Date"].max()

    summary_rows = []

    for (url, strategy, url_type), group in df.groupby(["URL", "Strategy", "URL Type"]):

        group_sorted = group.sort_values("Run Date", ascending=False)

        current = group_sorted.iloc[0]["Performance"]
        last5 = group_sorted.head(5)["Performance"]

        summary_rows.append({
            "URL": url,
            "Strategy": strategy,
            "URL Type": url_type,
            "Current Performance": current,
            "5-Day Avg": round(last5.mean(), 2),
            "Best (Last 5)": last5.max(),
            "Worst (Last 5)": last5.min(),
            "Change vs Yesterday":
                current - group_sorted.iloc[1]["Performance"]
                if len(group_sorted) > 1 else 0
        })

    summary_df = pd.DataFrame(summary_rows)

    with pd.ExcelWriter(MASTER_FILE, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        summary_df.to_excel(writer, sheet_name="Summary_Dashboard", index=False)

    print("✅ Summary_Dashboard sheet generated")

# =========================================================
# DOM EXTRACTION
# =========================================================

def extract_from_dom(page):
    try:
        scores_container = page.locator(".lh-scores-container").last
        raw_scores = [
            x.strip()
            for x in scores_container.locator(
                ".lh-gauge__percentage"
            ).all_text_contents()
        ]

        perf_section = page.locator("#performance").last
        raw_metrics = [
            x.strip().replace('\xa0', ' ')
            for x in perf_section.locator(
                ".lh-metric__value"
            ).all_text_contents()
        ]

        print(f"    [DOM] scores={raw_scores}")
        print(f"    [DOM] metrics={raw_metrics}")

        performance    = raw_scores[0] if len(raw_scores) > 0 else "N/A"
        accessibility  = raw_scores[1] if len(raw_scores) > 1 else "N/A"
        best_practices = raw_scores[2] if len(raw_scores) > 2 else "N/A"
        seo            = raw_scores[3] if len(raw_scores) > 3 else "N/A"

        if strategy == "mobile" and len(raw_metrics) >= 5:
            fcp, speed_index, lcp, tbt, cls = raw_metrics[:5]
        elif len(raw_metrics) >= 5:
            fcp, lcp, tbt, cls, speed_index = raw_metrics[:5]
        else:
            fcp = lcp = tbt = cls = speed_index = "N/A"

        return {
            "Performance"   : performance,
            "Accessibility" : accessibility,
            "Best Practices": best_practices,
            "SEO"           : seo,
            "FCP"           : fcp,
            "LCP"           : lcp,
            "TBT"           : tbt,
            "CLS"           : cls,
            "Speed Index"   : speed_index,
        }

    except Exception as e:
        print(f"    [DOM] Failed: {e}")
        return _empty_metrics()

# =========================================================
# STRONG READINESS CHECK
# Multiple signals to confirm Lighthouse run is TRULY complete
# =========================================================

def detect_loading_indicators(page):
    """
    PSI shows loading spinners / shimmer placeholders while
    Lighthouse runs. If ANY are visible, the report isn't final.
    """
    loading_selectors = [
        ".lh-loading",
        ".lh-shimmer",
        "[class*='loading']",
        "[class*='shimmer']",
        "[class*='spinner']",
        "[aria-busy='true']",
        ".cdc-spinner",
        ".lh-progress",
    ]

    for sel in loading_selectors:
        try:
            loc = page.locator(sel)
            count = loc.count()
            if count > 0:
                # Check if any are actually visible
                for i in range(min(count, 5)):
                    try:
                        if loc.nth(i).is_visible(timeout=500):
                            return True, sel
                    except Exception:
                        pass
        except Exception:
            pass

    return False, None


def get_url_hash_state(page):
    """
    PSI updates the URL hash/state when the report finalizes.
    Some PSI versions add ?form_factor=... or update hash fragments
    once the run completes.
    """
    try:
        return page.evaluate("() => window.location.href")
    except Exception:
        return ""


def wait_for_report_truly_ready(page, max_wait=240):
    """
    MULTI-SIGNAL readiness check:

      Signal 1: Elements exist (>= 4 scores, >= 5 metrics)
      Signal 2: Values are non-empty
      Signal 3: NO loading indicators visible
      Signal 4: DOM stable across 3 consecutive reads (was 2)
      Signal 5: Each read has a MINIMUM gap of 5 seconds
                (catches slow Lighthouse runs)
      Signal 6: Must observe AT LEAST ONE change before declaring
                stable (prevents catching early estimate as final)
    """
    print("  [Ready] Stage 1: Wait for DOM elements...")

    start = time.time()

    # ── Stage 1: Wait for elements with non-empty values ────
    while time.time() - start < max_wait:
        try:
            scores_loc = page.locator(
                ".lh-scores-container"
            ).last.locator(".lh-gauge__percentage")

            metrics_loc = page.locator(
                "#performance"
            ).last.locator(".lh-metric__value")

            sc = scores_loc.count()
            mc = metrics_loc.count()

            if sc >= 4 and mc >= 5:
                texts = [
                    t.strip()
                    for t in scores_loc.all_text_contents()
                ]
                non_empty = sum(1 for t in texts if t and t != "—")

                if non_empty >= 4:
                    print(f"  [Ready] Elements present "
                          f"(scores={sc}, metrics={mc})")
                    break

            print(f"  [Ready] Waiting... scores={sc}, metrics={mc}")
        except Exception as e:
            print(f"  [Ready] Probe error: {e}")

        page.wait_for_timeout(3000)
    else:
        print("  [Ready] Stage 1 timed out")
        return False

    # ── Stage 2: Wait for NO loading indicators ─────────────
    print("  [Ready] Stage 2: Wait for loading indicators to clear...")
    spinner_start = time.time()
    spinner_timeout = 60

    while time.time() - spinner_start < spinner_timeout:
        is_loading, which = detect_loading_indicators(page)
        if not is_loading:
            print("  [Ready] No loading indicators detected")
            break
        print(f"  [Ready] Still loading: {which}")
        page.wait_for_timeout(3000)
    else:
        print("  [Ready] Loading indicators didn't clear (continuing)")

    # ── Stage 3: Stability with REQUIRED CHANGE detection ───
    print("  [Ready] Stage 3: Wait for FINAL stable values...")

    snapshots         = []
    required_stable   = 3       # 3 consecutive identical reads
    required_changes  = 1       # Must observe ≥1 change first
    change_count      = 0
    stable_count      = 0
    poll_interval     = 5000    # 5 sec between reads
    stab_start        = time.time()
    stab_timeout      = 90      # Up to 90 sec for final values

    prev_snapshot = None

    while time.time() - stab_start < stab_timeout:
        try:
            scores = page.locator(
                ".lh-scores-container"
            ).last.locator(
                ".lh-gauge__percentage"
            ).all_text_contents()

            metrics = page.locator(
                "#performance"
            ).last.locator(
                ".lh-metric__value"
            ).all_text_contents()

            snap = str(scores) + "|" + str(metrics)
            snapshots.append((scores, metrics))

            if prev_snapshot is None:
                print(f"  [Ready] Initial: scores={scores}")
                prev_snapshot = snap
                page.wait_for_timeout(poll_interval)
                continue

            if snap != prev_snapshot:
                change_count += 1
                stable_count = 0
                print(f"  [Ready] CHANGED ({change_count}): "
                      f"scores={scores}")
            else:
                stable_count += 1
                print(f"  [Ready] Stable {stable_count}/"
                      f"{required_stable} (changes seen: "
                      f"{change_count})")

            # Final condition: stable AND has seen a change
            # OR has been stable for a long time without change
            # (some reports legitimately don't change after first
            # read because they were already complete)
            if stable_count >= required_stable:
                if change_count >= required_changes:
                    print("  [Ready] ✓ Stable after observed change")
                    return True
                else:
                    # No change seen — could be either:
                    #   (a) report was already done when we started
                    #   (b) we caught an early estimate that's stuck
                    # Sanity-check the values
                    if _values_look_reasonable(scores, metrics):
                        print("  [Ready] ✓ Stable, no change "
                              "but values look reasonable")
                        return True
                    else:
                        print("  [Ready] ⚠ Values look suspicious, "
                              "waiting more...")
                        # Reset and keep waiting
                        stable_count = 0

            prev_snapshot = snap

        except Exception as e:
            print(f"  [Ready] Stability error: {e}")

        page.wait_for_timeout(poll_interval)

    print("  [Ready] Stage 3 timed out — using last observed values")
    return False


def _values_look_reasonable(scores, metrics):
    """
    Heuristic sanity check on the values:
      - All scores should be 0-100 integers
      - At least one metric should NOT be 0/empty
      - Numeric values shouldn't be absurd

    Specifically guards against the "stuck early estimate" case
    where everything looks plausible structurally but the
    Lighthouse run hasn't actually finished.
    """
    try:
        # Check scores are valid integers 0-100
        for s in scores:
            s = s.strip()
            if not s or s == "—":
                return False
            try:
                val = int(s)
                if val < 0 or val > 100:
                    return False
            except ValueError:
                return False

        # Check metrics have at least some non-zero values
        non_zero_metrics = 0
        for m in metrics:
            m = m.strip().replace('\xa0', ' ')
            if m and m != "0" and m != "0 s" and m != "0 ms":
                non_zero_metrics += 1

        if non_zero_metrics < 3:
            return False

        return True
    except Exception:
        return False

# =========================================================
# DOUBLE-CHECK BY RELOADING REPORT URL
# =========================================================

def verify_by_reload(page, original_metrics, report_url):
    """
    Final safety check: reload the SAME archived report URL.
    Archived reports are immutable — reloading returns the
    SAME data. If our extracted values match what we get on
    reload, they're definitely the final values.

    If they differ, we caught a transitional state. Use the
    reload values (these match the immutable archive = PDF).
    """
    print("  [Verify] Reloading archived report to confirm...")

    try:
        page.goto(
            report_url,
            wait_until="domcontentloaded",
            timeout=120000
        )

        # Wait for elements
        try:
            page.wait_for_selector(
                ".lh-gauge__percentage",
                timeout=90000,
                state="visible"
            )
        except Exception:
            pass

        # Wait for loading indicators to clear
        for _ in range(15):
            is_loading, _ = detect_loading_indicators(page)
            if not is_loading:
                break
            page.wait_for_timeout(2000)

        # Extra settle
        page.wait_for_timeout(5000)

        reload_metrics = extract_from_dom(page)

        if reload_metrics == original_metrics:
            print("  [Verify] ✓ Reload matches original")
            return reload_metrics, "VERIFIED"
        else:
            print("  [Verify] ⚠ Reload DIFFERS from original")
            print(f"     Original: {original_metrics}")
            print(f"     Reload  : {reload_metrics}")
            return reload_metrics, "CORRECTED_BY_RELOAD"

    except Exception as e:
        print(f"  [Verify] Reload failed: {e}")
        return original_metrics, "RELOAD_FAILED"

# =========================================================
# MAIN
# =========================================================

def analyze_and_save_pdf(url, pdf_path, retries=3):
    """
    Workflow:
      1. Open PSI analysis URL
      2. Wait for archived report URL
      3. Multi-signal readiness wait (stage 1-3)
      4. Extract metrics (initial read)
      5. Reload report URL → confirm immutable values
      6. Save PDF from the verified state
    """
    for attempt in range(retries):
        try:
            with sync_playwright() as p:

                browser = p.chromium.launch(
                    headless=True,
                    channel="chrome",
                    args=[
                        "--disable-dev-shm-usage",
                        "--no-sandbox",
                        "--disable-gpu",
                    ]
                )

                page = browser.new_page(
                    viewport={"width": 1400, "height": 900}
                )

                psi_url = (
                    f"https://pagespeed.web.dev/analysis"
                    f"?url={url}"
                    f"&form_factor={strategy}"
                )

                print(f"\n  [Run] Opening: {psi_url}")

                page.goto(
                    psi_url,
                    wait_until="domcontentloaded",
                    timeout=240000
                )

                page.wait_for_url(
                    "**/analysis/**",
                    timeout=240000
                )
                report_url = page.url
                print(f"  [Run] Report URL: {report_url}")

                page.wait_for_timeout(3000)
                page_content = page.content()
                if any(kw in page_content for kw in [
                    "Oops! Something went wrong",
                    "RPC::UNREACHABLE",
                    "UNABLE_TO_RETRY",
                    "Lighthouse returned error",
                ]):
                    print(f"  [Run] PSI error. Retry {attempt+1}")
                    browser.close()
                    time.sleep(15)
                    continue

                # ── Multi-signal readiness wait ───────────────
                wait_for_report_truly_ready(page, max_wait=240)

                dismiss_overlays(page)
                page.wait_for_timeout(1500)

                # ── Initial extraction ────────────────────────
                metrics_initial = extract_from_dom(page)
                print(f"  [Initial] {metrics_initial}")

                # ── RELOAD VERIFICATION ───────────────────────
                # Archived reports are immutable — reload gives
                # us the FINAL values stored in PSI's database
                final_metrics, verify_status = verify_by_reload(
                    page, metrics_initial, report_url
                )

                # ── Dismiss overlays again post-reload ────────
                dismiss_overlays(page)
                page.wait_for_timeout(1500)

                # ── Save PDF (from reloaded final state) ──────
                try:
                    page.pdf(
                        path=pdf_path,
                        format="A3",
                        print_background=True
                    )
                    print(f"  [PDF] Saved: {pdf_path}")
                except Exception as e:
                    print(f"  [PDF] Save failed: {e}")

                # ── Final extraction post-PDF (sanity) ────────
                post_pdf = extract_from_dom(page)

                if post_pdf != final_metrics:
                    print("  ⚠ [Final] Values changed AFTER PDF save")
                    print(f"     Pre-PDF : {final_metrics}")
                    print(f"     Post-PDF: {post_pdf}")
                    # Use post-PDF values as they reflect what the
                    # PDF actually captured
                    final_metrics = post_pdf
                    verify_status += "+POST_DRIFT"

                browser.close()

                return {
                    "Report Link"     : report_url,
                    "Extraction Flag" : verify_status,
                    **final_metrics
                }

        except Exception as e:
            print(f"  [Run] Attempt {attempt+1} failed: {e}")
            time.sleep(15)

    return {
        "Report Link"     : "N/A",
        "Extraction Flag" : "FAILED",
        **_empty_metrics()
    }

# =========================================================
# READ INPUT
# =========================================================

df = pd.read_excel(INPUT_FILE)
final_results = []

# =========================================================
# PROCESS EACH ROW
# =========================================================

for index, row in df.iterrows():

    live_url    = str(row['Live']).strip()
    replica_url = str(row['Live Replica']).strip()

    print(f"\n{'='*60}")
    print(f"Row {index + 1}")
    print(f"  Live    : {live_url}")
    print(f"  Replica : {replica_url}")
    print(f"{'='*60}")

    live_pdf_name = sanitize_filename(live_url) + f"_LIVE_{strategy}.pdf"
    live_pdf_path = os.path.join(PDF_FOLDER, live_pdf_name)

    print("\n[LIVE]")
    live_data = analyze_and_save_pdf(live_url, live_pdf_path)

    time.sleep(3)

    replica_pdf_name = sanitize_filename(replica_url) + f"_REPLICA_{strategy}.pdf"
    replica_pdf_path = os.path.join(PDF_FOLDER, replica_pdf_name)

    print("\n[REPLICA]")
    replica_data = analyze_and_save_pdf(replica_url, replica_pdf_path)

    final_results.append({
        "Live URL"               : live_url,
        "Replica URL"            : replica_url,

        "Live Report Link"       : live_data["Report Link"],
        "Replica Report Link"    : replica_data["Report Link"],

        "Live Status"            : live_data.get("Extraction Flag", ""),
        "Replica Status"         : replica_data.get("Extraction Flag", ""),

        "Live Performance"       : live_data["Performance"],
        "Replica Performance"    : replica_data["Performance"],

        "Live Accessibility"     : live_data["Accessibility"],
        "Replica Accessibility"  : replica_data["Accessibility"],

        "Live Best Practices"    : live_data["Best Practices"],
        "Replica Best Practices" : replica_data["Best Practices"],

        "Live SEO"               : live_data["SEO"],
        "Replica SEO"            : replica_data["SEO"],

        "Live FCP"               : live_data["FCP"],
        "Replica FCP"            : replica_data["FCP"],

        "Live LCP"               : live_data["LCP"],
        "Replica LCP"            : replica_data["LCP"],

        "Live TBT"               : live_data["TBT"],
        "Replica TBT"            : replica_data["TBT"],

        "Live CLS"               : live_data["CLS"],
        "Replica CLS"            : replica_data["CLS"],

        "Live Speed Index"       : live_data["Speed Index"],
        "Replica Speed Index"    : replica_data["Speed Index"],

        "Live PDF Path"          : live_pdf_path,
        "Replica PDF Path"       : replica_pdf_path,
    })

    time.sleep(3)

# =========================================================
# EXPORT
# =========================================================

output_df = pd.DataFrame(final_results)
output_df.to_excel(OUTPUT_FILE, index=False)


# ----------------------------------------
# UPDATE MASTER WORKBOOK
# ----------------------------------------

save_snapshot_sheet(final_results, strategy)
update_historical_data(final_results, strategy)
generate_url_trends()
generate_summary_dashboard()

print("✅ Master workbook updated successfully!")

print(f"\n{'='*60}")
print(f"Done! Output: {OUTPUT_FILE}")
print(f"{'='*60}")