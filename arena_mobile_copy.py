import pandas as pd
import time
import os
import re
from datetime import datetime
from urllib.parse import quote
from playwright.sync_api import sync_playwright
import pdfplumber
from openpyxl import load_workbook

# =========================================================
# CONFIG
# =========================================================

INPUT_FILE = "InputFile.xlsx"
MASTER_FILE = "PageSpeed_History.xlsx"
# "mobile" or "desktop"
strategy = "mobile"

# =========================================================
# FOLDER SETUP
# =========================================================

current_date = datetime.now().strftime("%Y-%m-%d")
folder_name = current_date
os.makedirs(folder_name, exist_ok=True)

PDF_FOLDER = os.path.join(folder_name, f"pdf_reports_{strategy}")
os.makedirs(PDF_FOLDER, exist_ok=True)

timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
OUTPUT_FILE = os.path.join(folder_name, f"{strategy}_results_{timestamp}.xlsx")

# =========================================================
# HELPERS
# =========================================================

def sanitize_filename(url):
    url = re.sub(r'https?://', '', url)
    url = re.sub(r'[^a-zA-Z0-9]', '_', url)
    return url[:150]


def empty_metrics():
    return {
        "Performance": "N/A",
        "Accessibility": "N/A",
        "Best Practices": "N/A",
        "SEO": "N/A",
        "FCP": "N/A",
        "LCP": "N/A",
        "TBT": "N/A",
        "CLS": "N/A",
        "Speed Index": "N/A",
    }


def all_values_found(data):
    required = [
        "Performance",
        "Accessibility",
        "Best Practices",
        "SEO",
        "FCP",
        "LCP",
        "TBT",
        "CLS",
        "Speed Index",
    ]

    return all(
        data.get(k) not in [None, "", "N/A"]
        for k in required
    )

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

def dismiss_overlays(page):
    try:
        page.keyboard.press("Escape")
    except Exception:
        pass

    selectors = [
        "button:has-text('Ok, Got it.')",
        "button:has-text('OK')",
        "button:has-text('Ok')",
        "button:has-text('Accept All')",
        "button:has-text('Accept')",
    ]

    for selector in selectors:
        try:
            loc = page.locator(selector)
            if loc.count() > 0:
                loc.first.click(timeout=2000)
                page.wait_for_timeout(500)
                break
        except Exception:
            pass

# =========================================================
# PDF PARSING HELPERS
# =========================================================

def norm_word(text):
    return re.sub(r'[^a-z0-9]+', '', str(text).lower())


def word_center_x(word):
    return (word["x0"] + word["x1"]) / 2


def word_center_y(word):
    return (word["top"] + word["bottom"]) / 2


def find_phrase_boxes(words, phrase, min_top=0, max_top=10**9):
    """
    Finds phrase locations in pdfplumber words.
    Example:
      phrase = "First Contentful Paint"
    """
    tokens = [norm_word(x) for x in phrase.split()]
    n = len(tokens)

    boxes = []

    sorted_words = sorted(words, key=lambda w: (w["top"], w["x0"]))

    for i in range(len(sorted_words) - n + 1):
        seq = sorted_words[i:i+n]

        if not (min_top <= seq[0]["top"] <= max_top):
            continue

        seq_tokens = [norm_word(w["text"]) for w in seq]

        if seq_tokens == tokens:
            boxes.append({
                "x0": min(w["x0"] for w in seq),
                "x1": max(w["x1"] for w in seq),
                "top": min(w["top"] for w in seq),
                "bottom": max(w["bottom"] for w in seq),
                "text": " ".join(w["text"] for w in seq),
            })

    return boxes


def find_number_above_label(words, label_box, min_top):
    """
    For score gauges:
      number is visually above label.
      Example:
        70
        Performance
    """
    label_cx = (label_box["x0"] + label_box["x1"]) / 2

    candidates = []

    for w in words:
        text = w["text"].strip()

        if not re.fullmatch(r'\d{1,3}', text):
            continue

        if w["top"] < min_top:
            continue

        if w["bottom"] >= label_box["top"]:
            continue

        vertical_gap = label_box["top"] - w["bottom"]

        if vertical_gap > 140:
            continue

        x_gap = abs(word_center_x(w) - label_cx)

        if x_gap > 120:
            continue

        candidates.append((vertical_gap, x_gap, text))

    if not candidates:
        return "N/A"

    candidates.sort(key=lambda x: (x[0], x[1]))
    return candidates[0][2]


def is_number_token(text):
    return re.fullmatch(r'\d+(?:[,.]\d+)?', text.strip()) is not None


def find_value_below_label(words, label_box):
    """
    For lab metrics:
      label is above value.
      Example:
        First Contentful Paint
        2.9 s
    """
    label_cx = (label_box["x0"] + label_box["x1"]) / 2

    candidates = []

    for w in words:
        text = w["text"].strip()

        if not is_number_token(text):
            continue

        if w["top"] <= label_box["bottom"]:
            continue

        vertical_gap = w["top"] - label_box["bottom"]

        if vertical_gap > 90:
            continue

        x_gap = abs(word_center_x(w) - label_cx)

        if x_gap > 180:
            continue

        candidates.append((vertical_gap, x_gap, w))

    if not candidates:
        return "N/A"

    candidates.sort(key=lambda x: (x[0], x[1]))
    number_word = candidates[0][2]

    value = number_word["text"].strip()

    # Check if unit exists immediately after number on same line.
    # Example: "2.9" + "s", "190" + "ms"
    unit = ""

    for w in words:
        if abs(w["top"] - number_word["top"]) > 4:
            continue

        if w["x0"] <= number_word["x1"]:
            continue

        if w["x0"] - number_word["x1"] > 30:
            continue

        t = w["text"].strip().lower()

        if t in ["s", "ms"]:
            unit = t
            break

    if unit:
        value = f"{value} {unit}"

    return value


# =========================================================
# PDF PARSER — SOURCE OF TRUTH
# =========================================================

def parse_pagespeed_pdf(pdf_path):
    """
    Extracts values from the generated PDF itself.

    This guarantees Excel values match the PDF because the PDF
    is the source of truth.
    """
    result = empty_metrics()

    try:
        with pdfplumber.open(pdf_path) as pdf:

            all_text_pages = []

            # -------------------------------------------------
            # 1. Find the page containing:
            #    "Diagnose performance issues"
            # -------------------------------------------------
            target_page = None
            target_words = None
            diagnose_box = None

            for page_index, page in enumerate(pdf.pages):
                text = page.extract_text(layout=True) or ""
                all_text_pages.append(text)

                words = page.extract_words(
                    x_tolerance=2,
                    y_tolerance=3,
                    keep_blank_chars=False
                )

                boxes = find_phrase_boxes(
                    words,
                    "Diagnose performance issues"
                )

                if boxes:
                    target_page = page
                    target_words = words
                    diagnose_box = boxes[0]
                    print(f"    [PDF] Found Diagnose section on page {page_index + 1}")
                    break

            if target_page is None:
                print("    [PDF] Diagnose performance issues section not found")
                save_pdf_debug_text(pdf_path, "\n\n".join(all_text_pages))
                return result

            words = target_words
            diag_bottom = diagnose_box["bottom"]

            # -------------------------------------------------
            # 2. Find METRICS header
            # -------------------------------------------------
            metric_header_boxes = find_phrase_boxes(
                words,
                "METRICS",
                min_top=diag_bottom
            )

            if metric_header_boxes:
                metrics_top = metric_header_boxes[0]["top"]
            else:
                metrics_top = target_page.height

            # -------------------------------------------------
            # 3. Extract scores from the score row
            # -------------------------------------------------
            score_labels = {
                "Performance": "Performance",
                "Accessibility": "Accessibility",
                "Best Practices": "Best Practices",
                "SEO": "SEO",
            }

            print("    [PDF] Extracting scores from PDF...")

            for output_name, label_text in score_labels.items():

                boxes = find_phrase_boxes(
                    words,
                    label_text,
                    min_top=diag_bottom,
                    max_top=metrics_top
                )

                if not boxes:
                    print(f"      {output_name}: label not found")
                    continue

                # Use the first occurrence after Diagnose header.
                # This corresponds to the score row shown in PDF.
                label_box = boxes[0]

                score = find_number_above_label(
                    words,
                    label_box,
                    min_top=diag_bottom
                )

                result[output_name] = score
                print(f"      {output_name}: {score}")

            # -------------------------------------------------
            # 4. Extract lab metrics from METRICS section
            # -------------------------------------------------
            metric_labels = {
                "FCP": "First Contentful Paint",
                "LCP": "Largest Contentful Paint",
                "TBT": "Total Blocking Time",
                "CLS": "Cumulative Layout Shift",
                "Speed Index": "Speed Index",
            }

            print("    [PDF] Extracting lab metrics from PDF...")

            for output_name, label_text in metric_labels.items():

                boxes = find_phrase_boxes(
                    words,
                    label_text,
                    min_top=metrics_top,
                    max_top=target_page.height
                )

                if not boxes:
                    print(f"      {output_name}: label not found")
                    continue

                label_box = boxes[0]

                value = find_value_below_label(
                    words,
                    label_box
                )

                result[output_name] = value
                print(f"      {output_name}: {value}")

            # -------------------------------------------------
            # 5. Fallback text parser if coordinate parser missed anything
            # -------------------------------------------------
            if not all_values_found(result):
                print("    [PDF] Coordinate parser missed some values, trying text fallback...")

                full_text = "\n\n".join(all_text_pages)

                fallback = parse_pdf_text_fallback(full_text)

                for key, value in fallback.items():
                    if result.get(key) in ["N/A", "", None] and value not in ["N/A", "", None]:
                        result[key] = value

            if not all_values_found(result):
                print("    [PDF] WARNING: Some values still missing")
                save_pdf_debug_text(pdf_path, "\n\n".join(all_text_pages))

            print(f"    [PDF] Final parsed values: {result}")
            return result

    except Exception as e:
        print(f"    [PDF] Parse failed: {e}")
        return result


def parse_pdf_text_fallback(full_text):
    """
    Text fallback parser.
    Used only if coordinate-based parsing misses something.
    """
    result = empty_metrics()

    text = full_text.replace("\xa0", " ")
    text = re.sub(r'[ \t]+', ' ', text)

    # Isolate lab data area
    m = re.search(
        r'Diagnose performance issues(.*?)(?:INSIGHTS|DIAGNOSTICS|PASSED AUDITS|PageSpeed Insights)',
        text,
        re.I | re.S
    )

    if m:
        lab = m.group(1)
    else:
        lab = text

    flat = re.sub(r'\s+', ' ', lab)

    # Scores pattern: 70 Performance 77 Accessibility 100 Best Practices 85 SEO
    score_patterns = [
        r'(\d{1,3})\s+Performance\s+(\d{1,3})\s+Accessibility\s+(\d{1,3})\s+Best\s+Practices\s+(\d{1,3})\s+SEO',
        r'(\d{1,3})\s+(\d{1,3})\s+(\d{1,3})\s+(\d{1,3})\s+Performance\s+Accessibility\s+Best\s+Practices\s+SEO',
    ]

    for pattern in score_patterns:
        sm = re.search(pattern, flat, re.I)
        if sm:
            result["Performance"] = sm.group(1)
            result["Accessibility"] = sm.group(2)
            result["Best Practices"] = sm.group(3)
            result["SEO"] = sm.group(4)
            break

    # Metric label followed by value fallback
    metric_patterns = {
        "FCP": r'First Contentful Paint\s+([0-9,.]+)\s*(s|ms)?',
        "LCP": r'Largest Contentful Paint\s+([0-9,.]+)\s*(s|ms)?',
        "TBT": r'Total Blocking Time\s+([0-9,.]+)\s*(s|ms)?',
        "CLS": r'Cumulative Layout Shift\s+([0-9,.]+)',
        "Speed Index": r'Speed Index\s+([0-9,.]+)\s*(s|ms)?',
    }

    for key, pattern in metric_patterns.items():
        mm = re.search(pattern, flat, re.I)
        if mm:
            value = mm.group(1)
            unit = mm.group(2) if len(mm.groups()) >= 2 else None
            if unit:
                value = f"{value} {unit}"
            result[key] = value

    return result


def save_pdf_debug_text(pdf_path, text):
    """
    Saves extracted PDF text for debugging if parser misses values.
    """
    try:
        debug_path = pdf_path + ".debug.txt"
        with open(debug_path, "w", encoding="utf-8") as f:
            f.write(text)
        print(f"    [PDF] Debug text saved: {debug_path}")
    except Exception as e:
        print(f"    [PDF] Could not save debug text: {e}")

# =========================================================
# WAIT FOR PSI REPORT BEFORE PDF
# =========================================================

def wait_for_report_visible(page, timeout=300):
    """
    Wait until enough visible text exists to indicate the report
    is rendered before creating PDF.
    """
    print("  [Wait] Waiting for PSI report to render...")

    start = time.time()

    required_texts = [
        "Diagnose performance issues",
        "First Contentful Paint",
        "Largest Contentful Paint",
        "Total Blocking Time",
        "Cumulative Layout Shift",
        "Speed Index",
    ]

    while time.time() - start < timeout:
        try:
            content = page.locator("body").inner_text(timeout=5000)

            if all(t in content for t in required_texts):
                print("  [Wait] Required report text found")
                return True

        except Exception:
            pass

        page.wait_for_timeout(3000)

    print("  [Wait] Timeout waiting for report text")
    return False

# =========================================================
# ANALYZE + SAVE PDF + PARSE PDF
# =========================================================

def analyze_and_save_pdf(url, pdf_path, retries=3):
    """
    Important:
      1. Open PSI UI
      2. Wait for report URL
      3. Save PDF
      4. Parse values from saved PDF
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

                encoded_url = quote(url, safe="")

                psi_url = (
                    f"https://pagespeed.web.dev/analysis"
                    f"?url={encoded_url}"
                    f"&form_factor={strategy}"
                )

                print(f"\n  [Run] Opening:")
                print(f"  {psi_url}")

                page.goto(
                    psi_url,
                    wait_until="domcontentloaded",
                    timeout=300000
                )

                page.wait_for_url(
                    "**/analysis/**",
                    timeout=300000
                )

                report_url = page.url

                print(f"  [Run] Report URL:")
                print(f"  {report_url}")

                # Detect errors
                page.wait_for_timeout(3000)
                page_content = page.content()

                error_keywords = [
                    "Oops! Something went wrong",
                    "RPC::UNREACHABLE",
                    "UNABLE_TO_RETRY",
                    "Lighthouse returned error",
                ]

                if any(k in page_content for k in error_keywords):
                    print(f"  [Run] PSI error detected. Retry {attempt + 1}")
                    browser.close()
                    time.sleep(20)
                    continue

                # Wait for report text
                wait_for_report_visible(
                    page,
                    timeout=360 if strategy == "mobile" else 240
                )

                # Extra settle time
                page.wait_for_timeout(10000 if strategy == "mobile" else 5000)

                dismiss_overlays(page)

                # -------------------------------------------------
                # Save PDF and parse it.
                # If parsing fails due to early rendering,
                # wait and regenerate PDF.
                # -------------------------------------------------
                final_metrics = empty_metrics()
                status = "PDF_PARSE_FAILED"

                for pdf_try in range(3):

                    print(f"  [PDF] Saving PDF attempt {pdf_try + 1}/3")

                    page.pdf(
                        path=pdf_path,
                        format="A3",
                        print_background=True
                    )

                    print(f"  [PDF] Saved:")
                    print(f"  {pdf_path}")

                    parsed = parse_pagespeed_pdf(pdf_path)

                    if all_values_found(parsed):
                        final_metrics = parsed
                        status = "PDF_PARSED"
                        break

                    print("  [PDF] Some values missing. Waiting and retrying PDF...")
                    page.wait_for_timeout(15000)

                browser.close()

                return {
                    "Report Link": report_url,
                    "Status": status,
                    **final_metrics
                }

        except Exception as e:
            print(f"  [Run] Attempt {attempt + 1} failed:")
            print(e)
            time.sleep(20)

    return {
        "Report Link": "N/A",
        "Status": "FAILED",
        **empty_metrics()
    }

# =========================================================
# READ INPUT
# =========================================================

df = pd.read_excel(INPUT_FILE)
final_results = []

# =========================================================
# PROCESS ROWS
# =========================================================

for index, row in df.iterrows():

    live_url = str(row["Live"]).strip()
    replica_url = str(row["Live Replica"]).strip()

    print(f"\n{'=' * 60}")
    print(f"Row {index + 1} ({strategy})")
    print(f"  Live    : {live_url}")
    print(f"  Replica : {replica_url}")
    print(f"{'=' * 60}")

    # -----------------------------------------------------
    # LIVE
    # -----------------------------------------------------
    live_pdf_name = sanitize_filename(live_url) + f"_LIVE_{strategy}.pdf"
    live_pdf_path = os.path.join(PDF_FOLDER, live_pdf_name)

    print("\n[LIVE]")
    live_data = analyze_and_save_pdf(
        live_url,
        live_pdf_path
    )

    time.sleep(3)

    # -----------------------------------------------------
    # REPLICA
    # -----------------------------------------------------
    replica_pdf_name = sanitize_filename(replica_url) + f"_REPLICA_{strategy}.pdf"
    replica_pdf_path = os.path.join(PDF_FOLDER, replica_pdf_name)

    print("\n[REPLICA]")
    replica_data = analyze_and_save_pdf(
        replica_url,
        replica_pdf_path
    )

    # -----------------------------------------------------
    # STORE RESULTS
    # -----------------------------------------------------
    final_results.append({

        "Live URL": live_url,
        "Replica URL": replica_url,

        "Live Report Link": live_data["Report Link"],
        "Replica Report Link": replica_data["Report Link"],

        "Live Status": live_data["Status"],
        "Replica Status": replica_data["Status"],

        "Live Performance": live_data["Performance"],
        "Replica Performance": replica_data["Performance"],

        "Live Accessibility": live_data["Accessibility"],
        "Replica Accessibility": replica_data["Accessibility"],

        "Live Best Practices": live_data["Best Practices"],
        "Replica Best Practices": replica_data["Best Practices"],

        "Live SEO": live_data["SEO"],
        "Replica SEO": replica_data["SEO"],

        "Live FCP": live_data["FCP"],
        "Replica FCP": replica_data["FCP"],

        "Live LCP": live_data["LCP"],
        "Replica LCP": replica_data["LCP"],

        "Live TBT": live_data["TBT"],
        "Replica TBT": replica_data["TBT"],

        "Live CLS": live_data["CLS"],
        "Replica CLS": replica_data["CLS"],

        "Live Speed Index": live_data["Speed Index"],
        "Replica Speed Index": replica_data["Speed Index"],

        "Live PDF Path": live_pdf_path,
        "Replica PDF Path": replica_pdf_path,
    })

    time.sleep(3)

# =========================================================
# EXPORT
# =========================================================

output_df = pd.DataFrame(final_results)

output_df.to_excel(
    OUTPUT_FILE,
    index=False
)



# ----------------------------------------
# UPDATE MASTER WORKBOOK
# ----------------------------------------

save_snapshot_sheet(final_results, strategy)
update_historical_data(final_results, strategy)
generate_url_trends()
generate_summary_dashboard()

print("✅ Master workbook updated successfully!")

print(f"\n{'=' * 60}")
print("Completed Successfully!")
print(f"Output saved:")
print(OUTPUT_FILE)
print(f"{'=' * 60}")