"""Logging for the PageSpeed runners.

Writes a timestamped log next to the results workbook and mirrors it on
the terminal. Existing print() calls are captured so the file shows the
same progress as the console, plus tracebacks for caught failures.
"""

import builtins
import logging
import os
import sys
import time
from urllib.parse import urlparse

LOGGER_NAME = "pagespeed"


def setup_run_log(log_dir, strategy, timestamp):
    os.makedirs(log_dir, exist_ok=True)
    log_path = os.path.join(log_dir, f"{strategy}_{timestamp}.log")

    logger = logging.getLogger(LOGGER_NAME)
    logger.setLevel(logging.DEBUG)
    logger.propagate = False
    logger.handlers.clear()

    formatter = logging.Formatter(
        "%(asctime)s | %(levelname)-7s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    file_handler = logging.FileHandler(log_path, encoding="utf-8")
    file_handler.setLevel(logging.DEBUG)
    file_handler.setFormatter(formatter)

    stream_handler = logging.StreamHandler(sys.stdout)
    stream_handler.setLevel(logging.INFO)
    stream_handler.setFormatter(formatter)

    logger.addHandler(file_handler)
    logger.addHandler(stream_handler)

    # PageSpeed PDFs omit FontBBox on some fonts. pdfminer warns once per
    # glyph and the scores are still parsed.
    logging.getLogger("pdfminer").setLevel(logging.ERROR)

    _install_print_hook(logger)
    logger.info("Logging to %s", log_path)
    return logger, log_path


def attach_page_logs(page):
    """Record browser console errors and uncaught page exceptions."""
    logger = logging.getLogger(LOGGER_NAME)

    def on_console(msg):
        try:
            if msg.type == "error":
                logger.error("Browser console: %s", msg.text)
        except Exception:
            logger.debug("Could not read a browser console message", exc_info=True)

    def on_page_error(err):
        logger.error("Page error: %s", err)

    page.on("console", on_console)
    page.on("pageerror", on_page_error)


def page_error_snippet(html, keywords, radius=180):
    """Return the matched PSI error keyword and a short slice of the page."""
    for keyword in keywords:
        index = html.find(keyword)
        if index == -1:
            continue
        start = max(0, index - radius)
        end = min(len(html), index + len(keyword) + radius)
        snippet = " ".join(html[start:end].split())
        return keyword, snippet
    return None, ""


PSI_ERROR_MARKERS = (
    "Unable to resolve",
    "Oops! Something went wrong",
    "RPC::UNREACHABLE",
    "UNABLE_TO_RETRY",
    "Lighthouse returned error",
)

PSI_REPORT_MARKERS = (
    "Diagnose performance issues",
    "First Contentful Paint",
    "Largest Contentful Paint",
    "Speed Index",
)


def is_archived_report_url(url):
    """True for /analysis/<page>/<id>, not the in-progress /analysis?url= form."""
    parsed = urlparse(url)
    parts = [part for part in parsed.path.split("/") if part]
    return (
        parsed.netloc.endswith("pagespeed.web.dev")
        and len(parts) >= 3
        and parts[0] == "analysis"
    )


# Minimum gap between two report loads (live, then replica).
REPORT_GAP_SECONDS = 45
_last_report_finished = 0.0
# A normal Chrome UA. HeadlessChrome is what PageSpeed rate-limited.
NORMAL_CHROME_UA = (
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/156.0.0.0 Safari/537.36"
)


def wait_for_pagespeed_slot():
    """Keep a gap between reports so the next website run is not rate limited."""
    logger = logging.getLogger(LOGGER_NAME)
    if not _last_report_finished:
        return
    remaining = REPORT_GAP_SECONDS - (time.time() - _last_report_finished)
    if remaining > 1:
        logger.info(
            "Waiting %ss before the next PageSpeed page so the website report can finish",
            int(remaining),
        )
        time.sleep(remaining)


def mark_pagespeed_finished():
    global _last_report_finished
    _last_report_finished = time.time()


def open_pagespeed_browser(playwright):
    """Open normal Chrome so the website PDF can be saved.

    HeadlessChrome was rate-limited (HTTP 429). A visible Chrome window
    matches the browser that produced the QA PDF.
    Set PSI_HEADLESS=1 to force headless Chrome with a normal user agent.
    """
    logger = logging.getLogger(LOGGER_NAME)
    force_headless = os.environ.get("PSI_HEADLESS", "").strip().lower() in (
        "1",
        "true",
        "yes",
    )
    modes = [True] if force_headless else [False, True]
    last_error = None

    for headless in modes:
        try:
            browser = playwright.chromium.launch(
                headless=headless,
                channel="chrome",
                ignore_default_args=["--enable-automation"],
                args=[
                    "--disable-blink-features=AutomationControlled",
                    "--disable-dev-shm-usage",
                ],
            )
        except Exception as exc:
            last_error = exc
            logger.warning("Could not open Chrome (headless=%s): %s", headless, exc)
            continue

        context = browser.new_context(
            user_agent=NORMAL_CHROME_UA,
            viewport={"width": 1400, "height": 900},
            locale="en-US",
        )
        page = context.new_page()
        page.add_init_script(
            "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
        )
        state = {"rate_limited": False}

        def on_response(response, state=state):
            if response.status == 429 and "PagespeedUi" in response.url:
                state["rate_limited"] = True
                logger.error(
                    "PageSpeed website returned HTTP 429 (rate limit) for the analysis request."
                )

        page.on("response", on_response)
        logger.info(
            "Opened %s Chrome. The PDF is still saved from the PageSpeed website.",
            "headless" if headless else "a visible",
        )
        return browser, page, state

    raise RuntimeError(f"Could not open Chrome: {last_error}")


def wait_for_pagespeed(page, timeout_s=300, response_state=None):
    """Wait until a report is ready, or return the PageSpeed error text.

    The old wait for '**/analysis/**' never matches the in-progress URL
    /analysis?url=... and also never arrives when the run fails, so the
    script sat for the full timeout.
    """
    logger = logging.getLogger(LOGGER_NAME)
    deadline = time.time() + timeout_s
    announced_running = False
    content_ready_at = None

    while time.time() < deadline:
        _dismiss_cookie_banner(page)
        current = page.url

        if response_state and response_state.get("rate_limited"):
            snippet = _page_snippet(page)
            message = (
                "PageSpeed rate limit (HTTP 429). "
                "No PDF was saved because the website refused to start the report. "
                f"Page says: {snippet}"
            )
            logger.error(message)
            return current, message, 0

        if is_archived_report_url(current):
            logger.info("PageSpeed report URL: %s", current)
            return current, None, 0

        text = _page_text(page)
        if text is None:
            page.wait_for_timeout(3000)
            continue

        for marker in PSI_ERROR_MARKERS:
            if marker in text:
                snippet = " ".join(text.split())[:500]
                message = f"PageSpeed error: {marker}. Page says: {snippet}"
                logger.error(message)
                return current, message, 20

        if all(marker in text for marker in PSI_REPORT_MARKERS):
            if content_ready_at is None:
                content_ready_at = time.time()
                logger.info("PageSpeed report content is visible; waiting for the final report URL")
            elif time.time() - content_ready_at >= 20:
                logger.info("Using the visible PageSpeed report at %s", current)
                return current, None, 0
        else:
            content_ready_at = None
            if "Running analysis" in text and not announced_running:
                logger.info("PageSpeed is running the analysis")
                announced_running = True

        page.wait_for_timeout(3000)

    snippet = _page_snippet(page)
    message = (
        f"Timed out after {timeout_s}s waiting for the PageSpeed report. "
        f"URL stayed {page.url}. Page says: {snippet}"
    )
    logger.error(message)
    return page.url, message, 20


def _dismiss_cookie_banner(page):
    for selector in (
        "button:has-text('Ok, Got it.')",
        "button:has-text('Accept all')",
        "button:has-text('Accept All')",
    ):
        try:
            button = page.locator(selector)
            if button.count() > 0:
                button.first.click(timeout=1500)
                return
        except Exception:
            continue


def _page_text(page):
    try:
        return page.locator("body").inner_text(timeout=5000)
    except Exception as exc:
        logging.getLogger(LOGGER_NAME).debug(
            "PageSpeed page text not readable yet: %s", exc
        )
        return None


def _page_snippet(page):
    text = _page_text(page)
    if not text:
        return ""
    return " ".join(text.split())[:500]


def _level_for(message):
    lowered = message.lower()
    if any(token in lowered for token in ("failed", "failure", "error", "exception")):
        return logging.ERROR
    if "⚠" in message or any(
        token in lowered
        for token in ("warning", "timed out", "timeout", "missing", "differs", "suspicious")
    ):
        return logging.WARNING
    return logging.INFO


def _install_print_hook(logger):
    original_print = builtins.print

    def logged_print(*args, **kwargs):
        file = kwargs.get("file")
        if file not in (None, sys.stdout):
            original_print(*args, **kwargs)
            return
        sep = kwargs.get("sep", " ")
        message = sep.join(str(arg) for arg in args)
        level = _level_for(message) if message.strip() else logging.INFO
        # Pass the text as an argument so % in URLs is not treated as a format.
        logger.log(level, "%s", message)

    builtins.print = logged_print
