"""
Facebook Direct Poster v1
Posts a photo/video + text as an original post into a list of Facebook groups
using Selenium + Chrome. No link-sharing — the media is uploaded directly
into each group's composer.

Folder layout expected next to this script:

    campaigns/
        <any_name>.jpg | .png | .mp4 | ...   (exactly one media file)
        post.txt                              (post text, any encoding)
    Page_Bard/groups.xlsx        (columns: Group Name | URL | Category)
    Page_Liberman/groups.xlsx    (columns: Group Name | URL | Category)

Run:
    python facebook_direct_poster_v1.py
"""

from __future__ import annotations

import sys
import time
import random
import logging
from dataclasses import dataclass
from pathlib import Path
from typing import Optional

import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import (
    TimeoutException,
    NoSuchElementException,
    ElementClickInterceptedException,
    StaleElementReferenceException,
)

# --------------------------------------------------------------------------- #
# Configuration
# --------------------------------------------------------------------------- #

BASE_DIR = Path(__file__).resolve().parent
CAMPAIGNS_DIR = BASE_DIR / "campaigns"
REPORTS_DIR = BASE_DIR / "posting_reports"
ERRORS_DIR = REPORTS_DIR / "errors"

PAGES = {
    "1": ("Bard", BASE_DIR / "Page_Bard" / "groups.xlsx"),
    "2": ("Liberman", BASE_DIR / "Page_Liberman" / "groups.xlsx"),
}

MEDIA_EXTENSIONS = {".jpg", ".jpeg", ".png", ".gif", ".mp4", ".mov", ".m4v"}
VIDEO_EXTENSIONS = {".mp4", ".mov", ".m4v"}

SKIP_CATEGORY = "Error_NoPostField"

# Confirmed via screenshots during a large batch run: the composer dialog
# opens fine and auto-focuses its textbox, but Facebook's UI responds much
# slower than usual after a big burst of automated activity on the account
# -- no block, no captcha, just genuinely slower rendering. 10s, which was
# plenty for the first clean batch, wasn't enough here and caused a 100%
# failure rate across an entire run. Raised to give the UI more room.
WAIT_TIMEOUT = 25
PHOTO_ATTACH_TIMEOUT = 20
VIDEO_ATTACH_TIMEOUT = 60
# Facebook doesn't gate the Post button on server-side video processing --
# once the local upload preview renders it's clickable, so this only needs
# to cover how long the button takes to appear/settle, not real processing.
POST_BUTTON_TIMEOUT = 30

# Pause between groups so posting doesn't look automated to Facebook (a
# perfectly steady interval is itself a bot signal, hence the jitter).
BETWEEN_GROUPS_DELAY_MIN = 30
BETWEEN_GROUPS_DELAY_MAX = 45

WRITE_SOMETHING_XPATH = (
    "//div[@role='button']"
    "[contains(., 'Write something') or contains(., 'Написать что-нибудь') "
    "or contains(., 'Что у вас нового') or contains(., \"What's on your mind\")]"
)

# A Facebook page typically has several hidden <input type="file"> elements
# scattered around (cover photo, avatar, other widgets, etc.), not just the
# composer's own. Everything below is scoped to the open "Create post"
# dialog specifically, so we never grab the wrong one.
DIALOG_XPATH = "//div[@role='dialog']"

# Facebook shows one of these once the file has actually finished
# uploading/processing and is attached to the composer. Waiting for this
# (rather than a fixed sleep) avoids clicking "Post" before a video is
# ready, which silently posts the text alone with no attachment.
MEDIA_ATTACHED_XPATH = (
    ".//div[@aria-label='Remove photo' or @aria-label='Remove video' "
    "or @aria-label='Удалить фото' or @aria-label='Удалить видео']"
    " | .//video"
    " | .//img[contains(@src, 'blob:')]"
)

# The post text contains a link (the tinyurl ticket link), and Facebook
# auto-generates a link preview for it -- pulling thumbnail images from
# the linked site itself, which is why unrelated pictures (from the
# ticketing page, not our video) showed up attached. Confirmed via
# DevTools: its close button is aria-label="Remove link preview from your
# post". A human has to click that to clear it before uploading the real
# file; skipping that step is exactly why our own "media attached" check
# above could pass instantly on the link preview instead of the file we
# actually sent.
REMOVE_ATTACHMENT_XPATH = (
    ".//div[@aria-label='Remove link preview from your post' "
    "or @aria-label='Remove photo' or @aria-label='Remove video' "
    "or @aria-label='Удалить фото' or @aria-label='Удалить видео' "
    "or @aria-label='Close' or @aria-label='Закрыть' "
    "or @aria-label='Delete' or @aria-label='Удалить']"
)

POST_BUTTON_XPATH = (
    ".//div[@aria-label='Post' or @aria-label='Опубликовать']"
    "[@role='button']"
)

# Some groups make posting conditional on answering membership questions
# first -- a dialog Facebook can show instead of (or on top of) the normal
# composer. There's no reliable single selector for it across groups, so
# this matches on visible page text instead, same approach as the earlier
# link-share script that established this convention.
ADMIN_QUESTIONS_INDICATORS = (
    "membership question",
    "answer the following",
    "do you live in",
    "answer question",
    "ответьте на вопрос",
    "вопросы администратора",
)


# --------------------------------------------------------------------------- #
# Logging
# --------------------------------------------------------------------------- #

def setup_logging() -> Path:
    # Windows terminals often default to a legacy ANSI codepage that can't
    # render Cyrillic; force UTF-8 so log output isn't garbled into "?".
    if sys.platform == "win32":
        for stream in (sys.stdout, sys.stderr):
            if hasattr(stream, "reconfigure"):
                stream.reconfigure(encoding="utf-8", errors="replace")

    REPORTS_DIR.mkdir(exist_ok=True)
    timestamp = time.strftime("%Y%m%d_%H%M%S")
    log_path = REPORTS_DIR / f"report_{timestamp}.log"

    logger = logging.getLogger()
    logger.setLevel(logging.INFO)
    logger.handlers.clear()

    fmt = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s", "%H:%M:%S")

    console = logging.StreamHandler(sys.stdout)
    console.setFormatter(fmt)
    logger.addHandler(console)

    file_handler = logging.FileHandler(log_path, encoding="utf-8")
    file_handler.setFormatter(fmt)
    logger.addHandler(file_handler)

    return log_path


# --------------------------------------------------------------------------- #
# Campaign discovery
# --------------------------------------------------------------------------- #

@dataclass
class Campaign:
    media_path: Path
    text: str
    is_video: bool


def find_campaign() -> Campaign:
    if not CAMPAIGNS_DIR.exists():
        raise FileNotFoundError(f"Campaigns folder not found: {CAMPAIGNS_DIR}")

    media_candidates = [
        p for p in CAMPAIGNS_DIR.iterdir()
        if p.is_file() and p.suffix.lower() in MEDIA_EXTENSIONS
    ]
    if not media_candidates:
        raise FileNotFoundError(f"No media file found in {CAMPAIGNS_DIR}")
    if len(media_candidates) > 1:
        raise ValueError(
            f"Expected exactly one media file in {CAMPAIGNS_DIR}, found: "
            f"{[p.name for p in media_candidates]}"
        )
    media_path = media_candidates[0]

    text_path = CAMPAIGNS_DIR / "post.txt"
    if not text_path.exists():
        raise FileNotFoundError(f"post.txt not found in {CAMPAIGNS_DIR}")

    text = read_text_any_encoding(text_path)
    is_video = media_path.suffix.lower() in VIDEO_EXTENSIONS

    return Campaign(media_path=media_path, text=text, is_video=is_video)


def read_text_any_encoding(path: Path) -> str:
    for encoding in ("utf-8-sig", "utf-8", "cp1251", "latin-1"):
        try:
            return path.read_text(encoding=encoding).strip()
        except (UnicodeDecodeError, UnicodeError):
            continue
    return path.read_bytes().decode("utf-8", errors="replace").strip()


# --------------------------------------------------------------------------- #
# Page / group selection
# --------------------------------------------------------------------------- #

def choose_page() -> tuple[str, Path]:
    print("\nSelect page:")
    for key, (name, _) in PAGES.items():
        print(f"  [{key}] {name}")

    choice = input("Page number: ").strip()
    if choice not in PAGES:
        print("Invalid choice.")
        sys.exit(1)

    name, groups_path = PAGES[choice]
    if not groups_path.exists():
        raise FileNotFoundError(f"groups.xlsx not found: {groups_path}")

    return name, groups_path


def load_groups(groups_path: Path) -> pd.DataFrame:
    df = pd.read_excel(groups_path, dtype=str)
    required = {"Group Name", "URL", "Category"}
    missing = required - set(df.columns)
    if missing:
        raise ValueError(f"{groups_path} is missing columns: {missing}")
    df = df.dropna(subset=["URL"]).reset_index(drop=True)
    return df


def choose_categories(df: pd.DataFrame) -> list[str]:
    categories = sorted(df["Category"].dropna().unique().tolist())

    print("\nAvailable categories:")
    print("  [0] All categories")
    for i, cat in enumerate(categories, start=1):
        count = (df["Category"] == cat).sum()
        print(f"  [{i}] {cat} ({count} groups)")

    raw = input("Select categories (e.g. 1,3,5 or 0 for all): ").strip()
    if raw == "0":
        return categories

    try:
        indices = [int(x.strip()) for x in raw.split(",") if x.strip()]
    except ValueError:
        print("Invalid input.")
        sys.exit(1)

    selected = []
    for i in indices:
        if not (1 <= i <= len(categories)):
            print(f"Invalid category number: {i}")
            sys.exit(1)
        selected.append(categories[i - 1])

    return selected


# --------------------------------------------------------------------------- #
# Selenium helpers
# --------------------------------------------------------------------------- #

def build_driver() -> webdriver.Chrome:
    options = Options()
    options.add_argument("--start-maximized")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option("useAutomationExtension", False)
    return webdriver.Chrome(options=options)


def js_click(driver, element) -> None:
    driver.execute_script("arguments[0].click();", element)


def get_dialog(driver):
    """
    Always fetch the "Create post" dialog fresh from the driver rather than
    reusing a cached WebElement. Facebook appears to first render a
    skeleton/loading version of the dialog and then swap in the fully
    loaded one shortly after (and further re-renders can happen when we
    dismiss the suggested-photo attachment) -- any reference held across
    those swaps throws "stale element reference", which is what testing
    kept hitting a few seconds into every group.
    """
    return driver.find_element(By.XPATH, DIALOG_XPATH)


def check_for_admin_questions(driver) -> bool:
    try:
        page_text = driver.page_source.lower()
    except (NoSuchElementException, StaleElementReferenceException):
        return False
    return any(indicator in page_text for indicator in ADMIN_QUESTIONS_INDICATORS)


def save_error_screenshot(driver, group_name: str) -> None:
    try:
        ERRORS_DIR.mkdir(parents=True, exist_ok=True)
        safe_name = "".join(c if c.isalnum() else "_" for c in group_name)[:40]
        path = ERRORS_DIR / f"error_{safe_name}_{time.strftime('%H%M%S')}.png"
        driver.save_screenshot(str(path))
        logging.info(f"  Error screenshot: {path}")
    except Exception as exc:  # noqa: BLE001 - screenshotting must never break the run
        logging.debug(f"  Could not save error screenshot: {exc}")


def open_composer(driver, wait: WebDriverWait) -> None:
    write_something = wait.until(
        EC.element_to_be_clickable((By.XPATH, WRITE_SOMETHING_XPATH))
    )
    js_click(driver, write_something)
    wait.until(EC.presence_of_element_located((By.XPATH, DIALOG_XPATH)))
    # Give the composer a moment to settle past its initial skeleton render
    # before anything tries to grab elements inside it.
    time.sleep(1.5)


def _pick_media_file_input(file_inputs):
    for inp in file_inputs:
        accept = (inp.get_attribute("accept") or "").lower()
        if "video" in accept or "image" in accept:
            return inp
    return file_inputs[0]


def clear_suggested_media(driver) -> None:
    """
    A link in the post text (e.g. a ticket URL) makes Facebook
    auto-generate a link preview with images pulled from that site --
    unrelated to our own media. Click every close/remove control in the
    dialog (confirmed via DevTools: aria-label="Remove link preview from
    your post") so our own upload starts from an empty slot, matching what
    a human has to do manually (click its X first).
    """
    try:
        close_buttons = get_dialog(driver).find_elements(By.XPATH, REMOVE_ATTACHMENT_XPATH)
    except (NoSuchElementException, StaleElementReferenceException):
        return

    for btn in close_buttons:
        try:
            js_click(driver, btn)
        except (StaleElementReferenceException, ElementClickInterceptedException):
            pass
    if close_buttons:
        time.sleep(1)


def attach_media(driver, media_path: Path, is_video: bool) -> None:
    # clear_suggested_media() is called separately by post_to_group, before
    # this -- observed manual behavior is: type the text first, then close
    # the suggested photo, then upload the real file.

    # Baseline count *after* clearing, so the wait below only succeeds once
    # our own file actually attaches, not on a leftover suggestion that
    # failed to clear or on a stale match. Re-fetch the dialog fresh right
    # before every use -- see get_dialog().
    baseline = len(get_dialog(driver).find_elements(By.XPATH, MEDIA_ATTACHED_XPATH))

    # Deliberately NOT clicking the "Attach a photo or video" button here.
    # Facebook's own click handler on that button calls .click() on the
    # underlying hidden <input type="file">, and a script-triggered click
    # on a file input still opens the real native OS "Open File" dialog --
    # a window outside the browser that Selenium can't see or control. It
    # steals OS-level focus, which is exactly why every click after the
    # upload (closing the suggestion, hitting Post) went nowhere even
    # though the file itself did attach (via send_keys below, which sets
    # the input's files directly through the WebDriver protocol and does
    # NOT open that dialog). Instead, find the hidden input straight away
    # -- Facebook renders it whether or not the button was clicked -- and
    # send the file path to it directly.
    file_inputs = WebDriverWait(driver, WAIT_TIMEOUT).until(
        lambda d: get_dialog(d).find_elements(By.CSS_SELECTOR, "input[type='file']") or False
    )
    file_input = _pick_media_file_input(file_inputs)
    file_input.send_keys(str(media_path.resolve()))

    # This only confirms the upload *started* (a preview/thumbnail appeared),
    # not that it's fully processed -- see wait_for_post_button_ready for that.
    timeout = VIDEO_ATTACH_TIMEOUT if is_video else PHOTO_ATTACH_TIMEOUT
    WebDriverWait(driver, timeout).until(
        lambda d: len(get_dialog(d).find_elements(By.XPATH, MEDIA_ATTACHED_XPATH)) > baseline
    )


def find_post_textbox(driver, wait: WebDriverWait):
    def _locate(d):
        try:
            dialog = get_dialog(d)
        except (NoSuchElementException, StaleElementReferenceException):
            return False
        candidates = dialog.find_elements(
            By.CSS_SELECTOR, "div[contenteditable='true'][role='textbox']"
        )
        for el in candidates:
            try:
                placeholder = (el.get_attribute("aria-placeholder") or "")
            except StaleElementReferenceException:
                continue
            if "comment" not in placeholder.lower() and "коммент" not in placeholder.lower():
                return el
        return False

    return wait.until(_locate)


def type_post_text(driver, textbox, text: str) -> None:
    # Typed natively via Selenium's key events rather than the OS clipboard.
    # A clipboard round-trip (pyperclip + Ctrl+V) goes through Windows'
    # legacy ANSI codepage on some machines and silently mangles Cyrillic
    # into "?", even though the source text and log file are correct UTF-8.
    js_click(driver, textbox)
    time.sleep(0.5)
    textbox.send_keys(text)
    time.sleep(0.5)


def wait_for_post_button_ready(driver, group_name: str, timeout: int):
    """
    Wait for the Post button to be present and clickable.

    Confirmed live (aria-disabled='None' logged on every run, video or
    photo) that Facebook does not set aria-disabled on this button at all
    in this UI -- it stays absent whether or not the button is actually
    usable, so waiting for it to become "false" never resolves and always
    burns the full timeout. Facebook also doesn't require the video to
    finish server-side processing before posting: once the local upload
    preview renders (which attach_media already waits for), Post is
    clickable immediately and processing continues after the post goes
    through. So just wait for clickability, with a short settle delay
    for the preview to finish rendering.
    """
    WebDriverWait(driver, timeout).until(
        lambda d: get_dialog(d).find_elements(By.XPATH, POST_BUTTON_XPATH) or False
    )
    time.sleep(3)
    # Re-fetch right before returning -- the settle delay above is enough
    # time for a re-render to make an earlier reference stale.
    button = get_dialog(driver).find_element(By.XPATH, POST_BUTTON_XPATH)
    logging.info(f"  Post button ready in '{group_name}'")
    return button


def click_post_button(driver, button) -> None:
    js_click(driver, button)


# --------------------------------------------------------------------------- #
# Posting flow
# --------------------------------------------------------------------------- #

def post_to_group(driver, campaign: Campaign, group_name: str, group_url: str) -> str:
    """
    Never raises -- every failure is caught, logged, screenshotted, and
    reported back as one of a fixed set of status strings so the run keeps
    going to the next group. This mirrors the earlier link-share script's
    convention: specific Error_* statuses the operator can later copy into
    groups.xlsx's Category column to make future runs skip that group via
    SKIP_CATEGORY, rather than one generic "Error" that hides the reason.
    """
    wait = WebDriverWait(driver, WAIT_TIMEOUT)

    try:
        driver.get(group_url)
        time.sleep(2)

        if check_for_admin_questions(driver):
            logging.warning(f"Admin membership questions detected in '{group_name}'")
            save_error_screenshot(driver, group_name)
            return "Error_NeedsQuestions"

        try:
            open_composer(driver, wait)
        except TimeoutException:
            logging.warning(f"Post field never opened in '{group_name}' (group likely closed or restricted)")
            save_error_screenshot(driver, group_name)
            return "Error_NoPostField"

        if check_for_admin_questions(driver):
            logging.warning(f"Admin membership questions appeared after opening composer in '{group_name}'")
            save_error_screenshot(driver, group_name)
            return "Error_NeedsQuestions"

        # Observed manual order: text first, then dismiss the suggested
        # photo, then upload the real file -- not the other way around.
        textbox = find_post_textbox(driver, wait)
        type_post_text(driver, textbox, campaign.text)

        clear_suggested_media(driver)

        try:
            attach_media(driver, campaign.media_path, campaign.is_video)
        except TimeoutException:
            logging.warning(
                f"Media upload never started in '{group_name}' "
                f"(timed out after {VIDEO_ATTACH_TIMEOUT if campaign.is_video else PHOTO_ATTACH_TIMEOUT}s) "
                "-- skipping to avoid posting without the attachment"
            )
            save_error_screenshot(driver, group_name)
            return "Error"

        try:
            post_button = wait_for_post_button_ready(driver, group_name, POST_BUTTON_TIMEOUT)
        except TimeoutException:
            logging.warning(
                f"Post button never appeared in '{group_name}' "
                f"(timed out after {POST_BUTTON_TIMEOUT}s) -- skipping"
            )
            save_error_screenshot(driver, group_name)
            return "Error_NoButton"

        click_post_button(driver, post_button)
        time.sleep(3)

        return "Posted"

    except TimeoutException:
        logging.warning(f"Timeout waiting for an element in '{group_name}'")
        save_error_screenshot(driver, group_name)
        return "Error"
    except (NoSuchElementException, ElementClickInterceptedException, StaleElementReferenceException) as exc:
        logging.warning(f"Selenium error in '{group_name}': {exc}")
        save_error_screenshot(driver, group_name)
        return "Error"
    except Exception as exc:  # noqa: BLE001 - last-resort catch to keep the run going
        logging.warning(f"Unexpected error in '{group_name}': {exc}")
        save_error_screenshot(driver, group_name)
        return "Error"


# --------------------------------------------------------------------------- #
# Main
# --------------------------------------------------------------------------- #

def main() -> None:
    log_path = setup_logging()
    logging.info(f"Log file: {log_path}")

    try:
        campaign = find_campaign()
    except (FileNotFoundError, ValueError) as exc:
        logging.error(str(exc))
        sys.exit(1)

    logging.info(f"Media: {campaign.media_path.name} ({'video' if campaign.is_video else 'photo'})")
    logging.info(f"Text: {campaign.text[:80]}{'...' if len(campaign.text) > 80 else ''}")

    page_name, groups_path = choose_page()
    df = load_groups(groups_path)
    categories = choose_categories(df)

    targets = df[df["Category"].isin(categories)]
    targets = targets[targets["Category"] != SKIP_CATEGORY]
    skipped_count = df[df["Category"].isin(categories) & (df["Category"] == SKIP_CATEGORY)].shape[0]

    if targets.empty:
        logging.error("No groups to post to after filtering.")
        sys.exit(1)

    print(f"\nPage: {page_name}")
    print(f"Groups to post: {len(targets)}")
    if skipped_count:
        print(f"Groups skipped ({SKIP_CATEGORY}): {skipped_count}")
    confirm = input("Proceed? [y/N]: ").strip().lower()
    if confirm != "y":
        print("Aborted.")
        sys.exit(0)

    driver = build_driver()
    driver.get("https://www.facebook.com/")
    input("\nLog in to Facebook in the opened browser window, then press ENTER to start posting...")

    results = []
    group_rows = list(targets.iterrows())
    for i, (_, row) in enumerate(group_rows):
        group_name = row["Group Name"]
        group_url = row["URL"]
        logging.info(f"Posting to '{group_name}' ({group_url})")

        status = post_to_group(driver, campaign, group_name, group_url)
        results.append({"Group Name": group_name, "URL": group_url, "Status": status})
        logging.info(f"  -> {status}")

        if i < len(group_rows) - 1:
            delay = random.uniform(BETWEEN_GROUPS_DELAY_MIN, BETWEEN_GROUPS_DELAY_MAX)
            logging.info(f"  Waiting {delay:.0f}s before the next group...")
            time.sleep(delay)

    driver.quit()

    report_df = pd.DataFrame(results)
    csv_path = REPORTS_DIR / f"report_{time.strftime('%Y%m%d_%H%M%S')}.csv"
    report_df.to_csv(csv_path, index=False, encoding="utf-8-sig")

    posted = (report_df["Status"] == "Posted").sum()
    errors = report_df["Status"].str.startswith("Error").sum()
    logging.info(f"Done. Posted: {posted}, Errors: {errors}, Total: {len(report_df)}")
    logging.info(f"CSV report: {csv_path}")

    # Break down errors by exact status so groups can be copied straight
    # into groups.xlsx's Category column (matching SKIP_CATEGORY's
    # convention of "Error_..." categories that future runs skip).
    error_rows = report_df[report_df["Status"].str.startswith("Error")]
    if not error_rows.empty:
        logging.info("Groups that need attention (copy into groups.xlsx's Category column):")
        for status, group in error_rows.groupby("Status"):
            logging.info(f"  {status} ({len(group)}):")
            for name in group["Group Name"]:
                logging.info(f"    - {name}")


if __name__ == "__main__":
    main()
