# Facebook Direct Poster v1

Posts a photo/video + text directly into a list of Facebook groups via Selenium.
No link-sharing — each group gets an original post with the media attached.

## Folder layout

```
campaigns/
    <any_name>.jpg|.png|.mp4|...   # exactly one media file
    post.txt                       # post text, any encoding
Page_Bard/groups.xlsx              # columns: Group Name | URL | Category
Page_Liberman/groups.xlsx          # columns: Group Name | URL | Category
facebook_direct_poster_v1.py
```

Rows whose `Category` is `Error_NoPostField` are always skipped.

## Setup

```
pip install -r requirements.txt
```

Requires Chrome + a matching chromedriver on PATH (Selenium 4 manages this
automatically via Selenium Manager in most environments).

## Run

```
python facebook_direct_poster_v1.py
```

1. Choose the page (`1` Bard / `2` Liberman)
2. Choose categories from `groups.xlsx` (comma list, or `0` for all)
3. Confirm the run
4. A Chrome window opens — log in to Facebook manually, then press ENTER
5. The script posts to each selected group and writes a log + CSV report to
   `posting_reports/`

## Notes on this workflow

Since this script is iterated on with Claude Code directly (rather than
pasted into PyCharm), edits and test runs happen in this session. Real
Facebook posting still needs to run somewhere with a visible Chrome browser
and your logged-in session — copy the script + folders to your machine (or
run this session with a local runtime) to do live posting runs, then report
back any log/report output here for further debugging.
