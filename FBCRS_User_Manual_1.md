# FBCRS Document Analysis & Recommendation Engine
## User Manual — Version 8.5

---

## Contents

1. [What this tool does](#1-what-this-tool-does)
2. [What this tool does *not* do](#2-what-this-tool-does-not-do)
3. [Installation](#3-installation)
4. [Quick start](#4-quick-start)
5. [Reading the report — all 21 columns](#5-reading-the-report--all-21-columns)
6. [How classification works](#6-how-classification-works)
7. [Retention and expiry dates](#7-retention-and-expiry-dates)
8. [Duplicate detection](#8-duplicate-detection)
9. [The review workflow](#9-the-review-workflow)
10. [Maintaining the rulebook](#10-maintaining-the-rulebook)
11. [Performance tuning](#11-performance-tuning)
12. [Troubleshooting](#12-troubleshooting)
13. [Building and sharing an .exe](#13-building-and-sharing-an-exe)
14. [Known limitations](#14-known-limitations)
15. [Appendix A — configuration reference](#appendix-a--configuration-reference)
16. [Appendix B — file inventory](#appendix-b--file-inventory)

---

## 1. What this tool does

The engine scans a folder of documents and produces an Excel inventory that
suggests, for each file, which records series it belongs to and what should
happen to it.

For every file it reports:

- a **records series code** from your classification rulebook, or `UNCLASSIFIED`
- the **retention period** and calculated **expiry date**
- a **recommended action** — Destroy, Do Nothing, Manual Review, or Transfer to Archives
- a **confidence level** and a plain-English explanation of why
- whether the file is an exact **duplicate** of another file in the scan
- up to three **alternate codes** when the top match is uncertain

It reads inside `.docx`, `.pdf`, `.xlsx`, `.xlsm`, `.pptx`, and `.txt` files.
Other file types are still inventoried, but classified on filename alone.

---

## 2. What this tool does *not* do

Read this section before acting on any output.

**It does not delete, move, or modify a single file.** The tool is read-only.
Every action in the report is a *recommendation* for a human to approve. Nothing
happens to your documents.

**`DESTROY` is a suggestion, never an instruction.** Do not run a bulk deletion
from this report. Every destroy recommendation needs review by someone with the
authority to approve disposition.

**It does not detect near-duplicates.** Only byte-for-byte identical files. Two
versions of the same document with a one-word change are treated as unrelated.

---

## 3. Installation

### 3.1 What you need

Two files, **in the same folder**:

```
C:\FBCRS\
├── FBCRS_Analyzer.exe          (the application)
├── FBCRS_Master_Full.xlsx      (the classification rulebook)
└── repair_rulebook.exe         (rulebook checker - optional)
```

The rulebook **must** be named exactly `FBCRS_Master_Full.xlsx` and **must** sit
beside the program. This is the single most common setup mistake.

### 3.2 Where to install it

**Install to a local drive — `C:\FBCRS\` — not a network share.**

The program writes temporary checkpoint files thousands of times during a scan.
Running it from a mapped drive or UNC path (`Z:\...` or `\\SERVER\...`) causes
intermittent "Access is denied" failures partway through a scan, and is
dramatically slower. You can *scan* network folders; just don't *run* from one.

### 3.3 Python setup (skip if using the .exe)

Requires Python 3.10 or later. Install the dependencies once:

```
pip install customtkinter openpyxl pypdf python-docx python-pptx
```

| Package | Needed for |
|---|---|
| `customtkinter` | the interface — **required** |
| `openpyxl` | reading the rulebook and writing reports — **required** |
| `pypdf` | reading inside PDF files — optional |
| `python-docx` | reading inside Word files — optional |
| `python-pptx` | reading inside PowerPoint files — optional |

The three optional packages degrade gracefully. Without `pypdf`, for example,
PDFs are still inventoried and classified on filename, just not on content.

### 3.4 Where reports are saved

Automatically, to the current user's Documents folder:

```
C:\Users\<YourName>\Documents\FBCRS_Reports\
```

This is resolved at runtime, so on a colleague's machine it saves to *their*
Documents folder. If Documents isn't writable, the program falls back to
`%LOCALAPPDATA%\FBCRS_Reports`, then to the system temp folder. The active path
is displayed in small grey text at the bottom of the window.

---

## 4. Quick start

1. **Close the rulebook.** `FBCRS_Master_Full.xlsx` must not be open in Excel.
   Excel locks the file and the program will refuse to start.

2. **Launch.** Double-click `FBCRS_Analyzer.exe`. If you are running from source
   instead, use:
   ```
   python C:\FBCRS\FBCRS_Record_Inventory_V8_5.py
   ```

3. **Click "Browse Folder"** and choose the folder to scan. Subfolders are
   included automatically.

4. **Click "Start Analysis."** The status line reports progress as
   `Scanning files: 1,240 / 8,900`.

5. **Wait.** The report opens in Excel automatically when finished, and a summary
   dashboard shows elapsed time, files found, and files classified.

### First run advice

**Start with 30–50 files you already know the answers for.** This tells you
whether the rulebook is tuned correctly for your content before you commit to a
multi-hour scan. Check the **Confidence Reason** column on those known files — if
it says the match came from generic terms, the rulebook needs work, not the
folder.

---

## 5. Reading the report — all 21 columns

The report has one sheet, `Inventory`, with the top row frozen and filters
enabled. A hidden sheet named `Lists` holds the dropdown values — leave it alone
unless you're deliberately editing the vocabularies.

### Identification

| Col | Header | Contents |
|---|---|---|
| **A** | File Name | Filename with extension |
| **B** | Extension | Lowercase extension, e.g. `.docx` |
| **C** | Content Category | `Document`, `Spreadsheet`, `PDF`, or `Other` |
| **D** | Date Created | From the filesystem |
| **E** | Date Modified | From the filesystem — **drives the expiry calculation** |
| **O** | Full File Path | Complete path, clickable hyperlink |

> **Note on Date Created:** Windows resets creation dates when files are copied or
> moved. A 2011 record copied to a new share in 2024 will show as created in 2024.
> Treat this column as unreliable for anything legal.

### Classification

| Col | Header | Contents |
|---|---|---|
| **F** | Code | Records series code, `UNCLASSIFIED`, or `N/A` for skipped files |
| **G** | Series Title | Series name from the rulebook |
| **Q** | Confidence | `Very High` / `High` / `Medium` / `Low` — **editable dropdown** |
| **R** | Confidence Reason | Plain-English explanation of the confidence level |
| **S–U** | Alternate Code 1–3 | Runner-up codes, when genuinely competitive |

### Retention and disposition

| Col | Header | Contents |
|---|---|---|
| **H** | Retention Period | Raw retention text from the rulebook, e.g. `CCY + 2 years` |
| **I** | Retention Expiry Date | Calculated date, or blank if it couldn't be determined |
| **J** | Disposition | Raw disposition text from the rulebook |
| **K** | Recommended Action | **Editable dropdown** — the decision to act on |
| **L** | Reason | **Editable dropdown** — why that action was recommended |

### Review flags

| Col | Header | Contents |
|---|---|---|
| **M** | Duplicate Detected | `YES` if an identical file appeared earlier in the scan |
| **N** | SME Review Required | `YES` for low-confidence or unclassified files |
| **P** | Comments | Error details, skip reasons, and other notes |

### The three dropdown columns

Columns **K**, **L**, and **Q** are constrained dropdowns. Typing anything not on
the list is rejected. Validation extends 500 rows past your data, so pasting in
more rows keeps the dropdowns working.

**Column K — Recommended Action**

| Value | Meaning |
|---|---|
| `DESTROY` | Retention has lapsed, or the file is an exact duplicate or empty |
| `DO NOTHING` | Still within its retention period — leave it alone |
| `MANUAL REVIEW REQUIRED` | The engine could not decide; a person must |
| `TRANSFER TO ARCHIVES` | The rulebook disposition calls for archival transfer |

**Column L — Reason**

| Value | Meaning |
|---|---|
| `WITHIN RETENTION` | Expiry date is in the future |
| `PASSED RETENTION` | Expiry date has passed |
| `DUPLICATE` | Byte-identical to a file seen earlier in this scan |
| `EMPTY/DRAFT` | Zero bytes, or a temporary Office lock file (`~$…`) |
| `UNSUPPORTED FORMAT` | File type that cannot contain records (image, video, archive) |
| `MISSING DISPOSITION RULE` | Retention lapsed but the rulebook has no disposition for that code |
| `NEEDS REVIEW` | Not confidently classified, or unreadable |

**Column Q — Confidence**

`Very High`, `High`, `Medium`, `Low`. Editable so a reviewer can record their own
assessment after inspecting a file.

### Order of precedence

When several rules could apply, the first match wins:

1. **Duplicate** → `DESTROY` / `DUPLICATE`
2. **Unclassified** → `MANUAL REVIEW REQUIRED` / `NEEDS REVIEW`
3. **Expired, disposition known** → the rulebook's disposition / `PASSED RETENTION`
4. **Expired, no disposition** → `MANUAL REVIEW REQUIRED` / `MISSING DISPOSITION RULE`
5. **Low confidence** → `MANUAL REVIEW REQUIRED` / `NEEDS REVIEW`
6. **Otherwise** → `DO NOTHING` / `WITHIN RETENTION`

> **Important:** duplicate detection sits at the top. A file that is both
> unclassified *and* a duplicate is marked `DESTROY`, not sent for review. If you
> want unclassified duplicates reviewed instead, filter on
> `Reason = DUPLICATE` **and** `Code = UNCLASSIFIED` before acting.

---

## 6. How classification works

### 6.1 Scoring

Each file's filename and content are matched against the rulebook's keywords and
phrases. Points accumulate per records series:

| Signal | Points |
|---|---|
| Keyword match | 2 |
| Phrase match (multi-word) | 10 |
| Match in the **filename** | × 3 |

Filenames get triple weight because a filename is a deliberate human label,
whereas body text is full of boilerplate.

**Rarity weighting.** Each term's points are divided by the number of series that
claim it. A keyword unique to one series contributes its full 2 points; a keyword
shared by 20 series contributes 0.1. This is essential — auto-generated rulebooks
tend to contain boilerplate terms attached to dozens of codes, which would
otherwise swamp the real signal.

**Terms shared by more than 20 series are ignored entirely** as having no
discriminating value.

**Each distinct term counts once per field.** A document repeating "financial"
400 times scores no higher than one mentioning it once.

### 6.2 Confidence bands

| Score | Confidence | Typically means |
|---|---|---|
| 30 or more | Very High | An exact phrase matched in the filename |
| 12–29 | High | Several distinctive terms matched |
| 4–11 | Medium | Moderate evidence |
| below 4 | Low | Only weak or generic terms matched |

Confidence is downgraded to `Medium` when two series score within 15% of each
other — a near-tie is not high confidence regardless of the raw score.

### 6.3 When the engine refuses to guess

A file becomes `UNCLASSIFIED` if:

- **nothing matched** → *"No rulebook keywords or phrases found — context is not clear"*
- **score below 2** → *"Only generic terms matched, shared across many series"*
- **score below 6 with no filename support** → *"Weak match in document body only, nothing in the file name"*

**Expect a meaningful number of `UNCLASSIFIED` results, and treat that as the tool
working correctly.** Refusing to guess is the point. A version that assigns a code
to everything is not more accurate — it just hides its uncertainty, which is far
more dangerous when the output drives destruction decisions.

### 6.4 Alternate codes (S–U)

Runner-up codes appear only when genuinely competitive: within 40% of the winning
score, and scoring at least 2 in their own right.

This means **high-confidence rows usually show no alternates at all**, which is
intentional. Listing second, third, and fourth place unconditionally would hand
reviewers three plausible-looking wrong answers on every clear match, and people
do pick them.

Alternates are always blank for `UNCLASSIFIED` files — if the top match is noise,
so are the runners-up.

---

## 7. Retention and expiry dates

### 7.1 Supported formats

The engine reads two patterns from the rulebook's Retention column:

| Pattern | Meaning | Expiry |
|---|---|---|
| `CCY + N` | Current Calendar Year plus N years | 31 December, year + N |
| `CFY + N` | Current Fiscal Year plus N years | 31 March, fiscal year end + N |

Fiscal year is treated as April–March: files modified January to March fall in
the fiscal year ending that same calendar year; April onwards fall in the fiscal
year ending the next.

Surrounding text is ignored, so `CCY + 2 years`, `CCY+2`, and `ccy + 2` all parse
identically.

### 7.2 Important limitations

**Expiry is calculated from Date Modified.** Records retention normally runs from
file *closure* — case closed, project completed, employee departed — which the
filesystem does not know. A file edited last week shows a retention clock that
started last week, regardless of the underlying record's actual age.

**Retention text the engine can't parse produces no expiry date.** Entries like
`Within a maximum of 1 month` or `Superseded + 1` are left blank, and the file is
reported as `WITHIN RETENTION` — indefinitely. Blank expiry dates are not a
guarantee of currency; they mean *unknown*. Filter on blank column I to find
these.

**Decimals are truncated.** `CCY+0.3` parses as `CCY+0`, giving an expiry of
31 December in the year the file was modified. Review any rulebook entries with
fractional retention.

---

## 8. Duplicate detection

### 8.1 How it works

Duplicates are found in two stages. First, file **sizes** are compared — identical
files must have identical sizes, so this eliminates almost everything without
reading any file content. Only files whose size collides with another are then
**hashed** (SHA-256 over the first 5 MB).

This is why scans are fast: on a typical folder, only a small percentage of files
are ever opened for hashing.

### 8.2 Which copy gets flagged

Files are processed in sorted path order, and **the first copy encountered is
kept**; later copies are marked `DUPLICATE` / `DESTROY`. This is deterministic —
re-scanning the same folder flags the same copies every time.

> **"First in sorted order" is alphabetical, not a judgement about which copy is
> the real record.** `Archive\report.docx` sorts before `Current\report.docx`, so
> the archived copy would be kept and the current one marked for destruction.
> Always confirm which copy you actually want before acting on a duplicate.

### 8.3 What is not detected

- **Near-duplicates.** One changed word makes two files unrelated to this tool.
- **Files differing beyond the first 5 MB.** Two large files with identical first
  5 MB and identical total size are reported as duplicates. Rare, but possible
  with padded or templated files.
- **Duplicates among skipped types.** Images, videos, and archives are never
  hashed, so identical copies of them are not flagged.
- **Duplicates across separate scans.** Detection is per-scan only.

---

## 9. The review workflow

A practical order for working through a report.

**Step 1 — Fix the plumbing first.** Filter column P (Comments) for
`ACCESS DENIED` and `Scan failed`. These files were never actually examined.
Resolve permissions and re-scan before drawing conclusions from the totals.

**Step 2 — Clear the safe wins.** Filter column L for `EMPTY/DRAFT`. Zero-byte
files and `~$` Office lock files are genuinely disposable and need little scrutiny.

**Step 3 — Review every `DESTROY` recommendation.** Filter column K for
`DESTROY`. Sort by column M so duplicates group together. For each, confirm the
copy being kept is the one you want. This is where irreversible mistakes happen.

**Step 4 — Work the review queue by confidence.** Filter column N for `YES`, then
sort column Q ascending so `Low` appears first. Use columns R, S, T, and U —
Confidence Reason tells you *why* it's uncertain, and the alternates give you
candidate codes to pick from instead of searching the full series list.

**Step 5 — Find the silent gaps.** Filter for blank column I (no expiry date) and
for `MISSING DISPOSITION RULE`. These are rulebook problems masquerading as file
problems, and they will recur on every future scan until fixed in the rulebook.

**Step 6 — Record your decisions in the dropdowns.** Override columns K, L, and Q
as you review. The constrained vocabularies keep the report analysable afterwards
— you can pivot it, count it, and hand it on.

> **Keep the original report unedited.** Save your reviewed version under a new
> name. Reports are timestamped, so a re-scan won't overwrite anything, but an
> unedited baseline is worth having.

---

## 10. Maintaining the rulebook

### 10.1 Structure

`FBCRS_Master_Full.xlsx` needs four sheets, each with a header row:

| Sheet | Column A | Column B | Columns C–D |
|---|---|---|---|
| **Records** | Code | Series Title | Retention, Disposition |
| **Keywords** | Code | Keyword (one per row) | — |
| **Phrases** | Code | Phrase (one per row) | — |
| **Synonyms** | Term | Replacement | — |

Additional sheets are ignored, so a `FunctionCodes` reference sheet or similar can
live in the same workbook safely.

### 10.2 Editing

1. Close the app.
2. Edit the workbook and save it.
3. **Save and close the workbook.**
4. Restart the app — the rulebook is read once at startup.

### 10.3 Writing effective keywords

**Specific beats general.** `reconciliation` identifies a series;
`report` identifies nothing. Rarity weighting means generic terms contribute
almost nothing anyway, so they're wasted effort.

**Phrases outperform keywords** — 10 points versus 2, and far fewer false
positives. `key performance indicator` is worth more than `key` plus
`performance` plus `indicator` separately.

**Terms are matched as whole words.** `plan` will not match `plans` or
`planning`. Add each form you need. Conversely this is protective: `bi` matches
only the standalone word, never the inside of `liability`.

**Never put multi-line content in a Synonyms cell.** Each synonym replacement
must be a short single-line phrase. A pasted block in that column expands into
every document the engine reads and destroys all classification accuracy — see
section 12.

### 10.4 Auditing the rulebook

`repair_rulebook.exe` checks a rulebook and reports problems.

**Double-click it.** A file dialog opens — select the rulebook you want checked.
The findings appear in a console window, which stays open until you press Enter.

You can also pass the path directly if you prefer:

```
repair_rulebook.exe "C:\FBCRS\FBCRS_Master_Full.xlsx"
```

### What it reports

- **Corrupt Synonyms entries** — multi-line or overlong replacement values, the
  fault that makes every file classify as the same code
- **Over-shared keywords** — terms claimed by so many series that they carry no
  signal, listed worst first
- **Over-shared phrases** — the same problem in the Phrases sheet
- **How many keywords are genuinely distinctive** — owned by exactly one series.
  This is the single best measure of rulebook health

### What it writes

Two files, beside the rulebook you selected:

| File | Contents |
|---|---|
| `<name>_REPAIRED.xlsx` | The cleaned rulebook — **use this one** |
| `<name>_AUDIT.txt` | The full report, saved so you can keep or circulate it |

**The original is never modified.** Records, Keywords, and Phrases are copied
through untouched; only the Synonyms sheet is repaired, and anything removed from
it is moved to a `FunctionCodes` sheet rather than deleted, so nothing is lost.

To put a repaired rulebook into service, rename it to `FBCRS_Master_Full.xlsx`
and place it beside `FBCRS_Analyzer.exe`.

Run this whenever you make significant rulebook edits, and any time
classification quality drops.

---

## 11. Performance tuning

Expect roughly **30–40 ms per file** on local storage. Network shares are slower
and vary with the link.

If scans are too slow, adjust these constants near the top of the script, in this
order of impact:

```python
XLSX_ROWS = 500     # rows read per spreadsheet sheet
TEXT_CAP  = 100000  # characters of content fed to the matcher
PDF_PAGES = 10      # pages read per PDF
```

**`XLSX_ROWS` is the biggest lever.** Parsing spreadsheets dominates runtime.
Dropping it to `150` roughly halves total scan time and rarely changes results —
if a spreadsheet's subject isn't evident in its first 150 rows and its filename,
another 350 rows seldom help.

Other settings:

| Constant | Default | Effect |
|---|---|---|
| `MAX_WORKERS` | `min(8, cores × 2)` | Parallel threads. Raising helps on high-latency shares |
| `MAX_SIZE_BYTES` | 15 MB | Files above this are classified on filename only |
| `HASH_BYTES` | 5 MB | Bytes hashed for duplicate comparison |
| `DOCX_PARAS` | 200 | Paragraphs read per Word file |
| `PPTX_SLIDES` | 10 | Slides read per presentation |

### General advice

- **Run from a local drive.** The single biggest factor.
- **Break very large directories into sub-folders** and scan separately. Reports
  are named after the folder, so they won't collide.
- **Scan outside business hours** for large network folders.

---

## 12. Troubleshooting

### "Fatal Error: Missing Rulebook"

`FBCRS_Master_Full.xlsx` is not in the same folder as the program. The dialog
shows the exact folder being searched. Check for a renamed file
(`FBCRS_Master_Full_REPAIRED.xlsx` won't be found) or a hidden double extension
(`.xlsx.xlsx`).

### "Fatal Error: Rulebook Read Error"

Usually the rulebook is open in Excel. Close it. If it persists, the file may be
corrupt — open it in Excel and re-save.

### "Access is denied" partway through a scan

You're running the program from a network share. Move it to `C:\FBCRS\`. See
section 3.2.

### "The output Excel file is open"

A previous report is open in Excel. Close it. The program retries three times
before giving up.

### Every file gets the same code

A corrupt **Synonyms** sheet. If a synonym's replacement value contains multiple
lines or a long block of text, that block gets injected into every document the
engine reads, so every file matches the same series.

Run `repair_rulebook.exe` (section 10.4) and use the repaired copy. Verify by
checking the **Confidence Reason** column — if unrelated files all claim to match
the same phrase, this is the cause.

### Nearly everything is UNCLASSIFIED

Check **Confidence Reason** to distinguish two very different situations:

- *"No rulebook keywords or phrases found"* → your rulebook lacks terms for this
  content. Add keywords for the series that actually appear in this folder.
- *"Only generic terms matched"* → the rulebook's terms are too generic. Add
  distinctive phrases.

If review volume is genuinely unmanageable, `MIN_SCORE` and `BODY_ONLY_MIN`
control the thresholds — but read the evidence before loosening them. Lowering
thresholds doesn't create accuracy, it just converts visible uncertainty into
invisible uncertainty.

### Files are missing from the report

Files and folders whose names begin with `~` or `.` are skipped by design (Office
temp files, hidden folders). Also check column P for `ACCESS DENIED`.

### Where to look when nothing else explains it

```
Documents\FBCRS_Reports\FBCRS_error_log.txt
```

Full technical tracebacks for any failure, with timestamps. Also check for
leftover `*_temp.csv` files in that folder — if a scan was interrupted, the
partial results survive there and can be opened in Excel.

---

## 13. Building and sharing an .exe

### 13.1 Building

```
pip install pyinstaller
cd /d C:\FBCRS
```

**The main application** — `--windowed` suppresses the console, since it has its
own interface:

```
pyinstaller --onefile --windowed --name FBCRS_Analyzer FBCRS_Record_Inventory_V8_5.py
```

**The rulebook checker** — `--console` here, *not* `--windowed`, because its
output is the console window. Build it windowed and the findings go nowhere:

```
pyinstaller --onefile --console --name repair_rulebook repair_rulebook.py
```

Both results appear in `C:\FBCRS\dist\`.

If `FBCRS_Analyzer.exe` builds but crashes on launch with a `customtkinter` path
error, rebuild it with `--collect-all customtkinter` added.

### 13.2 Packaging for colleagues

The rulebook is deliberately **not** bundled into the exe, so you can update
keywords without rebuilding. Ship a folder:

```
FBCRS_Analyzer\
├── FBCRS_Analyzer.exe          (the application)
├── FBCRS_Master_Full.xlsx      (the rulebook)
├── repair_rulebook.exe         (rulebook checker)
└── FBCRS_User_Manual.pdf       (this manual)
```

Zip it. Tell recipients to unzip it and keep the files together — the exe cannot
run without the rulebook beside it.

### 13.3 What to expect on their machines

**Reports save to their own Documents folder**, not yours. Nothing is hardcoded
to your user profile.

**Windows SmartScreen and antivirus will likely complain.** Unsigned PyInstaller
executables routinely trigger "Windows protected your PC" warnings and are
sometimes quarantined outright. This is a known false positive, but recipients
will need to click through it — and on a locked-down corporate machine it may be
blocked entirely. Warn people in advance, and check with IT before wide
distribution.

**Test on a machine without Python installed** before sending it widely. An exe
that works on the build machine can still be missing components that only surface
on a clean one.

---

## 14. Known limitations

**Classification is keyword matching, not comprehension.** No understanding of
context, business purpose, or the difference between an official record and a
working copy.

**Retention runs from Date Modified, not file closure.** See section 7.2.

**Unparseable retention text produces no expiry date and a `WITHIN RETENTION`
verdict.** Blank expiry means *unknown*, not *current*.

**Duplicate "keeper" selection is alphabetical.** Not a judgement about which copy
is authoritative. See section 8.2.

**Only the beginning of each file is read** — 10 PDF pages, 200 Word paragraphs,
500 spreadsheet rows per sheet, 100,000 characters total. Content past those
limits is invisible. Configurable, at a cost in speed.

**Scanned documents and images are not OCR'd.** A scanned PDF with no text layer
yields no content, and classifies on filename alone.

**Files over 15 MB are classified on filename only.**

**Encrypted and password-protected files cannot be read.** They appear with
`Scan failed` in Comments.

**Very long paths may fail on Windows.** Paths beyond 260 characters can produce
`ACCESS DENIED` unless long path support is enabled. Excel also refuses
hyperlinks over 255 characters, so those cells are left as plain text.

**Duplicate detection is per-scan.** No memory between runs.

**Date Created is unreliable** after files are copied or moved.

---

## Appendix A — configuration reference

All constants sit near the top of the script.

### Scoring

| Constant | Default | Purpose |
|---|---|---|
| `KW_BASE` | 2.0 | Points per keyword match |
| `PH_BASE` | 10.0 | Points per phrase match |
| `FN_BOOST` | 3.0 | Multiplier for filename matches |
| `MAX_OWNERS` | 20 | Terms claimed by more series than this are ignored |
| `MIN_SCORE` | 2.0 | Below this → `UNCLASSIFIED` |
| `BODY_ONLY_MIN` | 6.0 | Threshold when nothing matched in the filename |
| `CONF_BANDS` | 30 / 12 / 4 | Very High / High / Medium cutoffs |

### Alternates

| Constant | Default | Purpose |
|---|---|---|
| `N_ALTERNATES` | 3 | Number of runner-up columns |
| `ALT_MIN_SCORE` | 2.0 | Minimum score for a runner-up to be shown |
| `ALT_REL_MIN` | 0.40 | Runner-up must reach this fraction of the winner. Set to `0.0` to always show three |

### Reading and performance

| Constant | Default | Purpose |
|---|---|---|
| `MAX_WORKERS` | `min(8, cores × 2)` | Parallel worker threads |
| `TEXT_CAP` | 100000 | Max characters fed to the matcher |
| `XLSX_ROWS` | 500 | Rows read per sheet — **largest speed lever** |
| `XLSX_SHEETS` | 3 | Sheets read per workbook |
| `PDF_PAGES` | 10 | Pages read per PDF |
| `DOCX_PARAS` | 200 | Paragraphs read per Word file |
| `PPTX_SLIDES` | 10 | Slides read per presentation |
| `HASH_BYTES` | 5 MB | Bytes hashed for duplicate comparison |
| `MAX_SIZE_BYTES` | 15 MB | Above this, filename-only classification |

### File handling

| Constant | Contents |
|---|---|
| `SKIP_EXTS` | `.jpg .jpeg .png .gif .bmp .tiff .mp4 .mov .avi .wmv .mkv .zip .7z .rar .exe .dll .msi` |
| `ACTION_OPTIONS` | The four Recommended Action values |
| `REASON_OPTIONS` | The seven Reason values |
| `CONF_OPTIONS` | The four Confidence values |

Adding a value to `ACTION_OPTIONS`, `REASON_OPTIONS`, or `CONF_OPTIONS`
automatically extends the corresponding Excel dropdown.

> **Note:** `EXTS` exists in the code but is not used for filtering. Every file
> not in `SKIP_EXTS` is inventoried; content extraction simply only works for the
> supported types.

---

## Appendix B — file inventory

| File | Purpose |
|---|---|
| `FBCRS_Analyzer.exe` | The application |
| `FBCRS_Master_Full.xlsx` | Classification rulebook — must sit beside the app |
| `repair_rulebook.exe` | Rulebook checker and Synonyms repair tool |

Source files, if you are running or rebuilding from Python rather than the exes:

| File | Builds |
|---|---|
| `FBCRS_Record_Inventory_V8_5.py` | `FBCRS_Analyzer.exe` |
| `repair_rulebook.py` | `repair_rulebook.exe` |

### Generated output

| Location | Contents |
|---|---|
| `Documents\FBCRS_Reports\<folder> summary <timestamp>.xlsx` | The report |
| `<rulebook>_REPAIRED.xlsx` | Cleaned rulebook from `repair_rulebook.exe` |
| `<rulebook>_AUDIT.txt` | Rulebook audit report |
| `Documents\FBCRS_Reports\<folder>_<timestamp>_temp.csv` | Checkpoint; deleted on success, survives a crash |
| `Documents\FBCRS_Reports\FBCRS_error_log.txt` | Technical error log, appended to |

---

*Version 8.5. This manual describes the tool's behaviour as built; it is not a
records management policy document. Disposition decisions remain the
responsibility of the reviewing records professional.*
