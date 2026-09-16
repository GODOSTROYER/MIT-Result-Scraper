# MIT Result Scraper

A simple Python program that uses **Selenium** and **Pandas** to scrape SGPA and CGPA results from the MIT student portal (`student.mitapps.in`) for a batch of accounts and save them to an Excel file.

## What it does

Given a spreadsheet of student logins, the script:

1. Opens Chrome and visits the portal's landing page.
2. Logs in with each username/password in turn via the **Student Login** form.
3. Opens the result page and reads the **SGPA** (per semester) and **CGPA** values.
4. Collects everything into a table and writes it out as an Excel workbook.

If a login fails or times out, that row is still recorded with empty results, and the script moves on to the next account.

## How it works

- **`init_driver()`** — starts a Chrome WebDriver. It expects `chromedriver.exe` at `C:\chromedriver-win64\chromedriver.exe`. The browser runs visibly by default; a commented `--headless` line can be enabled for headless runs.
- **`login(driver, username, password)`** — navigates to `https://student.mitapps.in/landing`, clicks *Student Login*, submits the credentials, and waits for the results link to appear.
- **`scrape_results(driver)`** — opens `https://student.mitapps.in/student_result` and extracts the SGPA and CGPA values from the page.
- **`main(input_excel, output_excel)`** — reads the input workbook (no header row), loops over every row, and writes a `DataFrame` with columns `Username`, `Password`, `SGPA`, `CGPA` to the output workbook.

## Requirements

- Python 3
- [Google Chrome](https://www.google.com/chrome/) and a matching [ChromeDriver](https://developer.chrome.com/docs/chromedriver/)
- Python packages: `selenium`, `pandas`, `openpyxl`

```bash
pip install selenium pandas openpyxl
```

## Input / output

- **Input:** an `.xlsx` file with **no header row** — column 1 is the username, column 2 is the password (one account per row).
- **Output:** an `.xlsx` file with the columns `Username`, `Password`, `SGPA`, `CGPA`.

Both paths are set near the bottom of the script (in the `__main__` block) and default to `D:\MIT Result Scraper\Blank.xlsx` (input) and `D:\MIT Result Scraper\cute.xlsx` (output).

## Running it

1. Install the requirements above and place `chromedriver.exe` where the script expects it (or edit the path in `init_driver()`).
2. Prepare your input `.xlsx` and update `input_excel` / `output_excel` in the script to your own paths.
3. Run:

```bash
python "scraper-nonheadless v1.py"
```

## Project structure

```
MIT-Result-Scraper/
└── scraper-nonheadless v1.py   # the entire scraper (login + scrape + export)
```

## Responsible use

This tool automates logging into a student portal with real credentials, so please use it responsibly:

- Only use it with accounts you own or have explicit permission to access.
- The input and output files store usernames and **passwords in plain text** — keep them private and delete them when you're done.
- Respect the portal's terms of service, and don't hammer it with large batches or rapid repeated runs.

This is a personal utility, provided as-is with no warranty.

## Author

**Arnav Bule** — [arnavbule.in](https://www.arnavbule.in) · [github.com/GODOSTROYER](https://github.com/GODOSTROYER)
