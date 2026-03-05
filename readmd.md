Async Website Email Scraper

A high-speed asynchronous Python crawler that extracts email addresses from websites listed in an Excel file and writes the results back to a new Excel file.

The script checks:

Homepage

Contact pages

About pages

Support pages

Internal website pages

It uses async requests + intelligent email selection to scrape emails quickly.

 Features

Asynchronous scraping (aiohttp + asyncio)

 Reads websites from Excel

 Extracts emails from:

visible text

mailto links

JavaScript

Base64 encoded strings

 Crawls contact/about pages automatically

 Crawls internal pages if needed
 Picks the best email automatically

 Saves results back to Excel

 Requirements

Install dependencies:

pip install aiohttp pandas beautifulsoup4 openpyxl

Python version recommended:

Python 3.9+
📁 Project Structure
email-scraper/
│
├── scraper.py
├── work sheet 2.xlsx
├── input_updated.xlsx
└── README.md
Excel Input Format
The script reads websites from column A.

Example:

A (Website)
https://example.com

https://company.com

https://startup.io
 Excel Output

The script writes the email into Column B.

Website	Email
https://example.com
	info@example.com

https://company.com
	support@company.com
▶️ Running the Script

Run:

python scraper.py

After completion you will see:

Done! Saved to input_updated.xlsx
⚙️ Configuration

Inside the script:

INPUT_FILE = "work sheet 2.xlsx"
OUTPUT_FILE = "input_updated.xlsx"
START_ROW = 1
INPUT_FILE

Excel file containing website URLs.

OUTPUT_FILE

File where results will be saved.

START_ROW

Row where scraping begins.

Example:

START_ROW = 5

This skips the first 4 rows.

 Using Different Excel Layouts

Different Excel files may store websites in different columns.

You must change this line:

url = str(df.iloc[i, 0]).strip()
Column Index Guide
Excel Column	Index
A	0
B	1
C	2
D	3
E	4
Example 1

Websites in Column C

A	B	C
Name	Industry	Website

Change:

url = str(df.iloc[i, 2]).strip()
Example 2

Websites in Column D

Change:

url = str(df.iloc[i, 3]).strip()
📌 Changing Email Output Column

Currently the script writes email to Column B:

ws.cell(row=row + 1, column=2, value=email)
Column index reference
Excel Column	Number
A	1
B	2
C	3
D	4
Example

Save email to Column E

ws.cell(row=row + 1, column=5, value=email)
 Performance Settings

You can tune scraping speed.

CONCURRENT_REQUESTS = 50
MAX_PAGES_CRAWL = 5
Recommended
Setting	Value
CONCURRENT_REQUESTS	50–100
MAX_PAGES_CRAWL	3–5

Higher values = faster scraping but more server load.

Pages Automatically Checked

The script checks common contact pages:

/contact
/contact-us
/about
/about-us
/support
/help
/team
/careers
Email Selection Logic

If multiple emails are found, the script prioritizes:

info@
support@
contact@
hello@
sales@
business@

It ignores:

noreply@
no-reply@
donotreply@
🕷 Crawling Behavior
