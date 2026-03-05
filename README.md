# Excel Format Adjustment Guide

This guide explains what needs to be changed in the script to work with the provided Excel file (`work1.xlsx`).

---

## File Structure Overview

In `work1.xlsx`:

- The actual data starts from **Excel row 5**
- Column A contains **URLs**
- Column B is where **emails should be written**
- The first 4 rows contain headers or unrelated content and must be skipped

Example structure:

| Column A (Category) | Column B (Source) |
|---------------------|-------------------|
| https://example.com | email@example.com |

---

## Required Script Changes

### 1. Change START_ROW

Since the real data starts at Excel row 5:

```python
START_ROW = 5
```

This prevents the script from trying to process invalid header rows.

---

### 2. Keep `header=None`

Do NOT remove this:

```python
df = pd.read_excel(INPUT_FILE, header=None)
```

Reason:
The file does not contain a clean structured header row. Using `header=None` ensures pandas treats all rows as raw data.

---

### 3. Keep URL Column as Column A

Since URLs are in Column A (index 0):

```python
url = str(df.iloc[i, 0]).strip()
```

No changes required here.

---

### 4. Keep Email Output as Column B

Since emails should be written to Column B (column index 2 in openpyxl):

```python
ws.cell(row=row + 1, column=2, value=email)
```

No changes required here.

---

## Final Required Modification

Only this line must be changed:

```python
START_ROW = 5
```

Everything else in the script can remain the same.

---

## Why This Change Is Necessary

The first 4 rows of the Excel file contain:

- Category titles
- Extra formatting rows
- Non-URL content

If the script starts at row 1, it will attempt to scrape invalid values and may fail or waste requests.

Setting `START_ROW = 5` ensures only valid URLs are processed.
