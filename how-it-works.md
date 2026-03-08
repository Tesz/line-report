# How It Works - LINE Report Generator

## Overview

This program converts LINE chat export (chatlog + images) into a structured markdown report.

---

## 1. File Search

### Chatlog Detection

The program searches for chatlog file in this order:

1. **First priority:** `[LINE]Keep Memo.txt`
2. **Second priority:** If not found, search for any `*.txt` file
3. **Error:** If no `.txt` files found → stop with error
4. **Error:** If multiple `.txt` files found (without `[LINE]Keep Memo.txt`) → stop with error

### Image Detection

The program lists **all image files** in the folder:
- Supported formats: `.jpg`, `.jpeg`, `.png`, `.gif`
- No filename pattern required (works with any naming)
- Examples: `photo1.jpg`, `IMG_001.png`, `image123.jpeg`

---

## 2. Image Sorting

### Sort Options

| Option | Description | Order |
|--------|-------------|-------|
| `n` | Sort by filename | Alphabetical (ASC) |
| `d` | Sort by creation date | Oldest → Newest (ASC) |

### Example

```
# Filename sort (n)
photo1.jpg, photo2.jpg, photo3.jpg, ...

# Date sort (d)
photo_modified_earliest.jpg, ..., photo_modified_latest.jpg
```

---

## 3. Parse Chatlog

### Message Format Detection

The program uses regex to parse each line:

```
HH:MM Name Message
```

**Regex pattern:** `^(\d{2}:\d{2})\s+(.+?)\s+(.+)$`

| Group | Captures | Example |
|-------|----------|---------|
| Group 1 | Timestamp | `10:56` |
| Group 2 | Sender name | `Trin・ティン` |
| Group 3 | Message | `ข้อความ...` |

### Photo Detection

The program detects photos by `PHOTO_MARKER`:
- Thai: `รูป`
- English: `[photo]`

Two formats supported:

1. **Same line as timestamp:**
   ```
   10:56 Trin・ティン รูป
   ```

2. **Separate line:**
   ```
   10:56 Trin・ティン ข้อความ
   รูป
   ```

---

## 4. Image Assignment

### Entry Grouping

Each entry contains:
- **Message:** Text before first photo
- **Photo count:** Number of photos for this entry

### Image Ordering (Per Entry)

When an entry has multiple images, they are **interleaved**:

```
Input:  IMG_001, IMG_002, IMG_003, IMG_004
Output: IMG_001, IMG_003 (column 1)
        IMG_002, IMG_004 (column 2)
```

**Pattern:** Odd images first → Even images

---

## 5. Report Structure

### Output Format

```markdown
# ReportName

## หัวข้อที่ 1
ข้อความก่อนรูปภาพ
````col
```col-md
photo1.jpg
photo3.jpg
```
```col-md
photo2.jpg
photo4.jpg
```
````

## หัวข้อที่ 2
...
```

---

## 6. Configuration

### PHOTO_MARKER

Located at the top of `line_report.py`:

```python
# === CONFIGURATION ===
# Change this to match your LINE language
# Thai: รูป
# English: [photo]
PHOTO_MARKER = "รูป"
# ====================
```

---

## 7. Usage

```bash
python line_report.py <Report Name> <Folder> <Sort>

# Example
python line_report.py 20260308 ./20260308 n
python line_report.py 20260308 C:\data\20260308 d
```

### Parameters

| Parameter | Description | Example |
|-----------|-------------|---------|
| Report Name | Output filename (without .md) | `20260308` |
| Folder | Path to folder with chatlog + images | `./20260308` |
| Sort | `n` (name) or `d` (date) | `n` or `d` |

---

## 8. Program Flow Diagram

```
START
  │
  ├─► Find chatlog file
  │     ├─ [LINE]Keep Memo.txt → USE IT
  │     └─ *.txt (1 file) → USE IT
  │         └─ 0 or >1 files → ERROR
  │
  ├─► Get all images in folder
  │     └─ Sort by name or date (ASC)
  │
  ├─► Parse chatlog
  │     ├─ Extract: timestamp, sender, message
  │     └─ Count photos (รูป or [photo])
  │
  ├─► Build entries
  │     ├─ Message + photo count per entry
  │     └─ Interleave images (odd → even)
  │
  └─► Generate markdown report
        └─ Save as {ReportName}.md
END
```
