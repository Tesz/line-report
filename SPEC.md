# LINE Report Generator - Specification

## 1. Project Overview

This is a Python program that generates a report from LINE chat export files and image files. Designed to run on Windows as the primary platform.

**Purpose:** Convert LINE Keep/Chat export (chatlog + images) into a structured report.

**Supported Output Formats:**
- Markdown (`.md`) - Default
- Excel (`.xlsx`)

---

## 2. System Input

| Parameter | Description | Example |
|-----------|-------------|---------|
| Report Name | Output filename (without extension) | `20260308` |
| Folder | Path to folder containing chatlog and images | `./20260308` or `C:\data\20260308` |
| Sort Image | Image sorting method: `d` (date) or `n` (name) | `d` or `n` |
| Output Format | Output file format: `md` (markdown) or `xlsx` (Excel) | `md` or `xlsx` (default: `md`) |

---

## 3. Functionality

### 3.1 Image & Video Sorting
- **`d` (date):** Sort files by creation date, ascending (oldest → newest)
- **`n` (name):** Sort files by filename (IMG_001 → IMG_999), ascending

### 3.2 Media File Detection
The program supports both **images** and **videos**:

**Image formats:** `.jpg`, `.jpeg`, `.png`, `.gif`
**Video formats:** `.mp4`, `.mov`, `.avi`

### 3.3 Chatlog File Detection
The system searches for chatlog file in this order:
1. Look for `[LINE]Keep Memo.txt`
2. If not found, search for `*.txt` (any `.txt` file)
3. If exactly one `.txt` file is found, use it for processing
4. If no `.txt` files are found, **stop execution with error**

### 3.3 Chatlog Format
The chatlog exported from LINE Keep has this format:

```text
2026.03.08 วันอาทิตย์
10:56 Trin・ティン Pro. ยกแป้นพินมาผิดทำให้รถเลี้ยวผิด
ทางออก booth wax
ทำการถอยรถ และ manual สับราง
Reset หน้าตู้และ run c/v ok.
เวลา 21.40-21.42 = 2 นาที
10:56 Trin・ティン [Photo]
10:56 Trin・ティン [Photo]
14:21 Trin・ティン สภาพจาระบี Roller PP205
- Roller DFK มาร์คเขียว( 2 อาทิตย์) จาระบี molykote สภาพเริ่มเข้าไปในลูกปืน สภาพยังดีสีเริ่มเทาเข้ม
- Roller DFK มาร์คชมพู(1 อาทิตย์) จาระบี molykote สภาพยังดี ขาวเหลือง
- Roller Overhaul (1 อาทิตย์)
จาระบี Lumax สภาพเปลี่ยนจากสีขาวเป็นดำ แต่ยังไม่แห้งครับ
ต่อไปจะขอถอดเช็คจาระบีใน Roller ทุกๆ 1 เดือนก่อน เพื่อประเมินความเสี่ยงครับ
14:21 Trin・ティン [Photo]
14:21 Trin・ティン [Photo]
14:21 Trin・ティン [Photo]
14:21 Trin・ティン [Photo]
```

**Key patterns:**
- Date line: `YYYY.MM.DD วัน[วัน]`
- Timestamp: `HH:MM Name [Photo]` or `HH:MM Name Message...`
- Multi-line message: continues until next timestamp
- Media markers: `[Photo]`, `Photos`, `รูป`, `写真`, `照片`, `วิดีโอ`, `Video`

### 3.4 Entry Grouping (Important!)
Entries are grouped by **content continuity**, NOT by timestamp.

- If the next message has **NO message text** (only media), it will be grouped with the **previous entry**
- This allows multiple media-only messages (same job, different timestamps) to be grouped together
- Example: Job description at 14:21, followed by 4 photos at 14:21, 16:42, 16:43, 16:44 → All grouped as 1 entry

### 3.5 Media Marker Configuration
The program supports multiple language markers as an **array**:

```python
MEDIA_MARKERS = [
    "รูป",        # Thai
    "[Photo]",    # English
    "Photos",     # English
    "写真",       # Japanese
    "照片",       # Chinese
    "วิดีโอ",     # Thai video
    "Video",      # English video
]
```

### 3.6 Report Generation
- Output file: `{Report Name}.md` (e.g., `20260308.md`)
- Location: Same folder as input folder

### 3.7 Report Structure

``````markdown
# 20260308

## หัวข้อที่ 1
ข้อความทั้งหมดก่อนถึงรูปภาพ
````col
```col-md
![[ชื่อรูปที่1.jpg]]
```
```col-md
![[ชื่อรูปที่2.jpg]]
```
````

## หัวข้อที่ 2
ข้อความทั้งหมดก่อนถึงรูปภาพ
````col
```col-md
![[ชื่อรูปที่3.jpg]]
![[ชื่อรูปที่5.jpg]]
![[ชื่อรูปที่7.jpg]]
```
```col-md
![[ชื่อรูปที่4.jpg]]
![[ชื่อรูปที่6.jpg]]
![[ชื่อรูปที่8.jpg]]
```
````
``````

### 3.8 Media Ordering for Multiple Files
When an entry has multiple images, the images are **interleaved** (สลับเรียง):

```
Input order: IMG_001, IMG_002, IMG_003, IMG_004, IMG_005, IMG_006
Output order: IMG_001, IMG_003, IMG_005, IMG_002, IMG_004, IMG_006
```

This creates a zigzag pattern: odd numbers first, then even numbers.

### 3.9 Output Format Selection
The program supports two output formats:

- **`md` (Markdown):** Default format, generates `.md` file with image links
- **`xlsx` (Excel):** Generates `.xlsx` file with structured data table

**Default:** If not specified, output format defaults to `md`.

When `xlsx` is selected, the output will be `{Report Name}.xlsx` instead of `.md`.

### 3.10 Excel Output Structure (xlsx)
When output format is `xlsx`, the Excel file contains **2 sheets**:

#### Sheet 1: "Cover"
| Column | Description |
|--------|-------------|
| A: Entry No. | Sequential entry number (1, 2, 3...) |
| B: Date | Date from chatlog (e.g., "2026.03.08 วันอาทิตย์") |
| C: Time | Timestamp from chatlog (e.g., "10:56") |
| D: Sender | Person who sent the message (e.g., "Trin・ティン Pro.") |
| E: Text Content | Message text (all lines combined) |
| F: Image Count | Number of images in this entry |
| G: Image Filenames | Comma-separated list of image filenames (e.g., "IMG_001.jpg, IMG_002.jpg") |
| H: Image Path | Relative path to images folder (e.g., "./images/") |

**Cover Sheet Features:**
- Header row with bold formatting
- Auto-adjust column width
- Text wrapping enabled for long content
- Images are **not embedded** (only filenames listed)

#### Sheet 2: "Detail"
The Detail sheet contains the full report with images embedded:

**Structure per entry:**
```
[Entry No.] - [Date] [Time] - [Sender]
[Text Content]

[[Image 1]]
[[Image 2]]
[[Image 3]]
... (all images for this entry, stacked vertically)

[Spacing row for next entry]
[Entry No.2] - [Date] [Time] - [Sender]
...
```

**Detail Sheet Features:**
- **Image embedding:** All images are actually inserted into Excel cells
- **Image sizing:** 
  - Maximum width: 10cm
  - Height: Auto-adjusted based on original aspect ratio
  - Images stacked vertically within each entry
- **Spacing:** Empty rows between entries for separation
- **Text formatting:**
  - Entry header: Bold, font size 12
  - Text content: Normal, font size 10, text wrapping enabled
- **Flow:** Continue from entry 1 to entry N sequentially

---

## 4. Technical Considerations

### 4.1 Thai Language Handling
- Chatlog contains Thai characters
- Use UTF-8 encoding for reading and writing files
- Ensure proper Unicode handling throughout the program

### 4.2 Windows Compatibility
- Use `pathlib` or `os.path` for cross-platform path handling
- Support both forward slash (`/`) and backslash (`\`) in folder paths

### 4.3 Error Handling
- Display clear error messages in Thai/English when:
  - No `.txt` file found
  - Multiple `.txt` files found (when `[LINE]Keep Memo.txt` doesn't exist)
  - Image folder not found
  - Other file access errors

### 4.4 Excel Output Dependencies
- When output format is `xlsx`, requires `xlsxwriter` library
- Install via: `pip install xlsxwriter`
- If `xlsxwriter` is not installed and `xlsx` format is selected, display error message
- Note: xlsxwriter is more stable than openpyxl and has better support for various image formats

---

## 5. Example Usage

```bash
# Sort by date (d), output markdown (md)
python line_report.py 20260308 ./20260308 d md

# Sort by name (n), output markdown (md)
python line_report.py 20260308 C:\data\20260308 n md

# Sort by date (d), output Excel (xlsx)
python line_report.py 20260308 ./20260308 d xlsx

# Sort by name (n), output Excel (xlsx)
python line_report.py 20260308 C:\data\20260308 n xlsx

# Default: output markdown if format not specified
python line_report.py 20260308 ./20260308 d
```

---

## 6. Expected Output

- `{Report Name}.md` file (when format=`md`) in the input folder
- `{Report Name}.xlsx` file (when format=`xlsx`) in the input folder
- Console output showing processing status and selected format
