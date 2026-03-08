# LINE Report Generator - Specification

## 1. Project Overview

This is a Python program that generates a markdown report (`.md`) from LINE chat export files and image files. Designed to run on Windows as the primary platform.

**Purpose:** Convert LINE Keep/Chat export (chatlog + images) into a structured markdown report.

---

## 2. System Input

| Parameter | Description | Example |
|-----------|-------------|---------|
| Report Name | Output filename (without extension) | `20260308` |
| Folder | Path to folder containing chatlog and images | `./20260308` or `C:\data\20260308` |
| Sort Image | Image sorting method: `d` (date) or `n` (name) | `d` or `n` |

---

## 3. Functionality

### 3.1 Image Sorting
- **`d` (date):** Sort images by creation date, ascending (oldest → newest)
- **`n` (name):** Sort images by filename (IMG_001 → IMG_999), ascending

### 3.2 Chatlog File Detection
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
- `[Photo]` marker indicates image

### 3.3 Report Generation
- Output file: `{Report Name}.md` (e.g., `20260308.md`)
- Location: Same folder as input folder

### 3.4 Report Structure

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

### 3.5 Image Ordering for Multiple Photos
When an entry has multiple images, the images are **interleaved** (สลับเรียง):

```
Input order: IMG_001, IMG_002, IMG_003, IMG_004, IMG_005, IMG_006
Output order: IMG_001, IMG_003, IMG_005, IMG_002, IMG_004, IMG_006
```

This creates a zigzag pattern: odd numbers first, then even numbers.

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

---

## 5. Example Usage

```bash
# Sort by date (d)
python line_report.py 20260308 ./20260308 d

# Sort by name (n)
python line_report.py 20260308 C:\data\20260308 n
```

---

## 6. Expected Output

- `{Report Name}.md` file in the input folder
- Console output showing processing status
