# LINE Report Generator

Python program to generate markdown report from LINE chat export files and images.

## Requirements

- Python 3.x

## Configuration

Edit `PHOTO_MARKER` in `line_report.py` to match your LINE language:

```python
# Change this to match your LINE language
# Thai: รูป
# English: [photo] or Photos
PHOTO_MARKER = "Photos"
JOB_TITLE = "งานที่"
```

## Usage

```bash
python line_report.py <Report Name> <Folder> <Sort>
```

### Parameters

| Parameter | Description | Example |
|-----------|-------------|---------|
| Report Name | Output filename (without extension) | `20260308` |
| Folder | Path to folder containing chatlog and images | `./20260308` or `C:\data\20260308` |
| Sort | Image sorting: `d` (date) or `n` (name) | `d` or `n` |

### Examples

```bash
# Sort by date (d)
python line_report.py 20260308 ./20260308 d

# Sort by name (n)
python line_report.py 20260308 C:\data\20260308 n
```

## Input Folder Structure

```
20260308/
├── [LINE] Keep Memo.txt    # Chatlog file (or [LINE]Keep Memo.txt)
├── 20260308-1-001.jpg      # Image files
├── 20260308-1-002.jpg
└── 20260308-2-001.jpg
```

## Output Format

The program generates markdown with bootstrap columns:

``````markdown
# 20260308

## หัวข้อที่ 1
ข้อความก่อนรูปภาพ
````col
```col-md
![[image1.jpg]]
![[image3.jpg]]
'```
```col-md
![[image2.jpg]]
![[image4.jpg]]
````
``````

Images are automatically ordered:
- Odd images (1,3,5) → Column 1
- Even images (2,4,6) → Column 2

---

## Build EXE (Optional)

### Install PyInstaller

```bash
pip install pyinstaller
```

### Create EXE

```bash
pyinstaller --onefile line_report.py
```

The EXE will be in the `dist/` folder.

---

## Batch File (Optional)

Create a `.bat` file to run the program with interactive input:

```batch
@echo off
echo ========================================
echo LINE Report Generator
echo ========================================
echo.
echo Enter report name:
set /p name=
echo Enter folder (or . for current folder):
set /p folder=
echo Sort by (d = date, n = name):
set /p sort=
echo.
echo Running...
line_report.exe %name% %folder% %sort%
echo.
pause
```

### Usage

1. Place `line_report.exe` and the batch file in the same folder
2. Double-click the batch file
3. Enter report name, folder path, and sort option
4. Press Enter to run

---

## Supported LINE Languages

| Language | PHOTO_MARKER |
|----------|--------------|
| English  | `[Photo]` or `Photos` |
| Thai     | `รูป` |
| Japanese | `写真` |
| Chinese  | `照片` |

Change `PHOTO_MARKER` in `line_report.py` to match your LINE app language.
