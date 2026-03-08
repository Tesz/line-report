# LINE Report Generator

Python program to generate markdown report from LINE chat export files and images.

## Requirements

- Python 3.x

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
├── [LINE]Keep Memo.txt    # Chatlog file
├── IMG_001.jpg             # Image files
├── IMG_002.jpg
└── IMG_003.jpg
```

## Output

Generates `{Report Name}.md` in the same folder.
