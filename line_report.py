#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
LINE Report Generator
Generate report (markdown or excel) from LINE chat export files and images.

Usage: python line_report.py <Report Name> <Folder> <Sort> [Format]
  Sort: d (date) or n (name)
  Format: md (markdown) or xlsx (excel), default: md

Example: 
  python line_report.py 20260308 ./20260308 d        # output markdown
  python line_report.py 20260308 ./20260308 d xlsx   # output excel
"""

# === CONFIGURATION ===
# Media markers - supports multiple languages (array)
MEDIA_MARKERS = [
    "รูป",        # Thai
    "[Photo]",    # English
    "Photos",     # English
    "写真",       # Japanese
    "照片",       # Chinese
    "วิดีโอ",     # Thai video
    "Videos",      # English video
]

# Supported media extensions
IMAGE_EXTENSIONS = ['.jpg', '.jpeg', '.png', '.gif']
# Note: .mpo, .bmp, .tiff, .tif, .webp are excluded as openpyxl doesn't support them
VIDEO_EXTENSIONS = ['.mp4', '.mov', '.avi']

# Skip these extensions entirely (not just for Excel)
UNSUPPORTED_IMAGE_EXTENSIONS = ['.mpo', '.bmp', '.tiff', '.tif', '.webp']

# Job title prefix
JOB_TITLE = "งานที่"
# ====================

import os
import sys
import re
from pathlib import Path

# For image dimension reading (to calculate scaling)
try:
    from PIL import Image
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False

# For Excel output
try:
    import xlsxwriter
    XLSXWRITER_AVAILABLE = True
except ImportError:
    XLSXWRITER_AVAILABLE = False

# Image sizing for Excel (10cm max width)
MAX_IMAGE_WIDTH_CM = 10


def find_chatlog(folder):
    """Find chatlog file: [LINE]Keep Memo.txt or any single .txt file."""
    folder_path = Path(folder)
    
    # Try [LINE]Keep Memo.txt first (various patterns)
    patterns = [
        "[LINE]Keep Memo.txt",
        "[LINE] Chat with Keep Memo.txt",
        "[LINE] Chat history with Keep Memo.txt"
    ]
    
    for pattern in patterns:
        keep_memo = folder_path / pattern
        if keep_memo.exists():
            return keep_memo
    
    # Also try finding any file starting with [LINE]
    for f in folder_path.iterdir():
        if f.is_file() and f.suffix.lower() == '.txt' and f.name.startswith('[LINE]'):
            return f
    
    # Find all .txt files
    txt_files = list(folder_path.glob("*.txt"))
    
    if len(txt_files) == 0:
        print("Error: No .txt file found in folder")
        sys.exit(1)
    
    if len(txt_files) > 1:
        print("Error: Multiple .txt files found. Please use [LINE]Keep Memo.txt")
        sys.exit(1)
    
    return txt_files[0]


def get_media_by_name(folder):
    """Get all media files (images + videos) sorted by name (alphabetically, ascending)."""
    folder_path = Path(folder)
    media = []
    
    all_extensions = IMAGE_EXTENSIONS + VIDEO_EXTENSIONS
    
    for f in folder_path.iterdir():
        if f.is_file() and f.suffix.lower() in all_extensions:
            # Skip unsupported image types
            if f.suffix.lower() in UNSUPPORTED_IMAGE_EXTENSIONS:
                continue
            media.append((f.name, f.name))  # (sort_key, original_name)
    
    # Sort alphabetically by filename
    media.sort(key=lambda x: x[0])
    return media


def get_media_by_date(folder):
    """Get all media files sorted by creation date (oldest first)."""
    folder_path = Path(folder)
    media = []
    
    all_extensions = IMAGE_EXTENSIONS + VIDEO_EXTENSIONS
    
    for f in folder_path.iterdir():
        if f.is_file() and f.suffix.lower() in all_extensions:
            # Skip unsupported image types
            if f.suffix.lower() in UNSUPPORTED_IMAGE_EXTENSIONS:
                continue
            # Get creation time
            ctime = f.stat().st_ctime
            # Use original name as secondary sort for stability
            media.append((ctime, f.name))
    
    # Sort by date (ASC), then by name for stability
    media.sort(key=lambda x: (x[0], x[1]))
    return media


def extract_entries(chat_content):
    """Parse LINE chat log and extract entries with messages and media counts.
    
    Entries are grouped by content continuity:
    - If next message has NO text (only media), it will be grouped with previous entry
    
    Returns list of dict with keys: date, time, sender, message, media_count
    """
    entries = []
    
    lines = chat_content.split('\n')
    current_date = ""
    current_time = ""
    current_sender = ""
    current_message = ""
    media_count = 0
    
    for line in lines:
        line_stripped = line.strip()
        if not line_stripped:
            continue
        
        # Check for date line: YYYY.MM.DD วัน[วัน]
        date_match = re.match(r'^(\d{4}\.\d{2}\.\d{2})\s+วัน', line_stripped)
        if date_match:
            current_date = date_match.group(1)
            continue
        
        # Check if this line starts with a timestamp (entry start)
        # Format: HH:MM Name Message or HH:MM Name รูป
        timestamp_match = re.match(r'^(\d{2}:\d{2})\s+(.+?)\s+"(.+)"', line_stripped)
        timestamp_no_quote_match = re.match(r'^(\d{2}:\d{2})\s+(.+?)\s+(.+)$', line_stripped)
        
        if timestamp_match:
            # Save previous entry if exists
            if current_message or media_count > 0:
                entries.append({
                    'date': current_date,
                    'time': current_time,
                    'sender': current_sender,
                    'message': current_message.strip(),
                    'media_count': media_count
                })
            
            # Start new entry with quoted message
            current_time = timestamp_match.group(1)
            current_sender = timestamp_match.group(2)
            current_message = timestamp_match.group(3)
            media_count = 0
            
        elif timestamp_no_quote_match:
            # Check if it's a media marker (photo/video)
            if timestamp_no_quote_match.group(3).strip() in MEDIA_MARKERS:
                # Media-only message: ADD to current entry (don't create new entry)
                if not current_time:  # First media has no time yet
                    current_time = timestamp_no_quote_match.group(1)
                    current_sender = timestamp_no_quote_match.group(2)
                media_count += 1
            else:
                # New message with text: save previous and start new
                if current_message or media_count > 0:
                    entries.append({
                        'date': current_date,
                        'time': current_time,
                        'sender': current_sender,
                        'message': current_message.strip(),
                        'media_count': media_count
                    })
                
                # Start new entry without quotes
                current_time = timestamp_no_quote_match.group(1)
                current_sender = timestamp_no_quote_match.group(2)
                current_message = timestamp_no_quote_match.group(3)
                media_count = 0
                
        elif line_stripped in MEDIA_MARKERS:
            # Media marker on separate line: ADD to current entry
            media_count += 1
            
        elif current_message:
            # Continuation of message (multi-line)
            if current_message and not current_message.endswith('\n'):
                current_message += '\n'
            current_message += line_stripped
    
    # Add last entry
    if current_message or media_count > 0:
        entries.append({
            'date': current_date,
            'time': current_time,
            'sender': current_sender,
            'message': current_message.strip(),
            'media_count': media_count
        })
    
    return entries


def interleave_entry_images(images):
    """Reorder images for a single entry: odd first, then even."""
    if len(images) <= 1:
        return images
    
    odd_images = []
    even_images = []
    
    for i, img in enumerate(images):
        if i % 2 == 0:  # 0, 2, 4 -> odd positions (1st, 3rd, 5th)
            odd_images.append(img)
        else:  # 1, 3, 5 -> even positions (2nd, 4th, 6th)
            even_images.append(img)
    
    # Odd then even
    return odd_images + even_images


def generate_markdown(entries, image_files, report_name, output_path):
    """Generate markdown report."""
    
    # Get image names in order (already sorted by name or date)
    # image_files is list of (sort_key, original_name)
    ordered_images = [img[1] for img in image_files]
    
    # Track image index
    img_idx = 0
    
    # Build markdown content
    md_content = f"# {report_name}\n\n"
    
    for i, entry in enumerate(entries, 1):
        md_content += f"## {JOB_TITLE} {i}\n"
        md_content += f"{entry['message']}\n\n"
        
        media_count = entry['media_count']
        
        if media_count > 0:
            # Get images for this entry
            entry_images = []
            for _ in range(media_count):
                if img_idx < len(ordered_images):
                    entry_images.append(ordered_images[img_idx])
                    img_idx += 1
                else:
                    entry_images.append("[ไม่พบรูป]")
            
            # Interleave within this entry
            entry_images = interleave_entry_images(entry_images)
            
            # Split into odd and even columns
            half = (len(entry_images) + 1) // 2
            odd_imgs = entry_images[:half]
            even_imgs = entry_images[half:]
            
            # Build markdown with proper col blocks
            md_content += "````col\n"
            
            # Column 1: odd images
            md_content += "```col-md\n"
            for img in odd_imgs:
                md_content += f"![[{img}]]\n"
            md_content += "```\n"
            
            # Column 2: even images
            if even_imgs:
                md_content += "```col-md\n"
                for img in even_imgs:
                    md_content += f"![[{img}]]\n"
                md_content += "```\n"
            
            md_content += "````\n\n"
    
    # Write to file
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(md_content)
    
    print(f"✓ Report generated: {output_path}")
    print(f"  - Entries: {len(entries)}")
    print(f"  - Images used: {img_idx}")


def generate_excel(entries, image_files, report_name, output_path, folder):
    """Generate Excel report with Cover and Detail sheets using xlsxwriter."""
    
    if not XLSXWRITER_AVAILABLE:
        print("Error: xlsxwriter is not installed. Please install with: pip install xlsxwriter")
        sys.exit(1)
    
    # Get image files sorted (name or date)
    ordered_images = [img[1] for img in image_files]
    
    # Track image index
    img_idx = 0
    
    # Get all image paths for the folder
    folder_path = Path(folder)
    
    # Create workbook
    wb = xlsxwriter.Workbook(output_path)
    
    # ============================================
    # Sheet 1: "Cover" - Data table
    # ============================================
    ws_cover = wb.add_worksheet("Cover")
    
    # Formats
    header_format = wb.add_format({
        'bold': True,
        'text_wrap': True,
        'valign': 'top'
    })
    cell_format = wb.add_format({
        'text_wrap': True,
        'valign': 'top'
    })
    
    # Header row
    headers = ["No.", "Date", "Time", "Sender", "Text Content", "Image Count", "Image Filenames", "Image Path"]
    for col, header in enumerate(headers):
        ws_cover.write(0, col, header, header_format)
    
    # Set column widths
    ws_cover.set_column(0, 0, 5)   # No.
    ws_cover.set_column(1, 1, 12)  # Date
    ws_cover.set_column(2, 2, 8)   # Time
    ws_cover.set_column(3, 3, 20)  # Sender
    ws_cover.set_column(4, 4, 40)  # Text Content
    ws_cover.set_column(5, 5, 10) # Image Count
    ws_cover.set_column(6, 6, 30)  # Image Filenames
    ws_cover.set_column(7, 7, 20)  # Image Path
    
    # Add data rows
    for row_idx, entry in enumerate(entries, 1):
        # Get images for this entry
        entry_images = []
        for _ in range(entry['media_count']):
            if img_idx < len(ordered_images):
                entry_images.append(ordered_images[img_idx])
                img_idx += 1
        
        # Write row data
        ws_cover.write(row_idx, 0, row_idx, cell_format)  # No.
        ws_cover.write(row_idx, 1, entry.get('date', ''), cell_format)
        ws_cover.write(row_idx, 2, entry.get('time', ''), cell_format)
        ws_cover.write(row_idx, 3, entry.get('sender', ''), cell_format)
        ws_cover.write(row_idx, 4, entry.get('message', ''), cell_format)
        ws_cover.write(row_idx, 5, entry['media_count'], cell_format)
        ws_cover.write(row_idx, 6, ", ".join(entry_images) if entry_images else "", cell_format)
        ws_cover.write(row_idx, 7, f"{folder}/", cell_format)
    
    # ============================================
    # Sheet 2: "Detail" - Text + embedded images
    # ============================================
    ws_detail = wb.add_worksheet("Detail")
    
    # Formats for Detail sheet
    header_format_detail = wb.add_format({
        'bold': True,
        'font_size': 12
    })
    text_format_detail = wb.add_format({
        'font_size': 10,
        'text_wrap': True,
        'valign': 'top'
    })
    placeholder_format = wb.add_format({
        'font_size': 9,
        'italic': True,
        'font_color': '#808080'
    })
    
    # Set column widths
    ws_detail.set_column(0, 0, 60)  # A column wider
    ws_detail.set_column(1, 1, 40)
    
    # Reset image index for Detail sheet
    img_idx = 0
    current_row = 0
    
    for entry_idx, entry in enumerate(entries, 1):
        # Get images for this entry
        entry_images = []
        for _ in range(entry['media_count']):
            if img_idx < len(ordered_images):
                entry_images.append(ordered_images[img_idx])
                img_idx += 1
        
        # === Entry Header ===
        header_text = f"งานที่ {entry_idx}"
        if entry.get('date'):
            header_text += f" - {entry['date']}"
        if entry.get('time'):
            header_text += f" {entry['time']}"
        if entry.get('sender'):
            header_text += f" - {entry['sender']}"
        
        ws_detail.write(current_row, 0, header_text, header_format_detail)
        current_row += 1
        
        # === Text Content ===
        if entry.get('message'):
            ws_detail.write(current_row, 0, entry['message'], text_format_detail)
            ws_detail.set_row(current_row, 60)  # Set row height for text
            current_row += 1
        
        # === Images (embedded with xlsxwriter, arranged horizontally) ===
        if entry_images:
            # Calculate positions for each image (horizontal layout)
            img_col = 0  # Start from column 0
            max_height = 0  # Track max height for row height
            
            for img_name in entry_images:
                img_path = folder_path / img_name
                if img_path.exists():
                    try:
                        # Calculate scale to fit max width 10cm (~284 points at 96 DPI)
                        max_width_pt = 10 * 28.35  # 283.5 points
                        
                        if PIL_AVAILABLE:
                            try:
                                with Image.open(str(img_path)) as img:
                                    img_width_px = img.width
                                    img_height_px = img.height
                                    
                                    # Calculate scale to fit max width
                                    if img_width_px > max_width_pt:
                                        x_scale = max_width_pt / img_width_px
                                        y_scale = x_scale  # Keep aspect ratio
                                    else:
                                        x_scale = 1
                                        y_scale = 1
                                    
                                    # Calculate actual width after scaling (in points)
                                    actual_width_pt = img_width_px * x_scale
                            except:
                                x_scale = 1
                                y_scale = 1
                                actual_width_pt = 100  # Default fallback
                        else:
                            x_scale = 1
                            y_scale = 1
                            actual_width_pt = 100
                        
                        # Calculate x_offset to prevent overlap
                        # Add minimum 5 points spacing between images
                        x_offset = max(5, int(actual_width_pt * 0.05))  # 5% of width or minimum 5pt
                        
                        # Insert image at current column position
                        ws_detail.insert_image(
                            current_row, img_col,
                            str(img_path),
                            {'x_scale': x_scale, 'y_scale': y_scale, 'x_offset': x_offset, 'y_offset': 1}
                        )
                        
                        # Track max height for row setting
                        if PIL_AVAILABLE:
                            try:
                                with Image.open(str(img_path)) as img:
                                    scaled_height = img.height * y_scale
                                    if scaled_height > max_height:
                                        max_height = scaled_height
                            except:
                                pass
                        
                        # Move to next column position (approximate)
                        img_col += 1
                        
                    except Exception as e:
                        # If image fails, show placeholder
                        ws_detail.write(current_row, img_col, f"[โหลดรูปไม่ได้: {img_name}]", placeholder_format)
                        img_col += 1
                else:
                    # Image not found
                    ws_detail.write(current_row, img_col, f"[ไม่พบรูป: {img_name}]", placeholder_format)
                    img_col += 1
            
            # Set row height to accommodate images
            if max_height > 0:
                ws_detail.set_row(current_row, int(max_height * 1.1))  # Add 10% buffer
            current_row += 1
        
        # === Spacing row between entries ===
        current_row += 1
    
    # Close workbook
    wb.close()
    
    print(f"✓ Excel report generated: {output_path}")
    print(f"  - Entries: {len(entries)}")
    print(f"  - Images used: {img_idx}")
    print(f"  - Sheets: Cover, Detail")


def main():
    if len(sys.argv) < 4 or len(sys.argv) > 5:
        print("Usage: python line_report.py <Report Name> <Folder> <Sort> [Format]")
        print("  Sort: d (date) or n (name)")
        print("  Format: md (markdown) or xlsx (excel), default: md")
        print("Example:")
        print("  python line_report.py 20260308 ./20260308 d        # output markdown")
        print("  python line_report.py 20260308 ./20260308 d xlsx   # output excel")
        sys.exit(1)
    
    report_name = sys.argv[1]
    folder = sys.argv[2]
    sort_option = sys.argv[3].lower()
    output_format = sys.argv[4].lower() if len(sys.argv) == 5 else 'md'
    
    # Validate sort option
    if sort_option not in ['d', 'n']:
        print("Error: Sort must be 'd' (date) or 'n' (name)")
        sys.exit(1)
    
    # Validate format option
    if output_format not in ['md', 'xlsx']:
        print("Error: Format must be 'md' (markdown) or 'xlsx' (excel)")
        sys.exit(1)
    
    # Check xlsx dependency
    if output_format == 'xlsx' and not XLSXWRITER_AVAILABLE:
        print("Error: xlsxwriter is not installed. Please install with: pip install xlsxwriter")
        sys.exit(1)
    
    # Find chatlog
    print(f"Finding chatlog in: {folder}")
    chatlog_path = find_chatlog(folder)
    print(f"Using chatlog: {chatlog_path.name}")
    
    # Read chatlog
    with open(chatlog_path, 'r', encoding='utf-8') as f:
        chat_content = f.read()
    
    # Parse entries
    entries = extract_entries(chat_content)
    print(f"Found {len(entries)} entries")
    
    # Get images based on sort option
    if sort_option == 'd':
        image_files = get_media_by_date(folder)
        print(f"Found {len(image_files)} images (sorted by date)")
    else:
        image_files = get_media_by_name(folder)
        print(f"Found {len(image_files)} images (sorted by name)")
    
    # Generate output based on format
    if output_format == 'xlsx':
        output_path = Path(folder) / f"{report_name}.xlsx"
        generate_excel(entries, image_files, report_name, output_path, folder)
    else:
        output_path = Path(folder) / f"{report_name}.md"
        generate_markdown(entries, image_files, report_name, output_path)


if __name__ == "__main__":
    main()