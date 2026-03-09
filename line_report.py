#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
LINE Report Generator
Generate markdown report from LINE chat export files and images.

Usage: python line_report.py <Report Name> <Folder> <Sort>
  Sort: d (date) or n (name)

Example: python line_report.py 20260308 ./20260308 d
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
    "Video",      # English video
]

# Supported media extensions
IMAGE_EXTENSIONS = ['.jpg', '.jpeg', '.png', '.gif']
VIDEO_EXTENSIONS = ['.mp4', '.mov', '.avi']

# Job title prefix
JOB_TITLE = "งานที่"
# ====================

import os
import sys
import re
from pathlib import Path


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
    """
    entries = []
    
    lines = chat_content.split('\n')
    current_message = ""
    media_count = 0
    
    for line in lines:
        line_stripped = line.strip()
        if not line_stripped:
            continue
            
        # Check if this line starts with a timestamp (entry start)
        # Format: HH:MM Name Message or HH:MM Name รูป
        timestamp_match = re.match(r'^(\d{2}:\d{2})\s+(.+?)\s+"(.+)"', line_stripped)
        timestamp_no_quote_match = re.match(r'^(\d{2}:\d{2})\s+(.+?)\s+(.+)$', line_stripped)
        
        if timestamp_match:
            # Save previous entry if exists
            if current_message or media_count > 0:
                entries.append({
                    'message': current_message.strip(),
                    'media_count': media_count
                })
            
            # Start new entry with quoted message
            current_message = timestamp_match.group(3)
            media_count = 0
            
        elif timestamp_no_quote_match:
            # Check if it's a media marker (photo/video)
            if timestamp_no_quote_match.group(3).strip() in MEDIA_MARKERS:
                # Media-only message: ADD to current entry (don't create new entry)
                media_count += 1
            else:
                # New message with text: save previous and start new
                if current_message or media_count > 0:
                    entries.append({
                        'message': current_message.strip(),
                        'media_count': media_count
                    })
                
                # Start new entry without quotes
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


def generate_report(entries, image_files, report_name, output_path):
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


def main():
    if len(sys.argv) != 4:
        print("Usage: python line_report.py <Report Name> <Folder> <Sort>")
        print("  Sort: d (date) or n (name)")
        print("Example: python line_report.py 20260308 ./20260308 d")
        sys.exit(1)
    
    report_name = sys.argv[1]
    folder = sys.argv[2]
    sort_option = sys.argv[3].lower()
    
    # Validate sort option
    if sort_option not in ['d', 'n']:
        print("Error: Sort must be 'd' (date) or 'n' (name)")
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
    
    # Generate output path
    output_path = Path(folder) / f"{report_name}.md"
    
    # Generate report
    generate_report(entries, image_files, report_name, output_path)


if __name__ == "__main__":
    main()
