#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
LINE Report Generator
Generate markdown report from LINE chat export files and images.

Usage: python line_report.py <Report Name> <Folder> <Sort>
  Sort: d (date) or n (name)

Example: python line_report.py 20260308 ./20260308 d
"""

import os
import sys
import re
from pathlib import Path


def find_chatlog(folder):
    """Find chatlog file: [LINE]Keep Memo.txt or any single .txt file."""
    folder_path = Path(folder)
    
    # Try [LINE]Keep Memo.txt first
    keep_memo = folder_path / "[LINE]Keep Memo.txt"
    if keep_memo.exists():
        return keep_memo
    
    # Find all .txt files
    txt_files = list(folder_path.glob("*.txt"))
    
    if len(txt_files) == 0:
        print("Error: No .txt file found in folder")
        sys.exit(1)
    
    if len(txt_files) > 1:
        print("Error: Multiple .txt files found. Please use [LINE]Keep Memo.txt")
        sys.exit(1)
    
    return txt_files[0]


def get_images_by_name(folder):
    """Get all image files sorted by name (IMG_001 -> IMG_999)."""
    folder_path = Path(folder)
    images = []
    
    for f in folder_path.iterdir():
        if f.is_file() and f.suffix.lower() in ['.jpg', '.jpeg', '.png', '.gif']:
            if f.name.upper().startswith('IMG_'):
                # Extract number from IMG_XXX
                match = re.search(r'IMG_(\d+)', f.name.upper())
                if match:
                    num = int(match.group(1))
                    images.append((num, f.name))
    
    images.sort(key=lambda x: x[0])  # Ascending by name
    return images


def get_images_by_date(folder):
    """Get all image files sorted by creation date (oldest first)."""
    folder_path = Path(folder)
    images = []
    
    for f in folder_path.iterdir():
        if f.is_file() and f.suffix.lower() in ['.jpg', '.jpeg', '.png', '.gif']:
            if f.name.upper().startswith('IMG_'):
                # Get creation time
                ctime = f.stat().st_ctime
                images.append((ctime, f.name))
    
    images.sort(key=lambda x: x[0])  # Ascending by date
    return images


def extract_entries(chat_content):
    """Parse LINE chat log and extract entries with messages and photo counts."""
    entries = []
    
    lines = chat_content.split('\n')
    current_message = ""
    photo_count = 0
    last_timestamp = None
    
    for line in lines:
        line_stripped = line.strip()
        if not line_stripped:
            continue
            
        # Check if this line starts with a timestamp (entry start)
        # Format: HH:MM Name Message or HH:MM Name [Photo]
        # Can be with or without quotes around message
        timestamp_match = re.match(r'^(\d{2}:\d{2})\s+(.+?)\s+"(.+)"', line_stripped)
        timestamp_no_quote_match = re.match(r'^(\d{2}:\d{2})\s+(.+?)\s+(.+)$', line_stripped)
        photo_only_match = re.match(r'^(\d{2}:\d{2})\s+(.+?)\s+\[Photo\]$', line_stripped)
        
        if timestamp_match:
            # Save previous entry if exists
            if current_message or photo_count > 0:
                entries.append({
                    'message': current_message.strip(),
                    'photo_count': photo_count
                })
            
            # Start new entry with quoted message
            current_message = timestamp_match.group(3)
            photo_count = 0
            last_timestamp = timestamp_match.group(1)
            
        elif timestamp_no_quote_match:
            # Check if it's [Photo]
            if timestamp_no_quote_match.group(3).strip() == '[Photo]':
                ts = timestamp_no_quote_match.group(1)
                
                # Save previous entry if exists AND timestamp changed
                if (current_message or photo_count > 0) and ts != last_timestamp:
                    entries.append({
                        'message': current_message.strip(),
                        'photo_count': photo_count
                    })
                    current_message = ""
                    photo_count = 0
                
                photo_count += 1
                last_timestamp = ts
            else:
                # Save previous entry if exists
                if current_message or photo_count > 0:
                    entries.append({
                        'message': current_message.strip(),
                        'photo_count': photo_count
                    })
                
                # Start new entry without quotes
                current_message = timestamp_no_quote_match.group(3)
                photo_count = 0
                last_timestamp = timestamp_no_quote_match.group(1)
                
        elif photo_only_match:
            ts = photo_only_match.group(1)
            
            # Save previous entry if exists AND timestamp changed
            if (current_message or photo_count > 0) and ts != last_timestamp:
                entries.append({
                    'message': current_message.strip(),
                    'photo_count': photo_count
                })
                current_message = ""
                photo_count = 0
            
            photo_count += 1
            last_timestamp = ts
            
        elif line_stripped == '[Photo]':
            # [Photo] on separate line - add to current entry's photo count
            photo_count += 1
        elif current_message:
            # Continuation of message (multi-line)
            if current_message and not current_message.endswith('\n'):
                current_message += '\n'
            current_message += line_stripped
    
    # Add last entry
    if current_message or photo_count > 0:
        entries.append({
            'message': current_message.strip(),
            'photo_count': photo_count
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
    ordered_images = [img[1] for img in image_files]
    
    # Track image index
    img_idx = 0
    
    # Build markdown content
    md_content = f"# {report_name}\n\n"
    
    for i, entry in enumerate(entries, 1):
        md_content += f"## หัวข้อที่ {i}\n"
        md_content += f"{entry['message']}\n\n"
        
        photo_count = entry['photo_count']
        
        if photo_count > 0:
            # Get images for this entry
            entry_images = []
            for _ in range(photo_count):
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
                md_content += f"{img}\n"
            md_content += "```\n"
            
            # Column 2: even images
            if even_imgs:
                md_content += "```col-md\n"
                for img in even_imgs:
                    md_content += f"{img}\n"
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
        image_files = get_images_by_date(folder)
        print(f"Found {len(image_files)} images (sorted by date)")
    else:
        image_files = get_images_by_name(folder)
        print(f"Found {len(image_files)} images (sorted by name)")
    
    # Generate output path
    output_path = Path(folder) / f"{report_name}.md"
    
    # Generate report
    generate_report(entries, image_files, report_name, output_path)


if __name__ == "__main__":
    main()
