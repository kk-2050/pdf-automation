# =============================================================================
# Program Name : processor.py
# Author       : Kaori Kashiwagi
# Date         : 2026-04-03
# Purpose      : Core PDF processing engine for the PDF Keyword Highlight Tool.
#                Handles the following tasks for each PDF file:
#                  - Open the PDF using PyMuPDF (fitz)
#                  - Search for keywords within the specified page range
#                  - Apply color highlights with adjustable opacity
#                  - Move the processed PDF to the Destination folder
#                  - Write processing results to the shared Excel log file
#                  - Support overwrite mode to reprocess existing files
# =============================================================================

import os
import shutil
import getpass
import fitz  # PyMuPDF — library for reading and annotating PDF files

from utils import (color_to_rgb01, is_whole_word_match,
                   append_log, build_log_row, HIGHLIGHT_OPACITY)

# Name of the shared Excel log file saved in C:\PDFs\
LOG_FILENAME = "pdf_highlight_log.xlsx"

# Temporary file suffix used while processing (avoids overwriting originals mid-process)
OUTPUT_SUFFIX = "_highlighted_tmp"


def highlight_pdf(in_pdf, out_pdf, keywords, color_rgb, opacity,
                  case_sensitive, whole_word, begin_page, end_page):
    """
    Open a PDF, search for keywords within the specified page range,
    and apply highlight annotations with the given color and opacity.

    Parameters:
        in_pdf       : Path to the input PDF file
        out_pdf      : Path where the highlighted PDF will be saved
        keywords     : List of keyword strings to search for
        color_rgb    : Highlight color as RGB tuple (values 0.0–1.0)
        opacity      : Highlight transparency (0.0=transparent, 1.0=solid)
        case_sensitive: Whether to match uppercase/lowercase exactly
        whole_word   : Whether to match whole words only
        begin_page   : First page to search (1-based)
        end_page     : Last page to search (1-based), or large number for all pages

    Returns:
        (pages_scanned, total_hits) — number of pages searched and total keyword matches
    """
    doc = fitz.open(in_pdf)
    total_pages = doc.page_count

    # Convert 1-based page numbers to 0-based index used by PyMuPDF
    p_start = max(0, begin_page - 1)
    p_end = min(total_pages - 1, end_page - 1) if end_page else total_pages - 1
    pages_scanned = p_end - p_start + 1
    total_hits = 0

    # Loop through each page in the specified range
    for page_idx in range(p_start, p_end + 1):
        page = doc[page_idx]
        page_text = page.get_text("text")  # Extract plain text for whole-word matching

        # Search for each keyword on this page
        for kw in keywords:
            if not kw:
                continue

            # Find all positions (rectangles) where the keyword appears
            rects = page.search_for(kw)
            if not rects:
                continue

            # Apply a highlight annotation to each match
            for rect in rects:
                # Skip if whole-word mode is ON and the match is not a whole word
                if whole_word and not is_whole_word_match(page_text, kw):
                    continue

                annot = page.add_highlight_annot(rect)
                annot.set_colors(stroke=color_rgb)  # Set highlight color
                annot.set_opacity(opacity)           # Set transparency (40% = readable)
                annot.update()
                total_hits += 1

    # Save the annotated PDF to the temporary output path
    doc.save(out_pdf, garbage=4, deflate=True)  # garbage=4 removes unused objects
    doc.close()
    return pages_scanned, total_hits


def process_customer(rule, input_folder, cancel_flag=None,
                     progress_callback=None, shared_log_path=None,
                     overwrite=False):
    """
    Process all PDF files in the given Input folder for one customer.

    For each PDF:
      1. Check if it already exists in the Destination folder
         - If Overwrite is OFF → skip and log SKIPPED_EXISTS
         - If Overwrite is ON  → delete existing file and reprocess
      2. Run highlight_pdf() to add keyword highlights
      3. Move the highlighted PDF to the Destination folder
      4. Append the result to the shared Excel log file

    Parameters:
        rule             : CustomerRule object (from excel_master.py)
        input_folder     : Full path to the Input folder containing PDFs
        cancel_flag      : Dict with {"cancel": bool} — set True to stop mid-run
        progress_callback: Function called after each file for progress bar updates
        shared_log_path  : Full path to the shared Excel log file (C:\\PDFs\\pdf_highlight_log.xlsx)
        overwrite        : If True, delete and reprocess existing files in Done folder

    Returns:
        (results, log_path)
        results  : List of (filename, status) tuples
        log_path : Path where the log was written
    """
    dest_folder = rule.destination_folder

    # Ensure Input and Destination folders exist
    os.makedirs(input_folder, exist_ok=True)
    os.makedirs(dest_folder, exist_ok=True)

    # Determine where to write the log
    if shared_log_path:
        log_path = shared_log_path
        log_dir = os.path.dirname(log_path)
        if log_dir:
            os.makedirs(log_dir, exist_ok=True)  # Create log directory if needed
    else:
        # Fallback: save log in the Destination folder
        log_path = os.path.join(dest_folder, LOG_FILENAME)

    user = getpass.getuser()  # Get current Windows username for the log
    color_rgb = color_to_rgb01(rule.color)  # Convert color name to RGB tuple

    # Use a very large number if End_Page is not specified (process all pages)
    end_page = rule.end_page if rule.end_page else 999999

    # Get list of all PDF files in the Input folder
    pdfs = [f for f in os.listdir(input_folder) if f.lower().endswith(".pdf")]
    results = []

    # Process each PDF file one by one
    for i, pdf_name in enumerate(pdfs):

        # Check if user clicked Stop — exit loop if so
        if cancel_flag and cancel_flag.get("cancel"):
            results.append((pdf_name, "CANCELLED"))
            break

        in_path   = os.path.join(input_folder, pdf_name)
        dest_path = os.path.join(dest_folder, pdf_name)

        # Update progress bar via callback
        if progress_callback:
            progress_callback(pdf_name, i + 1, len(pdfs))

        # ── Check for existing file in Destination folder ──────────────────
        if os.path.exists(dest_path) and not overwrite:
            # Overwrite is OFF → skip this file
            row = build_log_row(
                rule.customer, pdf_name, ",".join(rule.keywords), 0,
                rule.color, rule.begin_page, rule.end_page,
                dest_path, "SKIPPED_EXISTS", user)
            append_log(log_path, row)
            results.append((pdf_name, "SKIPPED_EXISTS"))
            continue

        if os.path.exists(dest_path) and overwrite:
            # Overwrite is ON → delete the existing file before reprocessing
            try:
                os.remove(dest_path)
            except Exception as e:
                row = build_log_row(
                    rule.customer, pdf_name, ",".join(rule.keywords), 0,
                    rule.color, rule.begin_page, rule.end_page,
                    dest_path, f"ERROR:Cannot remove existing file: {e}", user)
                append_log(log_path, row)
                results.append((pdf_name, f"ERROR:Cannot remove: {e}"))
                continue

        # ── Process the PDF ────────────────────────────────────────────────
        base, ext = os.path.splitext(pdf_name)
        temp_out = os.path.join(input_folder, f"{base}{OUTPUT_SUFFIX}{ext}")

        try:
            # Step 1: Apply highlights and save to a temporary file
            pages_scanned, hits = highlight_pdf(
                in_pdf=in_path,
                out_pdf=temp_out,
                keywords=rule.keywords,
                color_rgb=color_rgb,
                opacity=HIGHLIGHT_OPACITY,
                case_sensitive=rule.case_sensitive,
                whole_word=rule.whole_word,
                begin_page=rule.begin_page,
                end_page=end_page)

            # Step 2: Move the temporary file to the Destination folder
            shutil.move(temp_out, dest_path)

            # Step 3: Log the success result
            status = f"OK_HITS_{hits}"
            row = build_log_row(
                rule.customer, pdf_name, ",".join(rule.keywords), pages_scanned,
                rule.color, rule.begin_page, rule.end_page,
                dest_path, status, user)
            append_log(log_path, row)
            results.append((pdf_name, status))

        except Exception as e:
            # Clean up temp file if something went wrong
            if os.path.exists(temp_out):
                try:
                    os.remove(temp_out)
                except Exception:
                    pass
            # Log the error
            row = build_log_row(
                rule.customer, pdf_name, ",".join(rule.keywords), 0,
                rule.color, rule.begin_page, rule.end_page,
                dest_folder, f"ERROR:{e}", user)
            append_log(log_path, row)
            results.append((pdf_name, f"ERROR:{e}"))

    return results, log_path
