#!/usr/bin/env python3
"""
Bib Generator
=============
Reads a list of runners from Excel and produces a PowerPoint file with one
bib slide per runner, using an existing PPTX slide as the visual template.

Usage:
    python generate_bibs.py runners.xlsx template.pptx output.pptx

Excel file must have these columns (case-insensitive, order doesn't matter):
    Date, First Name, Last Name, Distance

Template PPTX must have at least one slide containing:
    - A text box with "Month" somewhere in it (e.g. "Monthly Run Month 2026")
    - A text box containing a bib number placeholder (e.g. "500")
    - A text box containing "FirstName LastName"
    (Any images/logos on that slide are copied as-is onto every bib.)
"""

import copy
import re
import sys
from datetime import datetime

import pandas as pd
from pptx import Presentation

# ---------------------------------------------------------------------------
# CONFIG — edit these to match your event
# ---------------------------------------------------------------------------

# Which slide (0-indexed) in the template file to use as the master design.
TEMPLATE_SLIDE_INDEX = 0

# Starting bib number for each distance category.
# If a distance isn't listed here, the script tries to auto-derive a start
# number from the leading digits of the distance text (e.g. "5K" -> 500,
# "10K" -> 1000, "21K" -> 2100) by multiplying by 100.
# Add explicit overrides here for anything non-numeric (e.g. "Fun Run").
CATEGORY_START_OVERRIDES = {
    # "Fun Run": 100,
    # "Half Marathon": 2100,
    # "Marathon": 4200,
}

# Placeholder text in the template to find-and-replace.
NAME_PLACEHOLDER = "FirstName LastName"
BIB_NUMBER_PLACEHOLDERS = {"500", "1000"}  # any of these found = bib number box
MONTH_TOKEN = "Month"  # literal word to replace with the actual month name

# ---------------------------------------------------------------------------


def load_runners(excel_path):
    df = pd.read_excel(excel_path)
    # Normalize column names: case-insensitive, ignore extra spaces
    colmap = {c.lower().strip(): c for c in df.columns}

    def find_col(*aliases):
        for a in aliases:
            if a in colmap:
                return colmap[a]
        raise ValueError(
            f"Could not find a column matching any of {aliases}. "
            f"Found columns: {list(df.columns)}"
        )

    date_col = find_col("date")
    first_col = find_col("first name", "firstname")
    last_col = find_col("last name", "lastname")
    dist_col = find_col("distance")

    df = df.rename(
        columns={
            date_col: "Date",
            first_col: "First Name",
            last_col: "Last Name",
            dist_col: "Distance",
        }
    )
    df["Date"] = pd.to_datetime(df["Date"])
    df["Distance"] = df["Distance"].astype(str).str.strip()
    return df[["Date", "First Name", "Last Name", "Distance"]]


def get_category_start(distance):
    if distance in CATEGORY_START_OVERRIDES:
        return CATEGORY_START_OVERRIDES[distance]
    match = re.match(r"\s*(\d+)", distance)
    if match:
        return int(match.group(1)) * 100
    raise ValueError(
        f"Can't auto-derive a starting bib number for distance '{distance}'. "
        f"Add it to CATEGORY_START_OVERRIDES in generate_bibs.py."
    )


def assign_bib_numbers(df):
    """Assign a sequential bib number per distance category, in row order."""
    df = df.copy()
    next_number = {}
    bib_numbers = []
    for _, row in df.iterrows():
        dist = row["Distance"]
        if dist not in next_number:
            next_number[dist] = get_category_start(dist)
        bib_numbers.append(next_number[dist])
        next_number[dist] += 1
    df["Bib Number"] = bib_numbers
    return df


R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
R_EMBED = f"{{{R_NS}}}embed"
R_LINK = f"{{{R_NS}}}link"


def duplicate_slide(prs, source_slide):
    """Duplicate a slide (shapes + images) and append it to the presentation."""
    layout = source_slide.slide_layout
    new_slide = prs.slides.add_slide(layout)

    # Remove any placeholder shapes the layout auto-added
    for shp in list(new_slide.shapes):
        shp._element.getparent().remove(shp._element)

    # Copy every shape from the source slide, remapping any image/media
    # relationship IDs (r:embed / r:link) to new relationships on this slide.
    for shape in source_slide.shapes:
        new_el = copy.deepcopy(shape._element)
        for attr in (R_EMBED, R_LINK):
            for el in new_el.iter():
                old_rid = el.get(attr)
                if old_rid:
                    rel = source_slide.part.rels[old_rid]
                    new_rid = new_slide.part.rels.get_or_add(rel.reltype, rel.target_part)
                    el.set(attr, new_rid)
        new_slide.shapes._spTree.append(new_el)

    return new_slide


def set_run_text(shape, new_text):
    """Replace the text of the first run in a text-frame shape, keeping formatting."""
    tf = shape.text_frame
    tf.paragraphs[0].runs[0].text = new_text
    # clear any extra runs in that paragraph so old text doesn't linger
    for extra_run in tf.paragraphs[0].runs[1:]:
        extra_run.text = ""


def fill_bib_slide(slide, first_name, last_name, bib_number, month_name):
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        text = shape.text_frame.text
        if text == NAME_PLACEHOLDER:
            set_run_text(shape, f"{first_name} {last_name}")
        elif text.strip() in BIB_NUMBER_PLACEHOLDERS:
            set_run_text(shape, str(bib_number))
        elif re.search(rf"\b{MONTH_TOKEN}\b", text):
            set_run_text(shape, re.sub(rf"\b{MONTH_TOKEN}\b", month_name, text))


def generate(excel_path, template_path, output_path):
    df = load_runners(excel_path)
    df = assign_bib_numbers(df)

    prs = Presentation(template_path)
    template_slide = prs.slides[TEMPLATE_SLIDE_INDEX]

    # Record how many slides existed before we start duplicating, and which
    # template slide to clone from — we clone from a fresh copy of it each time.
    template_xml_slides = list(prs.slides._sldIdLst)

    for _, row in df.iterrows():
        month_name = row["Date"].strftime("%B")
        new_slide = duplicate_slide(prs, template_slide)
        fill_bib_slide(
            new_slide,
            row["First Name"],
            row["Last Name"],
            row["Bib Number"],
            month_name,
        )

    # Remove the original template slide(s) so only generated bibs remain
    xml_slides = prs.slides._sldIdLst
    for sldId in template_xml_slides:
        xml_slides.remove(sldId)

    prs.save(output_path)
    print(f"Generated {len(df)} bibs -> {output_path}")
    print(df[["First Name", "Last Name", "Distance", "Bib Number"]].to_string(index=False))


if __name__ == "__main__":
    if len(sys.argv) != 4:
        print("Usage: python generate_bibs.py runners.xlsx template.pptx output.pptx")
        sys.exit(1)
    generate(sys.argv[1], sys.argv[2], sys.argv[3])
