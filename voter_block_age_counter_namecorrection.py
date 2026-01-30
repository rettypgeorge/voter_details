import os
import glob
import re
import cv2
import numpy as np
import pytesseract
from pdf2image import convert_from_path
from openpyxl import Workbook
from datetime import datetime

# ---------------- POPPLER PATH ----------------
os.environ["PATH"] += os.pathsep + r"C:\poppler\poppler-24.02.0\Library\bin"

# ---------------- CONFIG ----------------
PDF_FOLDER = "fixed_pdfs"
DPI = 300
START_PAGE = 3   # voter data starts from page 3

# ---------------- NAME CLEANING ----------------
def clean_malayalam_name(raw):
    if not raw:
        return None

    name = raw.strip()

    # remove label noise
    prefixes = ["പേര്", "പേ", "പര", ":"]
    for p in prefixes:
        if name.startswith(p):
            name = name[len(p):].strip()

    # cut relation text if merged
    for r in ["പിതാവ്", "ഭർത്താവ്", "അമ്മ", "അച്ഛൻ"]:
        if r in name:
            name = name.split(r)[0].strip()

    # keep Malayalam only
    name = re.sub(r"[^അ-ഹ ാ-ൗ ്]", "", name)

    # normalize spaces
    name = re.sub(r"\s+", " ", name).strip()

    # common OCR corrections (SAFE ONLY)
    corrections = {
        "ലക്ഷമി": "ലക്ഷ്മി",
        "സനിൽ": "സുനിൽ",
        "സനില": "സുനില",
        "രമന്": "രാമൻ",
        "രാമന്": "രാമൻ",
        "സിത": "സീത",
    }

    for w, c in corrections.items():
        if w in name:
            name = name.replace(w, c)

    if len(name) < 2:
        return None

    return name

# ---------------- LOAD PDFs ----------------
pdfs = sorted(glob.glob(os.path.join(PDF_FOLDER, "*.pdf")))
print("Found PDFs:", len(pdfs))

# ---------------- EXCEL SETUP ----------------
wb = Workbook()
ws = wb.active
ws.title = "Voters"

ws.append([
    "House No",
    "Name",
    "Age",
    "Gender",
    "EPIC ID",
    "Page No"
])

total_saved = 0

# ---------------- PROCESS ----------------
for pdf in pdfs:

    pages = convert_from_path(pdf, dpi=DPI)
    print(f"\nPROCESSING {os.path.basename(pdf)} → pages:", len(pages))

    last_house = None

    for page_no in range(START_PAGE - 1, len(pages)):

        img = np.array(pages[page_no])
        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

        th = cv2.adaptiveThreshold(
            gray, 255,
            cv2.ADAPTIVE_THRESH_MEAN_C,
            cv2.THRESH_BINARY_INV,
            35, 15
        )

        contours, _ = cv2.findContours(
            th, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE
        )

        blocks = []
        for c in contours:
            x, y, w, h = cv2.boundingRect(c)
            if w > 140 and h > 95:
                blocks.append((x, y, w, h))

        blocks.sort(key=lambda b: (b[1], b[0]))
        print(f"Page {page_no+1} → Blocks:", len(blocks))

        for x, y, w, h in blocks:

            block = gray[y:y+h, x:x+w]

            # ---------- HOUSE NO ----------
            hx1, hx2 = int(w*0.02), int(w*0.25)
            hy1, hy2 = int(h*0.02), int(h*0.25)

            house_crop = block[hy1:hy2, hx1:hx2]
            house_crop = cv2.threshold(house_crop, 150, 255, cv2.THRESH_BINARY)[1]

            house_text = pytesseract.image_to_string(
                house_crop, lang="eng", config="--psm 7 digits"
            )

            m = re.search(r"\d{1,4}", house_text)
            if m:
                house = m.group()
                last_house = house
            else:
                house = last_house

            # ---------- FULL OCR ----------
            text = pytesseract.image_to_string(
                block, lang="mal+eng", config="--psm 6"
            )

            lines = [l.strip() for l in text.splitlines() if l.strip()]
            if not lines:
                continue

            # ---------- NAME ----------
            name = None
            for ln in lines:
                if re.search(r"[അ-ഹ]", ln):
                    name = clean_malayalam_name(ln)
                    if name:
                        break

            # ---------- AGE ----------
            age = None
            age_match = re.search(r"(വയ|പ്രായ)[^\d]{0,6}(\d{1,2})", text)
            if age_match:
                age = int(age_match.group(2))

            # ---------- GENDER ----------
            gender = None
            if "ആൺ" in text:
                gender = "Male"
            elif "സ്ത്രീ" in text:
                gender = "Female"

            # ---------- EPIC ID ----------
            epic = None
            epic_match = re.search(r"[A-Z]{3}\d{7}", text)
            if epic_match:
                epic = epic_match.group()

            # ---------- WRITE ----------
            ws.append([
                house,
                name,
                age,
                gender,
                epic,
                page_no + 1
            ])

            total_saved += 1

# ---------------- SAVE ----------------
outfile = f"voter_member_wise_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
wb.save(outfile)

print("\n✅ DONE")
print("Total members saved:", total_saved)
print("Excel saved as:", outfile)
