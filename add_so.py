import fitz
import pandas as pd
import os
import re


# ============================================================
# PATHS
# ============================================================

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

INPUT_FOLDER = os.path.join(BASE_DIR, "Input")
OUTPUT_FOLDER = os.path.join(BASE_DIR, "Output")
EXCEL_PATH = os.path.join(BASE_DIR, "SO_Numbers.xlsx")

INPUT_PDF = os.path.join(
    INPUT_FOLDER,
    "jobbags_20260910172927.pdf"
)

OUTPUT_PDF = os.path.join(
    OUTPUT_FOLDER,
    "jobbags_20260910172927_SO.pdf"
)


# ============================================================
# CHECK OUTPUT FOLDER
# ============================================================

os.makedirs(OUTPUT_FOLDER, exist_ok=True)


# ============================================================
# READ EXCEL
# ============================================================

df = pd.read_excel(EXCEL_PATH)

required_columns = ["JOBBAG_NUMBER", "SO_NUMBER"]

for column in required_columns:
    if column not in df.columns:
        raise Exception(
            f"Excel must contain column: {column}"
        )


# ============================================================
# CONVERT COLUMNS TO STRING
# ============================================================

df["JOBBAG_NUMBER"] = (
    df["JOBBAG_NUMBER"]
    .astype(str)
    .str.replace(".0", "", regex=False)
    .str.strip()
)

df["SO_NUMBER"] = (
    df["SO_NUMBER"]
    .astype(str)
    .str.replace(".0", "", regex=False)
    .str.strip()
)


# ============================================================
# CREATE JOBBAG → SO MAPPING
# ============================================================

so_mapping = dict(
    zip(
        df["JOBBAG_NUMBER"],
        df["SO_NUMBER"]
    )
)


print("Excel mapping loaded:")
print()

for jobbag, so in so_mapping.items():
    print(f"{jobbag}  →  SO {so}")

print()


# ============================================================
# OPEN PDF
# ============================================================

doc = fitz.open(INPUT_PDF)

print(f"PDF pages: {len(doc)}")
print()


# ============================================================
# PROCESS EVERY PAGE
# ============================================================

processed_count = 0
missing_count = 0

for page_number, page in enumerate(doc, start=1):

    print(f"Processing page {page_number}...")

    words = page.get_text("words")

    # --------------------------------------------------------
    # FIND 8-DIGIT JOBBAG NUMBERS
    # --------------------------------------------------------

    jobbags = []

    for word in words:

        x0, y0, x1, y1, text = word[:5]

        text = text.strip()

        # Jobbag number = exactly 8 digits
        if re.fullmatch(r"\d{8}", text):

            # Jobbag numbers are near the top of each section
            if y0 < 100:

                jobbags.append({
                    "jobbag": text,
                    "x": x0,
                    "y": y0
                })


    # --------------------------------------------------------
    # PROCESS EACH JOBBAG
    # --------------------------------------------------------

    for item in jobbags:

        jobbag = item["jobbag"]

        x = item["x"]
        y = item["y"]

        print(
            f"  Found Jobbag: {jobbag}"
        )


        # ----------------------------------------------------
        # FIND CORRESPONDING SO NUMBER
        # ----------------------------------------------------

        if jobbag not in so_mapping:

            print(
                f"    WARNING: SO not found for {jobbag}"
            )

            missing_count += 1

            continue


        so_number = so_mapping[jobbag]

        print(
            f"    SO: {so_number}"
        )


        # ----------------------------------------------------
        # SO NUMBER POSITION
        # ----------------------------------------------------
        #
        # Shifted further RIGHT
        # Slightly adjusted vertically
        # Larger area for the SO number
        #

        box_x = x + 108
        box_y = y + 12

        box_width = 65
        box_height = 22


        # ----------------------------------------------------
        # CREATE SO TEXT BOX
        # ----------------------------------------------------

        so_rect = fitz.Rect(
            box_x,
            box_y,
            box_x + box_width,
            box_y + box_height
        )


        # ----------------------------------------------------
        # INSERT SO NUMBER
        # ----------------------------------------------------

        page.insert_textbox(
            so_rect,
            so_number,
            fontsize=10,
            fontname="helv",
            color=(0, 0, 0),
            align=1
        )


        processed_count += 1


# ============================================================
# SAVE OUTPUT PDF
# ============================================================

doc.save(OUTPUT_PDF)

doc.close()


# ============================================================
# FINAL RESULT
# ============================================================

print()
print("========================================")
print("PROCESS COMPLETED")
print("========================================")
print(f"SO numbers added : {processed_count}")
print(f"Missing SO       : {missing_count}")
print()
print("Output:")
print(OUTPUT_PDF)