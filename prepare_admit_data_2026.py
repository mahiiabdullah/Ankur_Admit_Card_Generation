"""
Validate and clean 'Student data.xlsx' (2026 registrations) and export
'Admit_Card_Data_2026.xlsx' with the fields used by the 'Demo Admit.docx' card.

Sheets written:
  Admit_Data  - one row per order, cleaned card fields (+ Remarks for problems)
  Review      - every issue found, one line per issue
  Centre_Map  - area -> exam centre name used
  Raw_Data    - untouched copy of the source sheet
"""
import argparse
import os
import re

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

CENTRE_MAP = {
    "dewanhat": "চট্টগ্রাম কলেজিয়েট স্কুল",
    "halishahar": "সিলভার বেলস কিন্ডারগার্টেন এন্ড গার্লস হাই স্কুল",
    "agrabad": "সিলভার বেলস কিন্ডারগার্টেন এন্ড গার্লস হাই স্কুল",
    "patenga": "পতেঙ্গা উচ্চ বিদ্যালয়",
}

# Orders held out of the card batch (moved to the 'Separate' sheet), with reason
SEPARATE_ORDERS = {
    3625: "Kept separate on request (class not given - email entered instead)",
}

# Likely re-registrations the automatic check misses (different mobile).
# new order -> earlier order. Only the new order is marked '_Duplicate'
# (the earlier card was already sent, so its name stays the same).
MANUAL_DUPLICATES = {
    3843: 3561,  # Safwan Zaman Ahnaf - same name, father, class, centre; other mobile
}

# Student names corrected by the user (registration had the wrong name)
NAME_OVERRIDES = {
    3578: "Sarah Nur Shawli",  # registered as "Hossain Mohiuddin" (father's name); roll 51226058
}

# Classes confirmed by the user for rows where the entered class was unusable
CLASS_OVERRIDES = {
    3748: 6,  # Tamiha Sarkar - email entered as class
}

BANGLA_DIGITS = str.maketrans("০১২৩৪৫৬৭৮৯", "0123456789")

# Checked in order; first match wins. Bangla patterns match after digit translation.
CLASS_PATTERNS = [
    (10, r"\bten\b|^10$|\b10\b|দশম"),
    (9, r"\bnine\b|\bix\b|\b0?9\b|নবম"),
    (8, r"\beight\b|\bviii\b|\b0?8\b|অষ্টম|8ম"),
    (7, r"\bseven\b|\bsaven\b|\bvii\b|\b0?7\b|সপ্তম|7ম|সেভেন"),
    (6, r"\bsix\b|\bvi\b|\b0?6\b|ষষ্ঠ|6ষ্ঠ"),
    (5, r"\bfive\b|\bv\b|\b0?5\b|পঞ্চম|5ম"),
    (4, r"\bfour\b|\biv\b|\b0?4\b|চতুর্থ|চর্তুথ|4র্থ|4থ"),
]

BANGLA_RE = re.compile(r"[ঀ-৿]")


def clean_text(value) -> str:
    if pd.isna(value):
        return ""
    s = str(value).replace(" ", " ")
    s = re.sub(r"\s+", " ", s).strip()
    # Drop stray Bangla combining marks typed before Latin text (e.g. '্Md Tareq')
    s = re.sub(r"^[়-ৗ]+(?=[A-Za-z])", "", s)
    return s


def normalize_class(value):
    s = clean_text(value).lower().translate(BANGLA_DIGITS)
    if not s or "@" in s:
        return None
    for cls, pattern in CLASS_PATTERNS:
        if re.search(pattern, s):
            return cls
    return None


def normalize_mobile(value):
    """Return (+8801XXXXXXXXX, None) or (original, problem)."""
    raw = clean_text(value)
    all_digits = re.sub(r"\D", "", raw)
    digits = all_digits[2:] if all_digits.startswith("88") else all_digits
    if re.fullmatch(r"01[3-9]\d{8}", digits):
        return "+88" + digits, None
    return raw, f"invalid mobile ({len(all_digits)} digits)"


def main(src: str, out: str) -> None:
    raw = pd.read_excel(src, dtype=str)
    issues = []

    def flag(order_id, field, value, problem):
        issues.append({"Order ID": order_id, "Field": field, "Value": value, "Problem": problem})

    rows = []
    for _, r in raw.iterrows():
        oid = clean_text(r["Order ID"])
        remarks = []

        if int(oid) in NAME_OVERRIDES:
            flag(oid, "Name (English)", r["Name (English)"], f"name corrected to {NAME_OVERRIDES[int(oid)]} (by user)")

        cls = normalize_class(r["Class"])
        if int(oid) in CLASS_OVERRIDES:
            cls = CLASS_OVERRIDES[int(oid)]
            flag(oid, "Class", r["Class"], f"class set to {cls} (confirmed by user)")
        elif cls is None:
            flag(oid, "Class", r["Class"], "class could not be read - fill in manually")
            remarks.append("Class missing")

        mobile, problem = normalize_mobile(r["Mobile (Online)"])
        if problem:
            # Fall back to the bKash, then Nagad, payment number
            for field in ("Bkash number", "Nagad Number"):
                alt, alt_problem = normalize_mobile(r[field])
                if not alt_problem:
                    entered = re.sub(r"\D", "", clean_text(r["Mobile (Online)"]))
                    similar = sum(a == b for a, b in zip(entered.lstrip("1"), alt[3:])) >= 8
                    note = f"{problem}; replaced with {field} {alt}"
                    if similar:
                        remarks.append(f"Mobile taken from {field}")
                    else:
                        note += " - different number, please confirm"
                        remarks.append(f"Mobile taken from {field} (different number, confirm)")
                    flag(oid, "Mobile (Online)", r["Mobile (Online)"], note)
                    mobile = alt
                    break
            else:
                flag(oid, "Mobile (Online)", r["Mobile (Online)"], problem)
                remarks.append("Check mobile")

        area = clean_text(r["Exam Centre Last Added"]).lower()
        centre = CENTRE_MAP.get(area, "")
        if not centre:
            flag(oid, "Exam Centre Last Added", r["Exam Centre Last Added"], "unknown area")
            remarks.append("Centre missing")

        for field in ("Fathers Name", "Name of School"):
            if BANGLA_RE.search(clean_text(r[field])):
                flag(oid, field, r[field], "written in Bangla (kept as entered)")

        for field in ("Name (English)", "Fathers Name", "Name of School"):
            if not clean_text(r[field]):
                flag(oid, field, "", "empty")
                remarks.append(f"{field} empty")

        rows.append({
            "Order ID": int(oid),
            "Exam Roll": "",
            "Gender": clean_text(r["Gender"]),
            "Name (English)": NAME_OVERRIDES.get(int(oid), clean_text(r["Name (English)"])),
            "Fathers Name": clean_text(r["Fathers Name"]),
            "Name of School": clean_text(r["Name of School"]),
            "Class": cls if cls is not None else "",
            "Mobile (Online)": mobile,
            "Area": area.capitalize(),
            "Exam Centre": centre,
            "Class (as entered)": clean_text(r["Class"]),
            "Remarks": "",
            "_remarks": remarks,
        })

    df = pd.DataFrame(rows)

    # Duplicate students: same English name + father + mobile. Kept, but marked.
    key = (df["Name (English)"].str.lower().str.replace(r"[^a-z]", "", regex=True)
           + "|" + df["Fathers Name"].str.lower().str.replace(r"\W", "", regex=True)
           + "|" + df["Mobile (Online)"])
    for _, grp in df[key.duplicated(keep=False)].groupby(key):
        ids = grp["Order ID"].tolist()
        for i in grp.index:
            others = ", ".join(str(x) for x in ids if x != df.at[i, "Order ID"])
            df.at[i, "_remarks"].append(f"Duplicate of order {others}")
            flag(df.at[i, "Order ID"], "Student", df.at[i, "Name (English)"],
                 f"same student as order {others} - kept both, marked in file name")

    for new_id, old_id in MANUAL_DUPLICATES.items():
        for i in df.index[df["Order ID"] == new_id]:
            df.at[i, "_remarks"].append(f"Duplicate of order {old_id} (same name & father, different mobile)")
            flag(new_id, "Student", df.at[i, "Name (English)"],
                 f"likely same student as order {old_id} - kept, marked in file name")

    # Keep rolls already issued (by assign_rolls_2026.py or edited by hand) - they must never change.
    # The roll's 2nd digit fixes the centre, so a student with a roll also keeps their Exam Centre
    # (hand edits moved some Halishahar/Agrabad students to Collegiate).
    if os.path.isfile(out):
        prev = pd.read_excel(out, sheet_name="Admit_Data", dtype=str)
        prev = prev.dropna(subset=["Exam Roll"])
        issued = dict(zip(prev["Order ID"].astype(int), prev["Exam Roll"]))
        centre = dict(zip(prev["Order ID"].astype(int), prev["Exam Centre"]))
        df["Exam Roll"] = df["Order ID"].map(issued).fillna("")
        has_roll = df["Exam Roll"] != ""
        df.loc[has_roll, "Exam Centre"] = df.loc[has_roll, "Order ID"].map(centre)

    df["File Name"] = df.apply(
        lambda r: f"Admit_{r['Exam Roll'] or r['Order ID']}"
        + ("_Duplicate" if any("Duplicate" in x for x in r["_remarks"]) else ""),
        axis=1)
    df["Remarks"] = df["_remarks"].map("; ".join)
    df = df.drop(columns="_remarks").sort_values("Order ID")
    df = df[["Order ID", "Exam Roll", "File Name", "Gender", "Name (English)", "Fathers Name",
             "Name of School", "Class", "Mobile (Online)", "Area", "Exam Centre",
             "Class (as entered)", "Remarks"]]

    held = df["Order ID"].isin(SEPARATE_ORDERS)
    separate = df[held].copy()
    separate.insert(1, "Reason", separate["Order ID"].map(SEPARATE_ORDERS))
    df = df[~held]

    review = pd.DataFrame(issues, columns=["Order ID", "Field", "Value", "Problem"])
    review["Order ID"] = review["Order ID"].astype(int)
    review = review.sort_values(["Problem", "Order ID"])
    centre_map = pd.DataFrame(
        [{"Area": a.capitalize(), "Exam Centre": c, "Students": int((df["Area"] == a.capitalize()).sum())}
         for a, c in CENTRE_MAP.items()])

    with pd.ExcelWriter(out, engine="openpyxl") as xw:
        df.to_excel(xw, sheet_name="Admit_Data", index=False)
        separate.to_excel(xw, sheet_name="Separate", index=False)
        review.to_excel(xw, sheet_name="Review", index=False)
        centre_map.to_excel(xw, sheet_name="Centre_Map", index=False)
        raw.to_excel(xw, sheet_name="Raw_Data", index=False)

    # Formatting: bold header, freeze, filter, widths, highlight rows needing attention
    wb = load_workbook(out)
    yellow = PatternFill("solid", fgColor="FFF2CC")
    for ws in wb.worksheets:
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        for c in ws[1]:
            c.font = Font(bold=True)
        for col in ws.columns:
            width = max(len(str(c.value or "")) for c in col)
            ws.column_dimensions[get_column_letter(col[0].column)].width = min(max(width + 2, 10), 60)
    ws = wb["Admit_Data"]
    remarks_col = [c.value for c in ws[1]].index("Remarks") + 1
    for row in ws.iter_rows(min_row=2):
        if row[remarks_col - 1].value:
            for c in row:
                c.fill = yellow
    wb.save(out)

    print(f"Rows exported: {len(df)}  ->  {out}")
    print(f"Kept separate: {len(separate)} ({', '.join(separate['Name (English)'])})")
    print(f"Rows needing attention: {(df['Remarks'] != '').sum()}")
    print(review.groupby("Problem").size().to_string())
    print("\nClass counts:\n" + df["Class"].astype(str).value_counts().sort_index().to_string())
    print("\nCentre counts:\n" + df["Exam Centre"].value_counts().to_string())


if __name__ == "__main__":
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("--data", default="Student data.xlsx")
    p.add_argument("--out", default="Admit_Card_Data_2026.xlsx")
    a = p.parse_args()
    main(a.data, a.out)
