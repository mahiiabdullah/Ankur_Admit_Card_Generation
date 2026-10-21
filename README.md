# Aunkur Scholarship — Admit Card Generator

Makes one PDF admit card per student from an Excel sheet and a Word design,
gives every student an exam roll, and helps send the cards to parents on WhatsApp.

> ⚠️ This repo contains children's names and parents' phone numbers.
> **Keep it private.** Never make it public or share the files outside the team.

---

## 1. Setup on a new laptop (one time)

1. Install **Python 3.11+** and **Git**.
2. Install **WPS Office** or **Microsoft Word** (used to turn the cards into PDF).
3. Install the Bangla fonts **Kalpurush** and **Li Shamim Cholontika Unicode**
   (right-click the .ttf → *Install*). Without them Bangla text breaks in the PDFs.
4. Get the project and install the Python packages:
   ```powershell
   git clone https://github.com/mahiiabdullah/Ankur_Admit_Card_Generation
   cd Ankur_Admit_Card_Generation
   python -m pip install -r requirements.txt
   ```
5. For sending: install **WhatsApp Desktop** and log in with the sending number.

---

## 2. What each file is (2026)

| File / folder | What it is |
|---|---|
| `Student data.xlsx` | Raw registrations (orders) — the input |
| `Demo Admit.docx` | The card design (a filled sample). Never edited by the scripts |
| `Admit_Card_Data_2026.xlsx` | **Cleaned data** — one row per student, with Exam Roll. Main working file |
| `ROLL_NUMBER_RULES.md` | Roll number format + **latest roll used in each group** |
| `Admit_Cards_2026/PDF/` | The finished admit cards, named `Admit_<Roll>.pdf` |
| `Admit_Cards_2026/DOCX/` | Same cards as Word files (for manual edits) |
| `Send_Admit_Cards_WhatsApp.bat` | **Double-click to send cards on WhatsApp** |
| `whatsapp_send_log.csv` | Who was sent / skipped, when, and from which number |
| `whatsapp_message.txt` | The WhatsApp message text (editable) |
| `prepare_admit_data_2026.py` | Step 1 — check & clean the data |
| `assign_rolls_2026.py` | Step 2 — give exam rolls |
| `generate_admit_cards_2026.py` | Step 3 — make the PDF cards |
| `whatsapp_sender.py` | Step 4 — the sending app (started by the .bat file) |

### Sheets inside `Admit_Card_Data_2026.xlsx`
- **Admit_Data** – the students who get a card. Yellow rows need attention (see *Remarks*).
- **Separate** – students held back on purpose (no roll, no card).
- **Review** – every problem found in the raw data.
- **Centre_Map** – area → exam centre.
- **Raw_Data** – untouched copy of the input.

---

## 3. The full process (in order)

Run these in PowerShell inside the project folder.

### Step 1 — Check & clean the data
```powershell
python prepare_admit_data_2026.py
```
- Converts the messy *Class* values (`Six`, `৬ষ্ঠ`, `VIII`, `05(Five)` …) to numbers.
- Formats mobiles as `+8801XXXXXXXXX`. If a mobile is invalid, it uses the bKash, then Nagad number.
- Sets the Exam Centre from the area.
- Marks duplicate students (same name + father + mobile) with `_Duplicate` in the file name.
- **Safe to re-run:** rolls already given are kept.

To fix a student, correct `Student data.xlsx` and run this step again.
(Fixes typed directly into `Admit_Card_Data_2026.xlsx` are overwritten — except Exam Roll.)

### Step 2 — Give exam rolls
```powershell
python assign_rolls_2026.py --dry-run   # preview only
python assign_rolls_2026.py             # assign and save
```
Only students **without** a roll get one, so running it twice never creates duplicates.
It continues from the numbers in `ROLL_NUMBER_RULES.md` and updates that file.

### Step 3 — Make the admit cards
```powershell
python generate_admit_cards_2026.py                  # all cards (rebuilds both folders)
python generate_admit_cards_2026.py --only 3679      # just one student (Order ID or Roll)
```
All 269 cards take about 10 minutes. Close WPS/Word documents first.

### Step 4 — Send on WhatsApp
Double-click **`Send_Admit_Cards_WhatsApp.bat`**, then for each number:

1. Click **“1. Open chat + copy PDF”** → the chat opens with the message typed.
2. In WhatsApp press **Enter** (message) → **Ctrl+V** → **Enter** (PDF).
3. Click **“2. Sent ✓ – next”**.

- Brothers/sisters on the same number get one message with all their cards.
- Progress is saved — close any time; it continues from the first unsent number.
- **Skip for now** puts a number aside; **◀ Back** returns to it.
- Red warning = check before sending (unconfirmed number, duplicate student).

### Step 5 — Save progress to GitHub (when using two laptops)
```powershell
git add whatsapp_send_log.csv
git commit -m "send progress"
git push
```
Before starting on the other laptop: `git pull`.
**Send from only one laptop at a time.**

---

## 4. Exam roll format (short version)

8 digits: **Class · Centre · Gender · 26 · Serial** — e.g. `61226044` = Class 6, Centre 1, Girl, 2026, serial 044.

| Class | 4–9 → same digit, **10 → 1** |
|---|---|
| Centre | 1 = চট্টগ্রাম কলেজিয়েট স্কুল (Dewanhat) · 2 = পতেঙ্গা উচ্চ বিদ্যালয় (Patenga) · 3 = সিলভার বেলস কিন্ডারগার্টেন এন্ড গার্লস হাই স্কুল (Halishahar & Agrabad) |
| Gender | 1 = Boy · 2 = Girl |

Each Class + Centre + Gender group counts separately. Full rules and the latest
number used per group: **`ROLL_NUMBER_RULES.md`** (never reuse or skip a roll).

---

## 5. Common tasks

| I want to… | Do this |
|---|---|
| Add new students | Add them to `Student data.xlsx` → Step 1 → Step 2 → `generate_admit_cards_2026.py --only <OrderIDs>` |
| Fix one student's name/school | Fix in `Student data.xlsx` → Step 1 → Step 3 with `--only <OrderID>` |
| Change one roll | Edit Exam Roll + File Name in `Admit_Card_Data_2026.xlsx`, update the group in `ROLL_NUMBER_RULES.md`, delete the old PDF/DOCX, run Step 3 with `--only` |
| Hold a student back | Add their Order ID to `SEPARATE_ORDERS` in `prepare_admit_data_2026.py` → Step 1 |
| Set a class by hand | Add to `CLASS_OVERRIDES` in `prepare_admit_data_2026.py` → Step 1 |
| Change the WhatsApp text | Edit `whatsapp_message.txt` (keep `{students}` and `{centre}`) |
| Change the sending number | Edit `SENDER = "+880…"` at the top of `whatsapp_sender.py` |
| Change exam date / card design | Edit `Demo Admit.docx` (keep the sample values the same), then Step 3 |

---

## 6. Problems & fixes

| Problem | Fix |
|---|---|
| `Permission denied` when saving the Excel file | It is open in WPS/Excel — close it and run again |
| Bangla letters broken in PDF | Install the Kalpurush and Li Shamim fonts, then regenerate |
| PDF conversion fails / “RPC server unavailable” | The script restarts WPS and retries automatically; if it still fails, close all WPS windows and run again |
| “Can't send messages to new people” on WhatsApp | The number is temporarily restricted (usually 24h). Stop, use another number or wait. Next time send slower and take the break every 15 |
| PDF paste doesn't work in WhatsApp | Click **Copy PDF again**, then Ctrl+V |

### WhatsApp safety
WhatsApp limits accounts that message many unsaved numbers quickly.
Send about **15 at a time, then break**, max ~60–80 per day per number.
Never use “bulk WhatsApp sender” tools — they get numbers banned.
For thousands of messages use bulk SMS with download links or the official WhatsApp Business API.

---

## 7. Status (21 Oct 2026)

- 285 registrations (270 + 15 added on 21/10 from `orders-2026-10-20-22-46-13.xlsx`) → **284 cards** (1 kept separate: order 3625).
- Exam: **23 October 2026, 10:00 AM**.
- WhatsApp: 238 of 255 numbers sent; left = 3 skipped + the 14 numbers of the new students — see `whatsapp_send_log.csv` for the live count.
- To confirm before sending: order 3648 (mobile taken from Nagad), orders 3506 / 3750 (same student, two rolls),
  order 3843 (likely same student as 3561, who already got roll 41126040).
- If a new card is added for a number that was already messaged, the sender shows it again as “new card added”.
- 22/10: added the 14 completed new orders from `orders-2026-10-21-17-40-53.xlsx`
  (3695, 3786, 3834, 3862–3869, 3871–3873); rolls set by hand; Halishahar ones go to Collegiate.
  Orders 3565, 3751, 3798 were “On hold” there and were not added.

---

## 8. Older (2025) scripts

`new_file.py`, `generate_admit_cards_custom.py`, `generate_specific_rolls.py`,
`convert_*.py`, `Convert-DocxToPdf.ps1`, `extract_mobile_numbers.py`, `create_mobile_xml.py`,
`check_*.py`, `verify_*.py` and the `Generated_Admit_Cards_*` folders are from the 2025 batch.
They use hard-coded paths from the old PC and are kept for reference only.
