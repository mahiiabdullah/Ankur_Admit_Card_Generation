"""
WhatsApp admit card sender (you press Send yourself - nothing is sent automatically).

For each mobile number, one click:
  1. opens that chat in WhatsApp Desktop with the message already typed
  2. copies that number's admit card PDF(s) to the clipboard

Then in WhatsApp:  Enter (sends the message)  ->  Ctrl+V  ->  Enter (sends the PDF).
Come back here and click "Sent - next".

Progress is saved to whatsapp_send_log.csv, so you can close the app and
continue later from where you stopped. Message text: whatsapp_message.txt
"""
import csv
import os
import subprocess
import time
import tkinter as tk
from datetime import datetime
from tkinter import messagebox
from urllib.parse import quote

import pandas as pd

HERE = os.path.dirname(os.path.abspath(__file__))
XLSX = os.path.join(HERE, "Admit_Card_Data_2026.xlsx")
PDF_DIR = os.path.join(HERE, "Admit_Cards_2026", "PDF")
MESSAGE = os.path.join(HERE, "whatsapp_message.txt")
LOG = os.path.join(HERE, "whatsapp_send_log.csv")
# WhatsApp number currently logged in to WhatsApp Desktop (recorded in the log).
# History: +8801723088207 (Ankur ccs South) restricted 24h on 18/10/2026 after 52 sends; +8801856851751 used
# in between (50 sends); back on +8801723088207 from 21/10/2026; positions 201-241 sent from another
# laptop; this laptop on +8801575808883 from 21/10/2026.
SENDER = "+8801575808883"
BREAK_EVERY = 15  # suggest a break after this many sends in one session
MIN_GAP = 0       # seconds to wait before the next chat can be opened (0 = no wait)

# Rows to double-check before sending (shown in red)
WARNINGS = {
    "3648": "Mobile taken from Nagad number - differs from the one entered. Confirm first.",
    "3506": "Same student as order 3750 (two rolls). Decide which card to send.",
    "3750": "Same student as order 3506 (two rolls). Decide which card to send.",
    "3843": "Likely same student as order 3561 (roll 41126040, already sent to another number). Confirm first.",
}


def load_queue():
    df = pd.read_excel(XLSX, sheet_name="Admit_Data", dtype=str)
    queue = []
    for mobile, grp in df.sort_values("Order ID", key=lambda s: s.astype(int)).groupby("Mobile (Online)", sort=False):
        students = grp.to_dict("records")
        queue.append({
            "mobile": mobile,
            "phone": mobile.lstrip("+"),
            "students": students,
            "pdfs": [os.path.join(PDF_DIR, s["File Name"] + ".pdf") for s in students],
            "warnings": [f"Order {s['Order ID']}: {WARNINGS[s['Order ID']]}" for s in students if s["Order ID"] in WARNINGS],
        })
    return queue


def build_message(item, template):
    lines, centres = [], {s["Exam Centre"] for s in item["students"]}
    for s in item["students"]:
        line = f"• {s['Name (English)']} — Roll: *{s['Exam Roll']}* (Class {s['Class']})"
        if len(centres) > 1:
            line += f"\n   পরীক্ষা কেন্দ্র: {s['Exam Centre']}"
        lines.append(line)
    centre = f"পরীক্ষা কেন্দ্র: {centres.pop()}" if len(centres) == 1 else ""
    return template.format(students="\n".join(lines), centre=centre).strip()


def read_log(queue=None):
    status, sent_rolls = {}, set()
    if os.path.isfile(LOG):
        with open(LOG, encoding="utf-8-sig") as f:
            for row in csv.DictReader(f):
                status[row["Mobile"]] = row["Status"]  # latest entry wins
                if row["Status"] == "sent":
                    sent_rolls.update(row["Rolls"].split(" | "))
    # A number already messaged gets a new card later (e.g. a sibling registers) -> send again
    for item in queue or []:
        if status.get(item["mobile"]) == "sent" and any(s["Exam Roll"] not in sent_rolls for s in item["students"]):
            status[item["mobile"]] = "new card added"
    return status


def write_log(item, status):
    new = not os.path.isfile(LOG)
    with open(LOG, "a", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        if new:
            w.writerow(["Time", "Sender", "Mobile", "Status", "Students", "Rolls"])
        w.writerow([datetime.now().strftime("%Y-%m-%d %H:%M:%S"), SENDER, item["mobile"], status,
                    " | ".join(s["Name (English)"] for s in item["students"]),
                    " | ".join(s["Exam Roll"] for s in item["students"])])


def copy_files_to_clipboard(paths):
    quoted = ",".join("'" + p.replace("'", "''") + "'" for p in paths)
    subprocess.run(["powershell", "-NoProfile", "-Command", f"Set-Clipboard -Path {quoted}"],
                   check=True, creationflags=subprocess.CREATE_NO_WINDOW)


class App:
    def __init__(self, root):
        self.root = root
        self.queue = load_queue()
        self.template = open(MESSAGE, encoding="utf-8").read()
        self.status = read_log(self.queue)
        self.session_sent = 0
        self.i = self._first_pending()

        root.title("Admit Card - WhatsApp Sender")
        root.geometry("640x520")
        root.attributes("-topmost", True)

        self.last_open = 0.0
        tk.Label(root, text=f"Sending from: 0{SENDER[4:8]}-{SENDER[8:]}  (WhatsApp Desktop must be logged in to this number)",
                 font=("Segoe UI", 9, "bold"), fg="#075E54").pack(pady=(6, 0))
        self.progress = tk.Label(root, font=("Segoe UI", 11, "bold"))
        self.progress.pack(pady=(10, 0))
        self.number = tk.Label(root, font=("Segoe UI", 20, "bold"))
        self.number.pack(pady=4)
        self.state = tk.Label(root, font=("Segoe UI", 10, "bold"))
        self.state.pack()
        self.students = tk.Label(root, font=("Segoe UI", 11), justify="left")
        self.students.pack(pady=6)
        self.warn = tk.Label(root, font=("Segoe UI", 10, "bold"), fg="#c00000", wraplength=600, justify="left")
        self.warn.pack()

        self.open_btn = tk.Button(root, text="1.  Open chat + copy PDF", font=("Segoe UI", 13, "bold"),
                                  bg="#25D366", fg="white", width=28, command=self.open_chat)
        self.open_btn.pack(pady=(12, 4))
        tk.Label(root, text="In WhatsApp:  Enter  →  Ctrl+V  →  Enter", font=("Segoe UI", 11)).pack()
        tk.Button(root, text="2.  Sent ✓ - next", font=("Segoe UI", 13, "bold"),
                  bg="#128C7E", fg="white", width=28, command=self.mark_sent).pack(pady=(10, 4))

        row = tk.Frame(root)
        row.pack(pady=8)
        tk.Button(row, text="◀ Back", width=10, command=lambda: self.move(-1)).pack(side="left", padx=4)
        tk.Button(row, text="Skip for now", width=12, command=self.skip).pack(side="left", padx=4)
        tk.Button(row, text="Next ▶", width=10, command=lambda: self.move(1)).pack(side="left", padx=4)
        tk.Button(row, text="Copy PDF again", width=14, command=self.copy_again).pack(side="left", padx=4)

        self.info = tk.Label(root, font=("Segoe UI", 9), fg="#555")
        self.info.pack(side="bottom", pady=8)
        self.show()

    def _first_pending(self):
        for i, item in enumerate(self.queue):
            if self.status.get(item["mobile"]) != "sent":
                return i
        return len(self.queue) - 1

    def show(self):
        item = self.queue[self.i]
        sent = sum(1 for q in self.queue if self.status.get(q["mobile"]) == "sent")
        self.progress.config(text=f"Number {self.i + 1} of {len(self.queue)}     |     Sent: {sent}     |     This session: {self.session_sent}")
        self.number.config(text="0" + item["phone"][3:])
        st = self.status.get(item["mobile"], "not sent")
        self.state.config(text=f"Status: {st}", fg={"sent": "#128C7E", "skipped": "#c07000"}.get(st, "#555"))
        self.students.config(text="\n".join(
            f"{s['Name (English)']}  —  Roll {s['Exam Roll']}  —  Class {s['Class']}" for s in item["students"])
            + f"\n\n{len(item['pdfs'])} PDF: " + ", ".join(os.path.basename(p) for p in item["pdfs"]))
        self.warn.config(text="\n".join("⚠ " + w for w in item["warnings"]))
        self.info.config(text="Progress is saved automatically. You can close and continue later.")

    def _countdown(self):
        left = int(MIN_GAP - (time.time() - self.last_open))
        if left > 0:
            self.open_btn.config(state="disabled", text=f"Wait {left}s (keeps your number safe)", bg="#9e9e9e")
            self.root.after(1000, self._countdown)
        else:
            self.open_btn.config(state="normal", text="1.  Open chat + copy PDF", bg="#25D366")

    def open_chat(self):
        if time.time() - self.last_open < MIN_GAP:
            return
        item = self.queue[self.i]
        missing = [p for p in item["pdfs"] if not os.path.isfile(p)]
        if missing:
            messagebox.showerror("PDF missing", "\n".join(missing))
            return
        if item["warnings"] and not messagebox.askyesno("Check first", "\n\n".join(item["warnings"]) + "\n\nOpen chat anyway?"):
            return
        text = build_message(item, self.template)
        os.startfile(f"whatsapp://send?phone={item['phone']}&text={quote(text)}")
        copy_files_to_clipboard(item["pdfs"])
        self.last_open = time.time()
        self._countdown()
        self.info.config(text="Chat opened and PDF copied. In WhatsApp: Enter → Ctrl+V → Enter")

    def copy_again(self):
        copy_files_to_clipboard(self.queue[self.i]["pdfs"])
        self.info.config(text="PDF copied again - press Ctrl+V in the WhatsApp chat")

    def mark_sent(self):
        item = self.queue[self.i]
        write_log(item, "sent")
        self.status[item["mobile"]] = "sent"
        self.session_sent += 1
        if self.session_sent % BREAK_EVERY == 0:
            messagebox.showinfo("Take a break",
                                f"{self.session_sent} sent this session.\n\nTake a 30-60 minute break to keep "
                                "your WhatsApp account safe (about 60-80 per day in total). Progress is saved.")
        self.next_pending()

    def skip(self):
        item = self.queue[self.i]
        write_log(item, "skipped")
        self.status[item["mobile"]] = "skipped"
        self.next_pending()

    def next_pending(self):
        for j in range(self.i + 1, len(self.queue)):
            if self.status.get(self.queue[j]["mobile"]) != "sent":
                self.i = j
                self.show()
                return
        skipped = sum(1 for q in self.queue if self.status.get(q["mobile"]) == "skipped")
        self.show()
        messagebox.showinfo("Done", f"Reached the end of the list.\nSkipped numbers left: {skipped}\n"
                                    "Use ◀ Back to go to them.")

    def move(self, step):
        self.i = max(0, min(len(self.queue) - 1, self.i + step))
        self.show()


if __name__ == "__main__":
    root = tk.Tk()
    App(root)
    root.mainloop()
