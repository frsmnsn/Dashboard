# -*- coding: utf-8 -*-
"""
dq_dashboard_app.py - หน้าต่างโปรแกรมสำหรับทีม ไม่ต้องแตะโค้ดหรือ Terminal

เปิดด้วยการดับเบิลคลิก "เปิดโปรแกรม.bat" (แนะนำ) หรือรัน:
    pythonw dq_dashboard_app.py     (ไม่โชว์หน้าต่างดำ)
    python dq_dashboard_app.py      (โชว์หน้าต่างดำไว้ดู error ได้)

ไฟล์นี้ต้องอยู่โฟลเดอร์เดียวกับ dq_dashboard.py เพราะดึงฟังก์ชันมาใช้โดยตรง
(ไม่ได้เขียนตรรกะซ้ำ - ใช้ตัวอ่าน/วิเคราะห์/เทมเพลต HTML เดียวกับตอนรันบรรทัดคำสั่ง)
"""
import io
import json
import os
import sys
import threading
import traceback
import webbrowser
from datetime import datetime
from tkinter import Tk, StringVar, filedialog, messagebox
from tkinter import ttk, scrolledtext

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import dq_dashboard as dq  # noqa: E402  (ใช้ตัวอ่าน/วิเคราะห์/เทมเพลตเดียวกับ CLI)

HERE = os.path.dirname(os.path.abspath(__file__))
BG = "#eef3f0"
GREEN = "#0b4a2a"
PINK = "#c2185b"


def build_dashboard_file(it_path, control_path=None, log=print):
    """ทำตามขั้นตอนเดียวกับ dq_dashboard.py main(): อ่านไฟล์ -> วิเคราะห์ -> เขียน HTML
    คืนค่า path ของไฟล์ผลลัพธ์"""
    log(f"อ่านไฟล์: {os.path.basename(it_path)}")
    data, _used = dq.load_all(it_path)
    for code, df in data.items():
        n = 0 if df is None else len(df)
        log(f"  {code}: {n:,} แถว" if df is not None else f"  {code}: ไม่พบชีตนี้ (ข้าม)")

    ctrl = dq.load_control(control_path) if control_path else None
    allsn = [s for k in data if data[k] is not None and "snapshot" in data[k]
             for s in data[k].snapshot.dropna().unique()]
    latest = max(allsn) if allsn else None
    meta = {
        "title": "รายงานการประเมินและตรวจสอบคุณภาพข้อมูล",
        "demo": False,
        "source": os.path.basename(it_path),
        "snapshot": dq.th_date(dq.pd.Timestamp(latest)) if latest is not None else "-",
        "snapshots": [dq.th_date(dq.pd.Timestamp(s), True) for s in sorted(set(allsn))],
        "generated": dq.th_date(dq.pd.Timestamp(datetime.now())),
    }
    log("กำลังวิเคราะห์และสร้างหน้า Dashboard ...")
    payload = dq.san(dq.build(data, ctrl, meta))
    js = json.dumps(payload, ensure_ascii=False, allow_nan=False).replace("</", "<\\/")
    html = dq.HTML.replace("__DATA__", js)

    stamp = latest.strftime("%Y-%m-%d") if latest is not None else datetime.now().strftime("%Y-%m-%d")
    out_dir = os.path.join(HERE, "ผลลัพธ์ Dashboard")
    os.makedirs(out_dir, exist_ok=True)
    dated = os.path.join(out_dir, f"dashboard_{stamp}.html")
    latest_copy = os.path.join(out_dir, "dashboard_ล่าสุด.html")
    with open(dated, "w", encoding="utf-8") as f:
        f.write(html)
    with open(latest_copy, "w", encoding="utf-8") as f:
        f.write(html)
    log(f"เสร็จแล้ว -> {os.path.basename(dated)}")
    return latest_copy


class App:
    def __init__(self, root):
        self.root = root
        root.title("Dashboard คุณภาพข้อมูล — บสส.")
        root.geometry("720x560")
        root.configure(bg=BG)
        root.minsize(620, 480)

        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure("TFrame", background=BG)
        style.configure("Head.TLabel", background=BG, foreground=GREEN, font=("Tahoma", 17, "bold"))
        style.configure("Sub.TLabel", background=BG, foreground="#61726a", font=("Tahoma", 10))
        style.configure("Card.TLabelframe", background="white", relief="flat")
        style.configure("Card.TLabelframe.Label", background="white", foreground=GREEN, font=("Tahoma", 11, "bold"))
        style.configure("TButton", font=("Tahoma", 10), padding=8)
        style.map("Accent.TButton",
                  background=[("!disabled", GREEN)], foreground=[("!disabled", "white")])
        style.configure("Accent.TButton", font=("Tahoma", 11, "bold"), padding=10)

        outer = ttk.Frame(root, padding=18)
        outer.pack(fill="both", expand=True)

        ttk.Label(outer, text="Dashboard คุณภาพข้อมูล", style="Head.TLabel").pack(anchor="w")
        ttk.Label(outer, text="เลือกไฟล์ผลตรวจจาก IT แล้วกดปุ่ม ไม่ต้องแตะโค้ดหรือ Terminal",
                  style="Sub.TLabel").pack(anchor="w", pady=(2, 14))

        card1 = ttk.Labelframe(outer, text="  ขั้นตอนหลัก  ", style="Card.TLabelframe", padding=14)
        card1.pack(fill="x", pady=(0, 12))
        self.file_var = StringVar(value="ยังไม่ได้เลือกไฟล์")
        row = ttk.Frame(card1, style="TFrame")
        row.pack(fill="x")
        ttk.Button(row, text="1) เลือกไฟล์ผลตรวจจาก IT (.xlsx)", command=self.pick_file).pack(side="left")
        ttk.Label(row, textvariable=self.file_var, style="Sub.TLabel").pack(side="left", padx=10)
        ttk.Button(card1, text="2) สร้าง Dashboard และเปิดดูผล", style="Accent.TButton",
                  command=self.run_build).pack(fill="x", pady=(10, 0))

        card2 = ttk.Labelframe(outer, text="  B1 / B2 (เกณฑ์ที่ต้องตรวจเอง) — ไม่บังคับ  ",
                               style="Card.TLabelframe", padding=14)
        card2.pack(fill="x", pady=(0, 12))
        r2 = ttk.Frame(card2, style="TFrame")
        r2.pack(fill="x")
        ttk.Button(r2, text="สร้างไฟล์เทมเพลตให้กรอก", command=self.make_template).pack(side="left")
        ttk.Button(r2, text="รวมผลที่กรอกแล้ว + สร้าง Dashboard", command=self.run_with_control
                  ).pack(side="left", padx=8)

        ttk.Label(outer, text="สถานะ", style="Sub.TLabel").pack(anchor="w")
        self.log = scrolledtext.ScrolledText(outer, height=12, font=("Consolas", 9),
                                             bg="#0b1f14", fg="#d7f5e3", insertbackground="white")
        self.log.pack(fill="both", expand=True, pady=(4, 0))
        self.log.configure(state="disabled")

        self.it_path = None
        self._log("พร้อมใช้งาน — เริ่มจากขั้นตอนที่ 1")

    # ---------- ตัวช่วย ----------
    def _log(self, msg):
        self.log.configure(state="normal")
        self.log.insert("end", msg + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")
        self.root.update_idletasks()

    def pick_file(self):
        p = filedialog.askopenfilename(title="เลือกไฟล์ผลตรวจจาก IT",
                                        filetypes=[("Excel files", "*.xlsx")])
        if p:
            self.it_path = p
            self.file_var.set(os.path.basename(p))
            self._log(f"เลือกไฟล์แล้ว: {os.path.basename(p)}")

    def _need_file(self):
        if not self.it_path:
            messagebox.showwarning("ยังไม่ได้เลือกไฟล์", "กรุณากดปุ่ม 1) เลือกไฟล์ผลตรวจจาก IT ก่อน")
            return False
        return True

    def _run_async(self, fn):
        threading.Thread(target=self._safe(fn), daemon=True).start()

    def _safe(self, fn):
        def wrap():
            try:
                fn()
            except Exception:
                err = traceback.format_exc()
                self._log("เกิดข้อผิดพลาด:\n" + err)
                self.root.after(0, lambda: messagebox.showerror(
                    "เกิดข้อผิดพลาด",
                    "โปรแกรมทำงานไม่สำเร็จ\nคัดลอกข้อความในช่องสถานะ ส่งให้ทีมพัฒนาเพื่อแก้ไข"))
        return wrap

    # ---------- การทำงานหลัก ----------
    def run_build(self):
        if not self._need_file():
            return
        def job():
            out = build_dashboard_file(self.it_path, log=self._log)
            self._log("กำลังเปิด Dashboard ในเบราว์เซอร์ ...")
            webbrowser.open("file://" + os.path.abspath(out))
        self._run_async(job)

    def make_template(self):
        if not self._need_file():
            return
        def job():
            self._log("กำลังอ่านไฟล์เพื่อสร้างเทมเพลต B1/B2 ...")
            data, _ = dq.load_all(self.it_path)
            out_dir = os.path.join(HERE, "ผลลัพธ์ Dashboard")
            os.makedirs(out_dir, exist_ok=True)
            out = os.path.join(out_dir, "b_control_template.xlsx")
            dq.make_template(data, out)
            self._log(f"สร้างเทมเพลตแล้ว -> {out}")
            self.root.after(0, lambda: messagebox.showinfo(
                "สร้างเทมเพลตสำเร็จ",
                f"ไฟล์อยู่ที่:\n{out}\n\nกรอกคอลัมน์ ref_count / ref_value แล้วกลับมากดปุ่ม\n"
                "'รวมผลที่กรอกแล้ว + สร้าง Dashboard'"))
        self._run_async(job)

    def run_with_control(self):
        if not self._need_file():
            return
        ctrl_path = filedialog.askopenfilename(
            title="เลือกไฟล์ b_control_template.xlsx ที่กรอกแล้ว",
            filetypes=[("Excel files", "*.xlsx")])
        if not ctrl_path:
            return
        def job():
            out = build_dashboard_file(self.it_path, control_path=ctrl_path, log=self._log)
            self._log("กำลังเปิด Dashboard ในเบราว์เซอร์ ...")
            webbrowser.open("file://" + os.path.abspath(out))
        self._run_async(job)


def main():
    root = Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()