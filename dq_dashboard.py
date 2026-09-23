#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
dq_dashboard.py  -  ใส่ไฟล์ผลตรวจจาก IT แล้วได้ Dashboard report (HTML ไฟล์เดียว)

วิธีใช้
  python dq_dashboard.py ผลตรวจ_IT.xlsx                          -> dashboard.html
  python dq_dashboard.py ผลตรวจ_IT.xlsx --template               -> b_control_template.xlsx (ไว้กรอก B1/B2)
  python dq_dashboard.py ผลตรวจ_IT.xlsx --control b_control.xlsx -> รวมผลตรวจ B1/B2 ที่ตรวจเอง
  python dq_dashboard.py --demo                                  -> dashboard_demo.html (ข้อมูลสมมติ)

ต้องใช้แค่  pip install pandas openpyxl   (ไม่ต้องใช้อินเทอร์เน็ต ไฟล์ผลลัพธ์เปิดในเบราว์เซอร์ได้เลย)

เกณฑ์ตัดสิน (threshold) ทั้งหมดอยู่ใน CONFIG ด้านล่าง ปรับให้ตรงกับนิยามของ บสส. ได้
"""
import argparse
import json
import math
import os
import re
import sys
import tempfile
from collections import Counter
from datetime import date, datetime

import numpy as np
import pandas as pd

# =====================================================================
# CONFIG  (ค่าเริ่มต้นเป็นข้อเสนอ ควรยืนยันกับพี่เลี้ยง/ทีม IT)
# =====================================================================
CONFIG = {
    # ชื่อชีตที่จะหา (ลองตามลำดับ ไม่สนตัวพิมพ์เล็ก/ใหญ่)
    "sheets": {
        "A1": ["A1 Final", "A1"],
        "A2": ["A2", "A2 Final"],
        "A3": ["A3 Final", "A3"],
        "C1": ["C1 Final", "C1"],       # C1 ดิบมีทุก key อาจเป็นล้านแถว จึงลอง Final ก่อน
        "C2": ["C2"],
        "C3": ["C3 Final1", "C3"],
        "C4": ["C4"],
        "D1": ["D1"],
        "E1": ["E1"],
    },
    "a1_min_missing": 1.0,   # A1: ตัวแปรถือว่า "มีปัญหา" เมื่อ %missing >= ค่านี้ (1.0 = 100%)
    "c2_dominant": 0.95,     # C2: ค่าเดียวครอบคลุม >= 95% ถือว่ามีปัญหา
    "c3_high_missing": 0.70, # C3: missing สูงเมื่อ >= 70% (ใช้เมื่อชีตไม่มี Flag_HighMissing)
    "c4_missing": 0.05,      # C4: แสดง Missing > 5% เป็นข้อมูลประกอบ
    "e1_min_change": 0.01,   # E1: เปลี่ยนแปลง >= 1% ของแถว ถือว่าควรสังเกต
    "b1_tol": 0.001,         # B1: ยอมให้ต่างได้ 0.1%
    "b2_tol": 0.0,           # B2: จำนวนรายการต้องเท่ากัน
    "max_rows": 2500,        # จำนวนแถวสูงสุดต่อตารางรายละเอียด (กันไฟล์ใหญ่เกิน)
    # เทียบกับรอบก่อน/ที่ปรึกษา (ไม่บังคับ) เช่น {"A1": {"tables": 27, "items": 116}}
    "baseline": {},
}
CFG = CONFIG

CHECKS = [
    ("A1", "ความครบถ้วน", "การตรวจสอบการว่างของข้อมูลในแต่ละ Snapshot",
     "ตรวจสอบว่าไม่มีข้อมูลที่เว้นว่างในบาง Snapshot และข้อมูลที่ได้รับนั้นมีความคงที่ตามช่วงเวลา"),
    ("A2", "ความครบถ้วน", "การตรวจสอบ Primary Key",
     "ตรวจสอบว่าข้อมูลที่ได้รับไม่มีรายการซ้ำกันในแต่ละ Primary Key"),
    ("A3", "ความครบถ้วน", "การตรวจสอบระยะเวลาที่มีข้อมูลของแต่ละฟิลด์",
     "ตรวจสอบว่าตัวแปรแต่ละตัวมีข้อมูลเริ่มต้นที่ Snapshot ใด เพราะบางฟิลด์เพิ่มเติมภายหลังทำให้ไม่มีข้อมูลในอดีต"),
    ("B1", "ความถูกต้อง", "การกระทบยอดค่าของข้อมูล",
     "เปรียบเทียบค่าข้อมูลใน Data Model กับข้อมูลหรือรายงานของ บสส. ณ Snapshot เดียวกัน (ต้องตรวจเอง)"),
    ("B2", "ความถูกต้อง", "การกระทบจำนวนรายการข้อมูล",
     "เปรียบเทียบจำนวนรายการข้อมูลกับข้อมูลหรือรายงานของ บสส. ณ Snapshot เดียวกัน (ต้องตรวจเอง)"),
    ("C1", "ความสมเหตุสมผล", "การตรวจสอบ Primary Key ในแต่ละช่วงเวลา",
     "ตรวจสอบว่าแต่ละ Primary Key มีข้อมูลหายไปในบาง Snapshot หรือไม่"),
    ("C2", "ความสมเหตุสมผล", "การตรวจสอบ Categorical variables",
     "ตรวจสอบค่าทั้งหมดของข้อมูล Categorical ว่ามีค่าตามที่คาดการณ์ไว้หรือมีค่าว่างหรือไม่"),
    ("C3", "ความสมเหตุสมผล", "การตรวจสอบ Continuous variables",
     "ตรวจสอบการกระจายตัวของข้อมูลเพื่อหาค่าผิดปกติ ค่าที่ไม่สมเหตุสมผล หรือค่าว่าง"),
    ("C4", "ความสมเหตุสมผล", "การตรวจสอบ Date variables",
     "ตรวจสอบค่าวันที่ว่ามีค่าตามที่คาดการณ์ไว้หรือไม่ และเป็นปี พ.ศ. หรือ ค.ศ. ที่ถูกต้อง"),
    ("D1", "ความสอดคล้อง", "การตรวจสอบความสอดคล้องกันระหว่างฟิลด์ข้อมูล",
     "ตรวจสอบว่าค่าของข้อมูลในแต่ละฟิลด์เป็นไปตามเงื่อนไขและสอดคล้องกับฟิลด์อื่นที่ควรมีความสัมพันธ์กัน"),
    ("E1", "การเปลี่ยนแปลง", "การตรวจสอบการเปลี่ยนแปลงของข้อมูลในแต่ละช่วงเวลา",
     "ตรวจสอบว่าค่าของข้อมูลในแต่ละฟิลด์มีการเปลี่ยนแปลงหรือไม่เมื่อเทียบกับ Snapshot ก่อนหน้า"),
]
TAGTONE = {"Critical": "bad", "High": "warn", "Partial": "info", "Complete": "good",
           "เริ่มมีข้อมูลภายหลัง": "warn", "ข้อมูลหายในล่าสุด": "bad",
           "ไม่มีข้อมูลเลย": "mut", "ปกติ": "good", "ตรง": "good", "ไม่ตรง": "bad"}
TH_M = ["ม.ค.", "ก.พ.", "มี.ค.", "เม.ย.", "พ.ค.", "มิ.ย.", "ก.ค.", "ส.ค.", "ก.ย.", "ต.ค.", "พ.ย.", "ธ.ค."]


# =====================================================================
# ตัวช่วยทั่วไป
# =====================================================================
def th_date(ts, short=False):
    if short:
        return f"{ts.day} {TH_M[ts.month - 1]} {(ts.year + 543) % 100:02d}"
    return f"{ts.day} {TH_M[ts.month - 1]} {ts.year + 543}"


def clean(df):
    df = df.copy()
    df.columns = [c if isinstance(c, (datetime, date)) else re.sub(r"\s+", " ", str(c)).strip()
                  for c in df.columns]
    df = df.loc[:, [not (isinstance(c, str) and c.startswith("Unnamed")) for c in df.columns]]
    return df.dropna(how="all").dropna(axis=1, how="all")


def find(df, *prefixes, required=True):
    for p in prefixes:
        for c in df.columns:
            if isinstance(c, str) and c.strip().lower().startswith(p.lower()):
                return c
    if required:
        raise KeyError(f"ไม่พบคอลัมน์ {prefixes} ใน {list(df.columns)[:14]}")
    return None


def to_ts(x):
    """แปลงเป็นวันที่ (ค.ศ.) ปี พ.ศ. (>2400) จะลบ 543 ให้ ถ้าแปลงไม่ได้คืน NaT"""
    try:
        if isinstance(x, (datetime, date)):
            if pd.isna(x):
                return pd.NaT
            y = x.year - 543 if x.year > 2400 else x.year
            return pd.Timestamp(y, x.month, x.day)
        if isinstance(x, str):
            s = x.strip()
            m = re.match(r"^(\d{1,2})[/-](\d{1,2})[/-](\d{4})", s)
            if m:
                d, mo, y = map(int, m.groups())
                return pd.Timestamp(y - 543 if y > 2400 else y, mo, d)
            m = re.match(r"^(\d{4})[/-](\d{1,2})[/-](\d{1,2})", s)
            if m:
                y, mo, d = map(int, m.groups())
                return pd.Timestamp(y - 543 if y > 2400 else y, mo, d)
    except Exception:
        pass
    return pd.NaT


def raw_year(x):
    """ปีตามที่เก็บจริง (ไม่แก้ พ.ศ.) ใช้ตรวจ BE error"""
    if isinstance(x, (datetime, date)) and not pd.isna(x):
        return x.year
    if isinstance(x, str):
        m = re.search(r"(\d{4})", x)
        return int(m.group(1)) if m else None
    return None


def frac(s):
    s = pd.to_numeric(s, errors="coerce")
    return s.where(s.isna() | (s <= 1), s / 100)


def num(s):
    return pd.to_numeric(s, errors="coerce")


def san(o):
    if o is pd.NaT or o is None:
        return None
    if isinstance(o, dict):
        return {str(k): san(v) for k, v in o.items()}
    if isinstance(o, (list, tuple, set)):
        return [san(v) for v in o]
    if isinstance(o, (bool, np.bool_)):
        return bool(o)
    if isinstance(o, (int, np.integer)):
        return int(o)
    if isinstance(o, (float, np.floating)):
        return None if (math.isnan(o) or math.isinf(o)) else float(o)
    if isinstance(o, (pd.Timestamp, datetime, date)):
        return str(o)[:10]
    return o


# =====================================================================
# ตัวอ่านแต่ละชีต -> ตารางมาตรฐาน
# =====================================================================
def wide_or_long(df, value_name, long_col):
    """ชีตที่วันที่เป็นหัวคอลัมน์ (Final) หรือมีคอลัมน์ snapshot (ดิบ)"""
    df = clean(df)
    port, var = find(df, "port"), find(df, "variable")
    dc = [c for c in df.columns if c not in (port, var) and pd.notna(to_ts(c))]
    if dc:
        out = df.melt(id_vars=[port, var], value_vars=dc, var_name="snapshot", value_name=value_name)
    else:
        out = df[[port, var, find(df, "snapshot"), find(df, long_col)]].copy()
    out.columns = ["port", "variable", "snapshot", value_name]
    out["port"] = out["port"].astype(str)
    out["snapshot"] = out["snapshot"].map(to_ts)
    out[value_name] = frac(out[value_name])
    return out.dropna(subset=["snapshot"])


def ld_A1(df):
    return wide_or_long(df, "miss", "pct_miss")


def ld_A3(df):
    return wide_or_long(df, "exist", "%variable")


def ld_A2(df):
    df = clean(df)
    dup = find(df, "n_obs_dup", required=False)
    if dup:
        out = pd.DataFrame({"port": df[find(df, "port")].astype(str),
                            "snapshot": df[find(df, "snapshot")].map(to_ts),
                            "n_obs": num(df[find(df, "n_obs")]), "dup": num(df[dup])})
    else:  # A2 Final: ชุดข้อมูล x วันที่ (จำนวนซ้ำ)
        idc = df.columns[0]
        dc = [c for c in df.columns if c != idc and pd.notna(to_ts(c))]
        out = df.melt(id_vars=[idc], value_vars=dc, var_name="snapshot", value_name="dup")
        out.columns = ["port", "snapshot", "dup"]
        out["port"] = out["port"].astype(str)
        out["snapshot"] = out["snapshot"].map(to_ts)
        out["n_obs"] = np.nan
        out["dup"] = num(out["dup"])
    return out.dropna(subset=["snapshot"])


def ld_C1(df):
    df = clean(df)
    port = find(df, "port")
    miss = find(df, "n_missing", required=False)
    if miss:  # ดิบ: 1 แถว = 1 key
        g = pd.DataFrame({"port": df[port].astype(str), "gap": num(df[miss]) > 0})
        out = g.groupby("port")["gap"].agg(["size", "sum"]).reset_index()
        out.columns = ["port", "n_keys", "n_gap"]
        out["pct"] = out["n_gap"] / out["n_keys"]
    else:     # Final: PORT, %discontinue
        out = pd.DataFrame({"port": df[port].astype(str),
                            "pct": frac(df[find(df, "%disc", "pct_disc")]),
                            "n_keys": np.nan, "n_gap": np.nan})
    return out


def ld_C2(df):
    df = clean(df)
    val = df[find(df, "description")].astype(object).where(df[find(df, "description")].notna(), "(ว่าง)")
    return pd.DataFrame({"port": df[find(df, "port")].astype(str), "variable": df[find(df, "variable")].astype(str),
                         "value": val.astype(str), "n_obs": num(df[find(df, "n_obs")]),
                         "share": frac(df[find(df, "pct_obs")])})


def ld_C3(df):
    df = clean(df)
    out = pd.DataFrame({"port": df[find(df, "port")].astype(str), "variable": df[find(df, "variable")].astype(str),
                        "snapshot": df[find(df, "snapshot")].map(to_ts), "n_obs": num(df[find(df, "n_obs")]),
                        "pct_miss": frac(df[find(df, "pct_miss")])})
    for k in ("mean", "min", "max"):
        c = find(df, k, required=False)
        out[k] = num(df[c]) if c else np.nan
    for c in df.columns:
        if isinstance(c, str) and c.lower().startswith("flag_"):
            out["flag_" + re.sub(r"[^a-z0-9]", "", c[5:].lower())] = num(df[c])
    return out.dropna(subset=["snapshot"])


def ld_C4(df):
    df = clean(df)
    snap = df[find(df, "snapshot")].map(to_ts)
    out = pd.DataFrame({"port": df[find(df, "port")].astype(str), "variable": df[find(df, "variable")].astype(str),
                        "snapshot": snap, "n_obs": num(df[find(df, "n_obs")]),
                        "pct_miss": frac(df[find(df, "pct_miss")])})
    nu = find(df, "n_obs_unique", required=False)
    out["n_unique"] = num(df[nu]) if nu else np.nan
    dcols = [c for c in (find(df, "min", required=False), find(df, "max", required=False),
                         find(df, "percentile_5", required=False), find(df, "percentile_95", required=False)) if c]
    be = pd.Series(False, index=df.index)
    for c in dcols:
        be = be | df[c].map(lambda v: (raw_year(v) or 0) > 2400)
    out["be_error"] = be.values
    mx = find(df, "max", required=False)
    maxdt = df[mx].map(to_ts) if mx else pd.Series(pd.NaT, index=df.index)
    out["max_dt"] = maxdt.values
    out["future"] = (maxdt.values > snap.values) & pd.notna(maxdt.values)
    out["const"] = (out["n_unique"] <= 1) & (out["pct_miss"] < 1)
    return out.dropna(subset=["snapshot"])


def ld_D1(df):
    df = clean(df)
    return pd.DataFrame({"port": df[find(df, "port")].astype(str), "rule": df[find(df, "rules")].astype(str),
                         "snapshot": df[find(df, "as_of")].map(to_ts), "n": num(df[find(df, "n_obs")]),
                         "pct": frac(df[find(df, "pct_obs")])}).dropna(subset=["snapshot"])


def ld_E1(df):
    df = clean(df)
    return pd.DataFrame({"port": df[find(df, "port")].astype(str), "variable": df[find(df, "variable")].astype(str),
                         "snapshot": df[find(df, "snapshot")].map(to_ts), "n_obs": num(df[find(df, "n_obs")]),
                         "n_diff": num(df[find(df, "n_diff")]),
                         "pct_diff": frac(df[find(df, "pct_diff")])}).dropna(subset=["snapshot"])


LOADERS = {"A1": ld_A1, "A2": ld_A2, "A3": ld_A3, "C1": ld_C1, "C2": ld_C2,
           "C3": ld_C3, "C4": ld_C4, "D1": ld_D1, "E1": ld_E1}


def load_all(path):
    xl = pd.ExcelFile(path)
    names = {n.strip().lower(): n for n in xl.sheet_names}
    data, used = {}, {}
    for code, loader in LOADERS.items():
        nm = next((names[c.lower()] for c in CFG["sheets"][code] if c.lower() in names), None)
        if nm is None:
            print(f"  {code}: ไม่พบชีต {CFG['sheets'][code]} (ข้าม)")
            data[code] = None
            continue
        try:
            data[code] = loader(xl.parse(nm))
            used[code] = nm
            print(f"  {code}: อ่านชีต '{nm}' ได้ {len(data[code]):,} แถว")
        except Exception as e:
            print(f"  {code}: อ่านชีต '{nm}' ไม่สำเร็จ -> {e}")
            data[code] = None
    return data, used


def load_control(path):
    xl = pd.ExcelFile(path)
    out = {}
    for code in ("B1", "B2"):
        nm = next((s for s in xl.sheet_names if s.strip().upper() == code), None)
        if not nm:
            continue
        df = clean(xl.parse(nm))
        ref, dq = find(df, "ref", required=False), find(df, "dq", required=False)
        if not ref or not dq:
            continue
        sn, vr = find(df, "snapshot", required=False), find(df, "variable", required=False)
        o = pd.DataFrame({"port": df[find(df, "port")].astype(str),
                          "snapshot": df[sn].map(to_ts) if sn else pd.NaT,
                          "variable": df[vr].astype(str) if vr else "",
                          "dq": num(df[dq]), "ref": num(df[ref])}).dropna(subset=["dq", "ref"])
        if len(o):
            out[code] = o
    return out


# =====================================================================
# ตัวช่วยสร้างผลวิเคราะห์
# =====================================================================
def K(label, value, tone="info", fmt="n", sub=None):
    return {"label": label, "value": value, "tone": tone, "fmt": fmt, "sub": sub}


def Col(label, typ="text"):
    return {"label": label, "type": typ}


def R(code, status="ok", kpis=None, cols=None, rows=None, charts=None, notes=None,
      by_port=None, ports=None, fvars=None):
    return dict(code=code, status=status, kpis=kpis or [], cols=cols or [], rows=rows or [],
                charts=charts or [], notes=notes or [], by_port={k: int(v) for k, v in (by_port or {}).items()},
                ports=sorted(ports or []), fvars=fvars or set())


def bars(title, series, n=12, fmt="n", tone="info", sub=None):
    s = series[series > 0].sort_values(ascending=False).head(n)
    return {"t": "bars", "title": title, "labels": [str(i) for i in s.index], "values": s.tolist(),
            "fmt": fmt, "tone": tone, "sub": sub}


def donut(title, parts):
    return {"t": "donut", "title": title, "parts": [{"l": l, "v": int(v), "tone": t} for l, v, t in parts]}


def trend(title, labels, values, fmt="n"):
    return {"t": "trend", "title": title, "labels": labels, "values": values, "fmt": fmt}


def share(a, b):
    return f"{(a / b * 100):.1f}%" if b else None


def cap(rows):
    return rows[:CFG["max_rows"]]


def snaps_of(df):
    return sorted(df["snapshot"].dropna().unique())


def sl(ts):
    return th_date(pd.Timestamp(ts), short=True)


# =====================================================================
# วิเคราะห์แต่ละเกณฑ์
# =====================================================================
def an_A1(a):
    sn = snaps_of(a)
    L = a[a.snapshot == sn[-1]].copy()
    L["level"] = np.select([L.miss <= 0, L.miss < 0.7, L.miss < 1], ["Complete", "Partial", "High"], "Critical")
    bad = L[L.miss >= CFG["a1_min_missing"] - 1e-9]
    by = bad.groupby("port").size()
    cnt = L.level.value_counts().to_dict()
    n = len(L)
    rows = [[r.port, r.variable, r.miss, r.level] for r in
            L[L.miss > 0].sort_values(["miss", "port", "variable"], ascending=[False, True, True]).itertuples()]
    return R("A1", kpis=[K("ตัวแปรที่ตรวจ", n), K("Critical (missing 100%)", cnt.get("Critical", 0), "bad", sub=share(cnt.get("Critical", 0), n)),
                         K("High (70–99%)", cnt.get("High", 0), "warn", sub=share(cnt.get("High", 0), n)),
                         K("Complete (0%)", cnt.get("Complete", 0), "good", sub=share(cnt.get("Complete", 0), n))],
             cols=[Col("ตาราง"), Col("ตัวแปร"), Col("% Missing", "pct"), Col("ระดับ", "tag")], rows=cap(rows),
             charts=[donut("สัดส่วนตัวแปรตามระดับ Missing", [("Complete", cnt.get("Complete", 0), "good"), ("Partial", cnt.get("Partial", 0), "info"),
                                                            ("High", cnt.get("High", 0), "warn"), ("Critical", cnt.get("Critical", 0), "bad")]),
                     bars("ตัวแปรที่ Missing 100% แยกตามตาราง", by, tone="bad"),
                     trend("% Missing เฉลี่ยของทุกตัวแปร", [sl(s) for s in sn], [float(a[a.snapshot == s].miss.mean()) for s in sn], "p")],
             notes=["เกณฑ์ระดับ: Complete = 0% | Partial < 70% | High < 100% | Critical = 100% (ใช้ Snapshot ล่าสุด)",
                    "ตารางที่ถือว่า 'พบปัญหา' คือมีตัวแปร missing ตั้งแต่ " + f"{CFG['a1_min_missing'] * 100:.0f}% ขึ้นไป (ปรับได้ที่ a1_min_missing)"],
             by_port=by.to_dict(), ports=set(L.port), fvars=set(zip(bad.port, bad.variable)))


def an_A2(a):
    sn = snaps_of(a)
    L = a[a.snapshot == sn[-1]].groupby("port").agg(n_obs=("n_obs", "sum"), dup=("dup", "sum")).reset_index()
    has = L[L.dup > 0]
    tot_obs = L.n_obs.sum()
    rows = [[r.port, r.n_obs if r.n_obs == r.n_obs else None, r.dup, (r.dup / r.n_obs) if r.n_obs else None]
            for r in L.sort_values("dup", ascending=False).itertuples()]
    return R("A2", kpis=[K("ตารางที่ตรวจ", len(L)), K("ตารางที่มีรายการซ้ำ", len(has), "bad" if len(has) else "good", sub=share(len(has), len(L))),
                         K("รายการซ้ำรวม", int(L.dup.sum()), "warn"),
                         K("อัตราซ้ำโดยรวม", (L.dup.sum() / tot_obs) if tot_obs else None, "info", "p")],
             cols=[Col("ตาราง"), Col("จำนวนรายการ", "num"), Col("รายการซ้ำ", "num"), Col("% ซ้ำ", "pct")], rows=cap(rows),
             charts=[donut("ตารางที่มี/ไม่มีรายการซ้ำ", [("ไม่มีรายการซ้ำ", len(L) - len(has), "good"), ("มีรายการซ้ำ", len(has), "bad")]),
                     bars("จำนวนรายการซ้ำแยกตามตาราง", L.set_index("port").dup, tone="bad"),
                     trend("รายการซ้ำรวมในแต่ละ Snapshot", [sl(s) for s in sn], [float(a[a.snapshot == s].dup.sum()) for s in sn])],
             notes=["นับเฉพาะ Snapshot ล่าสุด ไม่บวกรวมทุกเดือน (ไม่เช่นนั้นรายการซ้ำเดิมจะถูกนับซ้ำ 3 เท่า)"],
             by_port=L.set_index("port").dup.to_dict(), ports=set(L.port))


def an_A3(a):
    sn = snaps_of(a)
    pv = a.pivot_table(index=["port", "variable"], columns="snapshot", values="exist", aggfunc="first")
    f, l = pv[sn[0]].fillna(0), pv[sn[-1]].fillna(0)
    allzero = (pv.fillna(0) <= 0).all(axis=1)
    cat = np.select([allzero, (f <= 0) & (l > 0), (f > 0) & (l <= 0)], ["ไม่มีข้อมูลเลย", "เริ่มมีข้อมูลภายหลัง", "ข้อมูลหายในล่าสุด"], "ปกติ")
    t = pd.DataFrame({"f": f, "l": l, "cat": cat}).reset_index()
    issue = t[t.cat.isin(["เริ่มมีข้อมูลภายหลัง", "ข้อมูลหายในล่าสุด"])]
    cnt = t.cat.value_counts().to_dict()
    order = {"ข้อมูลหายในล่าสุด": 0, "เริ่มมีข้อมูลภายหลัง": 1, "ไม่มีข้อมูลเลย": 2}
    rows = [[r.port, r.variable, r.f, r.l, r.cat] for r in
            t[t.cat != "ปกติ"].assign(o=lambda x: x.cat.map(order)).sort_values(["o", "port", "variable"]).itertuples()]
    by = issue.groupby("port").size()
    return R("A3", kpis=[K("ตัวแปรที่ตรวจ", len(t)), K("ปกติ", cnt.get("ปกติ", 0), "good"),
                         K("เริ่มมีข้อมูลภายหลัง", cnt.get("เริ่มมีข้อมูลภายหลัง", 0), "warn"),
                         K("ไม่มีข้อมูลเลย", cnt.get("ไม่มีข้อมูลเลย", 0), "info")],
             cols=[Col("ตาราง"), Col("ตัวแปร"), Col("% มีข้อมูล Snapshot แรก", "pct"), Col("% มีข้อมูล Snapshot ล่าสุด", "pct"), Col("สถานะ", "tag")],
             rows=cap(rows),
             charts=[donut("สถานะช่วงเวลาที่มีข้อมูลของตัวแปร", [("ปกติ", cnt.get("ปกติ", 0), "good"), ("เริ่มมีข้อมูลภายหลัง", cnt.get("เริ่มมีข้อมูลภายหลัง", 0), "warn"),
                                                             ("ข้อมูลหายในล่าสุด", cnt.get("ข้อมูลหายในล่าสุด", 0), "bad"), ("ไม่มีข้อมูลเลย", cnt.get("ไม่มีข้อมูลเลย", 0), "mut")]),
                     bars("ตัวแปรที่เริ่มมีข้อมูลภายหลัง/หายไป แยกตามตาราง", by, tone="warn")],
             notes=["A3 = % ของแถวที่มีข้อมูลในแต่ละ Snapshot ตัวแปรที่ 'ไม่มีข้อมูลเลย' ถูกนับไว้ที่ A1 แล้ว จึงไม่นับเป็นปัญหาซ้ำใน A3"],
             by_port=by.to_dict(), ports=set(t.port), fvars=set(zip(issue.port, issue.variable)))


def an_C1(c):
    fl = c[c.pct > 0]
    items = {r.port: (int(r.n_gap) if r.n_gap == r.n_gap else 1) for r in fl.itertuples()}
    rows = [[r.port, r.n_keys if r.n_keys == r.n_keys else None, r.n_gap if r.n_gap == r.n_gap else None, r.pct]
            for r in c.sort_values("pct", ascending=False).itertuples()]
    return R("C1", kpis=[K("ตารางที่ตรวจ", len(c)), K("ตารางที่ Key หายบาง Snapshot", len(fl), "bad" if len(fl) else "good", sub=share(len(fl), len(c))),
                         K("% Key หายสูงสุด", float(c.pct.max()) if len(c) else None, "warn", "p")],
             cols=[Col("ตาราง"), Col("จำนวน Key", "num"), Col("Key ที่หาย", "num"), Col("% Discontinue", "pct")], rows=cap(rows),
             charts=[donut("ตารางที่ Key ต่อเนื่อง/ไม่ต่อเนื่อง", [("ต่อเนื่องครบ", len(c) - len(fl), "good"), ("มี Key หาย", len(fl), "bad")]),
                     bars("% Key ที่หายไปแยกตามตาราง", c.set_index("port").pct, fmt="p", tone="bad")],
             by_port=items, ports=set(c.port))


def an_C2(c):
    thr = CFG["c2_dominant"]
    bad = c[c.share >= thr - 1e-9]
    by = bad.groupby("port").size()
    band = pd.cut(c.share, [-1, .5, .8, thr - 1e-9, 2], labels=["< 50%", "50–79%", f"80–{thr * 100 - 1:.0f}%", f"≥ {thr * 100:.0f}%"]).value_counts()
    rows = [[r.port, r.variable, r.value, r.share] for r in c[c.share >= .8].sort_values("share", ascending=False).itertuples()]
    return R("C2", kpis=[K("ตัวแปร Categorical", len(c)), K("ค่าเดียว ≥ 80%", int((c.share >= .8).sum()), "warn", sub=share((c.share >= .8).sum(), len(c))),
                         K(f"ค่าเดียว ≥ {thr * 100:.0f}%", len(bad), "bad", sub=share(len(bad), len(c)))],
             cols=[Col("ตาราง"), Col("ตัวแปร"), Col("ค่าที่พบมากสุด"), Col("สัดส่วน", "pct")], rows=cap(rows),
             charts=[donut("สัดส่วนของ 'ค่าที่พบมากสุด' ต่อตัวแปร", [(str(k), band.get(k, 0), t) for k, t in zip(band.index, ["good", "info", "warn", "bad"][:len(band)])]),
                     bars("ตัวแปรที่ค่าเดียวเกือบทั้งหมด แยกตามตาราง", by, tone="bad"),
                     bars("ค่าที่พบบ่อยในกลุ่มนี้ (นับตัวแปร)", bad.groupby("value").size(), n=10, tone="info")],
             notes=["ตัวแปรที่มีค่าเดียวเกือบทั้งหมด อาจไม่มีประโยชน์ต่อโมเดล หรือสะท้อนว่าข้อมูลไม่ถูกบันทึกจริง"],
             by_port=by.to_dict(), ports=set(c.port), fvars=set(zip(bad.port, bad.variable)))


FLAG_TH = {"negative": "ค่าติดลบ", "outlier": "Outlier", "skew": "เบ้ (Skew)", "highmissing": "Missing สูง"}


def an_C3(c):
    sn = snaps_of(c)
    L = c[c.snapshot == sn[-1]].reset_index(drop=True)
    fl = {}
    for col in [x for x in L.columns if x.startswith("flag_")]:
        fl[col[5:]] = L[col].fillna(0) > 0
    if not any("neg" in k for k in fl):
        fl["negative"] = (L["min"] < 0).fillna(False)
    if not any("miss" in k for k in fl):
        fl["highmissing"] = (L["pct_miss"] >= CFG["c3_high_missing"]).fillna(False)
    anyf = pd.concat(list(fl.values()), axis=1).any(axis=1)
    bad = L[anyf]
    by = bad.groupby("port").size()
    lab = lambda k: FLAG_TH.get(k, k)
    rows = []
    for i in bad.index:
        r = L.loc[i]
        rows.append([r.port, r.variable, r["mean"], r["min"], r["max"], r.pct_miss, ", ".join(lab(k) for k in fl if fl[k][i])])
    rows.sort(key=lambda x: (-(x[5] if x[5] == x[5] else 0), x[0], x[1]))
    ks = [K("ตัวแปร Continuous", len(L))] + [K(lab(k), int(v.sum()), "bad" if k != "skew" else "warn", sub=share(v.sum(), len(L))) for k, v in fl.items()]
    notes = ["ใช้ Snapshot ล่าสุด | Flag มาจากผลตรวจของ IT ถ้าไม่มีคอลัมน์ Flag ระบบจะคำนวณเฉพาะ 'ค่าติดลบ' และ 'Missing สูง' ให้"]
    if "outlier" in fl and fl["outlier"].mean() > 0.5:
        notes.append(f"ข้อสังเกต: Outlier ถูก flag {fl['outlier'].mean() * 100:.0f}% ของตัวแปร สูงผิดปกติ อาจเกิดจากข้อมูลมีค่า 0 จำนวนมาก (percentile = 0) ควรยืนยันนิยามกับ IT ก่อนสรุป")
    return R("C3", kpis=ks[:6], cols=[Col("ตาราง"), Col("ตัวแปร"), Col("ค่าเฉลี่ย", "num"), Col("ต่ำสุด", "num"), Col("สูงสุด", "num"),
                                     Col("% Missing", "pct"), Col("ประเด็นที่พบ")],
             rows=cap(rows), charts=[bars("จำนวนตัวแปรแยกตามประเภทประเด็น", pd.Series({lab(k): int(v.sum()) for k, v in fl.items()}), tone="warn"),
                                     bars("ตัวแปรที่พบประเด็น แยกตามตาราง", by, tone="bad")],
             notes=notes, by_port=by.to_dict(), ports=set(L.port), fvars=set(zip(bad.port, bad.variable)))


def an_C4(c):
    sn = snaps_of(c)
    L = c[c.snapshot == sn[-1]].reset_index(drop=True)
    anyi = L.be_error | L.future | L.const
    bad = L[anyi]
    by = bad.groupby("port").size()
    rows = []
    for r in bad.itertuples():
        tags = [t for t, f in (("ปี พ.ศ./ค.ศ. ผิด", r.be_error), ("วันที่เกิน Snapshot", r.future), ("วันที่เดียวทั้งคอลัมน์", r.const)) if f]
        rows.append([r.port, r.variable, r.pct_miss, str(r.max_dt)[:10] if pd.notna(r.max_dt) else None, ", ".join(tags)])
    rows.sort(key=lambda x: (x[0], x[1]))
    m5 = int((L.pct_miss > CFG["c4_missing"]).sum())
    return R("C4", kpis=[K("ตัวแปรวันที่", len(L)), K("Missing > 5%", m5, "info", sub=share(m5, len(L))),
                         K("วันที่เกิน Snapshot", int(L.future.sum()), "warn"), K("ปี พ.ศ./ค.ศ. ผิด", int(L.be_error.sum()), "bad"),
                         K("วันที่เดียวทั้งคอลัมน์", int(L.const.sum()), "warn")],
             cols=[Col("ตาราง"), Col("ตัวแปร"), Col("% Missing", "pct"), Col("วันที่สูงสุด"), Col("ประเด็นที่พบ")], rows=cap(rows),
             charts=[bars("จำนวนตัวแปรแยกตามประเภทประเด็น", pd.Series({"ปี พ.ศ./ค.ศ. ผิด": int(L.be_error.sum()), "วันที่เกิน Snapshot": int(L.future.sum()),
                                                                    "วันที่เดียวทั้งคอลัมน์": int(L.const.sum())}), tone="warn"),
                     bars("ตัวแปรที่พบประเด็น แยกตามตาราง", by, tone="bad")],
             notes=["ประเด็นที่นับเป็นปัญหา: ปีผิด (ปี > 2400 ในค่าต่ำสุด/สูงสุด/percentile) วันที่สูงสุดเกิน Snapshot และค่าวันที่เดียวทั้งคอลัมน์ | Missing > 5% แสดงเป็นข้อมูลประกอบ"],
             by_port=by.to_dict(), ports=set(L.port), fvars=set(zip(bad.port, bad.variable)))


def an_D1(d):
    sn = snaps_of(d)
    L = d[d.snapshot == sn[-1]]
    v = L[L.n > 0]
    by = v.groupby("port").size()
    rows = [[r.port, r.rule, r.n, r.pct] for r in v.sort_values("n", ascending=False).itertuples()]
    top = v.sort_values("n", ascending=False).head(10)
    return R("D1", kpis=[K("กฎที่ตรวจ", len(L)), K("กฎที่พบรายการผิด", len(v), "bad" if len(v) else "good", sub=share(len(v), len(L))),
                         K("รายการที่ผิดรวม", int(v.n.sum()), "warn"), K("% ผิดสูงสุด", float(v.pct.max()) if len(v) else None, "bad", "p")],
             cols=[Col("ตาราง"), Col("กฎที่ตรวจ"), Col("รายการที่ผิด", "num"), Col("% ของทั้งหมด", "pct")], rows=cap(rows),
             charts=[bars("กฎที่พบรายการผิดมากที่สุด", pd.Series(top.n.values, index=[f"{p} · {r[:44]}" for p, r in zip(top.port, top.rule)]), n=10, tone="bad"),
                     bars("จำนวนกฎที่ผิดแยกตามตาราง", by, tone="warn"),
                     trend("รายการที่ผิดรวมในแต่ละ Snapshot", [sl(s) for s in sn], [float(d[d.snapshot == s].n.sum()) for s in sn])],
             by_port=by.to_dict(), ports=set(L.port))


def an_E1(e):
    sn = snaps_of(e)
    L = e[e.snapshot == sn[-1]]
    thr = CFG["e1_min_change"]
    ch = L[L.n_diff > 0]
    big = L[L.pct_diff >= thr - 1e-12]
    by = big.groupby("port").size()
    rows = [[r.port, r.variable, r.n_diff, r.pct_diff] for r in ch.sort_values("pct_diff", ascending=False).itertuples()]
    top = ch.sort_values("pct_diff", ascending=False).head(10)
    return R("E1", kpis=[K("ตัวแปรที่เทียบ", len(L)), K("มีการเปลี่ยนแปลง (> 0)", len(ch), "info", sub=share(len(ch), len(L))),
                         K(f"เปลี่ยน ≥ {thr * 100:.0f}% ของแถว", len(big), "warn"), K("เปลี่ยนสูงสุด", float(L.pct_diff.max()) if len(L) else None, "bad", "p")],
             cols=[Col("ตาราง"), Col("ตัวแปร"), Col("แถวที่เปลี่ยน", "num"), Col("% เปลี่ยน", "pct")], rows=cap(rows),
             charts=[bars("ตัวแปรที่เปลี่ยนแปลงมากที่สุด (% ของแถว)", pd.Series(top.pct_diff.values, index=[f"{p} · {v}" for p, v in zip(top.port, top.variable)]), n=10, fmt="p", tone="warn"),
                     trend("จำนวนตัวแปรที่มีการเปลี่ยนแปลงในแต่ละ Snapshot", [sl(s) for s in sn], [float((e[e.snapshot == s].n_diff > 0).sum()) for s in sn])],
             notes=[f"เทียบ Snapshot ล่าสุดกับก่อนหน้า | นับเป็นประเด็นเมื่อเปลี่ยนตั้งแต่ {thr * 100:.0f}% ของแถว (ปรับได้ที่ e1_min_change) เพราะการเปลี่ยนแปลงบางส่วนเป็นเรื่องปกติของระบบ"],
             by_port=by.to_dict(), ports=set(L.port), fvars=set(zip(big.port, big.variable)))


PENDING_NOTES = {
    "B1": ["B1/B2 ต้องตรวจเอง เพราะ IT ไม่ได้รันให้ (ต้องเทียบกับข้อมูลหรือรายงานของ บสส. เช่น ไฟล์ EMFI ที่ใช้พัฒนาแบบจำลอง ECL)",
           "วิธีทำ: 1) รัน  python dq_dashboard.py ไฟล์IT.xlsx --template   2) กรอกคอลัมน์ ref_value (ค่าจากรายงานอ้างอิง) ในชีต B1",
           "3) รัน  python dq_dashboard.py ไฟล์IT.xlsx --control b_control_template.xlsx   หน้านี้จะแสดงผลกระทบยอดให้อัตโนมัติ"],
    "B2": ["B2 เทียบจำนวนรายการ (เช่น จำนวนบัญชี ลูกหนี้ ทรัพย์หลักประกัน) กับรายงานอ้างอิง ณ Snapshot เดียวกัน",
           "ช่อง dq_count ในเทมเพลตจะเติมจำนวนรายการจากผลตรวจของ IT (A2) ให้แล้ว เหลือกรอก ref_count จากรายงานอ้างอิง"],
}


def an_B(code, ctrl):
    if ctrl is None or ctrl.empty:
        return R(code, "pending", notes=PENDING_NOTES[code])
    df = ctrl.copy()
    df["diff"] = df.dq - df.ref
    df["pct"] = np.where(df.ref != 0, df["diff"] / df.ref.abs(), np.where(df["diff"] == 0, 0.0, 1.0))
    tol = CFG["b1_tol"] if code == "B1" else CFG["b2_tol"]
    df["ok"] = df.pct.abs() <= tol + 1e-12
    bad = df[~df.ok]
    by = bad.groupby("port").size()
    stt = lambda o: "ตรง" if o else "ไม่ตรง"
    if code == "B2":
        cols = [Col("ตาราง"), Col("Snapshot"), Col("จำนวนจาก IT", "num"), Col("จำนวนอ้างอิง", "num"), Col("ผลต่าง", "num"), Col("% ต่าง", "pct"), Col("ผล", "tag")]
        rows = [[r.port, th_date(r.snapshot) if pd.notna(r.snapshot) else "", r.dq, r.ref, r.diff, r.pct, stt(r.ok)] for r in df.sort_values("pct", key=abs, ascending=False).itertuples()]
    else:
        cols = [Col("ตาราง"), Col("ตัวแปร"), Col("Snapshot"), Col("ค่าจาก IT", "num"), Col("ค่าอ้างอิง", "num"), Col("ผลต่าง", "num"), Col("% ต่าง", "pct"), Col("ผล", "tag")]
        rows = [[r.port, r.variable, th_date(r.snapshot) if pd.notna(r.snapshot) else "", r.dq, r.ref, r.diff, r.pct, stt(r.ok)]
                for r in df.sort_values("pct", key=abs, ascending=False).itertuples()]
    return R(code, kpis=[K("รายการที่กระทบ", len(df)), K("ตรง", int(df.ok.sum()), "good"), K("ไม่ตรง", len(bad), "bad" if len(bad) else "good")],
             cols=cols, rows=cap(rows),
             charts=[donut("ผลการกระทบ", [("ตรง", int(df.ok.sum()), "good"), ("ไม่ตรง", len(bad), "bad")]),
                     bars("รายการที่ไม่ตรง แยกตามตาราง", by, tone="bad")],
             notes=[f"เกณฑ์ยอมรับ: ต่างไม่เกิน {(CFG['b1_tol'] if code == 'B1' else CFG['b2_tol']) * 100:.2f}% (ปรับได้ที่ {code.lower()}_tol) | ตรวจเฉพาะรายการที่กรอกค่าอ้างอิงแล้ว"],
             by_port=by.to_dict(), ports=set(df.port) | set(bad.port))


ANALYZERS = {"A1": an_A1, "A2": an_A2, "A3": an_A3, "C1": an_C1, "C2": an_C2, "C3": an_C3, "C4": an_C4, "D1": an_D1, "E1": an_E1}


# =====================================================================
# ประกอบข้อมูลทั้งหมดสำหรับหน้าเว็บ
# =====================================================================
def build(data, ctrl, meta):
    res = {}
    for code, dim, title, desc in CHECKS:
        if code in ("B1", "B2"):
            r = an_B(code, (ctrl or {}).get(code))
        elif data.get(code) is None or len(data[code]) == 0:
            r = R(code, "nodata", notes=[f"ไม่พบข้อมูลของเกณฑ์นี้ในไฟล์ (ชีต {CFG['sheets'][code]})"])
        else:
            r = ANALYZERS[code](data[code])
        r.update(dim=dim, title=title, desc=desc)
        res[code] = r
    ports = sorted(set().union(*[set(r["ports"]) for r in res.values()]))
    matrix, score = {}, {}
    for p in ports:
        matrix[p] = {}
        ok = iss = 0
        for code in res:
            r = res[code]
            if r["status"] == "pending":
                cell = {"s": "pending", "n": 0}
            elif r["status"] == "nodata" or p not in r["ports"]:
                cell = {"s": "na", "n": 0}
            else:
                n = r["by_port"].get(p, 0)
                cell = {"s": "issue" if n > 0 else "ok", "n": n}
                ok += n == 0
                iss += n > 0
            matrix[p][code] = cell
        score[p] = round(100 * ok / (ok + iss), 1) if ok + iss else 100.0
    active = [c for c in res if res[c]["status"] == "ok"]
    for c in res:
        r = res[c]
        r["n_tables"] = sum(1 for p in r["ports"] if r["by_port"].get(p, 0) > 0) if r["status"] == "ok" else None
        r["n_items"] = sum(r["by_port"].values()) if r["status"] == "ok" else None
        b = CFG["baseline"].get(c)
        r["delta"] = ({"tables": r["n_tables"] - b["tables"], "items": r["n_items"] - b["items"]} if b and r["status"] == "ok" else None)
    cells_ok = sum(1 for p in ports for c in res if matrix[p][c]["s"] == "ok")
    cells_iss = sum(1 for p in ports for c in res if matrix[p][c]["s"] == "issue")
    overall = 100 * cells_ok / (cells_ok + cells_iss) if cells_ok + cells_iss else 100.0
    n_vars = len(data["A1"][data["A1"].snapshot == data["A1"].snapshot.max()]) if data.get("A1") is not None else None
    any_issue = [p for p in ports if any(matrix[p][c]["s"] == "issue" for c in res)]
    issue_checks = [c for c in active if res[c]["n_tables"]]
    worst = sorted(ports, key=lambda p: (score[p], -sum(matrix[p][c]["n"] for c in res)))[:5]

    summary = []
    summary.append(f"พบประเด็นใน {len(issue_checks)} จาก {len(active)} เกณฑ์ที่ตรวจแล้ว และมี {len(any_issue)} จาก {len(ports)} ตารางที่พบอย่างน้อย 1 ประเด็น")
    byt = sorted([c for c in active if res[c]["n_tables"]], key=lambda c: -res[c]["n_tables"])
    if byt:
        s = f"จำนวนตารางที่พบสูงสุด: เกณฑ์ {byt[0]} ({res[byt[0]]['n_tables']} ตาราง)"
        if len(byt) > 1:
            s += f" รองลงมา เกณฑ์ {byt[1]} ({res[byt[1]]['n_tables']} ตาราง)"
        summary.append(s)
    byi = sorted([c for c in active if res[c]["n_items"]], key=lambda c: -res[c]["n_items"])
    if byi:
        s = f"จำนวนรายการสูงสุด: เกณฑ์ {byi[0]} ({res[byi[0]]['n_items']:,} รายการ)"
        if len(byi) > 1:
            s += f" รองลงมา เกณฑ์ {byi[1]} ({res[byi[1]]['n_items']:,} รายการ)"
        summary.append(s)
    if worst:
        summary.append("ตารางที่ควรดูก่อน (คะแนนต่ำสุด): " + ", ".join(worst[:3]))
    pend = [c for c in res if res[c]["status"] == "pending"]
    if pend:
        summary.append(f"เกณฑ์ {', '.join(pend)} ยังรอผลตรวจที่ทำเอง จึงยังไม่นับรวมในคะแนน")
    for c in active:
        d = res[c]["delta"]
        if d and (d["tables"] or d["items"]):
            summary.append(f"เกณฑ์ {c} เทียบกับ baseline: ตาราง {d['tables']:+d} รายการ {d['items']:+d}")

    cnt = Counter()
    for c in ("A1", "A3", "C2", "C3", "C4", "E1"):
        for v in res[c]["fvars"]:
            cnt[v] += 1
    tv = []
    for v, n in cnt.most_common():
        if n >= 2:
            tv.append({"port": v[0], "variable": v[1], "n": n,
                       "codes": [c for c in ("A1", "A3", "C2", "C3", "C4", "E1") if v in res[c]["fvars"]]})
    tv = sorted(tv, key=lambda x: (-x["n"], x["port"], x["variable"]))[:10]

    checks = []
    for code, dim, title, desc in CHECKS:
        r = dict(res[code])
        r.pop("fvars", None)
        checks.append(r)
    return {"meta": meta, "checks": checks, "ports": ports, "matrix": matrix, "scores": score, "tagtone": TAGTONE,
            "overall": {"score": overall, "ports": len(ports), "vars": n_vars, "issue_checks": len(issue_checks),
                        "active_checks": len(active), "ports_issue": len(any_issue)},
            "summary": summary, "topvars": tv, "worst": worst}


# =====================================================================
# B1 / B2 template
# =====================================================================
def make_template(data, path):
    b2, b1 = pd.DataFrame(), pd.DataFrame()
    a = data.get("A2")
    if a is not None and a["n_obs"].notna().any():
        sn = a.snapshot.max()
        x = a[a.snapshot == sn].groupby("port")["n_obs"].sum().reset_index()
        b2 = pd.DataFrame({"port": x.port, "snapshot": sn.strftime("%Y-%m-%d"), "dq_count": x.n_obs, "ref_count": None, "ref_source": None})
    c = data.get("C3")
    if c is not None:
        L = c[c.snapshot == c.snapshot.max()].copy()
        L["dq_value"] = L["mean"] * L["n_obs"] * (1 - L["pct_miss"].fillna(0))
        L = L.dropna(subset=["dq_value"])
        b1 = pd.DataFrame({"port": L.port, "variable": L.variable, "snapshot": L.snapshot.dt.strftime("%Y-%m-%d"),
                           "dq_value": L.dq_value.round(2), "ref_value": None, "ref_source": None})
    guide = pd.DataFrame({"วิธีใช้": [
        "1) กรอก ref_count (B2) และ ref_value (B1) จากรายงาน/ระบบต้นทางของ บสส. ณ Snapshot เดียวกัน (กรอกเฉพาะรายการที่ตรวจ ที่เหลือปล่อยว่างได้)",
        "2) ระบุแหล่งอ้างอิงในคอลัมน์ ref_source เช่น ชื่อรายงาน/ไฟล์ EMFI",
        "3) หมายเหตุ B1: dq_value เป็นค่าประมาณ = ค่าเฉลี่ย x จำนวนแถว x (1 - %missing) ควรยืนยันวิธีคำนวณกับ IT หรือขอให้ IT เพิ่มคอลัมน์ผลรวมจริง แล้วแทนที่ค่านี้",
        "4) รัน: python dq_dashboard.py ไฟล์IT.xlsx --control ไฟล์นี้.xlsx"]})
    with pd.ExcelWriter(path, engine="openpyxl") as w:
        guide.to_excel(w, sheet_name="README", index=False)
        b1.to_excel(w, sheet_name="B1", index=False)
        b2.to_excel(w, sheet_name="B2", index=False)
        for ws in w.book.worksheets:
            for col in ws.columns:
                ws.column_dimensions[col[0].column_letter].width = 26 if ws.title != "README" else 120
    print(f"สร้างเทมเพลต B1/B2 แล้ว -> {path}")


# =====================================================================
# ข้อมูลสมมติ (--demo) เขียนเป็นไฟล์ Excel หน้าตาเหมือนไฟล์จาก IT แล้วอ่านผ่านตัวอ่านจริง
# =====================================================================
def make_demo_workbook(path):
    rng = np.random.default_rng(7)
    tables = {"PORT_CLIENT": 80283, "PORT_COST": 35657, "PORT_CTRL": 201, "BUSS_TYPE": 476, "ONL_PERSON": 141046, "PORT_PERSON": 140900,
              "PORT_DEBT": 281395, "RESTRUCT_HD": 920, "RESTRUCT_PV": 5100, "PORT_TR_SCH": 83622, "REC_HD": 1195379, "REC_TR": 1195376,
              "REC_TEMP_PAY_DB": 1195130, "REC_EXPENSE": 7711293, "PORT_EXPENSE": 4601160, "COL_SUM": 92775, "COL_SUM_DAY1": 92775,
              "CASE_ASSET": 9667, "CASE_COURT": 1959941, "NPA_SALE": 36970, "ADJ_DEBT_HD": 732582, "APP_NEW_PRICE_TR": 1233356}
    snaps = [datetime(2025, 10, 8), datetime(2025, 11, 5), datetime(2025, 12, 7)]
    skeys = [s.strftime("%Y-%m-%d") for s in snaps]
    stems = ["ORG", "OS", "COST", "INT", "PRIN", "PAID", "DEBT", "COLL", "FEE", "TAX", "VAT", "TR", "ACC", "LINE", "CUS", "AO", "LAST", "BRANCH"]
    vars_ = {}
    for p in tables:
        k = int(rng.integers(14, 30))
        vs = []
        for i in range(k):
            st = stems[int(rng.integers(0, len(stems)))]
            kind = rng.choice(["AMT", "RATE", "DATE", "STATUS", "TYPE", "NO", "FLAG"], p=[.28, .08, .2, .14, .1, .12, .08])
            vs.append((f"{st}_{kind}" if (st, kind) not in [(x.split('_')[0], x.split('_')[1]) for x, _ in vs] else f"{st}_{kind}_{i}", kind))
        vars_[p] = vs
    A1, A3, C2, C3, C4, E1, A2, D1 = [], [], [], [], [], [], [], []
    for p, n in tables.items():
        for v, kind in vars_[p]:
            m = rng.choice([0, 1, 2, 3], p=[.25, .30, .28, .17])
            base = [0.0, float(rng.uniform(.01, .69)), float(rng.uniform(.7, .999)), 1.0][m]
            mm = [min(1, max(0, base + float(rng.normal(0, .002)))) if 0 < base < 1 else base for _ in snaps]
            A1.append({"PORT": p, "variable": v, **dict(zip(skeys, mm)), "Average of pct_missing": float(np.mean(mm)), "Column7": None})
            ex = [max(0.0, 1 - x) for x in mm]
            r = rng.random()
            if r < .06:
                ex[0] = 0.0
            elif r < .08:
                ex[-1] = 0.0
            A3.append({"PORT": p, "Variable": v, **dict(zip(skeys, ex)), "%Variable Exist": float(np.mean(ex))})
            if kind in ("STATUS", "TYPE", "FLAG"):
                C2.append({"PORT": p, "variable": v, "description": rng.choice(["N", "N/A", "0", "NEW", "NPL", "Y", "-", None]),
                           "n_obs": n, "pct_obs": float(rng.choice([1.0, .99, .97, .85, .6, .4], p=[.3, .1, .1, .2, .2, .1]))})
            if kind in ("AMT", "RATE"):
                for s, x in zip(snaps, mm):
                    mn = float(rng.choice([0, 0, -25000.0, -281005983.0], p=[.5, .2, .2, .1]))
                    C3.append({"PORT": p, "variable": v, "snapshot": s, "n_obs": n, "n_obs_unique": int(n * rng.uniform(.01, .5)),
                               "mean": float(rng.lognormal(9, 2)), "median": 0.0, "min": mn, "max": float(rng.lognormal(16, 2)),
                               "percentile_5": 0.0, "percentile_95": float(rng.lognormal(11, 2)), "pct_missing": x,
                               "Flag_Negative": int(mn < 0), "Flag_Outlier": int(rng.random() < .62), "Flag_Skew": int(rng.random() < .1),
                               "Flag_HighMissing": int(x >= .7)})
            if kind == "DATE":
                for s, x in zip(snaps, mm):
                    be = rng.random() < .05
                    fut = rng.random() < .08
                    const = rng.random() < .06
                    yr = 2568 if be else 2025
                    mx = datetime(yr, 12, 31) if (fut or be) else datetime(2025, 9, int(rng.integers(1, 28)))
                    mn_ = datetime(2002, 3, 31) if not const else mx
                    C4.append({"PORT": p, "variable": v, "snapshot": s, "n_obs": n, "n_obs_unique": 1 if const else int(rng.integers(200, 5000)),
                               "median": mn_, "min": mn_, "max": mx, "percentile_5": mn_, "percentile_95": mx, "pct_missing": x})
            if rng.random() < .55:
                for s in snaps[1:]:
                    nd = int(n * float(rng.choice([0, 0, 0, .0004, .002, .011])))
                    E1.append({"PORT": p, "variable": v, "snapshot": s, "n_obs": n, "n_diff": nd, "pct_diff": nd / n})
    dups = {"ONL_PERSON": [1, 1, 1], "PORT_PERSON": [256] * 3, "RESTRUCT_PV": [56] * 3, "PORT_TR_SCH": [123, 122, 122],
            "REC_TEMP_PAY_DB": [2] * 3, "REC_EXPENSE": [550] * 3, "COL_SUM": [2969] * 3}
    for p, n in tables.items():
        for i, s in enumerate(snaps):
            d = dups.get(p, [0, 0, 0])[i]
            A2.append({"Port": p, "snapshot": s, "n_obs": n + i * 30, "n_obs_unique": n + i * 30 - d, "n_obs_duplicate": d})
    rules = ["ORG_OS < 0", "CLOSE_DATE > AS_OF_DATE", "START_DATE > AS_OF_DATE", "INST_AMT != PRIN_DB + INT_DB",
             "มี SIGN_DATE แต่ไม่มี RESULT_MEETING_DATE", "PAID_AMT ใน PORT_EXPENSE และ REC_EXPENSE ไม่ตรงกัน", "TRAN_OUT_DATE < SALE_DATE"]
    for p, n in tables.items():
        for rl in rng.choice(rules, size=int(rng.integers(2, 5)), replace=False):
            f = float(rng.choice([0, 0, .0144, .08, .15, .9661], p=[.3, .2, .2, .15, .1, .05]))
            for s in snaps:
                D1.append({"PORT": p, "rules": rl, "as_of_date": s, "N_Obs": int(n * f), "Pct_Obs": f})
    C1F = [{"PORT": p, "%discontinue": float(rng.choice([0, 0, 0, .0021]))} for p in tables]
    with pd.ExcelWriter(path) as w:
        for nm, rows in [("A1 Final", A1), ("A2", A2), ("A3 Final", A3), ("C1 Final", C1F), ("C2", C2), ("C3 Final1", C3),
                         ("C4", C4), ("D1", D1), ("E1", E1)]:
            pd.DataFrame(rows).to_excel(w, sheet_name=nm, index=False)


def demo_control(data):
    a = data["A2"]
    sn = a.snapshot.max()
    x = a[a.snapshot == sn].groupby("port")["n_obs"].sum()
    b2 = pd.DataFrame({"port": x.index, "snapshot": sn, "variable": "", "dq": x.values, "ref": x.values.astype(float)})
    b2.loc[[2, 9, 14], "ref"] += [3, -120, 45]
    b1 = pd.DataFrame({"port": ["PORT_CLIENT", "PORT_DEBT", "REC_EXPENSE", "NPA_SALE"], "snapshot": sn,
                       "variable": ["ORG_COST_AMT", "OS_AMT", "PAID_AMT", "SALE_AMT"], "dq": [1.2e9, 8.4e10, 3.3e9, 5.1e8],
                       "ref": [1.2e9, 8.4e10, 3.35e9, 5.1e8]})
    return {"B1": b1, "B2": b2}


# =====================================================================
# main
# =====================================================================
def main():
    ap = argparse.ArgumentParser(description="สร้าง Dashboard คุณภาพข้อมูลจากไฟล์ผลตรวจของ IT")
    ap.add_argument("xlsx", nargs="?", help="ไฟล์ผลตรวจจาก IT (.xlsx)")
    ap.add_argument("--control", help="ไฟล์ผลตรวจ B1/B2 ที่ตรวจเอง")
    ap.add_argument("--template", action="store_true", help="สร้างเทมเพลตสำหรับกรอก B1/B2")
    ap.add_argument("--demo", action="store_true", help="สร้างจากข้อมูลสมมติ")
    ap.add_argument("-o", "--out", help="ชื่อไฟล์ผลลัพธ์ (ค่าเริ่มต้น dashboard.html)")
    a = ap.parse_args()
    if not a.xlsx and not a.demo:
        ap.error("ต้องระบุไฟล์ .xlsx หรือใช้ --demo")
    src = a.xlsx
    if a.demo:
        src = os.path.join(tempfile.gettempdir(), "demo_it_results.xlsx")
        make_demo_workbook(src)
    print(f"อ่านไฟล์: {src}")
    data, used = load_all(src)
    if a.template:
        make_template(data, "b_control_template.xlsx")
        return
    ctrl = demo_control(data) if a.demo else (load_control(a.control) if a.control else None)
    allsn = [s for k in data if data[k] is not None and "snapshot" in data[k] for s in data[k].snapshot.dropna().unique()]
    latest = max(allsn) if allsn else None
    meta = {"title": "รายงานการประเมินและตรวจสอบคุณภาพข้อมูล", "demo": bool(a.demo),
            "source": "ข้อมูลตัวอย่าง" if a.demo else os.path.basename(src),
            "snapshot": th_date(pd.Timestamp(latest)) if latest is not None else "-",
            "snapshots": [th_date(pd.Timestamp(s), True) for s in sorted(set(allsn))],
            "generated": th_date(pd.Timestamp(datetime.now()))}
    payload = san(build(data, ctrl, meta))
    js = json.dumps(payload, ensure_ascii=False, allow_nan=False).replace("</", "<\\/")
    html = HTML.replace("__DATA__", js)
    if a.out:
        out = a.out
    elif a.demo:
        out = "dashboard_demo.html"
    else:
        stamp = latest.strftime("%Y-%m-%d") if latest is not None else datetime.now().strftime("%Y-%m-%d")
        out = f"dashboard_{stamp}.html"
        # เก็บสำเนาชื่อคงที่ไว้เปิดง่าย ๆ เสมอ (เขียนหลังบันทึกไฟล์หลักด้านล่าง)
    with open(out, "w", encoding="utf-8") as f:
        f.write(html)
    if not a.out:
        with open("dashboard_ล่าสุด.html", "w", encoding="utf-8") as f:
            f.write(html)
        print(f"เสร็จแล้ว -> {out}  ({os.path.getsize(out) / 1024:.0f} KB)  (และคัดลอกไว้ที่ dashboard_ล่าสุด.html ให้เปิดง่าย ๆ)")
    else:
        print(f"เสร็จแล้ว -> {out}  ({os.path.getsize(out) / 1024:.0f} KB)")


HTML = r'''<!DOCTYPE html>
<html lang="th">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>รายงานการประเมินและตรวจสอบคุณภาพข้อมูล</title>
<style>
:root{--bg:#eef3f0;--card:#fff;--ink:#12261b;--mut:#61726a;--line:#e1e9e4;--g1:#0b4a2a;--g2:#167a45;--g3:#c9e6d4;--pink:#c2185b;
--bad:#d63a30;--warn:#eb9a12;--good:#1a9a58;--info:#3673c9;--sh:0 8px 28px rgba(11,74,42,.10);--r:18px}
[data-theme=dark]{--bg:#09120e;--card:#101d16;--ink:#e6f2ea;--mut:#8ba297;--line:#1d3226;--g3:#1b3a29;--sh:0 8px 28px rgba(0,0,0,.45)}
*{box-sizing:border-box}
html{scroll-behavior:smooth}
body{margin:0;background:var(--bg);color:var(--ink);font-family:Sarabun,"Noto Sans Thai","Segoe UI",Tahoma,sans-serif;font-size:15px;line-height:1.5}
header{background:linear-gradient(120deg,var(--g1),#0f6a3b 55%,#128a4c);color:#fff;padding:22px 28px 64px;position:relative;overflow:hidden}
header:before{content:"";position:absolute;right:-80px;top:-120px;width:420px;height:420px;border-radius:50%;background:radial-gradient(circle,rgba(255,255,255,.14),transparent 65%)}
header:after{content:"";position:absolute;left:0;bottom:0;height:8px;width:100%;background:linear-gradient(90deg,var(--pink) 0 22%,transparent 22%)}
.top{display:flex;gap:16px;align-items:center;flex-wrap:wrap;position:relative;z-index:1}
.brand{flex:1;min-width:260px}
.brand small{letter-spacing:.14em;text-transform:uppercase;opacity:.75;font-size:12px}
.brand h1{margin:2px 0 0;font-size:26px;font-weight:700}
.chips{display:flex;gap:8px;flex-wrap:wrap;align-items:center}
.chip{background:rgba(255,255,255,.16);border:1px solid rgba(255,255,255,.28);padding:6px 12px;border-radius:999px;font-size:13px;backdrop-filter:blur(6px)}
select,button.btn,input[type=search]{font:inherit;color:inherit}
header select,header button.btn{background:rgba(255,255,255,.16);border:1px solid rgba(255,255,255,.3);color:#fff;padding:6px 12px;border-radius:999px;cursor:pointer}
header select option{color:#111}
.demo{background:var(--pink);color:#fff;text-align:center;padding:6px;font-size:13px;letter-spacing:.02em}
nav#tabs{position:sticky;top:0;z-index:20;margin:-40px 18px 0;display:flex;gap:6px;overflow-x:auto;padding:8px;background:var(--card);border-radius:16px;box-shadow:var(--sh);border:1px solid var(--line)}
nav button{flex:0 0 auto;border:0;background:transparent;color:var(--mut);padding:9px 14px;border-radius:12px;font:inherit;font-weight:600;cursor:pointer;display:flex;align-items:center;gap:7px;transition:.2s}
nav button:hover{background:var(--g3);color:var(--g1)}
nav button.on{background:linear-gradient(120deg,var(--g1),var(--g2));color:#fff}
.dot{width:8px;height:8px;border-radius:50%;background:var(--mut)}
.dot.issue{background:var(--bad)}.dot.ok{background:var(--good)}.dot.pending{background:var(--warn)}.dot.nodata{background:#aab}
main{padding:22px 18px 60px;max-width:1320px;margin:0 auto}
.card{background:var(--card);border:1px solid var(--line);border-radius:var(--r);box-shadow:var(--sh);padding:20px;animation:up .5s both}
@keyframes up{from{opacity:0;transform:translateY(12px)}to{opacity:1;transform:none}}
.grid{display:grid;gap:16px}
.g2{grid-template-columns:repeat(auto-fit,minmax(340px,1fr))}
.g3{grid-template-columns:repeat(auto-fit,minmax(280px,1fr))}
.gk{grid-template-columns:repeat(auto-fit,minmax(190px,1fr))}
h2{margin:0 0 4px;font-size:18px}h3{margin:0 0 12px;font-size:15px}
.sub{color:var(--mut);font-size:13px}
.hero{display:grid;grid-template-columns:auto 1fr;gap:26px;align-items:center;margin-bottom:16px}
.ring{width:170px;height:170px}
.rbg{fill:none;stroke:var(--line);stroke-width:13}
.rfg{fill:none;stroke:url(#rg);stroke-width:13;stroke-linecap:round;transition:stroke-dashoffset 1.4s cubic-bezier(.2,.8,.2,1)}
.rt{font-size:38px;font-weight:800;fill:var(--ink)}.rs{font-size:9.5px;fill:var(--mut)}
.sum{margin:8px 0 0;padding:0;list-style:none}
.sum li{padding:7px 0 7px 26px;position:relative;border-bottom:1px dashed var(--line)}
.sum li:before{content:"";position:absolute;left:4px;top:15px;width:10px;height:10px;border-radius:3px;background:linear-gradient(120deg,var(--pink),var(--g2))}
.kpi{position:relative;overflow:hidden}
.kpi .l{color:var(--mut);font-size:13px}
.kpi .v{font-size:32px;font-weight:800;line-height:1.15;margin-top:4px}
.kpi .s{font-size:12px;color:var(--mut)}
.kpi:after{content:"";position:absolute;right:-18px;top:-18px;width:74px;height:74px;border-radius:50%;background:var(--tc,var(--info));opacity:.13}
.kpi.good{--tc:var(--good)}.kpi.bad{--tc:var(--bad)}.kpi.warn{--tc:var(--warn)}.kpi.info{--tc:var(--info)}.kpi.mut{--tc:var(--mut)}
.kpi.good .v{color:var(--good)}.kpi.bad .v{color:var(--bad)}.kpi.warn .v{color:var(--warn)}
.vc{display:flex;align-items:flex-end;gap:8px;height:190px;padding-top:18px}
.vc>div{flex:1;display:flex;flex-direction:column;align-items:center;justify-content:flex-end;height:100%;cursor:pointer}
.vc .bar{width:100%;max-width:44px;border-radius:9px 9px 3px 3px;background:linear-gradient(180deg,var(--pink),#e45c93);transform-origin:bottom;animation:gh .9s cubic-bezier(.2,.8,.2,1) both;min-height:3px}
.vc .bar.t2{background:linear-gradient(180deg,var(--g2),#4cbf83)}
.vc .bar.na{background:repeating-linear-gradient(45deg,var(--line),var(--line) 4px,transparent 4px,transparent 8px)}
.vc b{font-size:12px;margin-bottom:3px}.vc span{font-size:12px;color:var(--mut);margin-top:5px;font-weight:600}
@keyframes gh{from{transform:scaleY(0)}to{transform:scaleY(1)}}
.bars .brow{display:grid;grid-template-columns:minmax(90px,38%) 1fr auto;gap:10px;align-items:center;padding:4px 0}
.bl{font-size:13px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.bt{height:12px;background:var(--line);border-radius:99px;overflow:hidden}
.bf{height:100%;border-radius:99px;transform-origin:left;animation:gw .9s cubic-bezier(.2,.8,.2,1) both}
@keyframes gw{from{transform:scaleX(0)}to{transform:scaleX(1)}}
.bv{font-size:13px;font-weight:700;min-width:44px;text-align:right}
.donut{display:flex;align-items:center;gap:18px;flex-wrap:wrap}
.donut svg{width:150px;height:150px;flex:0 0 auto}
.dt{font-size:20px;font-weight:800;fill:var(--ink)}.ds{font-size:8px;fill:var(--mut)}
.donut ul{list-style:none;margin:0;padding:0;flex:1;min-width:160px}
.donut li{display:flex;gap:8px;align-items:center;font-size:13px;padding:2px 0}
.donut li i{width:10px;height:10px;border-radius:3px}.donut li b{margin-left:auto}.donut li em{color:var(--mut);font-style:normal;font-size:12px;min-width:46px;text-align:right}
.trend svg{width:100%;height:auto}
.heat{overflow-x:auto}
.hg{display:grid;gap:4px;min-width:760px;align-items:center}
.hh{font-size:12px;font-weight:700;color:var(--mut);text-align:center;padding-bottom:4px}
.hp{font-size:13px;font-weight:600;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;padding-right:6px}
.hp.sel{color:var(--pink)}
.hc{border:0;border-radius:9px;height:30px;font:inherit;font-size:12px;font-weight:700;cursor:pointer;transition:transform .15s}
.hc:hover{transform:scale(1.12);box-shadow:0 4px 12px rgba(0,0,0,.2);position:relative;z-index:2}
.hc.ok{background:color-mix(in srgb,var(--good) 22%,transparent);color:var(--good)}
.hc.na{background:transparent;color:var(--mut);opacity:.5;cursor:default}
.hc.pending{background:color-mix(in srgb,var(--warn) 22%,transparent);color:var(--warn);font-size:11px}
.hc.i1{background:color-mix(in srgb,var(--bad) 30%,transparent);color:var(--bad)}
.hc.i2{background:color-mix(in srgb,var(--bad) 60%,transparent);color:#fff}
.hc.i3{background:var(--bad);color:#fff}
.sc{font-size:12px;font-weight:800;text-align:center;border-radius:9px;padding:6px 0}
.legend{display:flex;gap:14px;flex-wrap:wrap;font-size:12px;color:var(--mut);margin-top:10px}
.legend i{display:inline-block;width:12px;height:12px;border-radius:4px;margin-right:5px;vertical-align:-2px}
.tv{list-style:none;margin:0;padding:0}
.tv li{display:flex;gap:10px;align-items:center;padding:8px 0;border-bottom:1px dashed var(--line)}
.tv .n{width:26px;height:26px;border-radius:8px;background:linear-gradient(120deg,var(--pink),var(--g2));color:#fff;font-weight:800;display:grid;place-items:center;font-size:13px}
.tv small{color:var(--mut)}
.pill{display:inline-block;padding:2px 9px;border-radius:99px;font-size:12px;font-weight:700;background:var(--line)}
.pill.good{background:color-mix(in srgb,var(--good) 20%,transparent);color:var(--good)}
.pill.bad{background:color-mix(in srgb,var(--bad) 20%,transparent);color:var(--bad)}
.pill.warn{background:color-mix(in srgb,var(--warn) 25%,transparent);color:#a26a00}
.pill.info{background:color-mix(in srgb,var(--info) 20%,transparent);color:var(--info)}
.pill.mut{color:var(--mut)}
.ch{display:flex;gap:14px;align-items:center;flex-wrap:wrap;margin-bottom:16px}
.code{font-size:30px;font-weight:800;background:linear-gradient(120deg,var(--pink),var(--g2));-webkit-background-clip:text;background-clip:text;color:transparent}
.dim{background:var(--g3);color:var(--g1);padding:3px 11px;border-radius:99px;font-size:12px;font-weight:700}
[data-theme=dark] .dim{color:#bfe8cf}
.note{background:color-mix(in srgb,var(--warn) 12%,transparent);border-left:4px solid var(--warn);padding:10px 14px;border-radius:10px;font-size:13px;margin:6px 0}
.note.pend{border-color:var(--pink);background:color-mix(in srgb,var(--pink) 8%,transparent)}
.tools{display:flex;gap:10px;flex-wrap:wrap;margin-bottom:12px}
input[type=search]{flex:1;min-width:200px;padding:9px 14px;border-radius:12px;border:1px solid var(--line);background:var(--bg)}
button.btn2{padding:9px 14px;border-radius:12px;border:1px solid var(--line);background:var(--card);cursor:pointer;font:inherit;font-weight:600;color:var(--g1)}
[data-theme=dark] button.btn2{color:#bfe8cf}
button.btn2:hover{background:var(--g3)}
.tw{overflow:auto;max-height:560px;border:1px solid var(--line);border-radius:12px}
table{width:100%;border-collapse:collapse;font-size:13.5px}
th{position:sticky;top:0;background:var(--g1);color:#fff;text-align:left;padding:10px 12px;cursor:pointer;white-space:nowrap;font-weight:600}
th.r,td.r{text-align:right}
td{padding:8px 12px;border-top:1px solid var(--line)}
tr:hover td{background:var(--g3)}
.dlt{font-size:12px;font-weight:700;margin-left:6px}
.foot{color:var(--mut);font-size:12px;text-align:center;margin-top:30px}
@media(max-width:760px){.hero{grid-template-columns:1fr;justify-items:center;text-align:left}header{padding:18px 16px 60px}.brand h1{font-size:21px}}
@media print{header select,header button,nav,.tools,.demo~*button{display:none!important}body{background:#fff}.card{box-shadow:none;break-inside:avoid}}
</style>
</head>
<body>
<div id="demo"></div>
<header>
  <div class="top">
    <div class="brand"><small>Data Quality Cockpit</small><h1 id="ttl"></h1></div>
    <div class="chips" id="chips"></div>
  </div>
</header>
<nav id="tabs"></nav>
<main id="main"></main>
<div class="foot" id="foot"></div>
<svg width="0" height="0" style="position:absolute"><defs><linearGradient id="rg" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="#c2185b"/><stop offset="1" stop-color="#1a9a58"/></linearGradient></defs></svg>
<script>
const D = __DATA__;
const TONE={good:'var(--good)',warn:'var(--warn)',bad:'var(--bad)',info:'var(--info)',mut:'var(--mut)'};
const esc=s=>String(s==null?'':s).replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
const nf=n=>n==null?'–':(Math.abs(n)>=1e9?(n/1e9).toFixed(2)+'B':Math.abs(n)>=1e6?(n/1e6).toFixed(2)+'M':Number(n).toLocaleString('en-US',{maximumFractionDigits:2}));
const pf=(v,d)=>v==null?'–':(v*100).toFixed(d==null?1:d)+'%';
const FM={n:nf,p:pf,t:esc};
const st={tab:'overview',port:'ALL',q:{},sort:{},lim:{},flip:false};
const CH={};D.checks.forEach(c=>CH[c.code]=c);
const $=s=>document.querySelector(s);

/* ---------- ส่วนประกอบพื้นฐาน ---------- */
function kpi(k){
  const cnt=(typeof k.value==='number')?`data-v="${k.value}" data-f="${k.fmt}"`:'';
  return `<div class="card kpi ${k.tone}"><div class="l">${esc(k.label)}</div><div class="v" ${cnt}>${k.value==null?'–':(typeof k.value==='number'?'0':esc(k.value))}</div>${k.sub?`<div class="s">${esc(k.sub)}</div>`:'<div class="s">&nbsp;</div>'}</div>`;
}
function bars(c){
  if(!c.labels.length) return '<div class="sub">ไม่พบรายการ</div>';
  const m=Math.max(...c.values,1e-9),f=FM[c.fmt||'n'];
  return '<div class="bars">'+c.labels.map((l,i)=>`<div class="brow"><div class="bl" title="${esc(l)}">${esc(l)}</div><div class="bt"><div class="bf" style="width:${Math.max(2,c.values[i]/m*100)}%;background:${TONE[c.tone]||TONE.info};animation-delay:${i*.05}s"></div></div><div class="bv">${f(c.values[i])}</div></div>`).join('')+'</div>';
}
function donut(c){
  const R=42,C=2*Math.PI*R,tot=c.parts.reduce((a,p)=>a+p.v,0)||1;let acc=0;
  const segs=c.parts.filter(p=>p.v>0).map(p=>{const len=C*p.v/tot;const s=`<circle cx="60" cy="60" r="${R}" fill="none" stroke="${TONE[p.tone]||p.tone}" stroke-width="16" stroke-dasharray="${len} ${C-len}" stroke-dashoffset="${-acc}" transform="rotate(-90 60 60)"/>`;acc+=len;return s}).join('');
  const lg=c.parts.map(p=>`<li><i style="background:${TONE[p.tone]||p.tone}"></i>${esc(p.l)}<b>${nf(p.v)}</b><em>${(p.v/tot*100).toFixed(1)}%</em></li>`).join('');
  return `<div class="donut"><svg viewBox="0 0 120 120"><circle cx="60" cy="60" r="${R}" fill="none" stroke="var(--line)" stroke-width="16"/>${segs}<text x="60" y="60" text-anchor="middle" class="dt">${nf(tot)}</text><text x="60" y="72" text-anchor="middle" class="ds">รวม</text></svg><ul>${lg}</ul></div>`;
}
function trend(c){
  const W=320,H=130,px=34,py=22,f=FM[c.fmt||'n'];
  const vs=c.values.map(v=>v==null?0:v),mx=Math.max(...vs,1e-9),mn=Math.min(...vs,0);
  const x=i=>px+(vs.length<2?(W-2*px)/2:i*(W-2*px)/(vs.length-1));
  const y=v=>H-py-(v-mn)/((mx-mn)||1)*(H-2*py);
  const pts=vs.map((v,i)=>x(i)+','+y(v)).join(' ');
  const area=`${x(0)},${H-py} ${pts} ${x(vs.length-1)},${H-py}`;
  const dots=vs.map((v,i)=>`<circle cx="${x(i)}" cy="${y(v)}" r="4.5" fill="var(--card)" stroke="var(--pink)" stroke-width="2.5"/><text x="${x(i)}" y="${y(v)-10}" text-anchor="middle" font-size="10" font-weight="700" fill="var(--ink)">${f(v)}</text><text x="${x(i)}" y="${H-5}" text-anchor="middle" font-size="9.5" fill="var(--mut)">${esc(c.labels[i])}</text>`).join('');
  return `<div class="trend"><svg viewBox="0 0 ${W} ${H}"><polygon points="${area}" fill="var(--pink)" opacity=".10"/><polyline points="${pts}" fill="none" stroke="var(--pink)" stroke-width="3" stroke-linejoin="round" stroke-linecap="round"/>${dots}</svg></div>`;
}
function chartCard(c){
  const body=c.t==='bars'?bars(c):c.t==='donut'?donut(c):trend(c);
  return `<div class="card"><h3>${esc(c.title)}</h3>${body}</div>`;
}
function pill(v){const t=D.tagtone[v]||'info';return `<span class="pill ${t}">${esc(v)}</span>`}
function delta(c){if(!c.delta)return'';const f=(a)=>(a>0?'▲ ':'▼ ')+Math.abs(a);return `<span class="dlt" style="color:${c.delta.tables>0?'var(--bad)':'var(--good)'}" title="เทียบกับ baseline">${f(c.delta.tables)} ตาราง</span>`}

/* ---------- หน้าภาพรวม ---------- */
function vcols(vals,cls,sq){
  const arr=vals.map(v=>v==null?null:v),mx=Math.max(...arr.map(v=>v||0),1);
  return '<div class="vc">'+D.checks.map((c,i)=>{const v=arr[i];
    if(v==null) return `<div><b>รอ</b><div class="bar na" style="height:26px"></div><span>${c.code}</span></div>`;
    const h=Math.max(2,(sq?Math.sqrt(v/mx):v/mx)*100);
    return `<div data-go="${c.code}"><b>${nf(v)}</b><div class="bar ${cls}" style="height:${h}%;animation-delay:${i*.05}s"></div><span>${c.code}</span></div>`}).join('')+'</div>';
}
function heat(){
  const codes=D.checks.map(c=>c.code);
  const ports=[...D.ports].sort((a,b)=>st.flip?a.localeCompare(b):(D.scores[a]-D.scores[b])||a.localeCompare(b));
  const cols=`minmax(130px,1.5fr) repeat(${codes.length},minmax(38px,1fr)) 60px`;
  let h=`<div class="hg" style="grid-template-columns:${cols}"><div class="hh" style="text-align:left">ตาราง</div>`+codes.map(c=>`<div class="hh">${c}</div>`).join('')+'<div class="hh">คะแนน</div>';
  ports.forEach(p=>{
    h+=`<div class="hp ${st.port===p?'sel':''}" title="${esc(p)}">${esc(p)}</div>`;
    codes.forEach(c=>{const m=D.matrix[p][c];let cls,txt;
      if(m.s==='issue'){cls=m.n>50?'i3':m.n>5?'i2':'i1';txt=nf(m.n)}else if(m.s==='ok'){cls='ok';txt='✓'}else if(m.s==='pending'){cls='pending';txt='รอ'}else{cls='na';txt='–'}
      h+=`<button class="hc ${cls}" ${m.s==='issue'||m.s==='ok'?`data-cell="${c}|${esc(p)}"`:''} title="${esc(p)} · ${c}${m.s==='issue'?' · พบ '+m.n:''}">${txt}</button>`});
    const s=D.scores[p],col=s>=80?'good':s>=50?'warn':'bad';
    h+=`<div class="sc pill ${col}">${s.toFixed(0)}</div>`;
  });
  return h+'</div><div class="legend"><span><i style="background:var(--good);opacity:.4"></i>ผ่าน</span><span><i style="background:var(--bad);opacity:.35"></i>พบ 1–5</span><span><i style="background:var(--bad);opacity:.65"></i>6–50</span><span><i style="background:var(--bad)"></i>&gt; 50</span><span><i style="background:var(--warn);opacity:.5"></i>รอตรวจ</span><span>– ไม่มีข้อมูลในเกณฑ์นั้น</span><span>ตัวเลขในช่อง = จำนวนรายการที่พบ (คลิกเพื่อดูรายละเอียด)</span></div>';
}
function overview(){
  const o=D.overall,R=54,C=2*Math.PI*R;
  const ring=`<svg viewBox="0 0 140 140" class="ring"><circle cx="70" cy="70" r="${R}" class="rbg"/><circle cx="70" cy="70" r="${R}" class="rfg" id="rfg" style="stroke-dasharray:${C};stroke-dashoffset:${C}" data-off="${C*(1-o.score/100)}" transform="rotate(-90 70 70)"/><text x="70" y="72" class="rt" text-anchor="middle">${o.score.toFixed(0)}</text><text x="70" y="92" class="rs" text-anchor="middle">คะแนนรวม</text></svg>`;
  const tb=D.checks.map(c=>c.n_tables),it=D.checks.map(c=>c.n_items);
  const tv=D.topvars.length?'<ul class="tv">'+D.topvars.map(t=>`<li><div class="n">${t.n}</div><div style="flex:1"><b>${esc(t.variable)}</b><br><small>${esc(t.port)}</small></div><div>${t.codes.map(c=>`<span class="pill warn" style="margin-left:3px">${c}</span>`).join('')}</div></li>`).join('')+'</ul>':'<div class="sub">ไม่พบตัวแปรที่ติดหลายเกณฑ์พร้อมกัน</div>';
  const wr='<div class="bars">'+D.worst.map(p=>`<div class="brow"><div class="bl">${esc(p)}</div><div class="bt"><div class="bf" style="width:${100-D.scores[p]}%;background:var(--bad)"></div></div><div class="bv">${D.scores[p].toFixed(0)}</div></div>`).join('')+'</div>';
  return `<div class="card hero">${ring}<div><h2>สรุปผลการตรวจสอบคุณภาพข้อมูล</h2><div class="sub">Snapshot ล่าสุด ${esc(D.meta.snapshot)}</div><ul class="sum">${D.summary.map(s=>`<li>${esc(s)}</li>`).join('')}</ul></div></div>
  <div class="grid gk" style="margin-bottom:16px">${kpi({label:'ตารางที่ตรวจ',value:o.ports,tone:'info',fmt:'n'})}${kpi({label:'ตัวแปรที่ตรวจ (A1)',value:o.vars,tone:'info',fmt:'n'})}${kpi({label:'เกณฑ์ที่พบประเด็น',value:o.issue_checks,tone:o.issue_checks?'bad':'good',fmt:'n',sub:'จาก '+o.active_checks+' เกณฑ์ที่ตรวจแล้ว'})}${kpi({label:'ตารางที่พบอย่างน้อย 1 ประเด็น',value:o.ports_issue,tone:'warn',fmt:'n',sub:'จาก '+o.ports+' ตาราง'})}</div>
  <div class="grid g2" style="margin-bottom:16px"><div class="card"><h3>จำนวนตารางที่พบประเด็น แยกตามเกณฑ์</h3>${vcols(tb,'')}<div class="sub">คลิกแท่งเพื่อดูรายละเอียด</div></div>
  <div class="card"><h3>จำนวนรายการที่พบ แยกตามเกณฑ์</h3>${vcols(it,'t2',true)}<div class="sub">ความสูงแท่งใช้สเกลรากที่สอง เพื่อให้เห็นเกณฑ์ที่มีจำนวนน้อย</div></div></div>
  <div class="card heat" style="margin-bottom:16px"><div style="display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:8px"><div><h2>แผนที่ความเสี่ยง: ตาราง × เกณฑ์</h2><div class="sub">คะแนนของแต่ละตาราง = สัดส่วนเกณฑ์ที่ผ่าน จากเกณฑ์ที่ตรวจแล้ว (ไม่รวม B1/B2 ที่ยังรอ)</div></div><button class="btn2" id="flip">${st.flip?'เรียงตามคะแนน':'เรียง A–Z'}</button></div><div style="margin-top:12px">${heat()}</div></div>
  <div class="grid g2"><div class="card"><h3>ตารางที่ควรดูก่อน (คะแนนต่ำสุด)</h3>${wr}</div><div class="card"><h3>ตัวแปรที่ติดหลายเกณฑ์พร้อมกัน</h3>${tv}</div></div>`;
}

/* ---------- หน้ารายละเอียดแต่ละเกณฑ์ ---------- */
function rowsOf(c){
  let r=c.rows;
  if(st.port!=='ALL') r=r.filter(x=>x[0]===st.port);
  const q=(st.q[c.code]||'').toLowerCase();
  if(q) r=r.filter(x=>x.some(v=>String(v==null?'':v).toLowerCase().includes(q)));
  const s=st.sort[c.code];
  if(s){const t=c.cols[s.i].type,k=(t==='num'||t==='pct')?1:0;
    r=[...r].sort((a,b)=>{let x=a[s.i],y=b[s.i];if(x==null)return 1;if(y==null)return -1;return (k?x-y:String(x).localeCompare(String(y),'th'))*s.d});}
  return r;
}
function cell(v,t){
  if(v==null||v==='') return '<td>–</td>';
  if(t==='num') return `<td class="r">${nf(v)}</td>`;
  if(t==='pct') return `<td class="r">${pf(v,2)}</td>`;
  if(t==='tag') return `<td>${pill(v)}</td>`;
  return `<td>${esc(v)}</td>`;
}
function tbody(c){
  const rs=rowsOf(c),lim=st.lim[c.code]||150;
  const head='<tr>'+c.cols.map((k,i)=>{const s=st.sort[c.code];const a=s&&s.i===i?(s.d>0?' ▲':' ▼'):'';return `<th class="${k.type==='num'||k.type==='pct'?'r':''}" data-sort="${i}">${esc(k.label)}${a}</th>`}).join('')+'</tr>';
  const body=rs.slice(0,lim).map(r=>'<tr>'+r.map((v,i)=>cell(v,c.cols[i].type)).join('')+'</tr>').join('');
  return `<div class="tw"><table><thead>${head}</thead><tbody>${body||`<tr><td colspan="${c.cols.length}" class="sub">ไม่พบรายการ</td></tr>`}</tbody></table></div><div class="sub" style="margin-top:8px">แสดง ${Math.min(lim,rs.length).toLocaleString()} จาก ${rs.length.toLocaleString()} รายการ ${rs.length>lim?'<button class="btn2" id="more" style="margin-left:8px">แสดงเพิ่ม</button>':''}</div>`;
}
function checkPage(c){
  let h=`<div class="ch"><div class="code">${c.code}</div><div><h2>${esc(c.title)}</h2><div class="sub">${esc(c.desc)}</div></div><span class="dim">${esc(c.dim)}</span>${c.status==='ok'?`<span class="pill ${c.n_tables?'bad':'good'}">${c.n_tables?'พบ '+c.n_tables+' ตาราง · '+nf(c.n_items)+' รายการ':'ผ่านทุกตาราง'}</span>${delta(c)}`:''}</div>`;
  if(c.status!=='ok'){
    return h+`<div class="card">${c.notes.map(n=>`<div class="note ${c.status==='pending'?'pend':''}">${esc(n)}</div>`).join('')}${c.status==='pending'?'<div class="sub" style="margin-top:10px">ระหว่างรอ เกณฑ์นี้จะไม่ถูกนับรวมในคะแนนคุณภาพ</div>':''}</div>`;
  }
  h+=`<div class="grid gk" style="margin-bottom:16px">${c.kpis.map(kpi).join('')}</div>`;
  h+=`<div class="grid g2" style="margin-bottom:16px">${c.charts.map(chartCard).join('')}</div>`;
  if(c.notes.length) h+=c.notes.map(n=>`<div class="note">${esc(n)}</div>`).join('');
  h+=`<div class="card" style="margin-top:16px"><h3>รายละเอียด${st.port!=='ALL'?' · '+esc(st.port):''}</h3><div class="tools"><input type="search" id="q" placeholder="ค้นหาตาราง / ตัวแปร / ค่า..." value="${esc(st.q[c.code]||'')}"><button class="btn2" id="csv">ดาวน์โหลด CSV</button></div><div id="tb">${tbody(c)}</div></div>`;
  return h;
}

/* ---------- ควบคุมหน้า ---------- */
function countUp(){
  document.querySelectorAll('[data-v]').forEach(el=>{
    const v=+el.dataset.v,f=FM[el.dataset.f]||nf,t0=performance.now(),d=900;
    (function step(t){const p=Math.min(1,(t-t0)/d),e=1-Math.pow(1-p,3);el.textContent=f(v*e);if(p<1)requestAnimationFrame(step);else el.textContent=f(v)})(t0);
  });
  const r=document.getElementById('rfg');if(r)requestAnimationFrame(()=>requestAnimationFrame(()=>{r.style.strokeDashoffset=r.dataset.off}));
}
function go(tab,port){if(port)st.port=port;st.tab=tab;$('#psel').value=st.port;render();window.scrollTo({top:0,behavior:'smooth'})}
function tabs(){
  $('#tabs').innerHTML=[['overview','ภาพรวม','']].concat(D.checks.map(c=>[c.code,c.code,c.status==='ok'?(c.n_tables?'issue':'ok'):c.status])).map(([k,l,s])=>`<button data-tab="${k}" class="${st.tab===k?'on':''}">${s?`<span class="dot ${s}"></span>`:''}${l}</button>`).join('');
  document.querySelectorAll('[data-tab]').forEach(b=>b.onclick=()=>go(b.dataset.tab));
}
function render(){
  tabs();
  const m=$('#main');
  m.innerHTML=st.tab==='overview'?overview():checkPage(CH[st.tab]);
  countUp();
  const fl=$('#flip');if(fl)fl.onclick=()=>{st.flip=!st.flip;render()};
  document.querySelectorAll('[data-go]').forEach(e=>e.onclick=()=>go(e.dataset.go));
  document.querySelectorAll('[data-cell]').forEach(e=>e.onclick=()=>{const [c,p]=e.dataset.cell.split('|');go(c,p)});
  if(st.tab!=='overview'&&CH[st.tab].status==='ok') bindTable(CH[st.tab]);
}
function bindTable(c){
  const redo=()=>{$('#tb').innerHTML=tbody(c);wire()};
  const wire=()=>{
    document.querySelectorAll('[data-sort]').forEach(th=>th.onclick=()=>{const i=+th.dataset.sort,s=st.sort[c.code];st.sort[c.code]={i,d:s&&s.i===i?-s.d:1};redo()});
    const mo=$('#more');if(mo)mo.onclick=()=>{st.lim[c.code]=(st.lim[c.code]||150)+300;redo()};
  };
  wire();
  $('#q').oninput=e=>{st.q[c.code]=e.target.value;st.lim[c.code]=150;redo()};
  $('#csv').onclick=()=>{
    const rs=rowsOf(c),q=v=>'"'+String(v==null?'':v).replace(/"/g,'""')+'"';
    const csv='\ufeff'+[c.cols.map(k=>q(k.label)).join(',')].concat(rs.map(r=>r.map(q).join(','))).join('\r\n');
    const a=document.createElement('a');a.href=URL.createObjectURL(new Blob([csv],{type:'text/csv;charset=utf-8'}));a.download='DQ_'+c.code+'.csv';a.click();
  };
}
/* ---------- เริ่มต้น ---------- */
$('#ttl').textContent=D.meta.title;
$('#chips').innerHTML=`<span class="chip">Snapshot ล่าสุด ${esc(D.meta.snapshot)}</span><span class="chip">${esc(D.meta.snapshots.join(' · '))}</span><select id="psel"><option value="ALL">ทุกตาราง</option>${D.ports.map(p=>`<option>${esc(p)}</option>`).join('')}</select><button class="btn" id="th">โหมดมืด</button><button class="btn" id="pr">พิมพ์ / PDF</button>`;
$('#psel').onchange=e=>{st.port=e.target.value;render()};
$('#th').onclick=()=>{const d=document.documentElement;d.dataset.theme=d.dataset.theme==='dark'?'':'dark'};
$('#pr').onclick=()=>window.print();
if(D.meta.demo) $('#demo').innerHTML='<div class="demo">ข้อมูลตัวอย่าง (DEMO) — ตัวเลขทั้งหมดเป็นข้อมูลสมมติเพื่อดูหน้าตา ไม่ใช่ข้อมูลจริงของ บสส.</div>';
$('#foot').textContent='สร้างเมื่อ '+D.meta.generated+' · แหล่งข้อมูล: '+D.meta.source+' · เกณฑ์ตัดสินปรับได้ใน CONFIG ของ dq_dashboard.py';
render();
</script>
</body>
</html>
'''


if __name__ == "__main__":
    main()