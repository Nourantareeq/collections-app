# -*- coding: utf-8 -*-
"""
ربط النشاط بالمكالمات - نسخة المطابقة اليومية (v2 rules)

قواعد المطابقة:
- المطابقة برقم الهاتف المُطبّع (9 أرقام) -> العميل، من كلتا ورقتي ملف الأرقام
  (بما فيها الأرقام المسندة لأكثر من عميل)، وبنفس اليوم فقط (تاريخ Created on
  = تاريخ Call Time). لا يوجد أي تراكم بين الأيام.
- مكالمات Call Script (711) الصادرة تُحسب محاولة فقط، ولا تُحسب أبداً
  مكالمة مردوداً عليها ولا مكالمة فعلية. وفي أداء المحصّلين تُنسب للمحصّل
  صاحب العميل (نفس اليوم) وتُحسب محاولة.
- المكالمة الفعلية = صادر تم الرد عليه، أو وارد تم الرد عليه بمدة تحدث > 0.
- عدد المكالمات الواردة يُحسب بالأرجل (legs) شاملاً أرجل الطابور
  (Inbound Queue): المكالمة الواحدة قد تظهر كأكثر من رجل.
- يقبل ملف المكالمات CSV أو XLSX.
- كل المعالجة تتم محلياً على الجهاز، لا يُرفع أي شيء لأي خادم.
"""
import io
import re
import pandas as pd

# ---------- أسماء الأعمدة المضافة (مطابقة لملف activity_23days_v2) ----------
COL_STATUS = "حالة الاتصال"
COL_ATT = "عدد محاولات الاتصال (صادر)"
COL_ANS = "عدد المكالمات المردود عليها"
COL_UNANS = "عدد غير المردود عليها"
COL_TOT = "إجمالي مدة المكالمات"
COL_MAX = "أطول مكالمة"
COL_INB = "مكالمات واردة من العميل"
COL_ACTUAL = "مكالمة فعلية؟"
COL_ATTEMPT = "محاولة اتصال؟"
ADDED_COLS = [COL_STATUS, COL_ATT, COL_ANS, COL_UNANS, COL_TOT, COL_MAX,
              COL_INB, COL_ACTUAL, COL_ATTEMPT]

SCRIPT_NOTE = "Call Script auto-call"
NOROUTE_NOTE = "No route to destination - no Dialed pattern, clean number used"


# ---------- التطبيع والمساعدات ----------
def norm_phone(x):
    """تطبيع الرقم إلى 9 أرقام سعودية (5xxxxxxxx) أو None."""
    d = re.sub(r"\D", "", str(x))
    if d.startswith("0"):
        d = d.lstrip("0")
    if len(d) > 9:
        d = d[-9:]
    return d if len(d) == 9 else None


def talk_secs(s):
    try:
        h, m, sec = str(s).split(":")
        return int(h) * 3600 + int(m) * 60 + int(float(sec))
    except Exception:
        return 0


def fmt_dur(s):
    return "%d:%02d:%02d" % (s // 3600, s % 3600 // 60, s % 60)


def client_id(x):
    return str(x).split(".")[0].strip()


# ---------- ملف الأرقام: رقم -> مجموعة عملاء (كل الأوراق) ----------
def build_phone_map(map_file):
    sheets = pd.read_excel(map_file, sheet_name=None, dtype=str)
    num2clients = {}
    for name, df in sheets.items():
        df = df.rename(columns=str.strip)
        if "Phone" not in df.columns or "Debitor" not in df.columns:
            continue
        for ph, deb in zip(df["Phone"], df["Debitor"]):
            n = norm_phone(ph)
            if n and pd.notna(deb) and str(deb).strip():
                num2clients.setdefault(n, set()).add(client_id(deb))
    return num2clients


# ---------- قراءة ملف المكالمات (CSV أو XLSX) ----------
def read_calls(calls_file, filename=""):
    name = (filename or getattr(calls_file, "name", "") or "").lower()
    if name.endswith(".csv"):
        return pd.read_csv(calls_file, dtype=str, encoding="utf-8-sig")
    xl = pd.ExcelFile(calls_file)
    sheet = xl.sheet_names[0]
    for s in xl.sheet_names:
        if "call_reports" in s:
            sheet = s
            break
    return xl.parse(sheet)


# ---------- تنظيف المكالمات وتصنيفها ----------
def clean_calls(calls, num2clients):
    """يرجع DataFrame منظم: صف لكل مكالمة (غير Internal) مع الرقم المطبّع
    والعملاء المطابقون والنوع (out/script/in/inq) والتاريخ."""
    rows = []
    for _, r in calls.iterrows():
        d = str(r["Direction"])
        if d not in ("Outbound", "Inbound", "Inbound Queue"):
            continue
        det = str(r["Call Activity Details"])
        to = str(r["To"])
        note = ""
        if "Call Script" in to and d == "Outbound":
            # مكالمة سكربت آلية: الرقم في From، تُحسب محاولة فقط
            num, kind = norm_phone(r["From"]), "script"
            note = SCRIPT_NOTE
        elif d == "Outbound":
            m2 = re.search(r"Dialed:\s*\((\d+)\)", det)
            if m2:
                num = norm_phone(m2.group(1))
                if num is None:
                    note = "رقم غير قياسي: " + m2.group(1)
            else:
                num = norm_phone(r["To"])
                if num and "No route" in det:
                    note = NOROUTE_NOTE
            kind = "out"
        else:
            num = norm_phone(r["From"])
            kind = "in" if d == "Inbound" else "inq"
            if num is None:
                note = "رقم غير قياسي: " + str(r["From"])
        ext = r["To"] if (d != "Outbound" or kind == "script") else r["From"]
        debs = sorted(num2clients.get(num, ())) if num else []
        rows.append({
            "Call Time": r["Call Time"],
            "Direction": d,
            "Status": str(r["Status"]),
            "Talking": r["Talking"],
            "Ringing": r["Ringing"],
            "Extension": ext,
            "Phone (9 digits)": num or "",
            "Debtor(s)": ",".join(debs),
            "Note": note,
            "kind": kind,
            "talk": talk_secs(r["Talking"]),
            "date": pd.to_datetime(r["Call Time"]).date(),
        })
    return pd.DataFrame(rows)


# ---------- التجميع لكل (عميل, يوم) ----------
def aggregate(cleaned):
    cols = ["client", "date", "att", "ans", "unans", "tot", "mx", "inb", "inans"]
    if len(cleaned) == 0:
        return pd.DataFrame(columns=cols)
    cl = cleaned.rename(columns={"Debtor(s)": "debtors"})
    recs = []
    for r in cl.itertuples(index=False):
        for c in r.debtors.split(","):
            if c:
                recs.append((c, r.date, r.kind, r.Status, r.talk))
    df = pd.DataFrame(recs, columns=["client", "date", "kind", "status", "talk"])
    if len(df) == 0:
        return pd.DataFrame(columns=cols)

    def agg(x):
        out_mask = x["kind"] == "out"
        ans_mask = out_mask & (x["status"] == "Answered")
        in_ans = x["kind"].isin(["in"]) & (x["status"] == "Answered") & (x["talk"] > 0)
        return pd.Series({
            # المحاولات = الصادر + السكربت (السكربت محاولة فقط)
            "att": x["kind"].isin(["out", "script"]).sum(),
            # المردود/غير المردود: الصادر العادي فقط (بدون السكربت)
            "ans": ans_mask.sum(),
            "unans": (out_mask & (x["status"] != "Answered")).sum(),
            "tot": x.loc[ans_mask, "talk"].sum(),
            "mx": x.loc[ans_mask, "talk"].max() if ans_mask.any() else 0,
            # الوارد بالأرجل: Inbound + Inbound Queue
            "inb": x["kind"].isin(["in", "inq"]).sum(),
            # وارد فعلي: تم الرد عليه بمدة تحدث > 0
            "inans": in_ans.sum(),
        })

    return df.groupby(["client", "date"]).apply(agg, include_groups=False).reset_index()


def enrich_activity(act, agg):
    """يرجع نسخة من النشاط مع الأعمدة التسعة المضافة، بمطابقة (عميل, يوم)."""
    act = act.copy()
    act["_date"] = pd.to_datetime(act["Created on"]).dt.date
    act["_client"] = act["Debtor"].map(client_id)
    m = act.merge(agg, left_on=["_client", "_date"], right_on=["client", "date"],
                  how="left")
    for c in ["att", "ans", "unans", "tot", "mx", "inb", "inans"]:
        m[c] = m[c].fillna(0).astype(int)
    m[COL_STATUS] = ["تم الرد" if a > 0 else ("محاولة اتصال - لم يرد" if t > 0
                     else "لا يوجد اتصال") for a, t in zip(m["ans"], m["att"])]
    m[COL_ATT] = m["att"]
    m[COL_ANS] = m["ans"]
    m[COL_UNANS] = m["unans"]
    m[COL_TOT] = m["tot"].map(fmt_dur)
    m[COL_MAX] = m["mx"].map(fmt_dur)
    m[COL_INB] = m["inb"]
    m[COL_ACTUAL] = ["مكالمة فعلية" if (a > 0 or ia > 0) else "لا يوجد مكالمة فعلية"
                     for a, ia in zip(m["ans"], m["inans"])]
    m[COL_ATTEMPT] = ["فيه محاولة اتصال" if (t > 0 or i > 0) else "لا يوجد محاولة"
                      for t, i in zip(m["att"], m["inb"])]
    m = m.drop(columns=["_date", "_client", "client", "date",
                        "att", "ans", "unans", "tot", "mx", "inb", "inans"])
    return m


def missing_map_view(act, num2clients):
    """عملاء موجودون في النشاط وليس لهم أي رقم في ملف الأرقام."""
    known = set()
    for s in num2clients.values():
        known |= s
    act = act.copy()
    act["_client"] = act["Debtor"].map(client_id)
    miss = act[~act["_client"].isin(known)]
    if len(miss) == 0:
        return pd.DataFrame(columns=["Debtor", "Debitor Name",
                                     "عدد صفوف النشاط", "أيام النشاط"])
    g = miss.groupby("_client").agg(
        **{"Debitor Name": ("Debitor Name", "first"),
           "عدد صفوف النشاط": ("_client", "size"),
           "أيام النشاط": ("Created on", lambda s: s.astype(str).str[:10].nunique())}
    ).reset_index().rename(columns={"_client": "Debtor"})
    return g[["Debtor", "Debitor Name", "عدد صفوف النشاط", "أيام النشاط"]]


# ---------- المعالجة الكاملة (تُستخدم من الواجهة ومن الاختبارات) ----------
def process(activity_file, map_file, calls_file, calls_filename=""):
    """يرجع (enriched_df, cleaned_df, missing_df, summary_dict)."""
    num2clients = build_phone_map(map_file)
    calls = read_calls(calls_file, calls_filename)
    cleaned = clean_calls(calls, num2clients)
    agg = aggregate(cleaned)
    act = pd.read_excel(activity_file, dtype=str)
    enriched = enrich_activity(act, agg)
    missing = missing_map_view(act, num2clients)
    summary = {
        "rows": len(enriched),
        "calls_clean": len(cleaned),
        "script_calls": int((cleaned["kind"] == "script").sum()) if len(cleaned) else 0,
        "unmatched_calls": int((cleaned["Debtor(s)"] == "").sum()) if len(cleaned) else 0,
        "missing_clients": len(missing),
        "actual": int((enriched[COL_ACTUAL] == "مكالمة فعلية").sum()),
        "attempts": int((enriched[COL_ATTEMPT] == "فيه محاولة اتصال").sum()),
    }
    return enriched, cleaned, missing, summary


# ---------- بناء ملف الإكسل الناتج ----------
def build_output_workbook(enriched, cleaned, missing):
    """يبني الملف من الصفر بسرعة (مثل ملف v2 نفسه): Sheet1 = النشاط +
    الأعمدة التسعة، ورقة المكالمات بعد التنظيف (غير المطابق أصفر)، وورقة
    العملاء بدون أرقام."""
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="xlsxwriter") as writer:
        enriched.to_excel(writer, sheet_name="Sheet1", index=False)
        book = writer.book
        bold = book.add_format({"bold": True})
        yl = book.add_format({"bg_color": "#FFFF00"})
        wsen = writer.sheets["Sheet1"]
        for i, h in enumerate(enriched.columns):
            wsen.write(0, i, h, bold)
            wsen.set_column(i, i, 22)

        ws = book.add_worksheet("المكالمات بعد التنظيف")
        heads = ["Call Time", "Direction", "Status", "Talking", "Ringing",
                 "Extension", "Phone (9 digits)", "Debtor(s)", "Note"]
        for i, h in enumerate(heads):
            ws.write(0, i, h, bold)
        ws.set_column(0, 8, 20)
        r = 1
        for row in cleaned.itertuples(index=False):
            vals = ["" if pd.isna(v) else str(v) for v in (
                row[0], row[1], row[2], row[3], row[4], row[5], row[6],
                row[7], row[8])]
            fill = yl if not row[7] else None
            for i, v in enumerate(vals):
                ws.write(r, i, v, fill)
            r += 1

        if len(missing):
            missing.to_excel(writer, sheet_name="عملاء بدون أرقام في الشيت",
                             index=False)
            wsm = writer.sheets["عملاء بدون أرقام في الشيت"]
            for i, h in enumerate(missing.columns):
                wsm.write(0, i, h, bold)
        else:
            wsm = book.add_worksheet("عملاء بدون أرقام في الشيت")
            wsm.write(0, 0, "كل عملاء النشاط لهم أرقام في ملف الأرقام", bold)
    return out.getvalue()


# ======================================================================
# أداء المحصّلين (Collectors performance)
# ======================================================================
COLLECTOR_COL = "Created by"


def ring_secs(s):
    return talk_secs(s)


def attribute_calls(cleaned, act):
    """يربط كل مكالمة صادرة (غير السكربت) بالمحصّل عبر: المكالمة -> العميل
    (من ملف الأرقام) -> صفوف نشاط نفس العميل في نفس اليوم -> Created by.
    - المكالمة تُحسب مرة واحدة لكل محصّل (حتى لو الرقم يخص أكثر من عميل له).
    - لو عميل له صفوف عند أكثر من محصّل في نفس اليوم: تُنسب المكالمة للمحصّل
      صاحب أول صف نشاط في ذلك اليوم (الأقدم وقتاً)، وتُسجَّل الحالة.
    يرجع (calls_attr, info). calls_attr: صف لكل (مكالمة، عميل، محصّل)."""
    a = act.copy()
    a["_date"] = pd.to_datetime(a["Created on"]).dt.date
    a["_client"] = a["Debtor"].map(client_id)
    a["_ts"] = pd.to_datetime(a["Created on"])
    a = a.sort_values("_ts")
    # لكل (عميل، يوم): المحصّل الأول + هل فيه أكثر من محصّل
    grp = a.groupby(["_client", "_date"])[COLLECTOR_COL]
    first = grp.first()
    nuniq = grp.nunique()
    owner = pd.DataFrame({"collector": first, "n_collectors": nuniq}).reset_index()

    out = cleaned[cleaned["kind"].isin(["out", "script"])].copy()
    out["call_id"] = out.index
    out["ring"] = out["Ringing"].map(ring_secs)
    out["answered"] = (out["Status"] == "Answered") & (out["kind"] != "script")
    out["debtors_l"] = out["Debtor(s)"].str.split(",")
    ex = out.explode("debtors_l").rename(columns={"debtors_l": "client"})
    ex = ex[ex["client"].notna() & (ex["client"] != "")]
    m = ex.merge(owner, left_on=["client", "date"],
                 right_on=["_client", "_date"] if False else ["_client", "_date"],
                 how="left") if False else ex.merge(
        owner.rename(columns={"_client": "client", "_date": "date"}),
        on=["client", "date"], how="left")
    matched = m[m["collector"].notna()]
    # مكالمة واحدة لكل محصّل
    dedup = matched.sort_values("client").drop_duplicates(["call_id", "collector"])
    info = {
        "out_calls": int((out["kind"] == "out").sum()),
        "out_no_client": int(((out["Debtor(s)"] == "") & (out["kind"] == "out")).sum()),
        "out_client_no_activity_that_day": int(
            out.loc[out["kind"] == "out", "call_id"].isin(ex["call_id"]).sum()
            - matched.loc[matched["kind"] == "out", "call_id"].nunique()),
        "out_attributed": int(matched.loc[matched["kind"] == "out", "call_id"].nunique()),
        "script_attributed": int(matched.loc[matched["kind"] == "script", "call_id"].nunique()),
        "calls_to_multi_collector_client_day": int(
            matched.loc[matched["n_collectors"] > 1, "call_id"].nunique()),
        "client_days_multi_collector": int((owner["n_collectors"] > 1).sum()),
        "script_calls": int((cleaned["kind"] == "script").sum()),
        "script_ring_secs": int(cleaned.loc[cleaned["kind"] == "script",
                                            "Ringing"].map(ring_secs).sum()),
    }
    return matched, owner, info


def _fmt_hms(s):
    s = int(s)
    return "%d:%02d:%02d" % (s // 3600, s % 3600 // 60, s % 60)


def collector_performance(cleaned, act, num2clients):
    """يرجع (daily, total, info). كل عميل يُحسب مرة واحدة."""
    matched, owner, info = attribute_calls(cleaned, act)
    a = act.copy()
    a["date"] = pd.to_datetime(a["Created on"]).dt.date
    a["client"] = a["Debtor"].map(client_id)
    known = set()
    for s in num2clients.values():
        known |= s
    a["no_num"] = ~a["client"].isin(known)

    # كل عميل-يوم يُنسب لكل محصّل عنده صف (للعدّ في "بدون أرقام")
    act_cd = a.drop_duplicates(["client", "date", COLLECTOR_COL])

    # لكل (مكالمة، محصّل): صف واحد للمحاولات/الوقت
    calls1 = matched.drop_duplicates(["call_id", "collector"])
    calls1 = calls1.assign(total=calls1["ring"] + calls1["talk"])
    # للعملاء: ردّ على العميل في اليوم لو في أي مكالمة صادرة Answered
    cd = matched.groupby(["collector", "date", "client"]).agg(
        ans=("answered", "any")).reset_index()

    def build(keys):
        g = calls1.groupby(keys).agg(
            attempts=("call_id", "size"),
            answered=("answered", "sum"),
            ring=("ring", "sum"), talk=("talk", "sum")).reset_index()
        g["unanswered"] = g["attempts"] - g["answered"]
        g["total"] = g["ring"] + g["talk"]
        return g

    daily = build(["collector", "date"])
    c = cd.groupby(["collector", "date"]).agg(
        clients_contacted=("client", "nunique"),
        clients_answered=("ans", "sum")).reset_index()
    daily = daily.merge(c, on=["collector", "date"], how="outer")
    nn = act_cd.groupby([COLLECTOR_COL, "date"]).agg(
        activity_rows=("client", "size"),
        clients_no_numbers=("no_num", "sum")).reset_index().rename(
        columns={COLLECTOR_COL: "collector"})
    daily = nn.merge(daily, on=["collector", "date"], how="outer")

    # الإجمالي: كل عميل مرة واحدة على مستوى الفترة
    tot = build(["collector"])
    cdt = cd.groupby(["collector", "client"]).agg(ans=("ans", "any")).reset_index()
    ct = cdt.groupby("collector").agg(
        clients_contacted=("client", "nunique"),
        clients_answered=("ans", "sum")).reset_index()
    nnt = act_cd.drop_duplicates(["client", COLLECTOR_COL]).groupby(
        COLLECTOR_COL).agg(
        clients_in_activity=("client", "nunique"),
        clients_no_numbers=("no_num", "sum")).reset_index().rename(
        columns={COLLECTOR_COL: "collector"})
    days = daily[daily["attempts"].fillna(0) > 0].groupby("collector")["date"].nunique()
    total = nnt.merge(tot, on="collector", how="outer").merge(
        ct, on="collector", how="outer")
    total["active_days"] = total["collector"].map(days).fillna(0).astype(int)

    for df in (daily, total):
        for col in ["attempts", "answered", "unanswered", "ring", "talk", "total",
                    "clients_contacted", "clients_answered", "clients_no_numbers",
                    "activity_rows", "clients_in_activity"]:
            if col in df.columns:
                df[col] = df[col].fillna(0).astype(int)
        df["clients_not_answered"] = df["clients_contacted"] - df["clients_answered"]
        for col in ["ring", "talk", "total"]:
            df[col + "_hms"] = df[col].map(_fmt_hms)
    total["avg_total_per_day_hms"] = [
        _fmt_hms(t / d) if d else "0:00:00"
        for t, d in zip(total["total"], total["active_days"])]
    daily = daily.sort_values(["collector", "date"]).reset_index(drop=True)
    total = total.sort_values("total", ascending=False).reset_index(drop=True)
    return daily, total, info


AR = {
    "collector": "المحصّل", "date": "اليوم", "attempts": "عدد المحاولات",
    "answered": "تم الرد", "unanswered": "لم يتم الرد",
    "total_hms": "الوقت الكلي (رنين + مكالمة)", "ring_hms": "وقت الرنين",
    "talk_hms": "وقت المكالمة", "clients_contacted": "العملاء المتصل بهم (مرة لكل عميل)",
    "clients_answered": "عملاء ردّوا", "clients_not_answered": "عملاء لم يردّوا",
    "clients_no_numbers": "عملاء بدون أرقام تواصل", "activity_rows": "صفوف النشاط",
    "active_days": "أيام العمل", "avg_total_per_day_hms": "متوسط الوقت الكلي/يوم",
    "clients_in_activity": "عملاء في النشاط",
}
DAILY_COLS = ["collector", "date", "total_hms", "ring_hms", "talk_hms", "attempts",
              "answered", "unanswered", "clients_contacted", "clients_answered",
              "clients_not_answered", "clients_no_numbers", "activity_rows"]
TOTAL_COLS = ["collector", "active_days", "total_hms", "avg_total_per_day_hms",
              "ring_hms", "talk_hms", "attempts", "answered", "unanswered",
              "clients_contacted", "clients_answered", "clients_not_answered",
              "clients_in_activity", "clients_no_numbers"]


def short_name(n):
    return str(n).split(" II ")[-1].strip()


def build_collectors_workbook(daily, total, info):
    """ملف إكسل منسّق للمشاركة: ملخص الفترة + يوم بيوم + ملاحظات التغطية."""
    out = io.BytesIO()
    d = daily.copy()
    t = total.copy()
    d["collector"] = d["collector"].map(short_name)
    t["collector"] = t["collector"].map(short_name)
    d["date"] = d["date"].astype(str)
    with pd.ExcelWriter(out, engine="xlsxwriter") as w:
        wb = w.book
        hdr = wb.add_format({"bold": True, "bg_color": "#1F4E78",
                             "font_color": "white", "align": "center",
                             "valign": "vcenter", "text_wrap": True, "border": 1})
        cell = wb.add_format({"align": "center", "border": 1})
        name = wb.add_format({"bold": True, "border": 1, "align": "right"})
        key = wb.add_format({"bold": True, "bg_color": "#FFF2CC",
                             "align": "center", "border": 1})
        for sheet, df, cols in (("ملخص الفترة", t, TOTAL_COLS),
                                ("يوم بيوم", d, DAILY_COLS)):
            out_df = df[cols].rename(columns=AR)
            out_df.to_excel(w, sheet_name=sheet, index=False)
            ws = w.sheets[sheet]
            ws.right_to_left()
            ws.freeze_panes(1, 1)
            ws.set_row(0, 36)
            for i, h in enumerate(out_df.columns):
                ws.write(0, i, h, hdr)
                ws.set_column(i, i, 24 if i == 0 else 16)
            for r, row in enumerate(out_df.itertuples(index=False), start=1):
                for i, v in enumerate(row):
                    f = name if i == 0 else (key if "الوقت الكلي" in out_df.columns[i] else cell)
                    ws.write(r, i, v, f)
            ws.autofilter(0, 0, len(out_df), len(out_df.columns) - 1)
        ws = wb.add_worksheet("ملاحظات التغطية")
        ws.right_to_left()
        ws.set_column(0, 0, 70)
        ws.set_column(1, 1, 14)
        notes = [
            ("مكالمات صادرة (بدون Call Script)", info["out_calls"]),
            ("منسوبة لمحصّل", info["out_attributed"]),
            ("بدون عميل مطابق في ملف الأرقام", info["out_no_client"]),
            ("العميل بلا صف نشاط في نفس اليوم (غير منسوبة)", info["out_client_no_activity_that_day"]),
            ("عميل-يوم له أكثر من محصّل", info["client_days_multi_collector"]),
            ("مكالمات Call Script الآلية", info["script_calls"]),
            ("منها منسوبة للمحصّل صاحب العميل (محاولات)", info["script_attributed"]),
        ]
        ws.write(0, 0, "البند", hdr); ws.write(0, 1, "العدد", hdr)
        for i, (k, v) in enumerate(notes, start=1):
            ws.write(i, 0, k, name); ws.write(i, 1, int(v), cell)
        ws.write(len(notes) + 2, 0,
                 "الوقت الكلي = وقت الرنين + وقت المكالمة لكل المكالمات الصادرة، "
                 "حتى التي لم يُرد عليها.")
    return out.getvalue()


def render_collectors_page(st, cleaned, act, num2clients):
    daily, total, info = collector_performance(cleaned, act, num2clients)
    st.subheader("أداء المحصّلين")
    c = st.columns(4)
    c[0].metric("عدد المحصّلين", len(total))
    c[1].metric("إجمالي المحاولات", int(total["attempts"].sum()))
    c[2].metric("مكالمات منسوبة", f'{info["out_attributed"]:,} / {info["out_calls"]:,}')
    c[3].metric("مكالمات Call Script (منسوبة للمحصّل)", f'{info["script_attributed"]:,} / {info["script_calls"]:,}')
    st.caption("الوقت الكلي = وقت الرنين + وقت المكالمة، شامل المكالمات التي لم يُرد عليها. "
               "كل عميل يُحسب مرة واحدة لكل محصّل.")
    st.markdown("#### ملخص الفترة (مرتّب بالوقت الكلي)")
    t = total.copy(); t["collector"] = t["collector"].map(short_name)
    st.dataframe(t[TOTAL_COLS].rename(columns=AR), use_container_width=True,
                 hide_index=True)
    try:
        st.bar_chart(total.assign(c=total["collector"].map(short_name),
                                  hours=total["total"] / 3600).set_index("c")["hours"])
    except Exception:
        pass
    st.markdown("#### يوم بيوم")
    names = ["الكل"] + [short_name(x) for x in total["collector"]]
    pick = st.selectbox("المحصّل", names)
    d = daily.copy(); d["short"] = d["collector"].map(short_name)
    if pick != "الكل":
        d = d[d["short"] == pick]
    d["date"] = d["date"].astype(str)
    d = d.assign(collector=d["short"])
    st.dataframe(d[DAILY_COLS].rename(columns=AR), use_container_width=True,
                 hide_index=True)
    with st.expander("التغطية والملاحظات"):
        st.write(f"- مكالمات صادرة (بدون Call Script): {info['out_calls']:,}")
        st.write(f"- منسوبة لمحصّل: {info['out_attributed']:,}")
        st.write(f"- بدون عميل مطابق في ملف الأرقام: {info['out_no_client']:,}")
        st.write(f"- العميل بلا صف نشاط في نفس اليوم (غير منسوبة): {info['out_client_no_activity_that_day']:,}")
        st.write(f"- عميل-يوم له أكثر من محصّل: {info['client_days_multi_collector']} "
                 f"(المكالمات تُنسب للمحصّل صاحب أول صف نشاط في اليوم)")
        st.write(f"- مكالمات Call Script الآلية: {info['script_calls']:,} (منسوبة للمحصّل صاحب العميل: {info['script_attributed']:,}، وتُحسب محاولات)")
    st.download_button("تحميل ملخص المحصّلين (Excel)",
                       build_collectors_workbook(daily, total, info),
                       "collectors_performance.xlsx",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


THEME_CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700&display=swap');
:root { --p:#4F008C; --p2:#7B2FBE; --soft:#F3EAFB; --accent:#FF375E; }
html, body, [class*="css"], .stApp { font-family:'Cairo','Segoe UI',sans-serif; direction:rtl; }
.stApp { background:#FAF7FD; }
header[data-testid="stHeader"] { background:transparent; }
.hdr { background:linear-gradient(135deg,#4F008C 0%,#7B2FBE 100%); color:#fff;
  padding:26px 32px; border-radius:16px; margin-bottom:24px;
  border-bottom:4px solid var(--accent); box-shadow:0 6px 18px rgba(79,0,140,.25); }
.hdr-t { font-size:2rem; font-weight:700; }
.hdr-s { font-size:1rem; opacity:.85; margin-top:4px; }
h2, h3, h4 { color:var(--p) !important; font-weight:700 !important; }
section[data-testid="stSidebar"] { background:var(--soft); border-left:1px solid #E1CDF3; }
.stButton > button, .stDownloadButton > button { background:var(--p); color:#fff; border:0;
  border-radius:10px; padding:.55rem 1.6rem; font-weight:600; }
.stButton > button:hover, .stDownloadButton > button:hover { background:var(--p2); color:#fff; }
[data-testid="stFileUploader"] section { background:#fff; border:1.5px dashed var(--p2); border-radius:12px; }
[data-testid="stFileUploader"] label p { color:var(--p); font-weight:600; }
[data-testid="stDataFrame"] { border:1px solid #E1CDF3; border-radius:10px; }
div[role="radiogroup"] label[data-baseweb="radio"] div:first-child { border-color:var(--p) !important; }
</style>
"""


def apply_theme(st):
    st.markdown(THEME_CSS, unsafe_allow_html=True)


# ---------- واجهة Streamlit ----------
def main():
    import streamlit as st
    st.set_page_config(page_title="نظام متابعة التحصيل", page_icon="📞", layout="wide")
    apply_theme(st)
    st.markdown(
        '<div class="hdr"><div class="hdr-t">نظام متابعة التحصيل</div>'
        '<div class="hdr-s">تقرير النشاط والمكالمات</div></div>',
        unsafe_allow_html=True)

    page = st.sidebar.radio("الصفحة", ["ربط النشاط بالمكالمات", "أداء المحصّلين"])
    act_f = st.file_uploader("1) شيت النشاط (xlsx)", type=["xlsx"])
    phones_f = st.file_uploader("2) شيت الأرقام: رقم -> عميل (xlsx)", type=["xlsx"])
    calls_f = st.file_uploader("3) ملف المكالمات (csv أو xlsx)", type=["csv", "xlsx"])

    if page == "أداء المحصّلين":
        if act_f and phones_f and calls_f and st.button("تشغيل"):
            with st.spinner("جاري المعالجة..."):
                n2c = build_phone_map(phones_f)
                cleaned = clean_calls(read_calls(calls_f, calls_f.name), n2c)
                act = pd.read_excel(io.BytesIO(act_f.getvalue()), dtype=str)
                st.session_state["coll"] = (cleaned, act, n2c)
        if "coll" in st.session_state:
            render_collectors_page(st, *st.session_state["coll"])
        return

    if act_f and phones_f and calls_f and st.button("تشغيل"):
        with st.spinner("جاري المعالجة..."):
            enriched, cleaned, missing, summary = process(
                io.BytesIO(act_f.getvalue()), phones_f, calls_f, calls_f.name)
            data = build_output_workbook(enriched, cleaned, missing)

        st.subheader("معاينة النتيجة")
        st.dataframe(enriched[ADDED_COLS].head(50))

        st.subheader("عملاء في النشاط بدون أرقام في شيت الأرقام")
        if len(missing):
            st.dataframe(missing)
        else:
            st.write("لا يوجد — كل العملاء لهم أرقام.")

        st.download_button(
            "تحميل الشيت المحدّث", data, "activity_updated_daymatch.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


if __name__ == "__main__":
    main()
