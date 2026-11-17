# /// script
# requires-python = ">=3.11"
# dependencies = ["openpyxl>=3.1"]
# ///
"""Write the Nameer Water Company file pack into labs/files/. Fixed seed: the same files on every run.

    uv run labs/_tools/prepare_files.py
"""
import csv
import datetime as dt
import random
import re
import zipfile
import sys
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font

sys.path.insert(0, str(Path(__file__).parent))
import texts as T  # noqa: E402

SEED = 2202
LABS = Path(__file__).resolve().parents[1]
OUT = LABS / "files"
FIXED = dt.datetime(2026, 11, 1)

DEPTS = [  # department, respondents, mean per dimension D1..D6 on a 1..5 agreement scale
    ("العمليات والشبكات", 92, [2.4, 2.1, 2.6, 2.1, 2.3, 1.8]),
    ("خدمة العملاء والفوترة", 70, [2.6, 2.5, 2.3, 2.9, 2.5, 2.3]),
    ("تقنية المعلومات", 24, [2.8, 2.6, 3.4, 2.8, 2.7, 2.2]),
    ("المالية", 21, [2.7, 3.1, 2.6, 2.6, 2.5, 1.9]),
    ("الموارد البشرية والخدمات المشتركة", 16, [2.5, 2.4, 2.2, 2.4, 2.9, 1.8]),
    ("الاستراتيجية والتميز المؤسسي", 9, [3.0, 2.5, 2.4, 2.6, 2.8, 2.6]),
    ("مكتب البيانات والذكاء الاصطناعي", 4, [3.6, 2.4, 2.8, 3.4, 3.5, 4.3]),
]
ITEMS = [
    ("Q01", "D1", "أعرف أهداف الشركة التي تحتاج إلى البيانات"),
    ("Q02", "D1", "أعرف من يقرر في بيانات إدارتي ومن يصححها"),
    ("Q03", "D2", "أثق في دقة البيانات التي أستخدمها في عملي"),
    ("Q04", "D2", "حين أجد خطأ في البيانات أعرف لمن أبلغ عنه"),
    ("Q05", "D3", "أصل إلى البيانات التي أحتاجها دون طلب يدوي من إدارة أخرى"),
    ("Q06", "D3", "الأنظمة التي أستخدمها تتبادل البيانات دون نسخ يدوي"),
    ("Q07", "D4", "أعرف مستوى تصنيف البيانات التي أتعامل معها"),
    ("Q08", "D4", "أعرف ما يجوز مشاركته من بيانات العملاء وما لا يجوز"),
    ("Q09", "D5", "تلقيت تدريبًا على استخدام البيانات خلال السنة الماضية"),
    ("Q10", "D5", "يُطلب مني أن أدعم قراراتي بالبيانات"),
    ("Q11", "D6", "أعرف استخدامًا واحدًا على الأقل للذكاء الاصطناعي في الشركة يعمل اليوم"),
    ("Q12", "D6", "أعرف لمن أرجع إذا أخطأ نظام ذكاء اصطناعي في عملي"),
]
MANAGERS = [  # role (no names: the survey of managers is by role), self-rating D1..D6 on the 0..5 scale
    ("مدير العمليات والشبكات", [3, 4, 3, 3, 3, 3]),
    ("مديرة خدمة العملاء والفوترة", [2, 2, 2, 3, 2, 2]),
    ("مدير تقنية المعلومات", [3, 3, 4, 3, 3, 2]),
    ("المدير المالي", [3, 3, 3, 3, 2, 1]),
    ("مدير الموارد البشرية والخدمات المشتركة", [3, 2, 2, 3, 3, 1]),
    ("مدير الاستراتيجية والتميز المؤسسي", [3, 3, 3, 3, 3, 3]),
    ("مديرة مكتب البيانات والذكاء الاصطناعي", [2, 2, 2, 2, 2, 1]),
]


def book():
    wb = Workbook()
    wb.properties.created = FIXED
    wb.properties.modified = FIXED
    wb.properties.creator = "SDAIA Academy"
    wb.remove(wb.active)
    return wb


def sheet(wb, title, header, rows, widths=None):
    ws = wb.create_sheet(title)
    ws.sheet_view.rightToLeft = True
    ws.append(header)
    for c in ws[1]:
        c.font = Font(bold=True)
    for r in rows:
        ws.append(list(r))
    for i, w in enumerate(widths or [18] * len(header)):
        ws.column_dimensions[chr(65 + i)].width = w
    for row in ws.iter_rows():
        for c in row:
            c.alignment = Alignment(wrap_text=True, vertical="top")
    return ws


def save(wb, path):
    """Save, then rewrite the zip with fixed dates so the same workbook gives the same bytes."""
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    with zipfile.ZipFile(path) as z:
        parts = [(i.filename, z.read(i.filename)) for i in z.infolist()]
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in parts:
            if name == "docProps/core.xml":
                data = re.sub(rb"(<dcterms:modified[^>]*>)[^<]*", rb"\g<1>2026-11-01T00:00:00Z", data)
            z.writestr(zipfile.ZipInfo(name, date_time=(2026, 11, 1, 0, 0, 0)), data, zipfile.ZIP_DEFLATED)


def write(path, text):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text.strip() + "\n", encoding="utf-8", newline="\n")


def notes_md(title, notes):
    lines = [f"# {title}", "", "ملاحظات مرجعية كما يدوّنها مقيّم دقيق. كل سطر: الرقم، والعلامة، والملاحظة.", ""]
    lines += [f"- {nid} [{tag}] {text}" for nid, tag, text in notes]
    return "\n".join(lines)


def survey(rng):
    rows = []
    rid = 0
    for dept, n, means in DEPTS:
        for _ in range(n):
            rid += 1
            band = "مشرف" if rng.random() < 0.15 else "موظف"
            ans = []
            for _q, dim, _t in ITEMS:
                m = means[int(dim[1]) - 1] + (0.4 if band == "مشرف" else 0.0)
                v = round(rng.gauss(m, 0.85))
                ans.append(min(5, max(1, v)))
            rows.append([f"R{rid:04d}", dept, band] + ans)
    return rows


def main():
    rng = random.Random(SEED)
    write(OUT / "model_extracts.md", T.MODEL_EXTRACTS)
    write(OUT / "org_chart.md", T.ORG_CHART)
    write(OUT / "sponsor_brief.md", T.SPONSOR_BRIEF)
    write(OUT / "roles" / "role_ops_director.md", T.ROLE_OPS)
    write(OUT / "roles" / "role_cs_director.md", T.ROLE_CS)
    write(OUT / "start" / "interview_guide.md", T.INTERVIEW_GUIDE)
    write(OUT / "start" / "interview_notes_ops.md", notes_md("ملاحظات مقابلة مدير العمليات والشبكات", T.NOTES_OPS))
    write(OUT / "start" / "interview_notes_cs.md", notes_md("ملاحظات مقابلة مديرة خدمة العملاء والفوترة", T.NOTES_CS))
    write(OUT / "challenge_cards.md", T.CHALLENGES)
    write(OUT / "followup_template.md", T.FOLLOWUP)

    for d in T.DOCS:
        did, title, typ, date, owner, status, ver, dims, body = d
        head = f"# {did} · {title}\n\n| النوع | التاريخ | المالك | الحالة | الإصدار |\n|---|---|---|---|---|\n| {typ} | {date} | {owner} | {status} | {ver or '-'} |\n\n"
        write(OUT / "documents" / f"{did}.md", head + body)

    wb = book()
    sheet(wb, "الفهرس", ["الرمز", "العنوان", "النوع", "التاريخ", "المالك", "الحالة", "الإصدار"],
          [d[:7] for d in T.DOCS], [8, 48, 14, 12, 30, 12, 8])
    save(wb, OUT / "document_index.xlsx")

    # Lab 1.1: the plan template
    wb = book()
    sheet(wb, "مقارنة النماذج", ["المعيار", "النموذج أ", "النموذج ب", "النموذج ج"],
          [("يغطي إدارة البيانات", "", "", ""), ("يغطي الذكاء الاصطناعي", "", "", ""),
           ("يعطي مستويات أستطيع القياس عليها", "", "", ""), ("يناسب جهة سعودية وتنظيماتها", "", "", ""),
           ("يسمح ترخيصه بالتكييف والنشر الداخلي", "", "", ""), ("يجيب عن أسئلة الراعي الثلاثة", "", "", "")],
          [34, 22, 22, 22])
    sheet(wb, "الأبعاد بعد التكييف", ["الرمز", "البُعد", "ما يغطيه", "من أي نموذج"], [("", "", "", "")] * 6, [8, 30, 50, 30])
    sheet(wb, "خطة الأدلة", ["الرقم", "البُعد", "المصدر (مقابلة، وثيقة، استبانة، جولة)", "مع من أو من أين", "السؤال أو الطلب"],
          [], [8, 10, 26, 26, 60])
    save(wb, OUT / "evidence_plan.xlsx")

    # Lab 1.3: evidence log template
    wb = book()
    sheet(wb, "سجل الأدلة", ["الرمز", "البُعد", "ما يثبته", "قوة الدليل (قوي، متوسط، ضعيف، ليس دليلًا)", "الاتجاه (يرفع، يخفض)", "يتعارض مع"],
          [], [10, 8, 50, 22, 14, 16])
    save(wb, OUT / "evidence_log.xlsx")

    # Lab 1.4: survey and self-assessment
    with open(OUT / "staff_survey.csv", "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f, lineterminator="\n")
        w.writerow(["respondent", "department", "band"] + [q for q, _d, _t in ITEMS])
        srows = survey(rng)
        w.writerows(srows)
    wb = book()
    sheet(wb, "البنود", ["البند", "البُعد", "العبارة (1 لا أوافق إطلاقًا، 5 أوافق تمامًا)"], ITEMS, [8, 8, 70])
    sheet(wb, "قاعدة العرض", ["القاعدة"], [("لا تُعرض نتيجة أي مجموعة يقل عدد المستجيبين فيها عن 5. تُكتب «أقل من 5» مكان الأرقام.",),
                                            ("الفرق بين متوسطي مجموعتين لا يُقرأ نتيجةً ما لم يكن في كل منهما 30 مستجيبًا على الأقل.",)], [90])
    save(wb, OUT / "survey_items.xlsx")
    wb = book()
    sheet(wb, "التقييم الذاتي", ["المدير"] + [f"{d[0]} {d[1]}" for d in T.DIMENSIONS],
          [[m] + r for m, r in MANAGERS], [36] + [16] * 6)
    save(wb, OUT / "manager_self_assessment.xlsx")

    # Reference states for the later labs
    wb = book()
    sheet(wb, "المستويات", ["المستوى", "الاسم", "الوصف"], T.LEVELS, [8, 18, 80])
    sheet(wb, "القواعد", ["القاعدة"], [(r,) for r in T.RULES], [100])
    sheet(wb, "الثقة", ["المستوى", "متى"], T.CONFIDENCE, [12, 80])
    sheet(wb, "الأبعاد", ["الرمز", "البُعد", "ما يغطيه", "من النموذج أ", "من النموذج ب"], T.DIMENSIONS, [8, 28, 50, 40, 34])
    sheet(wb, "الملف", ["البُعد", "الدرجة (0 إلى 5)", "الثقة", "مراجع الأدلة", "لماذا لا تكون أعلى"],
          [(f"{d[0]} {d[1]}", "", "", "", "") for d in T.DIMENSIONS], [30, 14, 10, 40, 50])
    save(wb, OUT / "start" / "scoring_guide.xlsx")

    ev = []
    for nid, tag, text in T.NOTES_OPS + T.NOTES_CS:
        ev.append((nid, "مقابلة", tag, text))
    for d in T.DOCS:
        ev.append((d[0], "وثيقة", d[5], d[1]))
    for k, (q, _dim, text) in enumerate(ITEMS):
        m = sum(r[3 + k] for r in srows) / len(srows)
        ev.append((f"S-{q}", "استبانة", f"المتوسط {m:.2f} من {len(srows)} إجابة", text))
    wb = book()
    sheet(wb, "الأدلة", ["الرمز", "النوع", "الحالة أو العلامة", "المحتوى"], ev, [12, 10, 14, 90])
    save(wb, OUT / "start" / "evidence_log_reference.xlsx")

    wb = book()
    sheet(wb, "درجات صالح", ["البُعد", "الدرجة"], [(f"{d} {n}", s) for (d, s), (_x, n, *_r) in zip(T.PEER, T.DIMENSIONS)], [34, 10])
    save(wb, OUT / "start" / "peer_scores.xlsx")

    wb = book()
    sheet(wb, "الملف", ["البُعد", "الدرجة", "الثقة", "مراجع الأدلة", "الخلاصة", "ما ينقص للمستوى التالي"],
          [(f"{p[0]} {d[1]}",) + p[1:] for p, d in zip(T.PROFILE, T.DIMENSIONS)], [30, 8, 10, 40, 60, 40])
    save(wb, OUT / "start" / "profile_reference.xlsx")

    gaps = []
    for (dim, tgt, by, why), p, d in zip(T.TARGETS, T.PROFILE, T.DIMENSIONS):
        gaps.append((f"{dim} {d[1]}", p[1], tgt, by, tgt - p[1], why, p[5]))
    wb = book()
    sheet(wb, "سجل الفجوات", ["البُعد", "الحالي", "المستهدف", "بحلول", "الفجوة", "لماذا هذا المستهدف", "ما يسد الفجوة"], gaps,
          [30, 8, 8, 8, 8, 60, 50])
    save(wb, OUT / "start" / "gap_register_reference.xlsx")

    wb = book()
    sheet(wb, "المبادرات", ["الرمز", "المبادرة", "الأبعاد", "التكلفة (ألف ريال)", "جهد مكتب البيانات (شهر عمل)", "تعتمد على"],
          [(i[0], i[1], i[2], i[3], i[4], "، ".join(i[5]) or "-") for i in T.INITIATIVES], [8, 56, 10, 14, 16, 14])
    save(wb, OUT / "initiatives.xlsx")
    wb = book()
    sheet(wb, "القدرة", ["الأفق", "الوصف", "المدة (شهر)", "الميزانية (ألف ريال)", "جهد مكتب البيانات المتاح (شهر عمل)"],
          T.HORIZONS, [8, 30, 12, 18, 22])
    save(wb, OUT / "capacity.xlsx")

    wb = book()
    order = {"H1": 1, "H2": 2, "H3": 3}
    rows = sorted(((T.ROADMAP[i[0]], i[0], i[1], i[3], i[4], "، ".join(i[5]) or "-", i[6], i[7]) for i in T.INITIATIVES),
                  key=lambda r: (order[r[0]], r[1]))
    sheet(wb, "خارطة الطريق", ["الأفق", "الرمز", "المبادرة", "التكلفة", "الجهد", "تعتمد على", "المسؤول", "المؤشر"], rows,
          [6, 6, 50, 10, 8, 12, 30, 36])
    save(wb, OUT / "start" / "roadmap_reference.xlsx")
    print("wrote", sum(1 for p in OUT.rglob("*") if p.is_file()), "files into", OUT.name)


if __name__ == "__main__":
    main()
