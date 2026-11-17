# /// script
# requires-python = ">=3.11"
# dependencies = ["openpyxl>=3.1"]
# ///
"""Recompute every number the labs and their debrief slides quote, from labs/files, and check it.

    uv run labs/solutions/check.py          # all labs
    uv run labs/solutions/check.py 1.4 2.2  # some labs

Writes nothing to disk and calls no model: the model steps use the scripted stand-in in standin.json.
"""
import csv
import json
import re
import sys
from collections import defaultdict
from pathlib import Path

from openpyxl import load_workbook

HERE = Path(__file__).resolve().parent
FILES = HERE.parent / "files"
STANDIN = json.loads((HERE / "standin.json").read_text(encoding="utf-8"))
DIMS = ["D1", "D2", "D3", "D4", "D5", "D6"]
NAMES = ["يزيد", "وسن", "جُمانة", "رُواء"]
PSEUDONYMS = [("م. يزيد", "مدير العمليات"), ("يزيد", "مدير العمليات"), ("وسن", "مديرة خدمة العملاء"),
              ("جُمانة", "محللة الفوترة"), ("المهندسة رُواء", "مهندسة الشبكات (أ)"), ("رُواء", "مهندسة الشبكات (أ)")]
MIN_GROUP = 5


def rows(path, sheet=None):
    wb = load_workbook(FILES / path, read_only=True)
    ws = wb[sheet] if sheet else wb.worksheets[0]
    out = [r for r in ws.iter_rows(values_only=True)]
    return out[0], [r for r in out[1:] if any(c is not None for c in r)]


def notes():
    out = {}
    for f in ("interview_notes_ops.md", "interview_notes_cs.md"):
        for line in (FILES / "start" / f).read_text(encoding="utf-8").splitlines():
            m = re.match(r"- (N-[A-Z]+-\d\d) \[(.+?)\] (.*)", line)
            if m:
                out[m.group(1)] = (m.group(2), m.group(3))
    return out


def lab_1_1():
    text = (FILES / "model_extracts.md").read_text(encoding="utf-8")
    assert text.count("## النموذج") == 3
    assert "ما لا يغطيه: استخدام الذكاء الاصطناعي" in text, "model A must say it leaves AI out"
    assert "CC BY-NC-SA" in text
    # The reference plan: (dimension, source type, with whom or what)
    plan = [
        ("D1", "مقابلة", "المديران"), ("D1", "وثيقة", "استراتيجيات ومواثيق ومحاضر اللجنة"), ("D1", "استبانة", "من يقرر في البيانات"),
        ("D2", "مقابلة", "المديران: مثال خطأ وتصحيحه"), ("D2", "وثيقة", "سياسة الجودة وتقارير القياس ومستخلصات الأنظمة ودليل البيانات"), ("D2", "استبانة", "الثقة في الدقة والإبلاغ عن الخطأ"),
        ("D3", "مقابلة", "مديرة خدمة العملاء: كيف تصل التقارير"), ("D3", "وثيقة", "المعمارية ومواصفات التكامل وسجلاتها"), ("D3", "استبانة", "الوصول دون طلب يدوي"), ("D3", "جولة", "تقنية المعلومات: تشغيل التكامل"),
        ("D4", "مقابلة", "المديران: هل صُنّفت البيانات"), ("D4", "وثيقة", "سجل التصنيف وقائمة الامتثال وإشعار الخصوصية"), ("D4", "استبانة", "معرفة مستوى التصنيف"),
        ("D5", "مقابلة", "من يحلل البيانات"), ("D5", "وثيقة", "خطة التدريب والحضور والوظائف"), ("D5", "استبانة", "التدريب وطلب الدليل في القرار"),
        ("D6", "مقابلة", "المديران: استخدام يعمل اليوم ومالكه"), ("D6", "وثيقة", "تقارير التجارب والعروض"), ("D6", "استبانة", "معرفة استخدام يعمل ولمن يُرجع"),
    ]
    by = defaultdict(set)
    for d, src, _ in plan:
        by[d].add(src)
    for d in DIMS:
        assert len(by[d]) >= 2, d
    print(f"1.1  reference plan: {len(plan)} items, every dimension has 2 or more source types "
          f"(min {min(len(v) for v in by.values())}); model A leaves AI out; model C licence NC-SA")


def lab_1_2():
    n = notes()
    tags = defaultdict(int)
    for t, _ in n.values():
        tags[t] += 1
    assert len(n) == 20, len(n)
    assert "المستوى 4" in n["N-OPS-01"][1] and "لا تقرير جودة" in n["N-OPS-02"][1]
    print(f"1.2  reference notes: {len(n)}; tags {dict(tags)}; claim N-OPS-01 (level 4) has no artefact behind it (N-OPS-02)")


def lab_1_3():
    head, docs = rows("document_index.xlsx")
    status = defaultdict(list)
    for d in docs:
        status[d[5]].append(d[0])
    assert len(docs) == 20
    assert sorted(status["مسودة"]) == ["D02", "D05"]
    body = {p.stem: p.read_text(encoding="utf-8") for p in (FILES / "documents").glob("D*.md")}
    owned, total = (int(x) for x in (re.search(r"لها مالك معيّن: (\d+)", body["D08"]).group(1),
                                     re.search(r"المسجلة في الدليل: (\d+)", body["D08"]).group(1)))
    est = int(re.search(r"التقديرية: ([\d,]+)", body["D07"]).group(1).replace(",", ""))
    allr = int(re.search(r"المسجلة في الربع الثالث 2026: ([\d,]+)", body["D07"]).group(1).replace(",", ""))
    ticked, spec = 120, 140
    noev = int(re.search(r"بلا أي دليل مرفق: (\d+)", body["D12"]).group(1))
    assert "140 مواصفة" in body["D12"] and f": {ticked} (" in body["D12"]
    held = re.search(r"عُقد اجتماعان من ستة", body["D04"]) is not None
    assert held and owned == 41 and total == 312 and noev == 48
    not_evidence = ["D19", "D20", "D09"]
    print(f"1.3  20 documents; drafts {status['مسودة']}; not evidence {not_evidence} (claim, invitation, outdated); "
          f"catalogue {owned}/{total} owned ({owned/total:.0%}); estimated readings {est/allr:.0%}; "
          f"compliance ticked {ticked}/{spec} ({ticked/spec:.0%}), {noev} of {ticked} without evidence ({noev/ticked:.0%}); committee 2 of 6")
    return ticked / spec


def lab_1_4():
    _, items = rows("survey_items.xlsx", "البنود")
    dim_of = {q: d for q, d, _ in items}
    with open(FILES / "staff_survey.csv", encoding="utf-8-sig") as f:
        data = list(csv.DictReader(f))
    groups = defaultdict(list)
    for r in data:
        groups[r["department"]].append(r)
    print(f"1.4  survey: {len(data)} responses in {len(groups)} departments")
    table = {}
    for dept, rs in groups.items():
        n = len(rs)
        if n < MIN_GROUP:
            print(f"     {dept}: n={n} -> أقل من 5 (suppressed)")
            table[dept] = None
            continue
        means = []
        for d in DIMS:
            vals = [int(r[q]) for r in rs for q in dim_of if dim_of[q] == d]
            means.append(sum(vals) / len(vals))
        table[dept] = means
        print(f"     {dept}: n={n} " + " ".join(f"{d}={m:.1f}" for d, m in zip(DIMS, means)))
    small = [d for d, rs in groups.items() if len(rs) < MIN_GROUP]
    assert small == ["مكتب البيانات والذكاء الاصطناعي"], small
    allm = []
    for d in DIMS:
        vals = [int(r[q]) for r in data for q in dim_of if dim_of[q] == d]
        allm.append(sum(vals) / len(vals))
    print("     all staff: " + " ".join(f"{d}={m:.2f}" for d, m in zip(DIMS, allm)))
    # Inflation: managers' self-rating against the reference profile
    _, man = rows("manager_self_assessment.xlsx")
    _, prof = rows("start/profile_reference.xlsx")
    ref = [p[1] for p in prof]
    mean_self = [sum(m[i + 1] for m in man) / len(man) for i in range(6)]
    gaps = [s - r for s, r in zip(mean_self, ref)]
    print("     managers' mean self-rating minus reference: " + " ".join(f"{d}={g:+.1f}" for d, g in zip(DIMS, gaps)))
    ops = man[0]
    print(f"     largest single gap: {ops[0]} on D2, {ops[2]} against {ref[1]}")
    assert ops[2] - ref[1] == 3 and sum(g > 0 for g in gaps) == 5
    # Pseudonymise, then check the stand-in synthesis
    n = notes()
    text = "\n".join(f"{k} {v[1]}" for k, v in n.items())
    for a, b in PSEUDONYMS:
        text = text.replace(a, b)
    left = [x for x in NAMES if x in text]
    assert not left, left
    claims = STANDIN["D1_4_synthesis"]["claims"]
    uncited = [c["text"] for c in claims if not c["cites"]]
    bad = []
    for nid, words in STANDIN["D1_4_synthesis"]["unsupported"].items():
        for c in claims:
            if nid in c["cites"] and not any(w in n[nid][1] for w in words) and any(w in c["text"] for w in words):
                bad.append(nid)
    assert len(uncited) == 2 and bad == ["N-OPS-08"], (uncited, bad)
    kept = len(claims) - len(uncited) - len(bad)
    print(f"     names left after pseudonymisation: 0; stand-in synthesis: {len(claims)} claims, "
          f"{len(uncited)} without a note, {len(bad)} citing a note that does not say it -> {kept} kept")
    return ref


def lab_1_5(ref):
    _, peer = rows("start/peer_scores.xlsx")
    p = [r[1] for r in peer]
    exact = sum(a == b for a, b in zip(p, ref))
    within = sum(abs(a - b) <= 1 for a, b in zip(p, ref))
    diffs = {d: b - a for d, a, b in zip(DIMS, ref, p) if a != b}
    # Chance agreement for a uniform guess over 0..5 (the simulation in the work notes gives the same)
    chance = [sum(abs(g - r) <= 1 for g in range(6)) / 6 for r in ref]
    always2 = sum(abs(2 - r) <= 1 for r in ref)
    print(f"1.5  reference profile {dict(zip(DIMS, ref))}; peer {dict(zip(DIMS, p))}")
    print(f"     exact agreement {exact}/6, within one level {within}/6; peer above reference on {diffs}")
    print(f"     chance of within one by a uniform guess, per reference: {[round(c, 2) for c in chance]}; "
          f"an assessor who always says 2 is within one on {always2}/6")
    assert exact == 3 and within == 4 and diffs == {"D2": 2, "D4": 2, "D6": 1} and always2 == 6


def lab_2_1(ref):
    _, gaps = rows("start/gap_register_reference.xlsx")
    g = {r[0].split()[0]: (r[1], r[2], r[4]) for r in gaps}
    assert [g[d][0] for d in DIMS] == ref
    total = sum(v[2] for v in g.values())
    print(f"2.1  gaps {{{', '.join(f'{d}: {v[0]}->{v[1]}' for d, v in g.items())}}}; total {total} levels; "
          f"D3 holds at 3; D6 target 2, the sponsor's 4 would be a gap of {4 - g['D6'][0]} in 18 months")
    assert g["D3"][2] == 0 and g["D6"][1] == 2


def roadmap_problems(plan, inits, caps):
    order = {"H1": 1, "H2": 2, "H3": 3}
    probs = []
    for i, (cost, fte, deps) in inits.items():
        for d in deps:
            if order[plan[d]] >= order[plan[i]]:
                probs.append(f"{i} before or with its dependency {d}")
    for h, (budget, cap) in caps.items():
        c = sum(inits[i][0] for i in inits if plan[i] == h)
        f = sum(inits[i][1] for i in inits if plan[i] == h)
        if c > budget:
            probs.append(f"{h} budget {c} > {budget}")
        if f > cap:
            probs.append(f"{h} effort {f} > {cap}")
    return probs


def lab_2_2():
    _, ini = rows("initiatives.xlsx")
    inits = {r[0]: (r[3], r[4], [] if r[5] == "-" else r[5].split("، ")) for r in ini}
    _, cap = rows("capacity.xlsx")
    caps = {r[0]: (r[3], r[4]) for r in cap}
    _, rm = rows("start/roadmap_reference.xlsx")
    plan = {r[1]: r[0] for r in rm}
    assert all(r[6] and r[7] for r in rm), "every initiative has an owner and an indicator"
    probs = roadmap_problems(plan, inits, caps)
    assert not probs, probs
    for h in caps:
        c = sum(inits[i][0] for i in inits if plan[i] == h)
        f = sum(inits[i][1] for i in inits if plan[i] == h)
        print(f"2.2  {h}: {sorted(i for i in plan if plan[i] == h)} cost {c}/{caps[h][0]} effort {f}/{caps[h][1]}")
    ai_first = dict(plan, I10="H1", I11="H1")
    bad = roadmap_problems(ai_first, inits, caps)
    print(f"     AI first (I10 and I11 in H1): {len(bad)} problems: {bad}")
    assert len(bad) >= 4
    print(f"     total cost {sum(v[0] for v in inits.values())}k SAR of {sum(v[0] for v in caps.values())}k; "
          f"effort {sum(v[1] for v in inits.values())} of {sum(v[1] for v in caps.values())} person-months")


def lab_2_3():
    slides = STANDIN["D2_3_draft_outline"]["slides"]
    asks = [s for s in slides if "قرار" in s]
    answer_first = "النتائج" in slides[0] or "درج" in slides[0]
    print(f"2.3  stand-in draft: {len(slides)} slides, {len(asks)} asking for a decision, answer first: {answer_first}")
    assert not asks and not answer_first


LABS = {"1.1": lab_1_1, "1.2": lab_1_2, "1.3": lab_1_3}


def main(want):
    run = lambda k: not want or k in want  # noqa: E731
    if run("1.1"): lab_1_1()
    if run("1.2"): lab_1_2()
    if run("1.3"): lab_1_3()
    ref = lab_1_4() if run("1.4") or run("1.5") or run("2.1") else None
    if run("1.5"): lab_1_5(ref)
    if run("2.1"): lab_2_1(ref)
    if run("2.2"): lab_2_2()
    if run("2.3"): lab_2_3()
    print("PASS")


if __name__ == "__main__":
    main(set(sys.argv[1:]))
