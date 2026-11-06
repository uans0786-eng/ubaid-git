#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
bottom_align.py — صفحات کو نیچے سے برابر کرنے کا ٹول (Word سے بنی PDF کے لیے)

ہر صفحے پر:
  * پہلی لائن اپنی جگہ رہتی ہے،
  * آخری لائن ہدف (target) والی baseline پر آ جاتی ہے — یعنی وہیں جہاں Word کے
    بھرے ہوئے صفحات کی آخری لائن ہوتی ہے،
  * درمیانی لائنوں میں اضافی جگہ برابر بانٹ دی جاتی ہے۔

لائنوں کو کاٹا نہیں جاتا: ہر حرف کی baseline پڑھ کر اسے اس کی لائن کے گروپ میں
ڈالا جاتا ہے اور پورا گروپ vector حالت میں نیچے کھسکایا جاتا ہے، اس لیے نستعلیق کے
دامن، فونٹ، معیار اور selectable متن سب جوں کے توں رہتے ہیں۔

Usage:
  python bottom_align.py input.pdf output.pdf --bottom-margin 0.5in [--pages 5-120] [--skip 23,47]
"""

import argparse
import math
import re
import sys
from collections import Counter

import pikepdf
from pikepdf import Operator, ContentStreamInstruction

# --------------------------------------------------------------------------- #
#  چھوٹی matrix لائبریری (PDF matrix = [a b c d e f])
# --------------------------------------------------------------------------- #
IDENTITY = (1.0, 0.0, 0.0, 1.0, 0.0, 0.0)


def mmul(m1, m2):
    """m1 × m2 (PDF row-vector convention: point × m1 × m2)."""
    a1, b1, c1, d1, e1, f1 = m1
    a2, b2, c2, d2, e2, f2 = m2
    return (a1 * a2 + b1 * c2, a1 * b2 + b1 * d2,
            c1 * a2 + d1 * c2, c1 * b2 + d1 * d2,
            e1 * a2 + f1 * c2 + e2, e1 * b2 + f1 * d2 + f2)


def apply(m, x, y):
    a, b, c, d, e, f = m
    return (x * a + y * c + e, x * b + y * d + f)


def nums(operands):
    return [float(o) for o in operands]


# --------------------------------------------------------------------------- #
#  اکائیاں اور صفحات کی رینج
# --------------------------------------------------------------------------- #
UNITS = {"in": 72.0, "inch": 72.0, "cm": 72.0 / 2.54, "mm": 72.0 / 25.4, "pt": 1.0}


def parse_length(text):
    """'0.5in', '1.27cm', '12.7mm', '36pt' → points. بغیر اکائی کے = انچ."""
    m = re.fullmatch(r"\s*([0-9]*\.?[0-9]+)\s*([a-zA-Z]*)\s*", text)
    if not m:
        raise argparse.ArgumentTypeError(f"غلط پیمائش: {text!r}")
    unit = (m.group(2) or "in").lower()
    if unit not in UNITS:
        raise argparse.ArgumentTypeError(f"نامعلوم اکائی: {unit!r} (in, cm, mm, pt)")
    return float(m.group(1)) * UNITS[unit]


def parse_ranges(text, npages):
    """'1-10,15,20-30' → {0..9, 14, 19..29} (zero-based)."""
    pages = set()
    for part in text.split(","):
        part = part.strip()
        if not part:
            continue
        if "-" in part:
            a, b = part.split("-", 1)
            a = int(a) if a.strip() else 1
            b = int(b) if b.strip() else npages
            pages.update(range(a - 1, min(b, npages)))
        else:
            pages.add(int(part) - 1)
    return {p for p in pages if 0 <= p < npages}


# --------------------------------------------------------------------------- #
#  Content stream کا تجزیہ
# --------------------------------------------------------------------------- #
SHOW_OPS = {"Tj", "TJ", "'", '"'}
PAINT_OPS = {"S", "s", "f", "F", "f*", "B", "B*", "b", "b*"}
PATH_OPS = {"m", "l", "c", "v", "y", "re", "h"}
POSITION_OPS = {"Td", "TD", "Tm", "T*", "BT"}


def minv(m):
    a, b, c, d, e, f = m
    det = a * d - b * c
    if abs(det) < 1e-12:
        return None
    ia, ib, ic, id_ = d / det, -b / det, -c / det, a / det
    return (ia, ib, ic, id_, -(e * ia + f * ic), -(e * ib + f * id_))


class Item:
    """content stream کی ایک ایسی ہدایت جو صفحے پر کچھ بناتی ہے۔"""
    __slots__ = ("index", "kind", "y", "size", "bbox", "line", "tm", "ctm",
                 "run", "span", "start", "clip")

    def __init__(self, index, kind, **kw):
        self.index = index
        self.kind = kind            # 'text' | 'graphic'
        self.y = kw.get("y")        # text: baseline (PDF y، نیچے سے)
        self.size = kw.get("size")  # text: فونٹ سائز
        self.bbox = kw.get("bbox")  # graphic: (x0, y0, x1, y1)
        self.tm = kw.get("tm")      # text: اس وقت کا text matrix
        self.ctm = kw.get("ctm")    # اس وقت کا CTM
        self.run = kw.get("run")    # text: مسلسل حروف کا سلسلہ (بیچ میں کوئی positioning نہیں)
        self.span = kw.get("span")  # text: ActualText والا گروپ (ایک حرف کے اجزا)
        self.start = kw.get("start", index)  # graphic: path کہاں سے شروع ہوا
        self.clip = kw.get("clip", False)
        self.line = None            # کس لائن سے تعلق ہے؛ None = ساکن


def page_resources(page_obj):
    obj = page_obj
    while obj is not None:
        if "/Resources" in obj:
            return obj.Resources
        obj = obj.get("/Parent")
    return pikepdf.Dictionary()


def analyse(instructions, resources):
    """ہر ڈرائنگ ہدایت کی جگہ اور اس وقت کی حالت (state) معلوم کرو۔"""
    items = []
    ctm, stack = IDENTITY, []
    tm = tlm = IDENTITY
    tfs, tl, trise = 0.0, 0.0, 0.0
    path_pts, path_start, path_clip = [], None, False
    run, new_run = 0, True
    span_stack, span_counter = [], 0
    xobjects = resources.get("/XObject", pikepdf.Dictionary())

    def text_item(i):
        nonlocal run, new_run
        if new_run:
            run += 1
            new_run = False
        trm = mmul(tm, ctm)
        _, y = apply(trm, 0.0, trise)
        scale = math.hypot(trm[2], trm[3])
        span = next((s for s in reversed(span_stack) if s is not None), None)
        items.append(Item(i, "text", y=y, size=abs(tfs) * scale, tm=tm, ctm=ctm,
                          run=run, span=span))

    def next_line():
        nonlocal tm, tlm
        tlm = mmul((1, 0, 0, 1, 0, -tl), tlm)
        tm = tlm

    def box_item(i, m, corners, start=None):
        pts = [apply(m, x, y) for x, y in corners]
        xs, ys = zip(*pts)
        items.append(Item(i, "graphic", bbox=(min(xs), min(ys), max(xs), max(ys)),
                          ctm=ctm, start=i if start is None else start))

    unit_square = [(0, 0), (1, 0), (0, 1), (1, 1)]
    for i, ins in enumerate(instructions):
        if isinstance(ins, pikepdf.ContentStreamInlineImage):
            box_item(i, ctm, unit_square)
            continue
        op = str(ins.operator)
        ops = ins.operands
        if op in POSITION_OPS:
            new_run = True
        if op == "q":
            stack.append(ctm)
        elif op == "Q":
            ctm = stack.pop() if stack else IDENTITY
        elif op == "cm":
            ctm = mmul(tuple(nums(ops)), ctm)
        elif op == "BT":
            tm = tlm = IDENTITY
        elif op == "Tf":
            tfs = float(ops[1])
        elif op == "TL":
            tl = float(ops[0])
        elif op == "Ts":
            trise = float(ops[0])
        elif op == "Td":
            tx, ty = nums(ops)
            tlm = mmul((1, 0, 0, 1, tx, ty), tlm)
            tm = tlm
        elif op == "TD":
            tx, ty = nums(ops)
            tl = -ty
            tlm = mmul((1, 0, 0, 1, tx, ty), tlm)
            tm = tlm
        elif op == "Tm":
            tm = tlm = tuple(nums(ops))
        elif op == "T*":
            next_line()
        elif op in ("Tj", "TJ"):
            text_item(i)
        elif op in ("'", '"'):
            next_line()
            new_run = True
            text_item(i)
        elif op in ("BDC", "BMC"):
            props = ops[1] if op == "BDC" and len(ops) > 1 else None
            if isinstance(props, pikepdf.Dictionary) and "/ActualText" in props:
                span_counter += 1
                span_stack.append(span_counter)
            else:
                span_stack.append(None)
        elif op == "EMC":
            if span_stack:
                span_stack.pop()
        elif op in PATH_OPS:
            if path_start is None:
                path_start = i
            v = nums(ops)
            if op == "re":
                x, y, w, h = v
                path_pts += [apply(ctm, x, y), apply(ctm, x + w, y + h),
                             apply(ctm, x + w, y), apply(ctm, x, y + h)]
            else:
                path_pts += [apply(ctm, v[k], v[k + 1]) for k in range(0, len(v) - 1, 2)]
        elif op in ("W", "W*"):
            path_clip = True
        elif op in PAINT_OPS or op == "n":
            if op != "n" and path_pts:
                xs, ys = zip(*path_pts)
                it = Item(i, "graphic", bbox=(min(xs), min(ys), max(xs), max(ys)),
                          ctm=ctm, start=path_start, clip=path_clip)
                items.append(it)
            path_pts, path_start, path_clip = [], None, False
        elif op == "Do":
            xo = xobjects.get(str(ops[0]))
            if xo is not None and xo.get("/Subtype") == "/Form":
                x0, y0, x1, y1 = nums(xo.get("/BBox", [0, 0, 1, 1]))
                m = mmul(tuple(nums(xo.get("/Matrix", IDENTITY))), ctm)
                box_item(i, m, [(x0, y0), (x1, y0), (x0, y1), (x1, y1)])
            else:
                box_item(i, ctm, unit_square)
    return items


# --------------------------------------------------------------------------- #
#  baselines کی شناخت
# --------------------------------------------------------------------------- #
def find_baselines(text_items):
    """
    نستعلیق میں ایک ہی لائن کے حروف مختلف اونچائی پر ہوتے ہیں (ترچھا پن)، مگر
    اکثر حروف اصل baseline پر بیٹھتے ہیں۔ اس لیے ہر لائن کی baseline وہ اونچائی ہے
    جس پر سب سے زیادہ حروف ہوں۔
    """
    if not text_items:
        return []
    counts = Counter(round(it.y, 1) for it in text_items)
    sizes = sorted(it.size for it in text_items)
    fsize = sizes[len(sizes) // 2] or 10.0
    min_sep = 0.8 * fsize
    ranked = sorted(counts.items(), key=lambda kv: (-kv[1], kv[0]))

    # پہلا مرحلہ: مضبوط (بار بار آنے والی) baselines
    accepted = []
    for y, c in ranked:
        if c < 4:
            break
        if all(abs(y - a) >= min_sep for a in accepted):
            accepted.append(y)
    accepted.sort()

    # لائنوں کا عمومی فاصلہ (pitch)
    if len(accepted) >= 3:
        diffs = sorted(b - a for a, b in zip(accepted, accepted[1:]))
        pitch = diffs[len(diffs) // 2]
    else:
        pitch = 1.8 * fsize

    # دوسرا مرحلہ: کمزور امیدوار (مثلاً ایک دو لفظوں والی آخری سطر، صفحہ نمبر)
    for y, c in ranked:
        if c >= 4:
            continue
        if all(abs(y - a) >= 0.7 * pitch for a in accepted):
            accepted.append(y)
    return sorted(accepted, reverse=True)   # اوپر سے نیچے


def nearest(baselines, y):
    return min(range(len(baselines)), key=lambda k: abs(baselines[k] - y))


def majority(items, key):
    groups = {}
    for it in items:
        k = key(it)
        if k is not None:
            groups.setdefault(k, []).append(it)
    for members in groups.values():
        votes = Counter(it.line for it in members)
        winner = votes.most_common(1)[0][0]
        for it in members:
            it.line = winner


def num(v):
    return round(float(v), 4)


def ins(op, *operands):
    return ContentStreamInstruction(list(operands), Operator(op))


# --------------------------------------------------------------------------- #
#  ایک صفحے کی منصوبہ بندی
# --------------------------------------------------------------------------- #
class PagePlan:
    def __init__(self, pdf_page, bottom_y, top_y):
        self.page = pdf_page
        self.instructions = list(pikepdf.parse_content_stream(pdf_page))
        self.resources = page_resources(pdf_page.obj)
        self.items = analyse(self.instructions, self.resources)
        texts = [it for it in self.items if it.kind == "text"]
        self.baselines = find_baselines(texts)
        # متن کی لائنیں: نیچے کے مارجن سے اوپر اور (اگر دیا ہو) اوپر کے مارجن سے نیچے
        self.body = [k for k, b in enumerate(self.baselines)
                     if b > bottom_y and (top_y is None or b < top_y)]
        self.dy = {}

    @property
    def last_baseline(self):
        return self.baselines[self.body[-1]] if self.body else None

    def pitch(self):
        bs = [self.baselines[k] for k in self.body]
        d = sorted(a - b for a, b in zip(bs, bs[1:]))
        return d[len(d) // 2] if d else None

    def plan(self, target, mode):
        """ہر لائن کتنی نیچے جائے؟ (منفی = نیچے)"""
        body = self.body
        n = len(body)
        shift = self.last_baseline - target
        if n < 2 or shift <= 0.05:
            return 0.0
        top, last = self.baselines[body[0]], self.baselines[body[-1]]
        for rank, k in enumerate(body):
            if mode == "proportional":
                frac = (top - self.baselines[k]) / (top - last)
            else:  # equal: ہر دو لائنوں کے درمیان برابر اضافہ
                frac = rank / (n - 1)
            self.dy[k] = -shift * frac
        return shift

    def assign(self, pitch):
        texts = [it for it in self.items if it.kind == "text"]
        for it in texts:
            it.line = nearest(self.baselines, it.y)
        # ایک حرف کے تمام اجزا (ActualText) اور ایک مسلسل سلسلے کے حروف ایک ہی لائن میں
        majority(texts, lambda it: it.span)
        majority(texts, lambda it: it.run)
        for it in self.items:
            if it.kind == "graphic":
                x0, y0, x1, y1 = it.bbox
                if not it.clip and (y1 - y0) <= 1.5 * pitch:
                    # خط کشیدہ (underline) baseline سے ذرا نیچے ہوتا ہے
                    it.line = nearest(self.baselines, (y0 + y1) / 2 + 0.15 * pitch)

    def rewrite(self, pdf):
        """ہر لائن کو اپنی جگہ پر ہی (stream کے اندر) کھسکاؤ؛ کوئی نقل نہیں بنتی۔"""
        def dy_of(it):
            return self.dy.get(it.line, 0.0) if it.line is not None else 0.0

        def shifted(m_ctm, dy):
            inv = minv(m_ctm)
            if inv is None:
                return None
            return mmul(mmul(m_ctm, (1, 0, 0, 1, 0, dy)), inv)

        before, after, replace = {}, {}, {}
        seen_runs = set()
        for it in self.items:
            dy = dy_of(it)
            if it.kind == "text":
                if it.run in seen_runs:
                    continue          # سلسلے کا اگلا حرف: پچھلے کے ساتھ خود کھسک جاتا ہے
                seen_runs.add(it.run)
                m = shifted(it.ctm, dy)
                if m is None:
                    continue
                new_tm = mmul(it.tm, m)
                tm_ins = ins("Tm", *[num(v) for v in new_tm])
                orig = self.instructions[it.index]
                op = str(orig.operator)
                if op == "'":
                    replace[it.index] = [ins("T*"), tm_ins, ins("Tj", orig.operands[0])]
                elif op == '"':
                    aw, ac, s = orig.operands
                    replace[it.index] = [ins("Tw", aw), ins("Tc", ac), ins("T*"),
                                         tm_ins, ins("Tj", s)]
                else:
                    before.setdefault(it.index, []).append(tm_ins)
            elif abs(dy) > 1e-6:
                m = shifted(it.ctm, dy)
                if m is None:
                    continue
                before.setdefault(it.start, []).append(ins("q"))
                before[it.start].append(ins("cm", *[num(v) for v in m]))
                after.setdefault(it.index, []).append(ins("Q"))

        out = []
        for i, instr in enumerate(self.instructions):
            out.extend(before.get(i, []))
            out.extend(replace.get(i, [instr]))
            out.extend(after.get(i, []))
        self.page.obj.Contents = pikepdf.Stream(pdf, pikepdf.unparse_content_stream(out))


# --------------------------------------------------------------------------- #
#  مرکزی عمل
# --------------------------------------------------------------------------- #
def main(argv=None):
    try:  # Windows کنسول میں اردو صحیح دکھانے کے لیے
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass
    ap = argparse.ArgumentParser(description="Word سے بنی PDF کے صفحات کو نیچے سے برابر کریں۔")
    ap.add_argument("input")
    ap.add_argument("output")
    ap.add_argument("--bottom-margin", type=parse_length, required=True,
                    help="Word میں سیٹ کیا ہوا نیچے کا مارجن، مثلاً 0.5in یا 1.27cm")
    ap.add_argument("--top-margin", type=parse_length, default=None,
                    help="(اختیاری) اوپر کا مارجن، تاکہ ہیڈر کو متن نہ سمجھا جائے")
    ap.add_argument("--pages", default=None, help="کن صفحات پر کام ہو، مثلاً 5-120")
    ap.add_argument("--skip", default="", help="کون سے صفحات نہ چھیڑے جائیں، مثلاً 23,47,90-91")
    ap.add_argument("--target", type=parse_length, default=None,
                    help="(اختیاری) آخری لائن کی baseline صفحے کے نچلے کنارے سے کتنی اوپر ہو۔ "
                         "نہ دیں تو بھرے ہوئے صفحات سے خود معلوم کیا جائے گا")
    ap.add_argument("--max-lines", type=float, default=2.0,
                    help="اگر کمی اتنی لائنوں سے زیادہ ہو تو صفحہ نہ چھیڑو (پہلے سے 2)")
    ap.add_argument("--mode", choices=["equal", "proportional"], default="equal",
                    help="equal: ہر لائن کے درمیان برابر اضافہ؛ proportional: موجودہ فاصلوں کے تناسب سے")
    args = ap.parse_args(argv)

    pdf = pikepdf.open(args.input)
    npages = len(pdf.pages)
    selected = parse_ranges(args.pages, npages) if args.pages else set(range(npages))
    selected -= parse_ranges(args.skip, npages)

    plans = {}
    for p in sorted(selected):
        page = pdf.pages[p]
        mb = [float(v) for v in page.mediabox]
        bottom_y = mb[1] + args.bottom_margin
        top_y = mb[3] - args.top_margin if args.top_margin is not None else None
        plans[p] = PagePlan(page, bottom_y, top_y)

    # لائنوں کا عمومی فاصلہ پوری کتاب سے
    pitches = [pl.pitch() for pl in plans.values() if pl.pitch()]
    pitch = sorted(pitches)[len(pitches) // 2] if pitches else 20.0

    # ہدف: Word کبھی مارجن پار نہیں کرتا، اس لیے جس صفحے کی آخری لائن سب سے نیچے
    # ہے وہی "بھرا ہوا" صفحہ ہے، اور اس کی baseline ہی مارجن کے مطابق صحیح جگہ ہے۔
    if args.target is not None:
        target_of = lambda pl: float(pl.page.mediabox[1]) + args.target
        target_desc = f"{args.target:.2f}pt (دیا گیا)"
    else:
        lasts = [(pl.last_baseline - float(pl.page.mediabox[1]), p)
                 for p, pl in plans.items() if pl.last_baseline is not None]
        if not lasts:
            sys.exit("کسی صفحے پر متن نہیں ملا۔")
        tgt, ref_page = min(lasts)
        target_of = lambda pl: float(pl.page.mediabox[1]) + tgt
        target_desc = (f"{tgt:.2f}pt صفحے کے نچلے کنارے سے "
                       f"(مارجن سے {tgt - args.bottom_margin:.2f}pt اوپر؛ نمونہ: صفحہ {ref_page + 1})")

    print(f"لائنوں کا فاصلہ: {pitch:.2f}pt | ہدف baseline: {target_desc}")
    print(f"{'صفحہ':>5} {'لائنیں':>6} {'کمی(pt)':>8}  حالت")
    changed = 0
    for p, pl in plans.items():
        if pl.last_baseline is None:
            print(f"{p + 1:>5} {0:>6} {'-':>8}  خالی/متن نہیں")
            continue
        shortfall = pl.last_baseline - target_of(pl)
        n = len(pl.body)
        if shortfall > args.max_lines * pitch:
            print(f"{p + 1:>5} {n:>6} {shortfall:>8.2f}  چھوڑ دیا (کمی {shortfall / pitch:.1f} لائن، حد سے زیادہ)")
            continue
        shift = pl.plan(target_of(pl), args.mode)
        if shift <= 0.05:
            print(f"{p + 1:>5} {n:>6} {shortfall:>8.2f}  پہلے سے درست")
            continue
        pl.assign(pitch)
        pl.rewrite(pdf)
        changed += 1
        print(f"{p + 1:>5} {n:>6} {shortfall:>8.2f}  برابر کر دیا (ہر لائن +{shift / (n - 1):.2f}pt)")

    pdf.save(args.output)
    print(f"\n{changed} صفحات تبدیل ہوئے → {args.output}")


if __name__ == "__main__":
    main()
