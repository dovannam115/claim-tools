# -*- coding: utf-8 -*-
"""Column check for a Map Claim input sheet: which columns the engine reader will find, and
what each missing one costs. Copied from map_claim_page.py of the web app (same rules as
map_claim_ri.read_bangke / gen_pias) -- edit the engine, edit this too."""
import openpyxl

import map_claim_ri as core

# (label, key in core.BK_AL, key in core.ALIASES, group, what dies when the column is missing)
FIELD_SPEC = [
    ("STT", "stt", "no", "Nhận dòng", "Nhận dòng dữ liệu; trống cả cột thì nhận dòng theo IMEI/Serial"),
    ("Số yêu cầu", "gsc", "rr", "Nhận dòng", "Cột định danh hồ sơ trong preview"),
    ("IMEI/Serial", "serial", "serial", "Khóa map", "Khóa map 1, ưu tiên cao nhất"),
    ("Mã giao dịch", "mgd", "mgd", "Khóa map", "Khóa map 2"),
    ("Số điện thoại", "phone", "phone", "Khóa map", "Khóa map 3 (so 9 số cuối)"),
    ("Ngày hiệu lực", "eff", "eff", "Chọn đơn",
     "Tiêu chí ưu tiên 1 khi 1 IMEI có đơn ở nhiều chương trình"),
    ("Ngày hết hiệu lực", "exp", "exp", "Chọn đơn", "Đi kèm Ngày hiệu lực"),
    ("Số tiền bảo hiểm", "si", "si", "Chọn đơn", "Cột SI Check + 1 tiêu chí chấm điểm chọn đơn"),
    ("Ngày tổn thất", "incident", "incident", "Lưu ý HSBT",
     "3 cảnh báo: ngoài thời hạn BH · sát ngày bắt đầu HL · sát ngày hết hạn HL"),
    ("Ngày thông báo", "notif", "notif", "Lưu ý HSBT",
     "2 cảnh báo: báo ngoài thời hạn BH · báo trễ > 30 ngày"),
    ("Nguyên nhân hư hỏng", "cause", "cause", "Lưu ý HSBT", "Cảnh báo nguyên nhân có 'Pin'"),
    ("Phí trả GSC (+vat)", "__cost__", "cost", "Lưu ý HSBT",
     "3 cảnh báo: chi phí BT > 75% SI · > 90% SI · > 100% SI"),
]


def _cost_hdr(first):
    """The header core.find_cost_col() would pick -- same rule, but returns the name."""
    want = {core._sq(a) for a in core.COST_HDRS}
    for h in first:
        if core._sq(h) in want:
            return h
    for h in first:
        s = core._sq(h)
        if "trảgsc" in s and "+vat" in s:
            return h
    return None


def _pias_hdrs(ordered, first):
    """The four fee columns gen_pias needs, resolved exactly the way gen_pias resolves them."""
    def ex(*names):
        return next((n for n in names if n in first), None)
    def fc(*subs):
        return next((nm for nm, _c in ordered if all(s.lower() in nm.lower() for s in subs)), None)
    def vat_after(prec):
        if not prec:
            return None
        pc = first[prec]
        return next((nm for nm, c in ordered if c > pc and "vat" in nm.lower()), None)
    lab = fc("nhân công", "(-vat)")
    part = fc("linh kiện", "sau mt", "(-vat)") or fc("linh kiện", "(-vat)")
    return [("Phí nhân công (-vat)", lab, "Sheet Pias thiếu dòng phí nhân công"),
            ("VAT của nhân công", ex("VAT 8%") or vat_after(lab), "VAT nhân công tính bằng 0"),
            ("Phí linh kiện sau MT (-vat)", part, "Sheet Pias thiếu dòng phí sửa chữa/thay thế"),
            ("VAT của linh kiện", ex("VAT 10% & 8%") or vat_after(part), "VAT linh kiện tính bằng 0")]


def check_columns(path, sheet, kind, hr):
    """(fields, pias) -- which columns the reader will actually find on this sheet."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb[sheet]
    hrow = next(ws.iter_rows(min_row=hr, max_row=hr), ())
    ordered = [(core.norm(c.value), c.column) for c in hrow
               if getattr(c, "column", None) and c.value is not None]
    wb.close()
    first = {}
    for nm, c in ordered:
        if nm and nm not in first:
            first[nm] = c

    al = core.BK_AL if kind == "bangke" else core.ALIASES
    fields = []
    for label, bk, dk, group, lost in FIELD_SPEC:
        if kind == "bangke" and bk == "__cost__":
            hit = _cost_hdr(first)
        else:
            hit = next((a for a in al.get(bk if kind == "bangke" else dk, []) if a in first), None)
        fields.append(dict(group=group, label=label, found=hit, lost=lost))
    # gen_pias only runs for a bảng kê sheet -- an English 'Data' sheet never gets one.
    pias = [dict(group="Sheet Pias", label=lb, found=hit, lost=lost)
            for lb, hit, lost in _pias_hdrs(ordered, first)] if kind == "bangke" else []
    return fields, pias


