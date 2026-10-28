# One run of one tool, started by server.py as a child process:
#   python -u worker.py <tool> <job dir>
# Reads <job dir>\args.json, writes its files into <job dir>, ends with <job dir>\result.json.
# Lines it prints go to the run panel in the window. Markers the server reads:
#   [[P]] done|total|label     progress bar
#   [[CODE]] url|code          Fabric sign-in code to show the user
#   [[ERROR]] text             the message the user sees when the run fails
import os, sys, json, time, shutil, traceback, threading, datetime, decimal
from collections import Counter

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import common as C

try: sys.stdout.reconfigure(encoding="utf-8", line_buffering=True)
except Exception: pass

DIR = ""


def say(*a):
    print(" ".join(str(x) for x in a), flush=True)


def prog(done, total, label=""):
    print("[[P]] %d|%d|%s" % (done, total, str(label).replace("\n", " ")[:120]), flush=True)


class UserError(Exception):
    """A failure the user can fix (wrong file, missing column...): shown without a traceback."""


def jsonable(v):
    if v is None or isinstance(v, (bool, int, str)): return v
    if isinstance(v, float): return None if v != v else v
    if isinstance(v, decimal.Decimal): return float(v)
    if isinstance(v, (datetime.datetime, datetime.date)): return v.strftime("%d/%m/%Y")
    if isinstance(v, (bytes, bytearray)): return v.hex()
    return str(v)


def table(title, cols, rows, views=None, links=(), main=None, limit=2000):
    """A table for the result panel. rows = list of lists (same order as cols) or dicts.
    views = [(label, predicate on the row dict)]; links = columns holding a URL."""
    recs = [r if isinstance(r, dict) else dict(zip(cols, r)) for r in rows]
    out = {"title": title, "cols": cols, "total": len(recs), "links": list(links),
           "rows": [[jsonable(r.get(c)) for c in cols] for r in recs[:limit]]}
    if views:
        out["views"] = [{"label": lb, "rows": [i for i, r in enumerate(recs[:limit]) if f(r)]}
                        for lb, f in views]
    if main: out["main"] = [cols.index(c) for c in main if c in cols]
    return out


def stat(label, value, tone=""):
    return {"label": label, "value": value, "tone": tone}


def money(x):
    try: return "{:,.0f}".format(float(x)).replace(",", ".")
    except Exception: return str(x)


def stop_requested():
    return os.path.exists(os.path.join(DIR, "STOP"))


def out_name(path, suffix):
    return os.path.join(DIR, os.path.splitext(os.path.basename(path))[0] + suffix)


def fabric_connect(core):
    """Connect with the user's own saved sign-in. The engine's own message on a missing sign-in
    talks about the old web server, so check first and say what to click here."""
    if not C.driver():
        raise UserError("No SQL Server driver on this PC. Click 'Get driver' at the top of the page.")
    cred, _who = C.fabric_cred(core)
    if cred is None:
        raise UserError("Not signed in to Fabric yet. Click 'Sign in' at the top of the page.")
    try:
        cred.get_token(C.SCOPE)
    except Exception:
        raise UserError("The Fabric sign-in has expired. Click 'Sign in' at the top of the page.")
    try:
        return C.connect(core, cred)
    except Exception as e:
        raise UserError("Cannot connect to Fabric. Check the network and that your account can open the "
                        "Equipment workspace. (%s)" % str(e).split("\n")[0][:200])


# ------------------------------------------------------------------ Map Claim
def run_map(a):
    core = C.core()
    path, sheet, kind, hr = a["file"], a["sheet"], a["kind"], int(a["hr"])
    plabel, group, family = core.PRODUCTS[int(a.get("product", 0))]
    say("Reading sheet '%s' ..." % sheet)
    if kind == "data":
        claims, hr2 = core.read_data(path, sheet)
    else:
        rows = core.read_bangke(path, sheet, hr)
        claims = [dict(sheet_row=r["sheet_row"], no=None, rr=r["gsc"], serial=r["serial"],
                       phone=r["phone"], phone9=core.last9(r["phone"]), mgd=r["mgd"],
                       si=r["si"], incident=r["incident"], notif=r["notif"],
                       cause=r["cause"], detail=None, cost=r["cost"],
                       eff_file=r["eff_file"], exp_file=r["exp_file"], prod=r.get("prod")) for r in rows]
        hr2 = hr
    if not claims:
        raise UserError("The chosen sheet has no claim rows to map.")
    say("%d claim rows. Connecting to Fabric ..." % len(claims))
    cn = fabric_connect(core); cur = cn.cursor()
    say("Matching against the premium / RI database ...")
    results = core.match(cur, claims, group, family)
    cn.close()
    say("Writing the mapped columns (original format kept) ...")
    edits = [(sheet, hr2, {m["c"]["sheet_row"]: core._row_values(m) for m in results})]
    out = core.surgical_write(path, edits)
    pias_n = None; warns = []
    if kind == "bangke":
        core.PIAS_TPL = C.PIAS_TPL
        pb = core.gen_pias(out, sheet, hr2)
        if pb:
            core.graft_pias(out, pb[0]); pias_n = pb[1]
            say("Added the 'Pias' sheet (%d rows)." % pias_n)
        elif not os.path.exists(C.PIAS_TPL):
            warns.append("No 'Pias' sheet: the Pias template is missing from the app. Click Update.")
        else:
            warns.append("No 'Pias' sheet: the claim sheet has no STT column or no fee rows.")
    else:
        warns.append("No 'Pias' sheet: this sheet is read as an English 'Data' sheet. "
                     "Use the claim statement layout (Số yêu cầu + IMEI/Serial headers) to get Pias.")

    matched = sum(1 for m in results if m["best"])
    bykey = Counter(m["key"] for m in results if m["best"])
    si_mis = sum(1 for m in results if m["best"] and m["c"]["si"] is not None
                 and core.pnum(m["best"]["SI"]) is not None
                 and abs(core.pnum(m["best"]["SI"]) - m["c"]["si"]) >= 1)
    flagged = sum(1 for m in results if core._row_values(m)["hsbt"])
    recreated = sum(1 for m in results if "đã hủy" in (m.get("note_extra") or ""))
    n = len(results)

    cols = ["Row", "Số yêu cầu", "Serial", "Mã giao dịch", "Policy No", "Program", "Program Coverage",
            "Effective Date", "Expired Date", "SI", "RI Share Pct"] + list(core.REINSURERS) + \
           ["Match Key", "SI Check", "Lưu ý HSBT"]
    prev = []
    for m in results:
        v = core._row_values(m); c = m["c"]
        prev.append({"Row": c["sheet_row"], "Số yêu cầu": c["rr"], "Serial": c["serial"],
                     "Mã giao dịch": c["mgd"], "Policy No": v["Policy No"], "Program": v["Program"],
                     "Program Coverage": v["Program Coverage"],
                     "Effective Date": core.fmt_d(v["Effective Date"]),
                     "Expired Date": core.fmt_d(v["Expired Date"]),
                     "SI": v["SI"], "RI Share Pct": core.fmt_pct(v["RI Share Pct"]),
                     **{k: core.fmt_pct(v.get(k)) for k in core.REINSURERS},
                     "Match Key": v["key"], "SI Check": v["sicheck"], "Lưu ý HSBT": v["hsbt"]})
    views = [("All (%d)" % n, lambda r: True),
             ("With note (%d)" % flagged, lambda r: bool(r["Lưu ý HSBT"])),
             ("Not mapped (%d)" % (n - matched), lambda r: r["Match Key"] == "NO MATCH"),
             ("SI diff (%d)" % si_mis, lambda r: r["SI Check"] == "DIFF")]
    notes = ["Sheet '%s' · product %s · match key %s" % (sheet, plabel, dict(bykey))]
    if si_mis: notes.append("SI mismatch %d" % si_mis)
    if recreated: notes.append("Cancelled / recreated %d" % recreated)
    if pias_n is not None: notes.append("Pias sheet %d rows" % pias_n)
    return {"stats": [stat("Rows", n), stat("Mapped", matched, "good" if matched == n else "warn"),
                      stat("Rows with note", flagged, "warn" if flagged else "")],
            "files": [out], "notes": notes, "warns": warns,
            "tables": [table("Columns the tool appends", cols, prev, views)]}


# ------------------------------------------------------------------ Check Salvage
def run_salvage(a):
    import salvage_engine as core
    if not os.path.exists(C.RATE_SOURCE):
        raise UserError("Rate card file is missing: %s" % C.RATE_SOURCE)
    say("Loading rate cards ...")
    cards = core.load_cards(C.RATE_SOURCE)
    path = a["file"]
    if a.get("mode") == "eco":
        say("Reading sheet DT, grouping and pricing ...")
        results = core.reconcile_eco_file(path, cards)
        if not results: raise UserError("Không đọc được điện thoại nào (kiểm tra sheet DT).")
        out = out_name(path, "_ECO.xlsx")
        core.build_eco_workbook(results, out)
        s = core.summary_eco(results)
        cols = ["STT", "Model", "IMEI", "Nhóm", "%Loại", "Tháng", "%GRV Eco", "Giá Eco",
                "%GRV Celebi", "Giá Celebi", "Nhà HĐ", "Nhà cao hơn"]
        rows = [[r.get("stt"), r.get("model"), r.get("imei"), r.get("nhom"), r.get("loai"), r.get("thang"),
                 r.get("grEco"), r.get("ecoGross"), r.get("grCel"), r.get("celGross"), r.get("house"),
                 r.get("higher")] for r in results]
        return {"stats": [stat("Phones", s["n"]), stat("Higher = Ecoswap", s["nHigherEco"]),
                          stat("Higher = Celebi", s["nHigherCel"])],
                "files": [out],
                "notes": ["Theo hợp đồng: Ecoswap %s · Celebi %s · Tổng giá Eco %sđ · Tổng giá Celebi %sđ"
                          % (s["nEco"], s["nCel"], money(s["sEco"]), money(s["sCel"]))],
                "tables": [table("Detail", cols, rows, [
                    ("All (%d)" % len(rows), lambda r: True),
                    ("Higher = Ecoswap (%d)" % s["nHigherEco"], lambda r: r["Nhà cao hơn"] == "Ecoswap"),
                    ("Higher = Celebi (%d)" % s["nHigherCel"], lambda r: r["Nhà cao hơn"] == "Celebi")])]}

    say("Reading BBDS and classifying devices ...")
    results = core.reconcile_file(path, cards)
    if not results: raise UserError("Không đọc được dòng thiết bị nào (kiểm tra sheet BBDS).")
    say("%d devices. Writing the file ..." % len(results))
    out = out_name(path, "_AUTO.xlsx")
    core.build_output_workbook(results, out, cards)
    summ = core.summary(results); tot = summ[-1]
    kpi = [[s["nhom"], s["n"], s["ok"], s["bad"], int(s["auto"]), int(s["bb"]), int(s["diff"])] for s in summ]
    cols = ["STT", "Model", "IMEI", "Nhóm", "Loại TB", "Tháng", "AUTO", "BBDS", "Chênh", "Kết quả", "Ghi chú"]
    rows = [[r.get("stt"), r.get("model"), r.get("imei"), r.get("nhomTxt"), r.get("dt"), r.get("months"),
             r.get("auto"), r.get("bbVal"), r.get("diff"), r.get("symx"),
             core.make_note(str(r.get("code", "")), cards.notes)] for r in results]
    return {"stats": [stat("Devices", tot["n"]), stat("Match", tot["ok"], "good"),
                      stat("Mismatch", tot["bad"], "bad" if tot["bad"] else "")],
            "files": [out],
            "notes": ["AUTO %sđ · BBDS %sđ · chênh %sđ" % (money(tot["auto"]), money(tot["bb"]), money(tot["diff"]))],
            "tables": [table("By group", ["Nhóm", "SL", "Match", "Mismatch", "AUTO (đ)", "BBDS (đ)", "Chênh (đ)"], kpi),
                       table("Detail", cols, rows, [
                           ("All (%d)" % len(rows), lambda r: True),
                           ("Mismatch (%d)" % tot["bad"], lambda r: core.XMK in str(r["Kết quả"] or "")),
                           ("Match (%d)" % tot["ok"], lambda r: core.CHK in str(r["Kết quả"] or ""))])]}


# ------------------------------------------------------------------ Check Invoice
def run_invoice(a):
    import invoice_engine as engine
    say("Reading the device list ...")
    rows = engine.read_input(a["file"])
    if not rows:
        raise UserError("Không tìm thấy dòng nào có IMEI + số HĐ. Kiểm tra lại tiêu đề cột "
                        "(IMEI/Serial · SĐT · Số HĐ · Giá KH mua · Model).")
    speed = a.get("speed") or "Light"
    threads = engine.MODES.get(speed, 3)
    say("%d devices · %s (%d threads). Loading the captcha reader ..." % (len(rows), speed, threads))
    stop = threading.Event()

    def watch():
        while not stop.is_set():
            if stop_requested(): stop.set(); say("Stopping, waiting for the running lookups ..."); return
            time.sleep(0.5)
    threading.Thread(target=watch, daemon=True).start()
    t0 = time.time()
    prog(0, len(rows), "starting")
    res = engine.run_lookup(rows, threads, stop, lambda done, total, part: prog(done, total, "%d/%d devices" % (done, total)))
    halted = stop.is_set()
    stop.set()
    out = os.path.join(DIR, "Reconciliation_Result.xlsx")
    engine.write_excel(out, res)
    n_ok = sum(1 for r in res if r["ket_qua"] == "KHỚP")
    n_le = sum(1 for r in res if r["ket_qua"] == "LỆCH")
    n_no = sum(1 for r in res if r["ket_qua"] == "KHÔNG TRA ĐƯỢC")

    def num(v):
        try: return int(str(v).replace(",", "").strip())
        except Exception: return None
    cols = ["IMEI/Serial", "SĐT", "Số HĐ", "Model", "Giá KH khai", "Giá HĐ (gồm VAT)", "Chênh lệch", "Kết quả", "Cách match"]
    rows_ = [[r["imei"], r["phone"], r["sohd"], r["model"], num(r["gia_khai"]), num(r["gia_gomVAT"]),
              num(r["chenh_lech"]), r["ket_qua"], r["cach_match"]] for r in res]
    notes = ["%d devices · %.1f min" % (len(res), (time.time() - t0) / 60)]
    warns = ["Stopped part way: the file holds the devices looked up before the stop."] if halted else []
    return {"stats": [stat("KHỚP", n_ok, "good"), stat("LỆCH", n_le, "bad" if n_le else ""),
                      stat("KHÔNG TRA ĐƯỢC", n_no, "warn" if n_no else "")],
            "files": [out], "notes": notes, "warns": warns,
            "tables": [table("Detail", cols, rows_, [
                ("All (%d)" % len(rows_), lambda r: True),
                ("LỆCH (%d)" % n_le, lambda r: r["Kết quả"] == "LỆCH"),
                ("KHÔNG TRA ĐƯỢC (%d)" % n_no, lambda r: r["Kết quả"] == "KHÔNG TRA ĐƯỢC")])]}


# ------------------------------------------------------------------ Generate GCNBH
def run_gcngen(a):
    import gcnbh_gen_engine as core
    run_dir = os.path.join(DIR, "GCNBH")
    os.makedirs(run_dir, exist_ok=True)
    say("Filling the Word templates ...")
    try:
        res = core.generate(a["file"], run_dir, want_pdf=bool(a.get("pdf")),
                            progress=lambda d, t, lb: prog(d, t, lb), templates=C.GCN_TPL)
    except core.ToolError as e:
        raise UserError(str(e))
    say("Zipping ...")
    zp = shutil.make_archive(os.path.join(DIR, "GCNBH %s" % datetime.datetime.now().strftime("%Y-%m-%d %H%M")),
                             "zip", run_dir)
    goc = sum(x["rows"] for x in res["sheets"] if x["kind"] == "goc")
    sdbs = sum(x["rows"] for x in res["sheets"] if x["kind"] == "sdbs")
    det = ["%s → %s · %s · %d giấy" % (x["sheet"], x["prod"], "GCN gốc" if x["kind"] == "goc" else "SĐBS", x["rows"])
           for x in res["sheets"]]
    g = res["goc_src"]
    if sum(g.values()):
        det.append("Ngày cấp đơn gốc trên SĐBS: %d từ cột Excel · %d dò từ sheet GCN gốc · %d không tìm ra"
                   % (g["cột Excel"], g["dò sheet GCN gốc"], g[""]))
    det += [os.path.basename(p) for p in res["merged"]]
    return {"stats": [stat("Certificates", res["made"]), stat("GCN gốc", goc), stat("SĐBS", sdbs)],
            "files": [zp], "folder": run_dir, "warns": res["warns"],
            "details": "\n".join(det + [""] + res["log"])}


# ------------------------------------------------------------------ Search GCNBH
def run_gcnsearch(a):
    import gcnbh_search as G
    root = G.tim_kho()
    if not root:
        raise UserError("OneDrive folder '%s' not found on this PC. Ask P.TBHH to share it, then in "
                        "OneDrive on the web choose 'Add shortcut to My files'." % G.KHO_TEN)
    say("Scanning the GCN folder on OneDrive ...")
    kho = G.quet_kho(root)
    ket, src_name = {}, ""
    if a.get("file"):
        data = open(a["file"], "rb").read()
        src_name = os.path.basename(a["file"])
        sheets = a.get("sheets")
        ket = G.doc_file(data, set(sheets) if sheets else None)
    ket.update({k: v for k, v in G.doc_text(a.get("text") or "").items() if k not in ket})
    if not ket:
        raise UserError("Không thấy số hợp đồng nào dạng 25/45/09/KHAC/P000123.")
    goi_y = G.sp_tu_chuoi(os.path.splitext(src_name)[0], G.SP_TENFILE) if src_name else ""
    rows = G.chay(ket, kho["idx"], a.get("product") or G.TU_NHAN, goi_y, kho["xmp"])
    n_link = max([len(x["hits"]) for x in rows] + [1])
    name = (os.path.splitext(src_name)[0].replace("_MAPPED", "") if src_name else "Search") + "_GCNBH.xlsx"
    out = os.path.join(DIR, name)
    open(out, "wb").write(G.xuat_excel(rows, n_link))
    json.dump({"root": root, "rows": rows, "name": name}, open(os.path.join(DIR, "rows.json"), "w", encoding="utf-8"),
              ensure_ascii=False)
    dem = Counter(x["tt"] for x in rows)
    cols = ["Số hợp đồng", "SP hồ sơ", "Tình trạng"]
    for j in range(n_link): cols += ["GCN %d" % (j + 1), "SP %d" % (j + 1)]
    cols += ["XMP", "Số XMP", "Ghi chú"]
    recs = []
    for x in rows:
        d = {"Số hợp đồng": x["pol"], "SP hồ sơ": x["sp"], "Tình trạng": x["tt"]}
        for j in range(n_link):
            h = x["hits"][j] if j < len(x["hits"]) else None
            d["GCN %d" % (j + 1)] = G.weburl(h["path"]) if h else None
            d["SP %d" % (j + 1)] = h["sp"] if h else ""
        d["XMP"] = G.weburl(x["xmp"][0]) if x["xmp"] else None
        d["Số XMP"] = len(x["xmp"]) if x["tt"] == "Tìm thấy" else None
        d["Ghi chú"] = x["note"]
        recs.append(d)
    views = [("All (%d)" % len(recs), lambda r: True)] + \
            [("%s (%d)" % (k, dem[k]), (lambda k: lambda r: r["Tình trạng"] == k)(k)) for k in G.THU_TU if dem.get(k)]
    links = ["GCN %d" % (j + 1) for j in range(n_link)] + ["XMP"]
    det = ("Tên file tool không đọc được (%d):\n" % len(kho["khong_doc"]) + "\n".join(kho["khong_doc"])) \
        if kho["khong_doc"] else ""
    return {"stats": [stat(k, dem[k], "good" if k == "Tìm thấy" else "warn") for k in G.THU_TU if dem.get(k)]
                     + [stat("Total", len(rows))],
            "files": [out], "found": dem.get("Tìm thấy", 0),
            "tables": [table("Result", cols, recs, views, links=links)], "details": det}


def run_gcnprint(a):
    import gcnbh_search as G
    src = json.load(open(os.path.join(a["search_dir"], "rows.json"), encoding="utf-8"))
    say("Downloading the certificates from OneDrive ...")
    phan, so_giay, so_trang, loi = G.gop_pdf(src["root"], src["rows"], bool(a.get("xmp", True)),
                                             lambda d, t, lb: prog(d, t, lb), int(a.get("pages") or G.TRANG_MOI_FILE))
    base = src["name"].replace("_GCNBH.xlsx", "_GCNBH_print")
    files = []
    for i, (data, giay, trang) in enumerate(phan, 1):
        p = os.path.join(DIR, base + ("_%d.pdf" % i if len(phan) > 1 else ".pdf"))
        open(p, "wb").write(data); files.append(p)
        say("%s: %d certificates, %d pages" % (os.path.basename(p), giay, trang))
    return {"stats": [stat("Certificates", so_giay), stat("Pages", so_trang), stat("PDF files", len(files))],
            "files": files, "warns": loi}


# ------------------------------------------------------------------ Search Database
def run_searchdb(a):
    core = C.core()
    import search_db as S
    vals = S.tach_gia_tri(a.get("text") or "")[:S.MAX_VALS]
    if not vals: raise UserError("No value to look up.")
    mode, top = a.get("mode") or "auto", int(a.get("top") or S.MAX_ROWS)
    say("Connecting to Fabric ...")
    cn = fabric_connect(core); cur = cn.cursor()
    say("Loading the RI structure ...")
    core.load_ri_struct(cur)
    say("Looking up %d value(s) ..." % len(vals))
    recs, capped = S.tra_cuu(cur, mode, vals, top)
    cn.close()
    for _v, _k, r in recs: S.gan_co_cau_tai(r)
    for v, k, r in recs: r["Giá trị nhập"], r["Khoá tra"] = v, k
    cols = S.sap_cot([r for _v, _k, r in recs])
    order = {v: i for i, v in enumerate(vals)}
    recs.sort(key=lambda t: (order.get(t[0], 10 ** 9), core.stext(t[2].get("Policy No")) or "",
                             core.pdate(t[2].get("Effective Date")) or datetime.date(1900, 1, 1)))
    miss = [v for v in vals if v not in {v for v, _k, _r in recs}]
    files, tables = [], []
    if recs:
        out = os.path.join(DIR, "Tra_cuu_database.xlsx")
        open(out, "wb").write(S.xuat_excel(recs, cols)); files.append(out)
        tables.append(table("Premium rows", cols, S.bang_hien_thi(recs, cols), main=[c for c in cols if c in S.MAIN_COLS]))
    warns = []
    if capped:
        warns.append("Kết quả đã bị cắt ở %d dòng, database còn dòng khớp nữa. Nâng giới hạn số dòng "
                     "hoặc tra bằng khoá hẹp hơn (IMEI / mã giao dịch)." % len(recs))
    if not recs:
        warns.append("Không có dòng phí nào khớp. Kiểm tra lại khoá tra hoặc thử Auto.")
    npol = len({core.stext(r.get("Policy No")) or "" for _v, _k, r in recs})
    return {"stats": [stat("Values", len(vals)), stat("Found", len(vals) - len(miss), "good" if not miss else "warn"),
                      stat("Premium rows", len(recs)), stat("Policies", npol)],
            "files": files, "warns": warns, "tables": tables,
            "details": ("Không tìm thấy (%d):\n" % len(miss) + "\n".join(miss)) if miss else ""}


# ------------------------------------------------------------------ Reconcile TGDD
def run_tgdd(a):
    import reconcile_tgdd_engine as E
    rules = E.Rules(a.get("rules") or C.RULES)
    if a.get("mode") == "monthly":
        wk = a.get("weekly") or []
        say("Reconciling %d weekly file(s) with the ĐMX file ..." % len(wk))
        hdr, rows, counts, sums = E.reconcile_monthly([(os.path.basename(p), p) for p in wk], a["dmx"], rules)
        blob = E.write_monthly_xlsx(hdr, rows, counts, sums, rules)
        out = out_name(a["dmx"], " - doi soat.xlsx")
        open(out, "wb").write(blob)
        t = [[c[0], c[2], c[4], counts.get(c[0], 0), money(sums.get(c[0], 0.0))] for c in rules.cases]
        return {"stats": [stat("Match (case 1)", counts[1], "good"),
                          stat("Fee diff (2 + 9)", counts[2] + counts[9], "warn" if counts[2] + counts[9] else ""),
                          stat("One side only (4 + 5)", counts[4] + counts[5], "warn" if counts[4] + counts[5] else "")],
                "files": [out], "notes": ["%d weekly file(s) · %d result rows" % (len(wk), len(rows))],
                "tables": [table("By case", ["STT", "Trường hợp", "Kết quả", "Số dòng", "Tổng phí giảm + vat"], t)]}

    product = a.get("product") or "BHMR"
    say("Computing the discounted fee (%s) ..." % product)
    rows, info = E.calc_weekly(a["file"], product, rules)
    blob = E.write_weekly_xlsx(rows)
    out = os.path.join(DIR, E.weekly_out_name(os.path.basename(a["file"])))
    open(out, "wb").write(blob)
    warns = (["%d dòng chưa khớp rule STT, đã tô đỏ ở cột STT trong file, cần xem lại bằng tay."
              % info["unmatched_stt"]] if info["unmatched_stt"] else [])
    return {"stats": [stat("Rows", info["rows"]), stat("Phí gốc − VAT", money(info["sum_fee"])),
                      stat("Phí giảm + VAT", money(info["sum_giam_plus"])), stat("VAT", money(info["sum_vat"]))],
            "files": [out], "warns": warns,
            "notes": ["Product %s · sheet %s · header row %s" % (product, info["sheet"], info["header_row"])],
            "tables": [table("Rows", [h for h, _ in E.OUT_COLS], rows, limit=500)]}


# ------------------------------------------------------------------ Fabric sign-in, ODBC driver
def run_signin(a):
    core = C.core()
    from azure.identity import DeviceCodeCredential, TokenCachePersistenceOptions

    def show(url, code, expires):
        print("[[CODE]] %s|%s" % (url, code), flush=True)
        say("Open %s and enter the code %s" % (url, code))
        say("Sign in with your PVI account that can open the Fabric Equipment workspace.")
    cred = DeviceCodeCredential(client_id=core.CLIENT, prompt_callback=show, timeout=600,
                                cache_persistence_options=TokenCachePersistenceOptions(name="claude_fabric_equip"))
    rec = cred.authenticate(scopes=["https://database.windows.net/.default"])
    os.makedirs(C.DUMP, exist_ok=True)
    open(C.AUTH, "w", encoding="utf-8").write(rec.serialize())
    say("Signed in as %s." % rec.username)
    return {"stats": [stat("Signed in", rec.username, "good")]}


def run_odbc(a):
    """Get the SQL Server driver: the app's own copy (app\vendor) from the update source, else the
    system ODBC Driver 18 through winget (that one needs admin rights)."""
    import io, zipfile, subprocess
    c = C.update_cfg()
    if c["repo"]:
        say("Downloading the database driver ...")
        url = "https://codeload.github.com/%s/zip/refs/heads/%s" % (c["repo"], c["branch"] or "main")
        try:
            data = C.http_get(url, c["token"], timeout=300)
            n = 0
            with zipfile.ZipFile(io.BytesIO(data)) as z:
                for name in z.namelist():
                    rel = name.split("/", 1)[1] if "/" in name else ""
                    if not rel.startswith("app/vendor/") or name.endswith("/"): continue
                    dst = os.path.join(C.ROOT, rel.replace("/", os.sep))
                    os.makedirs(os.path.dirname(dst), exist_ok=True)
                    with open(dst, "wb") as f: f.write(z.read(name))
                    n += 1
            say("%d file(s) copied." % n)
        except Exception as e:
            say("Download failed: %s" % e)
        if os.path.exists(C.VENDOR_OK):
            return {"stats": [stat("Database driver", "ready", "good")]}
    say("Installing '%s' with winget. Windows may ask for permission, click Yes." % C.ODBC)
    p = subprocess.run(["winget", "install", "--id", "Microsoft.msodbcsql.18", "-e", "--silent",
                        "--accept-package-agreements", "--accept-source-agreements"],
                       capture_output=True, text=True, encoding="utf-8", errors="replace", creationflags=0x08000000)
    for ln in (p.stdout or "").splitlines():
        ln = ln.strip()
        if ln and not set(ln) <= set("-|/\\ █▒"): say(ln)
    if not C.driver():
        raise UserError("The driver is still missing. Ask IT to install 'Microsoft ODBC Driver 18 for SQL Server'.")
    return {"stats": [stat("Database driver", "ready", "good")]}


RUN = {"map": run_map, "salvage": run_salvage, "invoice": run_invoice, "gcngen": run_gcngen,
       "gcnsearch": run_gcnsearch, "gcnprint": run_gcnprint, "searchdb": run_searchdb, "tgdd": run_tgdd,
       "signin": run_signin, "odbc": run_odbc}


def main():
    global DIR
    tool, DIR = sys.argv[1], sys.argv[2]
    a = json.load(open(os.path.join(DIR, "args.json"), encoding="utf-8"))
    res = {"ok": False}
    try:
        res = RUN[tool](a) or {}
        res["ok"] = True
        rc = 0
    except UserError as e:
        print("[[ERROR]] " + str(e), flush=True); res["error"] = str(e); rc = 2
    except Exception as e:
        # MapError / ToolError from the engines are written for the user too
        user = type(e).__name__ in ("MapError", "ToolError", "ValueError")
        tb = traceback.format_exc()
        if not user: say(tb)
        msg = str(e) if user else "Unexpected error: %s" % e
        print("[[ERROR]] " + msg.replace("\n", " "), flush=True)
        res["error"] = msg; res["trace"] = "" if user else tb; rc = 2
    res["files"] = [{"name": os.path.basename(p), "path": p} for p in res.get("files", []) if os.path.exists(p)]
    json.dump(res, open(os.path.join(DIR, "result.json"), "w", encoding="utf-8"), ensure_ascii=False, default=jsonable)
    sys.exit(rc)


if __name__ == "__main__":
    main()
