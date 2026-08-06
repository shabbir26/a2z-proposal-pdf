"""
A2Z Proposal PDF service
-------------------------
A thin web wrapper around Shabbir's own a2z_proposals_fpdf.py generator.
The platform POSTs proposal data as JSON; this builds the in-memory workbook
his read_ltd / read_sa / read_partnership expect, then calls his UNCHANGED
build_ltd / build_sa / build_partnership and returns the byte-identical PDF.

His generator code is never modified - it is imported and called as-is.
"""
import io, os, tempfile, datetime, traceback, shutil, glob
from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import openpyxl

# --- Font self-heal ------------------------------------------------------
# His generator loads fonts from a "fonts/" subfolder next to it. When the .ttf
# files get uploaded loose at the repo root (GitHub drag-drop drops folders),
# they end up beside this file instead of inside fonts/. Make it robust: ensure
# a fonts/ folder exists and holds the 5 ttf, copying any that sit at the top
# level. Runs before importing his module so the path is ready. His file is
# NOT modified.
_HERE = os.path.dirname(os.path.abspath(__file__))
_FONTS_DIR = os.path.join(_HERE, "fonts")
_NEEDED = ["Cormorant-Bold.ttf", "Cormorant-SemiBold.ttf",
           "Nunito-Regular.ttf", "Nunito-SemiBold.ttf", "Nunito-Bold.ttf"]
try:
    os.makedirs(_FONTS_DIR, exist_ok=True)
    # index every .ttf anywhere under the project root, by filename
    found = {}
    for path in glob.glob(os.path.join(_HERE, "**", "*.ttf"), recursive=True):
        found.setdefault(os.path.basename(path), path)
    for fn in _NEEDED:
        dest = os.path.join(_FONTS_DIR, fn)
        if not os.path.exists(dest) and fn in found and os.path.abspath(found[fn]) != os.path.abspath(dest):
            shutil.copyfile(found[fn], dest)
    missing = [fn for fn in _NEEDED if not os.path.exists(os.path.join(_FONTS_DIR, fn))]
    if missing:
        print("FONT SELF-HEAL: still missing %s (looked under %s)" % (missing, _HERE), flush=True)
    else:
        print("FONT SELF-HEAL: all 5 fonts present in %s" % _FONTS_DIR, flush=True)
except Exception as _e:
    print("FONT SELF-HEAL failed: %s" % _e, flush=True)
# ------------------------------------------------------------------------

import a2z_proposals_fpdf as GEN  # his file, untouched

app = Flask(__name__)
# Allow the platform (and local testing) to call this endpoint from the browser.
CORS(app, resources={r"/*": {"origins": "*"}})


def _num(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return 0.0


def _set(ws, cell, val):
    ws[cell] = val


# ---- Build the 'LTD Proposal' sheet his read_ltd() reads ----
def _fill_ltd_sheet(ws, d):
    _set(ws, "C10", d.get("company") or "Your Company Ltd")
    _set(ws, "C11", d.get("contact") or "Director")
    _set(ws, "C12", d.get("email") or "")
    _set(ws, "C13", d.get("phone") or "")
    _set(ws, "C14", d.get("date") or datetime.date.today().strftime("%d %B %Y"))
    _set(ws, "C15", d.get("prepared_by") or "Shabbir Rahman FCCA")
    _set(ws, "C18", d.get("band") or "")
    _set(ws, "C25", int(_num(d.get("directors"))))
    _set(ws, "C35", int(_num(d.get("people"))))
    _set(ws, "K4", _num(d.get("discount")))
    _set(ws, "B72", d.get("notes") or "")
    _set(ws, "I10", d.get("internal_notes") or "")
    _set(ws, "J25", d.get("source") or "")
    _set(ws, "J26", d.get("referrer") or "")

    svc = d.get("services", {})
    # accounts+CT (his read_ltd sums G22+G28)
    _set(ws, "G22", _num(svc.get("accounts")))
    _set(ws, "G28", _num(svc.get("ct")))
    _set(ws, "G23", _num(svc.get("cs01")))
    _set(ws, "G24", _num(svc.get("address")))
    _set(ws, "C27", svc.get("book_freq") or "")
    _set(ws, "G27", _num(svc.get("book")))
    _set(ws, "C29", svc.get("vat_freq") or "")
    _set(ws, "G29", _num(svc.get("vat")))
    _set(ws, "C30", svc.get("software_name") or "")
    _set(ws, "G30", _num(svc.get("software")))
    _set(ws, "G31", _num(svc.get("dext")))
    _set(ws, "G32", _num(svc.get("contractor")))
    _set(ws, "C34", svc.get("payroll_freq") or "")
    _set(ws, "G34", _num(svc.get("payroll")))
    _set(ws, "G35", _num(svc.get("payroll_extra")))
    _set(ws, "G36", _num(svc.get("cis")))
    # management report tier: one of C39/C40/C41 = Quarterly/Monthly, fee in G39/40/41
    mt = (d.get("mgmt_tier") or "").upper()
    mf = _num(d.get("mgmt_freq_fee"))
    freq = d.get("mgmt_freq") or "Quarterly"
    if mt == "T1":
        _set(ws, "C39", freq); _set(ws, "G39", mf)
    elif mt == "T2":
        _set(ws, "C40", freq); _set(ws, "G40", mf)
    elif mt == "T3":
        _set(ws, "C41", freq); _set(ws, "G41", mf)

    # totals
    _set(ws, "G44", _num(d.get("sub")))
    _set(ws, "G45", _num(d.get("vat")))
    _set(ws, "G46", _num(d.get("gross")))

    # one-offs rows 50-57 (B desc, D detail, E when, F price, G note)
    for i, o in enumerate(d.get("oneoffs", [])[:8]):
        r = 50 + i
        _set(ws, "B%d" % r, o.get("label") or "")
        _set(ws, "D%d" % r, o.get("detail") or "")
        _set(ws, "F%d" % r, _num(o.get("amount")))
    _set(ws, "F67", _num(d.get("osub")))
    _set(ws, "F68", _num(d.get("ovat")))
    _set(ws, "F69", _num(d.get("ogross")))

    # setup/registration rows 60-66 (C=Required flag, F=fee)
    regs = d.get("registrations", {})
    rowmap = {"company": 60, "paye": 61, "vat": 62, "cis_sub": 63,
              "cis_con": 64, "sa": 65, "other": 66}
    for key, row in rowmap.items():
        if regs.get(key):
            _set(ws, "C%d" % row, "Required")
            _set(ws, "F%d" % row, _num((regs.get(key) or {}).get("fee")))


def _fill_rate_card(wb, ratecard):
    """Optional: platform can pass its rate card so page 8 renders. Skipped if absent."""
    if not ratecard:
        return
    ws = wb.create_sheet("Rate Card")
    r = 7
    for item in ratecard[:33]:
        ws.cell(r, 2, item.get("service") or "")
        ws.cell(r, 3, item.get("price") or "")
        ws.cell(r, 4, item.get("note") or "")
        r += 1


def _build_workbook_ltd(d):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "LTD Proposal"
    _fill_ltd_sheet(ws, d)
    _fill_rate_card(wb, d.get("ratecard"))
    return wb


# ---- Build the 'SA Proposal' sheet his read_sa() reads (sole trader / self-assessment) ----
def _fill_sa_sheet(ws, d):
    _set(ws, "C5", d.get("prepared_by") or "Shabbir Rahman FCCA")
    _set(ws, "C6", d.get("company") or d.get("contact") or "Client")
    _set(ws, "C7", d.get("contact") or "there")
    _set(ws, "C8", d.get("email") or "")
    _set(ws, "C9", d.get("phone") or "")
    _set(ws, "C10", d.get("date") or datetime.date.today().strftime("%d %B %Y"))
    _set(ws, "C13", d.get("ctype") or "")
    _set(ws, "C15", d.get("freq") or "Annually")
    _set(ws, "J30", d.get("source") or "")
    _set(ws, "J31", d.get("referrer") or "")
    _set(ws, "B60", d.get("notes") or "")

    # service lines - his read_sa reads 7 fixed rows (26-32); the platform's SA proposal is
    # a package + optional extras, so map by service name into the right fixed row.
    sa = d.get("services", {})
    monthly = (d.get("freq") or "Annually").lower() == "monthly"
    col = "E" if monthly else "D"   # E = monthly, D = annual
    rowmap = {"sa_return": 26, "property": 27, "mtd": 28,
              "book": 29, "vat": 30, "payroll": 31, "software": 32}
    for key, row in rowmap.items():
        v = _num(sa.get(key))
        if v > 0:
            _set(ws, "%s%d" % (col, row), v)

    # totals: his read_sa reads annual D34/D36, monthly E34/E36
    fee = _num(d.get("sub"))
    if monthly:
        _set(ws, "E34", fee); _set(ws, "E36", fee * 1.2)
    else:
        _set(ws, "D34", fee); _set(ws, "D36", fee * 1.2)
    _set(ws, "H34", _num(d.get("discount_annual")))
    _set(ws, "H35", _num(d.get("discount_monthly")))

    # one-offs rows 40-47 (B desc, E price)
    for i, o in enumerate(d.get("oneoffs", [])[:8]):
        r = 40 + i
        _set(ws, "B%d" % r, o.get("label") or "")
        _set(ws, "E%d" % r, _num(o.get("amount")))
    _set(ws, "E55", _num(d.get("osub")))
    _set(ws, "E56", _num(d.get("ovat")))
    _set(ws, "E57", _num(d.get("ogross")))

    # registration block rows 49-54 (C=Required, E=fee): sa/paye/vat/cis_sub/cis_con/other
    regs = d.get("registrations", {})
    sarow = {"sa": 49, "paye": 50, "vat": 51, "cis_sub": 52, "cis_con": 53, "other": 54}
    for key, row in sarow.items():
        if regs.get(key):
            _set(ws, "C%d" % row, "Required")
            _set(ws, "E%d" % row, _num((regs.get(key) or {}).get("fee")))


def _build_workbook_sa(d):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "SA Proposal"
    _fill_sa_sheet(ws, d)
    _fill_rate_card(wb, d.get("ratecard"))
    return wb


@app.route("/health")
def health():
    return jsonify(ok=True, service="a2z-proposal-pdf")


@app.route("/proposal", methods=["POST", "OPTIONS"])
def proposal():
    if request.method == "OPTIONS":
        return ("", 204)
    d = request.get_json(force=True, silent=True) or {}
    kind = (d.get("kind") or "LTD").upper()

    tmpdir = tempfile.mkdtemp()
    wb_path = os.path.join(tmpdir, "wb.xlsx")
    out_path = os.path.join(tmpdir, "proposal.pdf")
    try:
        if kind in ("LTD", "CIC", "CHARITY"):
            wb = _build_workbook_ltd(d)
            wb.save(wb_path)
            wb2 = GEN.safe_load_workbook(wb_path)
            GEN.build_ltd(wb2, out_path, ref=d.get("ref"))
        elif kind == "PARTNERSHIP":
            # partnership sheet uses the same layout under a different name
            wb = _build_workbook_ltd(d)
            wb.active.title = "Partnership Proposal"
            wb.save(wb_path)
            wb2 = GEN.safe_load_workbook(wb_path)
            GEN.build_partnership(wb2, out_path, ref=d.get("ref"))
        elif kind == "SA":
            wb = _build_workbook_sa(d)
            wb.save(wb_path)
            wb2 = GEN.safe_load_workbook(wb_path)
            GEN.build_sa(wb2, out_path, ref=d.get("ref"))
        else:
            return jsonify(error="Unknown kind: %s" % kind), 400
    except Exception as e:
        print("=== PROPOSAL PDF BUILD ERROR ===", flush=True)
        traceback.print_exc()
        print("=== END ERROR (kind=%s) ===" % (d.get("kind")), flush=True)
        return jsonify(error="PDF build failed", detail=str(e)), 500

    fname = (d.get("company") or "Proposal").replace("/", " ").replace("\\", " ")
    fname = " ".join(fname.split()) + " - proposal.pdf"
    return send_file(out_path, mimetype="application/pdf",
                     as_attachment=True, download_name=fname)


# ------------------------------------------------------------------
#  Direct email send via Microsoft 365 Graph (reuses the firm's
#  existing Entra app - Mail.Send application permission, admin
#  consent already granted for the Academy). Set these env vars on
#  the Render service (copy the values from the Academy setup):
#     MS_TENANT_ID, MS_CLIENT_ID, MS_CLIENT_SECRET
#     MAIL_FROM        (or INVITE_FROM) - the sending mailbox, e.g. info@a2zaccounting.co.uk
#     SEND_TOKEN       (optional but recommended) - a shared secret the
#                      platform must send in the X-A2Z-Token header
# ------------------------------------------------------------------
import base64 as _b64
import requests as _rq


def _mail_from():
    return os.environ.get("MAIL_FROM") or os.environ.get("INVITE_FROM") or ""


import re as _re


def _text_to_html(text):
    """Turn his plain-text email into simple, email-safe HTML:
    URLs become clickable links, bullet lines render as bullets, blank lines
    become paragraph breaks. Keeps his wording exactly."""
    def esc(s):
        return (s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))

    url_re = _re.compile(r"(https?://[^\s]+)")
    out = []
    out.append('<div style="font-family:Arial,Helvetica,sans-serif;font-size:14px;'
               'color:#1b2a38;line-height:1.6">')
    for raw in text.split("\n"):
        line = raw.rstrip()
        if line == "":
            out.append("<br>")
            continue
        safe = esc(line)
        safe = url_re.sub(lambda m: '<a href="%s" style="color:#1e6b47">%s</a>'
                          % (m.group(1), m.group(1)), safe)
        if line.startswith("\u2022 "):
            out.append('<div style="margin:2px 0 2px 8px">&bull; %s</div>' % safe[2:])
        else:
            out.append('<div>%s</div>' % safe)
    out.append("</div>")
    return "".join(out)


def _graph_token():
    tid = os.environ.get("MS_TENANT_ID", "")
    cid = os.environ.get("MS_CLIENT_ID", "")
    sec = os.environ.get("MS_CLIENT_SECRET", "")
    if not (tid and cid and sec):
        raise RuntimeError("Microsoft 365 not configured (MS_TENANT_ID / MS_CLIENT_ID / MS_CLIENT_SECRET missing).")
    r = _rq.post(
        "https://login.microsoftonline.com/%s/oauth2/v2.0/token" % tid,
        data={
            "client_id": cid,
            "client_secret": sec,
            "scope": "https://graph.microsoft.com/.default",
            "grant_type": "client_credentials",
        }, timeout=30)
    r.raise_for_status()
    return r.json()["access_token"]


def _pdf_bytes_for(d):
    """Build the proposal PDF (same path as /proposal) and return raw bytes."""
    kind = (d.get("kind") or "LTD").upper()
    tmpdir = tempfile.mkdtemp()
    wb_path = os.path.join(tmpdir, "wb.xlsx")
    out_path = os.path.join(tmpdir, "proposal.pdf")
    if kind == "SA":
        wb = _build_workbook_sa(d)
    else:
        wb = _build_workbook_ltd(d)
        if kind == "PARTNERSHIP":
            wb.active.title = "Partnership Proposal"
    wb.save(wb_path)
    wb2 = GEN.safe_load_workbook(wb_path)
    if kind == "SA":
        GEN.build_sa(wb2, out_path, ref=d.get("ref"))
    elif kind == "PARTNERSHIP":
        GEN.build_partnership(wb2, out_path, ref=d.get("ref"))
    else:
        GEN.build_ltd(wb2, out_path, ref=d.get("ref"))
    with open(out_path, "rb") as f:
        return f.read()


def _enhanced_email(d, kind, scenario):
    """Returns (subject, text_body, html_body). Text is his exact wording; html is a
    branded design (navy + GREEN, matching the proposal PDF, so it is clearly distinct
    from the welcome email's gold). His a2z_proposals_fpdf.py is NOT modified."""
    K = kind.upper()
    gbp = GEN.gbp
    num = GEN.num
    FORMS = GEN.FORMS

    def esc(s):
        return (str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
                .replace("\u00a3", "&pound;").replace("\u2014", "&mdash;")
                .replace("\u2013", "&ndash;").replace("\u2019", "&rsquo;")
                .replace("\u2018", "&lsquo;").replace("\u2026", "&hellip;"))

    name = (str(d.get("contact") or "there").strip()) or "there"
    company = (str(d.get("company") or "your business").strip()) or "your business"
    subject = "Your A2Z proposal for %s" % company
    svc = [str(r[0]).strip() for r in d.get("lines", []) if str(r[0]).strip()]
    regs = d.get("regs") or []
    keys = {r.get("key") for r in regs}
    intro = ("Thank you for the opportunity to look after %s \u2014 it's a pleasure to put this "
             "proposal together for you. I've attached the full proposal, and here's a quick summary." % company)

    L = ["Hi %s," % name, "", intro, "", "What we'd take care of for you:"]
    L += ["\u2022 %s" % x for x in (svc or ["the services we discussed"])]
    L += [""]

    # ---- fee ----
    fee_lines = []
    fee_display = ""
    fee_sub = "a single fixed amount, with no surprise bills along the way"
    if K == "LTD":
        fee_display = "%s + VAT" % gbp(d.get("sub", 0)); fee_period = "per month"
        fee_lines.append("Your fee would be %s + VAT a month \u2014 a single fixed amount, with no surprise bills along the way." % gbp(d.get("sub", 0)))
        if (num(d.get("discount", 0)) or 0) > 0:
            fee_lines.append("As an exceptional act of discretion, a goodwill discount of %s + VAT per month has been applied. Your fee above already reflects this." % gbp(d["discount"]))
        nd = int(num(d.get("directors", 0)) or 0)
        if nd == 1:
            fee_lines.append("Each director's personal tax return (self-assessment) is \u00a3120 + VAT a year, billed separately.")
        elif nd > 1:
            fee_lines.append("Each director's personal tax return (self-assessment) is \u00a3120 + VAT a year, billed separately \u2014 for %d directors that comes to %s + VAT a year." % (nd, gbp(nd * 120)))
    else:
        m = num(d.get("monthly", 0)) or 0
        a = num(d.get("annual", 0)) or 0
        if m > 0:
            fee_display = "%s + VAT" % gbp(d["monthly"]); fee_period = "per month"
            fee_lines.append("Your fee would be %s + VAT a month \u2014 a single fixed amount, with no surprise bills along the way." % gbp(d["monthly"]))
        elif a > 0:
            fee_display = "%s + VAT" % gbp(d["annual"]); fee_period = "per year"
            fee_lines.append("Your fee would be %s + VAT a year \u2014 a single fixed amount, with no surprise bills along the way." % gbp(d["annual"]))
        else:
            fee_period = ""
            fee_lines.append("Your fee is set out in the attached proposal.")
    for fl in fee_lines:
        L += [fl]
    L += [""]

    # ---- itemised setup ----
    setup_rows = []   # (label, value)
    setup_total = 0
    reg_labels = set()
    for r in regs:
        lab = str(r.get("label") or "").strip()
        reg_labels.add(lab.lower())
        fee = num(r.get("fee")) or 0
        if r.get("included") or fee <= 0:
            setup_rows.append((lab, "included"))
        else:
            setup_rows.append((lab, "%s + VAT" % gbp(fee))); setup_total += fee
    for o in d.get("oneoffs", []):
        desc = str(o[0]).strip() if o and len(o) > 0 else ""
        price = num(o[2]) if o and len(o) > 2 else 0
        if not desc:
            continue
        dl = desc.lower()
        if any(dl == rl or dl in rl or rl in dl for rl in reg_labels):
            continue
        if price and price > 0:
            setup_rows.append((desc, "%s + VAT" % gbp(price))); setup_total += price
        else:
            setup_rows.append((desc, "included"))
    osub = num(d.get("osub", 0)) or 0
    if osub > setup_total:
        setup_total = osub
    setup_total_line = ("One-off setup: %s + VAT" % gbp(setup_total)) if setup_total > 0 else "There's no charge for your setup \u2014 it's all included."
    if setup_rows:
        L += ["To get you set up, here's the one-off work at the start:"]
        L += ["\u2022 %s \u2014 %s" % (lab, val) for lab, val in setup_rows]
        L += [setup_total_line, ""]
    elif osub > 0:
        L += ["There's also a one-off setup of %s + VAT to get everything in place at the start." % gbp(osub), ""]

    # ---- onboarding scenario ----
    primary_key = "company" if K == "LTD" else "sa"
    primary_fee = 0
    for r in regs:
        if r.get("key") == primary_key:
            primary_fee = num(r.get("fee")) or 0
            break
    onboard_name, onboard_url = FORMS[K]["onboard"]
    reg_name, reg_url = FORMS[K]["reg"]
    sc = (scenario or "").lower()
    if sc not in ("newco", "switcher", "existing"):
        sc = "newco" if primary_key in keys else "existing"
    thing = "company" if K == "LTD" else "business"
    if sc == "newco":
        if K == "LTD":
            feebit = (" (the company formation fee is %s + VAT, charged once)" % gbp(primary_fee)) if primary_fee > 0 else ""
            onboard_text = "As we're forming %s for you, the first step is to complete our %s%s, which only takes a few minutes:" % (company, reg_name, feebit)
        else:
            onboard_text = "To get you registered, the first step is to complete our %s, which only takes a few minutes:" % reg_name
        cta_label, cta_url = "Open the " + reg_name, reg_url
    elif sc == "switcher":
        onboard_text = "You're moving to us from another accountant, so there's nothing for you to chase \u2014 just complete our %s and we'll write to your current accountant for professional clearance and handle the whole handover for you:" % onboard_name
        cta_label, cta_url = "Open the " + onboard_name, onboard_url
    else:
        if K == "LTD":
            onboard_text = "As your %s is already set up and you've not had an accountant before, getting started is simply completing our %s, which only takes a few minutes:" % (thing, onboard_name)
        else:
            onboard_text = "As you're already trading and you've not had an accountant before, getting started is simply completing our %s, which only takes a few minutes:" % onboard_name
        cta_label, cta_url = "Open the " + onboard_name, onboard_url
    L += [onboard_text, cta_url]

    # ---- extra registration links ----
    extra = []
    seen = set()
    for r in regs:
        if r.get("key") == primary_key:
            continue
        fm = r.get("form")
        if fm and fm[1] not in seen:
            lab = fm[0][:-5] if str(fm[0]).lower().endswith(" form") else fm[0]
            extra.append((lab, fm[1]))
            seen.add(fm[1])
    if extra:
        L += ["", "We'd also take care of a couple of registrations for you \u2014 you can complete those here as well:"]
        for lab, u in extra:
            L += ["%s: %s" % (lab, u)]
    closing = "If you have any questions, or would like to talk anything through, just let me know \u2014 I'd be glad to help."
    L += ["", closing]
    text_body = "\n".join(L)

    # ---------- branded HTML (navy + green) ----------
    NAVY = "#0D2B42"; GREEN = "#1E6B47"; CREAM = "#FBF8F3"; INK = "#243a4d"; MUTE = "#9fb3c8"
    svc_rows = "".join(
        '<tr><td valign="top" style="padding:3px 10px 3px 0;color:%s;font-weight:bold;">&#10003;</td>'
        '<td style="padding:3px 0;color:%s;font-size:15px;">%s</td></tr>' % (GREEN, INK, esc(x))
        for x in (svc or ["the services we discussed"]))
    setup_html = ""
    if setup_rows:
        rows = "".join(
            '<tr><td style="padding:7px 0;border-bottom:1px solid #eee6d8;color:%s;font-size:14px;">%s</td>'
            '<td align="right" style="padding:7px 0;border-bottom:1px solid #eee6d8;color:%s;font-size:14px;white-space:nowrap;">%s</td></tr>'
            % (INK, esc(lab), (GREEN if val == "included" else INK), esc(val))
            for lab, val in setup_rows)
        total_html = ('<tr><td style="padding:9px 0 0;font-weight:bold;color:%s;font-size:14px;">%s</td><td></td></tr>' % (GREEN, esc(setup_total_line))) if setup_total <= 0 else ('<tr><td style="padding:9px 0 0;font-weight:bold;color:%s;">One-off setup</td><td align="right" style="padding:9px 0 0;font-weight:bold;color:%s;white-space:nowrap;">%s + VAT</td></tr>' % (GREEN, GREEN, esc(gbp(setup_total))))
        setup_html = (
            '<div style="font-size:12px;letter-spacing:1.5px;text-transform:uppercase;color:%s;font-weight:bold;margin:26px 0 10px;">To get you set up</div>'
            '<table role="presentation" width="100%%" cellpadding="0" cellspacing="0" style="border:1px solid #eee6d8;border-radius:8px;padding:6px 16px;background:#FCFAF5;">%s%s</table>'
            % (GREEN, rows, total_html))
    fee_lines_html = "".join('<p style="margin:12px 0 0;color:%s;font-size:14px;line-height:1.6;">%s</p>' % (INK, esc(fl)) for fl in fee_lines[1:])
    extra_html = ""
    if extra:
        links = "".join('<div style="margin:6px 0;"><a href="%s" style="color:%s;font-weight:bold;text-decoration:none;font-size:14px;">%s &rarr;</a></div>' % (u, GREEN, esc(lab)) for lab, u in extra)
        extra_html = ('<p style="margin:22px 0 6px;color:%s;font-size:15px;">We\'d also take care of a couple of registrations for you &mdash; you can complete those here as well:</p>%s' % (INK, links))
    fee_box = ""
    if fee_display:
        fee_box = (
            '<table role="presentation" width="100%%" cellpadding="0" cellspacing="0" style="margin:22px 0;"><tr>'
            '<td style="background:%s;border-radius:8px;padding:20px 24px;">'
            '<div style="color:%s;font-size:11px;letter-spacing:1.5px;text-transform:uppercase;">Your fee</div>'
            '<div style="color:#ffffff;font-size:30px;font-family:Georgia,serif;margin:4px 0 2px;">%s</div>'
            '<div style="color:#cfe3d8;font-size:13px;">%s &middot; %s</div>'
            '</td></tr></table>' % (NAVY, MUTE, esc(fee_display), esc(fee_period), esc(fee_sub)))
    html_body = (
        '<div style="margin:0;padding:0;background:%s;">'
        '<table role="presentation" width="100%%" cellpadding="0" cellspacing="0" style="background:%s;"><tr><td align="center" style="padding:24px 12px;">'
        '<table role="presentation" width="600" cellpadding="0" cellspacing="0" style="max-width:600px;width:100%%;">'
        # header
        '<tr><td style="background:%s;padding:26px 32px;border-radius:10px 10px 0 0;">'
        '<div style="font-family:Georgia,serif;color:#EAF0F7;font-size:13px;letter-spacing:3px;">A2Z ACCOUNTING SOLUTIONS</div>'
        '<div style="height:3px;width:46px;background:%s;margin:14px 0;"></div>'
        '<div style="color:%s;font-size:11px;letter-spacing:2px;text-transform:uppercase;">Proposal of services</div>'
        '<div style="color:#ffffff;font-size:23px;font-family:Georgia,serif;margin-top:3px;">%s</div>'
        '</td></tr>'
        # body
        '<tr><td style="background:#ffffff;padding:30px 32px;font-family:Arial,Helvetica,sans-serif;">'
        '<p style="margin:0 0 16px;color:%s;font-size:15px;">Hi %s,</p>'
        '<p style="margin:0 0 20px;color:%s;font-size:15px;line-height:1.65;">%s</p>'
        '<div style="font-size:12px;letter-spacing:1.5px;text-transform:uppercase;color:%s;font-weight:bold;margin:0 0 10px;">What we\'d take care of for you</div>'
        '<table role="presentation" cellpadding="0" cellspacing="0">%s</table>'
        '%s'   # fee box
        '%s'   # fee lines (directors etc)
        '%s'   # setup
        '<p style="margin:24px 0 14px;color:%s;font-size:15px;line-height:1.65;">%s</p>'
        '<a href="%s" style="display:inline-block;background:%s;color:#ffffff;text-decoration:none;padding:13px 28px;border-radius:6px;font-weight:bold;font-size:14px;">%s</a>'
        '%s'   # extra links
        '<p style="margin:24px 0 0;color:%s;font-size:15px;line-height:1.65;">%s</p>'
        '</td></tr>'
        # footer
        '<tr><td style="background:%s;padding:18px 32px;border-radius:0 0 10px 10px;color:%s;font-size:12px;line-height:1.7;">'
        'A2Z Accounting Solutions &middot; 01224 042961 &middot; info@a2zaccounting.co.uk<br>'
        '1st Floor, 499 Union Street, Aberdeen, AB11 6DB &middot; Regulated by ACCA'
        '</td></tr>'
        '</table></td></tr></table></div>'
    ) % (CREAM, CREAM, NAVY, GREEN, MUTE, esc(company),
         INK, esc(name), INK, esc(intro), GREEN, svc_rows,
         fee_box, fee_lines_html, setup_html,
         INK, esc(onboard_text), cta_url, GREEN, esc(cta_label),
         extra_html, INK, esc(closing), NAVY, MUTE)

    return subject, text_body, html_body


def _email_for(d):
    """Produce the proposal email. LTD/SA use _enhanced_email (his wording + itemised
    setup + 3 onboarding scenarios); PARTNERSHIP keeps his original build_email."""
    kind = (d.get("kind") or "LTD").upper()
    scenario = d.get("scenario") or ""
    tmpdir = tempfile.mkdtemp()
    wb_path = os.path.join(tmpdir, "wb.xlsx")
    if kind == "SA":
        wb = _build_workbook_sa(d)
    else:
        wb = _build_workbook_ltd(d)
        if kind == "PARTNERSHIP":
            wb.active.title = "Partnership Proposal"
    wb.save(wb_path)
    wb2 = GEN.safe_load_workbook(wb_path)
    if kind == "SA":
        return _enhanced_email(GEN.read_sa(wb2), "SA", scenario)
    if kind == "PARTNERSHIP":
        psub, pbody = GEN.build_email(GEN.read_partnership(wb2), "PARTNERSHIP")
        return psub, pbody, _text_to_html(pbody)
    return _enhanced_email(GEN.read_ltd(wb2), "LTD", scenario)


@app.route("/email", methods=["POST", "OPTIONS"])
def email():
    """Return his exact proposal email as {subject, body} for the given contract."""
    if request.method == "OPTIONS":
        return ("", 204)
    d = request.get_json(force=True, silent=True) or {}
    try:
        subject, body, html = _email_for(d)
        return jsonify(ok=True, subject=subject, body=body, html=html)
    except Exception as e:
        return jsonify(error="Could not build the email", detail=str(e)), 500


@app.route("/send", methods=["POST", "OPTIONS"])
def send():
    if request.method == "OPTIONS":
        return ("", 204)
    # optional shared-secret guard
    want = os.environ.get("SEND_TOKEN")
    if want and request.headers.get("X-A2Z-Token") != want:
        return jsonify(error="Not authorised"), 401

    d = request.get_json(force=True, silent=True) or {}
    to = (d.get("to") or "").strip()
    subject = d.get("subject") or ""
    html = d.get("html") or ""
    if not to:
        return jsonify(error="No recipient (to) provided"), 400

    # For a proposal send, build the subject/body from HIS OWN build_email so it
    # matches his desktop tool exactly (warm tone, service list, all reg links).
    if d.get("attach_proposal") and d.get("proposal") and not d.get("html_override"):
        try:
            esubject, ebody, ehtml = _email_for(d["proposal"])
            subject = subject or esubject
            html = ehtml or _text_to_html(ebody)
        except Exception:
            pass  # fall back to whatever html/subject was supplied

    # Sender mailbox: the platform may choose which mailbox to send from (info@, payroll@, etc).
    # Falls back to MAIL_FROM. If ALLOWED_FROM is set (comma-separated), the chosen mailbox
    # must be on that list - a light guard so only your real mailboxes can be used.
    frm = (d.get("from") or "").strip() or _mail_from()
    if not frm:
        return jsonify(error="Sending mailbox not configured (MAIL_FROM)."), 500
    allow = [a.strip().lower() for a in (os.environ.get("ALLOWED_FROM") or "").split(",") if a.strip()]
    if allow and frm.lower() not in allow:
        return jsonify(error="That sending mailbox (%s) is not on the allowed list." % frm), 400

    message = {
        "subject": subject,
        "body": {"contentType": "HTML", "content": html},
        "toRecipients": [{"emailAddress": {"address": to}}],
    }
    cc = d.get("cc") or []
    if cc:
        message["ccRecipients"] = [{"emailAddress": {"address": a}} for a in cc]

    # optional proposal PDF attachment
    if d.get("attach_proposal") and d.get("proposal"):
        try:
            pdf = _pdf_bytes_for(d["proposal"])
            name = d.get("attachment_name") or ((d["proposal"].get("company") or "Proposal") + " - proposal.pdf")
            message["attachments"] = [{
                "@odata.type": "#microsoft.graph.fileAttachment",
                "name": name,
                "contentType": "application/pdf",
                "contentBytes": _b64.b64encode(pdf).decode("ascii"),
            }]
        except Exception as e:
            print("=== SEND: PDF ATTACHMENT BUILD ERROR ===", flush=True)
            traceback.print_exc()
            print("=== END ERROR ===", flush=True)
            return jsonify(error="Could not build the PDF attachment", detail=str(e)), 500

    try:
        token = _graph_token()
    except Exception as e:
        return jsonify(error=str(e)), 500

    try:
        r = _rq.post(
            "https://graph.microsoft.com/v1.0/users/%s/sendMail" % frm,
            headers={"Authorization": "Bearer " + token, "Content-Type": "application/json"},
            json={"message": message, "saveToSentItems": True}, timeout=60)
    except Exception as e:
        return jsonify(error="Send request failed", detail=str(e)), 502

    if r.status_code in (200, 202):
        return jsonify(ok=True, sent_to=to, attached=bool(message.get("attachments")))
    return jsonify(error="Microsoft 365 rejected the send", status=r.status_code, detail=r.text[:500]), 502


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8000))
    app.run(host="0.0.0.0", port=port)
