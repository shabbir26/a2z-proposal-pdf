"""
A2Z Proposal PDF service
-------------------------
A thin web wrapper around Shabbir's own a2z_proposals_fpdf.py generator.
The platform POSTs proposal data as JSON; this builds the in-memory workbook
his read_ltd / read_sa / read_partnership expect, then calls his UNCHANGED
build_ltd / build_sa / build_partnership and returns the byte-identical PDF.

His generator code is never modified - it is imported and called as-is.
"""
import io, os, tempfile, datetime
from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import openpyxl

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
    """His proposal email, reproduced verbatim for the unchanged parts, PLUS:
    an itemised setup-cost breakdown, and one of three onboarding scenarios
    (newco = we register / switcher = moving from another firm / existing =
    has a business but no previous accountant). His a2z_proposals_fpdf.py is
    NOT modified - this layers on top of the data his read_* returns."""
    K = kind.upper()
    gbp = GEN.gbp
    num = GEN.num
    FORMS = GEN.FORMS
    name = (str(d.get("contact") or "there").strip()) or "there"
    company = (str(d.get("company") or "your business").strip()) or "your business"
    subject = "Your A2Z proposal for %s" % company
    svc = [str(r[0]).strip() for r in d.get("lines", []) if str(r[0]).strip()]
    regs = d.get("regs") or []
    keys = {r.get("key") for r in regs}

    L = ["Hi %s," % name, "",
         ("Thank you for the opportunity to look after %s \u2014 it's a pleasure to put this "
          "proposal together for you. I've attached the full proposal, and here's a quick summary." % company),
         "", "What we'd take care of for you:"]
    L += ["\u2022 %s" % x for x in (svc or ["the services we discussed"])]
    L += [""]

    # ---- fee line (his exact wording) ----
    if K == "LTD":
        L += ["Your fee would be %s + VAT a month \u2014 a single fixed amount, with no surprise bills along the way." % gbp(d.get("sub", 0))]
        if (num(d.get("discount", 0)) or 0) > 0:
            L += ["As an exceptional act of discretion, a goodwill discount of %s + VAT per month has been applied. Your fee above already reflects this." % gbp(d["discount"])]
        nd = int(num(d.get("directors", 0)) or 0)
        if nd == 1:
            L += ["Each director's personal tax return (self-assessment) is \u00a3120 + VAT a year, billed separately."]
        elif nd > 1:
            L += ["Each director's personal tax return (self-assessment) is \u00a3120 + VAT a year, billed separately \u2014 for %d directors that comes to %s + VAT a year." % (nd, gbp(nd * 120))]
    else:  # SA
        m = num(d.get("monthly", 0)) or 0
        a = num(d.get("annual", 0)) or 0
        if m > 0:
            L += ["Your fee would be %s + VAT a month \u2014 a single fixed amount, with no surprise bills along the way." % gbp(d["monthly"])]
        elif a > 0:
            L += ["Your fee would be %s + VAT a year \u2014 a single fixed amount, with no surprise bills along the way." % gbp(d["annual"])]
        else:
            L += ["Your fee is set out in the attached proposal."]
    L += [""]

    # ---- itemised setup breakdown ----
    setup_items = []
    setup_total = 0
    reg_labels = set()
    for r in regs:
        lab = str(r.get("label") or "").strip().lower()
        reg_labels.add(lab)
        fee = num(r.get("fee")) or 0
        if r.get("included") or fee <= 0:
            setup_items.append("\u2022 %s \u2014 included" % r["label"])
        else:
            setup_items.append("\u2022 %s \u2014 %s + VAT" % (r["label"], gbp(fee)))
            setup_total += fee
    for o in d.get("oneoffs", []):
        desc = str(o[0]).strip() if o and len(o) > 0 else ""
        price = num(o[2]) if o and len(o) > 2 else 0
        if not desc:
            continue
        dl = desc.lower()
        # skip if this one-off is already shown as a registration (same or contained label)
        if any(dl == rl or dl in rl or rl in dl for rl in reg_labels):
            continue
        if price and price > 0:
            setup_items.append("\u2022 %s \u2014 %s + VAT" % (desc, gbp(price)))
            setup_total += price
        else:
            setup_items.append("\u2022 %s \u2014 included" % desc)
    osub = num(d.get("osub", 0)) or 0
    if osub > setup_total:   # trust the workbook subtotal if it is higher
        setup_total = osub
    if setup_items:
        L += ["To get you set up, here's the one-off work at the start:"]
        L += setup_items
        L += ["One-off setup: %s + VAT" % gbp(setup_total) if setup_total > 0 else "There's no charge for your setup \u2014 it's all included."]
        L += [""]
    elif osub > 0:
        L += ["There's also a one-off setup of %s + VAT to get everything in place at the start." % gbp(osub), ""]

    # ---- onboarding: one of three scenarios ----
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
        sc = "newco" if primary_key in keys else "existing"   # infer if not supplied
    thing = "company" if K == "LTD" else "business"

    if sc == "newco":
        if K == "LTD":
            feebit = (" (the company formation fee is %s + VAT, charged once)" % gbp(primary_fee)) if primary_fee > 0 else ""
            L += ["As we're forming %s for you, the first step is to complete our %s%s, which only takes a few minutes:" % (company, reg_name, feebit), reg_url]
        else:
            L += ["To get you registered, the first step is to complete our %s, which only takes a few minutes:" % reg_name, reg_url]
    elif sc == "switcher":
        L += ["You're moving to us from another accountant, so there's nothing for you to chase \u2014 just complete our %s and we'll write to your current accountant for professional clearance and handle the whole handover for you:" % onboard_name, onboard_url]
    else:  # existing
        if K == "LTD":
            L += ["As your %s is already set up and you've not had an accountant before, getting started is simply completing our %s, which only takes a few minutes:" % (thing, onboard_name), onboard_url]
        else:
            L += ["As you're already trading and you've not had an accountant before, getting started is simply completing our %s, which only takes a few minutes:" % onboard_name, onboard_url]

    # ---- extra registration links (non-primary), exactly as his build_email does ----
    extra = []
    seen = set()
    for r in regs:
        if r.get("key") == primary_key:
            continue
        fm = r.get("form")
        if fm and fm[1] not in seen:
            extra.append(fm)
            seen.add(fm[1])
    if extra:
        L += ["", "We'd also take care of a couple of registrations for you \u2014 you can complete those here as well:"]
        for nm_, u in extra:
            lab = nm_[:-5] if str(nm_).lower().endswith(" form") else nm_
            L += ["%s: %s" % (lab, u)]
    L += ["", "If you have any questions, or would like to talk anything through, just let me know \u2014 I'd be glad to help."]
    return subject, "\n".join(L)


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
        return GEN.build_email(GEN.read_partnership(wb2), "PARTNERSHIP")
    return _enhanced_email(GEN.read_ltd(wb2), "LTD", scenario)


@app.route("/email", methods=["POST", "OPTIONS"])
def email():
    """Return his exact proposal email as {subject, body} for the given contract."""
    if request.method == "OPTIONS":
        return ("", 204)
    d = request.get_json(force=True, silent=True) or {}
    try:
        subject, body = _email_for(d)
        return jsonify(ok=True, subject=subject, body=body)
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
            esubject, ebody = _email_for(d["proposal"])
            subject = subject or esubject
            html = _text_to_html(ebody)
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
