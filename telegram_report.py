"""Send the daily GEX / Max Pain summary (and alerts) to Telegram, to read on a phone.

Secrets live in telegram.json next to this file ({"token": ..., "chat_id": ...}), which is
git-ignored. Create it once with `python telegram_report.py --setup`. Without it every send
is a silent no-op, and a failed send only prints - a Telegram outage must never break the
daily run.
"""
import json
import sys
import tempfile
import time
import urllib.parse
import urllib.request
import uuid
from datetime import timedelta, timezone
from pathlib import Path

HERE = Path(__file__).parent
SECRETS = HERE / "telegram.json"
API = "https://api.telegram.org/bot{token}/{method}"


def load_secrets():
    try:
        s = json.loads(SECRETS.read_text(encoding="utf-8"))
        return s if s.get("token") and s.get("chat_id") else None
    except (OSError, ValueError):
        return None


def _call(token, method, data=None, files=None, timeout=20):
    url = API.format(token=token, method=method)
    if files:
        boundary = uuid.uuid4().hex
        body = b""
        for k, v in (data or {}).items():
            body += (f"--{boundary}\r\nContent-Disposition: form-data; name=\"{k}\"\r\n\r\n"
                     f"{v}\r\n").encode("utf-8")
        for k, (name, blob, ctype) in files.items():
            body += (f"--{boundary}\r\nContent-Disposition: form-data; name=\"{k}\"; "
                     f"filename=\"{name}\"\r\nContent-Type: {ctype}\r\n\r\n").encode("utf-8")
            body += blob + b"\r\n"
        body += f"--{boundary}--\r\n".encode("utf-8")
        req = urllib.request.Request(url, body, {"Content-Type": f"multipart/form-data; boundary={boundary}"})
    else:
        req = urllib.request.Request(url, urllib.parse.urlencode(data or {}).encode("utf-8"))
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return json.loads(r.read().decode("utf-8"))


def _call_retry(*args, attempts=3, **kw):
    """Telegram calls sometimes time out on this connection (seen 2026-10-07); retry with a
    longer timeout. A retried send can rarely arrive twice - better than not at all."""
    timeout = kw.pop("timeout", 20)
    for i in range(attempts):
        try:
            return _call(*args, timeout=timeout * (i + 1), **kw)
        except Exception:
            if i == attempts - 1:
                raise
            time.sleep(3)


def send_text(text):
    s = load_secrets()
    if not s:
        return False
    try:
        _call_retry(s["token"], "sendMessage", {"chat_id": s["chat_id"], "text": text[:4000]})
        return True
    except Exception as e:
        print(f"telegram: send failed ({type(e).__name__}: {e})")
        return False


def send_photo(png_path, caption):
    s = load_secrets()
    if not s:
        return False
    try:
        blob = Path(png_path).read_bytes()
        _call_retry(s["token"], "sendPhoto", {"chat_id": s["chat_id"], "caption": caption[:1024]},
                    {"photo": ("gex.png", blob, "image/png")}, timeout=40)
        return True
    except Exception as e:
        print(f"telegram: photo failed ({type(e).__name__}: {e}) - sending text only")
        return send_text(caption)


def _fmt(v, nd=0):
    if isinstance(v, (int, float)):
        return f"{v:,.{nd}f}"
    return "-" if v in (None, "") else str(v)


def _num_or_none(v):
    return v if isinstance(v, (int, float)) else None


def _lvl(k, price):
    """'4,200 (+12.9)' - a level with its distance from the price."""
    if not isinstance(k, (int, float)):
        return _fmt(k)
    return f"{k:,.0f} ({k - price:+,.1f})" if price else f"{k:,.0f}"


def _ranges(ks, step):
    """[4190, 4195, 4205] -> '4,190–4,195, 4,205'"""
    out, start, prev = [], None, None
    for k in ks:
        if start is None:
            start = prev = k
        elif k - prev <= step:
            prev = k
        else:
            out.append((start, prev))
            start = prev = k
    if start is not None:
        out.append((start, prev))
    return ", ".join(f"{a:,.0f}" if a == b else f"{a:,.0f}–{b:,.0f}" for a, b in out)


def dense_strikes(rows, price, window=0.015, n=2):
    """Strongest gross-gamma strikes just above / below the price (within ±window).

    Used instead of the workbook's จุดหนืด column, whose 2x-median rule breaks when most
    strikes are zero (median 0 -> every strike flagged, seen 2026-10-07)."""
    near = [(k, c + p) for k, c, p in rows if abs(k - price) <= price * window and c + p > 0]
    above = sorted((x for x in near if x[0] > price), key=lambda x: -x[1])[:n]
    below = sorted((x for x in near if x[0] <= price), key=lambda x: -x[1])[:n]
    return sorted(above), sorted(below, reverse=True)


def air_pockets(rows, lo, hi):
    """Strikes strictly between two levels with no gamma at all."""
    ks = [k for k, c, p in rows if lo < k < hi and c + p == 0]
    step = min((b[0] - a[0] for a, b in zip(rows, rows[1:])), default=5)
    return _ranges(ks, step)


def python_mode(rows, net):
    """GEX Calc!J5's rule, for when Excel can't be read: conviction = |NET| / sum|call-put|,
    under 0.05 is 'no mode' (a bare net>0 test called NET -10 'Negative', 2026-10-07)."""
    den = sum(abs(c - p) for _, c, p in rows)
    if not den or abs(net) / den < 0.05:
        return "ไม่มีโหมด"
    return "Positive GEX (นิ่ง)" if net > 0 else "Negative GEX (วิ่ง)"


def reading(mode, price, call_wall, put_wall):
    """One or two plain sentences on what the levels usually mean - not a trade signal."""
    out = []
    mode = mode or ""
    if "Positive" in mode:
        out.append("Positive GEX: dealer มักขายตอนขึ้น/ซื้อตอนลง → แกว่งแคบ ราคามักถูกดึงเข้าหา Pins")
    elif "Negative" in mode:
        out.append("Negative GEX: dealer มักไล่ตามทิศราคา → แกว่งได้แรง ใช้ Wall เป็นเส้นอ้างอิง")
    elif mode:
        out.append("Call/Put Gamma ใกล้กัน → สัญญาณ GEX ไม่ชัด อย่าพึ่งพา")
    if price and call_wall is not None and put_wall is not None:
        if put_wall <= price <= call_wall:
            out.append(f"ราคาอยู่ในกรอบ Put Wall–Call Wall (กว้าง {call_wall - put_wall:,.0f})")
        elif price > call_wall:
            out.append("ราคาอยู่เหนือ Call Wall")
        else:
            out.append("ราคาอยู่ใต้ Put Wall")
    return out


def format_report(mp, gex, xl):
    """Phone-sized summary. mp: Max Pain dict (or None), gex: update_gex's result (or None),
    xl: {cell: value} read back from Excel's own calculation of GEX Calc (may be empty)."""
    xl = xl or {}
    lines = []
    if gex:
        lines.append(f"📊 {gex['code']} ({gex['label']}) · ข้อมูล {gex['trade_date']} {gex.get('report') or ''}".rstrip())
    elif mp:
        lines.append(f"📊 {mp['contract']} · ข้อมูล {mp['trade_date']} {mp['report']} (GEX ไม่ได้อัปเดต)")
    price = (gex or mp or {}).get("price")
    top = f"ราคา {_fmt(price, 1)}"
    if mp and mp.get("dte") is not None:
        top += f" · DTE {mp['dte']}"
    lines.append(top)

    if gex:
        rows = gex["rows"]
        cw, pw = _num_or_none(xl.get("J9")), _num_or_none(xl.get("J10"))
        mode = xl.get("J5") or python_mode(rows, gex["net"])
        lines += ["", "— GEX —"]
        if not xl:
            lines.append("(อ่านค่าจาก Excel ไม่ได้ - ไม่มี Wall/Pins รอบนี้)")
        conv = xl.get("J18")
        conv = f" · Conviction {float(conv):.2f} {xl.get('J19', '')}".rstrip() if conv not in (None, "") else ""
        lines.append(f"NET GEX {gex['net']:+,.0f} → {mode}{conv}")
        lines.append(f"Call/Put Gamma {gex['calls']:,.0f} / {gex['puts']:,.0f}"
                     + (f" · Gamma Flip {_lvl(gex['flip'], price)}" if gex.get("flip") else ""))
        if xl.get("J27"):
            lines.append(f"แรงดูด: {xl['J27']}")
        if cw is not None or pw is not None:
            lines.append(f"Call Wall {_lvl(cw, price)} · Put Wall {_lvl(pw, price)}")
        lines.append(f"หนืดสุด {_lvl(_num_or_none(xl.get('J11')) or gex['peak'], price)}")
        pins = []
        for r in range(30, 35):
            k = _num_or_none(xl.get(f"J{r}"))
            if k is not None:
                side = xl.get(f"M{r}")
                pins.append(_lvl(k, price) + (f" {side}" if side and side != "(mixed)" else ""))
        if pins:
            lines.append("Pins: " + ", ".join(pins))
        if price:
            above, below = dense_strikes(rows, price)
            if above or below:
                fmt = lambda xs: ", ".join(f"{k:,.0f} ({g:,.0f})" for k, g in xs) or "-"
                lines.append(f"Gamma หนาแน่นใกล้ราคา: เหนือ {fmt(above)} · ใต้ {fmt(below)}")
            if cw is not None and pw is not None:
                gaps = air_pockets(rows, min(pw, price), max(cw, price))
                if gaps:
                    lines.append(f"ช่องว่าง (gamma 0) ในกรอบ Wall: {gaps}")
        read = reading(xl.get("J5") or mode, price, cw, pw)
        if read:
            lines.append("อ่านผล: " + " · ".join(read))

    if mp:
        lines += ["", "— Open Interest / Max Pain —"]
        dist = f" ({(mp['max_pain'] - price) / price:+.1%})" if price else ""
        z = f" · โซน {mp['zone']}" if mp.get("zone") else ""
        lines.append(f"Max Pain {_fmt(mp['max_pain'])}{dist}{z}")
        ct, pt = mp.get("call_total"), mp.get("put_total")
        if ct:
            lines.append(f"OI รวม Call {ct:,.0f} / Put {pt:,.0f} · P/C {pt / ct:.2f}")
        strikes = mp.get("strikes") or []
        for name, i in (("Call", 1), ("Put", 2)):
            top3 = sorted((s for s in strikes if s[i] > 0), key=lambda s: -s[i])[:3]
            if top3:
                lines.append(f"{name} OI สูงสุด: " + ", ".join(f"{s[0]:,.0f} ({s[i]:,.0f})" for s in top3))
    return "\n".join(lines)


def render_chart(rows, price, call_wall=None, put_wall=None, title="", path=None, window=0.05):
    """Call gamma up / put gamma down per strike around the price; returns the PNG path."""
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt

    if price:
        near = [r for r in rows if abs(r[0] - price) <= price * window]
        rows = near or rows
    ks = [r[0] for r in rows]
    width = min((b - a for a, b in zip(ks, ks[1:])), default=5) * 0.8
    fig, ax = plt.subplots(figsize=(6, 7.5), dpi=150)
    ax.barh(ks, [r[1] for r in rows], height=width, color="#2e7d32", label="Call gamma")
    ax.barh(ks, [-r[2] for r in rows], height=width, color="#c62828", label="Put gamma")
    ax.axvline(0, color="#555", lw=0.8)
    if price:
        ax.axhline(price, color="#1565c0", lw=1.4, ls="--", label=f"Price {price:,.1f}")
    for lvl, name, col in ((call_wall, "Call Wall", "#2e7d32"), (put_wall, "Put Wall", "#c62828")):
        if isinstance(lvl, (int, float)):
            ax.axhline(lvl, color=col, lw=1, ls=":")
            ax.annotate(f"{name} {lvl:,.0f}", (ax.get_xlim()[1], lvl), ha="right", va="bottom",
                        fontsize=8, color=col)
    ax.set_xlabel("Gamma (1 Pct)   put ←  → call")
    ax.set_ylabel("Strike")
    ax.set_title(title, fontsize=10)
    ax.legend(loc="lower right", fontsize=8)
    ax.grid(axis="x", alpha=0.3)
    fig.tight_layout()
    path = path or Path(tempfile.gettempdir()) / "maxpain_gex_chart.png"
    fig.savefig(path)
    plt.close(fig)
    return path


def render_maxpain_chart(strikes, price, max_pain, title="", path=None, window=0.05):
    """Call OI up / put OI down per strike, with the total-pain curve on a second axis."""
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    from update_workbooks import pain_by_strike

    pain = pain_by_strike(strikes)  # over every strike, so the curve matches the sheet
    shown = [s for s in strikes if price and abs(s[0] - price) <= price * window] or strikes
    ks = [s[0] for s in shown]
    width = min((b - a for a, b in zip(ks, ks[1:])), default=5) * 0.8
    fig, ax = plt.subplots(figsize=(7, 5), dpi=150)
    ax.bar(ks, [s[1] for s in shown], width=width, color="#2e7d32", label="Call OI")
    ax.bar(ks, [-s[2] for s in shown], width=width, color="#c62828", label="Put OI")
    ax.axhline(0, color="#555", lw=0.8)
    ax.set_ylabel("Open Interest   put ↓  ↑ call")
    ax2 = ax.twinx()
    ax2.plot(ks, [pain[k] / 1e3 for k in ks], color="#6a1b9a", lw=1.6, label="Total pain (k)")
    ax2.set_ylabel("Total pain (thousands)")
    if price:
        ax.axvline(price, color="#1565c0", lw=1.4, ls="--", label=f"Price {price:,.1f}")
    ax.axvline(max_pain, color="#6a1b9a", lw=1.2, ls=":", label=f"Max Pain {max_pain:,.0f}")
    ax.set_xlabel("Strike")
    ax.set_title(title, fontsize=10)
    h1, l1 = ax.get_legend_handles_labels()
    h2, l2 = ax2.get_legend_handles_labels()
    ax.legend(h1 + h2, l1 + l2, loc="upper left", fontsize=7)
    ax.grid(axis="y", alpha=0.3)
    fig.tight_layout()
    path = path or Path(tempfile.gettempdir()) / "maxpain_pain_chart.png"
    fig.savefig(path)
    plt.close(fig)
    return path


CARD_CSS = """
body{margin:0;background:#eef1f5;font-family:'Leelawadee UI','Segoe UI',Tahoma,sans-serif}
.card{width:440px;background:#fff;padding:14px 16px 12px;box-sizing:border-box;color:#1d2433;font-size:15px}
.hd{display:flex;justify-content:space-between;align-items:baseline}
.code{font-size:21px;font-weight:700}.sub{color:#68707d;font-size:13px}
.px{font-size:15px;margin-top:2px}.px b{font-size:19px}
.mode{margin:10px 0 4px;padding:8px 10px;border-radius:8px;color:#fff;font-weight:700;font-size:17px}
.pos{background:#2e7d32}.neg{background:#c62828}.none{background:#78808c}
.mode small{display:block;font-weight:400;font-size:13px;opacity:.95}
.warn{background:#fff3e0;color:#a84300;border:1px solid #ffb74d;border-radius:6px;padding:5px 8px;font-size:13px;margin:6px 0}
h3{font-size:13px;color:#68707d;margin:12px 0 4px;font-weight:600;letter-spacing:.3px}
table{width:100%;border-collapse:collapse}
td{padding:5px 6px;border-bottom:1px solid #edf0f3}
td.k{text-align:right;font-weight:700;font-variant-numeric:tabular-nums}
td.d{text-align:right;width:70px;font-variant-numeric:tabular-nums}
.up{color:#2e7d32}.dn{color:#c62828}
tr.price td{background:#e3f2fd;font-weight:700;color:#0d47a1;border-bottom:2px solid #90caf9}
.tag{display:inline-block;padding:1px 7px;border-radius:10px;font-size:12px;font-weight:700;color:#fff;margin-right:3px}
.cw{background:#2e7d32}.pw{background:#c62828}.mp{background:#6a1b9a}.pk{background:#ef6c00}.pin{background:#546e7a}.fl{background:#00838f}
.side{color:#68707d;font-size:12px}
.two{display:flex;gap:10px}.two>div{flex:1}
.note{font-size:13px;color:#3d4554;margin-top:8px;line-height:1.45}
.foot{font-size:11px;color:#9aa1ab;margin-top:8px;text-align:right}
"""


def price_note(gex, mp):
    """Where the reference price came from - a live quote vs the previous close is the usual
    reason it won't match the chart on the phone."""
    src = gex or mp or {}
    settle = (gex or {}).get("settle") or (mp or {}).get("price")
    t = (gex or {}).get("live_time")
    fut = f"GC {(mp or {}).get('price_month') or ''}".strip()
    if t:
        bkk = t.astimezone(timezone(timedelta(hours=7)))
        return (f"{fut} สด {bkk:%H:%M} น. (CME ดีเลย์ 10 นาที) · ปิด {src.get('trade_date')} "
                f"{_fmt(settle, 1)}")
    return f"ราคาปิด {fut} ของ {src.get('trade_date')} (ไม่มีราคาสด)"


def _mode_class(mode):
    return "pos" if "Positive" in (mode or "") else "neg" if "Negative" in (mode or "") else "none"


def card_html(mp, gex, xl):
    """The phone summary as one HTML card (rendered to PNG by render_card)."""
    from html import escape as e
    xl = xl or {}
    src = gex or mp or {}
    price = src.get("price")
    title = gex["code"] if gex else (mp or {}).get("contract", "")
    label = gex["label"] if gex else ""
    report = (src.get("report") or "").replace("PRELIMINARY", "PRELIM")
    out = [f"<div class='card'><div class='hd'><span class='code'>{e(str(title))}</span>"
           f"<span class='sub'>{e(label)}</span></div>",
           f"<div class='sub'>ข้อมูลวันที่ {e(str(src.get('trade_date', '')))} · {e(report)}</div>",
           f"<div class='px'>ราคา <b>{_fmt(price, 1)}</b>"
           + (f" &nbsp;·&nbsp; DTE {mp['dte']}" if mp and mp.get("dte") is not None else "") + "</div>",
           f"<div class='sub'>{e(price_note(gex, mp))}</div>"]

    def dist(k):
        if not (price and isinstance(k, (int, float))):
            return "<td class='d'></td>"
        d = k - price
        return f"<td class='d {'up' if d >= 0 else 'dn'}'>{d:+,.1f}</td>"

    if gex:
        rows = gex["rows"]
        mode = xl.get("J5") or python_mode(rows, gex["net"])
        conv = xl.get("J18")
        conv = f" · Conviction {float(conv):.2f} {e(str(xl.get('J19', '')))}" if conv not in (None, "") else ""
        pull = f"<br>แรงดูด: {e(str(xl['J27']))}" if xl.get("J27") else ""
        out.append(f"<div class='mode {_mode_class(mode)}'>{e(mode)}"
                   f"<small>NET GEX {gex['net']:+,.0f} (Call {gex['calls']:,.0f} / Put {gex['puts']:,.0f}){conv}{pull}</small></div>")
        if not xl:
            out.append("<div class='warn'>อ่านค่าจาก Excel ไม่ได้ - รอบนี้ไม่มี Wall/Pins</div>")

        levels = {}  # strike -> [tags]
        def add(k, cls, name):
            if isinstance(k, (int, float)):
                levels.setdefault(k, []).append(f"<span class='tag {cls}'>{e(name)}</span>")
        add(_num_or_none(xl.get("J9")), "cw", "Call Wall")
        add(_num_or_none(xl.get("J10")), "pw", "Put Wall")
        if mp:
            add(mp.get("max_pain"), "mp", "Max Pain")
        add(_num_or_none(xl.get("J11")) or gex["peak"], "pk", "หนืดสุด")
        add(gex.get("flip"), "fl", "Gamma Flip")
        for i, r in enumerate(range(30, 35), 1):
            k = _num_or_none(xl.get(f"J{r}"))
            if k is not None:
                side = xl.get(f"M{r}")
                levels.setdefault(k, []).append(
                    f"<span class='tag pin'>Pin {i}</span>"
                    + (f"<span class='side'>{e(side)}</span>" if side and side != "(mixed)" else ""))
        out.append("<h3>ระดับสำคัญ (เรียงตาม strike)</h3><table>")
        placed = False
        for k in sorted(levels, reverse=True):
            if price and not placed and k < price:
                out.append(f"<tr class='price'><td>◀ ราคาปัจจุบัน</td><td class='k'>{price:,.1f}</td><td class='d'></td></tr>")
                placed = True
            out.append(f"<tr><td>{' '.join(levels[k])}</td><td class='k'>{k:,.0f}</td>{dist(k)}</tr>")
        if price and not placed:
            out.append(f"<tr class='price'><td>◀ ราคาปัจจุบัน</td><td class='k'>{price:,.1f}</td><td class='d'></td></tr>")
        out.append("</table>")

        if price:
            above, below = dense_strikes(rows, price)
            if above or below:
                f = lambda xs: ", ".join(f"<b>{k:,.0f}</b> ({g:,.0f})" for k, g in xs) or "-"
                out.append(f"<div class='note'>Gamma หนาแน่นใกล้ราคา: เหนือ {f(above)} · ใต้ {f(below)}</div>")
            cw, pw = _num_or_none(xl.get("J9")), _num_or_none(xl.get("J10"))
            if cw is not None and pw is not None:
                gaps = air_pockets(rows, min(pw, price), max(cw, price))
                if gaps:
                    out.append(f"<div class='note'>ช่องว่าง (gamma 0) ในกรอบ Wall: {e(gaps)}</div>")
            read = reading(mode, price, cw, pw)
            if read:
                out.append(f"<div class='note'><b>อ่านผล:</b> {e(' · '.join(read))}</div>")
    elif mp:
        out.append("<div class='warn'>GEX ไม่ได้อัปเดตรอบนี้</div>")

    if mp:
        ct, pt = mp.get("call_total"), mp.get("put_total")
        dmp = f" ({(mp['max_pain'] - price) / price:+.1%})" if price else ""
        out.append("<h3>Open Interest / Max Pain</h3><table>")
        out.append(f"<tr><td>Max Pain</td><td class='k'>{_fmt(mp['max_pain'])}{dmp}</td>"
                   f"<td class='d'>{e(mp.get('zone') or '')}</td></tr>")
        if ct:
            out.append(f"<tr><td>OI รวม <span class='up'>Call</span> / <span class='dn'>Put</span></td>"
                       f"<td class='k'>{ct:,.0f} / {pt:,.0f}</td><td class='d'>P/C {pt / ct:.2f}</td></tr>")
        out.append("</table>")
        strikes = mp.get("strikes") or []
        cols = []
        for name, i, cls in (("Call OI สูงสุด", 1, "up"), ("Put OI สูงสุด", 2, "dn")):
            top3 = sorted((s for s in strikes if s[i] > 0), key=lambda s: -s[i])[:3]
            body = "".join(f"<tr><td class='k {cls}'>{s[0]:,.0f}</td><td class='d'>{s[i]:,.0f}</td></tr>" for s in top3)
            cols.append(f"<div><h3>{name}</h3><table>{body}</table></div>")
        out.append(f"<div class='two'>{''.join(cols)}</div>")
    out.append("<div class='foot'>CME / QuikStrike · ไม่ใช่คำแนะนำการลงทุน</div></div>")
    return (f"<!doctype html><html><head><meta charset='utf-8'><style>{CARD_CSS}</style></head>"
            f"<body>{''.join(out)}</body></html>")


def render_card(html_text, path=None):
    """Screenshot the card in headless Chrome (proper Thai shaping, unlike matplotlib).
    Persistent context because Browser.close() hangs on this machine (see fetch_cme.py)."""
    from playwright.sync_api import sync_playwright
    path = path or Path(tempfile.gettempdir()) / "maxpain_card.png"
    with sync_playwright() as p, tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as tmp:
        ctx = p.chromium.launch_persistent_context(tmp, channel="chrome", headless=True,
                                                   viewport={"width": 460, "height": 900},
                                                   device_scale_factor=2)
        try:
            pg = ctx.pages[0] if ctx.pages else ctx.new_page()
            pg.set_content(html_text, wait_until="load")
            pg.evaluate("document.fonts.ready")
            pg.locator(".card").screenshot(path=str(path))
        finally:
            ctx.close()
    return path


def headline(mp, gex, xl):
    """One-line caption - what the phone notification shows."""
    xl = xl or {}
    if not gex:
        return f"Max Pain {_fmt((mp or {}).get('max_pain'))} · {(mp or {}).get('contract', '')} · GEX ไม่ได้อัปเดต"
    mode = xl.get("J5") or python_mode(gex["rows"], gex["net"])
    parts = [gex["code"], mode]
    if _num_or_none(xl.get("J9")) is not None:
        parts.append(f"Call Wall {xl['J9']:,.0f} / Put Wall {_fmt(xl.get('J10'))}")
    if mp:
        parts.append(f"Max Pain {mp['max_pain']:,.0f}")
    return " · ".join(parts)


def send_album(paths, caption):
    """All pictures as one Telegram album, caption on the first; one by one if that fails."""
    s = load_secrets()
    if not s:
        return False
    paths = [Path(p) for p in paths]
    media = [{"type": "photo", "media": f"attach://p{i}", **({"caption": caption[:1024]} if i == 0 else {})}
             for i in range(len(paths))]
    try:
        _call_retry(s["token"], "sendMediaGroup",
                    {"chat_id": s["chat_id"], "media": json.dumps(media, ensure_ascii=False)},
                    {f"p{i}": (p.name, p.read_bytes(), "image/png") for i, p in enumerate(paths)},
                    timeout=60)
        return True
    except Exception as e:
        print(f"telegram: album failed ({type(e).__name__}: {e}) - sending photos one by one")
        ok = [send_photo(p, caption if i == 0 else "") for i, p in enumerate(paths)]
        return all(ok)


XL_CELLS = ["J5", "J9", "J10", "J11", "J18", "J19", "J27",
            *[f"{c}{r}" for r in range(30, 35) for c in "JM"]]


def send_daily_report(mp, gex):
    """End-of-run summary + gamma chart. Walls/Pins/mode come from Excel's own calculation of
    the saved GEX workbook (verify.excel_cells), so the phone shows exactly what the sheet
    shows; if Excel is unavailable or the workbook couldn't be saved, Python totals only."""
    if not load_secrets():
        return
    xl = {}
    if gex and gex.get("saved"):
        try:
            from verify import excel_cells
            cells = {}
            for _ in range(2):  # Excel COM occasionally returns nothing on the first try
                cells = excel_cells(Path(gex["path"]).resolve(), {"GEX Calc": XL_CELLS}) or {}
                if cells:
                    break
            xl = {k.split("!")[1]: v for k, v in cells.items()}
            for k in ("J9", "J10", "J11", "J18", *[f"J{r}" for r in range(30, 35)]):
                try:
                    xl[k] = float(xl[k])
                except (KeyError, TypeError, ValueError):
                    pass
        except Exception as e:
            print(f"telegram: Excel read-back failed ({type(e).__name__}: {e})")
    # one album: summary card, gamma chart, Max Pain chart; a one-line caption carries the
    # gist into the phone notification. The long text report is only the fallback now.
    pics = []
    try:
        pics.append(render_card(card_html(mp, gex, xl)))
    except Exception as e:
        print(f"telegram: summary card failed ({type(e).__name__}: {e})")
    if gex:
        try:
            pics.append(render_chart(gex["rows"], gex["price"], _num_or_none(xl.get("J9")),
                                     _num_or_none(xl.get("J10")), title=f"GEX {gex['code']} · {gex['trade_date']}"))
        except Exception as e:
            print(f"telegram: gamma chart failed ({type(e).__name__}: {e})")
    if mp and mp.get("strikes"):
        try:
            pics.append(render_maxpain_chart(mp["strikes"], mp.get("price"), mp["max_pain"],
                                             title=f"Max Pain {mp['contract']} · {mp['trade_date']}"))
        except Exception as e:
            print(f"telegram: max pain chart failed ({type(e).__name__}: {e})")
    if pics:
        send_album(pics, headline(mp, gex, xl))
    if not pics or "card" not in Path(pics[0]).name:  # no card -> the text report instead
        send_text(format_report(mp, gex, xl))


def setup():
    """One-time: ask for the BotFather token, wait for /start, save telegram.json."""
    print("1) In Telegram, open @BotFather, send /newbot, follow the steps, copy the token.")
    token = input("2) Paste the bot token here: ").strip()
    me = _call(token, "getMe")
    name = me["result"]["username"]
    print(f"3) Open https://t.me/{name} on your phone and press Start (or send /start). Waiting...")
    for _ in range(60):
        upd = _call(token, "getUpdates", {"timeout": 0})["result"]
        chats = [u["message"]["chat"]["id"] for u in upd if "message" in u]
        if chats:
            SECRETS.write_text(json.dumps({"token": token, "chat_id": chats[-1]}), encoding="utf-8")
            send_text("เชื่อมต่อแล้ว ✅ สรุป GEX / Max Pain จะส่งมาที่นี่ทุกครั้งที่รันเสร็จ")
            print(f"Saved {SECRETS.name} - a test message was sent to your phone.")
            return
        time.sleep(5)
    sys.exit("No /start received within 5 minutes - run --setup again.")


if __name__ == "__main__":
    if "--setup" in sys.argv:
        setup()
    else:
        print(send_text(" ".join(sys.argv[1:]) or "test from telegram_report.py"))
