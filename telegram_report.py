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
    text = format_report(mp, gex, xl)
    # pictures first with short captions, the full text last - it's what the phone's
    # notification shows, and it's too long for a photo caption (1024 chars)
    if gex:
        try:
            png = render_chart(gex["rows"], gex["price"], _num_or_none(xl.get("J9")),
                               _num_or_none(xl.get("J10")), title=f"GEX {gex['code']} · {gex['trade_date']}")
            send_photo(png, f"GEX {gex['code']} · {gex['trade_date']}")
        except Exception as e:
            print(f"telegram: gamma chart failed ({type(e).__name__}: {e})")
    if mp and mp.get("strikes"):
        try:
            png = render_maxpain_chart(mp["strikes"], mp.get("price"), mp["max_pain"],
                                       title=f"Max Pain {mp['contract']} · {mp['trade_date']}")
            send_photo(png, f"Max Pain {mp['max_pain']:,.0f} · {mp['contract']}")
        except Exception as e:
            print(f"telegram: max pain chart failed ({type(e).__name__}: {e})")
    send_text(text)


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
