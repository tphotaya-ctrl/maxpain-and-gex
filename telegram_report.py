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


def send_text(text):
    s = load_secrets()
    if not s:
        return False
    try:
        _call(s["token"], "sendMessage", {"chat_id": s["chat_id"], "text": text[:4000]})
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
        _call(s["token"], "sendPhoto", {"chat_id": s["chat_id"], "caption": caption[:1024]},
              {"photo": ("gex.png", blob, "image/png")}, timeout=40)
        return True
    except Exception as e:
        print(f"telegram: photo failed ({type(e).__name__}: {e}) - sending text only")
        return send_text(caption)


def _fmt(v, nd=0):
    if isinstance(v, (int, float)):
        return f"{v:,.{nd}f}"
    return "-" if v in (None, "") else str(v)


def format_report(mp, gex, xl):
    """Phone-sized summary. mp: Max Pain dict (or None), gex: update_gex's result (or None),
    xl: {cell: value} read back from Excel's own calculation of GEX Calc (may be empty)."""
    xl = xl or {}
    lines = []
    if gex:
        lines.append(f"GEX {gex['code']} ({gex['label']}) · ข้อมูล {gex['trade_date']} {gex.get('report', '')}".rstrip())
    elif mp:
        lines.append(f"{mp['contract']} · ข้อมูล {mp['trade_date']} {mp['report']} (GEX ไม่ได้อัปเดต)")
    price = (gex or mp or {}).get("price")
    top = f"ราคา {_fmt(price, 1)}"
    if mp:
        dist = f" ({(mp['max_pain'] - price) / price:+.1%})" if price else ""
        top += f" · Max Pain {_fmt(mp['max_pain'])}{dist}"
    lines.append(top)
    if gex:
        mode = xl.get("J5") or ("Positive GEX" if gex["net"] > 0 else "Negative GEX")
        conv = xl.get("J18")
        conv = f" · Conviction {float(conv):.2f} {xl.get('J19', '')}".rstrip() if conv not in (None, "") else ""
        lines.append(f"NET GEX {gex['net']:+,.0f} → {mode}{conv}")
        walls = [f"Call Wall {_fmt(xl.get('J9'))}", f"Put Wall {_fmt(xl.get('J10'))}"] if xl else []
        lines.append(" · ".join(walls + [f"หนืดสุด {_fmt(xl.get('J11') or gex['peak'])}"]))
        pins = [_fmt(xl.get(f"J{r}")) for r in range(30, 35) if xl.get(f"J{r}") not in (None, "")]
        if pins:
            lines.append("Pins " + ", ".join(pins))
        lines.append(f"Call/Put Gamma {gex['calls']:,.0f} / {gex['puts']:,.0f}"
                     + (f" · Gamma Flip {_fmt(gex['flip'])}" if gex.get("flip") else ""))
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


XL_CELLS = ["J5", "J9", "J10", "J11", "J18", "J19", "J30", "J31", "J32", "J33", "J34"]


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
            cells = excel_cells(Path(gex["path"]).resolve(), {"GEX Calc": XL_CELLS}) or {}
            xl = {k.split("!")[1]: v for k, v in cells.items()}
            for k in ("J9", "J10", "J11", "J18", *[f"J{r}" for r in range(30, 35)]):
                try:
                    xl[k] = float(xl[k])
                except (KeyError, TypeError, ValueError):
                    pass
        except Exception as e:
            print(f"telegram: Excel read-back failed ({type(e).__name__}: {e})")
    text = format_report(mp, gex, xl)
    if gex:
        try:
            png = render_chart(gex["rows"], gex["price"], xl.get("J9"), xl.get("J10"),
                               title=f"{gex['code']} · {gex['trade_date']}")
            send_photo(png, text)
            return
        except Exception as e:
            print(f"telegram: chart failed ({type(e).__name__}: {e})")
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
