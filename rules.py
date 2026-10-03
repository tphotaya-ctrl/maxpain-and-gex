"""Paper-trading rules, fixed on 2026-10-03 before any results were seen.

Changing a rule makes its earlier results in-sample: note the date and count from there.
Each rule takes one Outcome record and returns a Trade or None (no trade).

Assumptions (see README, "Paper trades"):
- entry = open of the first trading day after the data date (day T's data only exists after
  T closes), exit = futures settle on the expiry day, P&L in futures points;
- stops are checked worst-case: only the period's daily high/low are known, not the order
  things happened, so if the stop is inside high..low the trade is stopped at the stop.
"""
from dataclasses import dataclass

from update_gex import _mode

MP_MIN_GAP = 0.005  # R3 only trades when Max Pain is at least 0.5% from entry


@dataclass
class Trade:
    direction: int  # +1 long, -1 short
    entry: float
    exit: float
    stopped: bool

    @property
    def pnl(self):
        return (self.exit - self.entry) * self.direction


def _trade(rec, direction, stop=None):
    entry, settle, hi, lo = rec["entry"], rec["settle"], rec["hi"], rec["lo"]
    if entry is None or settle is None:
        return None
    if stop is not None:
        if hi is None or lo is None:
            return None  # can't judge the stop
        if (direction > 0 and lo <= stop) or (direction < 0 and hi >= stop):
            return Trade(direction, entry, stop, True)
    return Trade(direction, entry, settle, False)


def r1_range_fade(rec):
    """Positive GEX, entry inside (Put Wall, Call Wall): buy the lower half, sell the upper
    half; stop at the wall behind the trade."""
    pw, cw, entry = rec["put_wall"], rec["call_wall"], rec["entry"]
    if _mode(rec["mode"]) != "+" or pw is None or cw is None or entry is None or not pw < entry < cw:
        return None
    mid = (pw + cw) / 2
    if entry < mid:
        return _trade(rec, +1, stop=pw)
    if entry > mid:
        return _trade(rec, -1, stop=cw)
    return None


def r2_momentum(rec):
    """Negative GEX: follow the data day's futures change; stop at the wall behind the
    trade when there is one."""
    change, entry = rec["change"], rec["entry"]
    if _mode(rec["mode"]) != "-" or not change or entry is None:
        return None
    if change > 0:
        pw = rec["put_wall"]
        return _trade(rec, +1, stop=pw if pw is not None and pw < entry else None)
    cw = rec["call_wall"]
    return _trade(rec, -1, stop=cw if cw is not None and cw > entry else None)


def r3_max_pain(rec):
    """Any mode: trade toward Max Pain when it is at least MP_MIN_GAP away; no stop."""
    mp, entry = rec["mp"], rec["entry"]
    if mp is None or not entry or abs(mp - entry) / entry < MP_MIN_GAP:
        return None
    return _trade(rec, +1 if mp > entry else -1)


def baseline_long(rec):
    """What every rule has to beat: just buy at the same entry, sell at the same exit."""
    return _trade(rec, +1)


RULES = {
    "R1 Range fade (+GEX)": r1_range_fade,
    "R2 Momentum (-GEX)": r2_momentum,
    "R3 Max Pain magnet": r3_max_pain,
    "Baseline: always long": baseline_long,
}


def stats(trades):
    if not trades:
        return {"n": 0}
    pnls = [t.pnl for t in trades]
    return {
        "n": len(trades),
        "win_rate": sum(p > 0 for p in pnls) / len(pnls),
        "total": sum(pnls),
        "avg": sum(pnls) / len(pnls),
        "worst": min(pnls),
        "stopped": sum(t.stopped for t in trades),
    }
