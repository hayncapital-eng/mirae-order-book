"""
Shared data layer for the Mirae model.

Parses the authoritative data arrays out of index.html so the dashboard stays the
single source of truth — no second copy of CONTRACTS/ACTUALS to drift out of sync.
Used by build_dcf_xlsx.py.

The arrays are JS object literals (unquoted keys, trailing commas, // comments),
so they're normalised to JSON rather than eval'd.
"""

import json
import os
import re

REPO = os.path.dirname(os.path.abspath(__file__))
INDEX = os.path.join(REPO, "index.html")


def _extract_literal(src, name, open_ch, close_ch):
    """Pull `const NAME = <literal>;` out of src by brace-matching.

    Counting delimiters is enough here: the literals contain no strings with
    unbalanced braces/brackets. Guarded by _selftest below.
    """
    m = re.search(r"const\s+%s\s*=\s*\%s" % (re.escape(name), open_ch), src)
    if not m:
        raise ValueError("could not find `const %s =` in %s" % (name, INDEX))
    start = m.end() - 1
    depth = 0
    for i in range(start, len(src)):
        c = src[i]
        if c == open_ch:
            depth += 1
        elif c == close_ch:
            depth -= 1
            if depth == 0:
                return src[start:i + 1]
    raise ValueError("unbalanced %s%s for %s" % (open_ch, close_ch, name))


def _js_to_json(lit):
    lit = re.sub(r"/\*.*?\*/", "", lit, flags=re.S)      # block comments
    lit = re.sub(r"//[^\n]*", "", lit)                    # line comments
    lit = re.sub(r'(?<![\w"])([A-Za-z_$][\w$]*)\s*:', r'"\1":', lit)  # quote keys
    lit = lit.replace("'", '"')
    lit = re.sub(r",\s*([}\]])", r"\1", lit)              # trailing commas
    return json.loads(lit)


def _read_index():
    with open(INDEX, encoding="utf-8") as f:
        return f.read()


def load():
    """Return (contracts, actuals, cfg) exactly as index.html defines them."""
    src = _read_index()
    contracts = _js_to_json(_extract_literal(src, "CONTRACTS", "[", "]"))
    actuals = _js_to_json(_extract_literal(src, "ACTUALS", "{", "}"))
    cfg = _js_to_json(_extract_literal(src, "CFG", "{", "}"))
    return contracts, actuals, cfg


def quarter_of(iso_date):
    """'2026-08-25' -> '2026-Q3'. Matches index.html's delivery-quarter bucketing."""
    y, m, _ = iso_date.split("-")
    return "%s-Q%d" % (y, (int(m) - 1) // 3 + 1)


def order_book_by_quarter(contracts):
    """Recognised order-book revenue per delivery quarter (K-IFRS point-in-time).

    Mirrors index.html: every contract books 100% in its delivery-deadline quarter.
    """
    out = {}
    for c in contracts:
        q = quarter_of(c["delivery"])
        out[q] = round(out.get(q, 0.0) + c["value"], 6)
    return dict(sorted(out.items()))


def _selftest():
    contracts, actuals, cfg = load()

    assert len(contracts) >= 22, "expected >=22 contracts, got %d" % len(contracts)
    for c in contracts:
        for k in ("date", "party", "seg", "mem", "value", "delivery"):
            assert k in c, "contract missing %r: %r" % (k, c)
        assert c["delivery"] >= c["date"], "delivery precedes order: %r" % c
        assert c["seg"] in ("China-localization", "Korea-HBM"), c["seg"]
        assert c["mem"] in ("DRAM", "NAND", "HBM", "Mixed"), c["mem"]

    for q, a in actuals.items():
        got, want = a["ate"] + a["mai"], a["rev"]
        assert abs(got - want) < 0.02, "%s: ate+mai=%.3f != rev=%.3f" % (q, got, want)

    for k in ("annualRevenueBn", "marketCapBn", "sharesM", "netMargin"):
        assert k in cfg, "CFG missing %r" % k

    total = sum(c["value"] for c in contracts)
    print("contracts     : %d" % len(contracts))
    print("order book    : %.2f bn KRW" % total)
    print("actual qtrs   : %s" % ", ".join(sorted(actuals)))
    print("cfg           : %s" % cfg)
    print("\nrecognised order book by delivery quarter (bn KRW):")
    obq = order_book_by_quarter(contracts)
    for q, v in obq.items():
        print("  %-9s %6.2f" % (q, v))
    assert abs(sum(obq.values()) - total) < 1e-6, "quarter buckets lost value"
    print("\nself-test OK — buckets reconcile to %.2f bn" % total)


if __name__ == "__main__":
    _selftest()
