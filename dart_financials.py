"""
Pull Mirae's full financial statements (BS / IS / CF) from OpenDART and cache to
financials.json — the DCF inputs the dashboard never needed: D&A, capex, working
capital, cash, debt, tax.

    OPENDART_API_KEY=xxx python3 dart_financials.py --years 2022 2023 2024 2025

Cache is gitignored: it's re-fetchable, and the key must not leak via the repo.
Uses stdlib urllib to match update_from_dart.py.
"""

import argparse
import json
import os
import sys
import urllib.parse
import urllib.request

API = "https://opendart.fss.or.kr/api"
CORP_CODE = "00197759"  # Mirae Corporation (025560), cached in .dart_corp_code
REPO = os.path.dirname(os.path.abspath(__file__))
CACHE = os.path.join(REPO, "financials.json")

# reprt_code -> label. 11011 = 사업보고서 (annual, the full-year figure).
REPORTS = {"11013": "Q1", "11012": "H1", "11014": "9M", "11011": "FY"}

# Account names we care about, by statement. DART labels vary between filings, so
# each target lists the aliases seen in practice; first hit wins.
WANTED = {
    "BS": {
        "cash":            ["현금및현금성자산"],
        "short_term_inv":  ["단기금융상품", "당기손익-공정가치측정금융자산"],
        "receivables":     ["매출채권", "매출채권및기타채권", "매출채권 및 기타채권"],
        "inventory":       ["재고자산"],
        "ppe":             ["유형자산"],
        "payables":        ["매입채무", "매입채무및기타채무", "매입채무 및 기타채무"],
        "contract_liab":   ["계약부채", "선수금"],
        "st_debt":         ["단기차입금"],
        "lt_debt":         ["장기차입금", "사채"],
        "equity":          ["자본총계"],
        "total_assets":    ["자산총계"],
    },
    "IS": {
        "revenue":         ["매출액", "수익(매출액)", "영업수익"],
        "cogs":            ["매출원가"],
        "gross_profit":    ["매출총이익"],
        "sga":             ["판매비와관리비"],
        "op_profit":       ["영업이익", "영업이익(손실)"],
        "pretax":          ["법인세비용차감전순이익", "법인세비용차감전순이익(손실)"],
        "tax":             ["법인세비용", "법인세비용(수익)"],
        "net_income":      ["당기순이익", "당기순이익(손실)"],
    },
    "CF": {
        "cfo":             ["영업활동현금흐름", "영업활동으로인한현금흐름"],
        "depreciation":    ["감가상각비"],
        "amortization":    ["무형자산상각비"],
        "capex":           ["유형자산의취득", "유형자산의 취득"],
        "intangible_capex":["무형자산의취득", "무형자산의 취득"],
        "cfi":             ["투자활동현금흐름", "투자활동으로인한현금흐름"],
        "cff":             ["재무활동현금흐름", "재무활동으로인한현금흐름"],
    },
}


def need_key():
    k = os.environ.get("OPENDART_API_KEY")
    if not k:
        sys.exit("ERROR: set OPENDART_API_KEY (free key from https://opendart.fss.or.kr).")
    return k


def _get(url):
    req = urllib.request.Request(url, headers={"User-Agent": "mirae-order-book/1.0"})
    with urllib.request.urlopen(req, timeout=60) as r:
        return json.loads(r.read().decode("utf-8"))


def fetch(key, year, reprt_code, fs_div="CFS"):
    """One period's full account list. fs_div CFS=consolidated, OFS=separate."""
    q = urllib.parse.urlencode({
        "crtfc_key": key, "corp_code": CORP_CODE,
        "bsns_year": str(year), "reprt_code": reprt_code, "fs_div": fs_div,
    })
    d = _get("%s/fnlttSinglAcntAll.json?%s" % (API, q))
    status = d.get("status")
    if status == "013":
        return None                      # no data for this period — normal, skip
    if status != "000":
        raise RuntimeError("DART %s: %s" % (status, d.get("message")))
    return d.get("list", [])


def _num(s):
    if s is None:
        return None
    s = str(s).strip().replace(",", "")
    if s in ("", "-"):
        return None
    neg = s.startswith("(") and s.endswith(")")   # DART parenthesises negatives
    if neg:
        s = s[1:-1]
    try:
        v = float(s)
    except ValueError:
        return None
    return -v if neg else v


def _norm(name):
    return (name or "").replace(" ", "").strip()


def extract(rows):
    """Map DART's raw account list onto WANTED. Returns {sj: {field: KRW bn}}."""
    out = {sj: {} for sj in WANTED}
    unmatched = []
    for r in rows:
        sj = r.get("sj_div")                      # BS / IS / CIS / CF
        if sj not in WANTED:
            continue
        nm = _norm(r.get("account_nm"))
        val = _num(r.get("thstrm_amount"))
        if val is None:
            continue
        hit = False
        for field, aliases in WANTED[sj].items():
            if field in out[sj]:
                continue                          # first match wins
            if any(_norm(a) == nm for a in aliases):
                out[sj][field] = round(val / 1e9, 4)   # KRW -> bn, model's unit
                hit = True
                break
        if not hit:
            unmatched.append((sj, r.get("account_nm")))
    return out, unmatched


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--years", nargs="+", type=int, required=True)
    ap.add_argument("--fs-div", default="CFS", choices=["CFS", "OFS"])
    ap.add_argument("--show-unmatched", action="store_true",
                    help="list DART accounts that matched no WANTED alias")
    args = ap.parse_args()
    key = need_key()

    data, unmatched_all = {}, []
    for y in args.years:
        for code, label in REPORTS.items():
            try:
                rows = fetch(key, y, code, args.fs_div)
            except RuntimeError as e:
                print("  %d %-3s ERROR %s" % (y, label, e))
                continue
            if not rows:
                continue
            vals, unmatched = _extract_report(rows)
            key_ = "%d-%s" % (y, label)
            data[key_] = vals
            unmatched_all += unmatched
            got = {sj: len(v) for sj, v in vals.items()}
            print("  %-8s BS %2d  IS %2d  CF %2d" % (key_, got["BS"], got["IS"], got["CF"]))

    if not data:
        sys.exit("No data returned — check the key and that corp_code %s is right." % CORP_CODE)

    with open(CACHE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2, sort_keys=True)
    print("\nwrote %s (%d periods)" % (CACHE, len(data)))

    if args.show_unmatched:
        print("\nunmatched accounts (candidates for WANTED aliases):")
        for sj, nm in sorted(set(unmatched_all)):
            print("  %-4s %s" % (sj, nm))


def _extract_report(rows):
    return extract(rows)


if __name__ == "__main__":
    main()
