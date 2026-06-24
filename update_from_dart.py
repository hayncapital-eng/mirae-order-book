#!/usr/bin/env python3
"""
update_from_dart.py — pull Mirae Corp (025560) single-supply-contract filings
straight from DART via the OpenDART API, so the order-book model can be updated
without manual screenshotting.

WHAT IT DOES
  1. Resolves Mirae's 8-digit DART corp_code (cached to .dart_corp_code).
  2. Lists exchange disclosures (pblntf_ty=I) over a date range and keeps the
     single-supply-contract ones (단일판매ㆍ공급계약체결) + their progress/closeout
     updates (진행상황 / 정정).
  3. For each, downloads the filing document and prints the contract fields:
     counterparty, amount, contract period (start + END = delivery deadline),
     payment terms — i.e. everything the dashboard needs, including the delivery
     date that drives the revenue-recognition quarter.
  4. Flags filings whose delivery deadline has already passed → these are the
     ones to cross-check against reported quarterly revenue (validation step).

SETUP (one-time)
  - Get a FREE OpenDART API key: https://opendart.fss.or.kr  (Sign up → 인증키 신청/관리)
  - export OPENDART_API_KEY=xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx

USAGE
  python3 update_from_dart.py                 # last 180 days
  python3 update_from_dart.py --since 20250101 # explicit start date
  python3 update_from_dart.py --rcept 20260601800123   # one specific filing

Output is human-readable AND a JSON block (--json) that maps onto the CONTRACTS
list in index.html.
"""
import os, sys, io, re, json, zipfile, argparse, datetime, urllib.request, urllib.parse

API = "https://opendart.fss.or.kr/api"
KEY = os.environ.get("OPENDART_API_KEY", "").strip()
STOCK_CODE = "025560"
CACHE = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".dart_corp_code")

# DART single-supply-contract field labels -> our keys
LABELS = {
    "amount":   ["계약금액", "공급계약금액", "계약금액(원)"],
    "party":    ["계약상대", "계약상대방", "판매ㆍ공급계약 상대방", "매출처"],
    "start":    ["계약시작일", "계약기간 시작일", "시작일"],
    "end":      ["계약종료일", "계약기간 종료일", "종료일", "납기"],
    "terms":    ["주요계약조건", "대금지급", "결제조건", "주요 계약조건"],
    "product":  ["판매ㆍ공급계약 내용", "체결계약명", "계약내용", "판매ㆍ공급계약 구분"],
}


def _get(url):
    req = urllib.request.Request(url, headers={"User-Agent": "mirae-order-book/1.0"})
    with urllib.request.urlopen(req, timeout=60) as r:
        return r.read()


def need_key():
    if not KEY:
        sys.exit("ERROR: set OPENDART_API_KEY (free key from https://opendart.fss.or.kr).")


def resolve_corp_code():
    """Mirae's DART corp_code, cached. corpCode.xml is a zip of every listed company."""
    if os.path.exists(CACHE):
        return open(CACHE).read().strip()
    need_key()
    raw = _get(f"{API}/corpCode.xml?crtfc_key={KEY}")
    zf = zipfile.ZipFile(io.BytesIO(raw))
    xml = zf.read(zf.namelist()[0]).decode("utf-8", "ignore")
    # entries look like <list><corp_code>..</corp_code><corp_name>..</corp_name><stock_code>025560</stock_code>..</list>
    for m in re.finditer(r"<list>(.*?)</list>", xml, re.S):
        block = m.group(1)
        if f"<stock_code>{STOCK_CODE}</stock_code>" in block:
            cc = re.search(r"<corp_code>(\d+)</corp_code>", block).group(1)
            open(CACHE, "w").write(cc)
            print(f"[resolved corp_code {cc} for {STOCK_CODE}]", file=sys.stderr)
            return cc
    sys.exit(f"Could not find corp_code for stock {STOCK_CODE} in corpCode.xml")


def list_filings(corp_code, since, until):
    """All exchange disclosures (pblntf_ty=I) in [since,until], paginated."""
    need_key()
    out, page = [], 1
    while True:
        q = urllib.parse.urlencode({
            "crtfc_key": KEY, "corp_code": corp_code, "bgn_de": since, "end_de": until,
            "pblntf_ty": "I", "page_no": page, "page_count": 100,
        })
        data = json.loads(_get(f"{API}/list.json?{q}"))
        if data.get("status") != "000":
            if data.get("status") == "013":  # no data
                break
            sys.exit(f"list.json error {data.get('status')}: {data.get('message')}")
        out += data.get("list", [])
        if page >= int(data.get("total_page", 1)):
            break
        page += 1
    return out


def fetch_doc_text(rcept_no):
    """document.xml returns a zip of the filing; return its concatenated text."""
    need_key()
    raw = _get(f"{API}/document.xml?crtfc_key={KEY}&rcept_no={rcept_no}")
    try:
        zf = zipfile.ZipFile(io.BytesIO(raw))
    except zipfile.BadZipFile:
        return raw.decode("utf-8", "ignore")  # error XML, return as-is
    text = []
    for name in zf.namelist():
        body = zf.read(name).decode("utf-8", "ignore")
        body = re.sub(r"<[^>]+>", " ", body)        # strip tags
        body = re.sub(r"[ \t ]+", " ", body)
        text.append(body)
    return "\n".join(text)


def find_after(text, labels, maxlen=80):
    """Grab the value appearing just after any of the given Korean labels."""
    for lab in labels:
        m = re.search(re.escape(lab) + r"[^0-9A-Za-z가-힣]{0,6}([^\n]{1,%d})" % maxlen, text)
        if m:
            return m.group(1).replace("\r", "").strip(" :：\t-")
    return None


def norm_date(s):
    if not s:
        return None
    m = re.search(r"(20\d\d)[.\-/년 ]+(\d{1,2})[.\-/월 ]+(\d{1,2})", s)
    return f"{m.group(1)}-{int(m.group(2)):02d}-{int(m.group(3)):02d}" if m else None


def is_date(s):
    return bool(s) and bool(re.fullmatch(r"\d{4}-\d{2}-\d{2}", s))


def parse_contract(text):
    rec = {}
    # contract amount: first comma-grouped number right after 계약금액(원)
    m = re.search(r"계약금액[^\n]{0,12}\s*([\d]{1,3}(?:,\d{3})+)", text)
    if m:
        rec["amount_won"] = m.group(1)
        rec["amount_bn"] = round(int(m.group(1).replace(",", "")) / 1e9, 2)
    rec["party"]      = find_after(text, ["계약상대"])
    rec["start"]      = norm_date(find_after(text, ["시작일"]))
    rec["end"]        = norm_date(find_after(text, ["종료일"]))
    rec["terms"]      = find_after(text, ["대금지급 조건 등", "조건 등", "결제조건"])
    rec["order_date"] = norm_date(find_after(text, ["계약(수주)일자", "수주)일자"]))
    rec["product"]    = find_after(text, ["판매ㆍ공급계약 내용", "체결계약명", "계약내용"])
    return rec


def quarter(d):
    if not is_date(d):
        return "?"
    y, m, _ = d.split("-")
    return f"{y}-Q{(int(m)-1)//3 + 1}"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--since", default=(datetime.date.today() - datetime.timedelta(days=180)).strftime("%Y%m%d"))
    ap.add_argument("--until", default=datetime.date.today().strftime("%Y%m%d"))
    ap.add_argument("--rcept", help="inspect a single filing by receipt no")
    ap.add_argument("--json", action="store_true", help="also emit JSON for index.html")
    args = ap.parse_args()

    if args.rcept:
        rec = parse_contract(fetch_doc_text(args.rcept))
        print(json.dumps(rec, ensure_ascii=False, indent=2))
        return

    corp = resolve_corp_code()
    filings = list_filings(corp, args.since, args.until)
    contracts = [f for f in filings if "공급계약" in f.get("report_nm", "")]
    print(f"# {len(contracts)} supply-contract filings for Mirae {STOCK_CODE} "
          f"({args.since}–{args.until})\n")

    today = datetime.date.today().strftime("%Y-%m-%d")
    results = []
    for f in contracts:
        rec = parse_contract(fetch_doc_text(f["rcept_no"]))
        rec["rcept_no"], rec["report_nm"] = f["rcept_no"], f["report_nm"]
        rec["rcept_dt"] = f"{f['rcept_dt'][:4]}-{f['rcept_dt'][4:6]}-{f['rcept_dt'][6:]}"
        delivered = is_date(rec.get("end")) and rec["end"] <= today
        rec["delivery_passed"] = bool(delivered)
        rec["delivery_q"] = quarter(rec.get("end"))
        results.append(rec)
        print(f"- order {rec.get('order_date') or rec['rcept_dt']}  {f['report_nm']}  (rcept {f['rcept_no']})")
        print(f"    party:   {rec.get('party')}")
        print(f"    amount:  {rec.get('amount_bn')} bn KRW")
        print(f"    period:  {rec.get('start')} -> {rec.get('end')}"
              + (f"   [delivery quarter {rec['delivery_q']}]" if is_date(rec.get('end')) else "")
              + ("   << deadline passed: validate vs reported revenue" if delivered else ""))
        print(f"    terms:   {rec.get('terms')}")
        print(f"    product: {rec.get('product')}\n")

    if args.json:
        print("\n# ---- JSON (maps to CONTRACTS in index.html) ----")
        print(json.dumps(results, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
