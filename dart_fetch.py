"""
Keyless DART access — filing lists, filing bodies, and section-level fetches.

NO API KEY NEEDED. This is the route that actually works and that every pull in
this repo was built on. WebFetch against dart.fss.or.kr returns a frameset and is
useless; curl/urllib against the viewer.do TOC-offset route returns the real text.

    python3 dart_fetch.py list 20250101 20260716        # all filings, w/ [정정] flags
    python3 dart_fetch.py list 20250101 20260716 단일판매 # filter by report name
    python3 dart_fetch.py toc 20260319001166            # section offsets of a report
    python3 dart_fetch.py body 20260716800558           # full text of a filing
    python3 dart_fetch.py section 20260319001166 24     # one TOC section by eleId

How it works:
  1. dsab007/detailSearch.ax (POST)      -> filing list w/ rcpNo + report name
  2. dsaf001/main.do?rcpNo=<R>           -> viewDoc("<R>","<dcmNo>",...) + makeToc() nodes
  3. report/viewer.do?...&offset=&length -> the actual section text (needs a Referer)

Encoding varies by document type: supply-contract docs are EUC-KR, periodic
reports and 대량보유 docs are UTF-8. decode_any() tries in order.
"""

import re
import sys
import urllib.parse
import urllib.request

CORP = "미래산업"
UA = {"User-Agent": "Mozilla/5.0", "Referer": "https://dart.fss.or.kr/"}


def _get(url, referer=None):
    h = dict(UA)
    if referer:
        h["Referer"] = referer
    return urllib.request.urlopen(urllib.request.Request(url, headers=h), timeout=60).read()


def decode_any(raw):
    for enc in ("utf-8", "euc-kr", "cp949"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    return raw.decode("utf-8", "replace")


def flatten(html):
    """Tags -> | separators. Numbers survive even where labels garble."""
    t = re.sub(r"<(script|style)[^>]*>.*?</\1>", "", html, flags=re.S)
    t = re.sub(r"<[^>]+>", "|", t)
    t = t.replace("&nbsp;", " ").replace("&amp;", "&")
    t = re.sub(r"(\s*\|\s*)+", "|", t)
    return re.sub(r"[ \t]{2,}", " ", t)


def list_filings(start, end, name_filter=None, corp=CORP):
    """[(rcpNo, report_name, is_amendment)] newest first. Paginates until dry."""
    out = {}
    for page in range(1, 8):
        data = urllib.parse.urlencode({"currentPage": page, "maxResults": 100,
                                       "textCrpNm": corp, "startDate": start,
                                       "endDate": end}).encode()
        req = urllib.request.Request("https://dart.fss.or.kr/dsab007/detailSearch.ax",
                                     data=data, headers=UA)
        h = urllib.request.urlopen(req, timeout=60).read().decode("utf-8", "replace")
        found = 0
        for m in re.finditer(r'rcpNo=(\d{14})"[^>]*title="([^"]*?)\s*공시뷰어[^"]*"\s*>(.*?)</a>',
                             h, re.S):
            rcp, nm, inner = m.group(1), m.group(2).strip(), m.group(3)
            out[rcp] = (nm, "[기재정정]" in inner or "[첨부정정]" in inner)
            found += 1
        if not found:
            break
    rows = [(r, n, a) for r, (n, a) in out.items()]
    if name_filter:
        rows = [x for x in rows if name_filter in x[1]]
    return sorted(rows, reverse=True)


def doc_meta(rcp):
    """(dcmNo, dtd) from the first viewDoc(...) call on the report shell page."""
    h = decode_any(_get("https://dart.fss.or.kr/dsaf001/main.do?rcpNo=%s" % rcp))
    m = re.search(r'viewDoc\("(\d+)",\s*"(\d+)",\s*"([^"]*)",\s*"(\d+)",\s*"(\d+)",\s*"([^"]*)"', h)
    if not m:
        return None, None, h
    return m.group(2), m.group(6), h


def toc(rcp):
    """[(title, eleId, offset, length)] parsed from makeToc()'s node assignments."""
    _, _, h = doc_meta(rcp)
    if h is None:
        return []
    nodes = []
    for m in re.finditer(r"node\d+\['text'\]\s*=\s*\"([^\"]*)\";(.*?)(?=var node|treeData\.push|\Z)",
                         h, re.S):
        title, blk = m.group(1).strip(), m.group(2)
        d = {}
        for k in ("eleId", "offset", "length"):
            mm = re.search(k + r"'\]\s*=\s*\"([^\"]*)\"", blk)
            if mm:
                d[k] = mm.group(1)
        if "offset" in d:
            nodes.append((title, d["eleId"], d["offset"], d["length"]))
    return nodes


def section(rcp, ele_id, offset=None, length=None):
    """One TOC section as flattened text. Offsets are looked up if not supplied."""
    dcm, dtd, _ = doc_meta(rcp)
    if not dcm:
        return None
    if offset is None:
        for t, e, o, l in toc(rcp):
            if str(e) == str(ele_id):
                offset, length = o, l
                break
        else:
            return None
    url = ("https://dart.fss.or.kr/report/viewer.do?rcpNo=%s&dcmNo=%s&eleId=%s"
           "&offset=%s&length=%s&dtd=%s" % (rcp, dcm, ele_id, offset, length, dtd or "dart4.xsd"))
    return flatten(decode_any(_get(url, referer="https://dart.fss.or.kr/dsaf001/main.do?rcpNo=%s" % rcp)))


def body(rcp):
    """Whole filing (its first/primary document) as flattened text."""
    dcm, dtd, h = doc_meta(rcp)
    if not dcm:
        return None
    m = re.search(r'viewDoc\("(\d+)",\s*"(\d+)",\s*"([^"]*)",\s*"(\d+)",\s*"(\d+)",\s*"([^"]*)"', h)
    _, _, ele, off, ln, dtd = m.groups()
    return section(rcp, ele, off, ln)


def parse_contract(rcp):
    """Extract a 단일판매·공급계약 filing into a dict. Returns None if not that form."""
    t = body(rcp)
    if not t or "계약금액" not in t:
        return None

    def f(label):
        m = re.search(r"\|%s\|([^|]*)" % re.escape(label), t)
        return m.group(1).strip() if m else None

    def n(s):
        s = re.sub(r"[^\d]", "", s or "")
        return int(s) if s else None

    m = re.search(r"\|5\. 계약기간\|시작일\|([\d-]+)\|종료일\|([\d-]+)", t)
    # NB: DART writes 판매ㆍ공급지역 with U+318D (ㆍ), NOT the middle dot U+00B7 (·).
    # They look near-identical and silently break an exact-match label lookup.
    return {
        "rcp": rcp,
        "amount": n(f("계약금액(원)")) or n(f("확정 계약금액(원)")),
        "party": f("3. 계약상대") or f("계약상대방"),
        "region": f("4. 판매ㆍ공급지역") or f("4. 판매·공급지역"),
        "start": m.group(1) if m else None,
        "delivery": m.group(2) if m else None,     # 계약기간 종료일 = deliver-BY deadline
        "order_date": f("7. 계약(수주)일자") or f("계약(수주)일자"),
        "terms": f("대금지급 조건 등") or f("계약조건") or f("지급조건"),
        # "정정신고(보고)" is unspaced here; some EUC-KR forms render it "정 정 신 고".
        "is_amendment": bool(re.search(r"정\s*정\s*신\s*고", t)) or "정정관련 공시서류" in t
                        or "정정대상 공시서류" in t,
    }


def _cli():
    if len(sys.argv) < 2:
        print(__doc__)
        return
    cmd = sys.argv[1]
    if cmd == "list":
        start, end = sys.argv[2], sys.argv[3]
        filt = sys.argv[4] if len(sys.argv) > 4 else None
        rows = list_filings(start, end, filt)
        print("%d filings" % len(rows))
        for r, nm, am in rows:
            print("  %s  %s%s" % (r, "[정정] " if am else "       ", nm))
    elif cmd == "toc":
        for t, e, o, l in toc(sys.argv[2]):
            print("  ele=%-3s off=%-9s len=%-9s %s" % (e, o, l, t))
    elif cmd == "body":
        print(body(sys.argv[2]))
    elif cmd == "section":
        print(section(sys.argv[2], sys.argv[3]))
    elif cmd == "contract":
        import json
        print(json.dumps(parse_contract(sys.argv[2]), ensure_ascii=False, indent=1))
    else:
        print(__doc__)


if __name__ == "__main__":
    _cli()
