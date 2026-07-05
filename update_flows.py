#!/usr/bin/env python3
"""
update_flows.py — pull Mirae Corp (025560) SECONDARY-MARKET investor flows
(on-exchange net buying by 기관/외국인 + foreign ownership %) from Naver Finance,
so the "Secondary-market equity flows" panel in index.html can be refreshed.

This is DISTINCT from update_from_dart.py:
  - update_from_dart.py  -> primary/contract & ownership filings (DART)
  - update_flows.py      -> daily secondary-market trading flow (Naver/KRX)

WHY NAVER (not KRX直): KRX's data.krx.co.kr investor-flow endpoints reject
non-browser requests (return "LOGOUT"); Naver's item/frgn table redistributes the
same KRX data and is fetchable. Naver reports 기관 & 외국인 net share flow + foreign
holding shares/ratio; 개인(retail) is the residual and not separately reported.

USAGE
  python3 update_flows.py                 # ~7 pages (~140 trading days)
  python3 update_flows.py --pages 12      # deeper history
Output: the JS `const FLOWS = [ ... ];` block to paste into index.html.
Row = [date, close(KRW), instNetShares, frgnNetShares, frgnHoldShares, frgnPct]
"""
import re, sys, argparse, urllib.request

CODE = "025560"
UA = ("Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/125 Safari/537.36")


def fetch_page(page):
    url = f"https://finance.naver.com/item/frgn.naver?code={CODE}&page={page}"
    req = urllib.request.Request(url, headers={
        "User-Agent": UA,
        "Referer": f"https://finance.naver.com/item/main.naver?code={CODE}"})
    return urllib.request.urlopen(req, timeout=30).read().decode("euc-kr", "ignore")


def parse(html):
    out = []
    for r in re.findall(r"<tr[^>]*>(.*?)</tr>", html, re.S):
        c = [re.sub(r"<[^>]+>", "", x).replace("&nbsp;", " ").strip()
             for x in re.findall(r"<td[^>]*>(.*?)</td>", r, re.S)]
        if c and re.match(r"20\d\d\.\d\d\.\d\d", c[0]):
            num = lambda s: int(re.sub(r"[^\d-]", "", s.replace("+", "")) or 0)
            out.append([
                c[0].replace(".", "-"),          # date
                num(c[1]),                        # close
                num(c[5]),                        # 기관 net shares
                num(c[6]),                        # 외국인 net shares
                num(c[7]),                        # 외국인 holding shares
                float(re.sub(r"[^\d.]", "", c[8]) or 0),  # 외국인 %
            ])
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--pages", type=int, default=7)
    args = ap.parse_args()

    rows = {}
    for p in range(1, args.pages + 1):
        for row in parse(fetch_page(p)):
            rows[row[0]] = row
    data = sorted(rows.values(), key=lambda r: r[0])
    print(f"# {len(data)} trading days  {data[0][0]} -> {data[-1][0]}", file=sys.stderr)

    print("const FLOWS = [")
    for r in data:
        print(f'  ["{r[0]}",{r[1]},{r[2]},{r[3]},{r[4]},{r[5]}],')
    print("];")


if __name__ == "__main__":
    main()
