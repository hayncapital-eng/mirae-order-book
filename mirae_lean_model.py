"""
mirae_lean_model.py  --  ADDRESSABLE-CAPACITY signal (researched).

ONE question: is MORE Mirae-applicable fab capacity coming than what drove the recent surge?
We model ONLY capacity ADDED per year (kwpm). No capture rate, no revenue translation.

Two series:
  ADDRESSABLE = Mirae-customer fabs x commodity/package-testable capacity (ex-HBM).
                Mirae customers: CXMT, YMTC, SK hynix, (Samsung spec.). NOT Micron.
                HBM-bound DRAM bypasses Mirae's package handlers (wafer/bare-die test).
  INDUSTRY    = ALL makers (incl Micron + non-confirmed) and ALL product (incl HBM).
                The gap (industry - addressable) = HBM-bound + Micron + non-Mirae + Korea-NAND cuts.

All figures L-confidence rounded kwpm; sources in comments. Edit as data arrives.
Sources: SemiAnalysis/Omdia/Digitimes (CXMT 100->200->~300->~350); TrendForce/TheElec
(SK hynix M15X 40->80, Yongin Y1 60kwpm/cleanroom, 1st install Feb'27, HBM-skewed);
TrendForce (HBM ~23% of DRAM wafers 2026); TrendForce (YMTC Phase III +50kwpm by 2027);
TrendForce (Samsung 1c 200kwpm by end-2026, P4 done Apr'27, HBM4-focused).
"""

# --- ADDRESSABLE kwpm added/yr (Mirae customer x commodity ex-HBM) ---
ADDR = {
    2024: {"CXMT": 95, "YMTC": 25, "SK hynix": 15},                              # CXMT 100->~195
    2025: {"CXMT": 85, "YMTC": 35, "SK hynix": 15, "Samsung": 8},                # CXMT ->~280
    2026: {"CXMT": 76, "YMTC": 40, "SK hynix": 22, "Samsung": 15, "Korea NAND": -20},  # CXMT ->~350
    # UPCOMING (pipeline) -- SK hynix cut vs first draft: Yongin Y1 1st cleanroom installs
    # Feb'27 & is HBM-skewed, so 2027 commodity is light; bulk shifts to 2028.
    2027: {"CXMT": 83, "SK hynix M15X": 25, "SK hynix Yongin Y1": 32, "YMTC": 45},     # was SK 127
    2028: {"CXMT": 83, "SK hynix Yongin Y1": 70, "YMTC": 45, "Samsung P5 (spec.)": 30},
}

# --- TOTAL INDUSTRY kwpm added/yr (all makers, DRAM+NAND, incl HBM) ---
# DRAM: Samsung 7.47M->8.18M wafers/yr ('24->'25, +59 kwpm); SK 5.12M->6.39M (+107); CXMT +85;
# Micron +30. 2H26 Samsung & SK both ramp DRAM hard. NAND: YMTC +, Korea NAND -, others flat.
INDUSTRY = {2024: 270, 2025: 320, 2026: 360, 2027: 440, 2028: 440}

RECENT_YEARS = [2024, 2025, 2026]
PIPELINE_YEARS = [2027, 2028]

total = lambda m: sum(m.values())
addr = {y: total(ADDR[y]) for y in ADDR}


def show(title, years):
    print(f"{title}:")
    for y in years:
        drv = ", ".join(f"{k} {v:+d}" if v < 0 else f"{k} {v}" for k, v in ADDR[y].items())
        print(f"  {y}: addressable {addr[y]:>4.0f} | industry {INDUSTRY[y]:>4.0f} "
              f"| Mirae share {addr[y]/INDUSTRY[y]:.0%}  ({drv})")
    a = sum(addr[y] for y in years) / len(years)
    print(f"  avg addressable: {a:.0f} kwpm/yr\n")
    return a


if __name__ == "__main__":
    print("=== Mirae addressable vs industry capacity additions (kwpm/yr) ===\n")
    ra = show("RECENT (drove the surge)", RECENT_YEARS)
    ua = show("UPCOMING (pipeline)", PIPELINE_YEARS)
    print(f">>> Addressable pipeline is {ua/ra - 1:+.0%} vs recent ({ua:.0f} vs {ra:.0f} kwpm/yr).")
    print(f">>> Industry pipeline avg {sum(INDUSTRY[y] for y in PIPELINE_YEARS)/2:.0f} vs recent "
          f"{sum(INDUSTRY[y] for y in RECENT_YEARS)/3:.0f} kwpm/yr.")
    print(f">>> Mirae-addressable share of industry: "
          f"{sum(addr[y] for y in RECENT_YEARS)/sum(INDUSTRY[y] for y in RECENT_YEARS):.0%} recent "
          f"-> {sum(addr[y] for y in PIPELINE_YEARS)/sum(INDUSTRY[y] for y in PIPELINE_YEARS):.0%} pipeline")
    print("    (share dips as HBM + Micron take more of industry growth).")
