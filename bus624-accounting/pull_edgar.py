#!/usr/bin/env python3
"""
BUS 624 - Comprehensive Financial Statement Analysis
Step 1 of 2: pull raw XBRL facts from SEC EDGAR for Walmart + competitors.

WHY THIS EXISTS
    Prof. Antal requires that AI-assisted numbers be traced back to the original
    filing and verified by hand. So this script does NOT compute anything. It only
    fetches, and it stamps every single value with the accession number, form type
    and filing date it came from. That stamp is what makes hand-verification a
    5-second lookup instead of a scavenger hunt.

USAGE
    pip install requests
    python3 pull_edgar.py                    # writes edgar_facts.csv, 4 years
    python3 pull_edgar.py --years 5          # pull more history

    Then run build_workbook.py to turn edgar_facts.csv into the Excel deliverable.

NOTE ON THE USER-AGENT
    The SEC requires a real contact email in the User-Agent header or it will
    throttle/deny you. Put your own address in CONTACT below.
"""
import argparse, csv, json, sys, time
from collections import defaultdict

try:
    import requests
except ImportError:
    sys.exit("Need requests:  pip install requests")

CONTACT = "Micah Kosasa mkosasa@gmail.com"   # <-- SEC requires a real contact
HEADERS = {"User-Agent": CONTACT, "Accept-Encoding": "gzip, deflate"}

# --- Companies -------------------------------------------------------------
# The script prints the entityName EDGAR returns for each CIK so you can confirm
# the CIK actually resolves to the company you think it does. Verify on first run.
COMPANIES = {
    "WMT":  {"cik": 104169, "name": "Walmart Inc.",        "fye": "01-31", "role": "subject"},
    "COST": {"cik": 909832, "name": "Costco Wholesale",    "fye": "08-31", "role": "competitor"},
    "TGT":  {"cik":  27419, "name": "Target Corp",         "fye": "01-31", "role": "competitor"},
    "KR":   {"cik":  56873, "name": "Kroger Co",           "fye": "01-31", "role": "competitor"},
}

# --- Concept -> candidate us-gaap tags -------------------------------------
# Filers tag the same economic concept differently. We try each candidate in
# order and keep the first that returns data, recording WHICH tag matched.
# If a concept comes back empty for a company, that is a finding, not a bug:
# it usually means the company reports it on a different line. Check the 10-K.
CONCEPTS = {
    # ---- Income statement ----
    "Revenue": ["RevenueFromContractWithCustomerExcludingAssessedTax", "Revenues",
                "SalesRevenueNet", "RevenueFromContractWithCustomerIncludingAssessedTax"],
    "COGS": ["CostOfGoodsAndServicesSold", "CostOfRevenue", "CostOfGoodsSold"],
    "GrossProfit": ["GrossProfit"],
    "SGA": ["SellingGeneralAndAdministrativeExpense",
            "GeneralAndAdministrativeExpense", "OperatingCostsAndExpenses"],
    "OperatingIncome": ["OperatingIncomeLoss"],
    "InterestExpense": ["InterestExpense", "InterestExpenseNonoperating",
                        "InterestIncomeExpenseNet", "InterestExpenseDebt"],
    "PretaxIncome": ["IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest",
                     "IncomeLossFromContinuingOperationsBeforeIncomeTaxesMinorityInterestAndIncomeLossFromEquityMethodInvestments"],
    "TaxExpense": ["IncomeTaxExpenseBenefit"],
    "NetIncome": ["NetIncomeLoss", "ProfitLoss"],
    "EPSDiluted": ["EarningsPerShareDiluted"],
    "EPSBasic": ["EarningsPerShareBasic"],
    "SharesDiluted": ["WeightedAverageNumberOfDilutedSharesOutstanding"],

    # ---- Balance sheet ----
    "Cash": ["CashAndCashEquivalentsAtCarryingValue", "CashCashEquivalentsRestrictedCashAndRestrictedCashEquivalents"],
    "Receivables": ["ReceivablesNetCurrent", "AccountsReceivableNetCurrent"],
    "Inventory": ["InventoryNet"],
    "CurrentAssets": ["AssetsCurrent"],
    "PPENet": ["PropertyPlantAndEquipmentNet"],
    "Goodwill": ["Goodwill"],
    "TotalAssets": ["Assets"],
    "AccountsPayable": ["AccountsPayableCurrent", "AccountsPayableTradeCurrent"],
    "CurrentLiabilities": ["LiabilitiesCurrent"],
    "LongTermDebt": ["LongTermDebtNoncurrent", "LongTermDebt"],
    "LongTermDebtCurrent": ["LongTermDebtCurrent"],
    "OperatingLeaseLiabNoncurrent": ["OperatingLeaseLiabilityNoncurrent"],
    "OperatingLeaseLiabCurrent": ["OperatingLeaseLiabilityCurrent"],
    "TotalLiabilities": ["Liabilities"],
    "Equity": ["StockholdersEquity"],
    "EquityInclNCI": ["StockholdersEquityIncludingPortionAttributableToNoncontrollingInterest"],

    # ---- Cash flow ----
    "CFO": ["NetCashProvidedByUsedInOperatingActivities",
            "NetCashProvidedByUsedInOperatingActivitiesContinuingOperations"],
    "CFI": ["NetCashProvidedByUsedInInvestingActivities"],
    "CFF": ["NetCashProvidedByUsedInFinancingActivities"],
    "CapEx": ["PaymentsToAcquirePropertyPlantAndEquipment",
              "PaymentsToAcquireProductiveAssets"],
    "DepAmort": ["DepreciationDepletionAndAmortization", "DepreciationAmortizationAndAccretionNet"],
    "DividendsPaid": ["PaymentsOfDividendsCommonStock", "PaymentsOfDividends"],
    "Buybacks": ["PaymentsForRepurchaseOfCommonStock"],
}

BASE = "https://data.sec.gov/api/xbrl/companyfacts/CIK{:010d}.json"


def fetch_facts(cik, retries=4):
    """Fetch the full companyfacts blob, with backoff on transient failures."""
    url = BASE.format(cik)
    delay = 2
    for attempt in range(retries):
        try:
            r = requests.get(url, headers=HEADERS, timeout=60)
            if r.status_code == 200:
                return r.json()
            if r.status_code in (429, 503):
                time.sleep(delay); delay *= 2; continue
            r.raise_for_status()
        except requests.RequestException as e:
            if attempt == retries - 1:
                raise
            time.sleep(delay); delay *= 2
    raise RuntimeError(f"could not fetch CIK {cik}")


def annual_rows(facts, concept, candidates, ticker, n_years):
    """
    Return annual (10-K) observations for the first candidate tag that has data.

    Selection rules that matter for correctness:
      * form must be 10-K (not 10-Q) so we get audited annual figures
      * for flow concepts we require a ~365 day window (start->end), which drops
        the year-to-date quarterly stubs EDGAR also stores under fy/FY
      * when the same fiscal year appears more than once (original filing plus
        later comparatives), we keep the ORIGINAL - the earliest filed - because
        that is the number printed in that year's own 10-K, which is the one the
        group will be eyeballing when they verify by hand.
    """
    usgaap = facts.get("facts", {}).get("us-gaap", {})
    for tag in candidates:
        if tag not in usgaap:
            continue
        out = {}
        for unit, obs in usgaap[tag]["units"].items():
            for o in obs:
                if o.get("form") != "10-K" or o.get("fy") is None:
                    continue
                start, end = o.get("start"), o.get("end")
                if start:  # flow concept - enforce an annual window
                    from datetime import date
                    y0, m0, d0 = map(int, start.split("-"))
                    y1, m1, d1 = map(int, end.split("-"))
                    if not (330 <= (date(y1, m1, d1) - date(y0, m0, d0)).days <= 400):
                        continue
                key = (o["fy"], o.get("frame", ""))
                fy = o["fy"]
                cand = {
                    "ticker": ticker, "concept": concept, "us_gaap_tag": tag,
                    "unit": unit, "fiscal_year": fy, "period_start": start or "",
                    "period_end": end, "value": o["val"],
                    "form": o["form"], "filed": o.get("filed", ""),
                    "accession": o.get("accn", ""),
                }
                prev = out.get(fy)
                # keep earliest-filed observation for each fiscal year
                if prev is None or cand["filed"] < prev["filed"]:
                    out[fy] = cand
        if out:
            years = sorted(out, reverse=True)[:n_years]
            return [out[y] for y in sorted(years)]
    return []


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--years", type=int, default=4,
                    help="fiscal years to pull. Default 4, not 3: the assignment wants "
                         "3 years of RATIOS, but average-balance ratios (ROA, ROE, all "
                         "turnovers) need an opening balance from the year before the "
                         "earliest ratio year.")
    ap.add_argument("--out", default="edgar_facts.csv")
    args = ap.parse_args()

    rows, missing = [], defaultdict(list)

    for ticker, meta in COMPANIES.items():
        print(f"Fetching {ticker} (CIK {meta['cik']}) ...", flush=True)
        facts = fetch_facts(meta["cik"])
        entity = facts.get("entityName", "?")
        # Sanity check: does this CIK actually resolve to who we think?
        flag = "" if meta["name"].split()[0].lower() in entity.lower() else "  <-- CHECK THIS"
        print(f"    EDGAR says: {entity}{flag}")

        for concept, candidates in CONCEPTS.items():
            got = annual_rows(facts, concept, candidates, ticker, args.years)
            if got:
                rows.extend(got)
            else:
                missing[ticker].append(concept)
        time.sleep(0.2)   # be polite to SEC

    cols = ["ticker", "concept", "us_gaap_tag", "unit", "fiscal_year",
            "period_start", "period_end", "value", "form", "filed", "accession"]
    with open(args.out, "w", newline="") as f:
        w = csv.DictWriter(f, fieldnames=cols)
        w.writeheader()
        w.writerows(rows)

    print(f"\nWrote {len(rows)} facts to {args.out}")
    if missing:
        print("\nNOT FOUND - these need a manual look in the 10-K.")
        print("This is normal. It usually means the company uses a different line item,")
        print("not that the data is absent. Note what you find in the workbook's Notes column.")
        for t, cs in missing.items():
            print(f"  {t}: {', '.join(cs)}")


if __name__ == "__main__":
    main()
