"""Audit Q1/Q2 NCUA CSV exports against CU360 normalization assumptions."""

import csv
import re
from collections import Counter
from pathlib import Path

DEFAULT_FILES = (
    Path.home() / "Downloads" / "AI Labs" / "NCUA" / "Data" / "Q1 2026.csv",
    Path.home() / "Downloads" / "AI Labs" / "NCUA" / "Data" / "Q2 2026.csv",
)
FIELD_ALIASES = {
    "name": ("name", "creditunion", "creditunionname", "cuname", "institution", "institutionname"),
    "state": ("state", "st", "statecode", "stateabbr", "stateabbreviation", "custate"),
    "assetsM": ("assetsm", "assets", "totalassets", "totalassetsm", "assetsmillions", "totalassetsmillions", "acct010"),
    "members": ("members", "totalmembers", "membership", "nummembers", "numberofmembers", "membercount", "acct083"),
    "loansM": ("loansm", "loans", "amountoftotalloansleases", "totalloans", "totalloansandleases", "netloans", "loansandleases", "acct025b"),
    "depositsM": ("depositsm", "deposits", "totaldeposits", "totalsharesamount", "sharesamount", "totalshares", "totalsharesanddeposits", "sharesanddeposits", "shares", "acct018"),
    "equityM": ("equitym", "equity", "networthamount", "networth", "totalnetworth", "totalequity", "acct998"),
    "delinquentAmountM": ("delinquentamountm", "delinquentamount", "totalamountofdelinquentloansleases", "delinquentloansamount"),
    "niM": ("nim", "netincome", "netincomem", "netincomeloss", "ni", "acct661a"),
}
EQUITY_COMPONENT_WORDS = ("amount", "earnings", "reserve", "interest", "gain", "loss", "merger")


def norm_key(value):
    return re.sub(r"[^a-z0-9]", "", str(value or "").lower())


def number(value):
    text = str(value or "").strip()
    negative = text.startswith("-") or (text.startswith("(") and text.endswith(")"))
    cleaned = re.sub(r"[^0-9.eE]", "", text)
    try:
        result = float(cleaned or 0)
    except ValueError:
        return 0.0
    return -abs(result) if negative else result


def unique_headers(headers):
    seen = Counter()
    result = []
    for index, header in enumerate(headers, start=1):
        base = str(header or "").strip() or f"Column {index}"
        seen[base] += 1
        result.append(base if seen[base] == 1 else f"{base} ({seen[base]})")
    return result


def auto_map(headers):
    normalized = [norm_key(header) for header in headers]
    mapping = {}
    used = set()
    for field, aliases in FIELD_ALIASES.items():
        for alias in (norm_key(field), *aliases):
            index = next((i for i, header in enumerate(normalized) if i not in used and header == alias), None)
            if index is not None:
                mapping[field] = headers[index]
                used.add(index)
                break
    noise = re.compile(r"(prior|prev|pct|percent|change|growth|rank|avg|average|yoy|qoq)")
    for field, aliases in FIELD_ALIASES.items():
        if field in mapping:
            continue
        long_aliases = [norm_key(field), *(norm_key(alias) for alias in aliases)]
        for alias in (item for item in long_aliases if len(item) >= 9):
            index = next((i for i, header in enumerate(normalized) if i not in used and alias in header and not noise.search(header)), None)
            if index is not None:
                mapping[field] = headers[index]
                used.add(index)
                break
    return mapping


def load_period(path):
    with path.open(encoding="utf-8-sig", newline="") as source:
        reader = csv.reader(source)
        raw_headers = next(reader)
        raw_rows = [row for row in reader if any(value.strip() for value in row)]
    headers = unique_headers(raw_headers)
    records = [dict(zip(headers, row)) for row in raw_rows]
    return raw_headers, raw_rows, records, auto_map(headers)


def amount(record, column):
    return number(record.get(column, "")) if column else 0.0


def equity_components(record):
    return sum(
        number(value)
        for field, value in record.items()
        if norm_key(field).startswith("equity") and any(word in field.lower() for word in EQUITY_COMPONENT_WORDS)
    )


def summarize(path):
    raw_headers, raw_rows, records, mapping = load_period(path)
    ids = [record.get("CU_NUMBER", "").strip() for record in records if record.get("CU_NUMBER", "").strip()]
    duplicates = {header: count for header, count in Counter(raw_headers).items() if count > 1}
    duplicate_conflicts = {}
    for header in duplicates:
        indices = [index for index, item in enumerate(raw_headers) if item == header]
        if header:
            duplicate_conflicts[header] = sum(
                len({row[index] if index < len(row) else "" for index in indices}) > 1
                for row in raw_rows
            )
    required_source_fields = (
        "Total_Assets", "Amount_of_Total_Loans_Leases", "Total_Shares_Amount",
        "Total_Members", "Total_Amount_of_Delinquent_Loans_Leases", "Gross_Income",
    )
    missing = {
        field: sum(not record.get(field, "").strip() for record in records)
        for field in required_source_fields
    }
    denominators = {
        "assets": mapping.get("assetsM"),
        "loans": mapping.get("loansM"),
        "shares": mapping.get("depositsM"),
    }
    zero_denominators = {
        label: sum(amount(record, field) == 0 for record in records) if field else len(records)
        for label, field in denominators.items()
    }
    ratios = {"net_worth_pct": [], "delinquency_pct": [], "loan_to_share_pct": [], "unannualized_roa_pct": []}
    for record in records:
        assets = amount(record, mapping.get("assetsM"))
        loans = amount(record, mapping.get("loansM"))
        shares = amount(record, mapping.get("depositsM"))
        equity = equity_components(record) or amount(record, mapping.get("equityM"))
        delinquent = number(record.get("Total_Amount_of_Delinquent_Loans_Leases", ""))
        net_income = amount(record, mapping.get("niM"))
        ratios["net_worth_pct"].append(pct(equity, assets))
        ratios["delinquency_pct"].append(pct(delinquent, loans))
        ratios["loan_to_share_pct"].append(pct(loans, shares))
        ratios["unannualized_roa_pct"].append(pct(net_income, assets))
    affinity = next((
        record for record in records
        if record.get("CU_Name", "").strip().casefold() == "affinity"
        and record.get("STATE", "").strip().upper() == "NJ"
    ), None)
    result = {
        "path": path,
        "raw_headers": raw_headers,
        "raw_rows": len(raw_rows),
        "records": records,
        "mapping": mapping,
        "id_count": len(ids),
        "unique_ids": len(set(ids)),
        "duplicate_ids": len(ids) - len(set(ids)),
        "blank_headers": sum(not header.strip() for header in raw_headers),
        "duplicate_headers": duplicates,
        "duplicate_header_conflicts": duplicate_conflicts,
        "missing_required": missing,
        "zero_denominators": zero_denominators,
        "ratio_values": ratios,
        "affinity": affinity,
    }
    return result


def pct(numerator, denominator):
    return numerator / denominator * 100 if denominator else None


def distribution(values):
    ordered = sorted(value for value in values if value is not None)
    if not ordered:
        return None
    return {
        "min": ordered[0],
        "p50": ordered[(len(ordered) - 1) // 2],
        "p95": ordered[int((len(ordered) - 1) * 0.95)],
        "max": ordered[-1],
    }


def display(value):
    return "N/A" if value is None else f"{value:,.4f}%"


def main():
    summaries = [summarize(path) for path in DEFAULT_FILES]
    for summary in summaries:
        print(f"\n{summary['path'].name}: source rows={summary['raw_rows']}; unique IDs={summary['unique_ids']}; duplicate IDs={summary['duplicate_ids']}")
        print(f"  raw columns={len(summary['raw_headers'])}; blank headers={summary['blank_headers']}; duplicate headers={summary['duplicate_headers']}")
        print(f"  duplicate-header rows with conflicting values={summary['duplicate_header_conflicts']}")
        print(f"  mapped source fields={summary['mapping']}")
        print(f"  missing required values={summary['missing_required']}")
        print(f"  zero denominators={summary['zero_denominators']}")
        print(f"  source ratio distributions (ROA is unannualized)={ {name: distribution(values) for name, values in summary['ratio_values'].items()} }")
        record = summary["affinity"]
        if not record:
            print("  Affinity, NJ: not found")
            continue
        mapping = summary["mapping"]
        assets = amount(record, mapping.get("assetsM"))
        loans = amount(record, mapping.get("loansM"))
        shares = amount(record, mapping.get("depositsM"))
        equity = equity_components(record)
        delinquent = number(record.get("Total_Amount_of_Delinquent_Loans_Leases", ""))
        print(f"  Affinity NJ CU_NUMBER={record.get('CU_NUMBER')}; assets={assets:,.0f}; loans={loans:,.0f}; shares={shares:,.0f}; source NET_WORTH_AMOUNT={record.get('NET_WORTH_AMOUNT')}; equity components={equity:,.0f}")
        print(f"  independently recalculated: net worth={display(pct(equity, assets))}; delinquency={display(pct(delinquent, loans))}; loan-to-share={display(pct(loans, shares))}")
        print(f"  current alias result: deposits maps to {mapping.get('depositsM')!r}; capital source maps to {mapping.get('equityM')!r}")

    if len(summaries) == 2:
        q1, q2 = summaries
        q1_by_id = {row.get("CU_NUMBER", "").strip(): row for row in q1["records"] if row.get("CU_NUMBER", "").strip()}
        q2_by_id = {row.get("CU_NUMBER", "").strip(): row for row in q2["records"] if row.get("CU_NUMBER", "").strip()}
        common_ids = q1_by_id.keys() & q2_by_id.keys()
        print(f"\nCross-period CU_NUMBER continuity: common={len(common_ids)}; Q1-only={len(q1_by_id.keys() - q2_by_id.keys())}; Q2-only={len(q2_by_id.keys() - q1_by_id.keys())}")
        affinity_id = "857"
        if affinity_id in q1_by_id and affinity_id in q2_by_id:
            growth_fields = {
                "assets": "Total_Assets",
                "loans": "Amount_of_Total_Loans_Leases",
                "deposits": "Total_Shares_Amount",
                "members": "Total_Members",
            }
            for label, field in growth_fields.items():
                prior = number(q1_by_id[affinity_id].get(field, ""))
                current = number(q2_by_id[affinity_id].get(field, ""))
                change = (current / prior - 1) * 100 if prior else None
                print(f"Affinity NJ {label} Q1->Q2 growth={display(change)} (Q1={prior:,.0f}; Q2={current:,.0f})")


if __name__ == "__main__":
    main()
