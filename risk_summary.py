"""
risk_summary.py

Reads the risk register and produces a summary by risk rating band, using a
standard 5x5 likelihood/impact matrix (score = likelihood x impact, 1-25).

Rating bands:
    1-4   Low
    5-9   Medium
    10-14 High
    15-25 Critical

Usage:
    python3 risk_summary.py
"""

import csv


def rating(score):
    score = int(score)
    if score >= 15:
        return "Critical"
    elif score >= 10:
        return "High"
    elif score >= 5:
        return "Medium"
    else:
        return "Low"


def load_register(path="risk_register.csv"):
    with open(path, newline="") as f:
        return list(csv.DictReader(f))


def main():
    rows = load_register()

    print(f"Total risks tracked: {len(rows)}\n")

    print("=== Inherent risk by rating band ===")
    inherent_bands = {}
    for r in rows:
        band = rating(r["inherent_risk_score"])
        inherent_bands.setdefault(band, []).append(r["risk_id"])
    for band in ["Critical", "High", "Medium", "Low"]:
        ids = inherent_bands.get(band, [])
        print(f"{band:9s}: {len(ids):2d}  {', '.join(ids)}")

    print("\n=== Residual risk by rating band (after mitigation) ===")
    residual_bands = {}
    for r in rows:
        band = rating(r["residual_risk_score"])
        residual_bands.setdefault(band, []).append(r["risk_id"])
    for band in ["Critical", "High", "Medium", "Low"]:
        ids = residual_bands.get(band, [])
        print(f"{band:9s}: {len(ids):2d}  {', '.join(ids)}")

    print("\n=== Open or in-progress Critical/High risks ===")
    for r in rows:
        if r["status"] in ("Open", "In Progress") and rating(r["inherent_risk_score"]) in ("Critical", "High"):
            print(f"  {r['risk_id']}: {r['title']} (owner: {r['risk_owner']}, target: {r['target_date']})")


if __name__ == "__main__":
    main()
