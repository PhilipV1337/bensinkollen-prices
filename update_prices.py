#!/usr/bin/env python3
"""Import an EU Weekly Oil Bulletin workbook and matching ECB historical EUR/SEK rate.
Requires: python3 -m pip install openpyxl
Run: python3 scripts/update_prices.py /path/to/weekly-file.xlsx
The workbook is downloaded manually from the European Commission. This is NOT a scheduled updater.
"""
import csv, io, json, sys, urllib.request
from datetime import datetime
from pathlib import Path
from openpyxl import load_workbook

ECB_CSV = "https://www.ecb.europa.eu/stats/eurofxref/eurofxref-hist.csv"
DEST = Path(__file__).resolve().parents[1] / "prices.json"

def load_prices(path):
    workbook = load_workbook(path, read_only=True, data_only=True)
    try:
        sheet = workbook.active
        values = sheet.values
        header = next(values)
        units = next(values)
        if not isinstance(units[0], datetime) or str(units[1]).strip().lower() != "1000 l" or str(units[2]).strip().lower() != "1000 l":
            raise ValueError("Unexpected workbook date or unit columns; review format before updating")
        if "Euro-super 95" not in str(header[1]) or "Gas oil automobile" not in str(header[2]):
            raise ValueError("Fuel columns do not match expected EU workbook format")
        record = next((r for r in values if str(r[0]).strip().lower() == "sweden"), None)
        if record is None:
            raise ValueError("Sweden row not found")
        petrol, diesel = record[1:3]
        if not all(isinstance(x, (float, int)) and 200 < x < 5000 for x in (petrol, diesel)):
            raise ValueError("Unexpected fuel price values; expected EUR per 1000 litres")
        return units[0].date().isoformat(), petrol / 1000, diesel / 1000
    finally:
        workbook.close()

def ecb_rate(date):
    with urllib.request.urlopen(ECB_CSV, timeout=25) as response:
        body = response.read().decode("utf-8-sig")
    for row in csv.DictReader(io.StringIO(body)):
        if row.get("Date", "").strip() == date:
            rate = float(row["SEK"].strip())
            if not 1 < rate < 30:
                raise ValueError("Unexpected ECB exchange rate")
            return rate
    raise ValueError("ECB EUR/SEK reference rate missing for the workbook date; no update performed")

def main():
    if len(sys.argv) != 2:
        raise SystemExit("Usage: python3 scripts/update_prices.py weekly-file.xlsx")
    date, petrol, diesel = load_prices(sys.argv[1])
    old = json.loads(DEST.read_text(encoding="utf-8"))
    if date < old["date"]:
        raise ValueError("Workbook is older than the stored snapshot; refusing to overwrite")
    rate = ecb_rate(date)
    old.update(date=date, petrol95=petrol, diesel=diesel, updatedAutomatically=False)
    old["fx"].update(date=date, rate=rate)
    DEST.write_text(json.dumps(old, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(f"Updated {DEST}: {date}, B95 {petrol*rate:.2f} SEK/L, diesel {diesel*rate:.2f} SEK/L")
    print("Updated repository prices.json; workflow will publish this file through GitHub Pages.")

if __name__ == "__main__":
    main()
