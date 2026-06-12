#!/usr/bin/env python3
"""Gera data.json a partir do export Bloomberg (OB_OS.xlsx).

Estrutura esperada em cada aba:
  Coluna A: Date | Coluna B: MOV_AVG_200D | Coluna D: Date | Coluna E: PX_LAST
  O ticker é lido da primeira célula de texto da coluna A (ex.: "SPY US EQUITY").

Uso:  python build.py [caminho_do_xlsx]   (default: OB_OS.xlsx)
"""
import sys, json, statistics
from datetime import datetime
import openpyxl

SKIP_SHEETS = {"Charts"}

def main(path="OB_OS.xlsx"):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    out = {}
    for name in wb.sheetnames:
        if name in SKIP_SHEETS:
            continue
        sh = wb[name]
        ticker, ma, px = None, {}, {}
        for r in sh.iter_rows(min_row=1, max_col=5, values_only=True):
            a, b, _, d, e = (list(r) + [None] * 5)[:5]
            if ticker is None and isinstance(a, str) and "Date" not in a and a.strip():
                ticker = a.strip()
            if isinstance(a, datetime) and isinstance(b, (int, float)):
                ma[a.date()] = b
            if isinstance(d, datetime) and isinstance(e, (int, float)):
                px[d.date()] = e
        common = sorted(set(ma) & set(px))
        series = [(dt, px[dt] / ma[dt]) for dt in common if ma[dt] != 0]
        if not series:
            print(f"AVISO: aba '{name}' sem dados válidos — ignorada", file=sys.stderr)
            continue
        vals = [v for _, v in series]
        out[name] = {
            "ticker": ticker or name,
            "dates": [dt.isoformat() for dt, _ in series],
            "ratio": [round(v, 4) for v in vals],
            "mean": round(statistics.mean(vals), 4),
            "sd": round(statistics.pstdev(vals), 4),
            "last": round(vals[-1], 4),
            "lastDate": series[-1][0].isoformat(),
            "n": len(vals),
        }
        print(f"{name:12s} {ticker or '?':22s} {len(vals):>5d} obs  até {series[-1][0]}")
    with open("data.json", "w") as f:
        json.dump(out, f)
    print(f"\ndata.json gerado: {len(out)} ativos")

if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else "OB_OS.xlsx")
