#!/usr/bin/env python3
"""Gera data.json a partir do export Bloomberg (OB_OS.xlsx).

Estrutura esperada em cada aba:
  Coluna A: Date | Coluna B: MOV_AVG_200D | Coluna D: Date | Coluna E: PX_LAST
  O ticker é lido da primeira célula de texto da coluna A (ex.: "SPY US EQUITY").

Robusto a: datas armazenadas como texto (vários formatos), números como texto
(ponto ou vírgula decimal). O formato de data é detectado por coluna, exigindo
que todas as células parseiem e que a série fique em ordem cronológica.

Uso:  python build.py [caminho_do_xlsx]   (default: OB_OS.xlsx)
"""
import sys, json, statistics
from datetime import datetime, date
import openpyxl

SKIP_SHEETS = {"Charts"}
DATE_FORMATS = ["%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y", "%d.%m.%Y", "%d-%m-%Y", "%Y/%m/%d"]

def to_number(v):
    if isinstance(v, (int, float)):
        return float(v)
    if isinstance(v, str):
        s = v.strip().replace("'", "")
        if s.count(",") == 1 and s.count(".") == 0:
            s = s.replace(",", ".")
        else:
            s = s.replace(",", "")
        try:
            return float(s)
        except ValueError:
            return None
    return None

def parse_date_column(raw):
    """raw: lista de valores brutos. Retorna lista de date|None do mesmo tamanho."""
    out = [v.date() if isinstance(v, datetime) else (v if isinstance(v, date) else None) for v in raw]
    texts = [(i, v.strip()) for i, v in enumerate(raw) if isinstance(v, str) and v.strip() and "date" not in v.lower()]
    texts = [(i, s) for i, s in texts if any(c.isdigit() for c in s)]
    if not texts:
        return out
    best = None
    for fmt in DATE_FORMATS:
        parsed = []
        ok = True
        for i, s in texts:
            try:
                parsed.append((i, datetime.strptime(s[:10], fmt).date()))
            except ValueError:
                ok = False
                break
        if not ok:
            continue
        merged = list(out)
        for i, d in parsed:
            merged[i] = d
        seq = [d for d in merged if d]
        if seq == sorted(seq):  # série cronológica => formato correto
            best = parsed
            break
        if best is None:
            best = parsed  # fallback: primeiro formato que parseia tudo
    if best:
        for i, d in best:
            out[i] = d
    return out

def main(path="OB_OS.xlsx"):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    out = {}
    for name in wb.sheetnames:
        if name in SKIP_SHEETS:
            continue
        sh = wb[name]
        ticker = None
        colA, colB, colD, colE = [], [], [], []
        for r in sh.iter_rows(min_row=1, max_col=5, values_only=True):
            a, b, _, d, e = (list(r) + [None] * 5)[:5]
            if ticker is None and isinstance(a, str) and "Date" not in a and a.strip() and not any(c.isdigit() for c in a):
                ticker = a.strip()
            colA.append(a); colB.append(b); colD.append(d); colE.append(e)
        datesA = parse_date_column(colA)
        datesD = parse_date_column(colD)
        ma = {dt: to_number(v) for dt, v in zip(datesA, colB) if dt and to_number(v) is not None}
        px = {dt: to_number(v) for dt, v in zip(datesD, colE) if dt and to_number(v) is not None}
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
