#!/usr/bin/env python3
"""Mantém a aba Valuation (P/E) viva entre exports da Bloomberg.

Modo padrão — EXTENSÃO POR PREÇO (todas as séries do p_e.xlsx):
  Entre dois resultados trimestrais o P/E trailing só se move pelo preço.
  Então, para cada série Bloomberg: EPS implícito = preço ÷ P/E no último
  ponto do export, e daí em diante P/E(dia) = fechamento do Yahoo ÷ esse EPS
  congelado. Exato na âncora; deriva conforme lucros novos são reportados —
  por isso vale subir um p_e.xlsx novo a cada temporada de resultados, o que
  re-ancora tudo automaticamente. A metodologia de EPS continua sendo a da
  Bloomberg (before XO items) para todos os ativos, consistente com o
  histórico desde 2002. Cada série estendida ganha "extendedFrom" (mostrado
  no site como "EPS de <data>").

Modo opcional --sec — P/E COMPUTADO (ações americanas do UNIVERSE):
  P/E(dia) = fechamento (Yahoo) ÷ EPS diluído GAAP TTM (SEC EDGAR, 10-Q/10-K),
  desde 2009, point-in-time pela data de protocolo, ajustado a splits, Q4 =
  anual − 3 trimestres. Validado contra a Bloomberg (8/15 nomes em ±10%; o
  resto diverge por GAAP vs lucro ajustado — ex.: um charge único derruba o
  EPS GAAP e o P/E dispara). Não usado por padrão: optou-se pela consistência
  metodológica da série Bloomberg estendida.

Este script ATUALIZA o pe.json existente (gerado por build.py --no-tec a
partir do p_e.xlsx). SX7E não tem preço no Yahoo e fica sem extensão.

Uso:  python fetch_pe.py [--sec]     (SEC_USER_AGENT="Nome email" p/ --sec)
Só stdlib. Roda em CI depois de build.py --no-tec.
"""
import json
import os
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
from datetime import date, timedelta

START = "2009-01-01"          # início da série computada (XBRL confiável de 2009+)
MIN_OBS = 30                  # mesmo piso do build_pe

# (ticker Bloomberg como aparece no pe.json, símbolo Yahoo, CIKs SEC)
# CIK extra cobre predecessor legal (Google Inc. antes da Alphabet, 2009-2015).
UNIVERSE = [
    ("AAPL US EQUITY", "AAPL", [320193]),
    ("AMZN US EQUITY", "AMZN", [1018724]),
    ("GOOGL US EQUITY", "GOOGL", [1652044, 1288776]),
    ("IBM US EQUITY", "IBM", [51143]),
    ("IRM US EQUITY", "IRM", [1020569]),
    ("JBL US EQUITY", "JBL", [898293]),
    ("LLY US EQUITY", "LLY", [59478]),
    ("MELI US EQUITY", "MELI", [1099590]),
    ("META US EQUITY", "META", [1326801]),
    ("MSFT US EQUITY", "MSFT", [789019]),
    ("MU US EQUITY", "MU", [723125]),
    ("NFLX US EQUITY", "NFLX", [1065280]),
    ("NOW US EQUITY", "NOW", [1373715]),
    ("NVDA US EQUITY", "NVDA", [1045810]),
    ("VRT US EQUITY", "VRT", [1674910]),
]

# TODAS as séries Bloomberg são ESTENDIDAS diariamente por preço: entre dois
# trimestres o P/E trailing só se move pelo preço, então EPS implícito =
# preço ÷ P/E no último ponto Bloomberg, e daí em diante P/E(dia) = preço
# Yahoo ÷ EPS congelado. Exato na âncora; deriva conforme novos lucros são
# reportados — re-ancora sozinho quando um novo p_e.xlsx é publicado (por
# isso vale subir a planilha a cada temporada de resultados).
# Símbolo Yahoo = 1ª palavra do ticker ("/"->"-"); exceções abaixo.
# SX7E não tem preço no Yahoo: fica sem extensão, parada no último export.
EXTEND_OVERRIDES = {
    "SPX": "^GSPC", "NDX": "^NDX", "SOX": "^SOX",
    "S5ENRS": "^GSPE", "S5FINL": "^SP500-40",
    "ASML": "ASML.AS",              # série Bloomberg é da linha de Amsterdã (EUR)
    "2330": "2330.TW",              # TSMC na bolsa de Taiwan (TWD)
    "VALE3": "VALE3.SA",            # Vale na B3 (BRL)
    "SX7E": None,
}


def extend_symbol(ticker: str) -> str | None:
    short = ticker.split()[0].upper()
    if short in EXTEND_OVERRIDES:
        return EXTEND_OVERRIDES[short]
    return short.replace("/", "-")

EPS_TAGS = ["EarningsPerShareDiluted", "EarningsPerShareBasicAndDiluted"]
SEC_UA = os.environ.get("SEC_USER_AGENT") or "Socinvest fetch_pe (pedrof.amorim@gmail.com)"
SEC_URL = "https://data.sec.gov/api/xbrl/companyconcept/CIK{cik:010d}/us-gaap/{tag}.json"
YH_URL = ("https://query1.finance.yahoo.com/v8/finance/chart/{symbol}"
          "?interval=1d&range=20y&events=splits")
PAUSE_S = 0.3
RETRIES = 2


def get_json(url: str, headers: dict) -> dict:
    last_err = None
    for attempt in range(RETRIES + 1):
        try:
            req = urllib.request.Request(url, headers=headers)
            with urllib.request.urlopen(req, timeout=30) as resp:
                return json.load(resp)
        except Exception as e:  # noqa: BLE001 — retentado, depois exposto
            last_err = e
            if attempt < RETRIES:
                time.sleep(1.5 * (attempt + 1))
    raise RuntimeError(f"{url}: {last_err}")


def sec_eps_entries(ciks: list[int]) -> list[dict]:
    """Registros de EPS diluído (todas as durações) dos CIKs, deduplicados."""
    entries = []
    for cik in ciks:
        for tag in EPS_TAGS:
            try:
                data = get_json(SEC_URL.format(cik=cik, tag=tag),
                                {"User-Agent": SEC_UA})
            except RuntimeError:
                continue
            for unit, items in data.get("units", {}).items():
                if "/shares" not in unit:
                    continue
                entries.extend(items)
            time.sleep(PAUSE_S)
            break  # primeira tag com dados atende; fallback só se a 1ª faltar
    # dedup por período, ficando com o protocolo mais antigo (point-in-time:
    # o número ficou público na primeira vez em que foi arquivado)
    by_period: dict[tuple, dict] = {}
    for e in entries:
        if not (e.get("start") and e.get("end") and e.get("val") is not None
                and e.get("filed")):
            continue
        k = (e["start"], e["end"], round(float(e["val"]), 4))
        if k not in by_period or e["filed"] < by_period[k]["filed"]:
            by_period[k] = e
    return list(by_period.values())


def days(a: str, b: str) -> int:
    return (date.fromisoformat(b) - date.fromisoformat(a)).days


def quarterly_ttm(entries: list[dict],
                  splits: list[tuple[str, float]]) -> list[tuple[str, str, float]]:
    """[(data_efetiva=filed, fim_do_período, eps_ttm)] a partir dos registros.

    Trimestres = durações de ~3 meses; Q4 derivado = anual − 3 trimestres do
    exercício. TTM = soma de 4 trimestres consecutivos (janela de ~1 ano).

    Cada EPS é normalizado para a base de ações ATUAL antes de qualquer soma:
    o número arquivado está na base vigente na data de protocolo (splits
    anteriores ao protocolo já vêm aplicados retroativamente pela empresa),
    então divide-se pelos splits POSTERIORES ao protocolo. Sem isso, somar
    trimestres pré-split com o anual pós-split (ex.: NFLX/NOW em 2025) produz
    Q4 derivado absurdo — e o preço do Yahoo já é ajustado, tem de bater.
    """
    def pick(bucket: dict, e: dict) -> None:
        k = e["end"]
        if k not in bucket or e["filed"] < bucket[k]["filed"]:
            bucket[k] = e

    quarters: dict[str, dict] = {}
    annuals: dict[str, dict] = {}
    for e in entries:
        d = days(e["start"], e["end"])
        if 80 <= d <= 100:
            pick(quarters, e)
        elif 350 <= d <= 380:
            pick(annuals, e)
    for bucket in (quarters, annuals):
        for e in bucket.values():
            e["val"] = float(e["val"]) / split_factor_after(splits, e["filed"])

    # Q4 = anual − (Q1+Q2+Q3 dentro do exercício); efetivo no filed do 10-K
    for a in annuals.values():
        if a["end"] in quarters:
            continue
        inside = [q for q in quarters.values()
                  if q["start"] >= a["start"] and q["end"] <= a["end"]]
        if len(inside) == 3:
            q4 = {"start": max(q["end"] for q in inside), "end": a["end"],
                  "val": float(a["val"]) - sum(float(q["val"]) for q in inside),
                  "filed": a["filed"], "derived": True}
            quarters[a["end"]] = q4

    qs = sorted(quarters.values(), key=lambda q: q["end"])
    ttm = []
    for i in range(3, len(qs)):
        window = qs[i - 3:i + 1]
        span = days(window[0]["start"], window[3]["end"])
        if not 330 <= span <= 400:      # trimestres não consecutivos (buraco)
            continue
        eff = max(q["filed"] for q in window)
        ttm.append((eff, window[3]["end"],
                    sum(float(q["val"]) for q in window)))
    ttm.sort()
    return ttm


_YAHOO_CACHE: dict[str, tuple] = {}


def yahoo_closes(symbol: str) -> tuple[list[str], list[float], list[tuple[str, float]]]:
    """Datas, fechamentos e lista de splits [(data, razão)]. Cacheado por símbolo."""
    if symbol in _YAHOO_CACHE:
        return _YAHOO_CACHE[symbol]
    payload = get_json(YH_URL.format(symbol=urllib.parse.quote(symbol)),
                       {"User-Agent": "Mozilla/5.0"})
    result = payload["chart"]["result"][0]
    ts = result.get("timestamp") or []
    closes = result["indicators"]["quote"][0]["close"]
    dates, vals = [], []
    for t, c in zip(ts, closes):
        if c is None or c <= 0:
            continue
        dates.append(time.strftime("%Y-%m-%d", time.gmtime(t)))
        vals.append(float(c))
    splits = []
    for s in (result.get("events", {}) or {}).get("splits", {}).values():
        ratio = float(s["numerator"]) / float(s["denominator"])
        if ratio > 0:
            splits.append((time.strftime("%Y-%m-%d", time.gmtime(s["date"])), ratio))
    splits.sort()
    _YAHOO_CACHE[symbol] = (dates, vals, splits)
    return dates, vals, splits


def split_factor_after(splits: list[tuple[str, float]], period_end: str) -> float:
    """Produto das razões de split posteriores ao fim do período do EPS."""
    f = 1.0
    for d, r in splits:
        if d > period_end:
            f *= r
    return f


def robust_stats(vals: list[float]) -> tuple[float, float]:
    sv = sorted(vals)

    def pct(p: float) -> float:
        k = (len(sv) - 1) * p
        f = int(k)
        return sv[f] + (sv[min(f + 1, len(sv) - 1)] - sv[f]) * (k - f)

    return pct(0.5), (pct(0.84) - pct(0.16)) / 2


MAX_STALE_DAYS = 450                    # sem relatório novo há +15 meses => sem P/E


def build_series(bbg: str, symbol: str, ciks: list[int]) -> dict | None:
    dates, closes, splits = yahoo_closes(symbol)
    ttm = quarterly_ttm(sec_eps_entries(ciks), splits)
    if not ttm:
        print(f"FALHA {bbg}: sem EPS trimestral utilizável na SEC", file=sys.stderr)
        return None
    series_d, series_v = [], []
    j = -1
    for d, c in zip(dates, closes):
        if d < START:
            continue
        while j + 1 < len(ttm) and ttm[j + 1][0] <= d:
            j += 1
        if j < 0:
            continue
        eff, _, eps = ttm[j]
        if eps <= 0:                    # P/E não existe com lucro negativo
            continue
        if days(eff, d) > MAX_STALE_DAYS:   # TTM velho demais: fluxo quebrou
            continue
        series_d.append(d)
        series_v.append(round(c / eps, 3))
    if len(series_v) < MIN_OBS:
        print(f"FALHA {bbg}: só {len(series_v)} observações de P/E", file=sys.stderr)
        return None
    med, sd = robust_stats(series_v)
    return {
        "ticker": bbg.upper(),
        "dates": series_d,
        "ratio": series_v,
        "mean": round(med, 3),
        "sd": round(sd, 3),
        "last": series_v[-1],
        "lastDate": series_d[-1],
        "n": len(series_v),
        "robust": True,
        "src": "computed",              # Yahoo preço ÷ SEC EPS (GAAP)
    }


def extend_entry(v: dict, symbol: str) -> bool:
    """Estende uma série Bloomberg até hoje usando o preço do Yahoo."""
    # idempotente: se já foi estendida antes, volta ao trecho Bloomberg puro
    if v.get("extendedFrom"):
        cut = v["extendedFrom"]
        keep = [i for i, d in enumerate(v["dates"]) if d <= cut]
        v["dates"] = v["dates"][:len(keep)]
        v["ratio"] = v["ratio"][:len(keep)]
        v["last"], v["lastDate"] = v["ratio"][-1], v["dates"][-1]
        del v["extendedFrom"]
    dates, closes, _ = yahoo_closes(symbol)
    anchor_price = None
    for d, c in zip(dates, closes):
        if d <= v["lastDate"]:
            anchor_price = c
        else:
            break
    if anchor_price is None or v["last"] <= 0:
        return False
    implied_eps = anchor_price / v["last"]
    ext = [(d, round(c / implied_eps, 3))
           for d, c in zip(dates, closes) if d > v["lastDate"]]
    if not ext:
        return False
    v["extendedFrom"] = v["lastDate"]
    v["dates"] += [d for d, _ in ext]
    v["ratio"] += [r for _, r in ext]
    v["last"], v["lastDate"], v["n"] = v["ratio"][-1], v["dates"][-1], len(v["ratio"])
    med, sd = robust_stats(v["ratio"])
    v["mean"], v["sd"] = round(med, 3), round(sd, 3)
    return True


def main() -> int:
    try:
        with open("pe.json") as f:
            pe = json.load(f)
    except FileNotFoundError:
        pe = {}
    computed, failures = 0, []
    if "--sec" in sys.argv:
        for bbg, symbol, ciks in UNIVERSE:
            s = build_series(bbg, symbol, ciks)
            time.sleep(PAUSE_S)
            if s is None:
                failures.append(bbg)
                continue
            # substitui a entrada Bloomberg de mesmo ticker (qualquer variação de chave)
            for k in [k for k, v in pe.items()
                      if v.get("ticker", k).split()[0].upper() == bbg.split()[0]]:
                del pe[k]
            pe[bbg.split()[0]] = s
            computed += 1
            print(f"OK    {bbg:18s} {s['n']:>5d} obs  {s['dates'][0]} -> {s['lastDate']}"
                  f"  P/E atual {s['last']:.1f}x")
    extended = extend_all(pe, "pe.json")
    if computed == 0 and extended == 0:
        print("ABORTADO: nada computado nem estendido — pe.json preservado",
              file=sys.stderr)
        return 1
    with open("pe.json", "w") as f:
        json.dump(pe, f)
    static = [v.get("ticker", k) for k, v in pe.items()
              if v.get("src") != "computed" and not v.get("extendedFrom")]
    print(f"\npe.json: {extended} séries Bloomberg estendidas por preço"
          + (f", {computed} computadas via SEC" if computed else "")
          + (f"; sem extensão: {', '.join(sorted(static))}" if static else "")
          + (f"; falhas SEC: {', '.join(failures)}" if failures else ""))
    # série forward (pe_fwd.json), quando o export tem a aba "Forward PE":
    # mesma extensão por preço — o P/E projetado também só se move pelo preço
    # entre revisões de estimativa.
    try:
        with open("pe_fwd.json") as f:
            fwd = json.load(f)
    except FileNotFoundError:
        return 0
    n = extend_all(fwd, "pe_fwd.json")
    with open("pe_fwd.json", "w") as f:
        json.dump(fwd, f)
    print(f"pe_fwd.json: {n} séries forward estendidas por preço")
    return 0


def extend_all(pe: dict, label: str) -> int:
    """Estende todas as séries Bloomberg de um dict por preço. Retorna quantas."""
    extended = 0
    for k, v in pe.items():
        if v.get("src") == "computed":
            continue
        sym = extend_symbol(v.get("ticker", k))
        if not sym:
            print(f"EXT   {v.get('ticker', k):18s} sem preço no Yahoo — fica até {v['lastDate']}")
            continue
        try:
            ok = extend_entry(v, sym)
        except (RuntimeError, KeyError) as e:
            print(f"EXT   {v.get('ticker', k):18s} falhou ({e}) — fica até {v['lastDate']}",
                  file=sys.stderr)
            time.sleep(PAUSE_S)
            continue
        if ok:
            extended += 1
            print(f"EXT   {v.get('ticker', k):18s} ({sym:9s}) estendida "
                  f"{v['extendedFrom']} -> {v['lastDate']}  P/E {v['last']:.1f}x")
        time.sleep(PAUSE_S)
    return extended


if __name__ == "__main__":
    sys.exit(main())
