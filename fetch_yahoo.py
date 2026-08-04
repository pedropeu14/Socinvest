#!/usr/bin/env python3
"""Gera data.json (aba Técnico MM200) a partir do Yahoo Finance — sem Bloomberg.

Para cada ativo do UNIVERSE, baixa até 20 anos de fechamentos diários da API
de gráficos do Yahoo (sem chave), calcula a média móvel simples de 200 dias e
o ratio preço/MM200, e escreve data.json no MESMO schema que build.py produzia
a partir do OB_OS.xlsx — o index.html não muda nada.

Equivalências com o export Bloomberg:
  PX_LAST        -> close diário do Yahoo (ajustado por splits)
  MOV_AVG_200D   -> SMA de 200 pregões calculada aqui, sobre a mesma série

Os primeiros 199 pregões de cada série são descartados (aquecimento da média).
Os valores ficam muito próximos dos da Bloomberg, mas não idênticos bit a bit
(convenções de ajuste e fechamento de índice diferem marginalmente).

pe.json (Valuation) segue vindo do p_e.xlsx via build.py — não existe série
histórica gratuita de P/E trailing no Yahoo.

Uso:  python fetch_yahoo.py          (escreve data.json no diretório atual)
Só stdlib. Roda em CI (.github/workflows/build.yml) em dias úteis após o
fechamento americano, e localmente quando se quiser.
"""
import json
import sys
import time
import urllib.error
import urllib.parse
import urllib.request

# Universo = os mesmos ativos do export Bloomberg, identificados pelo ticker
# Bloomberg (o front-end deriva o nome de exibição da primeira palavra).
# yahoo=None usa a regra padrão: primeira palavra, com "/" -> "-".
UNIVERSE = [
    # (ticker Bloomberg, símbolo Yahoo ou None para regra padrão)
    ("000660 KS EQUITY", "000660.KS"),  # SK hynix — KRX, preços em KRW
    ("005930 KS EQUITY", "005930.KS"),  # Samsung Electronics — KRX, preços em KRW
    ("AAPL US EQUITY", None),
    ("AIRR US EQUITY", None),
    ("AMZN US EQUITY", None),
    ("ARA CN EQUITY", "ARA.TO"),        # Aclara Resources — Toronto
    ("ASML US EQUITY", None),
    ("BNK FP EQUITY", "BNK.PA"),        # Amundi Stoxx Europe 600 Banks — Paris
    ("BNT US EQUITY", None),
    ("BOTZ US EQUITY", None),
    ("BRK/B US EQUITY", "BRK-B"),
    ("COPX US EQUITY", None),
    ("DVY US EQUITY", None),
    ("DWX US EQUITY", None),
    ("EEM US EQUITY", None),
    ("EMXC US EQUITY", None),
    ("EQCH SW EQUITY", "EQCH.SW"),      # Invesco EQQQ CHF Hedged — SIX
    ("EWJ US EQUITY", None),
    ("EWL US EQUITY", None),
    ("EWY US EQUITY", None),
    ("EWZ US EQUITY", None),
    ("GDX US EQUITY", None),
    ("GOOGL US EQUITY", None),
    ("HYG US EQUITY", None),
    ("IBM US EQUITY", None),
    ("IEUR US EQUITY", None),
    ("IGV US EQUITY", None),
    ("IRM US EQUITY", None),
    ("ITUB US EQUITY", None),
    ("JBL US EQUITY", None),
    ("LLY US EQUITY", None),
    ("MELI US EQUITY", None),
    ("META US EQUITY", None),
    ("MSFT US EQUITY", None),
    ("MU US EQUITY", None),
    ("NFLX US EQUITY", None),
    ("NOW US EQUITY", None),
    ("NU US EQUITY", None),
    ("NVDA US EQUITY", None),
    ("ORCL US EQUITY", None),
    ("PBR US EQUITY", None),
    ("PICK US EQUITY", None),
    ("PLTR US EQUITY", None),
    ("QQQ US EQUITY", None),
    ("SOXX US EQUITY", None),
    ("SPOT US EQUITY", None),
    ("SPX INDEX", "^GSPC"),
    ("SPY US EQUITY", None),
    ("TIP US Equity", None),
    ("TLT US EQUITY", None),
    ("TSM US EQUITY", None),
    ("UBER US EQUITY", None),
    ("VALE US EQUITY", None),
    ("VRT US EQUITY", None),
    ("XAU Curncy", "GC=F"),             # ouro: futuro COMEX (spot XAU não há no Yahoo)
    ("XLE US EQUITY", None),
    ("XLF US EQUITY", None),
]

MM_WINDOW = 200
# range=20y com interval=1d devolve barras diárias; range=max degrada para
# barras mensais/semanais em séries longas — não usar.
CHART_URL = ("https://query1.finance.yahoo.com/v8/finance/chart/{symbol}"
             "?interval=1d&range=20y")
HEADERS = {"User-Agent": "Mozilla/5.0"}
PAUSE_S = 0.35
RETRIES = 2
# Abaixo disso o data.json anterior é preservado (não regredir o site por
# instabilidade pontual do Yahoo).
MIN_OK_FRACTION = 0.8


def yahoo_symbol(bbg: str, override: str | None) -> str:
    if override:
        return override
    return bbg.split()[0].replace("/", "-")


def fetch_closes(symbol: str) -> tuple[list[str], list[float]]:
    """Datas (ISO) e fechamentos diários. Levanta exceção após retries."""
    url = CHART_URL.format(symbol=urllib.parse.quote(symbol))
    last_err = None
    for attempt in range(RETRIES + 1):
        try:
            req = urllib.request.Request(url, headers=HEADERS)
            with urllib.request.urlopen(req, timeout=30) as resp:
                payload = json.load(resp)
            result = payload["chart"]["result"][0]
            ts = result.get("timestamp") or []
            closes = result["indicators"]["quote"][0]["close"]
            dates, vals = [], []
            for t, c in zip(ts, closes):
                if c is None or c <= 0:   # feriados/artefatos de meia-sessão
                    continue
                dates.append(time.strftime("%Y-%m-%d", time.gmtime(t)))
                vals.append(float(c))
            if len(vals) < MM_WINDOW + 10:
                raise ValueError(f"só {len(vals)} fechamentos — insuficiente p/ MM{MM_WINDOW}")
            return dates, vals
        except Exception as e:  # noqa: BLE001 — retentado, depois exposto
            last_err = e
            if attempt < RETRIES:
                time.sleep(1.5 * (attempt + 1))
    raise RuntimeError(f"{symbol}: {last_err}")


def load_previous(path: str = "data.json") -> dict:
    """data.json do run anterior, ou {} se não existir/estiver corrompido."""
    try:
        with open(path, encoding="utf-8") as f:
            prev = json.load(f)
        return prev if isinstance(prev, dict) else {}
    except (OSError, ValueError):
        return {}


def merge_series(prev_entry: dict | None, dates: list[str],
                 ratios: list[float]) -> tuple[list[str], list[float], list[str]]:
    """União da série nova com a anterior. Ponto novo vence; antigo preenche lacuna.

    Existe porque o Yahoo RETIRA fechamento que já publicou: em 2026-08-04 o
    fechamento de 2026-08-03 voltou como `close: None` em todos os papéis
    americanos, embora estivesse disponível na noite anterior. Sem união, o run
    seguinte republicava a série sem aquele dia — e o dado bom desaparecia do
    site com o build verde e nenhum alerta.

    Só resgata datas dentro da janela da série nova (>= primeira data dela).
    Sem esse limite, a janela deslizante de 20 anos faria o arquivo crescer
    para sempre, ressuscitando o começo da série a cada execução.
    """
    if not prev_entry:
        return dates, ratios, []
    prev_d = prev_entry.get("dates") or []
    prev_r = prev_entry.get("ratio") or []
    if not prev_d or len(prev_d) != len(prev_r) or not dates:
        return dates, ratios, []

    novos = set(dates)
    limite = dates[0]
    resgatados = sorted(d for d in prev_d if d not in novos and d >= limite)
    if not resgatados:
        return dates, ratios, []

    merged = {d: r for d, r in zip(prev_d, prev_r) if d >= limite}
    merged.update(zip(dates, ratios))          # o fresco sempre vence
    ordenado = sorted(merged)
    return ordenado, [merged[d] for d in ordenado], resgatados


def mm200_ratio(dates: list[str], closes: list[float]) -> tuple[list[str], list[float]]:
    """Ratio close/SMA200 a partir do 200º pregão (soma rolante, O(n))."""
    out_d, out_r = [], []
    rolling = sum(closes[:MM_WINDOW])
    for i in range(MM_WINDOW - 1, len(closes)):
        if i >= MM_WINDOW:
            rolling += closes[i] - closes[i - MM_WINDOW]
        ma = rolling / MM_WINDOW
        out_d.append(dates[i])
        out_r.append(closes[i] / ma)
    return out_d, out_r


def main() -> int:
    previous = load_previous()
    out, failures, resgates = {}, [], {}
    for bbg, override in UNIVERSE:
        sym = yahoo_symbol(bbg, override)
        try:
            dates, closes = fetch_closes(sym)
            rd, rr = mm200_ratio(dates, closes)
        except (RuntimeError, ValueError) as e:
            failures.append(bbg)
            print(f"FALHA {bbg:18s} ({sym}): {e}", file=sys.stderr)
            time.sleep(PAUSE_S)
            continue
        key = bbg.split()[0]
        rd, rr, recuperados = merge_series(previous.get(key), rd, rr)
        if recuperados:
            resgates[key] = recuperados
            print(f"      {bbg:18s} resgatou {len(recuperados)} data(s) do run "
                  f"anterior: {', '.join(recuperados[-3:])}")
        mean = sum(rr) / len(rr)
        sd = (sum((v - mean) ** 2 for v in rr) / len(rr)) ** 0.5
        out[key] = {
            "ticker": bbg,
            "dates": rd,
            "ratio": [round(v, 4) for v in rr],
            "mean": round(mean, 4),
            "sd": round(sd, 4),
            "last": round(rr[-1], 4),
            "lastDate": rd[-1],
            "n": len(rr),
        }
        print(f"OK    {bbg:18s} ({sym:8s}) {len(rr):>5d} obs  {rd[0]} -> {rd[-1]}")
        time.sleep(PAUSE_S)

    baixados = len(out)
    if baixados < MIN_OK_FRACTION * len(UNIVERSE):
        print(f"\nABORTADO: só {baixados}/{len(UNIVERSE)} ativos baixados "
              f"(mínimo {MIN_OK_FRACTION:.0%}) — data.json anterior preservado",
              file=sys.stderr)
        return 1

    # Ativo que falhou hoje mas existia antes é MANTIDO com o dado antigo, não
    # apagado: sumir do site é pior que aparecer defasado, e o rodapé do
    # index.html já sinaliza séries mais velhas que a moda. Sem isso, uma falha
    # pontual do Yahoo num ticker o removia do painel com o build verde.
    herdados = []
    for bbg, _ in UNIVERSE:
        key = bbg.split()[0]
        if key not in out and key in previous:
            out[key] = previous[key]
            herdados.append(f"{key}@{previous[key].get('lastDate', '?')}")

    with open("data.json", "w", encoding="utf-8") as f:
        json.dump(out, f)
    print(f"\ndata.json gerado: {baixados}/{len(UNIVERSE)} baixados agora"
          f" · {len(out)} no arquivo"
          f" · falhas: {', '.join(failures) or 'nenhuma'}")
    if resgates:
        total = sum(len(v) for v in resgates.values())
        print(f"  resgate de datas que o Yahoo retirou: {total} ponto(s) em "
              f"{len(resgates)} ativo(s) — {', '.join(sorted(resgates))}")
    if herdados:
        print(f"  mantidos com dado anterior (falharam hoje): {', '.join(herdados)}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
