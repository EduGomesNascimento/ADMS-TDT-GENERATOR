"""
build_modulewise.py — Gama de sinais POR MÓDULO, empírica, lida da base completa.
==============================================================================
Substitui a expansão "cega" (build_expand_signals despejava os 639+55 do padrão
em TODOS os tipos). Aqui a gama de cada tipo é descoberta EMPIRICAMENTE:

1. A base completa (Export_base_Full) tem ~237k sinais discretos + 132k analógicos,
   no formato {objid}_{suffix}. Cada `objid` é um device físico real.
2. Agrupamos os sufixos por objid → cada device vira um "saco de sinais" (a
   assinatura real daquele equipamento em campo).
3. Cada tipo do catálogo (alimentador, TR, BC, LT, barra, TSA) tem uma SEMENTE:
   os sufixos já aprendidos da TDT-fonte específica daquele tipo.
4. Classificamos cada objid ao tipo cuja semente ele mais cobre (Jaccard ponderado).
   Só conta se cobrir >= COVER_MIN da semente (senão é device de outro tipo).
5. A UNIÃO dos sufixos de todos os devices classificados num tipo = a gama real
   daquele tipo. A frequência (nº de devices que têm o sufixo) vira o "peso".
6. Inserimos no catálogo os sufixos que existem no sigla_index (têm linha-template
   clonável) e ainda não estão no tipo, grupo "Base Real", ordenados por frequência.

Resultado: cada equipamento oferece só os sinais que equipamentos DAQUELE tipo
realmente possuem na base — não os 639 despejados em todos.

Uso: python build_modulewise.py        (scan ~6 min na base de 98MB)
"""
from __future__ import annotations
import json
import time
from collections import defaultdict
from pathlib import Path
import openpyxl

ROOT = Path(__file__).resolve().parent
DATA = ROOT / "data"
BASE = ROOT.parent.parent / "Export_base_Full__27_fev_2026.xlsx"

MIN_OVERLAP = 5       # device precisa partilhar >=5 sinais com a semente do tipo
SEED_MIN = 4          # objid precisa ter >=4 sufixos p/ classificar (evita ruído)
NAME_COL = 0
AOR_COL = 5
CACHE = None          # data/_devices_cache.json — evita re-escanear a base (8 min)

SHEETS = [("DNP3_DiscreteSignals", "discrete"), ("DNP3_AnalogSignals", "analog")]


def _split(name: str, known: set) -> tuple[str, str | None]:
    """Separa (device_prefix, suffix) usando o vocabulário conhecido de sufixos.

    A base é MISTA: nomes `1212973_50F1` (objid numérico) e `FWB_AL13_52-13_50F1`
    (estruturado) convivem. Pega o sufixo CONHECIDO mais curto a partir da direita
    (take=1 antes de 2,3) — assim `50F1` casa antes do poluído `52-13_50F1`.
    """
    parts = name.split("_")
    for take in (1, 2, 3):
        if len(parts) > take:
            cand = "_".join(parts[-take:])
            if len(cand) > 1 and cand in known:
                return "_".join(parts[:-take]), cand
    # fallback: device = tudo menos último token (sufixo desconhecido, ignorado)
    if len(parts) >= 2:
        return "_".join(parts[:-1]), None
    return name, None


def scan_base(known: set):
    """device_prefix -> {'discrete': set(suffix), 'analog': set(suffix), aor counters}."""
    wb = openpyxl.load_workbook(BASE, read_only=True, data_only=True)
    devices = defaultdict(lambda: {"discrete": set(), "analog": set(), "distr": 0, "trans": 0})
    for sheet, klass in SHEETS:
        ws = wb[sheet]
        n = 0
        for row in ws.iter_rows(min_row=5, values_only=True):
            if not row or not row[NAME_COL]:
                continue
            name = str(row[NAME_COL])
            oid, suf = _split(name, known)
            if not suf:
                continue
            d = devices[oid]
            d[klass].add(suf)
            aor = row[AOR_COL] if len(row) > AOR_COL else None
            if aor:
                a = str(aor)
                if "Distr" in a:
                    d["distr"] += 1
                elif "Trans" in a:
                    d["trans"] += 1
            n += 1
        print(f"  {sheet}: {n} sinais lidos")
    wb.close()
    print(f"  devices físicos (objid) distintos: {len(devices)}")
    return devices


def main():
    t = time.time()
    cat = json.loads((DATA / "catalog.json").read_text(encoding="utf-8"))
    idx = json.loads((DATA / "sigla_index.json").read_text(encoding="utf-8"))
    known = {"discrete": idx["DNP3_DiscreteSignals"], "analog": idx["DNP3_AnalogSignals"]}

    # --- sementes: sufixos genuinamente aprendidos (não os despejados "Padrão ADMS") ---
    EXPANDED = ("Padrão ADMS", "Base Real")   # grupos adicionados por scripts (não-aprendidos)
    known_all = set(known["discrete"]) | set(known["analog"])

    def real_code(suf: str) -> str | None:
        """Código real de um sufixo de semente (ignora placeholders e caminho de
        sub-módulo): 'TR1AT_52-2_43LR'→'43LR', '<<MODULE>>_52-2_81'→'81'."""
        parts = [p for p in suf.split("_") if not p.startswith("<<")]
        for take in (1, 2):
            if len(parts) >= take:
                cand = "_".join(parts[-take:])
                if len(cand) > 1 and cand in known_all:
                    return cand
        return parts[-1] if parts else None

    # seeds = sufixos aprendidos (p/ expandir); seed_real = códigos reais (p/ casar na base)
    seeds = {}
    seed_real = {}
    expandable = {}
    for d in cat["deviceTypes"]:
        seed = {"discrete": set(), "analog": set()}
        for klass in ("discrete", "analog"):
            for s in d["signals"].get(klass, []):
                if s.get("group") not in EXPANDED:
                    seed[klass].add(s["suffix"])
        seeds[d["id"]] = seed
        all_suf = seed["discrete"] | seed["analog"]
        real = {real_code(s) for s in all_suf}
        real.discard(None)
        seed_real[d["id"]] = real
        # expansível só se a geração usa <<PREFIX>> simples (sufixos reais no índice).
        # TR/linha são consolidados (sufixos = caminho de sub-módulo) → não expande,
        # mas ainda classificam seus devices (via seed_real) p/ não poluir alimentador.
        frac = len(all_suf & known_all) / max(len(all_suf), 1)
        expandable[d["id"]] = frac >= 0.5
        print(f'semente {d["id"]:18s} aprendidos={len(all_suf):3d} reais={len(real):3d} '
              f'expansivel={expandable[d["id"]]}')

    cache_path = DATA / "_devices_cache.json"
    if cache_path.exists():
        print(f"usando cache {cache_path.name} (sem re-escanear a base)...")
        raw = json.loads(cache_path.read_text(encoding="utf-8"))
        devices = {oid: {"discrete": set(v["discrete"]), "analog": set(v["analog"])}
                   for oid, v in raw.items()}
    else:
        print("escaneando base completa (demora ~6 min)...")
        devices = scan_base(known_all)
        cache_path.write_text(json.dumps(
            {oid: {"discrete": sorted(v["discrete"]), "analog": sorted(v["analog"])}
             for oid, v in devices.items()}, ensure_ascii=False), encoding="utf-8")
        print(f"cache salvo em {cache_path.name}")

    # --- classifica cada device físico ao tipo cuja semente ele mais cobre ---
    # acumula frequência de cada sufixo por tipo
    freq = {d["id"]: {"discrete": defaultdict(int), "analog": defaultdict(int)} for d in cat["deviceTypes"]}
    matched_devices = defaultdict(int)
    seed_all = seed_real   # casa pelos CÓDIGOS REAIS (resolve TR/LT e des-incha alimentador)
    for oid, dev in devices.items():
        dev_all = dev["discrete"] | dev["analog"]
        if len(dev_all) < SEED_MIN:
            continue
        # overlap ABSOLUTO (nº de sinais partilhados) — não penaliza sementes grandes
        # (TR/LT são consolidados, espalhados por vários objids físicos)
        best_type, best_ov = None, 0
        for tid, sall in seed_all.items():
            ov = len(dev_all & sall)
            if ov > best_ov:
                best_ov, best_type = ov, tid
        if best_type and best_ov >= MIN_OVERLAP:
            matched_devices[best_type] += 1
            for klass in ("discrete", "analog"):
                for suf in dev[klass]:
                    freq[best_type][klass][suf] += 1

    # --- reconstrói as listas de sinais: aprendidos primeiro, depois "Base Real" ---
    added_total = 0
    for d in cat["deviceTypes"]:
        tid = d["id"]
        sig = d["signals"]
        for klass in ("discrete", "analog"):
            lst = sig.setdefault(klass, [])
            # remove expansões anteriores ("Padrão ADMS"/"Base Real") — mantém aprendidos
            lst[:] = [s for s in lst if s.get("group") not in EXPANDED]
            have = {s.get("suffix") for s in lst}
            if not expandable[tid]:
                continue   # TR/linha: consolidados, já ricos — não recebem sinais genéricos
            # candidatos = sufixos vistos em devices deste tipo, ordenados por frequência
            ranked = sorted(freq[tid][klass].items(), key=lambda kv: -kv[1])
            for suf, count in ranked:
                if suf in have or len(suf) <= 1 or suf not in known[klass]:
                    continue
                row = known[klass][suf]
                desc = row[2] if len(row) > 2 and row[2] else suf
                meas = row[4] if len(row) > 4 and row[4] else ("Valor Medido" if klass == "analog" else "Status")
                lst.append({
                    "suffix": suf,
                    "description": str(desc),
                    "aliasLabel": str(desc),
                    "klass": klass,
                    "measurementType": str(meas),
                    "signalSubType": None,
                    "phases": "ABC",
                    "hasCommand": False,
                    "command": None,
                    "group": "Base Real",
                    "occurrences": count,     # nº de devices reais com este sinal
                    "row": row,
                })
                have.add(suf)
                added_total += 1
        d["signalCount"] = {"discrete": len(sig.get("discrete", [])), "analog": len(sig.get("analog", []))}

    (DATA / "catalog.json").write_text(json.dumps(cat, ensure_ascii=False), encoding="utf-8")
    print(f"\n+{added_total} sinais empíricos somados. devices classificados por tipo:")
    for d in cat["deviceTypes"]:
        s = d["signals"]
        print(f'  {d["id"]:18s} devices={matched_devices[d["id"]]:5d}  dig={len(s.get("discrete",[]))} anl={len(s.get("analog",[]))}')
    print(f"tempo total: {round(time.time()-t,1)}s")


if __name__ == "__main__":
    main()
