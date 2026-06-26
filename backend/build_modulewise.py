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

COVER_MIN = 0.30      # device precisa cobrir >=30% da semente do tipo p/ ser daquele tipo
SEED_MIN = 4          # objid precisa ter >=4 sufixos p/ classificar (evita ruído)
NAME_COL = 0
AOR_COL = 5

SHEETS = [("DNP3_DiscreteSignals", "discrete"), ("DNP3_AnalogSignals", "analog")]


def _suffix(name: str) -> str | None:
    """objid_suffix → suffix (tudo após o 1º '_')."""
    i = name.find("_")
    return name[i + 1:] if i > 0 and i < len(name) - 1 else None


def _objid(name: str) -> str:
    i = name.find("_")
    return name[:i] if i > 0 else name


def scan_base():
    """objid -> {'discrete': set(suffix), 'analog': set(suffix), 'aor': Counter}."""
    wb = openpyxl.load_workbook(BASE, read_only=True, data_only=True)
    devices = defaultdict(lambda: {"discrete": set(), "analog": set(), "distr": 0, "trans": 0})
    for sheet, klass in SHEETS:
        ws = wb[sheet]
        n = 0
        for row in ws.iter_rows(min_row=5, values_only=True):
            if not row or not row[NAME_COL]:
                continue
            name = str(row[NAME_COL])
            suf = _suffix(name)
            if not suf:
                continue
            oid = _objid(name)
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
    seeds = {}
    for d in cat["deviceTypes"]:
        seed = {"discrete": set(), "analog": set()}
        for klass in ("discrete", "analog"):
            for s in d["signals"].get(klass, []):
                if s.get("group") != "Padrão ADMS":   # só os aprendidos da TDT-fonte
                    seed[klass].add(s["suffix"])
        seeds[d["id"]] = seed
        print(f'semente {d["id"]:18s} dig={len(seed["discrete"])} anl={len(seed["analog"])}')

    print("escaneando base completa (demora ~6 min)...")
    devices = scan_base()

    # --- classifica cada device físico ao tipo cuja semente ele mais cobre ---
    # acumula frequência de cada sufixo por tipo
    freq = {d["id"]: {"discrete": defaultdict(int), "analog": defaultdict(int)} for d in cat["deviceTypes"]}
    matched_devices = defaultdict(int)
    for oid, dev in devices.items():
        total_suf = len(dev["discrete"]) + len(dev["analog"])
        if total_suf < SEED_MIN:
            continue
        best_type, best_score = None, 0.0
        for tid, seed in seeds.items():
            seed_all = seed["discrete"] | seed["analog"]
            if not seed_all:
                continue
            dev_all = dev["discrete"] | dev["analog"]
            cover = len(dev_all & seed_all) / len(seed_all)   # cobertura da semente
            if cover > best_score:
                best_score, best_type = cover, tid
        if best_type and best_score >= COVER_MIN:
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
            # remove a expansão cega antiga ("Padrão ADMS") — mantém só os aprendidos
            lst[:] = [s for s in lst if s.get("group") != "Padrão ADMS"]
            have = {s.get("suffix") for s in lst}
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
