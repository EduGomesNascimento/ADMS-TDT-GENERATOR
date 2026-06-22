"""
build_expand_signals.py — Aumenta a GAMA de sinais por equipamento.

Cada tipo (alimentador, TR, BC, LT, barra, TSA) foi aprendido de UMA TDT só, então
tem poucos sinais. Este script adiciona, em cada tipo, TODOS os SIGLAs do índice
(sigla_index.json) que ainda não estão nele — num grupo "Padrão ADMS" — para que o
assistente ofereça a gama completa. Os sinais originais (aprendidos) ficam primeiro.

Idempotente: só adiciona o que falta. Uso: python build_expand_signals.py
"""
import json
from pathlib import Path

DATA = Path(__file__).parent / "data"


def _signal_from_index(suffix: str, row: list, klass: str) -> dict:
    desc = row[2] if len(row) > 2 and row[2] else suffix
    measurement = row[4] if len(row) > 4 and row[4] else ("Valor Medido" if klass == "analog" else "Status")
    return {
        "suffix": suffix,
        "description": str(desc),
        "aliasLabel": str(desc),
        "klass": klass,
        "measurementType": str(measurement),
        "signalSubType": None,
        "phases": "ABC",
        "hasCommand": False,
        "command": None,
        "group": "Padrão ADMS",
        "row": row,
    }


def main():
    cat = json.loads((DATA / "catalog.json").read_text(encoding="utf-8"))
    idx = json.loads((DATA / "sigla_index.json").read_text(encoding="utf-8"))
    pa = json.loads((DATA / "padrao_adms.json").read_text(encoding="utf-8"))
    disc_idx = idx["DNP3_DiscreteSignals"]
    anl_idx = idx["DNP3_AnalogSignals"]
    # só os SIGLAs do PADRÃO OFICIAL (gama relevante, evita inchar com 2994)
    off_disc = set(pa.get("discrete", {})); off_anl = set(pa.get("analog", {}))

    added_total = 0
    for d in cat["deviceTypes"]:
        sig = d["signals"]
        for klass, src, official in (("discrete", disc_idx, off_disc),
                                     ("analog", anl_idx, off_anl)):
            lst = sig.setdefault(klass, [])
            have = {s.get("suffix") for s in lst}
            before = len(lst)
            for suffix, row in src.items():
                if suffix in have or len(suffix) <= 1 or suffix not in official:
                    continue
                lst.append(_signal_from_index(suffix, row, klass))
            added_total += len(lst) - before
        d["signalCount"] = {
            "discrete": len(sig.get("discrete", [])),
            "analog": len(sig.get("analog", [])),
        }

    (DATA / "catalog.json").write_text(json.dumps(cat, ensure_ascii=False), encoding="utf-8")
    print(f"Expandido: +{added_total} sinais somados nos tipos.")
    for d in cat["deviceTypes"]:
        s = d["signals"]
        print(f"  {d['id']:18s} dig={len(s.get('discrete',[]))} anl={len(s.get('analog',[]))}")


if __name__ == "__main__":
    main()
