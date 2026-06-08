"""
build_barra.py — adiciona o tipo "Barra" ao catálogo. Barras (módulos BP*, ex.:
BP69, BP113.8) têm estrutura módulo==device (ALIAS_BPxx_BPxx_SIGLA) e são
digitais (proteção diferencial de barra 87B). As SIGLAs vêm da lista de pontos
padrão; cada template (campos fixos) é puxado do índice global da base ADMS.

Modo de identificação: "bus" (usuário informa alias + nome da barra; o módulo e o
device recebem o mesmo valor → prefixo ALIAS_BUS_BUS).
"""
from __future__ import annotations
import json
from pathlib import Path
import openpyxl
from build_catalog import _lbl_idx, detect_command, load_dictionary

ROOT = Path(__file__).resolve().parent
SRC = ROOT.parent.parent
CATALOG = ROOT / "data" / "catalog.json"
INDEX = ROOT / "data" / "sigla_index.json"
LIST = SRC.parent / "lista_resumida_Input_com_campos_padrao - 2.xlsx"


def main():
    cat = json.loads(CATALOG.read_text(encoding="utf-8"))
    cols = cat["columns"]
    idx = json.loads(INDEX.read_text(encoding="utf-8"))
    dictionary = load_dictionary()
    lw = openpyxl.load_workbook(LIST, data_only=True)

    sheet = "DNP3_DiscreteSignals"
    lab = _lbl_idx(cols[sheet])
    didx = idx[sheet]

    seen = set()
    signals = []
    for r in list(lw["Discreto"].iter_rows(values_only=True))[1:]:
        nm = r[1]
        if not nm:
            continue
        p = str(nm).split("_")
        if len(p) < 2 or not p[1].startswith("BP"):
            continue
        sigla = str(r[0]).strip()
        if sigla in seen or sigla not in didx:
            continue
        seen.add(sigla)
        row = didx[sigla]
        info = dictionary.get(sigla.upper(), {})
        command = detect_command(row, lab)
        signals.append({
            "suffix": sigla, "code": sigla, "group": "Barra",
            "description": info.get("description") or sigla,
            "aliasLabel": "", "klass": "discrete",
            "measurementType": row[lab.get("Measurement Type")] if lab.get("Measurement Type") is not None else None,
            "signalSubType": row[lab.get("Signal Type")] if lab.get("Signal Type") is not None else None,
            "phases": row[lab.get("Phases")] if lab.get("Phases") is not None else None,
            "hasCommand": command is not None, "command": command, "row": row,
        })

    dev = {
        "id": "barra", "label": "Barra",
        "description": "Barramento da subestação — proteção diferencial (87B) e estados de barra.",
        "source": "BP (lista padrão)", "consolidated": True, "paramKind": "bus",
        "defaults": {"module": "BP138"},
        "signalCount": {"discrete": len(signals), "analog": 0},
        "signals": {"discrete": signals, "analog": []},
    }
    cat["deviceTypes"] = [d for d in cat["deviceTypes"] if d["id"] != "barra"] + [dev]
    CATALOG.write_text(json.dumps(cat, ensure_ascii=False, indent=2, default=str), encoding="utf-8")
    print(f"Barra adicionada: {len(signals)} sinais digitais -> {[s['code'] for s in signals]}")


if __name__ == "__main__":
    main()
