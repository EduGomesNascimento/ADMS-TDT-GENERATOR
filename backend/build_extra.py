"""
build_extra.py — adiciona tipos de módulo "simples" (um módulo só) ao catálogo,
extraídos da base completa numa única varredura: Banco de Capacitores (BC),
Interligação de Barras (IB) e Barra. Reutiliza a tokenização padrão
(alias_modulo_device) do engine.
"""
from __future__ import annotations
import json, re
from pathlib import Path
import openpyxl
from build_catalog import _lbl_idx, detect_command, load_dictionary


def make_tokenizer(alias, module, device):
    """Tokenização por TOKEN INTEIRO (evita super-substituição de módulos curtos
    como 'IB'). Substitui apenas ocorrências delimitadas por não-alfanumérico."""
    prefix = f"{alias}_{module}_{device}"
    pairs = [(prefix, "<<PREFIX>>"), (alias, "<<ALIAS>>"), (module, "<<MODULE>>"), (str(device), "<<DEVICE>>")]

    def tok(v):
        if not isinstance(v, str):
            return v
        out = v
        for lit, ph in pairs:
            if lit:
                out = re.sub(r"(?<![A-Za-z0-9])" + re.escape(lit) + r"(?![A-Za-z0-9])", ph, out)
        return out

    return tok

ROOT = Path(__file__).resolve().parent
SRC = ROOT.parent.parent
BASE = SRC / "Export_base_Full__27_fev_2026.xlsx"
CATALOG = ROOT / "data" / "catalog.json"

SPECS = [
    {"id": "banco_capacitores", "label": "Banco de Capacitores", "re": re.compile(r"^BC\d+$"),
     "description": "Banco de capacitores com disjuntor e proteções."},
    {"id": "interligacao_barras", "label": "Interligação de Barras (IB)", "re": re.compile(r"^IB\d*$"),
     "description": "Disjuntor de interligação/acoplamento de barras."},
]


def main():
    cat = json.loads(CATALOG.read_text(encoding="utf-8"))
    cols = cat["columns"]
    dictionary = load_dictionary()
    wb = openpyxl.load_workbook(BASE, read_only=True, data_only=True)

    # 1 varredura: coleta linhas dos módulos de interesse, agrupadas por instância
    buckets = {s["id"]: {} for s in SPECS}  # id -> (alias,module,device) -> {sheet:[rows]}
    for sheet, klass in (("DNP3_DiscreteSignals", "discrete"), ("DNP3_AnalogSignals", "analog")):
        for r in wb[sheet].iter_rows(min_row=5, values_only=True):
            nm = r[0]
            if not nm:
                continue
            p = str(nm).split("_")
            if len(p) < 4:
                continue
            mod = p[1]
            for s in SPECS:
                if s["re"].match(mod):
                    key = (p[0], p[1], p[2])
                    d = buckets[s["id"]].setdefault(key, {"discrete": [], "analog": []})
                    d[klass].append(r)

    new_types = []
    for s in SPECS:
        inst = buckets[s["id"]]
        if not inst:
            print(f"{s['label']}: nenhuma instancia encontrada")
            continue
        # escolhe a instancia mais completa (mais sinais, prefere ter analog)
        best = max(inst.items(), key=lambda kv: (len(kv[1]["analog"]) > 0, len(kv[1]["discrete"]) + len(kv[1]["analog"])))
        (alias, module, device), data = best
        tokenize = make_tokenizer(alias, module, device)
        signals = {"discrete": [], "analog": []}
        for klass in ("discrete", "analog"):
            sheet = "DNP3_AnalogSignals" if klass == "analog" else "DNP3_DiscreteSignals"
            lab = _lbl_idx(cols[sheet])
            seen = set()
            for r in data[klass]:
                nm = str(r[0])
                suffix = nm.split("_", 3)[3] if len(nm.split("_")) > 3 else nm
                if suffix in seen:
                    continue
                seen.add(suffix)
                trow = [tokenize(c) for c in r]
                while len(trow) < len(cols[sheet]):
                    trow.append(None)
                code = nm.split("_")[-1]
                info = dictionary.get(code.upper(), {})
                ra = r[2] if len(r) > 2 else None
                command = detect_command(r, lab) if klass == "discrete" else None
                signals[klass].append({
                    "suffix": suffix, "code": code,
                    "description": info.get("description") or (str(ra) if ra else code),
                    "aliasLabel": str(ra) if ra else "", "klass": info.get("klass", klass),
                    "measurementType": r[4] if len(r) > 4 else None,
                    "signalSubType": r[14] if len(r) > 14 else None,
                    "phases": r[15] if len(r) > 15 else None,
                    "hasCommand": command is not None, "command": command, "row": trow,
                })
        new_types.append({
            "id": s["id"], "label": s["label"], "description": s["description"],
            "source": f"{alias}_{module}_{device}", "consolidated": False,
            "defaults": {"module": module, "device": device},
            "signalCount": {"discrete": len(signals["discrete"]), "analog": len(signals["analog"])},
            "signals": signals,
        })
        print(f"{s['label']}: {alias}_{module}_{device} -> {len(signals['discrete'])} dig, {len(signals['analog'])} anl")

    ids = {t["id"] for t in new_types}
    cat["deviceTypes"] = [d for d in cat["deviceTypes"] if d["id"] not in ids] + new_types
    CATALOG.write_text(json.dumps(cat, ensure_ascii=False, indent=2, default=str), encoding="utf-8")
    print("catalogo atualizado")


if __name__ == "__main__":
    main()
