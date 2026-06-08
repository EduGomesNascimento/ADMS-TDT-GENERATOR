"""
build_lt.py — adiciona "Linha de Transmissão (LT)" ao catálogo (consolidado).
O módulo É o nome da linha (LTSAS, LTKGT…); device varia (52-x disjuntor, 29-x/
89-x seccionadoras, LINHA_A/LINHA_P grupos de proteção). Usuário informa alias +
nome da linha. Tokeniza alias -> <<ALIAS>> e linha -> <<MODULE>> (por token inteiro).
"""
from __future__ import annotations
import json, re
from pathlib import Path
import openpyxl
from build_catalog import _lbl_idx, detect_command, load_dictionary

ROOT = Path(__file__).resolve().parent
SRC = ROOT.parent.parent
BASE = SRC / "Export_base_Full__27_fev_2026.xlsx"
CATALOG = ROOT / "data" / "catalog.json"
LINE_RE = re.compile(r"^LT[A-Z0-9]+$")


def make_tok(alias, line):
    pairs = [(alias, "<<ALIAS>>"), (line, "<<MODULE>>")]
    def tok(v):
        if not isinstance(v, str):
            return v
        out = v
        for lit, ph in pairs:
            out = re.sub(r"(?<![A-Za-z0-9])" + re.escape(lit) + r"(?![A-Za-z0-9])", ph, out)
        return out
    return tok


def group_of(device, suffix):
    d = str(device)
    if d.endswith("_A") or "_A_" in d or suffix.startswith("A"):
        pass
    if re.search(r"_A$|_A_", d):
        return "Proteção A"
    if re.search(r"_P$|_P_", d):
        return "Proteção P"
    if d.startswith("52"):
        return "Disjuntor"
    if d.startswith(("29", "89")):
        return "Seccionadora"
    return "Linha"


def main():
    cat = json.loads(CATALOG.read_text(encoding="utf-8"))
    cols = cat["columns"]
    dictionary = load_dictionary()
    wb = openpyxl.load_workbook(BASE, read_only=True, data_only=True)

    buckets = {}  # (alias, line) -> {discrete:[], analog:[]}
    for sheet, klass in (("DNP3_DiscreteSignals", "discrete"), ("DNP3_AnalogSignals", "analog")):
        for r in wb[sheet].iter_rows(min_row=5, values_only=True):
            nm = r[0]
            if not nm:
                continue
            p = str(nm).split("_")
            if len(p) < 4 or not LINE_RE.match(p[1]):
                continue
            buckets.setdefault((p[0], p[1]), {"discrete": [], "analog": []})[klass].append(r)

    if not buckets:
        print("nenhuma LT encontrada"); return
    (alias, line), data = max(buckets.items(),
        key=lambda kv: (len(kv[1]["analog"]) > 0, len(kv[1]["discrete"]) + len(kv[1]["analog"])))
    tokenize = make_tok(alias, line)
    print(f"LT escolhida: {alias}_{line} -> {len(data['discrete'])} dig, {len(data['analog'])} anl")

    signals = {"discrete": [], "analog": []}
    for klass in ("discrete", "analog"):
        sheet = "DNP3_AnalogSignals" if klass == "analog" else "DNP3_DiscreteSignals"
        lab = _lbl_idx(cols[sheet])
        seen = set()
        for r in data[klass]:
            nm = str(r[0])
            rel = nm[len(alias) + 1:]
            if rel in seen:
                continue
            seen.add(rel)
            trow = [tokenize(c) for c in r]
            while len(trow) < len(cols[sheet]):
                trow.append(None)
            p = nm.split("_")
            device = p[2]
            code = p[-1]
            info = dictionary.get(code.upper(), {})
            ra = r[2] if len(r) > 2 else None
            command = detect_command(r, lab) if klass == "discrete" else None
            signals[klass].append({
                "suffix": tokenize(rel), "code": code, "group": group_of(device, "_".join(p[3:])),
                "description": info.get("description") or (str(ra) if ra else code),
                "aliasLabel": str(ra) if ra else "", "klass": info.get("klass", klass),
                "measurementType": r[4] if len(r) > 4 else None,
                "signalSubType": r[14] if len(r) > 14 else None,
                "phases": r[15] if len(r) > 15 else None,
                "hasCommand": command is not None, "command": command, "row": trow,
            })

    dev = {
        "id": "linha", "label": "Linha de Transmissão (LT)",
        "description": "Linha de transmissão — disjuntor, seccionadoras e grupos de proteção (A/P) consolidados.",
        "source": f"{alias}_{line}", "consolidated": True, "paramKind": "line",
        "defaults": {"module": line},
        "signalCount": {"discrete": len(signals["discrete"]), "analog": len(signals["analog"])},
        "signals": signals,
    }
    cat["deviceTypes"] = [d for d in cat["deviceTypes"] if d["id"] != "linha"] + [dev]
    CATALOG.write_text(json.dumps(cat, ensure_ascii=False, indent=2, default=str), encoding="utf-8")
    print(f"LT adicionada: {len(signals['discrete'])} dig, {len(signals['analog'])} anl")


if __name__ == "__main__":
    main()
