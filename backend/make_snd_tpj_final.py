"""
make_snd_tpj_final.py — TDT DEFINITIVA do vão TPJ (SND) usando o CHANGESET como gabarito.

O changeset PT-MOD-SE-SND.xml contém os 102 sinais que já EXISTEM no modelo
(nomes SND_LTTPJ_* consolidados) e o Device Mapping exato de cada um.
A TDT deve usar ESSES nomes — assim o import casa com os objetos do modelo em
vez de criar sinais novos (que causavam os conflitos de client point).

Índices: re-sequenciados por tipo (UTR_SND_1 é nova). Comandáveis sem valor
real → filler 9599+. AOR: SAN Trans.

Uso: python make_snd_tpj_final.py
"""
from __future__ import annotations
import json, re
import xml.etree.ElementTree as ET
from pathlib import Path
import tdt_engine as E

XML = Path("C:/Users/egnpo/Downloads/PT-MOD-SE-SND.xml")
OUT = Path("C:/Users/egnpo/Downloads/TDT_SND_TPJ_final.xlsx")
DATA = Path(__file__).parent / "data"
AOR = "SAN Trans"
FILL_START = 9599

# sigla do MODELO sem template exato na base → template da variante mais próxima
# (o NOME do sinal continua o do modelo; só a linha-molde vem da variante)
FALLBACK = {
    "21_1": "67_1", "21_2": "67_2",
    "21_21FT": "21_21", "21_21NT": "21_21N",
    "21_50F1": "21_50_1", "21_50F2": "21_50_2",
    "21_51F": "21_51FN",
    "67_51FN": "21_51FN", "67_MTRF": "MTRF",
}


def load_changeset():
    """[(tipo A/D, nome, device_mapping)] — sinais que existem no modelo."""
    root = ET.parse(XML).getroot()
    out = []
    for rd in root.iter("ResourceDescription"):
        idel = rd.find("id")
        if idel is None or idel.get("type") not in ("DSIGNAL", "ASIGNAL"):
            continue
        props = {p.get("id"): p.get("value") for p in rd.iter("Property")}
        nm = props.get("IDOBJ_NAME")
        if nm:
            out.append(("A" if idel.get("type") == "ASIGNAL" else "D",
                        nm, props.get("SIGNAL_DEVICEMAPPING") or ""))
    return out


def main():
    idx = json.loads((DATA / "sigla_index.json").read_text(encoding="utf-8"))
    dig, anl = idx["DNP3_DiscreteSignals"], idx["DNP3_AnalogSignals"]

    cat = E.Catalog()
    SH = "DNP3_DiscreteSignals"
    i_odt = cat.label_index(SH, "Output Data Type")
    i_ctrl = cat.label_index(SH, "Control Codes")

    def cmd_count(sigla):
        t = dig.get(sigla) or []
        ctrl = t[i_ctrl] if i_ctrl is not None and i_ctrl < len(t) else None
        if ctrl not in (None, ""):
            return len(str(ctrl).split(";"))
        odt = t[i_odt] if i_odt is not None and i_odt < len(t) else None
        return 1 if odt not in (None, "") else 0

    signals = load_changeset()
    print(f"changeset: {len(signals)} sinais do modelo")

    discrete, analog, missing = [], [], []
    seq = {"di": 0, "ai": 0}
    fill = FILL_START
    for tipo, nome, dm in sorted(signals, key=lambda x: (x[0], x[1])):
        parts = nome.split("_")
        sigla = "_".join(parts[3:]) if len(parts) > 3 else parts[-1]
        if tipo == "D" and sigla not in dig and sigla in FALLBACK:
            sigla = FALLBACK[sigla]
        if tipo == "A":
            if sigla not in anl:
                missing.append((tipo, nome, sigla)); continue
            analog.append({"sigla": sigla, "nome": nome, "escala": "",
                           "inCoord": seq["ai"], "aor": AOR, "deviceMapping": dm})
            seq["ai"] += 1
        else:
            if sigla not in dig:
                missing.append((tipo, nome, sigla)); continue
            cc = cmd_count(sigla)
            out = ""
            if cc:
                out = f"{fill};{fill}" if cc == 2 else fill
                fill += 1
            discrete.append({"sigla": sigla, "nome": nome, "inCoord": seq["di"],
                             "outCoord": out, "aor": AOR, "deviceMapping": dm})
            seq["di"] += 1

    lista = {"discrete": discrete, "analog": analog, "discrete_analog": [],
             "inputErrors": [], "uncertain": set()}
    print(f"gerando: dig={len(discrete)} anl={len(analog)} | "
          f"SEM template na base: {len(missing)}")
    for t, nm, s in missing:
        print(f"   [FALTA NA BASE] {nm} (sigla {s})")

    tdt, report = E.generate_tdt_from_list(lista, protocol="dnp3", native=True)
    OUT.write_bytes(tdt)
    print(f"TDT salva: {OUT} ({len(tdt)} bytes) | "
          f"ok dig={report['discrete']['matched']} anl={report['analog']['matched']}")


if __name__ == "__main__":
    main()
