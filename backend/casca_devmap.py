"""
casca_devmap.py — Device Mapping da SE CASCA aprendido da TDT ATUAL (CASCA.xlsx).

"OS DISPOSITIVOS SEGUEM COM O MESMO DEVICE MAPPING DA SUBESTAÇÃO ORIGINAL":
a TDT atual (abas IEC101/IEC104/DNP3) já tem o Device Mapping certo de cada
sinal. O nome do sinal lá segue o MESMO padrão da lista nova
(CAS_{MODULO}_{DEVICE}_{SIGLA}), então dá pra aprender a REGRA:

    Device Mapping = CAS_{MODULO}_{DEVICE}_{SUFIXO(SIGLA)}

onde SUFIXO é o "tipo de dispositivo" que aquela sigla aponta:
    PROT_{SIGLA}  proteções (50F, 51N, 21, 87...)   → relé do dispositivo
    DJ            aux. de disjuntor (MOLA, SF6A, BBAB, CCCO, 79...)
    SEC           aux. de seccionadora (43LR, LIBM, SECF, DSEC...)
    TR            sinais de transformador (TOA, TED, 63TA, VF1...)
    COMTAP        comutador de tap (TAP, CDC, BCDC)
    TAP_REG       regulador (R90, FC90, DR90)
    TC / TP       medidas (IA, IB, IC, P, Q / V, VB)

Os módulos da lista NOVA (AL18, TR7AT, LT1...) não existem na TDT antiga —
só a REGRA é transferível, e é ela que o modelo do "Casca_Obra" espera.

Uso: from casca_devmap import build; sufixo = build()['sigla'][SIGLA]
"""
from __future__ import annotations
import collections
from pathlib import Path
import openpyxl

TDT_ATUAL = Path("C:/Users/egnpo/Downloads/CASCA.xlsx")
HEADER_ROWS = 4

# Regras de fallback para siglas que não aparecem na TDT atual, na ordem:
#   1) sufixo aprendido da própria sigla (build()['sigla'])
#   2) família da sigla (prefixo)               → _FAM
#   3) tipo do dispositivo pelo nome do device  → _POR_DEVICE
#   4) 'DJ'  (instrução do usuário: na dúvida, disjuntor)
_FAM = [
    # religador: na TDT atual TODO 79* aponta pro disjuntor, não pro relé
    #   CAS_AL12_52-2_79 / _79OK / _79_1 -> CAS_AL12_52-2_DJ
    #   CAS_LTPRI_LTPRI_P_79LO           -> CAS_LTPRI_52-21_DJ
    (("79", "79OK", "79LO", "79_1", "79_2", "79RE", "79EP", "79BL"), "DJ"),
    # medidas
    (("IA", "IB", "IC", "IN", "IACC", "IBCC", "ICCC", "INCC", "I"), "TC"),
    (("P", "Q", "S", "FP", "COS"), "TC"),
    (("V", "VA", "VB", "VC", "VA_B", "VB_B", "VC_B", "VA_L", "VB_L", "VC_L",
      "VAB", "VBC", "VCA", "F", "FREQ"), "TP"),
    # comutador / regulador
    (("TAP", "CDC", "BCDC", "DCDC", "CDMT", "MDCD", "CDAM", "CDLR", "RLCD",
      "63CA", "63CD", "20CA", "20CD", "27CD", "59CD", "86CC", "71C", "71HI",
      "71LO", "TOC"), "COMTAP"),
    (("R90", "FC90", "DR90", "R90A", "R90B"), "TAP_REG"),
    # transformador (ventilação forçada segue TR na TDT atual: VF1/VF2/CAVF/FVF1/FVF2)
    (("VF", "FVF", "TOA", "TOD", "TOLE", "TEA", "TED", "TENR", "63T", "63TA", "63TD", "71T",
      "20A", "20D", "20TA", "20TD", "SCAR", "MEMB", "TOED", "DRT1", "DRT2",
      "FCT1", "FCT2", "FCPL", "CDAM", "86", "87", "87_T"), "TR"),
    # seccionadora
    (("SECF", "SECC", "SECT", "SECG", "SECB", "SELF", "SEC", "LIBM", "DSEC",
      "SLIB", "SECD"), "SEC"),
]
_POR_DEVICE = [("52-", "DJ"), ("89-", "SEC"), ("29-", "SEC"), ("24-", "DJ")]

_MEDIDA_TC = {"IA", "IB", "IC", "IN", "IACC", "IBCC", "ICCC", "INCC", "P", "Q", "S"}
_MEDIDA_TP = {"V", "VA", "VB", "VC", "VA_B", "VB_B", "VC_B", "VA_L", "VB_L",
              "VC_L", "VAB", "VBC", "VCA", "F"}


def _aprender():
    """SIGLA → sufixo majoritário observado na TDT atual da CASCA."""
    if not TDT_ATUAL.exists():
        return {}, collections.Counter()
    wb = openpyxl.load_workbook(TDT_ATUAL, read_only=True, data_only=True)
    per = collections.defaultdict(collections.Counter)
    for sn in wb.sheetnames:
        if "Signals" not in sn and "DiscreteAnalog" not in sn:
            continue
        rows = list(wb[sn].iter_rows(values_only=True))
        if len(rows) <= HEADER_ROWS:
            continue
        hdr = [str(c or "").strip() for c in rows[HEADER_ROWS - 1]]
        ix = {n: i for i, n in enumerate(hdr) if n}
        cD = ix.get("Device Mapping")
        if cD is None:
            continue
        for r in rows[HEADER_ROWS:]:
            if not r or not r[0]:
                continue
            nome = str(r[0]).strip()
            dm = str(r[cD] or "").strip()
            p = nome.split("_")
            if not dm or not nome.startswith("CAS_") or len(p) < 4:
                continue
            pref = f"CAS_{p[1]}_{p[2]}_"
            if dm.startswith(pref):          # mesmo módulo+device → regra pura
                per["_".join(p[3:])][dm[len(pref):]] += 1
    wb.close()
    return {s: c.most_common(1)[0][0] for s, c in per.items()}, per


_CACHE = None


def build():
    global _CACHE
    if _CACHE is None:
        sigla, bruto = _aprender()
        _CACHE = {"sigla": sigla, "bruto": bruto}
    return _CACHE


def sufixo(sig: str, device: str, modulo: str = "") -> tuple[str, str]:
    """Retorna (sufixo, origem) — origem serve p/ auditoria no relatório."""
    tab = build()["sigla"]
    if sig in tab:
        return tab[sig], "TDT atual"
    for fam, suf in _FAM:
        if sig in fam:
            return suf, "familia"
    # pickup/alarme de proteção: CAS_X_Y_P_5FA -> CAS_X_Y_P_PROT_5FA
    if len(sig) > 2 and sig[:2] in ("P_", "A_"):
        resto = sig[2:]
        if resto in tab:                       # ex.: A_79LO -> DJ (segue o 79LO)
            return tab[resto], "TDT atual (base)"
        for fam, suf in _FAM:
            if resto in fam:
                return suf, "familia (base)"
        if resto[:1].isdigit():
            return f"{sig[0]}_PROT_{resto}", "pickup/alarme"
        return "DJ", "default"
    d = str(device or "")
    # medida em device de módulo (ex.: CAS_LT1_LT1_IA) → TC/TP
    if sig in _MEDIDA_TC:
        return "TC", "medida"
    if sig in _MEDIDA_TP:
        return "TP", "medida"
    for pre, suf in _POR_DEVICE:
        if d.startswith(pre):
            # proteção (sigla começa com dígito = código ANSI) → relé do device
            if sig[:1].isdigit() or sig.startswith(("P_", "A_")):
                return f"PROT_{sig}", "ANSI"
            return suf, "device"
    if sig[:1].isdigit() or sig.startswith(("P_", "A_")):
        return f"PROT_{sig}", "ANSI"
    return "DJ", "default"


def device_mapping(nome: str, sigla: str) -> tuple[str, str]:
    """CAS_{MOD}_{DEV}_{SUFIXO} a partir do NOME da lista (CAS_MOD_DEV_SIGLA)."""
    p = str(nome).split("_")
    alias = p[0] if p else "CAS"
    mod = p[1] if len(p) > 1 else ""
    dev = p[2] if len(p) > 2 else mod
    suf, origem = sufixo(sigla, dev, mod)
    return f"{alias}_{mod}_{dev}_{suf}", origem


if __name__ == "__main__":
    tab = build()["sigla"]
    print(f"siglas aprendidas da TDT atual: {len(tab)}")
    c = collections.Counter(v if not v.startswith("PROT_") else "PROT_*"
                            for v in tab.values())
    print(c.most_common())
