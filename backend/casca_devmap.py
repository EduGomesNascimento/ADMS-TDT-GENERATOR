"""
casca_devmap.py — Device Mapping da SE CASCA.

FONTE DA VERDADE: PT-MOD-SE-CASCA.xml (export do changeset do modelo
"Casca_Obra"). Cada elemento traz a propriedade 1224979098644840199, que é o
"ID de Mapeamento SCADA" — exatamente o texto que a coluna Device Mapping da
TDT tem que conter. Se o texto não existir lá, o ADMS responde
"Could not find any device that corresponds to Device Mapping: ...".

O QUE O MODELO TEM (clone da CASCA ATUAL):
  alimentadores AL12 AL13 AL14 AL15 AL21 · linhas LTSCO LTPRI LTMRU
  trafos TR1(3W) TR2(2W) + TR12AT TR1AT TR1BT TR2AT TR2BT
  barras B138 BP23 · TSA3 · LTGPR
O QUE A LISTA NOVA PEDE (subestação modernizada):
  AL12..AL15 AL21 AL24 AL25 AL26 · BC1(AL18) BC2(AL28) · LT1 LT2 LT3
  TR6 TR7 (+AT/BT) · IB20 · transferências 24-1(AL18) 24-2(TRF29)
  BP69 BP113.8 BP213.8 · TSA1 TSA2 (RET1 RET2)

Ou seja: boa parte dos dispositivos AINDA NÃO EXISTE no modelo. Isso não se
resolve na TDT — tem que ser criado no diagrama. O que este módulo faz:

  1) casa pelo texto exato quando ele já existe no modelo;
  2) casa por MÓDULO + SUFIXO quando só o número do equipamento mudou
     (o alimentador AL12 do modelo usa 52-2, a lista nova usa 52-12 — é o
     mesmo vão, então vale CAS_AL12_52-2_DJ);
  3) dentro de um módulo que existe, se o relé específico não existe cai para
     o relé genérico _PROT e depois para o disjuntor _DJ
     ("na dúvida, põe no disjuntor" — instrução do usuário);
  4) módulo inexistente: devolve o nome canônico e MARCA como pendente, para
     o relatório listar o que precisa ser criado no Casca_Obra.

Uso: from casca_devmap import resolver; dm, origem, pendente = resolver(nome, sigla)
"""
from __future__ import annotations
import collections
import re
from pathlib import Path
import openpyxl

MODELO = Path("C:/Users/egnpo/Downloads/PT-MOD-SE-CASCA.xml")
TDT_ATUAL = Path("C:/Users/egnpo/Downloads/CASCA.xlsx")
PROP_SCADA_ID = "1224979098644840199"
HEADER_ROWS = 4

# ── sufixo (tipo de dispositivo) por sigla — aprendido da TDT atual da CASCA ──
_FAM = [
    # religador: na TDT atual TODO 79* aponta pro disjuntor, não pro relé
    (("79", "79OK", "79LO", "79_1", "79_2", "79RE", "79EP", "79BL"), "DJ"),
    (("IA", "IB", "IC", "IN", "IACC", "IBCC", "ICCC", "INCC", "I"), "TC"),
    (("P", "Q", "S", "FP", "COS"), "TC"),
    (("V", "VA", "VB", "VC", "VA_B", "VB_B", "VC_B", "VA_L", "VB_L", "VC_L",
      "VAB", "VBC", "VCA", "F", "FREQ"), "TP"),
    (("TAP", "CDC", "BCDC", "DCDC", "CDMT", "MDCD", "CDAM", "CDLR", "RLCD",
      "63CA", "63CD", "20CA", "20CD", "27CD", "59CD", "86CC", "71C", "71HI",
      "71LO", "TOC"), "COMTAP"),
    (("R90", "FC90", "DR90", "R90A", "R90B"), "TAP_REG"),
    (("VF", "FVF", "TOA", "TOD", "TOLE", "TEA", "TED", "TENR", "63T", "63TA",
      "63TD", "71T", "20A", "20D", "20TA", "20TD", "SCAR", "MEMB", "TOED",
      "DRT1", "DRT2", "FCT1", "FCT2", "FCPL", "86", "87", "87_T"), "TR"),
    (("SECF", "SECC", "SECT", "SECG", "SECB", "SELF", "SEC", "LIBM", "DSEC",
      "SLIB", "SECD"), "SEC"),
]
_POR_DEVICE = [("52-", "DJ"), ("89-", "SEC"), ("29-", "SEC"), ("24-", "DJ")]
_MEDIDA_TC = {"IA", "IB", "IC", "IN", "IACC", "IBCC", "ICCC", "INCC", "P", "Q", "S"}
_MEDIDA_TP = {"V", "VA", "VB", "VC", "VA_B", "VB_B", "VC_B", "VA_L", "VB_L",
              "VC_L", "VAB", "VBC", "VCA", "F"}
# ordem de rebaixamento dentro de um módulo que existe no modelo
_DEGRADA = {
    "COMTAP": ("TR", "DJ"), "TAP_REG": ("TR", "DJ"), "TR": ("DJ",),
    "TC": ("TP", "DJ"), "TP": ("TC", "DJ"), "SEC": ("DJ",), "DJ": (),
}

_CACHE = {}


# ─── catálogo do modelo (XML do changeset) ───────────────────────────────────
def catalogo():
    if "cat" in _CACHE:
        return _CACHE["cat"]
    validos, por_mod, tipos = set(), collections.defaultdict(dict), {}
    if MODELO.exists():
        txt = MODELO.read_text(encoding="utf-8-sig", errors="replace")
        for b in re.findall(r"<ResourceDescription>(.*?)</ResourceDescription>",
                            txt, re.S):
            m = re.search(rf'id="{PROP_SCADA_ID}" value="([^"]*)"', b)
            if not m or not m.group(1):
                continue
            dm = m.group(1)
            t = re.search(r'type="([^"]+)"', b)
            validos.add(dm)
            tipos.setdefault(dm, t.group(1) if t else "?")
            p = dm.split("_")
            if len(p) >= 3:
                por_mod[p[1]].setdefault("_".join(p[3:]) if len(p) > 3 else "", dm)
    _CACHE["cat"] = {"validos": validos, "por_mod": dict(por_mod), "tipos": tipos}
    return _CACHE["cat"]


# ─── sufixo por sigla (aprendido da TDT atual) ───────────────────────────────
def _aprender_siglas():
    if "sig" in _CACHE:
        return _CACHE["sig"]
    per = collections.defaultdict(collections.Counter)
    if TDT_ATUAL.exists():
        wb = openpyxl.load_workbook(TDT_ATUAL, read_only=True, data_only=True)
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
                nome, dm = str(r[0]).strip(), str(r[cD] or "").strip()
                p = nome.split("_")
                if not dm or not nome.startswith("CAS_") or len(p) < 4:
                    continue
                pref = f"CAS_{p[1]}_{p[2]}_"
                if dm.startswith(pref):
                    per["_".join(p[3:])][dm[len(pref):]] += 1
        wb.close()
    _CACHE["sig"] = {s: c.most_common(1)[0][0] for s, c in per.items()}
    return _CACHE["sig"]


def sufixo(sig: str, device: str) -> tuple[str, str]:
    tab = _aprender_siglas()
    if sig in tab:
        return tab[sig], "TDT atual"
    for fam, suf in _FAM:
        if sig in fam:
            return suf, "familia"
    if len(sig) > 2 and sig[:2] in ("P_", "A_"):
        resto = sig[2:]
        if resto in tab:
            return tab[resto], "TDT atual (base)"
        for fam, suf in _FAM:
            if resto in fam:
                return suf, "familia (base)"
        if resto[:1].isdigit():
            return f"{sig[0]}_PROT_{resto}", "pickup/alarme"
        return "DJ", "default"
    if sig in _MEDIDA_TC:
        return "TC", "medida"
    if sig in _MEDIDA_TP:
        return "TP", "medida"
    d = str(device or "")
    for pre, suf in _POR_DEVICE:
        if d.startswith(pre):
            if sig[:1].isdigit():
                return f"PROT_{sig}", "ANSI"
            return suf, "device"
    if sig[:1].isdigit():
        return f"PROT_{sig}", "ANSI"
    return "DJ", "default"


# ─── resolução contra o modelo ───────────────────────────────────────────────
def resolver(nome: str, sigla: str) -> tuple[str, str, str]:
    """(device_mapping, origem_da_regra, pendencia) — pendencia vazia = ok."""
    cat = catalogo()
    p = str(nome).split("_")
    alias = p[0] if p else "CAS"
    mod = p[1] if len(p) > 1 else ""
    dev = p[2] if len(p) > 2 else mod
    suf, origem = sufixo(sigla, dev)
    canonico = f"{alias}_{mod}_{dev}_{suf}"

    if not cat["validos"]:                       # sem XML: fica no canônico
        return canonico, origem, ""
    if canonico in cat["validos"]:
        return canonico, f"{origem} + modelo (exato)", ""

    domod = cat["por_mod"].get(mod)
    if domod:
        # 1) mesmo sufixo, equipamento renumerado (AL12: 52-12 -> 52-2)
        if suf in domod:
            return domod[suf], f"{origem} + modelo (equip. renumerado)", ""
        # 2) relé específico não existe -> relé genérico do vão
        if suf.startswith("PROT_") and "PROT" in domod:
            return domod["PROT"], f"{origem} + modelo (rele generico)", \
                f"rele {suf} nao existe em {mod}"
        # 3) rebaixa por tipo de dispositivo
        for alt in _DEGRADA.get(suf, ()):
            if alt in domod:
                return domod[alt], f"{origem} + modelo (fallback {alt})", \
                    f"{suf} nao existe em {mod}"
        if "DJ" in domod:
            return domod["DJ"], f"{origem} + modelo (disjuntor)", \
                f"{suf} nao existe em {mod}"
    return canonico, origem, f"MODULO {mod} nao existe no Casca_Obra"


def device_mapping(nome: str, sigla: str) -> tuple[str, str]:
    dm, origem, _ = resolver(nome, sigla)
    return dm, origem


if __name__ == "__main__":
    cat = catalogo()
    print(f"modelo: {len(cat['validos'])} IDs de mapeamento SCADA")
    print(f"modulos: {sorted(cat['por_mod'])}")
    print(f"siglas aprendidas da TDT atual: {len(_aprender_siglas())}")
