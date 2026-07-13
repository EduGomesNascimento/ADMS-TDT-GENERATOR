"""
make_san2_tdt.py — Gera a TDT do SAN2 a partir da lista de pontos, aplicando as
regras do usuário:
  - SÓ sinais cujo SIGLA existe na FULL BASE (sigla_index)
  - família "21" (distância) desambiguada p/ as variantes reais da base
  - índice DNP3 e de comando RE-SEQUENCIADOS globais (um UTR, sem colisão)
  - Device Mapping conforme a TDT de referência (21→21_PROT, 25→PROT_25x);
    o que faltar/tiver dúvida → DISJUNTOR ({alias}_{module}_{52-xx}_DJ)

Uso: python make_san2_tdt.py <arquivo_san2.xlsx> [alias]
"""
from __future__ import annotations
import json, re, sys
from pathlib import Path
import ai_mapper as A
import tdt_engine as E

DATA = Path(__file__).parent / "data"
IDX = json.loads((DATA / "sigla_index.json").read_text(encoding="utf-8"))
BASE_DIG = set(IDX["DNP3_DiscreteSignals"]); BASE_ANL = set(IDX["DNP3_AnalogSignals"])

# ── família 21 (distância): descrição → SIGLA da base (variantes da foto) ──────
_R21 = [
    (re.compile(r'FASE\s*A\b|TRIP\s*A\b', re.I), '21FA'),
    (re.compile(r'FASE\s*B\b|TRIP\s*B\b', re.I), '21FB'),
    (re.compile(r'FASE\s*C\b|TRIP\s*C\b', re.I), '21FC'),
    (re.compile(r'\bN\b.*TRIP|TRIP.*\bN\b|NEUTR', re.I), '21N'),
    (re.compile(r'Z1.?SF|Z1.?MF|MONOF', re.I), '21MF'),
    (re.compile(r'Z1.?MO|MULTIF|TRIF', re.I), '21MO'),
    (re.compile(r'Z2\b', re.I), '21Z2'),
    (re.compile(r'Z3\b', re.I), '21Z3'),
    (re.compile(r'Z4\b', re.I), '21Z4'),
    (re.compile(r'BLOQUEAD|PROT\.?21', re.I), '21'),
]

def _fix_21(desc: str) -> str | None:
    if '21' not in str(desc):
        return None
    for rx, sig in _R21:
        if rx.search(desc):
            return sig
    return '21'

# ── Device Mapping: por família de SIGLA ──────────────────────────────────────
# O Device Mapping da TDT referencia o "ID de Mapeamento SCADA" do objeto no
# MODELO (não o nome de exibição). Confirmado na tela do modelo:
#   disjuntor do vão TPJ → SND_LTTPJ_52-4_DJ  (módulo do MODELO = LTTPJ,
#   consolidado; 52-4 SEM zero à esquerda). O disjuntor é COMUM aos relés.
# Proteção usa o objeto do relé no modelo (SND_LTTPJ_LTTPJ_21_PROT / _PROT_21FA
# / _PROT_25x). Seccionadoras: ID no modelo não confirmado → regra da dúvida:
# DISJUNTOR.

# aba da lista → módulo do MODELO (linha consolidada)
_MODEL_LINE = {'LT67TPJ2': 'LTTPJ', 'LT21TPJ2': 'LTTPJ'}

def _model_module(sheet_module: str) -> str:
    return _MODEL_LINE.get(sheet_module, sheet_module)

def _short_dev(dev: str) -> str:
    """52-04 → 52-4 (modelo nomeia sem zero à esquerda)."""
    return re.sub(r'-0+(\d)', r'-\1', dev)

def _device_mapping(alias: str, module: str, breaker: str, sigla: str) -> str:
    s = sigla.upper()
    line = _model_module(module)
    if s in ('21FA', '21FB', '21FC'):
        return f"{alias}_{line}_{line}_PROT_{s}"
    if s.startswith('21'):                       # zonas/elementos de distância
        return f"{alias}_{line}_{line}_21_PROT"
    if s.startswith('25'):                       # sincronismo
        return f"{alias}_{line}_{line}_PROT_{s}"
    # tudo o mais (incl. seccionadoras) / dúvida → DISJUNTOR comum do vão
    return f"{alias}_{line}_{_short_dev(breaker)}_DJ"

_CLEAN = A._clean_token
_BRK_RE = re.compile(r'\b(52-?\d{1,3})\b')

def build(path: str, alias_cli: str | None = None, only_sheet: str | None = None):
    data = Path(path).read_bytes()
    raw, alias_det = A.parse_raw_excel(data)
    alias = _CLEAN(alias_cli or alias_det or 'SND')
    if only_sheet:                               # filtra por módulo (ex.: 'TPJ')
        raw = [s for s in raw if only_sheet.upper() in s.source_sheet.upper()]
        print(f"filtro módulo '{only_sheet}': {len(raw)} sinais nas abas "
              f"{sorted({s.source_sheet for s in raw})}")

    # descobre o disjuntor (52-xx) de cada aba a partir das descrições
    breaker = {}
    for s in raw:
        m = _BRK_RE.search(str(s.description))
        if m and s.source_sheet not in breaker:
            breaker[s.source_sheet] = m.group(1).replace('52', '52-').replace('52--', '52-')
    default_brk = '52-01'

    mapped = A.map_signals(raw, protocol='dnp3', llm_cfg={'provider': 'none'})

    # aplica a regra 21 ANTES de filtrar
    for m in mapped:
        f = _fix_21(m.description)
        if f and f in BASE_DIG:
            m.sigla = f; m.confidence = max(m.confidence, 90); m.confidence_label = 'ALTA'

    # ── comando no TEMPLATE: se a SIGLA da base é comandável, o ADMS exige
    #    Output Coordinates preenchido ("cannot be left empty") ────────────────
    cat = E.Catalog()
    SH = 'DNP3_DiscreteSignals'
    i_odt = cat.label_index(SH, 'Output Data Type')
    i_ctrl = cat.label_index(SH, 'Control Codes')
    i_dir = cat.label_index(SH, 'Direction')

    def _tpl(sigla):
        return IDX[SH].get(sigla) or []

    def _cmd_count(sigla) -> int:
        """0 = não-comandável; 1/2 = nº de coords de comando do template."""
        t = _tpl(sigla)
        def g(i): return t[i] if i is not None and i < len(t) else None
        ctrl = g(i_ctrl)
        if ctrl not in (None, ''):
            return len(str(ctrl).split(';'))
        if g(i_odt) not in (None, '') or g(i_dir) in ('Write', 'ReadWrite'):
            return 1
        return 0

    # monta a lista: índice GLOBAL re-sequenciado + device mapping + AOR Trans.
    # REGRA: nada sem index/comando — o que não tem valor real ganha filler ALTO
    # (9599+) para o ADMS aceitar e ficar visível p/ ajuste posterior.
    AOR = 'SAN Trans'      # nome COMPLETO do grupo no modelo (com espaço = literal)
    FILL_START = 9599
    fill = {'in': FILL_START, 'out': FILL_START}
    def _filler(kind):
        v = fill[kind]; fill[kind] += 1
        return v

    def _fmt_out(sigla, v) -> str:
        return f"{v};{v}" if _cmd_count(sigla) == 2 else v

    discrete, analog = [], []
    seq = {'di': 0, 'ai': 0, 'co': 0}
    seen = {}
    kept = dropped = 0
    for m in mapped:
        if not m.sigla:
            dropped += 1; continue
        base = BASE_ANL if m.signal_type == 'analog' else BASE_DIG
        if m.sigla not in base:                  # SÓ o que existe na FULL BASE
            dropped += 1; continue
        kept += 1
        module = _CLEAN(m.source_sheet) or 'XX'
        brk = breaker.get(m.source_sheet, default_brk)
        nome = A._unique_name(seen, alias, module, m.sigla)
        dm = _device_mapping(alias, module, brk, m.sigla)
        if m.signal_type == 'analog':
            analog.append({'sigla': m.sigla, 'nome': nome, 'escala': '',
                           'inCoord': seq['ai'], 'aor': AOR, 'deviceMapping': dm})
            seq['ai'] += 1
        elif m.signal_type == 'command':
            # comando também precisa de Input Coordinates (status do template):
            # sem valor real na lista → filler alto
            discrete.append({'sigla': m.sigla, 'nome': nome,
                             'inCoord': _filler('in'),
                             'outCoord': _fmt_out(m.sigla, seq['co']),
                             'aor': AOR, 'deviceMapping': dm})
            seq['co'] += 1
        else:
            cc = _cmd_count(m.sigla)
            out = _fmt_out(m.sigla, _filler('out')) if cc else ''
            discrete.append({'sigla': m.sigla, 'nome': nome, 'inCoord': seq['di'],
                             'outCoord': out, 'aor': AOR, 'deviceMapping': dm})
            seq['di'] += 1

    lista = {'discrete': discrete, 'analog': analog, 'discrete_analog': [],
             'inputErrors': [], 'uncertain': set()}
    print(f"alias={alias} | brutos={len(raw)} | na base={kept} | fora/sem sigla={dropped}")
    print(f"  discretos={len(discrete)} analog={len(analog)} | "
          f"index dig 0-{seq['di']-1}, anl 0-{seq['ai']-1}, cmd 0-{seq['co']-1}")
    return lista, alias


if __name__ == "__main__":
    src = sys.argv[1] if len(sys.argv) > 1 else "C:/Users/egnpo/Downloads/SAN2_V05.xlsx"
    alias = sys.argv[2] if len(sys.argv) > 2 else None
    only = sys.argv[3] if len(sys.argv) > 3 else None
    lista, alias = build(src, alias, only)
    tdt, report = E.generate_tdt_from_list(lista, protocol='dnp3', native=True)
    tag = f"_{only}" if only else ""
    out = f"C:/Users/egnpo/Downloads/TDT_SAN2{tag}_gerada.xlsx"
    Path(out).write_bytes(tdt)
    print(f"TDT salva: {out} ({len(tdt)} bytes) | dig ok={report['discrete']['matched']} anl={report['analog']['matched']}")
