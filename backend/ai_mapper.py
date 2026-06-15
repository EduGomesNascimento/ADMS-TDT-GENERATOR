"""
ai_mapper.py — Mapeador inteligente de listas de pontos não-padrão.

Fluxo:
  1. parse_raw_excel()   — extrai sinais de qualquer Excel de UTR
  2. map_signals()       — heurística (tokens) + LLM para mapear a SIGLA ADMS
  3. to_lista_resumida() — converte para o formato de parse_points_list (para gerar TDT)
"""
from __future__ import annotations

import io
import json
import logging
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

import openpyxl

log = logging.getLogger(__name__)

# ─── Modelos de dados ────────────────────────────────────────────────────────

@dataclass
class RawSignal:
    utr_id: str
    description: str
    dnp3_addr: Optional[int]
    signal_type: str       # analog | discrete | command | unknown
    module: str
    source_sheet: str

@dataclass
class MappedSignal:
    utr_id: str
    description: str
    dnp3_addr: Optional[int]
    signal_type: str
    module: str
    # resultado do mapeamento
    sigla: Optional[str] = None
    sigla_desc: Optional[str] = None
    confidence: int = 0
    confidence_label: str = "SEM"   # ALTA | MÉDIA | BAIXA | SEM
    alternative: Optional[str] = None
    match_method: str = "none"

# ─── Parsing de Excel ────────────────────────────────────────────────────────

_MODULE_RE   = re.compile(r'mód|módulo|modulo|module|bay|vão', re.I)
_UTR_ID_RE   = re.compile(r'identif|ponto.?na.?utr|utr.?id|ident\.', re.I)
_DNP3_RE     = re.compile(r'dnp3?\.?0?$|endereço.?dnp|index.?dnp', re.I)
_DESC_RE     = re.compile(r'descriç|descric|description|nome|designaç', re.I)
_SKIP_SHEETS = re.compile(r'^(capa|calculados?|slot|saca|sumário|config|setup|cron)', re.I)
_PREFER_SHEETS = re.compile(r'^(digitais|analogicos|analógicos|discretos)', re.I)
_ANALOG_SHEETS = re.compile(r'^(analóg|analog)', re.I)

_UTR_ID_PATTERN = re.compile(r'^[A-Z][A-Z0-9]{1,10}_[A-Z0-9_]{2,40}$')


def _is_utr_id(val) -> bool:
    if not val or not isinstance(val, str):
        return False
    return bool(_UTR_ID_PATTERN.match(str(val).strip()))


def _find_header_row(ws, max_scan: int = 30) -> Optional[int]:
    for r_idx, row in enumerate(ws.iter_rows(max_row=max_scan, values_only=True)):
        cells = [str(c).strip() if c else '' for c in row]
        text = ' '.join(cells).lower()
        if (_MODULE_RE.search(text) or _UTR_ID_RE.search(text)) and _DNP3_RE.search(text):
            return r_idx
    return None


def _map_columns(header_vals: list) -> dict:
    cols: dict = {}
    for idx, cell in enumerate(header_vals):
        if cell is None:
            continue
        c = str(cell).strip()
        if not c:
            continue
        if _MODULE_RE.search(c) and 'module' not in cols:
            cols['module'] = idx
        if _UTR_ID_RE.search(c) and 'utr_id' not in cols:
            cols['utr_id'] = idx
        if _DNP3_RE.search(c) and 'dnp3' not in cols:
            cols['dnp3'] = idx
        if _DESC_RE.search(c) and 'desc' not in cols:
            cols['desc'] = idx
    return cols


def _find_utr_col_heuristic(ws, data_start: int, max_cols: int = 20) -> Optional[int]:
    """Detecta coluna com padrão UTR_ID contando ocorrências."""
    counts: dict[int, int] = {}
    for row in ws.iter_rows(min_row=data_start, max_row=data_start + 40, values_only=True):
        for idx, cell in enumerate(row[:max_cols]):
            if _is_utr_id(cell):
                counts[idx] = counts.get(idx, 0) + 1
    if counts:
        return max(counts, key=counts.get)
    return None


def _parse_dnp3(raw) -> Optional[int]:
    if raw is None:
        return None
    try:
        v = int(float(str(raw)))
        return v if 1 <= v <= 99_999 else None
    except (ValueError, TypeError):
        return None


def _infer_type(desc: str, sheet_name: str) -> str:
    dl = desc.lower()
    if any(k in dl for k in ['corrente', 'tensão', 'tensao', 'potência', 'potencia', 'temp', 'tap atual', 'medida tap', 'frequência', 'frequencia']):
        return 'analog'
    if any(k in dl for k in ['abrir', 'fechar', 'comando']):
        return 'command'
    if _ANALOG_SHEETS.search(sheet_name):
        return 'analog'
    return 'discrete'


def _extract_sheet(ws, sheet_name: str, default_type: Optional[str] = None) -> list[RawSignal]:
    hdr_idx = _find_header_row(ws)
    if hdr_idx is None:
        return []

    n_cols = ws.max_column or 20
    header = [ws.cell(row=hdr_idx + 1, column=c + 1).value for c in range(n_cols)]
    cols = _map_columns(header)

    if 'utr_id' not in cols:
        utr_col = _find_utr_col_heuristic(ws, hdr_idx + 2, min(n_cols, 20))
        if utr_col is not None:
            cols['utr_id'] = utr_col

    if 'utr_id' not in cols:
        return []

    signals: list[RawSignal] = []
    for row in ws.iter_rows(min_row=hdr_idx + 2, values_only=True):
        if not any(row):
            continue
        raw_id = row[cols['utr_id']] if cols['utr_id'] < len(row) else None
        utr_id = str(raw_id).strip() if raw_id else ''
        if not utr_id or not _is_utr_id(utr_id):
            continue

        desc_idx = cols.get('desc')
        description = str(row[desc_idx]).strip() if (desc_idx is not None and desc_idx < len(row) and row[desc_idx]) else ''
        module_idx = cols.get('module')
        module = str(row[module_idx]).strip() if (module_idx is not None and module_idx < len(row) and row[module_idx]) else ''
        dnp3_idx = cols.get('dnp3')
        dnp3 = _parse_dnp3(row[dnp3_idx] if (dnp3_idx is not None and dnp3_idx < len(row)) else None)

        sig_type = default_type or _infer_type(description, sheet_name)
        signals.append(RawSignal(utr_id=utr_id, description=description,
                                 dnp3_addr=dnp3, signal_type=sig_type,
                                 module=module, source_sheet=sheet_name))
    return signals


def _detect_alias(wb) -> str:
    """Tenta detectar o alias da subestação na aba CAPA."""
    for name in wb.sheetnames:
        if name.upper() == 'CAPA':
            ws = wb[name]
            for row in ws.iter_rows(max_row=30, values_only=True):
                for cell in row:
                    if cell and re.search(r'SIGLA\s*:\s*([A-Z]{2,6})', str(cell), re.I):
                        m = re.search(r'SIGLA\s*:\s*([A-Z]{2,6})', str(cell), re.I)
                        if m:
                            return m.group(1).upper()
    return ''


def parse_raw_excel(data: bytes) -> tuple[list[RawSignal], str]:
    """
    Parseia qualquer Excel de lista de pontos.
    Retorna (sinais, alias_detectado).
    """
    wb = openpyxl.load_workbook(io.BytesIO(data), read_only=True, data_only=True)
    alias = _detect_alias(wb)

    preferred = [n for n in wb.sheetnames if _PREFER_SHEETS.search(n)]
    others    = [n for n in wb.sheetnames if n not in preferred and not _SKIP_SHEETS.search(n)]

    signals: list[RawSignal] = []
    seen: set[str] = set()

    def process(sheet_names):
        for sn in sheet_names:
            default = 'analog' if _ANALOG_SHEETS.search(sn) else 'discrete'
            for sig in _extract_sheet(wb[sn], sn, default):
                if sig.utr_id not in seen:
                    seen.add(sig.utr_id)
                    signals.append(sig)

    process(preferred if preferred else others)
    if not signals:
        process(others if preferred else [])

    wb.close()
    log.info("parse_raw_excel: %d sinais de %d abas; alias='%s'",
             len(signals), len(preferred or others), alias)
    return signals, alias

# ─── Carrega sigla_index ─────────────────────────────────────────────────────

_SHEET_TYPE = {
    'DNP3_DiscreteSignals': 'discrete',
    'DNP3_AnalogSignals': 'analog',
    'DNP3_DiscreteAnalog': 'analog',
    'IEC104_DiscreteSignals': 'discrete',
    'IEC104_AnalogSignals': 'analog',
}


def _load_sigla_flat(protocol: str = 'dnp3') -> dict:
    """
    Retorna {SIGLA: {desc, type, sheet_name}} — dicionário plano de todos os SIGLAs.
    """
    fname = 'sigla_index.json' if protocol == 'dnp3' else 'sigla_index_iec104.json'
    path = Path(__file__).parent / 'data' / fname
    with open(path, encoding='utf-8') as f:
        raw = json.load(f)

    flat: dict = {}
    for sheet_name, entries in raw.items():
        sig_type = _SHEET_TYPE.get(sheet_name, 'discrete')
        for sigla, row in entries.items():
            desc = row[2] if len(row) > 2 else sigla
            flat[sigla] = {'desc': str(desc or sigla), 'type': sig_type,
                           'sheet_name': sheet_name}
    return flat

# ─── Mapeamento heurístico ───────────────────────────────────────────────────

def _token_match(utr_id: str, sigla_flat: dict) -> Optional[tuple[str, int]]:
    """
    Tenta encontrar um SIGLA como sub-sequência de tokens do UTR ID.
    Retorna (sigla, confidence) ou None.
    """
    tokens = utr_id.upper().split('_')
    n = len(tokens)
    # Janelas de 1 a 4 tokens (evita pegar o alias como SIGLA)
    for start in range(1, n):          # começa em 1 para pular o primeiro token (alias/módulo)
        for window in range(1, min(5, n - start + 1)):
            candidate = '_'.join(tokens[start:start + window])
            if candidate in sigla_flat:
                return candidate, 100
    return None


def _desc_match(description: str, sigla_flat: dict) -> Optional[tuple[str, int]]:
    """Match por descrição exata (normalizada)."""
    if not description:
        return None
    desc_norm = re.sub(r'\s+', ' ', description.upper().strip())
    # Remove acentos básicos
    for old, new in [('Ã','A'),('Â','A'),('Á','A'),('À','A'),('É','E'),
                     ('Ê','E'),('Í','I'),('Ó','O'),('Ô','O'),('Ú','U'),('Ç','C')]:
        desc_norm = desc_norm.replace(old, new)
    for sigla, info in sigla_flat.items():
        sd = re.sub(r'\s+', ' ', str(info['desc']).upper().strip())
        if sd and sd == desc_norm:
            return sigla, 96
    return None

# ─── Backend LLM ─────────────────────────────────────────────────────────────

_BATCH_SIZE = 40

_PROMPT = """\
Você é especialista em sistemas SCADA de subestações elétricas brasileiras.

Mapeie cada sinal da UTR ao código SIGLA ADMS correto do dicionário abaixo.

=== DICIONÁRIO SIGLA ADMS ===
{sigla_list}

=== SINAIS PARA MAPEAR (idx | utr_id | descrição | tipo) ===
{signal_list}

REGRAS:
- SIGLAs são códigos curtos: "IA", "50F1", "43TC", "2649", "TAP" etc.
- Combine pelo código de função embutido no utr_id OU pela descrição semântica
- Analógicos: IA=Corrente FA, IB=FB, IC=FC, P=Pot.Ativa, Q=Reativa, VA/VB/VC=Tensão, TAP=comutador
- Se não encontrar correspondência use null
- confidence: 100=certeza, 90=muito provável, 70=provável, 50=possível, 0=sem match

Retorne APENAS um array JSON (sem markdown, sem explicação):
[{{"idx":0,"sigla":"IA","confidence":100,"alt":null}}, ...]
"""


def _build_sigla_text(sigla_flat: dict, type_filter: Optional[str] = None) -> str:
    lines = []
    for sigla, info in sorted(sigla_flat.items()):
        if type_filter == 'analog' and info['type'] != 'analog':
            continue
        if type_filter == 'discrete' and info['type'] == 'analog':
            continue
        lines.append(f"{sigla}: {info['desc']} ({info['type']})")
    return '\n'.join(lines)


def _parse_llm_json(text: str, n: int) -> list[dict]:
    m = re.search(r'\[.*\]', text, re.DOTALL)
    if not m:
        return [{'idx': i, 'sigla': None, 'confidence': 0, 'alt': None} for i in range(n)]
    try:
        return json.loads(m.group(0))
    except json.JSONDecodeError:
        return [{'idx': i, 'sigla': None, 'confidence': 0, 'alt': None} for i in range(n)]


def _call_llm(prompt: str, cfg: dict) -> str:
    provider = cfg.get('provider', 'gemini')
    model    = cfg.get('model', 'gemini-2.0-flash')
    api_key  = cfg.get('api_key', '')

    if provider == 'gemini':
        import google.generativeai as genai  # type: ignore
        genai.configure(api_key=api_key)
        resp = genai.GenerativeModel(model).generate_content(prompt)
        return resp.text

    if provider == 'groq':
        from groq import Groq  # type: ignore
        client = Groq(api_key=api_key)
        resp = client.chat.completions.create(
            model=model,
            messages=[{'role': 'user', 'content': prompt}],
            temperature=0.05,
        )
        return resp.choices[0].message.content

    if provider == 'ollama':
        import requests
        base = cfg.get('base_url', 'http://localhost:11434')
        resp = requests.post(f'{base}/api/chat', json={
            'model': model,
            'messages': [{'role': 'user', 'content': prompt}],
            'stream': False,
        }, timeout=120)
        return resp.json()['message']['content']

    raise ValueError(f"Provider desconhecido: {provider}")


def _map_llm(signals: list[RawSignal], sigla_flat: dict, llm_cfg: dict) -> list[Optional[tuple[str, int]]]:
    results: list[Optional[tuple[str, int]]] = [None] * len(signals)
    for batch_start in range(0, len(signals), _BATCH_SIZE):
        batch = signals[batch_start:batch_start + _BATCH_SIZE]
        types = [s.signal_type for s in batch]
        dominant = max(set(types), key=types.count) if types else None
        sigla_text = _build_sigla_text(sigla_flat, dominant)
        signal_text = '\n'.join(
            f"{i} | {s.utr_id} | {s.description} | {s.signal_type}"
            for i, s in enumerate(batch)
        )
        prompt = _PROMPT.format(sigla_list=sigla_text, signal_list=signal_text)
        try:
            raw = _call_llm(prompt, llm_cfg)
            parsed = _parse_llm_json(raw, len(batch))
            for item in parsed:
                idx = item.get('idx', 0)
                sigla = item.get('sigla')
                conf = int(item.get('confidence') or 0)
                if sigla and sigla in sigla_flat and 0 <= idx < len(batch):
                    results[batch_start + idx] = (sigla, conf)
        except Exception as e:
            log.warning("LLM batch %d falhou: %s", batch_start, e)
    return results

# ─── Mapeamento principal ─────────────────────────────────────────────────────

def _label(conf: int) -> str:
    if conf >= 90:
        return 'ALTA'
    if conf >= 70:
        return 'MÉDIA'
    if conf > 0:
        return 'BAIXA'
    return 'SEM'


def map_signals(
    raw_signals: list[RawSignal],
    protocol: str = 'dnp3',
    llm_cfg: Optional[dict] = None,
) -> list[MappedSignal]:
    """
    Mapeia sinais brutos para SIGLAs ADMS.
    llm_cfg = {'provider': 'gemini'|'groq'|'ollama', 'model': ..., 'api_key': ...}
    """
    sigla_flat = _load_sigla_flat(protocol)
    heuristic: list[Optional[tuple[str, int, str]]] = []
    unmatched_idx: list[int] = []

    for i, sig in enumerate(raw_signals):
        match = _token_match(sig.utr_id, sigla_flat)
        if match:
            heuristic.append((match[0], match[1], 'token'))
        else:
            match2 = _desc_match(sig.description, sigla_flat)
            if match2:
                heuristic.append((match2[0], match2[1], 'desc'))
            else:
                heuristic.append(None)
                unmatched_idx.append(i)

    if llm_cfg and unmatched_idx:
        unmatched = [raw_signals[i] for i in unmatched_idx]
        llm_res = _map_llm(unmatched, sigla_flat, llm_cfg)
        for batch_i, orig_i in enumerate(unmatched_idx):
            if llm_res[batch_i]:
                sigla, conf = llm_res[batch_i]
                heuristic[orig_i] = (sigla, conf, 'llm')

    mapped: list[MappedSignal] = []
    for sig, match in zip(raw_signals, heuristic):
        ms = MappedSignal(
            utr_id=sig.utr_id, description=sig.description,
            dnp3_addr=sig.dnp3_addr, signal_type=sig.signal_type,
            module=sig.module,
        )
        if match:
            sigla, conf, method = match
            ms.sigla = sigla
            ms.sigla_desc = sigla_flat.get(sigla, {}).get('desc', '')
            ms.confidence = conf
            ms.confidence_label = _label(conf)
            ms.match_method = method
        mapped.append(ms)

    matched = sum(1 for m in mapped if m.sigla)
    log.info("map_signals: %d/%d mapeados (%d heurística, %d LLM)",
             matched, len(mapped),
             sum(1 for _, m in zip(raw_signals, heuristic) if m and m[2] != 'llm'),
             sum(1 for _, m in zip(raw_signals, heuristic) if m and m[2] == 'llm'))
    return mapped

# ─── Converte para lista resumida (entrada do generate_tdt_from_list) ─────────

def to_lista_resumida(mapped: list[MappedSignal], alias: str,
                      min_confidence: int = 60) -> dict:
    """
    Converte sinais mapeados para o formato de parse_points_list,
    para ser consumido por generate_tdt_from_list.
    """
    discrete: list[dict] = []
    analog:   list[dict] = []
    da:       list[dict] = []

    aor_suffix = 'Distr'   # default; pode ser refinado

    for m in mapped:
        if not m.sigla or m.confidence < min_confidence:
            continue
        module = m.module or 'XX'
        nome = f"{alias}_{module}_{module}_{m.sigla}"
        aor  = f"{alias} {aor_suffix}"
        addr = str(m.dnp3_addr) if m.dnp3_addr else ''

        if m.signal_type == 'analog':
            analog.append({'sigla': m.sigla, 'nome': nome,
                           'escala': '', 'indexDnp3': addr, 'aor': aor})
        elif m.signal_type == 'command':
            # comando: vai como discreto com o índice no campo comando
            discrete.append({'sigla': m.sigla, 'nome': nome,
                             'indexDnp3Entrada': '', 'indexDnp3Comando': addr,
                             'aor': aor})
        else:
            discrete.append({'sigla': m.sigla, 'nome': nome,
                             'indexDnp3Entrada': addr, 'indexDnp3Comando': '',
                             'aor': aor})

    return {'discrete': discrete, 'analog': analog,
            'discrete_analog': da, 'inputErrors': []}
