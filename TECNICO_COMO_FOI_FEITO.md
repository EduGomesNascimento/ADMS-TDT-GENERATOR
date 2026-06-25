# TÉCNICO — Como foi construído o Gerador de TDTs ADMS
> Para outro Claude assumir. Cada decisão de implementação explicada com código real.

---

## 1. O problema central e a abordagem de clonagem

O ADMS valida a TDT contra um schema XML interno muito estrito. A tentativa inicial de montar
o `.xlsx` do zero com openpyxl resultava em rejeição silenciosa. A solução foi **clonar uma
TDT real** — capturar cada linha da base como template tokenizado, e na geração substituir
apenas os tokens de identidade.

### 1.1 Estrutura interna do .xlsx (ZIP)

O `.xlsx` é um ZIP. Comparando o ZIP de uma TDT oficial contra o do openpyxl:

```
TDT oficial                         openpyxl
xl/comments1.xml          ≠         xl/comments/comment1.xml
xl/sharedStrings.xml      ≠         (ausente — strings inline)
worksheets/sheet1.xml     ≠         sem xr:uid, sem mc:Ignorable
```

O parser TDI do ADMS é estrito nessa estrutura. Não há como corrigir na mão.

### 1.2 sigla_index.json — o coração do clonador

`build_index.py` varreu a base real (`Export_base_Full__27_fev_2026.xlsx`, 98MB) e para cada
SIGLA única capturou a **linha inteira** de uma TDT real, trocando os valores de identidade
por tokens:

```python
# Tokens substituídos em cada campo:
PREFIX  = f"{alias}_{module}_{device}"   # <<PREFIX>>
ALIAS   = alias                           # <<ALIAS>>
MODULE  = module                          # <<MODULE>>
DEVICE  = device                          # <<DEVICE>>
N       = transformer_number              # <<N>>
```

Resultado: `sigla_index.json` com 2994 sinais discretos + 447 analógicos, cada um uma
lista de 43/61/48 valores pré-tokenizados. Na geração, basta `str.replace()` nos valores.

```python
# tdt_engine.py — geração de uma linha
def _write_row(ws, row_idx, cols, vals, prefix, alias, module, device):
    for c, val in enumerate(vals, 1):
        v = str(val) if val is not None else ""
        v = (v.replace("<<PREFIX>>", prefix)
              .replace("<<ALIAS>>",  alias)
              .replace("<<MODULE>>", module)
              .replace("<<DEVICE>>", device))
        ws.cell(row=row_idx, column=c).value = v or None
```

---

## 2. Re-save via Excel COM (excel_native.py)

**Único ponto onde o ADMS aceita o arquivo.** Sem isso, 100% de rejeição.

```python
import pythoncom, win32com.client, threading, tempfile, os
_EXCEL_LOCK = threading.Lock()

def resave_native(xlsx_bytes: bytes) -> bytes:
    with _EXCEL_LOCK:
        pythoncom.CoInitialize()
        try:
            xl = win32com.client.DispatchEx("Excel.Application")
            xl.Visible = False; xl.DisplayAlerts = False
            with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
                f.write(xlsx_bytes); tmp = f.name
            wb = xl.Workbooks.Open(tmp)
            # Redimensiona ListObjects (tabelas) para cobrir todas as linhas
            for ws in wb.Worksheets:
                for tbl in ws.ListObjects:
                    last = ws.UsedRange.Rows.Count
                    tbl.Resize(ws.Range(f"A4:{ _col_letter(tbl.Range.Columns.Count) }{last}"))
            wb.Save()
            wb.Close(False)
            xl.Quit()
            with open(tmp, "rb") as f:
                result = f.read()
            return result
        finally:
            pythoncom.CoUninitialize()
            try: os.unlink(tmp)
            except: pass
```

**Por que thread + lock:** `CoInitialize` é por-thread no Windows. Chamadas concorrentes
ao mesmo processo Excel causam conflito de COM apartment. O lock serializa.

**Por que `DispatchEx` e não `Dispatch`:** `Dispatch` reaproveita uma instância do Excel
que pode estar aberta pelo usuário. `DispatchEx` força instância nova (headless).

**Por que redimensionar ListObjects:** openpyxl adiciona linhas mas não atualiza o `ref`
da tabela Excel. O Excel re-salva com a tabela curta → linhas fora da tabela perdem
faixa/estilo. `tbl.Resize()` corrige.

---

## 3. Motor de regras (ai_mapper.py) — implementação camada por camada

### 3.1 Normalização `_norm()`

Aplicada a TODA string antes de qualquer comparação (lista E base):

```python
import re, unicodedata

_ABBREV = {
    "DISJ": "DISJUNTOR", "SEC": "SECCIONADORA", "SECC": "SECCIONADORA",
    "TEMP": "TEMPERATURA", "POT": "POTENCIA", "CORR": "CORRENTE",
    "TENS": "TENSAO", "FREQ": "FREQUENCIA", "PROT": "PROTECAO",
    "CMD": "COMANDO", "IND": "INDICACAO", "MED": "MEDICAO",
}
_ABBREV_RE = re.compile(
    r'\b(' + '|'.join(re.escape(k) for k in _ABBREV) + r')\b'
)

def _norm(s: str) -> str:
    s = unicodedata.normalize("NFD", s.upper())
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")  # remove acentos
    s = _ABBREV_RE.sub(lambda m: _ABBREV[m.group()], s)
    return re.sub(r'\s+', ' ', s).strip()
```

Word-boundary (`\b`) é essencial: sem ele, "SECC" dentro de "SECCIONAL" vira
"SECCIONADORAL" — testado e confirmado como problema real.

### 3.2 Camada 1 — Token (conf 100, ALTA)

Procura sufixo ADMS embutido no próprio nome do ponto (`_50F1`, `_87T`, etc.):

```python
_TOKEN_RE = re.compile(r'_([A-Z0-9]{2,8})$')  # no final do nome

def _token_match(name: str, sigla_flat: set) -> str | None:
    m = _TOKEN_RE.search(_norm(name))
    if m:
        cand = m.group(1)
        if len(cand) > 1 and cand in sigla_flat:  # guard: rejeita 1 char
            return cand
    return None
```

Guard `len > 1`: sem ele, nomes terminando em `_A` ou `_B` casavam `A`/`B` (falso positivo
confirmado em testes com lista FredW).

### 3.3 Camada 2 — Semântico (conf 92, ALTA)

Dicionário de medições físicas → SIGLA. Ordem importa (REATIVA antes de ATIVA):

```python
_SEMANTIC_RULES = [
    # Tensões
    (re.compile(r'TENSAO.*(AB|FAB|FASE.AB)'), "VAB"),
    (re.compile(r'TENSAO.*(BC|FBC)'),           "VBC"),
    (re.compile(r'TENSAO.*(CA|FCA)'),           "VCA"),
    (re.compile(r'TENSAO.*(A|FASE.A)'),         "VA"),
    # Correntes
    (re.compile(r'CORRENTE.*(NEUTRO|N\b)'),     "IN"),
    (re.compile(r'CORRENTE.*(A|FASE.A)'),       "IA"),
    # Potências — REATIVA antes de ATIVA (evita Q→P)
    (re.compile(r'POT(ENCIA)?.*(REAT|Q\b)'),   "Q"),
    (re.compile(r'POT(ENCIA)?.*(ATIVA|P\b)'),  "P"),
    (re.compile(r'POT(ENCIA)?.*(APAR|S\b)'),   "S"),
    # Outros
    (re.compile(r'TEMPERATURA'),                "TOLE"),
    (re.compile(r'TAP|COMUTADOR'),              "TAP"),
    (re.compile(r'FREQUENCIA'),                 "F"),
    (re.compile(r'FATOR.POTENCIA'),             "FP"),
]

def _semantic_match(desc: str) -> str | None:
    n = _norm(desc)
    for pattern, sigla in _SEMANTIC_RULES:
        if pattern.search(n):
            return sigla
    return None
```

**Por que REATIVA antes de ATIVA:** "Potência Reativa" normaliza para "POTENCIA REAT" —
se o padrão de ATIVA (`POT.*ATIVA`) vier primeiro e usar `.*`, pode casar antes.
Testado: ordem errada dava Q→P em 100% dos casos.

### 3.4 Camada 5 — Proteção (conf 82, MÉDIA)

```python
_PROT_RULES = [
    (re.compile(r'SUBTENSAO|27'),              "27_T"),
    (re.compile(r'SOBRETENSAO|59'),            "59"),
    (re.compile(r'RELIGAMENTO|79'),            "79"),
    (re.compile(r'BLOQUEIO.*(GERAL|86)'),      "86"),
    (re.compile(r'DIFERENCIAL.*TRAFO|87T'),    "87T"),
    (re.compile(r'DIFERENCIAL.*BARRA|87B'),    "87B"),
    (re.compile(r'FALHA.COMUNIC'),             "FCOM"),
    (re.compile(r'SF.?6'),                     "SF6A"),
    (re.compile(r'MOLA'),                      "MOLA"),
    (re.compile(r'50.?BF|FALHA.DISJ'),         "50BF"),
]
```

Fica em MÉDIA (82) porque a função está certa mas o VÃO (bay-specific numbering) não.
Ex.: "Proteção Diferencial" acerta `87T` mas não sabe se é `87T` do TR1 ou TR2.

### 3.5 Camada 6 — Fuzzy ANSI (conf ≤78, MÉDIA)

Extrai código ANSI numérico do nome/descrição e calcula Jaccard de tokens:

```python
_ANSI_RE = re.compile(r'(?<![0-9])([2-9][0-9])(?![0-9])')

def _fuzzy_match(desc: str, inv_index: dict) -> tuple[str, int] | None:
    norm = _norm(desc)
    ansi_codes = set(_ANSI_RE.findall(norm))
    tokens = set(norm.split())
    best_sigla, best_score = None, 0
    for sigla, sigla_tokens in inv_index.items():
        if len(sigla) <= 1: continue  # guard
        # ANSI bônus: +30 se o código ANSI da SIGLA está na descrição
        ansi_bonus = 30 if sigla_tokens & ansi_codes else 0
        jaccard = len(tokens & sigla_tokens) / len(tokens | sigla_tokens) * 100
        score = min(int(jaccard + ansi_bonus), 78)  # teto 78 = MÉDIA
        if score > best_score:
            best_score, best_sigla = score, sigla
    return (best_sigla, best_score) if best_score > 20 else None
```

Teto de 78 é hard: garante que fuzzy nunca alcança ALTA (90+).

### 3.6 Ranker de candidatos `_rank_candidates`

Para CADA sinal, independente de qual camada casou, computa top-5:

```python
def _rank_candidates(sig, sigla_flat, inv, det, top=6):
    norm = _norm(sig.description)
    tokens = set(norm.split())
    ansi = set(_ANSI_RE.findall(norm))
    scores = {}
    for sigla in sigla_flat:
        if len(sigla) <= 1: continue
        st = inv.get(sigla, set())
        j = len(tokens & st) / max(len(tokens | st), 1) * 100
        bonus = 30 if st & ansi else 0
        scores[sigla] = min(int(j + bonus), 99)
    ranked = sorted(scores.items(), key=lambda x: -x[1])[:top]
    return [{"sigla": s, "score": sc, "desc": det.get(s, "")} for s, sc in ranked]
```

Esses candidatos chegam no frontend como chips clicáveis (mesmo para sinais SEM match firme).

---

## 4. Leitura de listas não-padrão (parse_raw_excel)

### 4.1 Detecção de formato

```python
def _find_utr_column(ws) -> int | None:
    # Procura coluna com >= 10 valores no padrão ALIAS_TOKENS
    UTR_RE = re.compile(r'^[A-Z]{2,6}_[A-Z0-9_]{2,}$')
    for col in range(1, min(ws.max_column+1, 60)):
        hits = sum(1 for r in range(1, min(ws.max_row+1, 200))
                   if UTR_RE.match(str(ws.cell(r, col).value or "")))
        if hits >= 10:
            return col
    return None
```

Se encontra coluna UTR-ID → formato URA. Senão → busca cabeçalho estruturado
(Módulo/Tipo/Descrição) via `_extract_structured`.

### 4.2 O bug crítico dos 207 sinais perdidos

**Problema 1 — abas:**
```python
# ANTES (errado): escolhia preferred OU others, nunca os dois
preferred = [s for s in wb.sheetnames if s in PREFERRED_SHEETS]
sheets = preferred or [s for s in wb.sheetnames if not _SKIP_SHEETS.search(s)]

# DEPOIS (correto): lê TODAS as não-lixo
sheets = [n for n in wb.sheetnames if not _SKIP_SHEETS.search(n)]
```

**Problema 2 — dedup por utr_id (não-único):**
```python
# ANTES: dedup por utr_id → no formato estruturado, utr_id = "MODULE_MODULE" (igual p/ todos)
if sig.utr_id in seen: continue

# DEPOIS: dedup por tupla completa
key = (sn, sig.module, sig.signal_type, sig.description, sig.dnp3_addr, sig.utr_id)
if key in seen: continue
```

**Problema 3 — limites de scan:**
```python
# ANTES
_scan_grid(ws, max_rows=3000, max_cols=30)
# DEPOIS
_scan_grid(ws, max_rows=200000, max_cols=60)
```

FredW: 1222 → 1429 sinais (+207). Esses sinais simplesmente sumiam sem aviso.

---

## 5. Gama de sinais — build_expand_signals.py

Cada tipo de equipamento tinha os sinais de UMA TDT. O script:

1. Carrega `catalog.json` (tipos), `sigla_index.json` (índice completo), `padrao_adms.json` (filtro oficial).
2. Para cada tipo, para cada classe (discrete/analog):
   - `have = {s["suffix"] for s in tipo["signals"][klass]}`
   - Itera o índice: se sufixo não está em `have`, não é 1 char, e está no padrão oficial → adiciona com grupo "Padrão ADMS".

```python
for suffix, row in src.items():
    if suffix in have or len(suffix) <= 1 or suffix not in official:
        continue
    lst.append({
        "suffix": suffix,
        "description": row[2] if len(row) > 2 else suffix,
        "klass": klass,
        "group": "Padrão ADMS",
        "row": row,  # linha pré-tokenizada do índice
    })
```

**Idempotente:** `have` é recalculado a cada run; só adiciona o que falta.
**Filtro `padrao_adms.json`:** evita inchar com os 2994 sinais brutos — só os 639+55
do padrão oficial RGE entram. Resultado: catalog 832KB → 2.7MB.

---

## 6. Geração da TDT (tdt_engine.py)

### 6.1 Fluxo geral

```python
def generate_tdt_from_list(parsed: dict, alias: str, protocol: str,
                            native: bool = True) -> bytes:
    wb = load_workbook(TEMPLATE_PATH)          # clona o template
    uncertain = set(parsed.get("uncertain") or [])

    for sheet_name, signals in [
        ("DNP3_DiscreteSignals",  parsed["discrete"]),
        ("DNP3_AnalogSignals",    parsed["analog"]),
        ("DNP3_DiscreteAnalog",   parsed.get("discrete_analog", [])),
    ]:
        ws = wb[sheet_name]
        styles = _capture_row_styles(ws, 5)    # linha-modelo
        for i, sig in enumerate(signals):
            row_idx = 5 + i
            vals = _build_row(sig, sheet_name)  # substitui tokens
            _write_row(ws, row_idx, vals, styles)
            if sig["nome"] in uncertain:        # destaque amarelo
                _apply_fill(ws, row_idx, _UNCERTAIN_FILL)
        _fit_table(ws, 4 + len(signals))        # estende tabela

    if not parsed.get("discrete_analog"):
        del wb["DNP3_DiscreteAnalog"]           # remove aba vazia

    buf = BytesIO(); wb.save(buf)
    return _finalize(buf.getvalue(), native)    # re-save Excel COM
```

### 6.2 Captura e cópia de estilos

```python
from copy import copy

def _capture_row_styles(ws, row: int) -> dict:
    return {c: copy(ws.cell(row, c)._style)
            for c in range(1, ws.max_column + 1)}

def _apply_styles(ws, row_idx: int, styles: dict):
    for c, style in styles.items():
        ws.cell(row_idx, c)._style = copy(style)
```

`._style` é o objeto interno do openpyxl que carrega font/fill/border/number_format.
Copiar diretamente (sem usar `NamedStyle`) preserva tudo exatamente.

O destaque amarelo é aplicado DEPOIS dos estilos (sobrescreve só o fill):
```python
_UNCERTAIN_FILL = PatternFill("solid", fgColor="FFF2CC")

def _apply_fill(ws, row_idx: int, fill):
    for c in range(1, ws.max_column + 1):
        ws.cell(row_idx, c).fill = fill
```

### 6.3 Extensão da tabela

```python
from openpyxl.utils import get_column_letter

def _fit_table(ws, last_row: int):
    for name, tbl in list(ws.tables.items()):
        first_col = tbl.ref.split(":")[0]          # ex: "A4"
        ncols = tbl.headerRowCount                  # colunas da tabela
        last_col = get_column_letter(               # ex: "AQ"
            ws[first_col].column + ncols - 1)
        tbl.ref = f"{first_col}:{last_col}{last_row}"
```

O Excel COM re-faz o resize via `tbl.Resize()` para garantir que o ListObject
interno do Excel também está atualizado (openpyxl só atualiza o XML, não o objeto COM).

---

## 7. DiscreteAnalog — como foi corrigida

O `build_fix_da.py` (deletado) tentava copiar a sheet AnalogSignals via COM:
```python
# ERRADO — corrompia a workbook (apagava DNP3_DiscreteSignals)
xl.ActiveSheet.Copy(Before=wb.Worksheets("DNP3_AnalogSignals"))
```
COM cross-workbook copy muda o contexto ativo e a sheet de referência se perde.

**Fix correto (`build_template_from_ura.py`):** ler a TDT real linha a linha com openpyxl
e copiar célula por célula para o template novo:

```python
src = load_workbook("TDT_DNP3_UTR_URA_TR2.xlsx")
dst = load_workbook("reference_template.xlsx")

for sheet in ["DNP3_DiscreteSignals","DNP3_AnalogSignals","DNP3_DiscreteAnalog"]:
    ws_src = src[sheet]; ws_dst = dst[sheet]
    for r in range(1, 5):                          # só os 4 headers
        for c in range(1, ws_src.max_column + 1):
            ws_dst.cell(r, c).value = ws_src.cell(r, c).value
dst.save("reference_template.xlsx")
```

Resultado: DiscreteAnalog correto com 48 colunas, grupos L1/L2 corretos
(DSIGNAL / REMOTEPOINT / APROCESSING nas colunas certas).

---

## 8. Frontend — RawImportPanel.tsx (revisão de sinais)

### 8.1 Chips de candidatos

```tsx
// Cada sinal tem m.candidates: [{sigla, score, desc}]
{m.candidates.slice(0, 5).map(c => (
  <button
    key={c.sigla}
    title={c.desc}
    onClick={() => setSiglaOverride(m.id, c.sigla)}
    className={clsx(
      "px-1.5 py-0.5 rounded text-xs border",
      activeSigla === c.sigla
        ? "bg-blue-600 text-white border-blue-600"
        : "bg-zinc-800 text-zinc-300 border-zinc-600 hover:border-blue-400"
    )}
  >
    {c.sigla} <span className="text-zinc-500">{c.score}</span>
  </button>
))}
```

### 8.2 Cor semáforo por confiança

```tsx
const confColor = (conf: number | null) =>
  conf === null      ? "border-zinc-600"   // SEM
  : conf >= 90       ? "border-green-500"  // ALTA
  : conf >= 70       ? "border-yellow-400" // MÉDIA
  :                    "border-red-500";   // BAIXA
```

### 8.3 Geração dos sinais revisados

```tsx
async function genReviewed() {
  const signals = filtered.map(m => ({
    nome: siglaOverrides[m.id] ?? m.suggestedSigla,
    inCoord: m.inCoord, outCoord: m.outCoord, escala: m.escala,
    signalType: m.signalType, module: m.module, confidence: m.confidence,
  }));
  const res = await fetch("/api/raw/export_reviewed", {
    method: "POST",
    headers: {"Content-Type": "application/json"},
    body: JSON.stringify({ alias, protocol, signals }),
  });
  // download do blob...
}
```

---

## 9. Erros de produção e como foram superados

| Erro | Causa | Fix |
|---|---|---|
| `"Invalid TDI file format"` | openpyxl OOXML não-canônico | Excel COM re-save |
| `"IDOBJ_NAME column missing"` | DiscreteAnalog com 61 col erradas | Reconstruir do zero a partir de TDT real |
| LLM retorna 100% até no chute | Groq/Gemini sem calibração | Teto hard de 85 no LLM |
| `Gemini 429` | `gemini-2.0-flash` cota zero | Usar `gemini-2.5-flash` + SDK `google-genai` |
| `Groq 413` | Prompt com 2994 SIGLAs = 52K tokens | Índice invertido → manda só candidatos |
| 207 sinais sumindo | Lógica OR de abas + dedup errado | Ler todas + dedup tupla completa |
| `Tensão → FA` (tipo errado) | Regra semântica sem verificar tipo A/D | Verificação `sig.signal_type` antes de casar |
| Endereços DNP3 zerando na TDT | `to_lista_resumida` usava chaves `indexDnp3*` | Fix: `_add_to_lista` com chaves `inCoord/outCoord` |
| Tabela curta na TDT | openpyxl não atualiza `ref` da tabela | `_fit_table` + COM `tbl.Resize()` |
| `A`,`B` como SIGLA (falso positivo) | Token/fuzzy sem guard de comprimento | `if len(cand) <= 1: continue` |

---

## 10. O que foi desnecessário (não repita)

- `build_fix_da.py` — COM cross-workbook copy. CORROMPE. Deletado.
- Tentar corrigir OOXML do openpyxl manualmente (sharedStrings, comments) — cada versão
  do openpyxl muda os paths; inviável de manter.
- Groq/Gemini para produção — cota diária estoura em listas grandes; usuário pediu offline.
- Ollama `qwen2.5:3b` em CPU — 90s por 8 sinais, acurácia 1/8. Só válido como fallback.
- `padrao_adms.json` com os 2994 brutos — o catálogo ficava 15MB e o wizard inutilizável.
  Filtrar para os 639+55 oficiais reduziu para 2.7MB.

---

## 11. Estrutura de dados chave

### sigla_index.json
```json
{
  "DNP3_DiscreteSignals": {
    "50F1": ["<<PREFIX>>_50F1", null, "Sobrecorrente Fase Est1 Trip", ...],
    "87T":  ["<<PREFIX>>_87T",  null, "Diferencial Trafo Trip", ...]
  },
  "DNP3_AnalogSignals": {
    "VAB": ["<<PREFIX>>_VAB", null, "Tensão de Linha AB", ...]
  },
  "DNP3_DiscreteAnalog": {
    "TAP": ["<<PREFIX>>_TAP", null, "Posição TAP Comutador", ...]
  }
}
```
Cada lista tem 43/61/48 elementos (colunas da respectiva aba), pré-tokenizados.

### padrao_adms.json
```json
{
  "discrete": {"50F1": "Sobrecorrente Fase Est1", "87T": "Diferencial Trafo", ...},
  "analog":   {"VAB": "Tensão de Linha AB", "IA": "Corrente Fase A", ...}
}
```
639 sinais discretos + 55 analógicos. Fonte: `Pontos Padro ADMS_v1.xlsx` (RGE).

### catalog.json (após expansão)
```json
{
  "deviceTypes": [{
    "id": "alimentador",
    "signals": {
      "discrete": [
        {"suffix":"52","description":"Disjuntor","group":"Alimentador","row":[...]},
        {"suffix":"50F1","description":"Sobrecorrente...","group":"Padrão ADMS","row":[...]}
      ],
      "analog": [...]
    },
    "signalCount": {"discrete": 536, "analog": 50}
  }]
}
```

---

> **Para o próximo Claude:** os dados em `data/*.json` e `data/*.xlsx` NÃO são recriáveis
> sem a base de 98MB (`Export_base_Full__27_fev_2026.xlsx`) que está apenas em
> `C:/Users/egnpo/Downloads/UI DE TDT/`. Sempre distribua o zip COM esses arquivos.
> O HANDOFF_COMPLETO.md tem o contexto estratégico; este arquivo tem o técnico.
