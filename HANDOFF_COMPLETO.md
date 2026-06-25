<div align="center">

# 🧠 HANDOFF COMPLETO — Gerador de TDTs ADMS
### Tudo que foi feito, cada erro superado, cada decisão. Para outro Claude assumir.

`Schneider EcoStruxure ADMS` · Distribuidora RGE · documento de transferência de conhecimento

</div>

---

## 0. Como ler este documento

Eu (Claude) construí esta ferramenta ao longo de muitas sessões com o usuário (engenheiro
de SCADA da RGE). Este arquivo é a **memória técnica completa**: o quê, o porquê, os erros,
os becos sem saída, e as dicas que só se aprende apanhando. Se você é outro Claude pegando
isto — **leia inteiro antes de mexer**. Há armadilhas que custaram horas.

**Localização:** `C:/Users/egnpo/Downloads/UI DE TDT/adms-tdt-generator` (NÃO a pasta antiga
`UI DE TDT/` solta). GitHub: `github.com/EduGomesNascimento/ADMS-TDT-GENERATOR`.

---

## 1. O que o programa faz

Gera **TDTs** (Telemetry Data Template) — planilhas `.xlsx` que o **ADMS da Schneider**
importa para configurar a telemetria (sinais DNP3/IEC104 de subestações). Dois caminhos:

1. **Assistente por equipamento** (wizard): usuário escolhe o tipo (alimentador, TR, BC,
   LT, barra, TSA), preenche alias/módulo, marca sinais → gera a TDT.
2. **Importação de lista de pontos**: usuário sobe a lista bruta da subestação → a
   ferramenta mapeia descrição→SIGLA e gera a TDT.

**Filosofia central:** a ferramenta é um **clonador inteligente**. Ela NÃO inventa o formato
— ela **clona** uma TDT real (template), preservando 100% da estrutura/estilos/validações, e
só substitui os campos que dependem da identidade do equipamento.

---

## 2. Arquitetura

```
adms-tdt-generator/
├─ backend/        Python FastAPI (porta 8077)
│  ├─ app.py            endpoints REST
│  ├─ tdt_engine.py     CLONA o template, substitui tokens, gera a TDT
│  ├─ ai_mapper.py      motor de regras (lê listas não-padrão, mapeia → SIGLA)
│  ├─ excel_native.py   re-save via MS Excel COM (CRÍTICO — ver §4)
│  ├─ probability_report.py   relatório colorido por confiança
│  ├─ build_*.py        scripts que constroem data/ (catalog, indices, templates)
│  └─ data/             catalog.json, sigla_index.json, templates .xlsx, padrao_adms.json
├─ frontend/       React + Vite + TS + Tailwind (porta 5180, proxy /api)
│  └─ src/components/RawImportPanel.tsx   tela de importação + revisão
├─ ABRIR.BAT / abrir.ps1   sobe backend+frontend, abre Chrome --app
└─ chave_groq.txt / chave_gemini.txt   (legado — IA agora é local)
```

**Stack:** FastAPI + openpyxl (Python) / React+Vite (front). Decisão: web app local, não
desktop nativo — roda em qualquer PC com Python+Node, abre como "app" no Chrome/Edge.

---

## 3. O FORMATO DA TDT (o que descobri engenharia-reversa)

A TDT é um `.xlsx` com **3–4 abas**:
- `DMSMatchingTemplateInfo` — enums/validações fixas do ADMS (279×178). **NÃO mexer**, só
  preservar. É a "tabela de domínios" dos dropdowns.
- `DNP3_DiscreteSignals` (43 colunas)
- `DNP3_AnalogSignals` (61 colunas)
- `DNP3_DiscreteAnalog` (48 colunas) — só o TAP do comutador mora aqui.

**Estrutura das colunas (descoberta crucial):**
- **Linha 1:** título do grupo ("Signal Details", "Remote Points").
- **Linha 2:** tipo do objeto ("DSIGNAL", "REMOTEPOINT", "APROCESSING").
- **Linha 3:** IDs técnicos (`IDOBJ_NAME`, `REMOTEPOINT_TYPE`, `REMOTEINPUTDNP3_INCOORDS`…).
- **Linha 4:** rótulos humanos ("Signal Name", "Input Coordinates"…).
- **Linha 5+:** os dados.

`HEADER_ROWS = 4`. Dados começam na linha 5.

**Nome do sinal:** `{ALIAS}_{MÓDULO}_{DEVICE}_{SUFIXO}`. Ex.: `FWB_AL13_52-13_50F1`.
O SUFIXO é a **SIGLA** (o tipo do sinal). Ex.: `50F1` = sobrecorrente fase estágio 1.

**Tokenização (o coração do clonador):** capturei TDTs reais e troquei os tokens de
identidade por placeholders: `<<ALIAS>>`, `<<MODULE>>`, `<<DEVICE>>`, `<<PREFIX>>` (=
ALIAS_MÓDULO_DEVICE), `<<N>>` (número do trafo). Cada SIGLA vira uma "linha-template"
tokenizada guardada em `sigla_index.json`. Na geração, substituo os placeholders pelos
valores reais. **Round-trip da FWB original deu ZERO diffs de valor.**

**Campos identity-dependent** (os únicos que mudam por equipamento): Signal Name, Signal/
Remote Point Custom ID, Remote Point Name, Input Coordinates, Device Mapping, Remote Unit
(`UTR_{alias}_1`), Signal AOR Group (`{alias} Distr/Trans`). O resto é FIXO por SIGLA.

**Detalhe que pega:** o **Signal Custom ID** sai SEMPRE vazio (None) — o ADMS gera o GUID.
Se você deixar o GUID do template vazar, dá conflito. Limpe `idx_signal_custom`.

---

## 4. 🔴 O MAIOR ERRO: "Invalid TDI file format" (LEIA ISTO)

**Sintoma:** o ADMS recusava o arquivo: *"Invalid TDI file format. Please use official
Telemetry Data Template document and MS Excel or save file in MS Excel prior to import."*

**Causa-raiz (custou descobrir):** o **openpyxl grava o OOXML em formato não-canônico**.
Comparando o ZIP/XML de uma TDT oficial vs uma do openpyxl:
- comentários em `xl/comments/comment1.xml` (openpyxl) vs `xl/comments1.xml` (Excel real)
- openpyxl **não gera** `xl/sharedStrings.xml` (usa strings inline)
- worksheets sem declaração XML, sem `xr:uid`, sem `mc:Ignorable`

O parser TDI do ADMS é **estrito** e recusa.

**Solução (a mensagem do ADMS já dizia):** re-salvar pelo **MS Excel real via COM**
(`excel_native.py`, pywin32, `DispatchEx`, `FileFormat=51`). O Excel reescreve o pacote no
formato canônico. Roda numa **thread própria com CoInitialize** + lock global. Custo ~9s
(cold-start do Excel). **Exige MS Excel instalado** (todos os usuários têm — usam ADMS).

```python
# excel_native.resave_native(bytes) — abre no Excel, redimensiona tabelas, re-salva
# Ligado em tdt_engine._finalize(bytes, native=True), no fim de generate_tdt*.
```

**DICA pro próximo Claude:** SEMPRE valide a TDT gerada abrindo o ZIP e checando
`sharedStrings.xml` + `comments1.xml` (caminho nativo). Se não tiver, o ADMS recusa.
`/api/health` expõe `excelNative` (bool). Se faltar Excel, a UI avisa.

**Confirmado em campo:** depois do fix, o relatório do ADMS disse *"No errors or warnings
while parsing DNP3_DiscreteSignals/AnalogSignals"*. 🎉

---

## 5. FORMATAÇÃO (como preservei e destaquei)

### 5.1 Faixa banded (tabelas do Excel)
Cada aba tem uma **Tabela do Excel** (`MyTableStyle`, faixa verde, `rowStripes=1`). O `ref`
da tabela é fixo no template (ex.: `A4:AQ43`). **Bug de campo:** ao escrever mais linhas que
o template, elas ficavam FORA da tabela → perdiam a faixa.
**Fix:** `_fit_table(ws, last_row)` estende o `ref` dinamicamente (`A4:AQ{last_row}`). E o
Excel COM (`_resize_all_tables`) redimensiona o ListObject de verdade (openpyxl não consegue).

### 5.2 Estilo das novas linhas
Copio o `_style` da linha-modelo (linha 5) para cada linha nova: `cell._style = copy(styles[c])`.
Fonte, borda, preenchimento, formato numérico — idênticos.

### 5.3 Destaque de incertos (pedido do usuário: "só a cor da linha")
Sinais não-ALTA saem com fundo **amarelo claro** (`FFF2CC`): `cell.fill = _UNCERTAIN_FILL`,
aplicado DEPOIS do `_style`. **Só a cor muda** — formatação/fórmulas/validações preservadas.
Os nomes a destacar vêm em `parsed['uncertain']` (set), montado em `to_lista_resumida`.

### 5.4 A aba DiscreteAnalog (outro erro)
O `build_tap.py` antigo criou a DiscreteAnalog copiando a AnalogSignals e cortando colunas →
ficou com 61 colunas e **marcadores de grupo errados** (L2 dizia REMOTEPOINT na coluna AK,
mas as colunas técnicas estavam em X). O ADMS reclamava "IDOBJ_NAME column for REMOTEPOINT
missing". **Fix:** `build_template_from_ura.py` reconstruiu o template inteiro a partir de uma
TDT real válida (`TDT_DNP3_UTR_URA_TR2.xlsx`) — DiscreteAnalog correto com 48 colunas.
**Defensivo:** quando não há TAP, removo a aba DiscreteAnalog inteira (`del wb[da_sheet]`).
**NÃO use COM cross-workbook copy via ActiveSheet** — corrompe (apaga DiscreteSignals).

---

## 6. O MOTOR DE REGRAS (mapeamento sem IA)

`ai_mapper.py`. Mapeia DESCRIÇÃO/NOME → SIGLA ADMS por **camadas determinísticas**, da mais
forte à mais fraca. Regra de ouro: **só camadas exatas marcam ALTA**.

| # | Camada | Função | Conf. |
|---|---|---|---|
| 1 | Token | código ADMS embutido no nome (`_50F1`) | ALTA 100 |
| 2 | Semântico | medições (`Tensão→VAB`, `Corrente→IA`) regex | ALTA 92 |
| 3 | Base oficial | descrição = Pontos Padrão v1 | ALTA 94 |
| 4 | Descrição exata | descrição = índice | ALTA 96 |
| 5 | Proteção | palavra-chave→função (`subtensão→27`, `religamento→79`) | MÉDIA 82 |
| 6 | Fuzzy ANSI | código ANSI + sobreposição de tokens | MÉDIA/BAIXA ≤78 |
| 7 | IA (opcional) | só o que sobrou; LOCAL (Ollama) | MÉDIA ≤85 |
| — | sem match | SEM (revisar) | — |

Regras completas em `REGRAS.txt`. **Guarda anti-falso-positivo:** rejeita SIGLA de 1
caractere (A/B/D/F) em token/fuzzy — eram falsos positivos que sujavam o resultado.

**Ranker de candidatos** (`_rank_candidates`): para CADA sinal, lista os top-5 SIGLAs
possíveis com score (token + ANSI). Aparece na revisão como chips clicáveis — o usuário
escolhe o certo, mesmo nos sem match firme.

---

## 7. NORMALIZAÇÃO (`_norm`)

Aplicada a TODA descrição antes de comparar (lista E base — só assim se encontram):
- MAIÚSCULAS, sem acento, espaços colapsados.
- **Expansão de abreviações** (word-boundary): `DISJ→DISJUNTOR`, `SEC→SECCIONADORA`,
  `TEMP→TEMPERATURA`, `POT→POTENCIA`, `CORR→CORRENTE`, etc. (lista em REGRAS.txt).
- **Códigos ANSI** extraídos separados (`(?<![0-9])([2-9][0-9])(?![0-9])`) — pesam mais.

Plano detalhado de normalização para TF-IDF/embedding em `NORMALIZACAO_TFIDF_EMBEDDING.md`.

---

## 8. LER LISTAS NÃO-PADRÃO (2 formatos + o bug crítico)

`parse_raw_excel`. Detecta o formato:
1. **UTR-ID** (URA): coluna com mais valores no padrão `ALIAS_TOKENS` (`_find_utr_column`).
2. **Colunas estruturadas** (FredW): cabeçalho com `Módulo`+`Tipo`+`Descrição` (`_extract_structured`).
   Tipo A/D/C → analog/discreto/comando. Mapeamento 100% por descrição.

### 🔴 BUG CRÍTICO que o usuário pegou: NÃO lia a lista inteira
Tinha 3 falhas que **pulavam sinais**:
1. Processava abas "preferidas" **OU** "outras", nunca as duas → pulava abas.
2. Dedup por `utr_id` (não-único no formato estruturado — `MODULE_MODULE` igual p/ todos) →
   dropava sinais do mesmo módulo.
3. `_scan_grid` cortava em 3000 linhas / 30 colunas → truncava.

**Fix:** lê TODAS as abas não-lixo; dedup só de linha 100% idêntica (tupla completa);
limites altos (200k linhas / 60 col). FredW foi de **1222 → 1429 sinais** (+207 que sumiam!).
**DICA:** sempre conte `sinais extraídos vs linhas reais por aba` ao mexer no parser.

Abas-lixo descartadas (`_SKIP_SHEETS`): `Capa, Calculados, Slot, Saca, Sumário, Config…`.

---

## 9. A IA (e por que virou só LOCAL)

Testei **3 provedores grátis**:
- **Groq** (Llama 3.3 70B): melhor + rápido, MAS tier grátis = **100K tokens/DIA**. Lista
  grande estoura no meio.
- **Gemini 2.5 Flash**: funciona via SDK novo `google-genai`. **NÃO use `gemini-2.0-flash`**
  (cota ZERO em muitos projetos → 429). Use `gemini-2.5-flash` (250K tok/min, 20 req/dia) ou
  `gemini-3.1-flash-lite` (500/dia).
- **Gemma 4 31B** (via Gemini API): instável, deu 500 INTERNAL. Evitar.
- **Ollama local**: 100% offline, mas LENTO sem GPU (qwen2.5:3b ~90s/8 sinais, e fraco).

**Decisão final do usuário:** IA **só LOCAL/offline** (removi Groq/Gemini da UI). O padrão é
**"Sem IA — motor de regras"**. A ferramenta deve rodar sem internet.

**Truques que viabilizaram o LLM no tier grátis:**
- `_candidates` (índice invertido): manda só os ~160 SIGLAs candidatos no prompt (não os
  2994) → cabe em 12K TPM do Groq.
- Retry com backoff (8/16/24s) em 429/413/rate.
- **Teto de confiança 85 no LLM:** ele acerta a FUNÇÃO mas chuta o VÃO (52-10 vs 52-2 real),
  e devolve 100% até no chute. Por isso LLM nunca marca ALTA.

---

## 10. O QUE IDENTIFIQUEI NOS TESTES (achados de ouro)

- **Validação contra TDT real (Frederico/FWB):** o LLM AUMENTA cobertura mas ERRA precisão em
  sinais bay-específicos (numeração de vão). Só o determinístico é confiável → ALTA.
- **Chave Gemini formato `AQ.`:** autentica mas o `google.generativeai` (SDK velho) está
  deprecado; use `google-genai` (novo).
- **Bug de coordenadas:** `to_lista_resumida` usava chaves `indexDnp3*` mas o engine lê
  `inCoord/outCoord/escala` → os endereços DNP3 SE PERDIAM na TDT. Fix: `_add_to_lista` com
  as chaves certas.
- **Lista FredW pura-descrição** é o caso MAIS difícil (sem código nem UTR-ID): determinístico
  ~45%, com LLM ~65-79%, resto = revisão humana. É o limite natural.
- **OpenClaw** (que o colega citou) = agente de código que usa Ollama; a parte de IA já temos.

---

## 11. O QUE GEREI EXTRA (e por quê)

- `REGRAS.txt` — todas as regras do motor (auditoria).
- `PLANO_MOTOR_REGRAS.md` — pipeline, normalização, TF-IDF/embedding, IA local, .exe único.
- `NORMALIZACAO_TFIDF_EMBEDDING.md` — passo a passo da normalização.
- `SPEC_REVISAO_SINAIS.md` — spec da tela de revisão.
- `COMPARATIVO_IA.xlsx` — prós/contras IA paga + comparação de métodos.
- `FredW_PROBABILIDADES.xlsx` / `FredW_TDT_gerada.xlsx` — saídas de teste reais (no Downloads).
- `build_expand_signals.py` — expandiu a gama de sinais por módulo (639+55 padrão ADMS por
  tipo). catalog 832KB→2.7MB.

---

## 12. NECESSÁRIO vs DESNECESSÁRIO (aprendi apanhando)

**Necessário (não remova):**
- `excel_native.py` (re-save nativo) — SEM ISSO O ADMS RECUSA TUDO.
- `data/*.json` + templates — NÃO recriáveis sem a base de 98MB. Distribua junto.
- `_fit_table` + estilo por linha — senão a formatação quebra.
- Limpar Signal Custom ID (None).
- Ler TODAS as abas (o bug §8).

**Becos sem saída / desnecessário (não repita):**
- **`build_fix_da.py`** (COM cross-workbook copy via ActiveSheet) — CORROMPE o template.
  Deletado. Use `build_template_from_ura.py`.
- **IA paga** — não compensa (ganho pequeno sobre Groq grátis, e o chefe não autoriza).
- **Gemma via API** — instável (500).
- **`gemini-2.0-flash`** — cota zero.
- **Ollama em CPU sem GPU** — lento demais para listas grandes.
- Tentar fazer o openpyxl gerar OOXML "compatível" na mão — fútil; use o Excel COM.

---

## 13. ARMADILHAS DE AMBIENTE (Windows)

- `pip install pywin32` basta para o COM (não precisa do postinstall em geral).
- O backend roda COM numa thread própria com `CoInitialize`/`CoUninitialize` + lock global
  (`_EXCEL_LOCK`) — senão dá conflito em chamadas concorrentes.
- Bash do harness no Windows: `&` em background NÃO persiste; use `run_in_background: true`.
- `git commit -m @'...'@` é sintaxe PowerShell; no Bash use `-F arquivo` ou aspas normais.
- openpyxl: `ws.tables` é dict; iterar `.items()` às vezes devolve string — use `ws.tables[name]`.
- Mata processo na porta: `Get-NetTCPConnection -LocalPort 8077 | Stop-Process`.

---

## 14. ESTADO ATUAL E BACKLOG

**Feito e funcional:**
- Clonador (wizard + import), re-save nativo, motor de regras (6 camadas), normalização,
  ler-tudo, tela de revisão com candidatos, destaque amarelo, gama expandida por módulo,
  IA local opcional, relatório de probabilidades.

**Backlog (documentado nas specs):**
- TF-IDF offline (`idf.json`) — desempate de candidatos sem IA.
- Embedding local (MiniLM) — camada semântica reproduzível, offline.
- Questionário final de dúvidas (checkbox antes de gerar).
- Aprendizado (`aprendizado.json`) — salvar correções do usuário → camada 0.
- Empacotamento PyInstaller `--onefile` (backend+frontend+dados+IA num .exe).
- Agrupar sinais por categoria no assistente (lista grande agora).

---

## 15. COMO RODAR / DISTRIBUIR

- **Rodar:** `ABRIR.BAT` (instala deps, sobe backend 8077 + frontend 5180, abre Chrome --app).
- **Distribuir:** zip `TDT-ADMS-completo.zip` (exclui node_modules/.venv/.git). Requer Python
  3.11+ e Node 18+ no destino + MS Excel.
- **Chave IA:** `chave_groq.txt`/`chave_gemini.txt` (legado — IA agora é Ollama local).

---

> **Mensagem para o próximo Claude:** o usuário é técnico, direto e cobra precisão ("NÃO
> ERRE"). Ele valida em campo (importa no ADMS real). Sempre teste contra dados reais, conte
> os sinais, e seja HONESTO sobre limites (ele prefere "41% confiável" a "52% com lixo"). O
> maior valor é o **determinístico confiável** — a IA é só reforço. Boa sorte. 🤝
