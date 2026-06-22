<div align="center">

# 🧹 Normalização de Texto para TF-IDF + Embedding
### Plano passo a passo — do dado bruto ao vetor pronto para análise

`Gerador de TDTs ADMS` · documento de planejamento · revisar depois

</div>

---

## 🎯 Por que normalizar?

As descrições das listas de campo são **sujas e inconsistentes**. O mesmo sinal
aparece de mil jeitos:

| Como vem no campo | O que realmente é |
|---|---|
| `Tensão Barra AT P (06F1)` | Tensão fase (medição) |
| `TENS. BARRA A-B` | Tensão AB |
| `Disj. 06Q0 (AT: 52) - Abrir / Fechar` | Comando do disjuntor (52) |
| `27 - TRIP SUBTENSÃO` | Proteção 27 |

> **Sem normalizar**, o TF-IDF e o embedding tratam `Tensão`, `TENS.` e `tensao`
> como **palavras diferentes** → o matching erra. A normalização é o passo que
> faz o algoritmo "enxergar" que tudo isso é a mesma coisa.

**Regra de ouro:** o **mesmo `normalizar()`** roda nos dois lados — na descrição
do campo **E** na descrição da base ADMS. Só assim eles se encontram.

---

## 🗺️ Visão geral do fluxo

```
 texto bruto
     │
 ┌───▼────────────────────────────────────────────────┐
 │ ETAPA 1  Limpeza básica (caixa, acento, espaço)     │
 │ ETAPA 2  Expansão de abreviações (Disj→Disjuntor)   │
 │ ETAPA 3  Separar CÓDIGO (ANSI/nº) do TEXTO          │
 │ ETAPA 4  Tokenizar                                   │
 │ ETAPA 5  Remover stopwords de domínio (ruído)       │
 │ ETAPA 6  Stemming leve (plural/sufixo)               │
 │ ETAPA 7  Montar o "documento limpo"                  │
 └───┬─────────────────────────┬──────────────────────┘
     │                         │
 ┌───▼──────────┐        ┌─────▼──────────┐
 │  TF-IDF      │        │   EMBEDDING    │
 │ (palavras)   │        │ (significado)  │
 └──────────────┘        └────────────────┘
```

---

## 🔧 Passo a passo (com exemplos antes → depois)

### ETAPA 1 — Limpeza básica
Coloca tudo num "chão comum".

| Operação | Antes | Depois |
|---|---|---|
| Maiúsculas | `Tensão Fase A` | `TENSÃO FASE A` |
| Remover acentos | `TENSÃO` | `TENSAO` |
| Colapsar espaços / `\n` | `BARRA  AT\nP` | `BARRA AT P` |
| Trocar `/ - .` por espaço | `A-B / N` | `A B N` |

```python
def etapa1(s):
    s = unicodedata_sem_acento(s.upper())
    s = re.sub(r'[\/\-\.\(\)\:]', ' ', s)
    s = re.sub(r'\s+', ' ', s).strip()
    return s
```

---

### ETAPA 2 — Expansão de abreviações
Padroniza a forma reduzida na forma plena (lista no `REGRAS.txt`).

| Antes | Depois |
|---|---|
| `DISJ` | `DISJUNTOR` |
| `SEC` / `SECC` | `SECCIONADORA` |
| `TEMP` | `TEMPERATURA` |
| `POT` / `CORR` / `TENS` | `POTENCIA` / `CORRENTE` / `TENSAO` |

> ⚠️ Só por **palavra inteira** (word-boundary), senão `SEC` estraga
> `SECCIONADORA`. Use `\b(DISJ|SEC|...)\b`.

---

### ETAPA 3 — Separar CÓDIGO do TEXTO  ⭐ (o pulo do gato)
Códigos ANSI (`50`, `51`, `87`) e números de função são **ouro** — pesam mais
que palavras. Extraia-os num campo próprio.

```
"27 TRIP SUBTENSAO"  →  codigos = {27}   texto = "TRIP SUBTENSAO"
"DISJUNTOR 52 ABRIR FECHAR" → codigos = {52}  texto = "DISJUNTOR ABRIR FECHAR"
```

- `codigos` = regex `(?<![0-9])([2-9][0-9])(?![0-9])` → 20–99.
- Esses códigos entram **direto** no score (peso alto), fora do TF-IDF.
- O `texto` restante segue para TF-IDF/embedding.

---

### ETAPA 4 — Tokenizar
Quebra em palavras de **≥ 2 caracteres**.

```
"TENSAO BARRA AT P"  →  ["TENSAO", "BARRA", "AT", "P"]
```

---

### ETAPA 5 — Remover stopwords de DOMÍNIO  ⭐
Não são as stopwords comuns (de, da) — são os **ruídos elétricos** que aparecem
em quase todo sinal e não ajudam a distinguir.

| Categoria | Descartar |
|---|---|
| Conectivos | `DO DA DE NO NA E O A` |
| Localização | `FASE AT BT MT LADO BARRA` |
| Marcadores P/A | `P A` (relé principal/alternado) |
| Vãos isolados | `06Q0 01F1` (regex `\d+[A-Z]\d*`) |
| IED/path | `RPRE_MM MEAS PONTO LOGICO` |

```
["TENSAO","BARRA","AT","P"]  →  ["TENSAO"]      ✅ sobrou só o sinal
["DISJUNTOR","ABRIR","FECHAR"] → ["DISJUNTOR","ABRIR","FECHAR"]  ✅
```

> 🎯 É **aqui** que você "filtra apenas o relevante". Quanto melhor a lista de
> stopwords de domínio, mais limpo o vetor — e mais certo o matching.

---

### ETAPA 6 — Stemming leve (opcional)
Junta variações da mesma raiz (sem biblioteca pesada).

| Antes | Depois |
|---|---|
| `CORRENTES` | `CORRENTE` |
| `DESLIGAMENTO` `DESLIGADO` | `DESLIG` |
| `ATUADO` `ATUACAO` | `ATUA` |

Regra simples: cortar sufixos `S, MENTO, ADO, ACAO, AGEM`. Cuidado para não
encurtar demais (validar contra a base).

---

### ETAPA 7 — Documento limpo
Junta os tokens que sobraram. É o **input final** do TF-IDF/embedding.

```
bruto:   "Disj. 06Q0 (AT: 52) - Abrir / Fechar"
limpo:   codigos={52}   doc="DISJUNTOR ABRIR FECHAR"
```

---

## 📊 Preparando o TF-IDF

**TF-IDF** dá peso por **raridade**: `BUCHHOLZ` (raro) vale mais que `ALARME`
(comum). Serve para **desempatar** candidatos no motor de regras.

### Passo a passo
1. **Corpus** = todas as descrições do dicionário oficial, já passadas pela
   ETAPA 1–7 (`SIGLA → doc_limpo`).
2. **Treinar o IDF** uma vez (offline):
   ```python
   from sklearn.feature_extraction.text import TfidfVectorizer
   vec = TfidfVectorizer(ngram_range=(1,2))      # uni + bigramas
   M = vec.fit_transform([doc_limpo de cada SIGLA])
   ```
3. **Salvar** `vec` e `M` (ou só o IDF em `data/idf.json` para não depender do
   sklearn em produção).
4. **Em runtime:** vetoriza a descrição do campo (`vec.transform([doc])`) e
   compara (cosseno) com as linhas de `M` → ranqueia os SIGLAs candidatos.

> 💡 Sem sklearn no destino: pré-calcule o IDF de cada termo num `idf.json` e
> faça o score à mão (`tf * idf`), somando os termos em comum. Zero dependência.

---

## 🧠 Preparando o Embedding (semântica, local e offline)

**Embedding** captura **significado**: entende que `"religamento"` ≈
`"reclose"` ≈ `"79"` mesmo sem palavra em comum. Reproduzível e offline.

### Passo a passo
1. **Modelo local pequeno** (escolha um):
   - `sentence-transformers` → `paraphrase-multilingual-MiniLM-L12-v2` (~120 MB)
   - ou `nomic-embed-text` via Ollama (offline)
2. **Pré-calcular** o vetor de cada SIGLA (descrição limpa) → `data/sigla_vec.npy`
   (matriz N×384) + `data/sigla_ids.json` (ordem dos SIGLAs).
   ```python
   from sentence_transformers import SentenceTransformer
   model = SentenceTransformer('paraphrase-multilingual-MiniLM-L12-v2')
   vecs = model.encode([doc_limpo de cada SIGLA], normalize_embeddings=True)
   np.save('data/sigla_vec.npy', vecs)
   ```
3. **Em runtime:**
   ```python
   q = model.encode([doc_do_campo], normalize_embeddings=True)[0]
   scores = sigla_vec @ q          # cosseno (vetores já normalizados)
   melhor = sigla_ids[scores.argmax()]
   conf   = int(scores.max() * 100)   # vira a confiança (teto MÉDIA)
   ```
4. **Posição na pipeline:** camada **6.5** — entre o fuzzy e a IA generativa.
   Pega sinônimos que o fuzzo não pega, **sem chutar vão**, e é determinístico.

> 🔌 Tudo cabe no zip: o modelo (~120 MB) + os `.npy/.json` pré-calculados →
> **funciona sem internet**.

---

## 🧪 Como validar a normalização

1. Pegue 30 descrições reais (FredW) com a SIGLA correta conhecida.
2. Rode ETAPA 1–7 e confira o `doc_limpo` — sobrou só o que importa?
3. Meça o **top-1** e **top-3** do TF-IDF e do embedding contra o gabarito.
4. Ajuste as **stopwords de domínio** (ETAPA 5) — é o que mais move o resultado.
5. Repita. Meta: top-3 ≥ 90% (o usuário escolhe entre 3 na tela de revisão).

---

## ✅ Resumo (cola rápida)

| Etapa | O que faz | Impacto |
|---|---|---|
| 1 Limpeza | caixa, acento, espaço | base |
| 2 Abreviação | Disj→Disjuntor | alinha campo×base |
| 3 **Código** | separa 50/51/87 | **alto** (peso direto) |
| 4 Tokenizar | quebra em palavras | base |
| 5 **Stopwords domínio** | tira FASE/AT/vão/IED | **alto** (limpa ruído) |
| 6 Stemming | junta raízes | médio |
| 7 Doc limpo | input do vetor | — |
| TF-IDF | peso por raridade | desempate |
| Embedding | significado/sinônimo | semântica offline |

> **Prioridade de esforço:** ETAPA 3 (códigos) e ETAPA 5 (stopwords de domínio)
> são as que mais melhoram o resultado. Comece por elas.
