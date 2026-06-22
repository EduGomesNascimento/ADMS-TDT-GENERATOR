# Plano — Motor de Regras, Leitura sem IA e Empacotamento Único

> Documento de planejamento. Consolida a pipeline (fluxograma do colega), as
> decisões de normalização, a estratégia TF-IDF/embedding e como tornar a
> ferramenta **offline e transportável num arquivo único**.

---

## 1. Objetivo

Ler **listas de pontos NÃO padronizadas** e gerar a TDT do ADMS:

1. **Sem IA por padrão** — um **motor de regras** determinístico resolve o máximo.
2. **IA só como scan final OPCIONAL** — roda **localmente** (offline), apenas no
   que as regras não acharam.
3. **Tudo no zip** — IA + plataforma juntas; funciona sem internet.
4. **Revisão humana fácil** — incertos destacados, checkbox para o usuário
   confirmar, e cor na linha da TDT para o que ficou duvidoso.

---

## 2. Pipeline (baseada no fluxograma)

```
1.0 INPUT (lista homogênea OU não-homogênea)
        │
   [CLASSIFICADOR]  homogênea? ──Sim──► caminho SEM IA direto (já casa por SIGLA)
        │ Não
   2.0 PRÉ-PROCESSAMENTO  (limpa abas-lixo: Capa, Slot, Índice, Lógicos)
        │
   2.X ANALISADOR DE COLUNAS  (detecta Módulo / Tipo / Descrição / Index / UTR-ID)
        │
   [TOKENIZAR]  (descrição → tokens normalizados; extrai códigos ANSI)
        │
   [PONTUAÇÃO]  (cada SIGLA candidata recebe um score por sobreposição+código)
        │
   [MOTOR DE REGRAS]  (camadas 1–6 do REGRAS.txt; aplica pesos)
        │
   [NORMALIZAÇÃO DOS NOMES]  (ALIAS_MOD_DEV_SIGLA; corrige variações)
        │
   [INSERÇÃO]  ──► TDT (linhas incertas destacadas em amarelo)
        │
   [AGENTE IA — REVISÃO (LOCAL, OPCIONAL)]  só os SEM/MÉDIA/BAIXA
        │
   [REVISÃO HUMANA]  (tabela lado a lado + checkbox) ──► TDT final + relatório
```

**Estado atual:** caminhos 2.X→Inserção já implementados (`ai_mapper.py` +
`tdt_engine.py`). Falta: classificador automático homogêneo/não, IA de revisão
local embarcada, e o aprendizado (gravar correções para reusar).

---

## 3. Normalização das colunas e dados relevantes (o ponto-chave)

As listas variam muito (FredW = Módulo/Tipo/Descrição; URA = UTR-ID; COS/GPR =
Slot/CTR). A normalização precisa **descobrir as colunas relevantes** sem
depender de uma posição fixa.

### 3.1 Detecção de cabeçalho (multi-estratégia)
1. **Por palavra-chave** — procura nas primeiras ~8 linhas um cabeçalho que
   contenha `MODULO` + `TIPO` + `DESCRI` (regex tolerante a acento e quebra de
   linha `\n`). → formato estruturado.
2. **Por padrão de UTR-ID** — acha a coluna com mais valores no padrão
   `ALIAS_TOKENS` (ex.: GPR21_43RC). → formato UTR-ID.
3. **Por densidade** — se nenhuma bater, escolhe a coluna com mais texto
   descritivo (fallback).

### 3.2 Colunas relevantes (mínimo necessário)
| Papel | Sinônimos aceitos | Uso |
|---|---|---|
| **Módulo** | Módulo, Modulo, Bay, Vão, Equipamento | nome do sinal |
| **Tipo** | Tipo, Tipo Ponto, A/D/C | roteia analog/discreto/comando |
| **Descrição** | Descrição, Descrição do ponto, Projeto-Descrição | **chave do matching** |
| **Index/DNP3** | Index, DNP3, Endereço, Coordinate | coordenada de entrada/saída |
| **UTR-ID** (opc.) | Identificação, Ponto na UTR | token match direto |

### 3.3 Limpeza de dados (antes de tokenizar)
- Descartar abas: `Capa, CAPA, Índice, Index_DNP3, Lógicos, SNMP, Slot*, Saúde,
  Calculados, A8000` (regex `_SKIP_SHEETS`).
- Remover linhas totalmente vazias e linhas de subtítulo (sem Tipo/Descrição).
- Aplicar `_norm` (maiúsculas, sem acento, abreviações expandidas).

---

## 4. Filtrar o relevante para TF-IDF + Embedding

A descrição tem **ruído** (bay, fase, "P/A", relé) que atrapalha. Estratégia:

### 4.1 O que MANTER (sinal) vs DESCARTAR (ruído)
- **Manter:** termos de função (TENSAO, CORRENTE, SUBTENSAO, RELIGAMENTO,
  BLOQUEIO…), códigos ANSI (50, 51, 87…), estado (ABERTO/FECHADO, ATUADO).
- **Descartar (stopwords de domínio):** DO, DA, DE, FASE, AT, BT, MT, P, A,
  números de vão isolados (06Q0), siglas de IED, "RPRE_MM", paths com "/".

### 4.2 TF-IDF (peso por raridade)
- Corpus = todas as descrições do dicionário oficial (SIGLA→descrição).
- Termos raros (ex.: "BUCHHOLZ", "DIFERENCIAL") pesam mais que comuns
  ("ALARME", "FASE"). Isso desempata candidatos no score do motor de regras.
- Implementação leve: `sklearn.feature_extraction.text.TfidfVectorizer`
  (offline) OU um IDF pré-computado salvo em `data/idf.json` (sem dependência).

### 4.3 Embedding (semântica, opcional, local)
- Modelo **pequeno e offline**: `sentence-transformers` multilíngue mini
  (ex.: `paraphrase-multilingual-MiniLM-L12-v2`, ~120 MB) OU embeddings do
  Ollama (`nomic-embed-text`).
- Pré-calcular o vetor de cada SIGLA (descrição) → `data/sigla_vectors.npy`.
- Em runtime: vetoriza a descrição do campo, busca o SIGLA mais próximo
  (cosseno). É **determinístico** (mesmo input → mesmo output) e **offline**.
- Entra como **camada 6.5** (entre fuzzy e IA): cobre sinônimos que o fuzzo
  não pega, sem chutar vão.

> Ordem recomendada: regras exatas → TF-IDF (desempate) → embedding local
> (semântica) → IA generativa local (só o resto, se o usuário quiser).

---

## 5. IA de revisão LOCAL (offline)

Sem internet, sem chave. Opções (da mais leve à mais pesada):
1. **Embedding local** (item 4.3) — não é "IA generativa", mas resolve a maior
   parte da semântica de forma reproduzível. **Recomendado como padrão.**
2. **Ollama embarcado** — `qwen2.5:3b`/`phi4-mini` rodando local. Já integrado
   (provider `ollama`). Lento em CPU, mas 100% offline.
3. **llama.cpp + modelo GGUF** dentro do zip — um binário + um `.gguf` (~2 GB),
   chamado via subprocess. Mais transportável que o Ollama (não precisa instalar).

A IA generativa fica **só no scan final**, sobre os SEM/BAIXA, e nunca marca ALTA.

---

## 6. Empacotamento em ARQUIVO ÚNICO (transportável)

Meta: um zip/exe que roda em qualquer PC, **sem instalar Python/Node/internet**.

### 6.1 Backend → executável único
- **PyInstaller** `--onefile` empacota Python + FastAPI + openpyxl + os dados
  (`catalog.json`, `sigla_index.json`, templates) num `tdt_backend.exe`.
- `--add-data` inclui a pasta `data/`. O `.exe` sobe o servidor em localhost.

### 6.2 Frontend → estático embutido
- `npm run build` gera `dist/` (HTML/JS estáticos). O FastAPI serve o `dist/`
  como arquivos estáticos (StaticFiles) → não precisa de Node no destino.

### 6.3 IA local embutida
- Embedding: incluir `sigla_vectors.npy` + o modelo MiniLM (ou só o IDF).
- (Opcional) `llama.cpp` + `.gguf` na pasta `ai/` do pacote.

### 6.4 Resultado
```
TDT-ADMS.exe   (backend + frontend + dados + embeddings)   <- clique e usa
  └─ (opcional) ai/modelo.gguf  para a IA generativa offline
```
Tamanho estimado: ~80 MB (sem IA generativa) / ~2 GB (com modelo GGUF).

---

## 7. Revisão humana — destaque + checkbox

- **TDT:** linhas incertas (não-ALTA) já saem **amarelas** (FFF2CC) — só a cor
  muda, formatação preservada. (implementado)
- **Tela de revisão:** SIGLA pré-marcada, editável com autocomplete, filtro "só
  pendentes". (implementado)
- **A FAZER — checkbox de confirmação:** uma coluna ☐ por linha incerta; ao
  marcar, vira ALTA (confirmado pelo humano) e sai do amarelo. Botão "confirmar
  todos os marcados". Gera relatório do que foi confirmado/alterado.
- **A FAZER — aprendizado:** salvar `descrição→SIGLA` confirmada em
  `data/aprendizado.json` e usá-la como camada 0 (a mais forte) nas próximas
  listas. Assim a ferramenta "aprende" o padrão da concessionária.

---

## 8. Roadmap (ordem sugerida)

1. [feito] Motor de regras (camadas 1–6) + destaque amarelo + tela de revisão.
2. Classificador automático homogênea/não-homogênea (já temos heurística).
3. TF-IDF offline (`data/idf.json`) para desempate de candidatos.
4. Embedding local (MiniLM) como camada semântica reproduzível.
5. Checkbox de confirmação + relatório + aprendizado (`aprendizado.json`).
6. Empacotamento PyInstaller `--onefile` + frontend estático no FastAPI.
7. (Opcional) IA generativa local via llama.cpp embarcado.
