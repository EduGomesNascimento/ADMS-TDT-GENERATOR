# Como usar — Fluxo completo TDT ADMS

Duas ferramentas que trabalham juntas, do ponto de campo até a TDT pronta para o ADMS:

```
┌─────────────────────────┐     ┌──────────────────────────┐     ┌──────────┐
│ 1. Gerador de Lista     │ →   │ 2. Gerador de TDTs ADMS  │ →   │  ADMS    │
│    Resumida (desktop)   │     │    (web app)             │     │ (import) │
│  planilha bruta → lista │     │  lista/assistente → TDT  │     │          │
└─────────────────────────┘     └──────────────────────────┘     └──────────┘
```

- **Pasta:** `gerador-lista-resumida/` (ferramenta 1) e `adms-tdt-generator/` (ferramenta 2).

---

## Ferramenta 1 — Gerador de Lista Resumida

Transforma a planilha **bruta de pontos** (várias abas por módulo, com a coluna
`UTILIZADO?`) em uma **lista resumida padronizada** (abas `Analogicos`,
`Discreto`, `DiscretoAnalogico`, `Erros`) — exatamente o formato que a ferramenta 2 consome.

### Rodar
```bat
cd gerador-lista-resumida
pip install -r requirements.txt
run.bat          (ou: python app\main.py)
```

### Passo a passo
1. **Arquivo de input** — a planilha bruta de pontos (a que tem `UTILIZADO?`, `TIPO`, `SIGLA SINAL`, `NOME`, `INDEX DNP3`…).
2. **Base ADMS** — `Pontos Padrao ADMS_v1.xlsx` (dicionário de SIGLAs válidas).
3. **Pasta de output**.
4. **PROCESSAR**.
5. Abra o resultado `lista_resumida_*.xlsx`:
   - **Analogicos / Discreto / DiscretoAnalogico** → pontos válidos (vão pra TDT).
   - **Erros** → o que ficou de fora e o porquê. Corrija no input e reprocesse até zerar.

> Só entram pontos com **`UTILIZADO? = SIM`**. O **TIPO** define o destino:
> `A`=Analógico, `D`=Discreto, `C`=Comando (vira a coluna *INDEX DNP3 - Comando*
> do par discreto), **`A/D`=Digital-Analógico** (ex.: **TAP** do comutador → aba
> `DiscretoAnalogico`).

---

## Ferramenta 2 — Gerador de TDTs ADMS

Gera a **TDT `.xlsx` idêntica ao formato do ADMS**. Dois caminhos.

### Rodar (como app)
```bat
cd adms-tdt-generator
ABRIR.BAT
```
Abre em janela de app (Chrome/Edge). Requer **Python 3.11+** e **Node 18+**.

### Caminho A — Importar a lista de pontos (recomendado)
1. Na tela inicial, escolha o **protocolo** (DNP3 ou IEC 60870-5-104).
2. Clique **Importar lista de pontos** e selecione o `lista_resumida_*.xlsx` da ferramenta 1.
3. Confira o resumo (casados / não-casados na base / **duplicados** / erros de origem).
4. **Baixar relatório** (opcional) — `.xlsx` com Resumo, **Revisão da Lista** (todos os pontos com status) e Problemas.
5. **Exportar TDT (.xlsx)** → importe no ADMS.

### Caminho B — Assistente por equipamento (DNP3)
1. Protocolo **DNP3** → escolha o tipo: Alimentador, TSA, Transformador (TR), Banco de Capacitores, Linha (LT) ou Barra.
2. Preencha a identificação (alias da subestação + módulo/nº do trafo/nome da linha/nome da barra).
3. Marque os sinais (digitais, analógicos, comandos e **digital-analógico TAP** quando houver). Defina os endereços DNP3 base.
4. Revise e **Exporte a TDT**.

---

## O que é preservado / substituído

- **Preservado** (idêntico à base ADMS): todos os campos fixos por SIGLA — tipo de
  sinal, message mapping, control codes, medição, unidade, etc., além da estrutura
  e formatação do Excel.
- **Substituído** por equipamento: Signal Name, Device Mapping, Remote Unit
  (`UTR_{alias}_1`), Custom IDs (re-sequenciados), AOR (`{alias} Trans/Distr`) e as
  coordenadas DNP3 (da lista ou auto-sequenciadas). O **Signal Custom ID** sai
  vazio (o ADMS gera).

---

## Dicas / solução de problemas

- **Sinal no Erros "Sigla X não encontrada na base ADMS"**: a SIGLA não está em
  `Pontos Padrao ADMS_v1.xlsx`. Adicione-a na aba correspondente (col. A = SIGLA,
  col. B = descrição) e reprocesse. *(Foi assim que o TAP passou a ser aceito.)*
- **TAP / A/D**: agora suportado de ponta a ponta — o resumidor o coloca na aba
  `DiscretoAnalogico` e o gerador escreve na sheet `DNP3_DiscreteAnalog`.
- **Nomes duplicados** na lista: o gerador avisa (gerariam linhas repetidas no ADMS).
- **Arquivos `data/*.json` e `data/*.xlsx`** do `adms-tdt-generator` devem ir junto
  ao distribuir — são o catálogo/índices/templates e não são recriáveis sem a base completa.
