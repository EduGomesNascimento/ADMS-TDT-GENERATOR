# Gerador Automático de TDTs · ADMS

Aplicação para montar dinamicamente os sinais de um equipamento elétrico e
exportar uma **TDT em Excel no formato exato aceito pelo ADMS** — funcionando
como um "clonador inteligente" das TDTs reais da base.

A lógica **não é genérica**: a estrutura (colunas, cabeçalhos, estilos,
enumerações e o significado de cada sinal) foi **aprendida diretamente das TDTs
reais e do dicionário oficial de sinais** fornecidos.

---

## Como funciona (o que foi descoberto nos arquivos reais)

**Estrutura da TDT.** Cada arquivo `.xlsx` tem até 3 sheets:
- `DMSMatchingTemplateInfo` — tabela fixa de enumerações/validações do ADMS
  (copiada *verbatim* em toda exportação).
- `DNP3_DiscreteSignals` — sinais digitais.
- `DNP3_AnalogSignals` — sinais analógicos.

Cada sheet de sinais tem **4 linhas de cabeçalho** (seção / tabela interna /
código do campo / rótulo legível) e os dados a partir da 5ª linha.

**Nomenclatura dos sinais:** `{ALIAS}_{MÓDULO}_{ID_EQUIPAMENTO}_{SUFIXO}`
— ex.: `FWB_AL13_52-13_IA`. O **sufixo** (após o último device token) é a chave
do tipo de sinal, mapeada no dicionário `RGE ADMS_Lista Sinais Padrão` para
descrição, tipo, medição, unidade e fases.

**Campos dependentes da identidade** (detectados automaticamente comparando o
mesmo sufixo entre 71 dispositivos reais — variam em 71/71): `Signal Name`,
`Signal Custom ID`, `Remote Point Name`, `Remote Point Custom ID`,
`Input Coordinates`, `Device Mapping`, `Remote Unit`. **Todo o resto é fixo por
sufixo** e copiado exatamente do template (formatação, tipos e valores).

A substituição é feita por **tokenização genérica**: ao construir o catálogo,
toda célula de texto que contém o prefixo/alias/módulo/device vira um placeholder
(`<<PREFIX>>`, `<<ALIAS>>`, …); na exportação os placeholders são substituídos
pelos dados do novo equipamento. Assim os campos variáveis são descobertos sem
hardcode.

**Fidelidade comprovada:** o round-trip que reconstrói a TDT FWB original tem
**zero diferenças de valor real** (apenas os campos de identidade, que são
re-sequenciados de propósito).

---

## Arquitetura

```
adms-tdt-generator/
├─ backend/                 # Python · FastAPI + openpyxl
│  ├─ build_catalog.py      # camada de APRENDIZAGEM (gera data/catalog.json)
│  ├─ tdt_engine.py         # motor de exportação (clona o template, substitui)
│  ├─ app.py                # API REST (/api/catalog, /preview, /export)
│  └─ data/
│     ├─ catalog.json       # base de conhecimento gerada
│     └─ reference_template.xlsx  # TDT real usada como template de estilos
└─ frontend/                # React + Vite + TypeScript + Tailwind
   └─ src/                  # wizard de 4 passos (UI/UX)
```

- **Backend em Python** porque a exportação **idêntica** ao original exige
  `openpyxl` operando sobre um workbook real (preserva estilos/validações).
- **Frontend React/TS/Tailwind** com wizard interativo e prévia em tempo real.

---

## Como rodar

Pré-requisitos: **Python 3.11+** e **Node 18+**.

### 1. Backend
```powershell
cd adms-tdt-generator\backend
pip install -r requirements.txt
python build_catalog.py            # gera/atualiza o catálogo (1x ou ao trocar TDTs)
uvicorn app:app --host 127.0.0.1 --port 8077
```

### 2. Frontend
```powershell
cd adms-tdt-generator\frontend
npm install
npm run dev                        # http://localhost:5180
```

O Vite faz proxy de `/api` para o backend em `:8077`.

Ou use o atalho na raiz do projeto: **`start.ps1`** (sobe os dois).

---

## Fluxo de uso (UI)

1. **Equipamento** — escolha o tipo (Alimentador, Transformador, …).
2. **Identificação** — preencha alias, módulo e identificador. Os sinais só são
   liberados após o preenchimento. Defina também o Custom ID inicial e os
   endereços DNP3 base (auto-sequenciados; desmarque para manter os do template).
3. **Sinais** — marque os sinais digitais/analógicos desejados (com busca e
   "selecionar todos"). Prévia do nome em tempo real. Sinais com **comando**
   recebem o selo **CMD**: ao marcá-los, abre-se a configuração de *output*
   (formato + Output Coordinates).
4. **Revisão** — confira a tabela final (incluindo Comando/Output) e **exporte a
   TDT (.xlsx)**.

### Protocolo (DNP3 / IEC104)

Na **tela inicial escolhe-se o protocolo** e todo o fluxo segue por ele:
- **DNP3** — assistente por equipamento (5 tipos) + importação de lista.
- **IEC 60870-5-104** — importação de lista (gera as sheets `IEC104_*`).

Cada protocolo tem seu template (`data/reference_template[_iec104].xlsx`) e seu
índice (`data/sigla_index[_iec104].json`, gerado por `build_index.py [iec104]`).
Os rótulos de coluna são quase idênticos entre protocolos; o motor lê os rótulos
do próprio template (tolera "Input Coordinate" singular do IEC104).

### Importar lista de pontos (entrada padrão)

Botão **"Importar lista de pontos"** na tela inicial. Aceita o `.xlsx` padrão com
abas **Discreto** (`SIGLA SINAL`, `NOME`, `INDEX DNP3 - Entrada`,
`INDEX DNP3 - Comando`, `AOR`) e **Analógicos** (`SIGLA`, `NOME`, `Escala`,
`INDEX DNP3`, `AOR`). Cada SIGLA é casada com um **índice global da base ADMS**
(`data/sigla_index.json`, 2994 siglas digitais + 447 analógicas) para herdar os
campos fixos; o `NOME` define o Signal Name, e os índices DNP3, Escala e AOR vêm
da lista (a lista é autoritativa nas coordenadas). Mostra resumo de casados/não
casados antes de exportar.

> `data/sigla_index.json` é gerado por `build_index.py` (1 varredura da base) e
> **deve ser distribuído junto** — não dá para recriar sem a base de 98MB.

### Sinais de comando (output)

Detectados automaticamente nas TDTs reais (Direction=ReadWrite/Write, Control
Codes, Output Data Type). Ao selecionar um sinal de comando você escolhe o
**formato** e as **Output Coordinates** (auto-sequenciadas a partir da *Coord.
base de comandos*, faixa própria como no índice FredW, ou digitadas manualmente):

| Formato | Control Codes | Coords |
|---|---|---|
| Manter do template | original do sinal | conforme template |
| Trip / Close (par) | `TripPulseOn;ClosePulseOn` | 2 (`c;c`) |
| Close / Close (par) | `ClosePulseOn;ClosePulseOn` | 2 (`c;c`) |
| Pulso único — Close | `ClosePulseOn` | 1 (`c`) |
| Pulso único — Trip | `TripPulseOn` | 1 (`c`) |
| Latch On / Off | `LatchOn;LatchOff` | 2 (`c;c`) |

---

## Estender para novos tipos de equipamento

Edite `SOURCE_DEVICES` em `backend/build_catalog.py` adicionando o equipamento
de referência (`alias`, `module`, `device`, arquivo-fonte e sheets) e rode
`python build_catalog.py`. O novo tipo aparece automaticamente na UI.

> **Tipos hoje (5):** Alimentador (39 dig + 13 anl), Transformador Auxiliar/TSA
> (43 dig), **Transformador principal TR** (consolidado corpo+AT+BT, 149 dig +
> 20 anl, `build_tr.py`), **Banco de Capacitores** (31 dig + 7 anl,
> `build_extra.py`) e **Linha de Transmissão LT** (consolidada: disjuntor +
> seccionadoras + grupos de proteção A/P, 158 dig + 17 anl, `build_lt.py`;
> usuário informa alias + nome da linha).
>
> **Extração da base completa:** `build_tr.py` (TR consolidado) e `build_extra.py`
> (tipos de módulo simples: BC, IB…) varrem `Export_base_Full…xlsx` e anexam ao
> catálogo. `build_extra.py` usa tokenização **por token inteiro** (evita
> super-substituição de módulos curtos como `IB`).
>
> **Barra** (`build_barra.py`): módulos `BP*` (proteção diferencial 87B), 13
> sinais digitais cujas SIGLAs vêm da lista padrão e os templates do índice da
> base. Estrutura `ALIAS_BUS_BUS_SIGLA`.
>
> **Pendentes:** IB (precisa de instância-fonte melhor); sinais
> "digital-analógico" (sheet `DNP3_DiscreteAnalog`, ex.: TAP).
>
> **Modos de identificação (UI):** `standard` (alias+módulo+device),
> `transformer` (alias+N → TR{N}), `line` (alias+nome da linha), `bus`
> (alias+nome da barra → ALIAS_BUS_BUS). Definidos por `paramKind`/`consolidated`.
>
> **Sinais "digital-analógico"** (ex.: TAP): o dicionário os classifica como
> classe própria; quando a base fornecer uma sheet dedicada, basta mapeá-la em
> `build_catalog.py` — a tokenização e o motor já suportam.
