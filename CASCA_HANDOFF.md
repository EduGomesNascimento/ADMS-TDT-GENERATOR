# SE CASCA — UTR_CAS_3 (DNP3) · handoff

Tudo que a próxima pessoa (ou o próximo Claude) precisa para continuar.
Última geração: `python backend/make_casca.py` seguido de `python backend/check_casca.py`.

---

## 1. O que é o trabalho

A CASCA opera hoje em **IEC 101/104**. Ela vai ser modernizada para **DNP3**.
Foi criada no ADMS uma cópia gráfica da subestação — o container **`Casca_Obra`** —
e dentro dela precisa nascer uma **UTR nova, a `UTR_CAS_3`**, com os sinais da
lista de pontos nova. As UTRs atuais (`UTR_CAS_1` IEC104 e os religadores `RAD_*`)
são só referência: **não** interagem com a nova.

### Regra que rege todo o mapeamento

> "seguir com as informações do unifilar, só usando os sinais da lista mesmo,
> mas o que está no campo real é o que está no ADMS/unifilar"

Traduzindo:

| Assunto | Quem manda |
|---|---|
| Quais sinais existem, nome, tipo, índice DNP3, comando, escala, níveis lógicos | **Lista de pontos** |
| Quais dispositivos existem (disjuntor, seccionadora, relé, TC, TP, trafo, barra) | **Unifilar / modelo do ADMS** |

Por isso a coluna **Device Mapping** aponta para o dispositivo do unifilar, mesmo
quando a lista usa outra numeração.

---

## 2. Arquivos

### Entradas (não são gerados — vêm do time)

| Arquivo | O que é |
|---|---|
| `RGE ADMS_Lista Pontos Casca.xlsx` | Lista de pontos nova. Fonte dos SINAIS. |
| `PT-MOD-SE-CASCA.xml` | Export do changeset do `Casca_Obra`. Fonte dos DISPOSITIVOS. |
| `CASCA.xlsx` | TDT atual da CASCA (IEC101/104/DNP3). De onde se aprende a convenção de Device Mapping. |
| `TDT_LVA_AL24.xlsx` | Esqueleto: TDT válida com a aba `DNP3_RTUs`, usada como molde. |
| `erros.csv` | Retorno do validador do ADMS. Usado para conferir o que ainda não mapeia. |
| `unifilar` (foto) | Diagrama da CASCA atual, com os números de equipamento. |

### Saídas (geradas)

| Arquivo | O que é |
|---|---|
| `TDT_CASCA_UTR_CAS_3.xlsx` | **A TDT.** 1282 sinais. É o que se importa no ADMS. |
| `RGE ADMS_Lista Pontos Casca_CORRIGIDA.xlsx` | A lista com a coluna INDEX DNP3 arrumada. Referência para parametrizar a UTR ELIPSE. |
| `CASCA_RELATORIO.xlsx` | 16 abas explicando tudo. Começa pelo `0-LEIA-ME`. |

### Código

| Arquivo | Papel |
|---|---|
| `backend/make_casca.py` | Gera os 3 arquivos de saída. É o programa principal. |
| `backend/casca_devmap.py` | Decide o Device Mapping de cada sinal. **A tabela `MODULO_EQUIV` fica aqui.** |
| `backend/check_casca.py` | Confere lista × TDT de forma independente. Rodar sempre depois de gerar. |
| `backend/excel_native.py` | Regrava o .xlsx pelo Excel (COM). **Obrigatório** — o ADMS rejeita OOXML do openpyxl. |
| `backend/data/sigla_index.json` | Linha-molde de cada sigla, extraída da base completa. |

---

## 3. Decisões que já foram tomadas (e por quê)

### 3.1 Todo ponto da lista recebe coordenada

A lista tem 2535 linhas de sinal: 1453 com `Utilizado? = SIM`, 1082 com `NÃO` e
algumas reservas sem sigla. **Todas** recebem índice DNP3.

As `NÃO` **não entram na TDT** (não são sinais ativos), mas **reservam o endereço** —
assim, quando forem ativadas, a numeração não muda e nada precisa ser renumerado.

### 3.2 Re-sequenciamento contínuo das coordenadas

A lista original tinha **704 ocorrências de coordenada repetida** e **1145 pontos
com `#REF!`** (fórmula quebrada). O DNP3 exige índice único por grupo, então o
ADMS rejeitava com *"does not have unique input coordinates"*.

Solução: redistribuir em ordem, sem buraco e sem repetição, preservando o arranjo
original (ordena pelo índice antigo). Onde havia `490, 490, 491` virou `490, 491, 492`.
Cada módulo continua de onde o anterior parou. Os `#REF!` entram no fim da fila do
seu grupo. O de-para completo está na aba `2`.

**Grupos de índice (espaços independentes no DNP3):**

- `D` — binary input · 0..1904
- `A` + `A/D` — analog input · 1..361 — **atenção:** o TAP (DiscreteAnalog) lê um
  ponto analógico, então divide a numeração com os analógicos comuns. Isso causava
  colisão silenciosa (TAP do TR6 = 70 = IA do TR7AT).
- `C` — binary output (comandos) · 1..300, no formato `n;n`

### 3.3 Nome duplicado é renomeado, não descartado

13 sinais tinham NOME repetido (o ADMS exige nome único). O 2º recebe sufixo no
device — `CAS_TR7_TR7_TOD` → `CAS_TR7_TR7-2_TOD` — e **entra na TDT**. O Device
Mapping continua apontando para o device real. Aba `6`.

O nome limpo fica com a linha que tinha índice válido na origem (ex.: `CAS_AL18_AL18_IA`
existe em `BC 1` com `#REF!` e em `TRANSFERENCIA 24-01` com índice — o limpo ficou
com a TRANSFERENCIA).

### 3.4 Comandos

176 linhas `C` na lista. Cada uma precisa achar o sinal `D` que vai carregá-la:

1. **NOME idêntico** — caso normal.
2. **Mesmo módulo + mesma sigla** — a lista às vezes põe o comando no disjuntor
   (`CAS_LT1_52-1_25IE`) e o status no módulo (`CAS_LT1_LT1_25IE`). 3 casos.
3. **Comando puro** (sem status nenhum): vira sinal discreto próprio com
   **Input de preenchimento 9599+**, conforme convenção do usuário
   ("se for impossível achar index, bote números beeem altos"). 5 casos:
   `CDC` do TR6/TR7 (comutador) e `86RM` de BP69/BP113.8/BP213.8 (rearme 86).

Sinal `D` sem comando tem os campos de comando **limpos** e `Direction = Read` —
senão o validador exige `Output Coordinates`. Aba `12`.

### 3.5 Device Mapping — segue o unifilar

O `Casca_Obra` é um clone da CASCA **atual**, e a lista renumerou tudo:

| | Unifilar / ADMS | Lista nova |
|---|---|---|
| Barras AT | P1 e P2 — 138 kV | BP69 — 69 kV |
| Barras BT | P3 e T1 — 23 kV | BP1 e BP2 — 13,8 kV |
| Linhas | LT KVM, LT PRI, LT SCO | LT1, LT2, LT3 |
| Trafos | TR1 15/20/25 MVA · TR2 10/12,5 MVA | TR6, TR7 |
| Alimentadores | AL12 AL13 AL14 AL15 AL21 | + AL24 AL25 AL26, BC1, BC2, IB20, 24-1, 24-2 |
| Serviço auxiliar | TSA-3 | TSA1, TSA2 |

Até os vãos com o mesmo nome mudaram de número: o AL12 é `52-12 / 29-48 / 29-50 / 29-52`
na lista e `52-02 / 29-06 / 29-08 / 29-10` no unifilar.

**Onde o ADMS busca:** cada elemento do XML tem a propriedade
`1224979098644840199` = *"ID de Mapeamento SCADA"*. A coluna Device Mapping da TDT
precisa conter **exatamente** esse texto. Se não bater:
`Could not find any device that corresponds to Device Mapping: ...`

**Ordem de resolução** (`casca_devmap.resolver`):

1. Texto idêntico a um ID do modelo.
2. **Equivalência de vão** — `MODULO_EQUIV`, aba `16`.
3. Mesmo vão, equipamento renumerado (`52-12` → `52-2`).
4. Relé específico inexistente → relé genérico `_PROT` do vão.
5. Dispositivo inexistente → o equivalente do vão. Ex.: o TR1 do unifilar entra
   por seccionadora `89-12` e **não tem disjuntor AT**, então os sinais do "52-4"
   da lista vão para a seccionadora. Aba `14`.
6. Vão que não existe no unifilar → nome canônico + aba `13`.

**Resultado: 858 de 1282 sinais mapeiam** (331 com dispositivo rebaixado).

**Os 424 restantes** estão em vãos que não existem no unifilar — não é erro de TDT,
é o diagrama que precisa crescer: `AL18` (BC 1 + transf. 24-1), `AL24`, `AL25`,
`AL26`, `TRF29` (24-2), `AL28` (BC 2), `IB20` (interbarras BT), `BP213.8`, `TSA2`.
A aba `13` lista dispositivo por dispositivo, com tipo e quantos sinais dependem.

Assim que o dispositivo for criado no `Casca_Obra` com esse ID, o sinal casa
sozinho — **nada precisa ser refeito na TDT**.

### 3.6 Equivalência de vãos — onde mexer

Está em **um lugar só**: `MODULO_EQUIV`, em `backend/casca_devmap.py`.
Cada entrada é `(destino, confiança, evidência)`.

| Lista | Unifilar | Confiança | Evidência |
|---|---|---|---|
| LT1 | LTSCO | **ALTA** | SECG `29-1` na lista = `29-01` do LT SCO |
| LT2 | LTPRI | **ALTA** | SECG `29-3` = `29-03` do LT PRI |
| LT3 | LTMRU | **ALTA** | SECG `29-5` = `29-05` do LT KVM |
| TR6 / TR6AT / TR6BT | TR1 / TR1AT / TR1BT | MÉDIA | aba "TR 1" da lista |
| TR7 / TR7AT / TR7BT | TR2 / TR2AT / TR2BT | MÉDIA | aba "TR 2" da lista |
| BP69 | B138 | MÉDIA | barra de alta |
| BP113.8 | BP23 | MÉDIA | barra de baixa 1 |
| TSA1, TSA | TSA3 | BAIXA | único TSA do unifilar |

**Se alguma estiver errada, corrija a tabela e rode de novo.** Nada mais precisa mudar.

Chave `REAPROVEITAR_DISPOSITIVO_ANTIGO` no mesmo arquivo: `False` faz tudo voltar
ao nome canônico (nenhum reaproveitamento).

### 3.7 Outros parâmetros fixos

```
RU        = UTR_CAS_3          AOR = CAS Trans
CONTAINER = Casca_Obra         FABRICANTE = ELIPSE
```

### 3.8 A lista corrigida abre sem pedir reparo

A lista original tem **vínculo externo** (`xl/externalLinks/externalLink1.xml`).
O openpyxl reescrevia essa parte sem o cache de valores e o Excel abria pedindo
reparo. Solução: congelar toda fórmula no valor calculado, remover os vínculos
externos e os nomes definidos que apontam pra fora, e regravar em formato nativo
via COM.

Se o arquivo estiver **aberto no Excel** na hora de gerar, a saída vira
`..._CORRIGIDA_NOVA.xlsx` — feche o Excel e rode de novo para gravar com o nome final.

---

## 4. Erros da planilha de origem (aba 15)

23 achados de severidade **ALTA**, detectados automaticamente pelo cabeçalho de
cada aba (`MÓDULO / DJ / SECC / SECF / SECT`). Os mais graves:

- **`TR 1` e `TR 2` declaram os mesmos equipamentos** (DJ AT `52-4`, seccionadoras
  `89-20/22/24/28`). Dois transformadores não dividem disjuntor de alta — a aba
  `TR 2` foi copiada da `TR 1` sem trocar os números. Foi isso que gerou os nomes
  duplicados.
- **`BC 1` e `BC 2` têm a célula MÓDULO escrita `AL`** (e não `BC`), então o NOME
  sai `CAS_AL18_...` e `CAS_AL28_...`. O `AL18` colide com a `TRANSFERENCIA 24-01`,
  que é o alimentador 18 de verdade.
- **`INTERBARRAS BT`** gera `CAS_IB20_...` num cabeçalho que diz ALIMENTADOR.
- Seccionadoras repetidas entre vãos diferentes: `AL 24`×`AL 15`, `TRANSF 24-2`×`AL 13`,
  `AL 25`×`TRANSFERENCIA 24-01`, `BC 2`×`AL 21`.

As abas `(FUTURO)` são reconhecidas como o mesmo vão e saem como **INFO**, não como erro.

---

## 5. Armadilhas já pagas (não repita)

- **Regravar pelo Excel é obrigatório.** `excel_native.resave_native` (pywin32
  `DispatchEx`, `FileFormat=51`). O .xlsx do openpyxl é rejeitado com
  *"Invalid TDI file format"*.
- **Importar `DNP3_RTUs` reprocessa a UTR inteira** e arrasta todo sinal filho
  para o changeset. Aqui é desejado (UTR nova). Para atualizar só alguns vãos de
  uma UTR existente, remova essa aba (ver `backend/slim_tdt.py`).
- **Ponto de posição** (Multi Coord / DoubleBit) precisa de **duas** coordenadas
  distintas (`n;n+1`). Comando usa a **mesma** repetida (`n;n`).
- **`RelayTrip` não pode apontar para dispositivo físico** (BREAKER etc.) — o
  validador rejeita. Por isso `81P/81F/81C` usam o molde `FALH` (Custom) e não `81E1`.
- **`TOC` existe na base, mas é outro sinal** (`87 GROUND TOC FAULT`, RelayTrip).
  O certo para "26 - ALARME TEMP ÓLEO CDC" é o molde `TOA`.
- **AOR com espaço é o nome literal do grupo** (`CAS Trans`). Sem espaço, o motor
  prefixa o alias.
- **Device Mapping dos moldes é lixo** — vem da subestação de origem. Sempre
  sobrescrever.
- **Casar linha da lista com linha da TDT por NOME não funciona** — a coluna NOME
  é fórmula e há nomes repetidos. Casar por `(aba, linha)`.

---

## 6. Como rodar e conferir

```bash
cd backend
python make_casca.py     # gera TDT + lista corrigida + relatório
python check_casca.py    # confere tudo, de forma independente
```

`check_casca.py` valida em 7 etapas e precisa terminar com
`OK — nenhum problema encontrado`:

1. Lista: nenhuma linha de sinal sem índice; coordenadas únicas e contíguas por grupo.
2. TDT: nomes únicos, coordenadas únicas por espaço DNP3, nada vazio, RTU/AOR/container.
3. TDT × lista: todo sinal `SIM` está na TDT com a **mesma** coordenada.
4. Comandos: ida e volta (176/176), nenhum `Output` sobrando.
5. Descrição, escala e Multi Coord.
6. Nenhum sinal órfão na TDT.
7. Contagem por aba (as 21 abas batem).

### Estado atual

```
LISTA:  2535 linhas · 0 sem index
        D 0..1904 · A 1..361 · C 1..300 — contíguos, 0 duplicadas
TDT:    1282 sinais · 0 nome duplicado · 0 coord duplicada · 0 campo vazio
        858 mapeiam no unifilar (331 rebaixados) · 424 sem dispositivo
COMANDOS: 176/176 · 5 com Input de preenchimento (9599..9603)
```

---

## 7. Guia do relatório (`CASCA_RELATORIO.xlsx`)

| Aba | Conteúdo |
|---|---|
| `0-LEIA-ME` | O problema, a solução e as decisões. **Comece aqui.** |
| `1-Resumo` | Os números de uma olhada. |
| `2-DePara coordenadas` | Índice antigo → novo, sinal por sinal (2535). |
| `3-Coords REPETIDAS` | Evidência célula a célula das colisões da lista original. |
| `4-INDEX invalido` | Os pontos que estavam com `#REF!`. |
| `5-Siglas sem template` | Vazia = todas resolvidas. |
| `6-Renomeados` | Os 13 de nome duplicado: nome na lista × nome na TDT. |
| `7-Nomes duplicados` | Quais nomes se repetiam e em que abas. |
| `8-Siglas por equivalencia` | Sigla sem molde próprio e qual molde equivalente foi usado. |
| `9-Index antigo limpo` | Índices velhos removidos de linhas não utilizadas. |
| `10-Device Mapping` | Sinal por sinal: DM, origem da regra e situação no modelo. |
| `11-DM origem` | Resumo das regras, com o significado de cada uma. |
| `12-Comandos resolvidos` | Como cada comando achou seu sinal. |
| `13-CRIAR no Casca_Obra` | **Roteiro da obra:** cada dispositivo a criar, tipo e quantos sinais dependem. |
| `14-DM rebaixado` | Sinais que mapearam num dispositivo diferente do ideal — conferir. |
| `15-AVISOS p-o modelo` | Erros da planilha de origem. ALTA em amarelo. |
| `16-Equivalencia de modulos` | A tabela vão-da-lista → vão-do-unifilar, com evidência e confiança. |

---

## 8. O que fazer em seguida

1. **Conferir a aba 16** — as equivalências MÉDIA/BAIXA (TR6/TR7, barras, TSA)
   precisam do aval de quem conhece a subestação. Corrigir em `MODULO_EQUIV` e rodar de novo.
2. **Conferir a aba 14** — 331 sinais mapearam num dispositivo rebaixado.
3. **Levar a aba 15 pro autor da planilha** — os equipamentos repetidos entre `TR 1`
   e `TR 2` e o módulo `AL` nas abas de banco de capacitor mudam NOME e Device Mapping.
   Melhor acertar antes de criar os dispositivos.
4. **Criar os dispositivos da aba 13** no `Casca_Obra`. Conforme forem nascendo, os
   424 sinais pendentes casam sozinhos.
5. **Reimportar e comparar com o `erros.csv` novo** — a contagem de falhas de
   Device Mapping tem que cair para o número de sinais dos vãos ainda não criados.

---

## 9. Erro de importação já resolvido — leia antes de reimportar

O primeiro import (`ERROS2.csv`) devolveu **2711 falhas**. Parecia catástrofe, mas
era **um erro só, em cascata**:

```
[Falhou] UTR_CAS_3 · Container:
   Mandatory reference Casca_Obra not found in model.
   Element will not be inserted nor updated.
```

A RTU não entrou → **nenhum sinal dela entrou**. Todas as outras 2710 linhas eram
a mesma frase repetida: *"will not be imported in model since its connected (or
parent) element needed for its validity in model was not imported"*.

**Causa:** a aba `DNP3_RTUs` referenciava o container só pelo NOME.
No esqueleto da LVA o campo vinha preenchido em par — `Container Name =
'LAGOA VERMELHA 1'` **e** `Container Custom ID` — e eu estava zerando o Custom ID.

**Correção:** `casca_devmap.container_da_subestacao()` lê o elemento
`type="SUBSTATION"` do XML do modelo e devolve nome + `IDOBJ_CUSTOMID`
(`Casca_Obra` / `ba13733b-569f-43e6-a2fc-01f9a438096d`). Os dois vão na TDT.

### Segundo achado do mesmo CSV

```
[Informação] UTR_LVA_2_Link1__SAT Hughes:
   Found one TCP/IP Link with the same name in model. It will be updated with new values.
```

A aba `DNP3_TCPLinks` ainda carregava a linha do **esqueleto da LVA** — importar a
CASCA teria **alterado o link da UTR da LVA**. É o mesmo tipo de vazamento que já
tinha acontecido antes com AL21/AL22.

`DNP3_TCPLinks`, `DNP3_UDPLinks` e `DNP3_ScanGroups` agora são **esvaziadas**. A
lista da CASCA não traz IP nem Address (estão como `X` na aba `Informações`, só a
porta 20000 está definida), então não dá para montar o link certo — **o time de
comunicação cria depois**, ou entra numa TDT posterior quando os IPs saírem.

### Guarda no `check_casca.py`

A etapa 2 agora reprova se:

- a RTU estiver **sem Container Custom ID**;
- sobrar em `DNP3_TCPLinks` / `UDPLinks` / `ScanGroups` qualquer linha de **outra
  subestação**.

### Se o erro de container voltar

Se o ADMS continuar dizendo que `Casca_Obra` não existe mesmo com o Custom ID,
o motivo é outro: o changeset **`PT-MOD-SE-CASCA` está em DRAFT** (ver
`Changesets.csv`, estado `DRAFT` / "Loaded para edição"). Enquanto ele não for
aplicado, a subestação `Casca_Obra` só existe dentro daquele rascunho e nenhum
outro changeset a enxerga. Nesse caso: aplicar o changeset do modelo primeiro,
ou importar a TDT **dentro dele**.

---

## 10. Segunda rodada de import — `errpsmapping.csv`

Com o container resolvido, a RTU entrou e os sinais foram processados. Sobraram
**1000 erros, todos de `Device Mapping`**, em duas classes.

### 10.1 "Found multiple devices" — 665 sinais

O `Casca_Obra` é uma **cópia** da `CASCA`. Os dispositivos clonados **herdaram o
mesmo "ID de Mapeamento SCADA"** dos originais, então `CAS_AL12_52-2_DJ` responde
por **dois** disjuntores — o da CASCA e o do Casca_Obra — e o ADMS se recusa a
escolher.

**Correção:** preencher a coluna **`Substation`** de todo sinal com `Casca_Obra`.
Ela existe exatamente para restringir a busca a uma subestação. Nas TDTs antigas
essa coluna vinha vazia porque não havia clone e não havia ambiguidade.

### 10.2 Ambiguidade que a coluna `Substation` NÃO resolve — 63 sinais

**7 IDs estão repetidos dentro do próprio `Casca_Obra`** — os dois candidatos
moram na mesma subestação, então nenhuma coluna da TDT desempata:

| ID de Mapeamento SCADA | Dispositivos que respondem | Sinais |
|---|---|---|
| `CAS_LTPRI_52-21_DJ` | disjuntores `52-21_CAS` e `52-22_CAS` | 38 |
| `CAS_LTMRU_29-5_SEC` | seccionadoras `29-05_CAS` e `29-07_CAS` | 11 |
| `CAS_LTMRU_LTMRU_TC` | `CAS_LTKVM_LTKVM_TC` e `CAS_LTKVM_TC` | 10 |
| `CAS_LTMRU_LTMRU_TP` | `CAS_LTKVM_TP` e `CAS_LTKVM_LTKVM_TP` | 4 |
| `CAS_LTMRU_89-112_SEC` · `CAS_LTPRI_LTPRI_PROT_81_U1` · `CAS_LTPRI_LTPRI_PROT_81SU` | dois cada | — |

**Isso é resolvido no modelo, não na TDT:** cada dispositivo precisa de um ID
único (ex.: `CAS_LTPRI_52-22_DJ` para o segundo disjuntor). Aba **`17-ID duplicado
no modelo`** do relatório lista sinal por sinal.

### 10.3 Alias de outra subestação — 4 sinais

Quatro linhas da lista saíam com o alias **`IMA_`** em vez de `CAS_`:

```
BARRA!L51  IMA_BP113.8_BP113.8_FGOO      RET 1!L14  IMA_TSA_RET_NEGT
BARRA!L52  IMA_BP213.8_BP213.8_FGOO      RET 1!L15  IMA_TSA_RET_POST
```

A coluna SUBESTAÇÃO dessas linhas ficou de outra SE. O NOME e o Device Mapping
saíam errados. O gerador agora **normaliza o alias para `CAS`** e registra o
achado na aba `15`. Vale corrigir a célula na planilha de origem.

### 10.4 "Could not find any device" — 335 sinais

Os vãos que não existem no unifilar (`AL18`, `AL24`, `AL25`, `AL26`, `TRF29`,
`AL28`, `IB20`, `BP213.8`, `TSA2`). Já esperado — aba `13`.

### Resumo do que deve cair no próximo import

| Antes | Depois |
|---|---|
| 665 "multiple devices" | ~63 (só os 7 IDs duplicados dentro do modelo) |
| 4 com alias `IMA_` | 0 |
| 335 "could not find" | ~335, até os dispositivos da aba 13 serem criados |

---

## 11. A TDT original tem prioridade sobre qualquer regra

> "o device mapping é o mesmo da tdt original da subestação"

`casca_devmap.dm_da_tdt_original()` monta um índice **`(módulo, sigla) → Device
Mapping`** a partir da `CASCA.xlsx` (as abas IEC101/IEC104/DNP3 da subestação em
operação). Esse índice é consultado **antes de qualquer regra** — é o valor que
funciona em produção hoje.

**270 sinais** passaram a vir daí. Em **51 casos** o valor da TDT original é
diferente do que a regra deduziria — e a regra estava errada:

| Sinal | A regra deduzia | A TDT original usa |
|---|---|---|
| `IA` `IB` `IC` `P` `Q` dos alimentadores | `CAS_AL12_AL12_TC` | **`CAS_AL12_52-2_DJ`** |
| auxiliares de seccionadora do LT SCO | `CAS_LTSCO_89-102_SEC` | **`CAS_LTSCO_52-20_DJ`** |
| auxiliares do TR1 (AT e BT) | `CAS_TR1AT_89-12_SEC` | **`CAS_TR1_TR1_TR`** |
| TAP / comutador do TR1 | `CAS_TR1_TR1_TR` | **`CAS_TR1_TR1_COMTAP`** / `_TAP_REG` |

As medidas irem para o **disjuntor** do vão, e não para um TC, é justamente o
tipo de convenção que nenhuma regra adivinha.

### Efeito colateral importante

`CAS_TR1_TR1_COMTAP` e `CAS_TR1_TR1_TAP_REG` **não aparecem no XML do changeset**,
mas existem no modelo — o export é um **delta**, não o modelo inteiro. Por isso o
catálogo de dispositivos válidos agora é a **união** do XML com os Device Mappings
da TDT original. Sem isso, a aba `13` superestimava o que falta criar.

**Ordem final de resolução** (`casca_devmap.resolver`):

0. **TDT original da CASCA**, por `(módulo do unifilar, sigla)` ← novo
1. Texto idêntico a um ID do modelo
2. Equivalência de vão (`MODULO_EQUIV`, aba `16`)
3. Mesmo vão, equipamento renumerado
4. Relé específico inexistente → relé genérico `_PROT`
5. Dispositivo inexistente → equivalente do vão (aba `14`)
6. Vão inexistente → nome canônico (aba `13`)

---

## 12. Terceiro import — 1000 de 1282 sinais mapearam

`eros3.csv`: de 1000+ falhas para **286**. O que sobrou está em
`CASCA_STATUS_IMPORT.xlsx`, gerado por `backend/casca_status.py`, que cruza o
retorno do ADMS com a TDT.

> O CSV de erro é a **fonte mais confiável** sobre o modelo: se o sinal mapeou,
> o dispositivo existe e é único; se falhou, a mensagem diz se está duplicado ou
> ausente. Rode `python casca_status.py <erros.csv>` a cada import.

| Classe | Sinais | Onde se resolve |
|---|---|---|
| **MAPEOU** | **1000** | — |
| `Found multiple devices` | 201 | **modelo:** ID de Mapeamento SCADA único |
| `Could not find any device` | 77 | **modelo:** criar o dispositivo |
| `already has client points` | 4 | decidir (ver 12.3) |

### 12.1 Por vão

| Situação | Vãos |
|---|---|
| **completo** | TR7BT, IB20, TSA |
| quase (≤4 falhas) | AL12, AL24, AL25, AL26, TRF29, LT1, LT2, LT3, TR6, TR7, TR6AT, TR7AT, TR6BT, AL28 |
| **41 de 52 falhando** | **AL13, AL14, AL15, AL21** |
| barras | BP69, BP113.8, BP213.8 |
| serviço auxiliar | TSA1, TSA2 |

### 12.2 A pista mais importante: analógico passa, discreto não

No AL13, com o **mesmo** `CAS_AL13_52-3_DJ`:

```
OK      CAS_AL13_AL13_IA / IB / IC / P / Q      (analógicos)
FALHOU  CAS_AL13_52-13_DJF1 / 79 / FA / ...     (discretos)
```

A ambiguidade só derruba os **discretos**. E repare no contraste:

| Vão | Device Mapping usado | Discretos |
|---|---|---|
| AL24 (criado agora) | `CAS_AL24_52-24_DJ` — numeração **nova** | ✅ 21 mapearam |
| AL12 | `CAS_AL12_52-2_DJ` — numeração **antiga** | ✅ 24 mapearam |
| AL13/14/15/21 | `CAS_AL13_52-3_DJ` — numeração **antiga** | ❌ 20 falharam |

Ou seja: **os vãos que o César já ajustou no modelo passam**; AL13, AL14, AL15 e
AL21 ainda têm o ID duplicado entre a CASCA original e o Casca_Obra.

Isso decide uma pergunta que estava em aberto: **qual numeração o dispositivo do
Casca_Obra deve usar?** O AL24 mostra que os vãos novos usam a numeração **nova**
(`52-24`). Se AL13/14/15/21 forem criados com `52-13`/`52-14`/`52-15`/`52-21`,
some a ambiguidade de vez — mas aí a TDT precisa apontar pro nome canônico.
`MODULO_EQUIV` e a chave `REAPROVEITAR_DISPOSITIVO_ANTIGO` controlam isso.

### 12.3 `already has client points` — 4 sinais

```
CAS_TR6_TR6_CDC  ->  ja existe CAS_TR1_TR1_CDC em CAS_TR1_COMTAP
CAS_TR7_TR7_CDC  ->  ja existe CAS_TR2_TR2_CDC em CAS_TR2_COMTAP
CAS_TR6_TR6_R90  ->  ja existe CAS_TR1_TR1_R90 em RT1
CAS_TR7_TR7_R90  ->  ja existe CAS_TR2_TR2_R90 em RT2
```

O dispositivo já tem esse ponto vindo da **UTR antiga (IEC)**. O ADMS não
sobrescreve. Decidir se o ponto novo entra em outro dispositivo ou se o antigo
sai quando a UTR velha for desativada.

### 12.4 Como acompanhar

```bash
python backend/casca_status.py <erros.csv>
```

`CASCA_STATUS_IMPORT.xlsx`: `1-Resumo`, `2-Por modulo`, `3-Por dispositivo`
(cada Device Mapping com status e ação), `4-Sinais que falharam` e
`5-Dispositivos OK` (os que já estão certos — não mexer).

### 12.5 Coluna `Device` — a tentativa seguinte

> "é cas_obra a subestação. o dvc mapping continuou o mesmo pois é uma cópia da
> casca original"

Confirmado: o `Casca_Obra` **não renumerou nada** — herdou os mesmos IDs de
Mapeamento SCADA da CASCA. Então a ambiguidade é estrutural: o mesmo ID vale
para dois dispositivos, um em cada subestação. Nada disso é erro da TDT.

A prova de que o problema é do modelo, e não do nosso lado, é a simetria entre
AL12 e AL13 — sinais idênticos, estrutura idêntica, resultados opostos:

| Sigla | AL12 | AL13 |
|---|---|---|
| `DJF1` `79` `FA` `BBAB` … | ✅ `CAS_AL12_52-2_DJ` | ❌ `CAS_AL13_52-3_DJ` |
| `50F` `51F` `81U1..81U5` … | ✅ `CAS_AL12_52-2_PROT` | ❌ `CAS_AL13_52-3_PROT` |

O AL12 já está resolvido no modelo; AL13, AL14, AL15 e AL21 ainda não.

**O que foi acrescentado na TDT:** além de `Substation`, agora a coluna
**`Device`** leva o **nome do dispositivo** (`IDOBJ_NAME`) lido do XML do modelo
— `52-03_CAS`, `CAS_AL13_52-3_PROT`, etc. O par `Substation` + `Device` aperta a
busca muito mais que o ID sozinho.

- **782 sinais** saíram com a coluna `Device` preenchida
- **201 dos 282 que falharam** passaram a levar nome de dispositivo
- Nos **7 IDs repetidos dentro do próprio Casca_Obra** a coluna fica em branco
  de propósito: ali há dois candidatos legítimos (`52-21_CAS` e `52-22_CAS`,
  `29-05_CAS` e `29-07_CAS`…) e a lista não diz qual é qual — esses só se
  resolvem dando ID único no modelo.

Se mesmo com `Substation` + `Device` a ambiguidade persistir, aí não há mais o
que fazer pela TDT: cada dispositivo do `Casca_Obra` precisa de um ID de
Mapeamento SCADA próprio.

### 12.6 O nome da subestação estava errado

A árvore do ADMS mostra:

```
Planalto
 ├─ ARATIBA
 ├─ ARVOREZINHA
 ├─ Cas_Obra  (Substation)        <-- aqui
 │   └─ UTR_CAS_3
 └─ CASCA
```

A subestação chama **`Cas_Obra`**. O XML do changeset traz `Casca_Obra` no
`IDOBJ_NAME` — é o **rascunho**, não o que o modelo aplicado usa.

Isso passou despercebido porque a **RTU resolve o container pelo Custom ID**, e o
Custom ID estava certo — o container entrou mesmo com o nome errado. Mas a coluna
`Substation` dos **sinais** só tem o nome: com `Casca_Obra` o ADMS não conseguia
restringir a busca, e a ambiguidade `CASCA` × `Cas_Obra` continuava derrubando os
discretos.

Agora `CONTAINER = "Cas_Obra"` em `make_casca.py` alimenta os dois campos, e o
`check_casca.py` confere. **Lição:** o nome do modelo aplicado vem da árvore do
ADMS, não do XML do rascunho.

---

## 13. Conclusão: `Substation` e `Device` NÃO desempatam. É o modelo.

`erros4.csv` saiu com **exatamente as mesmas 286 falhas** de `eros3.csv` —
`0 passaram a mapear, 0 novas falhas` — mesmo com:

- `Substation = Cas_Obra` (o nome certo, da árvore do ADMS)
- coluna `Device` preenchida em 782 sinais com o nome do dispositivo

**O ADMS resolve o Device Mapping globalmente e ignora as colunas de
localização.** Não existe mais nada que a TDT possa fazer: enquanto o mesmo
"ID de Mapeamento SCADA" apontar para dois dispositivos, o sinal não mapeia.

As colunas ficam preenchidas (a informação está correta e não atrapalha), mas
não conte com elas para desempatar.

### 13.1 O que resolve — trabalho no modelo

**A) 42 dispositivos com ID duplicado → 201 sinais**

Dar um ID de Mapeamento SCADA **próprio** ao dispositivo do `Cas_Obra`.

| ID atual (repetido) | Sinais |
|---|---|
| `CAS_AL13_52-3_DJ` · `CAS_AL14_52-4_DJ` · `CAS_AL15_52-5_DJ` · `CAS_AL21_52-6_DJ` | 20 cada |
| `CAS_AL13_52-3_PROT` · `CAS_AL14_52-4_PROT` · `CAS_AL15_52-5_PROT` · `CAS_AL21_52-6_PROT` | 16 cada |
| `CAS_B138_B138_BP` | 12 |
| `CAS_BP23_29-36_SEC` | 11 |
| `CAS_BP23_BP23_BP` · `CAS_TSA3_RET_RET` | 2 cada |
| 30 relés/seccionadoras avulsos (`_PROT_2649`, `_PROT_51N`, `_PROT_50N`, `_PROT_51N1`, `29-14_SEC`…) | 1 cada |

> **O AL12 é a prova de que funciona.** Ele tem os mesmos sinais do AL13, com a
> mesma estrutura, e mapeia 48 de 52 — porque o ID dele já foi resolvido. Basta
> repetir em AL13, AL14, AL15 e AL21.

**B) 55 dispositivos a criar → 81 sinais**

Vãos que não existem no `Cas_Obra`: `AL18/24-1` (transferência), `TSA2/RET2`,
`BP213.8`, e relés soltos de `AL24`, `AL25`, `AL26`, `TRF29`, `AL18/52-18`,
`AL28` (`PROT_2649`, `PROT_51N`, `PROT_81SU`, `DJ`).

**C) 4 sinais com ponto já existente**

`CDC` e `R90` do TR6/TR7 — o dispositivo já tem o ponto vindo da UTR IEC antiga.

### 13.2 Placar

```
1282 sinais na TDT
1000 mapearam            (78%)
 201 ID duplicado no modelo   -> A
  81 dispositivo nao existe   -> B
   4 ponto ja existe          -> C
```

A lista completa, dispositivo por dispositivo, está em
`CASCA_STATUS_IMPORT.xlsx`, aba `3-Por dispositivo` (verde = pronto, amarelo =
parcial, vermelho = nada mapeou) e `4-Sinais que falharam`.

---

## 14. Remote Point Custom ID ordinal

Antes o campo era derivado do nome (`CAS_LT1_89-2_SECC_UTR_CAS_3`) — longo e com
risco de bater num remote point que já existe no modelo. Agora é uma sequência
limpa e própria da UTR nova:

```
Cas_obra_id_00001   CAS_LT1_89-2_SECC
Cas_obra_id_00002   CAS_LT1_89-2_43LR
...
Cas_obra_id_01282   CAS_TR7_TR7_TAP
```

- prefixo em `RPC_PREFIXO`, em `backend/make_casca.py`
- **1282 IDs, todos únicos**, faixa `00001..01282`, contígua
- o de-para completo está na aba **`18-Remote Point Custom ID`** do relatório
  (nome do sinal, sigla, aba e linha da lista de origem)
- `check_casca.py` reprova se algum fugir do padrão `Cas_obra_id_00000`,
  repetir, ou se a sequência tiver buraco

> Se a TDT anterior estiver **aberta no Excel** na hora de gerar, a saída vira
> `TDT_CASCA_UTR_CAS_3_NOVA.xlsx`. O `check_casca.py` detecta e confere a mais
> recente das duas — mas feche o Excel e rode de novo para consolidar o nome.

---

## 15. `erros6` — a fotografia real: 1282 de 1282 falharam

`erros6.1.csv` (282 linhas) e `erros6.2.csv` (1000 linhas) são **duas metades do
mesmo relatório** — o ADMS corta em 1000. Juntos: **1282 falhas, ou seja, TODOS
os sinais**. Interseção zero, união exata com a TDT.

```
848  Found multiple devices     (ID de Mapeamento SCADA duplicado)
424  Could not find any device  (dispositivo nao existe)
 10  outros                     (CDC/R90 e afins)
```

### 15.1 Por que piorou de 286 para 1282

A **única** mudança entre `erros4` e `erros6` foi o Remote Point Custom ID
(`{nome}_UTR_CAS_3` → `Cas_obra_id_00001`).

É por esse campo que o ADMS reconhece um remote point de importações
anteriores. Com o id novo, **todo sinal virou sinal novo** e o Device Mapping
foi re-resolvido do zero. Os ~1000 que "mapeavam" antes não estavam sendo
resolvidos a cada import — estavam **vinculados desde uma importação anterior**.

Os números fecham com o que o próprio gerador reporta:

| gerador | `erros6` |
|---|---|
| `sem dispositivo: 424` | **424** "Could not find any" |
| `mapeiam no unifilar: 858` | **848** multiple + **10** outros = 858 |

### 15.2 O que isso revela

**Todo Device Mapping que existe no modelo está duplicado.** Sem exceção — o
`Cas_Obra` é cópia integral da `CASCA` e herdou todos os IDs. Não existem 42
dispositivos com ID repetido como parecia: são **todos**.

E há uma pergunta que fica em aberto sobre os ~1000 vínculos antigos: em qual
das duas subestações eles caíram? Se caíram na `CASCA` (a que está em
operação), os sinais da UTR nova estão pendurados nos dispositivos errados e
ninguém veria. **Vale conferir alguns no ADMS antes de decidir.**

### 15.3 Duas TDTs para a decisão

Chave `RPC_ORDINAL` em `backend/make_casca.py`:

| Arquivo | Custom ID | Efeito |
|---|---|---|
| `TDT_CASCA_UTR_CAS_3.xlsx` | `Cas_obra_id_00001` | re-resolve tudo. Honesto, mas só entra depois que o modelo tiver IDs únicos |
| `TDT_CASCA_UTR_CAS_3_IDANTIGO.xlsx` | `{nome}_UTR_CAS_3` | reencontra os remote points antigos e devolve o placar de ~286 falhas |

As duas são idênticas em tudo o mais (mesmos 1282 sinais, mesmas coordenadas,
mesmos Device Mappings).

> **Recomendação:** o caminho de verdade continua sendo dar ID de Mapeamento
> SCADA único aos dispositivos do `Cas_Obra`. A `_IDANTIGO` só recupera um
> placar melhor no papel — os vínculos que ela preserva são de uma resolução
> antiga que ninguém auditou.

---

## 16. Signal Alias = data da leva

Convenção do usuário: **o Signal Alias é a data de hoje**, igual em todos os
sinais — serve para marcar a importação.

```
Signal Alias   23/07/2026     (os 1282 sinais)
Description    SECCIONADORA CARGA, 43 - CHAVE LOCAL REMOTO, ...
```

A **descrição do ponto**, que antes ia no Signal Alias, passou para a coluna
**`Description`** — senão o texto da lista se perderia. Os 1282 sinais têm as
duas colunas preenchidas.

`SIGNAL_ALIAS` em `backend/make_casca.py` recalcula na hora de gerar
(`date.today()`, formato `dd/mm/aaaa`). O `check_casca.py` confere que o alias
é uniforme, está no formato certo, e que a `Description` bate com a descrição
da lista de pontos.

---

## 17. Device Mapping com sufixo `_2`

Decisão para matar a ambiguidade de vez: **os dispositivos do `Cas_Obra`
recebem `_2` no ID de Mapeamento SCADA**, e a TDT aponta para o ID com sufixo.

```
CAS_AL13_52-3_DJ        ->  CAS_AL13_52-3_DJ_2
CAS_LTSCO_52-20_DJ      ->  CAS_LTSCO_52-20_DJ_2
CAS_AL24_52-24_DJ       ->  CAS_AL24_52-24_DJ_2
```

Funciona porque esse texto **só existe no `Cas_Obra`** — a `CASCA` original
continua com o ID sem sufixo, e o "Found multiple devices" deixa de acontecer.

- **1282 sinais, 100% com `_2`**, em **310 Device Mappings distintos**
- vale para todos: os que já existem no modelo (**renomear**) e os que ainda
  serão criados (**nascem com o sufixo**)
- constante `DM_SUFIXO` em `backend/make_casca.py` — pôr `""` volta ao anterior

### O que o César precisa fazer no modelo

A aba **`13-CRIAR no Casca_Obra`** já lista tudo com o `_2` no nome. E a aba
`3-Por dispositivo` do `CASCA_STATUS_IMPORT.xlsx`, cruzada com o próximo
`erros.csv`, mostra o que sobrou.

| Ação | Dispositivos |
|---|---|
| **Renomear** o ID para terminar em `_2` | os que hoje dão "Found multiple" |
| **Criar** já com `_2` | 185 da aba 13 |

Os 10 maiores por impacto:

```
CAS_LTSCO_52-20_DJ_2      53 sinais      CAS_TR1AT_89-12_SEC_2     35
CAS_LTPRI_52-21_DJ_2      41             CAS_TR1_TR1_TR_2          29
CAS_TR2AT_89-16_SEC_2     38             CAS_AL24_52-24_DJ_2       22
CAS_LTMRU_LTMRU_DJ_2      37             CAS_AL25_52-25_DJ_2       22
                                         CAS_AL26_52-26_DJ_2       22
                                         CAS_TRF29_24-2_DJ_2       22
```

---

## 18. TDT de um vão só

```bash
python make_casca.py --modulo LTPRI     # aceita o nome do unifilar
python make_casca.py --modulo LT2       # ou o da lista — dá no mesmo
```

Sai `TDT_CASCA_LT2.xlsx` + `CASCA_RELATORIO_LT2.xlsx`. **Não** regrava a lista
corrigida nem a TDT completa.

O recorte **não renumera nada**: coordenadas, Device Mapping, Remote Point
Custom ID e Signal Alias saem idênticos aos da TDT inteira. Conferido sinal a
sinal — 88/88 iguais. Por isso a parcial e a completa convivem no modelo.

> O ordinal do Remote Point Custom ID é calculado sobre a TDT **completa** e só
> depois filtrado (`rpc_por_linha`). Sem isso o recorte recomeçaria do
> `Cas_obra_id_00001` e colidiria com a TDT inteira.

### LT PRI (aba `LT 2` da lista) — 88 sinais

```
73 discretos + 15 analogicos
Input coords 15..157        88/88 com dispositivo no modelo, 0 pendente
```

| Device Mapping | Sinais |
|---|---|
| `CAS_LTPRI_52-21_DJ_2` | 41 |
| `CAS_LTPRI_LTPRI_PROT_2` | 10 |
| `CAS_LTPRI_LTPRI_TC_2` | 10 |
| `CAS_LTPRI_29-3_SEC_2` | 5 |
| `CAS_LTPRI_LTPRI_TP_2` | 4 |
| + 12 relés/seccionadoras com 1–2 sinais cada | 18 |

São **17 dispositivos** para renomear com `_2` no `Cas_Obra` — depois disso o
vão inteiro entra limpo. É o teste mais barato para validar a estratégia do
sufixo antes de fazer em massa.

---

## 19. `DM_ESTRITO` — só os Device Mappings que existem na CASCA

> "é problema de mapeamento. todos os device mapping necessarios estão aqui.
> não existe outros."

O `CASCA.xlsx` tem **332 Device Mappings `CAS_*`** (os outros ~158 são numéricos,
dos religadores `RAD_*`, e não servem). Com `DM_ESTRITO = True`:

- o Device Mapping sai **obrigatoriamente** desse catálogo — nada inventado
- **`DM_SUFIXO` voltou para `""`** (sem `_2`: aquele texto não existe na CASCA)
- o sinal que não acha alvo **fica de fora da TDT** e vai para a aba
  **`19-FORA (sem Device Mapping)`**

### Resultado

```
TDT_CASCA_UTR_CAS_3.xlsx   1282 sinais  (1040 D + 240 A + 2 A/D)
    858 em dispositivo do CASCA.xlsx   (101 DMs distintos)
    424 na UTR_CAS_3                   (vao ainda nao existe)
      0 Device Mapping fora do catalogo
```

Como cada DM foi escolhido:

| Origem | Sinais |
|---|---|
| a TDT atual usa esse DM para (vão, sigla) | 268 |
| mesmo vão e sufixo, equipamento renumerado | 248 |
| rebaixado para o disjuntor do vão | 158 |
| rebaixado para a seccionadora | 112 |
| barra / TC / trafo / retificador | 67 |
| exato | 1 |

### Os 424 sem dispositivo vão para a UTR

> "o q n tiver dvc mapping põe na utr"

Em vez de ficarem fora, esses sinais entram na TDT com
**`Device Mapping = UTR_CAS_3`** — pendurados na própria UTR. Depois é só
remapear cada um para o dispositivo certo quando ele existir, sem reimportar
nada. A aba `19-Mapeados na UTR` lista os 424, com o motivo de cada um.

São os vãos que **não existem na CASCA de hoje**:

| Vão | Sinais | O que é |
|---|---|---|
| `AL18` | 102 | BC 1 + transferência 24-1 |
| `AL24` `AL25` `AL26` `TRF29` | 55 cada | alimentadores e transf. 24-2 novos |
| `AL28` | 49 | BC 2 |
| `IB20` | 34 | interbarras BT |
| `BP213.8` | 14 | 2ª barra 13,8 kV |
| `TSA2` | 5 | 2º serviço auxiliar |

Os **71 comandos** desses sinais vêm junto — nada se perde. Quando os
dispositivos existirem, é só rodar `make_casca.py` de novo com um `CASCA.xlsx`
atualizado e eles saem da UTR para o dispositivo certo.

Chaves em `backend/make_casca.py`: `DM_ESTRITO` (só o catálogo) e `DM_NA_UTR`
(pôr na UTR em vez de deixar de fora).

> As coordenadas **não mudam** por causa do recorte — o sequenciamento continua
> global, sobre os 2535 pontos da lista. A lista corrigida segue valendo
> inteira, e o que entra na TDT agora é um subconjunto dela.
