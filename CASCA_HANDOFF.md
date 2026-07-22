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
