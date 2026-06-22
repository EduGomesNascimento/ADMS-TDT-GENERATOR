<div align="center">

# 📝 SPEC — Revisão de Sinais
### Como o usuário confere, escolhe candidatos e confirma antes de gerar a TDT

`Gerador de TDTs ADMS` · especificação funcional · para revisar

</div>

---

## 1. Objetivo

Depois que o **motor de regras** (sem IA) mapeia a lista, o usuário **revisa**
numa única tela: vê o que foi analisado, **escolhe entre candidatos**, corrige
dúvidas e **confirma**. A TDT só é gerada do que foi validado.

> Princípio: **a ferramenta propõe, o engenheiro decide.** Nada vai pra TDT sem
> passar pela confiança do usuário.

---

## 2. O que cada linha mostra (transparência total)

Para o usuário ENTENDER por que chegou naquele sinal, a linha exibe **tudo**:

| Coluna | O que é | Origem |
|---|---|---|
| **Aba** | planilha de origem | lista |
| **Módulo** | equipamento/vão | lista |
| **Tipo** | A / D / C | lista |
| **Descrição (campo)** | texto cru do ponto | lista |
| **DNP3** | endereço de entrada/saída | lista |
| **SIGLA escolhida** ✎ | sugestão editável | motor de regras |
| **Descrição ADMS** | descrição da SIGLA atual | base ADMS |
| **Candidatos** | top-5 SIGLAs + score | ranker |
| **Conf.** | confiança da sugestão | motor |

---

## 3. Candidatos — o coração da revisão

Para **cada** sinal, o ranker calcula os **top-5 SIGLAs** possíveis (score
determinístico: sobreposição de tokens + código ANSI). São exibidos como
**chips clicáveis**:

```
SIGLA escolhida: [ VAB ▼ ]   ← editável (autocomplete de todas as SIGLAs)
  TENSAO VAB
  candidatos:  ▣VAB 24   ▢VB 23   ▢VC 23   ▢VABL 19   ▢VAB_BT 18
                 ↑ clicar troca a SIGLA escolhida
```

- O chip **selecionado** fica destacado (azul).
- Mesmo sinais **SEM match firme** recebem candidatos → o usuário escolhe o certo.
- Número = score (quanto maior, mais provável). Tooltip mostra a descrição.

---

## 4. Cores e confiança (semáforo)

| Cor da borda | Faixa | Significado | Ação esperada |
|---|---|---|---|
| 🟢 Verde | **ALTA** (≥90) | determinístico exato | aceitar |
| 🟡 Amarelo | **MÉDIA** (70-89) | provável (proteção/IA) | conferir |
| 🔴 Vermelho | **BAIXA** (<70) | incerto | escolher candidato |
| ⬜ Cinza | **SEM** | sem sugestão | escolher na mão |

Filtro **"Só pendentes"** esconde os verdes → o usuário foca no que importa.

---

## 5. Fluxo passo a passo

```
1. Analisar lista  ──►  motor de regras mapeia + calcula candidatos
2. Revisão:
     • verdes (ALTA) → já prontos
     • amarelos/vermelhos/cinzas → usuário confere:
         - clica num candidato, OU
         - digita a SIGLA, OU
         - apaga (exclui o sinal)
3. (A FAZER) Questionário de dúvidas:
     • antes de gerar, lista os que ficaram SEM/BAIXA não resolvidos
     • "Tem N sinais sem SIGLA. Deseja: [Pular] [Marcar como dúvida] [Voltar]"
4. Confirmar e gerar TDT  ──►  só os confirmados entram
5. TDT sai com linhas incertas em AMARELO (só a cor muda)
6. (A FAZER) Relatório do que foi confirmado/alterado/excluído
```

---

## 6. Questionário final de dúvidas (A FAZER)

Antes de gerar, um passo de **confirmação** dos pendentes:

- Mostra só os **SEM** e **BAIXA** não tocados pelo usuário.
- Cada um com **checkbox**: ☐ confirmar / ☐ marcar como dúvida / ☐ excluir.
- Botão **"confirmar todos os marcados"** → viram ALTA (validado por humano).
- O que ficar como "dúvida" sai **destacado em amarelo na TDT** + numa aba
  `Dúvidas` do relatório, pra resolver depois.

---

## 7. Destaque na TDT (já implementado)

- Linhas de sinais **não-ALTA** saem com fundo **amarelo claro** (`FFF2CC`).
- **Só a cor da linha muda** — formatação, fórmulas e validações preservadas.
- O ADMS aceita normalmente (só estilo de célula).

---

## 8. Aprendizado (A FAZER)

Toda correção do usuário (`descrição → SIGLA`) é salva em
`data/aprendizado.json` e vira a **camada 0** (a mais forte) nas próximas listas.
A ferramenta "aprende" o padrão da concessionária e erra cada vez menos.

---

## 9. Endpoints (backend)

| Método | Rota | Função |
|---|---|---|
| POST | `/api/raw/preview` | mapeia + retorna sinais com `candidates` e infos da linha |
| GET | `/api/raw/siglas` | lista de SIGLAs válidas (autocomplete) |
| POST | `/api/raw/export_reviewed` | gera a TDT só dos sinais confirmados |
| POST | `/api/raw/report` | relatório de probabilidades |

---

## 10. Estado

| Item | Status |
|---|---|
| Tabela com todas as infos da linha | ✅ feito |
| Candidatos clicáveis por sinal | ✅ feito |
| SIGLA editável + autocomplete | ✅ feito |
| Filtro "só pendentes" | ✅ feito |
| Confirmar e gerar (só validados) | ✅ feito |
| Destaque amarelo na TDT | ✅ feito |
| **Questionário final de dúvidas** | ⏳ a fazer |
| **Relatório de confirmações** | ⏳ a fazer |
| **Aprendizado (aprendizado.json)** | ⏳ a fazer |
