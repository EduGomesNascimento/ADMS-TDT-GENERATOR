# IA local e gratuita (Ollama) — reconhecimento de listas não-padrão

A ferramenta "Lista não-padrão — Reconhecimento por IA" pode usar um **LLM local,
100% gratuito e offline** via **Ollama** — nenhuma chave de API, nenhum dado sai
da sua máquina.

## 1. Instalar o Ollama

**Opção A (recomendada) — winget:**
```powershell
winget install Ollama.Ollama
```

**Opção B — instalador:** baixe em https://ollama.com/download e execute.

Depois de instalar, o Ollama roda sozinho em segundo plano (ícone na bandeja) e
escuta em `http://localhost:11434`.

## 2. Baixar um modelo

Abra o PowerShell e rode **uma** das opções abaixo conforme o seu PC:

| Modelo | Comando | RAM | Velocidade | Qualidade |
|--------|---------|-----|-----------|-----------|
| **qwen2.5:7b** (recomendado) | `ollama pull qwen2.5:7b` | ~8 GB | média (CPU) | alta |
| qwen2.5:3b (PC fraco) | `ollama pull qwen2.5:3b` | ~4 GB | rápida | boa |
| llama3.1:8b | `ollama pull llama3.1:8b` | ~8 GB | média | alta |

## ⚠️ Realidade de desempenho (CPU sem GPU)

Em PC **sem placa de vídeo dedicada** (como este: i5-1334U + Intel Iris Xe), a
inferência local é **lenta** e os modelos pequenos são **imprecisos**. Medido na
lista URA real:

| Modelo | Velocidade (CPU) | Qualidade nos sinais difíceis |
|--------|------------------|-------------------------------|
| qwen2.5:3b | ~90 s / 8 sinais | ruim (1/8, com erros) |
| qwen2.5:7b | ~2× mais lento | melhor, mas inviável p/ listas grandes |
| **Groq (nuvem, grátis)** | **~7 s / 25 sinais** | **boa (~21/25)** |

**Recomendação:** neste PC, use o **Groq** (nuvem, grátis e rápido) para listas
reais. O LLM **local** só compensa para listas **pequenas**, uso **offline
obrigatório**, ou em uma máquina **com GPU**. Sem LLM nenhum, as regras locais
(token + semântico) já resolvem ~57% **na hora**.

## 3. Usar na ferramenta

1. Abra o app (ABRIR.BAT) → tela inicial → **"Lista não-padrão — Reconhecimento por IA"**.
2. Selecione o Excel da lista de pontos.
3. Em **Modelo de IA**, escolha **"Ollama (local, grátis, offline)"**.
4. Confirme o modelo (`qwen2.5:7b`) e a URL (`http://localhost:11434`).
5. **Analisar lista** → revise o mapeamento → **Baixar Probabilidades** / **Exportar TDT**.

## Como funciona (camadas de mapeamento)

A IA só é chamada para o que as regras locais **não** resolvem sozinhas:

1. **Token** — código ADMS embutido no nome (ex.: `..._50F1` → `50F1`), respeitando o tipo.
2. **Semântico** — medições comuns por descrição (Tensão/Corrente/Potência →
   `VAB/VA/VB/VC/IA/IB/IC/IN/P/Q/S/F/TAP`). Funciona **sem** LLM.
3. **LLM (Ollama)** — só os restantes (proteções/comandos com nomes descritivos).

Resultado: a maior parte é resolvida sem IA; o LLM local fecha o resto de graça.

## Solução de problemas

- **"falha no mapeamento" / timeout:** o 1º uso baixa/carrega o modelo (pode demorar).
  Rode antes `ollama run qwen2.5:7b "ok"` no PowerShell para pré-carregar.
- **Ollama em outra máquina/porta:** ajuste a URL no painel, ou defina a variável de
  ambiente `OLLAMA_HOST` antes de subir o backend.
- **Sem internet só para o LLM:** depois do `ollama pull`, tudo roda offline.
