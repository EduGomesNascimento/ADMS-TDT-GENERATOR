import { useRef, useState } from "react";
import {
  Upload, Loader2, Bot, Download, FileSpreadsheet,
  CheckCircle2, AlertTriangle, HelpCircle, ChevronDown, ChevronUp, ArrowLeft,
} from "lucide-react";
import clsx from "clsx";

const BASE = "/api";

// ─── Tipos ───────────────────────────────────────────────────────────────────

interface MappedSignal {
  utrId: string;
  description: string;
  dnp3Addr: number | null;
  signalType: string;
  module: string;
  sigla: string | null;
  siglaDesc: string | null;
  confidence: number;
  confidenceLabel: "ALTA" | "MÉDIA" | "BAIXA" | "SEM";
  matchMethod: string;
  sourceSheet?: string;
  candidates?: { sigla: string; desc: string; score: number }[];
}

interface RawPreview {
  detectedAlias: string;
  alias: string;
  discrete: { total: number; alta: number; media: number; baixa: number; sem: number };
  analog: { total: number; alta: number; media: number; baixa: number; sem: number };
  signals: MappedSignal[];
  llmNote?: string | null;
  llmSkipped?: number;
  usedLLM?: boolean;
}

// Opções de confiança mínima para incluir na TDT (tratamento de exceções)
const CONF_OPTIONS = [
  { value: 90, label: "Só ALTA — máxima confiabilidade",  hint: "Apenas os determinísticos exatos. Zero risco de erro, menor cobertura." },
  { value: 70, label: "ALTA + MÉDIA (recomendado)",        hint: "Inclui também os prováveis (IA/fuzzy). Bom equilíbrio." },
  { value: 60, label: "Tudo (≥60%) — máxima cobertura",   hint: "Inclui os incertos. Revise tudo antes de importar no ADMS." },
];

// ─── Helpers ─────────────────────────────────────────────────────────────────

const CONF_COLORS: Record<string, string> = {
  ALTA:  "bg-emerald-900/30 border-emerald-700/30",
  MÉDIA: "bg-amber-900/30  border-amber-700/30",
  BAIXA: "bg-red-900/30    border-red-700/30",
  SEM:   "bg-slate-800/40  border-slate-700/20",
};
const CONF_BADGE: Record<string, string> = {
  ALTA:  "text-emerald-300",
  MÉDIA: "text-amber-300",
  BAIXA: "text-red-300",
  SEM:   "text-slate-500",
};

// Apenas OFFLINE/LOCAL — a ferramenta funciona sem internet.
const PROVIDERS = [
  { value: "none",   label: "Sem IA — motor de regras (recomendado, offline)", model: "",          hint: "100% offline, reprodutível e sem limite. Mapeamento determinístico + escolha de candidatos na revisão. Não precisa de internet nem chave." },
  { value: "ollama", label: "IA local (Ollama, offline)",      model: "qwen2.5:7b",                hint: "Scan final OPCIONAL, 100% local. Requer Ollama instalado e 'ollama pull qwen2.5:7b'. Lento sem GPU. Nenhum dado sai da máquina." },
];

function downloadBlob(blob: Blob, filename: string) {
  const url = URL.createObjectURL(blob);
  const a   = document.createElement("a");
  a.href = url; a.download = filename;
  document.body.appendChild(a); a.click(); a.remove();
  URL.revokeObjectURL(url);
}

function StatBadge({ label, val, color }: { label: string; val: number; color: string }) {
  return (
    <div className={clsx("flex flex-col items-center rounded-lg border px-3 py-2 text-center", color)}>
      <span className="text-xs text-slate-400">{label}</span>
      <span className="text-xl font-bold">{val}</span>
    </div>
  );
}

// ─── Componente principal ─────────────────────────────────────────────────────

interface Props {
  onBack: () => void;
}

export function RawImportPanel({ onBack }: Props) {
  const fileRef = useRef<HTMLInputElement>(null);
  const [file,       setFile]       = useState<File | null>(null);
  const [alias,      setAlias]      = useState("");
  const [protocol,   setProtocol]   = useState("dnp3");
  const [providerIdx,setProviderIdx]= useState(0);
  const [apiKey,     setApiKey]     = useState("");
  const [ollamaUrl,  setOllamaUrl]  = useState("http://localhost:11434");
  const [loading,    setLoading]    = useState(false);
  const [error,      setError]      = useState<string | null>(null);
  const [result,     setResult]     = useState<RawPreview | null>(null);
  const [showAll,    setShowAll]    = useState(false);
  const [exporting,  setExporting]  = useState<"prob" | "tdt" | null>(null);
  const [minConfIdx, setMinConfIdx] = useState(1);   // padrão: ALTA+MÉDIA
  const [edits,      setEdits]      = useState<Record<number, string>>({});
  const [siglaOpts,  setSiglaOpts]  = useState<string[]>([]);
  const [onlyPending,setOnlyPending]= useState(false);
  const [genRev,     setGenRev]     = useState(false);

  const prov = PROVIDERS[providerIdx];
  const minConf = CONF_OPTIONS[minConfIdx];

  function makeForm(): FormData {
    const fd = new FormData();
    fd.append("file",     file!);
    fd.append("provider", prov.value);
    fd.append("model",    prov.model);
    fd.append("api_key",  apiKey);
    fd.append("alias",    alias.trim());
    fd.append("protocol", protocol);
    fd.append("min_confidence", String(minConf.value));
    if (prov.value === "ollama") fd.set("base_url", ollamaUrl);
    return fd;
  }

  async function analyse() {
    if (!file) return;
    setLoading(true); setError(null); setResult(null);
    try {
      const r = await fetch(`${BASE}/raw/preview`, { method: "POST", body: makeForm() });
      if (!r.ok) throw new Error((await r.text()) || "Falha na análise");
      const data: RawPreview = await r.json();
      setResult(data);
      setEdits({});
      if (!alias && data.detectedAlias) setAlias(data.detectedAlias);
      // carrega as SIGLAs válidas para o autocomplete da revisão
      fetch(`${BASE}/raw/siglas?protocol=${protocol}`)
        .then((rr) => rr.json())
        .then((d) => setSiglaOpts([...(d.discrete || []), ...(d.analog || [])].map((x: any) => x.sigla)))
        .catch(() => {});
    } catch (e: any) {
      setError(e.message);
    } finally {
      setLoading(false);
    }
  }

  // gera a TDT a partir dos sinais REVISADOS (sugestão + correções do usuário)
  async function genReviewed() {
    if (!result) return;
    setGenRev(true); setError(null);
    try {
      const reviewed = result.signals
        .map((m, idx) => ({
          module: m.module,
          signalType: m.signalType,
          sigla: (edits[idx] ?? m.sigla ?? "").trim(),
          dnp3Addr: m.dnp3Addr,
        }))
        .filter((s) => s.sigla);
      const r = await fetch(`${BASE}/raw/export_reviewed`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ alias: (alias.trim() || result.alias), protocol, signals: reviewed }),
      });
      if (!r.ok) throw new Error((await r.text()) || "Falha");
      const blob = await r.blob();
      const m = (r.headers.get("Content-Disposition") || "").match(/filename="?([^"]+)"?/);
      downloadBlob(blob, m ? m[1] : "TDT_revisada.xlsx");
    } catch (e: any) {
      setError(e.message);
    } finally {
      setGenRev(false);
    }
  }

  async function downloadProb() {
    if (!file) return;
    setExporting("prob");
    try {
      const r = await fetch(`${BASE}/raw/report`, { method: "POST", body: makeForm() });
      if (!r.ok) throw new Error((await r.text()) || "Falha");
      const blob = await r.blob();
      const disp = r.headers.get("Content-Disposition") || "";
      const m = disp.match(/filename="?([^"]+)"?/);
      downloadBlob(blob, m ? m[1] : "Probabilidades.xlsx");
    } catch (e: any) {
      setError(e.message);
    } finally {
      setExporting(null);
    }
  }

  async function downloadTdt() {
    if (!file) return;
    setExporting("tdt");
    try {
      const r = await fetch(`${BASE}/raw/export`, { method: "POST", body: makeForm() });
      if (!r.ok) throw new Error((await r.text()) || "Falha");
      const blob = await r.blob();
      const disp = r.headers.get("Content-Disposition") || "";
      const m = disp.match(/filename="?([^"]+)"?/);
      downloadBlob(blob, m ? m[1] : "TDT.xlsx");
    } catch (e: any) {
      setError(e.message);
    } finally {
      setExporting(null);
    }
  }

  const signals = result?.signals || [];
  // mantém o índice ORIGINAL (para rastrear edições) ao filtrar/cortar
  const rows = signals
    .map((m, idx) => ({ m, idx }))
    .filter(({ m }) => !onlyPending || m.confidenceLabel !== "ALTA");
  const displayed = showAll ? rows : rows.slice(0, 80);
  const reviewedCount = signals.filter((m, idx) => (edits[idx] ?? m.sigla ?? "").trim()).length;

  return (
    <div className="space-y-5">

      {/* Cabeçalho */}
      <div className="glass rounded-2xl p-6">
        <div className="mb-4 flex items-center justify-between">
          <h2 className="flex items-center gap-2 text-lg font-semibold">
            <Bot size={20} className="text-brand-400" />
            Lista não-padrão — Reconhecimento por IA
            <span className="ml-2 rounded-full bg-amber-500/20 px-2 py-0.5 text-[11px] font-semibold text-amber-300">em teste</span>
          </h2>
          <button onClick={onBack} className="btn-ghost text-sm">
            <ArrowLeft size={14} /> Voltar
          </button>
        </div>
        <p className="text-sm text-slate-400">
          Aceita <strong className="text-slate-200">qualquer planilha de UTR</strong> — sem formato específico.
          A IA identifica os sinais e os mapeia para a base ADMS com score de confiança.
        </p>
      </div>

      {/* Configuração */}
      <div className="glass rounded-2xl p-6 space-y-4">
        {/* Arquivo */}
        <div>
          <label className="mb-1 block text-xs font-semibold uppercase tracking-wide text-slate-400">
            Arquivo Excel (lista de pontos)
          </label>
          <div
            onClick={() => fileRef.current?.click()}
            className={clsx(
              "flex cursor-pointer items-center gap-3 rounded-xl border-2 border-dashed px-4 py-4 transition",
              file ? "border-brand-600/50 bg-brand-900/10" : "border-slate-700 hover:border-slate-500",
            )}
          >
            <FileSpreadsheet size={22} className="shrink-0 text-brand-400" />
            <div>
              <p className="text-sm font-medium text-slate-200">
                {file ? file.name : "Clique para selecionar o Excel"}
              </p>
              {file && (
                <p className="text-xs text-slate-500">{(file.size / 1024).toFixed(0)} KB</p>
              )}
            </div>
            <input
              ref={fileRef}
              type="file"
              accept=".xlsx,.xls"
              className="hidden"
              onChange={(e) => { setFile(e.target.files?.[0] ?? null); setResult(null); }}
            />
          </div>
        </div>

        {/* Alias + Protocolo */}
        <div className="grid grid-cols-2 gap-4">
          <div>
            <label className="mb-1 block text-xs font-semibold uppercase tracking-wide text-slate-400">
              Alias da subestação
            </label>
            <input
              type="text"
              value={alias}
              onChange={(e) => setAlias(e.target.value.toUpperCase())}
              placeholder="ex: GPR (auto-detectado)"
              className="w-full rounded-lg border border-slate-700 bg-slate-900 px-3 py-2 text-sm focus:border-brand-500 focus:outline-none"
            />
          </div>
          <div>
            <label className="mb-1 block text-xs font-semibold uppercase tracking-wide text-slate-400">
              Protocolo
            </label>
            <select
              value={protocol}
              onChange={(e) => setProtocol(e.target.value)}
              className="w-full rounded-lg border border-slate-700 bg-slate-900 px-3 py-2 text-sm focus:border-brand-500 focus:outline-none"
            >
              <option value="dnp3">DNP3</option>
              <option value="iec104">IEC 60870-5-104</option>
            </select>
          </div>
        </div>

        {/* Modelo */}
        <div>
          <label className="mb-1 block text-xs font-semibold uppercase tracking-wide text-slate-400">
            Modelo de IA
          </label>
          <select
            value={providerIdx}
            onChange={(e) => { setProviderIdx(Number(e.target.value)); setApiKey(""); }}
            className="w-full rounded-lg border border-slate-700 bg-slate-900 px-3 py-2 text-sm focus:border-brand-500 focus:outline-none"
          >
            {PROVIDERS.map((p, i) => (
              <option key={p.value} value={i}>{p.label}</option>
            ))}
          </select>
          <p className="mt-1 text-xs text-slate-500">{prov.hint}</p>
        </div>

        {/* API Key (Gemini / Groq) */}
        {prov.value !== "none" && prov.value !== "ollama" && (
          <div>
            <label className="mb-1 block text-xs font-semibold uppercase tracking-wide text-slate-400">
              API Key {<span className="normal-case font-normal text-slate-500">(ou use a variável de ambiente)</span>}
            </label>
            <input
              type="password"
              value={apiKey}
              onChange={(e) => setApiKey(e.target.value)}
              placeholder="Deixe vazio para usar variável de ambiente"
              className="w-full rounded-lg border border-slate-700 bg-slate-900 px-3 py-2 text-sm focus:border-brand-500 focus:outline-none font-mono"
            />
          </div>
        )}

        {/* Ollama URL */}
        {prov.value === "ollama" && (
          <div>
            <label className="mb-1 block text-xs font-semibold uppercase tracking-wide text-slate-400">
              URL do Ollama
            </label>
            <input
              type="text"
              value={ollamaUrl}
              onChange={(e) => setOllamaUrl(e.target.value)}
              className="w-full rounded-lg border border-slate-700 bg-slate-900 px-3 py-2 text-sm font-mono focus:border-brand-500 focus:outline-none"
            />
          </div>
        )}

        {/* Botão Analisar */}
        <button
          onClick={analyse}
          disabled={!file || loading}
          className="btn-primary w-full justify-center"
        >
          {loading ? (
            <><Loader2 size={16} className="animate-spin" /> Analisando…</>
          ) : (
            <><Bot size={16} /> Analisar lista</>
          )}
        </button>
      </div>

      {/* Erro */}
      {error && (
        <div className="flex items-start gap-2 rounded-xl border border-red-500/30 bg-red-500/10 px-4 py-3 text-sm text-red-300">
          <AlertTriangle size={16} className="mt-0.5 shrink-0" />
          <span>{error}</span>
        </div>
      )}

      {/* Resultado */}
      {result && (
        <>
          {/* Stats */}
          <div className="glass rounded-2xl p-6">
            <h3 className="mb-4 text-sm font-semibold uppercase tracking-wide text-slate-400">
              Resumo — {result.alias || result.detectedAlias}
            </h3>
            <div className="space-y-3">
              {(["discrete", "analog"] as const).map((k) => {
                const s = result[k];
                return (
                  <div key={k}>
                    <p className="mb-2 text-xs font-medium text-slate-400 uppercase tracking-wide">
                      {k === "discrete" ? "Discretos" : "Analógicos"} — {s.total} total
                    </p>
                    <div className="grid grid-cols-4 gap-2">
                      <StatBadge label="ALTA"  val={s.alta}  color="bg-emerald-900/30 border border-emerald-700/40 text-emerald-300" />
                      <StatBadge label="MÉDIA" val={s.media} color="bg-amber-900/30 border border-amber-700/40 text-amber-300" />
                      <StatBadge label="BAIXA" val={s.baixa} color="bg-red-900/30 border border-red-700/40 text-red-300" />
                      <StatBadge label="SEM"   val={s.sem}   color="bg-slate-800 border border-slate-700/40 text-slate-400" />
                    </div>
                  </div>
                );
              })}
            </div>

            {/* Aviso de exceção da IA (limite/erro) */}
            {result.llmNote && (
              <div className="mt-4 flex items-start gap-2 rounded-lg border border-amber-500/30 bg-amber-500/10 px-3 py-2 text-xs text-amber-300">
                <AlertTriangle size={14} className="mt-0.5 shrink-0" />
                <span>{result.llmNote}</span>
              </div>
            )}

            {/* OPÇÃO: confiança mínima para a TDT */}
            <div className="mt-4">
              <label className="mb-1 block text-xs font-semibold uppercase tracking-wide text-slate-400">
                O que incluir na TDT exportada
              </label>
              <select
                value={minConfIdx}
                onChange={(e) => setMinConfIdx(Number(e.target.value))}
                className="w-full rounded-lg border border-slate-700 bg-slate-900 px-3 py-2 text-sm focus:border-brand-500 focus:outline-none"
              >
                {CONF_OPTIONS.map((o, i) => (
                  <option key={o.value} value={i}>{o.label}</option>
                ))}
              </select>
              <p className="mt-1 text-xs text-slate-500">{minConf.hint}</p>
            </div>

            {/* Ações */}
            <div className="mt-4 flex gap-3">
              <button
                onClick={downloadProb}
                disabled={exporting !== null}
                className="btn-ghost flex-1 justify-center border border-slate-700"
              >
                {exporting === "prob"
                  ? <Loader2 size={14} className="animate-spin" />
                  : <Download size={14} />}
                Baixar Probabilidades
              </button>
              <button
                onClick={downloadTdt}
                disabled={exporting !== null}
                className="btn-primary flex-1 justify-center"
              >
                {exporting === "tdt"
                  ? <Loader2 size={14} className="animate-spin" />
                  : <FileSpreadsheet size={14} />}
                Exportar TDT (.xlsx)
              </button>
            </div>
            <p className="mt-2 text-xs text-slate-500">
              <strong className="text-emerald-400">ALTA</strong> = determinístico, confiável.{" "}
              <strong className="text-amber-400">MÉDIA/BAIXA</strong> = IA/aproximado, revise no relatório de Probabilidades antes de importar no ADMS.
            </p>
          </div>

          {/* Revisão dos sinais (pré-marcado, o usuário confere/corrige) */}
          <datalist id="siglaOpts">
            {siglaOpts.map((s) => <option key={s} value={s} />)}
          </datalist>
          <div className="glass rounded-2xl p-4">
            <div className="mb-3 flex flex-wrap items-center justify-between gap-2">
              <h3 className="text-sm font-semibold uppercase tracking-wide text-slate-400">
                Revisão ({rows.length} de {signals.length})
              </h3>
              <div className="flex items-center gap-3">
                <label className="flex cursor-pointer items-center gap-1.5 text-xs text-slate-400">
                  <input type="checkbox" checked={onlyPending} onChange={(e) => setOnlyPending(e.target.checked)} />
                  Só pendentes (MÉDIA/BAIXA/SEM)
                </label>
                <button
                  onClick={genReviewed}
                  disabled={genRev || reviewedCount === 0}
                  className="btn-primary text-xs"
                >
                  {genRev ? <Loader2 size={13} className="animate-spin" /> : <CheckCircle2 size={13} />}
                  Confirmar e gerar TDT ({reviewedCount})
                </button>
              </div>
            </div>
            <p className="mb-2 text-xs text-slate-500">
              A SIGLA já vem <strong className="text-slate-300">pré-marcada</strong>. Confira os{" "}
              <strong className="text-amber-400">amarelos/vermelhos</strong>; clique num{" "}
              <strong className="text-brand-300">candidato</strong> para escolher, digite outra SIGLA, ou apague para excluir.
            </p>
            <div className="overflow-x-auto">
              <table className="w-full text-xs">
                <thead>
                  <tr className="text-left text-slate-400 border-b border-slate-700/50">
                    <th className="pb-2 pr-2 font-semibold">Aba</th>
                    <th className="pb-2 pr-2 font-semibold">Módulo</th>
                    <th className="pb-2 pr-2 font-semibold">Tipo</th>
                    <th className="pb-2 pr-2 font-semibold">Descrição (campo)</th>
                    <th className="pb-2 pr-2 font-semibold">DNP3</th>
                    <th className="pb-2 pr-2 font-semibold">SIGLA escolhida ✎ + candidatos</th>
                    <th className="pb-2 font-semibold">Conf.</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-800/60">
                  {displayed.map(({ m, idx }) => {
                    const cur = edits[idx] ?? m.sigla ?? "";
                    const tipo = m.signalType === "analog" ? "A" : m.signalType === "command" ? "C" : "D";
                    return (
                    <tr
                      key={idx}
                      className={clsx(
                        "border-l-2 align-top",
                        m.confidenceLabel === "ALTA"  && "border-l-emerald-600",
                        m.confidenceLabel === "MÉDIA" && "border-l-amber-500",
                        m.confidenceLabel === "BAIXA" && "border-l-red-500",
                        m.confidenceLabel === "SEM"   && "border-l-slate-700",
                      )}
                    >
                      <td className="py-1.5 pr-2 text-slate-500 max-w-[80px] truncate" title={m.sourceSheet}>{m.sourceSheet || '—'}</td>
                      <td className="py-1.5 pr-2 text-slate-400">{m.module}</td>
                      <td className="py-1.5 pr-2 font-mono text-slate-500">{tipo}</td>
                      <td className="py-1.5 pr-2 text-slate-300 max-w-[220px]" title={m.description}>{m.description}</td>
                      <td className="py-1.5 pr-2 font-mono text-slate-400">{m.dnp3Addr ?? '—'}</td>
                      <td className="py-1 pr-2">
                        <input
                          list="siglaOpts"
                          value={cur}
                          onChange={(e) => setEdits((p) => ({ ...p, [idx]: e.target.value.toUpperCase() }))}
                          placeholder="—"
                          className={clsx(
                            "w-32 rounded border bg-slate-800/80 px-1.5 py-1 font-mono text-xs focus:border-brand-500 focus:outline-none",
                            cur ? "border-slate-700 text-brand-300" : "border-red-700/50 text-slate-500",
                            edits[idx] !== undefined && "border-amber-500/60",
                          )}
                        />
                        {m.siglaDesc && <div className="mt-0.5 text-[10px] text-slate-500 max-w-[200px] truncate" title={m.siglaDesc}>{m.siglaDesc}</div>}
                        {(m.candidates && m.candidates.length > 0) && (
                          <div className="mt-1 flex flex-wrap gap-1">
                            {m.candidates.slice(0, 5).map((c) => (
                              <button
                                key={c.sigla}
                                title={`${c.desc} · ${c.score}%`}
                                onClick={() => setEdits((p) => ({ ...p, [idx]: c.sigla }))}
                                className={clsx(
                                  "rounded px-1.5 py-0.5 font-mono text-[10px] transition",
                                  cur === c.sigla
                                    ? "bg-brand-600 text-white"
                                    : "bg-slate-700/60 text-slate-300 hover:bg-slate-600",
                                )}
                              >
                                {c.sigla}<span className="ml-1 opacity-60">{c.score}</span>
                              </button>
                            ))}
                          </div>
                        )}
                      </td>
                      <td className={clsx("py-1.5 font-semibold", CONF_BADGE[m.confidenceLabel])}>
                        {m.confidence > 0 ? `${m.confidence}%` : '—'}
                      </td>
                    </tr>
                  );})}
                </tbody>
              </table>
            </div>

            {rows.length > 80 && (
              <button
                onClick={() => setShowAll(!showAll)}
                className="mt-3 flex w-full items-center justify-center gap-1 text-xs text-slate-500 hover:text-slate-300"
              >
                {showAll
                  ? <><ChevronUp size={14} /> Mostrar menos</>
                  : <><ChevronDown size={14} /> Mostrar todos ({rows.length})</>}
              </button>
            )}
          </div>
        </>
      )}
    </div>
  );
}
