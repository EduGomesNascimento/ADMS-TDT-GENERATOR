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
}

interface RawPreview {
  detectedAlias: string;
  alias: string;
  discrete: { total: number; alta: number; media: number; baixa: number; sem: number };
  analog: { total: number; alta: number; media: number; baixa: number; sem: number };
  signals: MappedSignal[];
}

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

const PROVIDERS = [
  { value: "groq",   label: "Groq / Llama 3.3 70B (grátis, recomendado)", model: "llama-3.3-70b-versatile", hint: "Rápido e preciso. Chave grátis em console.groq.com — env: GROQ_API_KEY" },
  { value: "gemini", label: "Gemini 2.5 Flash (grátis)",     model: "gemini-2.5-flash",      hint: "Grátis em aistudio.google.com (250K tokens/min, 20 req/dia). Para mais volume use gemini-3.1-flash-lite (500/dia). NÃO use gemini-2.0-flash (cota zero)." },
  { value: "ollama", label: "Ollama (local, grátis, offline)", model: "qwen2.5:7b",               hint: "100% offline mas LENTO sem GPU. Em PC sem placa de vídeo, prefira o Groq. Requer 'ollama pull qwen2.5:7b'." },
  { value: "none",   label: "Só heurística (sem IA)",         model: "",                        hint: "Sem IA — token + semântico (~57%). Instantâneo e offline." },
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

  const prov = PROVIDERS[providerIdx];

  function makeForm(): FormData {
    const fd = new FormData();
    fd.append("file",     file!);
    fd.append("provider", prov.value);
    fd.append("model",    prov.model);
    fd.append("api_key",  apiKey);
    fd.append("alias",    alias.trim());
    fd.append("protocol", protocol);
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
      if (!alias && data.detectedAlias) setAlias(data.detectedAlias);
    } catch (e: any) {
      setError(e.message);
    } finally {
      setLoading(false);
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
  const displayed = showAll ? signals : signals.slice(0, 80);

  return (
    <div className="space-y-5">

      {/* Cabeçalho */}
      <div className="glass rounded-2xl p-6">
        <div className="mb-4 flex items-center justify-between">
          <h2 className="flex items-center gap-2 text-lg font-semibold">
            <Bot size={20} className="text-brand-400" />
            Lista não-padrão — Reconhecimento por IA
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

            {/* Ações */}
            <div className="mt-5 flex gap-3">
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
              A TDT inclui apenas sinais com confiança ≥ 60%. Revise o arquivo de Probabilidades antes de importar no ADMS.
            </p>
          </div>

          {/* Tabela de sinais */}
          <div className="glass rounded-2xl p-4">
            <h3 className="mb-3 text-sm font-semibold uppercase tracking-wide text-slate-400">
              Sinais mapeados ({signals.length})
            </h3>
            <div className="overflow-x-auto">
              <table className="w-full text-xs">
                <thead>
                  <tr className="text-left text-slate-400 border-b border-slate-700/50">
                    <th className="pb-2 pr-3 font-semibold">UTR ID</th>
                    <th className="pb-2 pr-3 font-semibold">Módulo</th>
                    <th className="pb-2 pr-3 font-semibold">Descrição</th>
                    <th className="pb-2 pr-3 font-semibold">SIGLA ADMS</th>
                    <th className="pb-2 pr-3 font-semibold">Descrição ADMS</th>
                    <th className="pb-2 pr-3 font-semibold">DNP3</th>
                    <th className="pb-2 font-semibold">Conf.</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-800/60">
                  {displayed.map((m, i) => (
                    <tr
                      key={i}
                      className={clsx(
                        "border-l-2",
                        m.confidenceLabel === "ALTA"  && "border-l-emerald-600",
                        m.confidenceLabel === "MÉDIA" && "border-l-amber-500",
                        m.confidenceLabel === "BAIXA" && "border-l-red-500",
                        m.confidenceLabel === "SEM"   && "border-l-slate-700",
                      )}
                    >
                      <td className="py-1.5 pr-3 font-mono text-slate-300">{m.utrId}</td>
                      <td className="py-1.5 pr-3 text-slate-400">{m.module}</td>
                      <td className="py-1.5 pr-3 text-slate-300 max-w-[180px] truncate" title={m.description}>{m.description}</td>
                      <td className="py-1.5 pr-3">
                        {m.sigla
                          ? <span className="rounded bg-brand-900/40 px-1.5 py-0.5 font-mono text-brand-300">{m.sigla}</span>
                          : <span className="text-slate-600">—</span>}
                      </td>
                      <td className="py-1.5 pr-3 text-slate-400">{m.siglaDesc || ''}</td>
                      <td className="py-1.5 pr-3 font-mono text-slate-400">{m.dnp3Addr ?? '—'}</td>
                      <td className={clsx("py-1.5 font-semibold", CONF_BADGE[m.confidenceLabel])}>
                        {m.confidence > 0 ? `${m.confidence}%` : '—'}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>

            {signals.length > 80 && (
              <button
                onClick={() => setShowAll(!showAll)}
                className="mt-3 flex w-full items-center justify-center gap-1 text-xs text-slate-500 hover:text-slate-300"
              >
                {showAll
                  ? <><ChevronUp size={14} /> Mostrar menos</>
                  : <><ChevronDown size={14} /> Mostrar todos ({signals.length})</>}
              </button>
            )}
          </div>
        </>
      )}
    </div>
  );
}
