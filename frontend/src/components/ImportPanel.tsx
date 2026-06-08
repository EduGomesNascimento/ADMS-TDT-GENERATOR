import { useState } from "react";
import { Upload, FileSpreadsheet, Download, Loader2, CircleAlert, CheckCircle2, ArrowLeft, FileText } from "lucide-react";
import { importPreview, importExport, importReport, type ImportSummary } from "../lib/api";

export function ImportPanel({ onBack, protocol = "dnp3" }: { onBack: () => void; protocol?: string }) {
  const [file, setFile] = useState<File | null>(null);
  const [summary, setSummary] = useState<ImportSummary | null>(null);
  const [loading, setLoading] = useState(false);
  const [exporting, setExporting] = useState(false);
  const [reporting, setReporting] = useState(false);
  const [error, setError] = useState<string | null>(null);

  async function handleFile(f: File) {
    setFile(f); setSummary(null); setError(null); setLoading(true);
    try {
      setSummary(await importPreview(f, protocol));
    } catch (e: any) {
      setError(e.message);
    } finally {
      setLoading(false);
    }
  }

  async function doExport() {
    if (!file) return;
    setExporting(true); setError(null);
    try {
      await importExport(file, protocol);
    } catch (e: any) {
      setError(e.message);
    } finally {
      setExporting(false);
    }
  }

  async function doReport() {
    if (!file) return;
    setReporting(true); setError(null);
    try {
      await importReport(file, protocol);
    } catch (e: any) {
      setError(e.message);
    } finally {
      setReporting(false);
    }
  }

  const unmatched = summary
    ? [...summary.discrete.unmatched, ...summary.analog.unmatched]
    : [];

  return (
    <section className="glass rounded-2xl p-6">
      <h2 className="mb-1 flex items-center gap-2 text-lg font-semibold">
        <Upload size={18} className="text-brand-400" /> Importar lista de pontos
      </h2>
      <p className="mb-5 text-sm text-slate-400">
        Envie a lista padrão (.xlsx com abas <b>Discreto</b> e <b>Analógicos</b>). Cada SIGLA é
        casada com os campos reais da base ADMS; o NOME, índices DNP3 e AOR vêm da lista.
      </p>

      {error && (
        <div className="mb-4 flex items-center gap-2 rounded-lg border border-red-500/30 bg-red-500/10 px-4 py-3 text-sm text-red-300">
          <CircleAlert size={16} /> {error}
        </div>
      )}

      <label className="flex cursor-pointer flex-col items-center justify-center gap-2 rounded-xl border-2 border-dashed border-slate-700 bg-slate-900/40 px-6 py-10 text-center transition hover:border-brand-500/60 hover:bg-slate-900">
        <FileSpreadsheet size={32} className="text-brand-400" />
        <span className="text-sm text-slate-300">
          {file ? file.name : "Clique para selecionar a lista de pontos (.xlsx)"}
        </span>
        <input
          type="file"
          accept=".xlsx"
          className="hidden"
          onChange={(e) => e.target.files?.[0] && handleFile(e.target.files[0])}
        />
      </label>

      {loading && (
        <div className="mt-4 flex items-center gap-2 text-slate-400">
          <Loader2 className="animate-spin" size={18} /> Lendo e casando sinais…
        </div>
      )}

      {summary && (
        <div className="mt-5 space-y-4">
          <div className="grid gap-3 sm:grid-cols-3">
            <Stat label="Digitais" matched={summary.discrete.matched} total={summary.discrete.total} />
            <Stat label="Analógicos" matched={summary.analog.matched} total={summary.analog.total} />
            <div className="rounded-xl border border-slate-800 bg-slate-900/40 p-4">
              <p className="text-xs uppercase tracking-wide text-slate-500">Subestações</p>
              <p className="mt-1 font-mono text-brand-300">{summary.aliases.join(", ") || "—"}</p>
            </div>
          </div>

          {unmatched.length > 0 && (
            <div className="rounded-xl border border-amber-500/30 bg-amber-500/5 p-4">
              <p className="mb-2 flex items-center gap-2 text-sm font-medium text-amber-300">
                <CircleAlert size={15} /> {unmatched.length} sinais sem correspondência na base (serão omitidos)
              </p>
              <div className="max-h-32 overflow-auto text-xs text-slate-400">
                {unmatched.slice(0, 100).map((u, i) => (
                  <div key={i}>
                    <code className="text-amber-300">{u.sigla}</code> — {u.nome}
                  </div>
                ))}
              </div>
            </div>
          )}

          {summary.duplicates.length > 0 && (
            <div className="rounded-xl border border-amber-500/30 bg-amber-500/5 p-4">
              <p className="mb-2 flex items-center gap-2 text-sm font-medium text-amber-300">
                <CircleAlert size={15} /> {summary.duplicates.length} nomes duplicados na lista (gerarão linhas repetidas — revise a origem)
              </p>
              <div className="max-h-28 overflow-auto font-mono text-xs text-slate-400">
                {summary.duplicates.slice(0, 50).map((n, i) => <div key={i}>{n}</div>)}
              </div>
            </div>
          )}

          {summary.inputErrors.length > 0 && (
            <div className="rounded-xl border border-red-500/30 bg-red-500/5 p-4">
              <p className="mb-2 flex items-center gap-2 text-sm font-medium text-red-300">
                <CircleAlert size={15} /> {summary.inputErrors.length} sinais rejeitados na lista de origem (aba "Erros")
              </p>
              <div className="max-h-40 overflow-auto text-xs">
                <table className="w-full text-left">
                  <thead className="text-slate-500">
                    <tr><th className="pr-3">SIGLA</th><th className="pr-3">Tipo</th><th>Motivo</th></tr>
                  </thead>
                  <tbody className="text-slate-400">
                    {summary.inputErrors.slice(0, 100).map((e, i) => (
                      <tr key={i}>
                        <td className="pr-3"><code className="text-red-300">{e.sigla}</code></td>
                        <td className="pr-3">{e.tipo}</td>
                        <td>{e.motivo}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          )}

          <div className="flex flex-col-reverse items-stretch justify-between gap-3 sm:flex-row sm:items-center">
            <button onClick={onBack} className="btn-ghost justify-center sm:justify-start"><ArrowLeft size={16} /> Voltar</button>
            <div className="flex flex-col gap-2 sm:flex-row">
              <button onClick={doReport} disabled={reporting} className="btn-ghost justify-center border border-slate-700">
                {reporting ? <Loader2 className="animate-spin" size={18} /> : <FileText size={18} />}
                Baixar relatório
              </button>
              <button onClick={doExport} disabled={exporting} className="btn-primary justify-center">
                {exporting ? <Loader2 className="animate-spin" size={18} /> : <Download size={18} />}
                {exporting ? "Gerando…" : "Exportar TDT (.xlsx)"}
              </button>
            </div>
          </div>
        </div>
      )}

      {!summary && !loading && (
        <div className="mt-6">
          <button onClick={onBack} className="btn-ghost"><ArrowLeft size={16} /> Voltar</button>
        </div>
      )}
    </section>
  );
}

function Stat({ label, matched, total }: { label: string; matched: number; total: number }) {
  const ok = matched === total;
  return (
    <div className="rounded-xl border border-slate-800 bg-slate-900/40 p-4">
      <p className="text-xs uppercase tracking-wide text-slate-500">{label}</p>
      <p className="mt-1 flex items-center gap-2 text-xl font-bold text-slate-100">
        {ok ? <CheckCircle2 size={18} className="text-emerald-400" /> : null}
        {matched}<span className="text-sm font-normal text-slate-500">/ {total}</span>
      </p>
    </div>
  );
}
