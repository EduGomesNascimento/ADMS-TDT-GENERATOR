import { useMemo, useState } from "react";
import { Search, CheckSquare, Square, Activity, ToggleLeft, Terminal } from "lucide-react";
import clsx from "clsx";
import { type Signal, COMMAND_FORMATS } from "../lib/api";

export interface CmdCfg {
  format: string;
  outputCoord: string;
}

interface Props {
  title: string;
  sheet: string;
  kind: "discrete" | "analog";
  signals: Signal[];
  selected: Set<string>;
  cmdConfig: Record<string, CmdCfg>;
  onToggle: (key: string) => void;
  onToggleMany: (keys: string[], on: boolean) => void;
  onCmdChange: (key: string, patch: Partial<CmdCfg>) => void;
}

const phaseLabel = (p: string | null) => (p && p !== "None" ? p : "");

export function SignalList({
  title, sheet, kind, signals, selected, cmdConfig, onToggle, onToggleMany, onCmdChange,
}: Props) {
  const [q, setQ] = useState("");
  const filtered = useMemo(() => {
    const s = q.trim().toLowerCase();
    if (!s) return signals;
    return signals.filter(
      (sig) => sig.suffix.toLowerCase().includes(s) || sig.description.toLowerCase().includes(s)
    );
  }, [q, signals]);

  const keysFiltered = filtered.map((s) => `${sheet}|${s.suffix}`);
  const allSelected = keysFiltered.length > 0 && keysFiltered.every((k) => selected.has(k));
  const selCount = signals.filter((s) => selected.has(`${sheet}|${s.suffix}`)).length;
  const Icon = kind === "analog" ? Activity : ToggleLeft;

  return (
    <div className="flex h-full flex-col rounded-xl border border-slate-800 bg-slate-900/40">
      <div className="flex items-center justify-between gap-3 border-b border-slate-800 px-4 py-3">
        <div className="flex items-center gap-2">
          <Icon size={18} className={kind === "analog" ? "text-emerald-400" : "text-amber-400"} />
          <h3 className="font-semibold text-slate-100">{title}</h3>
          <span className="rounded-full bg-slate-800 px-2 py-0.5 text-xs text-slate-400">
            {selCount}/{signals.length}
          </span>
        </div>
        <button
          onClick={() => onToggleMany(keysFiltered, !allSelected)}
          className="flex items-center gap-1.5 rounded-lg px-2 py-1 text-xs text-brand-300 hover:bg-brand-500/10"
        >
          {allSelected ? <CheckSquare size={14} /> : <Square size={14} />}
          {allSelected ? "Limpar" : "Selecionar todos"}
        </button>
      </div>

      <div className="px-3 pt-3">
        <div className="relative">
          <Search size={15} className="absolute left-3 top-2.5 text-slate-500" />
          <input
            value={q}
            onChange={(e) => setQ(e.target.value)}
            placeholder="Buscar por código ou descrição…"
            className="w-full rounded-lg border border-slate-700 bg-slate-950/60 py-2 pl-9 pr-3 text-sm outline-none placeholder:text-slate-600 focus:border-brand-500"
          />
        </div>
      </div>

      <div className="flex-1 overflow-y-auto p-3 space-y-1.5" style={{ maxHeight: "46vh" }}>
        {filtered.length === 0 && (
          <p className="py-6 text-center text-sm text-slate-600">Nenhum sinal encontrado.</p>
        )}
        {filtered.map((sig) => {
          const key = `${sheet}|${sig.suffix}`;
          const on = selected.has(key);
          const cfg = cmdConfig[key] || { format: "template", outputCoord: "" };
          return (
            <div
              key={key}
              className={clsx(
                "rounded-lg border transition",
                on ? "border-brand-500/60 bg-brand-500/10" : "border-transparent bg-slate-800/40 hover:bg-slate-800"
              )}
            >
              <button onClick={() => onToggle(key)} className="flex w-full items-center gap-3 px-3 py-2 text-left">
                <span
                  className={clsx(
                    "flex h-5 w-5 shrink-0 items-center justify-center rounded border",
                    on ? "border-brand-400 bg-brand-500 text-white" : "border-slate-600"
                  )}
                >
                  {on && <CheckSquare size={13} />}
                </span>
                <div className="min-w-0 flex-1">
                  <div className="flex items-center gap-2">
                    <code className="rounded bg-slate-950/70 px-1.5 py-0.5 font-mono text-xs text-brand-300">
                      {sig.code || sig.suffix}
                    </code>
                    {sig.group && (
                      <span className="rounded bg-slate-700/60 px-1.5 py-0.5 text-[10px] text-slate-300">{sig.group}</span>
                    )}
                    <span className="truncate text-sm text-slate-200">{sig.description}</span>
                    {sig.hasCommand && (
                      <span className="flex items-center gap-1 rounded bg-fuchsia-500/15 px-1.5 py-0.5 text-[10px] font-medium text-fuchsia-300">
                        <Terminal size={10} /> CMD
                      </span>
                    )}
                  </div>
                  <div className="mt-0.5 flex flex-wrap gap-1.5 text-[11px] text-slate-500">
                    {sig.signalSubType && <span>{sig.signalSubType}</span>}
                    {sig.measurementType && <span>· {sig.measurementType}</span>}
                    {phaseLabel(sig.phases) && <span>· fases {phaseLabel(sig.phases)}</span>}
                  </div>
                </div>
              </button>

              {on && sig.hasCommand && (
                <div className="flex flex-wrap items-center gap-2 border-t border-brand-500/20 px-3 py-2 pl-11">
                  <span className="flex items-center gap-1 text-[11px] text-fuchsia-300">
                    <Terminal size={11} /> Comando:
                  </span>
                  <select
                    value={cfg.format}
                    onChange={(e) => onCmdChange(key, { format: e.target.value })}
                    className="rounded-md border border-slate-700 bg-slate-950/70 px-2 py-1 text-xs text-slate-200 outline-none focus:border-brand-500"
                  >
                    {COMMAND_FORMATS.map((f) => (
                      <option key={f.value} value={f.value}>{f.label}</option>
                    ))}
                  </select>
                  <input
                    value={cfg.outputCoord}
                    onChange={(e) => onCmdChange(key, { outputCoord: e.target.value })}
                    placeholder="coord. auto"
                    title="Output Coordinates (vazio = auto-sequência a partir da base de comando)"
                    className="w-28 rounded-md border border-slate-700 bg-slate-950/70 px-2 py-1 font-mono text-xs text-slate-200 outline-none placeholder:text-slate-600 focus:border-brand-500"
                  />
                  {sig.command?.controlCodes && cfg.format === "template" && (
                    <code className="text-[10px] text-slate-500">{sig.command.controlCodes}</code>
                  )}
                </div>
              )}
            </div>
          );
        })}
      </div>
    </div>
  );
}
