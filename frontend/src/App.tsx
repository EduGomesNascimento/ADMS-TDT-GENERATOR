import { useEffect, useMemo, useState } from "react";
import {
  Zap, ArrowRight, ArrowLeft, Download, Loader2, FileSpreadsheet,
  Cpu, Settings2, ListChecks, Eye, CircleAlert, Hash, Upload, Network, Bot,
} from "lucide-react";
import clsx from "clsx";
import { Stepper } from "./components/Stepper";
import { SignalList, type CmdCfg } from "./components/SignalList";
import { ImportPanel } from "./components/ImportPanel";
import { RawImportPanel } from "./components/RawImportPanel";
import {
  fetchCatalog, fetchDevice, fetchPreview, exportTdt,
  type DeviceTypeSummary, type DeviceDetail, type PreviewResult, type ExportConfig, type SignalSel,
} from "./lib/api";

const STEPS = ["Equipamento", "Identificação", "Sinais", "Revisão"];

const typeIcon: Record<string, JSX.Element> = {
  alimentador: <Zap size={22} />,
  tsa: <Cpu size={22} />,
  transformador: <Cpu size={22} />,
};

export default function App() {
  const [protocol, setProtocol] = useState<"" | "dnp3" | "iec104">("");
  const [appMode, setAppMode] = useState<"wizard" | "import" | "raw">("wizard");
  const [step, setStep] = useState(1);
  const [maxReached, setMaxReached] = useState(1);
  const [catalog, setCatalog] = useState<DeviceTypeSummary[]>([]);
  const [loadingCat, setLoadingCat] = useState(true);
  const [error, setError] = useState<string | null>(null);

  const [typeId, setTypeId] = useState<string | null>(null);
  const [device, setDevice] = useState<DeviceDetail | null>(null);

  const [alias, setAlias] = useState("");
  const [moduleId, setModuleId] = useState("");
  const [deviceId, setDeviceId] = useState("");
  const [trNumber, setTrNumber] = useState("1");
  const [deviceName, setDeviceName] = useState("");

  const [customIdStart, setCustomIdStart] = useState(1);
  const [coordDiscrete, setCoordDiscrete] = useState(0);
  const [coordAnalog, setCoordAnalog] = useState(0);
  const [commandCoord, setCommandCoord] = useState(0);
  const [seqCoords, setSeqCoords] = useState(true);

  const [selected, setSelected] = useState<Set<string>>(new Set());
  const [cmdConfig, setCmdConfig] = useState<Record<string, CmdCfg>>({});
  const [preview, setPreview] = useState<PreviewResult | null>(null);
  const [exporting, setExporting] = useState(false);
  const [excelWarn, setExcelWarn] = useState<string | null>(null);

  useEffect(() => {
    fetchCatalog()
      .then((c) => setCatalog(c.deviceTypes))
      .catch((e) => setError(e.message))
      .finally(() => setLoadingCat(false));
    fetch("/api/health")
      .then((r) => r.json())
      .then((h) => { if (!h.excelNative) setExcelWarn(h.note || "MS Excel não detectado."); })
      .catch(() => {});
  }, []);

  const kind = device?.paramKind || (device?.consolidated ? "transformer" : "standard");
  const isTransformer = kind === "transformer";
  const isLine = kind === "line";
  const isBus = kind === "bus";
  const consolidated = isTransformer; // mantém compat. com envio do transformerNumber
  const identityValid = isTransformer
    ? !!(alias.trim() && trNumber.trim())
    : isLine || isBus
    ? !!(alias.trim() && moduleId.trim())
    : !!(alias.trim() && moduleId.trim() && deviceId.trim());
  const prefix = !identityValid
    ? ""
    : isTransformer
    ? `${alias.trim()}_TR${trNumber.trim()}`
    : isLine
    ? `${alias.trim()}_${moduleId.trim()}`
    : isBus
    ? `${alias.trim()}_${moduleId.trim()}_${moduleId.trim()}`
    : `${alias.trim()}_${moduleId.trim()}_${deviceId.trim()}`;

  const goto = (n: number) => {
    setStep(n);
    setMaxReached((m) => Math.max(m, n));
  };

  async function chooseType(t: DeviceTypeSummary) {
    setTypeId(t.id);
    setError(null);
    setSelected(new Set());
    setCmdConfig({});
    setModuleId(t.defaults?.module || "");
    setDeviceId(t.defaults?.device || "");
    setTrNumber(t.defaults?.transformerNumber || "1");
    try {
      const d = await fetchDevice(t.id);
      setDevice(d);
      goto(2);
    } catch (e: any) {
      setError(e.message);
    }
  }

  // ordena seleção conforme ordem do catálogo e injeta config de comando
  function orderedSignals(): SignalSel[] {
    if (!device) return [];
    const out: SignalSel[] = [];
    const push = (sheet: string, s: { suffix: string; hasCommand: boolean }) => {
      const k = `${sheet}|${s.suffix}`;
      if (!selected.has(k)) return;
      const sel: SignalSel = { sheet, suffix: s.suffix };
      if (s.hasCommand) {
        const cfg = cmdConfig[k];
        if (cfg?.format) sel.commandFormat = cfg.format;
        if (cfg?.outputCoord) sel.outputCoord = cfg.outputCoord;
      }
      out.push(sel);
    };
    device.signals.discrete.forEach((s) => push("DNP3_DiscreteSignals", s));
    device.signals.analog.forEach((s) => push("DNP3_AnalogSignals", s));
    (device.signals.discrete_analog || []).forEach((s) => push("DNP3_DiscreteAnalog", s));
    return out;
  }

  const config: ExportConfig | null = useMemo(() => {
    if (!typeId || !identityValid) return null;
    return {
      deviceTypeId: typeId,
      alias: alias.trim(),
      module: isTransformer ? undefined : moduleId.trim(),
      device: isBus ? moduleId.trim() : (isTransformer || isLine ? undefined : deviceId.trim()),
      transformerNumber: isTransformer ? trNumber.trim() : undefined,
      deviceName: deviceName.trim() || undefined,
      customIdStart,
      coordStart: seqCoords ? { discrete: coordDiscrete, analog: coordAnalog } : undefined,
      commandCoordStart: seqCoords ? commandCoord : undefined,
      signals: [],
    };
  }, [typeId, identityValid, consolidated, alias, moduleId, deviceId, trNumber, deviceName, customIdStart, coordDiscrete, coordAnalog, commandCoord, seqCoords]);

  async function loadPreview() {
    if (!config) return;
    try {
      const ordered = { ...config, signals: orderedSignals() };
      const p = await fetchPreview(ordered);
      setPreview(p);
      goto(4);
    } catch (e: any) {
      setError(e.message);
    }
  }

  async function doExport() {
    if (!config) return;
    setExporting(true);
    setError(null);
    try {
      await exportTdt({ ...config, signals: orderedSignals() });
    } catch (e: any) {
      setError(e.message);
    } finally {
      setExporting(false);
    }
  }

  function ensureCmd(key: string) {
    setCmdConfig((prev) => (prev[key] ? prev : { ...prev, [key]: { format: "template", outputCoord: "" } }));
  }
  function toggle(key: string) {
    setSelected((prev) => {
      const next = new Set(prev);
      next.has(key) ? next.delete(key) : next.add(key);
      return next;
    });
    ensureCmd(key);
  }
  function toggleMany(keys: string[], on: boolean) {
    setSelected((prev) => {
      const next = new Set(prev);
      keys.forEach((k) => (on ? next.add(k) : next.delete(k)));
      return next;
    });
    if (on) keys.forEach(ensureCmd);
  }
  function onCmdChange(key: string, patch: Partial<CmdCfg>) {
    setCmdConfig((prev) => ({ ...prev, [key]: { ...(prev[key] || { format: "template", outputCoord: "" }), ...patch } }));
  }

  const selCount = selected.size;

  return (
    <div className="min-h-full bg-gradient-to-b from-[#0a0f1e] via-[#0a0f1e] to-[#0c1226]">
      <div className="pointer-events-none fixed inset-0 bg-[radial-gradient(ellipse_at_top,rgba(37,99,235,0.12),transparent_55%)]" />
      <div className="relative mx-auto max-w-6xl px-4 py-8">
        {/* Header */}
        <header className="mb-8 flex flex-col items-center text-center">
          <div className="mb-3 flex items-center gap-3">
            <div className="flex h-11 w-11 items-center justify-center rounded-xl bg-brand-600 shadow-lg shadow-brand-600/30">
              <FileSpreadsheet size={22} />
            </div>
            <div className="text-left">
              <h1 className="text-2xl font-bold tracking-tight">Gerador de TDTs · ADMS</h1>
              <p className="text-sm text-slate-400">
                Monte os sinais de um equipamento e exporte a TDT no formato exato do ADMS.
              </p>
            </div>
          </div>
          {protocol && (
            <div className="mb-3 flex items-center gap-2 text-sm">
              <span className="rounded-full bg-brand-600/20 px-3 py-1 font-mono font-semibold text-brand-300">
                {protocol.toUpperCase()}
              </span>
              <button onClick={() => { setProtocol(""); setAppMode("wizard"); setStep(1); }} className="text-slate-500 hover:text-slate-300">
                trocar protocolo
              </button>
            </div>
          )}
          {protocol && appMode === "wizard" && (
            <Stepper steps={STEPS} current={step} maxReached={maxReached} onJump={goto} />
          )}
        </header>

        {!protocol && (
          <section className="glass rounded-2xl p-6">
            <h2 className="mb-1 flex items-center gap-2 text-lg font-semibold">
              <Network size={18} className="text-brand-400" /> Escolha o protocolo de comunicação
            </h2>
            <p className="mb-5 text-sm text-slate-400">
              O fluxo e o formato da TDT seguem o protocolo escolhido.
            </p>
            <div className="grid gap-4 sm:grid-cols-2">
              {[
                { id: "dnp3",   label: "DNP3",              desc: "Assistente por equipamento + importação de lista de pontos." },
                { id: "iec104", label: "IEC 60870-5-104",   desc: "Importação de lista de pontos (TDT IEC104)." },
              ].map((p) => (
                <button
                  key={p.id}
                  onClick={() => { setProtocol(p.id as "dnp3" | "iec104"); setAppMode(p.id === "iec104" ? "import" : "wizard"); }}
                  className="flex flex-col gap-2 rounded-xl border border-slate-800 bg-slate-900/40 p-5 text-left transition hover:border-brand-600/50 hover:bg-slate-900"
                >
                  <div className="flex items-center gap-3">
                    <div className="flex h-11 w-11 items-center justify-center rounded-lg bg-brand-600/20 text-brand-300">
                      <Network size={22} />
                    </div>
                    <h3 className="text-lg font-semibold text-slate-100">{p.label}</h3>
                  </div>
                  <p className="text-sm text-slate-400">{p.desc}</p>
                </button>
              ))}
            </div>
            {/* Lista não-padrão (IA) */}
            <div className="mt-4 border-t border-slate-800 pt-4">
              <button
                onClick={() => { setProtocol("dnp3"); setAppMode("raw"); }}
                className="flex w-full items-center gap-4 rounded-xl border border-slate-700/50 bg-slate-900/30 p-4 text-left transition hover:border-brand-600/40 hover:bg-slate-900"
              >
                <div className="flex h-10 w-10 shrink-0 items-center justify-center rounded-lg bg-violet-600/20 text-violet-300">
                  <Bot size={20} />
                </div>
                <div>
                  <p className="font-semibold text-slate-200">
                    Lista não-padrão — Reconhecimento por IA{" "}
                    <span className="ml-1 rounded-full bg-amber-500/20 px-2 py-0.5 align-middle text-[11px] font-semibold text-amber-300">em teste</span>
                  </p>
                  <p className="text-xs text-slate-400">
                    Aceita qualquer planilha de UTR (sem formato específico). A IA mapeia os sinais para a base ADMS com score de confiança.
                  </p>
                </div>
              </button>
            </div>
          </section>
        )}

        {excelWarn && (
          <div className="mb-4 flex items-start gap-2 rounded-lg border border-amber-500/30 bg-amber-500/10 px-4 py-3 text-sm text-amber-300">
            <CircleAlert size={16} className="mt-0.5 shrink-0" />
            <span>{excelWarn}</span>
          </div>
        )}

        {protocol && error && (
          <div className="mb-4 flex items-center gap-2 rounded-lg border border-red-500/30 bg-red-500/10 px-4 py-3 text-sm text-red-300">
            <CircleAlert size={16} /> {error}
          </div>
        )}

        {protocol && appMode === "import" && (
          <ImportPanel
            protocol={protocol}
            onBack={() => (protocol === "iec104" ? setProtocol("") : setAppMode("wizard"))}
          />
        )}

        {protocol && appMode === "raw" && (
          <RawImportPanel onBack={() => { setProtocol(""); setAppMode("wizard"); }} />
        )}

        {/* STEP 1 */}
        {protocol && appMode === "wizard" && step === 1 && (
          <section className="glass rounded-2xl p-6">
            <div className="mb-5 flex items-center justify-between gap-3">
              <div>
                <h2 className="mb-1 flex items-center gap-2 text-lg font-semibold">
                  <ListChecks size={18} className="text-brand-400" /> Escolha o tipo de equipamento
                </h2>
                <p className="text-sm text-slate-400">
                  Os sinais são aprendidos automaticamente das TDTs reais da base.
                </p>
              </div>
              <button onClick={() => setAppMode("import")} className="btn-ghost shrink-0 border border-slate-700">
                <Upload size={16} /> Importar lista de pontos
              </button>
            </div>
            {loadingCat ? (
              <div className="flex items-center gap-2 text-slate-400">
                <Loader2 className="animate-spin" size={18} /> Carregando catálogo…
              </div>
            ) : (
              <div className="grid gap-4 sm:grid-cols-2">
                {catalog.map((t) => (
                  <button
                    key={t.id}
                    onClick={() => chooseType(t)}
                    className={clsx(
                      "group flex flex-col gap-3 rounded-xl border p-5 text-left transition",
                      typeId === t.id
                        ? "border-brand-500 bg-brand-500/10"
                        : "border-slate-800 bg-slate-900/40 hover:border-brand-600/50 hover:bg-slate-900"
                    )}
                  >
                    <div className="flex items-center gap-3">
                      <div className="flex h-11 w-11 items-center justify-center rounded-lg bg-brand-600/20 text-brand-300">
                        {typeIcon[t.id] ?? <Cpu size={22} />}
                      </div>
                      <div>
                        <h3 className="font-semibold text-slate-100">{t.label}</h3>
                        <code className="font-mono text-xs text-slate-500">{t.source}</code>
                      </div>
                    </div>
                    <p className="text-sm text-slate-400">{t.description}</p>
                    <div className="mt-1 flex gap-2 text-xs">
                      <span className="rounded-full bg-amber-500/15 px-2 py-0.5 text-amber-300">
                        {t.signalCount.discrete} digitais
                      </span>
                      <span className="rounded-full bg-emerald-500/15 px-2 py-0.5 text-emerald-300">
                        {t.signalCount.analog} analógicos
                      </span>
                    </div>
                  </button>
                ))}
              </div>
            )}
          </section>
        )}

        {/* STEP 2 */}
        {step === 2 && device && (
          <section className="glass rounded-2xl p-6">
            <h2 className="mb-1 flex items-center gap-2 text-lg font-semibold">
              <Settings2 size={18} className="text-brand-400" /> Identificação do equipamento
            </h2>
            <p className="mb-5 text-sm text-slate-400">
              Preencha alias, módulo e identificador. Os sinais só são liberados após o preenchimento.
            </p>

            <div className="grid gap-5 md:grid-cols-2">
              <Field label="Nome do dispositivo (opcional)" hint="Rótulo descritivo, ex.: Alimentador Centro 05">
                <input className="inp" value={deviceName} onChange={(e) => setDeviceName(e.target.value)} placeholder="Alimentador Centro 05" />
              </Field>
              <Field label="Alias da subestação *" hint="Prefixo do sinal. Ex.: FWB, GTD, CNO_NEW">
                <input className="inp font-mono" value={alias} onChange={(e) => setAlias(e.target.value.toUpperCase())} placeholder="FWB" />
              </Field>
              {isTransformer ? (
                <Field label="Número do transformador (N) *" hint="Ex.: 1 → TR1/TR1AT/TR1BT; 2 → TR2…">
                  <input className="inp font-mono" value={trNumber} onChange={(e) => setTrNumber(e.target.value.replace(/[^0-9]/g, ""))} placeholder="1" />
                </Field>
              ) : isLine ? (
                <Field label="Nome da linha *" hint="Identificador da LT. Ex.: LTABC, LTKGT">
                  <input className="inp font-mono" value={moduleId} onChange={(e) => setModuleId(e.target.value.toUpperCase())} placeholder="LTABC" />
                </Field>
              ) : isBus ? (
                <Field label="Nome da barra *" hint="Identificador do barramento. Ex.: BP138, BP69">
                  <input className="inp font-mono" value={moduleId} onChange={(e) => setModuleId(e.target.value.toUpperCase())} placeholder="BP138" />
                </Field>
              ) : (
                <>
                  <Field label="Módulo / Alimentador *" hint="Ex.: AL05, TSA">
                    <input className="inp font-mono" value={moduleId} onChange={(e) => setModuleId(e.target.value.toUpperCase())} placeholder="AL05" />
                  </Field>
                  <Field label="Identificador do equipamento *" hint="Ex.: 52-05 (disjuntor), 24-1">
                    <input className="inp font-mono" value={deviceId} onChange={(e) => setDeviceId(e.target.value.toUpperCase())} placeholder="52-05" />
                  </Field>
                </>
              )}
            </div>

            <div className="mt-5 rounded-xl border border-slate-800 bg-slate-950/50 p-4">
              <div className="mb-3 flex items-center gap-2 text-sm font-medium text-slate-300">
                <Hash size={15} className="text-brand-400" /> Numeração & endereços DNP3
              </div>
              <div className="grid gap-4 sm:grid-cols-2 lg:grid-cols-4">
                <Field label="Custom ID inicial">
                  <input type="number" className="inp" value={customIdStart} onChange={(e) => setCustomIdStart(+e.target.value)} />
                </Field>
                <Field label="Coord. base (digitais)">
                  <input type="number" disabled={!seqCoords} className="inp disabled:opacity-40" value={coordDiscrete} onChange={(e) => setCoordDiscrete(+e.target.value)} />
                </Field>
                <Field label="Coord. base (analógicos)">
                  <input type="number" disabled={!seqCoords} className="inp disabled:opacity-40" value={coordAnalog} onChange={(e) => setCoordAnalog(+e.target.value)} />
                </Field>
                <Field label="Coord. base (comandos)">
                  <input type="number" disabled={!seqCoords} className="inp disabled:opacity-40" value={commandCoord} onChange={(e) => setCommandCoord(+e.target.value)} />
                </Field>
              </div>
              <label className="mt-3 flex cursor-pointer items-center gap-2 text-sm text-slate-400">
                <input type="checkbox" checked={seqCoords} onChange={(e) => setSeqCoords(e.target.checked)} className="accent-brand-500" />
                Auto-sequenciar coordenadas a partir da base (desmarque para manter os endereços do template)
              </label>
            </div>

            {identityValid && (
              <div className="mt-5 rounded-xl border border-brand-600/30 bg-brand-600/5 p-4">
                <p className="text-xs uppercase tracking-wide text-slate-500">Prévia do nome do sinal</p>
                <code className="mt-1 block font-mono text-brand-300">
                  {prefix}_<span className="text-slate-500">SUFIXO</span>
                </code>
                <p className="mt-1 text-xs text-slate-500">Remote Unit: <code className="font-mono text-slate-400">UTR_{alias.trim()}_1</code></p>
              </div>
            )}

            <NavButtons
              onBack={() => goto(1)}
              onNext={() => goto(3)}
              nextDisabled={!identityValid}
              nextLabel="Selecionar sinais"
            />
          </section>
        )}

        {/* STEP 3 */}
        {step === 3 && device && (
          <section className="glass rounded-2xl p-6">
            <div className="mb-4 flex flex-wrap items-center justify-between gap-3">
              <div>
                <h2 className="flex items-center gap-2 text-lg font-semibold">
                  <ListChecks size={18} className="text-brand-400" /> Sinais de <code className="font-mono text-brand-300">{prefix}</code>
                </h2>
                <p className="text-sm text-slate-400">{selCount} sinais selecionados</p>
              </div>
            </div>
            <div className="grid gap-4 lg:grid-cols-2">
              <SignalList
                title="Sinais Digitais"
                kind="discrete"
                sheet="DNP3_DiscreteSignals"
                signals={device.signals.discrete}
                selected={selected}
                cmdConfig={cmdConfig}
                onToggle={toggle}
                onToggleMany={toggleMany}
                onCmdChange={onCmdChange}
              />
              <SignalList
                title="Sinais Analógicos"
                kind="analog"
                sheet="DNP3_AnalogSignals"
                signals={device.signals.analog}
                selected={selected}
                cmdConfig={cmdConfig}
                onToggle={toggle}
                onToggleMany={toggleMany}
                onCmdChange={onCmdChange}
              />
              {device.signals.discrete_analog && device.signals.discrete_analog.length > 0 && (
                <div className="lg:col-span-2">
                  <SignalList
                    title="Sinais Digital-Analógicos (A/D)"
                    kind="analog"
                    sheet="DNP3_DiscreteAnalog"
                    signals={device.signals.discrete_analog}
                    selected={selected}
                    cmdConfig={cmdConfig}
                    onToggle={toggle}
                    onToggleMany={toggleMany}
                    onCmdChange={onCmdChange}
                  />
                </div>
              )}
            </div>
            <NavButtons
              onBack={() => goto(2)}
              onNext={loadPreview}
              nextDisabled={selCount === 0}
              nextLabel="Revisar & exportar"
            />
          </section>
        )}

        {/* STEP 4 */}
        {step === 4 && preview && (
          <section className="glass rounded-2xl p-6">
            <h2 className="mb-1 flex items-center gap-2 text-lg font-semibold">
              <Eye size={18} className="text-brand-400" /> Revisão da TDT
            </h2>
            <p className="mb-4 text-sm text-slate-400">
              {preview.count} sinais · prefixo <code className="font-mono text-brand-300">{preview.prefix}</code> · Remote Unit{" "}
              <code className="font-mono text-slate-400">{preview.remoteUnit}</code>
            </p>

            <div className="overflow-hidden rounded-xl border border-slate-800">
              <div className="max-h-[52vh] overflow-auto">
                <table className="w-full text-sm">
                  <thead className="sticky top-0 bg-slate-900 text-left text-xs uppercase tracking-wide text-slate-400">
                    <tr>
                      <th className="px-3 py-2">#</th>
                      <th className="px-3 py-2">Tipo</th>
                      <th className="px-3 py-2">Signal Name</th>
                      <th className="px-3 py-2">Descrição</th>
                      <th className="px-3 py-2">Custom ID</th>
                      <th className="px-3 py-2">Coord.</th>
                      <th className="px-3 py-2">Comando / Output</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-800/70">
                    {preview.rows.map((r, i) => (
                      <tr key={r.signalName} className="hover:bg-slate-800/30">
                        <td className="px-3 py-1.5 text-slate-500">{i + 1}</td>
                        <td className="px-3 py-1.5">
                          <span className={clsx(
                            "rounded-full px-2 py-0.5 text-[11px]",
                            r.klass === "analog" ? "bg-emerald-500/15 text-emerald-300" : "bg-amber-500/15 text-amber-300"
                          )}>
                            {r.klass === "analog" ? "ANL" : "DIG"}
                          </span>
                        </td>
                        <td className="px-3 py-1.5 font-mono text-brand-300">{r.signalName}</td>
                        <td className="px-3 py-1.5 text-slate-300">{r.description}</td>
                        <td className="px-3 py-1.5 font-mono text-slate-400">{r.customId}</td>
                        <td className="px-3 py-1.5 font-mono text-slate-400">{r.inputCoord ?? "—"}</td>
                        <td className="px-3 py-1.5">
                          {r.hasCommand ? (
                            <span className="flex items-center gap-1.5">
                              <span className="rounded bg-fuchsia-500/15 px-1.5 py-0.5 text-[10px] text-fuchsia-300">
                                {r.commandFormat}
                              </span>
                              <code className="font-mono text-slate-400">{r.outputCoord ?? "—"}</code>
                            </span>
                          ) : (
                            <span className="text-slate-700">—</span>
                          )}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>

            <div className="mt-6 flex items-center justify-between">
              <button onClick={() => goto(3)} className="btn-ghost">
                <ArrowLeft size={16} /> Voltar aos sinais
              </button>
              <button onClick={doExport} disabled={exporting} className="btn-primary">
                {exporting ? <Loader2 className="animate-spin" size={18} /> : <Download size={18} />}
                {exporting ? "Gerando…" : "Exportar TDT (.xlsx)"}
              </button>
            </div>
          </section>
        )}

        <footer className="mt-10 text-center text-xs text-slate-600">
          Formato e estrutura aprendidos das TDTs reais do ADMS · exportação fiel ao original.
        </footer>
      </div>
    </div>
  );
}

function Field({ label, hint, children }: { label: string; hint?: string; children: React.ReactNode }) {
  return (
    <label className="block">
      <span className="mb-1 block text-sm font-medium text-slate-300">{label}</span>
      {children}
      {hint && <span className="mt-1 block text-xs text-slate-500">{hint}</span>}
    </label>
  );
}

function NavButtons({ onBack, onNext, nextDisabled, nextLabel }: {
  onBack: () => void; onNext: () => void; nextDisabled?: boolean; nextLabel: string;
}) {
  return (
    <div className="mt-6 flex items-center justify-between">
      <button onClick={onBack} className="btn-ghost">
        <ArrowLeft size={16} /> Voltar
      </button>
      <button onClick={onNext} disabled={nextDisabled} className="btn-primary">
        {nextLabel} <ArrowRight size={16} />
      </button>
    </div>
  );
}
