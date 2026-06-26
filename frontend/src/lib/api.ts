export interface SignalCount {
  discrete: number;
  analog: number;
}

export interface DeviceTypeSummary {
  id: string;
  label: string;
  description: string;
  source: string;
  consolidated?: boolean;
  paramKind?: "transformer" | "line" | "bus" | null;
  defaults: { module?: string; device?: string; transformerNumber?: string };
  signalCount: SignalCount;
}

export interface CommandInfo {
  outputDataType: string | null;
  controlCodes: string | null;
  commandTimes: string | null;
  commandingMode: string | null;
  commandTimeout: number | string | null;
  coordCount: number;
  templateCoord: number | string | null;
}

export interface Signal {
  suffix: string;
  code?: string | null;
  group?: string | null;
  description: string;
  aliasLabel: string;
  klass: "discrete" | "analog" | "discrete_analog";
  measurementType: string | null;
  signalSubType: string | null;
  phases: string | null;
  hasCommand: boolean;
  command: CommandInfo | null;
}

export interface DeviceDetail extends DeviceTypeSummary {
  signals: { discrete: Signal[]; analog: Signal[]; discrete_analog?: Signal[] };
}

export interface SignalSel {
  sheet: string;
  suffix: string;
  inputCoord?: number | string | null;
  followFrom?: boolean;       // "próximos sinais seguem a partir deste endereço"
  customId?: string | null;
  commandFormat?: string | null;
  outputCoord?: number | string | null;
}

export interface ExportConfig {
  deviceTypeId: string;
  alias: string;
  module?: string;
  device?: string;
  transformerNumber?: string;
  deviceName?: string;
  remoteUnit?: string;
  customIdStart?: number;
  coordStart?: { discrete?: number; analog?: number };
  commandCoordStart?: number | null;
  signals: SignalSel[];
}

export const COMMAND_FORMATS: { value: string; label: string }[] = [
  { value: "template", label: "Manter do template" },
  { value: "trip_close", label: "Trip / Close (par)" },
  { value: "close_close", label: "Close / Close (par)" },
  { value: "single_close", label: "Pulso único — Close" },
  { value: "single_trip", label: "Pulso único — Trip" },
  { value: "latch", label: "Latch On / Off" },
];

export interface PreviewRow {
  sheet: string;
  klass: string;
  suffix: string;
  signalName: string;
  description: string;
  remoteUnit: string;
  customId: string;
  inputCoord: number | string | null;
  hasCommand: boolean;
  commandFormat: string | null;
  outputCoord: number | string | null;
}

export interface PreviewResult {
  prefix: string;
  remoteUnit: string;
  count: number;
  rows: PreviewRow[];
}

const BASE = "/api";

export async function fetchCatalog(): Promise<{ deviceTypes: DeviceTypeSummary[] }> {
  const r = await fetch(`${BASE}/catalog`);
  if (!r.ok) throw new Error("Falha ao carregar catálogo");
  return r.json();
}

export async function fetchDevice(id: string): Promise<DeviceDetail> {
  const r = await fetch(`${BASE}/device/${id}`);
  if (!r.ok) throw new Error("Falha ao carregar equipamento");
  return r.json();
}

export async function fetchPreview(cfg: ExportConfig): Promise<PreviewResult> {
  const r = await fetch(`${BASE}/preview`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(cfg),
  });
  if (!r.ok) throw new Error("Falha ao gerar prévia");
  return r.json();
}

export interface InputError {
  aba: string; linha: number | null; sigla: string; modulo: string; tipo: string; motivo: string;
}

export interface ImportSummary {
  discrete: { total: number; matched: number; unmatched: { sigla: string; nome: string }[] };
  analog: { total: number; matched: number; unmatched: { sigla: string; nome: string }[] };
  aliases: string[];
  inputErrors: InputError[];
  duplicates: string[];
}

export async function importPreview(file: File, protocol = "dnp3"): Promise<ImportSummary> {
  const r = await fetch(`${BASE}/import/preview?protocol=${protocol}`, {
    method: "POST",
    headers: { "Content-Type": "application/octet-stream" },
    body: file,
  });
  if (!r.ok) throw new Error((await r.text()) || "Falha ao ler a lista");
  return r.json();
}

export async function importExport(file: File, protocol = "dnp3"): Promise<void> {
  const r = await fetch(`${BASE}/import?protocol=${protocol}`, {
    method: "POST",
    headers: { "Content-Type": "application/octet-stream" },
    body: file,
  });
  if (!r.ok) throw new Error((await r.text()) || "Falha na geração");
  const blob = await r.blob();
  const disp = r.headers.get("Content-Disposition") || "";
  const m = disp.match(/filename="?([^"]+)"?/);
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = m ? m[1] : "TDT_importada.xlsx";
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

export async function importReport(file: File, protocol = "dnp3"): Promise<void> {
  const r = await fetch(`${BASE}/import/report?protocol=${protocol}`, {
    method: "POST",
    headers: { "Content-Type": "application/octet-stream" },
    body: file,
  });
  if (!r.ok) throw new Error((await r.text()) || "Falha no relatório");
  const blob = await r.blob();
  const disp = r.headers.get("Content-Disposition") || "";
  const m = disp.match(/filename="?([^"]+)"?/);
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url; a.download = m ? m[1] : "Relatorio_TDT.xlsx";
  document.body.appendChild(a); a.click(); a.remove();
  URL.revokeObjectURL(url);
}

export async function exportTdt(cfg: ExportConfig): Promise<void> {
  const r = await fetch(`${BASE}/export`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(cfg),
  });
  if (!r.ok) {
    const txt = await r.text();
    throw new Error(`Falha na exportação: ${txt}`);
  }
  const blob = await r.blob();
  const disp = r.headers.get("Content-Disposition") || "";
  const m = disp.match(/filename="?([^"]+)"?/);
  const fname = m ? m[1] : "TDT.xlsx";
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = fname;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}
