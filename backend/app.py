"""
app.py — API FastAPI do Gerador de TDTs ADMS.

Endpoints:
  GET  /api/catalog            -> tipos de equipamento + sinais disponíveis (slim)
  GET  /api/device/{id}        -> detalhe de um tipo (sinais + defaults)
  POST /api/preview            -> prévia em JSON das linhas que serão geradas
  POST /api/export             -> arquivo .xlsx final (idêntico ao formato ADMS)
"""
from __future__ import annotations
import datetime as dt
from pathlib import Path

import os

from fastapi import FastAPI, HTTPException, Request, UploadFile, File, Form
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
from pydantic import BaseModel
import io

import tdt_engine
from tdt_engine import (Catalog, generate_tdt, _subst, COMMAND_PRESETS,
                        parse_points_list, generate_tdt_from_list, build_import_report)
import ai_mapper
import probability_report

app = FastAPI(title="ADMS TDT Generator", version="1.0")
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)


def slim_signal(s):
    return {
        "suffix": s["suffix"],
        "code": s.get("code"),
        "group": s.get("group"),
        "description": s["description"],
        "aliasLabel": s.get("aliasLabel", ""),
        "klass": s["klass"],
        "measurementType": s.get("measurementType"),
        "signalSubType": s.get("signalSubType"),
        "phases": s.get("phases"),
        "hasCommand": s.get("hasCommand", False),
        "command": s.get("command"),
    }


def slim_device(d):
    return {
        "id": d["id"],
        "label": d["label"],
        "description": d["description"],
        "source": d["source"],
        "consolidated": d.get("consolidated", False),
        "paramKind": d.get("paramKind"),
        "defaults": d.get("defaults", {}),
        "signalCount": d["signalCount"],
        "signals": {
            "discrete": [slim_signal(s) for s in d["signals"].get("discrete", [])],
            "analog": [slim_signal(s) for s in d["signals"].get("analog", [])],
            "discrete_analog": [slim_signal(s) for s in d["signals"].get("discrete_analog", [])],
        },
    }


@app.get("/api/catalog")
def get_catalog():
    cat = Catalog()
    return {
        "meta": cat.raw["meta"],
        "deviceTypes": [
            {
                "id": d["id"], "label": d["label"], "description": d["description"],
                "source": d["source"], "defaults": d.get("defaults", {}),
                "consolidated": d.get("consolidated", False),
                "paramKind": d.get("paramKind"),
                "signalCount": d["signalCount"],
            }
            for d in cat.raw["deviceTypes"]
        ],
    }


@app.get("/api/device/{device_id}")
def get_device(device_id: str):
    cat = Catalog()
    d = cat.device(device_id)
    if not d:
        raise HTTPException(404, "tipo de equipamento não encontrado")
    return slim_device(d)


class SignalSel(BaseModel):
    sheet: str
    suffix: str
    inputCoord: int | str | None = None
    customId: str | None = None
    commandFormat: str | None = None
    outputCoord: int | str | None = None


class ExportConfig(BaseModel):
    deviceTypeId: str
    alias: str
    module: str | None = None
    device: str | None = None
    transformerNumber: str | None = None
    deviceName: str | None = None
    remoteUnit: str | None = None
    customIdStart: int | None = 1
    coordStart: dict | None = None
    commandCoordStart: int | None = None
    signals: list[SignalSel]


@app.post("/api/preview")
def preview(cfg: ExportConfig):
    """Retorna as linhas resultantes (campos-chave) sem gerar o arquivo."""
    cat = Catalog()
    dev = cat.device(cfg.deviceTypeId)
    if not dev:
        raise HTTPException(404, "tipo de equipamento não encontrado")
    alias = cfg.alias.strip()
    module = (cfg.module or "").strip()
    device = (cfg.device or "").strip()
    prefix = f"{alias}_{module}_{device}"
    remote_unit = cfg.remoteUnit or f"UTR_{alias}_1"
    mapping = {"<<PREFIX>>": prefix, "<<ALIAS>>": alias, "<<MODULE>>": module,
               "<<DEVICE>>": device, "<<N>>": str(cfg.transformerNumber or "1")}

    sig_index = {}
    for klass, sheet in (("discrete", "DNP3_DiscreteSignals"), ("analog", "DNP3_AnalogSignals"),
                         ("discrete_analog", "DNP3_DiscreteAnalog")):
        for s in dev["signals"].get(klass, []):
            sig_index[(sheet, s["suffix"])] = s

    rows = []
    seq = {"DNP3_DiscreteSignals": int(cfg.customIdStart or 1),
           "DNP3_AnalogSignals": int(cfg.customIdStart or 1)}
    coord_seq = dict((cfg.coordStart or {}))
    cmd_seq = cfg.commandCoordStart
    for sel in cfg.signals:
        s = sig_index.get((sel.sheet, sel.suffix))
        if not s:
            continue
        klass = "analog" if "Analog" in sel.sheet else "discrete"
        tag = "AS" if klass == "analog" else "DS"
        name = _subst(s["row"][0], mapping) if s.get("row") else f"{prefix}_{sel.suffix}"
        custom = sel.customId or f"{alias}_{tag}_{seq[sel.sheet]:05d}"
        if not sel.customId:
            seq[sel.sheet] += 1
        coord = sel.inputCoord
        if (coord is None or coord == "") and klass in (cfg.coordStart or {}):
            coord = coord_seq.get(klass)
            coord_seq[klass] = (coord_seq.get(klass) or 0) + 1

        # comando (output)
        out_coord = None
        cmd_fmt = None
        cmd = s.get("command")
        if cmd:
            cmd_fmt = sel.commandFormat or "template"
            preset = COMMAND_PRESETS.get(cmd_fmt)
            count = preset["count"] if preset else (cmd.get("coordCount") or 1)
            if sel.outputCoord not in (None, ""):
                out_coord = sel.outputCoord
            elif cfg.commandCoordStart is not None:
                out_coord = f"{cmd_seq};{cmd_seq}" if count == 2 else cmd_seq
                cmd_seq += 1
            else:
                out_coord = cmd.get("templateCoord")

        rows.append({
            "sheet": sel.sheet,
            "klass": klass,
            "suffix": sel.suffix,
            "signalName": name,
            "description": s["description"],
            "remoteUnit": remote_unit,
            "customId": custom,
            "inputCoord": coord,
            "hasCommand": cmd is not None,
            "commandFormat": cmd_fmt,
            "outputCoord": out_coord,
        })
    return {"prefix": prefix, "remoteUnit": remote_unit, "count": len(rows), "rows": rows}


@app.post("/api/export")
def export(cfg: ExportConfig):
    try:
        data = generate_tdt(cfg.model_dump())
    except Exception as e:
        raise HTTPException(400, str(e))
    stamp = dt.datetime.now().strftime("%Y%m%d_%H%M")
    mid = cfg.module or (f"TR{cfg.transformerNumber}" if cfg.transformerNumber else "")
    fname = f"TDT_{cfg.alias}_{mid}_{cfg.device or ''}_{stamp}.xlsx".replace(" ", "_").replace("__", "_")
    return StreamingResponse(
        io.BytesIO(data),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{fname}"'},
    )


@app.post("/api/import/preview")
async def import_preview(request: Request, protocol: str = "dnp3"):
    """Recebe a lista de pontos (.xlsx cru) e retorna resumo de casamento."""
    data = await request.body()
    if not data:
        raise HTTPException(400, "arquivo vazio")
    try:
        parsed = parse_points_list(data)
    except Exception as e:
        raise HTTPException(400, f"falha ao ler a lista: {e}")
    _, report = generate_tdt_from_list(parsed, protocol, native=False)
    # amostra de nomes por alias
    all_items = parsed["discrete"] + parsed["analog"]
    aliases = sorted({(it["nome"].split("_")[0]) for it in all_items if it["nome"]})
    # nomes duplicados (gerariam linhas duplicadas no ADMS)
    from collections import Counter
    cnt = Counter(it["nome"] for it in all_items if it["nome"])
    duplicates = sorted([n for n, c in cnt.items() if c > 1])
    return {
        "discrete": {"total": len(parsed["discrete"]), **report["discrete"]},
        "analog": {"total": len(parsed["analog"]), **report["analog"]},
        "aliases": aliases,
        "inputErrors": parsed.get("inputErrors", []),
        "duplicates": duplicates,
    }


@app.post("/api/import")
async def import_export(request: Request, protocol: str = "dnp3"):
    """Recebe a lista de pontos (.xlsx cru) e devolve a TDT gerada."""
    data = await request.body()
    if not data:
        raise HTTPException(400, "arquivo vazio")
    try:
        parsed = parse_points_list(data)
        xlsx, report = generate_tdt_from_list(parsed, protocol)
    except Exception as e:
        raise HTTPException(400, f"falha na geração: {e}")
    stamp = dt.datetime.now().strftime("%Y%m%d_%H%M")
    return StreamingResponse(
        io.BytesIO(xlsx),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="TDT_{protocol}_importada_{stamp}.xlsx"'},
    )


@app.post("/api/import/report")
async def import_report(request: Request, protocol: str = "dnp3"):
    """Gera um relatório .xlsx (Resumo, Revisão da Lista, Problemas)."""
    data = await request.body()
    if not data:
        raise HTTPException(400, "arquivo vazio")
    try:
        parsed = parse_points_list(data)
        xlsx = build_import_report(parsed, protocol)
    except Exception as e:
        raise HTTPException(400, f"falha no relatório: {e}")
    stamp = dt.datetime.now().strftime("%Y%m%d_%H%M")
    return StreamingResponse(
        io.BytesIO(xlsx),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="Relatorio_TDT_{protocol}_{stamp}.xlsx"'},
    )



# ─── Raw import (lista não-padrão com IA) ────────────────────────────────────

def _key_from_file(provider: str) -> str:
    """Lê a chave de um arquivo local (chave_groq.txt / chave_gemini.txt) na raiz
    do projeto, se existir — para o usuário só colar a chave num .txt."""
    fname = {'groq': 'chave_groq.txt', 'gemini': 'chave_gemini.txt'}.get(provider)
    if not fname:
        return ''
    for base in (Path(__file__).resolve().parent.parent, Path(__file__).resolve().parent):
        p = base / fname
        if p.exists():
            for line in p.read_text(encoding='utf-8', errors='ignore').splitlines():
                line = line.strip()
                if line and not line.startswith('#'):
                    return line
    return ''


def _llm_cfg(provider: str, model: str, api_key: str, base_url: str = "") -> dict | None:
    """Monta config LLM; retorna None se provider == 'none'."""
    if provider in ('none', ''):
        return None
    env_map = {'gemini': 'GEMINI_API_KEY', 'groq': 'GROQ_API_KEY'}
    key = api_key or os.environ.get(env_map.get(provider, ''), '') or _key_from_file(provider)
    if not key and provider != 'ollama':
        return None
    cfg = {'provider': provider, 'model': model, 'api_key': key}
    if provider == 'ollama':
        cfg['base_url'] = base_url or os.environ.get('OLLAMA_HOST', 'http://localhost:11434')
    return cfg


@app.post("/api/raw/preview")
async def raw_preview(
    file: UploadFile = File(...),
    model: str = Form("llama-3.3-70b-versatile"),
    provider: str = Form("groq"),
    api_key: str = Form(""),
    alias: str = Form(""),
    protocol: str = Form("dnp3"),
    base_url: str = Form(""),
):
    """Parseia lista não-padrão e mapeia sinais com IA. Retorna JSON."""
    data = await file.read()
    if not data:
        raise HTTPException(400, "arquivo vazio")
    try:
        raw_signals, detected_alias = ai_mapper.parse_raw_excel(data)
    except Exception as e:
        raise HTTPException(400, f"falha ao ler o arquivo: {e}")

    eff_alias = alias.strip() or detected_alias
    cfg = _llm_cfg(provider, model, api_key, base_url)

    try:
        mapped = ai_mapper.map_signals(raw_signals, protocol=protocol, llm_cfg=cfg)
    except Exception as e:
        raise HTTPException(500, f"falha no mapeamento: {e}")

    def _stats(lst):
        return {
            'total': len(lst),
            'alta':  sum(1 for m in lst if m.confidence_label == 'ALTA'),
            'media': sum(1 for m in lst if m.confidence_label == 'MÉDIA'),
            'baixa': sum(1 for m in lst if m.confidence_label == 'BAIXA'),
            'sem':   sum(1 for m in lst if m.confidence_label == 'SEM'),
        }

    discrete = [m for m in mapped if m.signal_type != 'analog']
    analog   = [m for m in mapped if m.signal_type == 'analog']

    return {
        'detectedAlias': detected_alias,
        'alias': eff_alias,
        'discrete': _stats(discrete),
        'analog': _stats(analog),
        'signals': [
            {
                'utrId': m.utr_id,
                'description': m.description,
                'dnp3Addr': m.dnp3_addr,
                'signalType': m.signal_type,
                'module': m.module,
                'sigla': m.sigla,
                'siglaDesc': m.sigla_desc,
                'confidence': m.confidence,
                'confidenceLabel': m.confidence_label,
                'matchMethod': m.match_method,
            }
            for m in mapped
        ],
    }


@app.post("/api/raw/report")
async def raw_report(
    file: UploadFile = File(...),
    model: str = Form("llama-3.3-70b-versatile"),
    provider: str = Form("groq"),
    api_key: str = Form(""),
    alias: str = Form(""),
    protocol: str = Form("dnp3"),
    base_url: str = Form(""),
):
    """Gera o arquivo de probabilidades .xlsx."""
    data = await file.read()
    if not data:
        raise HTTPException(400, "arquivo vazio")
    raw_signals, detected_alias = ai_mapper.parse_raw_excel(data)
    eff_alias = alias.strip() or detected_alias
    cfg = _llm_cfg(provider, model, api_key, base_url)
    mapped = ai_mapper.map_signals(raw_signals, protocol=protocol, llm_cfg=cfg)
    xlsx = probability_report.build_probability_xlsx(
        mapped, alias=eff_alias, source_file=file.filename or '')
    stamp = dt.datetime.now().strftime("%Y%m%d_%H%M")
    fname = f"Probabilidades_{eff_alias}_{stamp}.xlsx"
    return StreamingResponse(
        io.BytesIO(xlsx),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{fname}"'},
    )


@app.post("/api/raw/export")
async def raw_export(
    file: UploadFile = File(...),
    model: str = Form("llama-3.3-70b-versatile"),
    provider: str = Form("groq"),
    api_key: str = Form(""),
    alias: str = Form(""),
    min_confidence: int = Form(60),
    protocol: str = Form("dnp3"),
    base_url: str = Form(""),
):
    """Gera a TDT .xlsx a partir da lista não-padrão mapeada."""
    data = await file.read()
    if not data:
        raise HTTPException(400, "arquivo vazio")
    raw_signals, detected_alias = ai_mapper.parse_raw_excel(data)
    eff_alias = alias.strip() or detected_alias
    if not eff_alias:
        raise HTTPException(400, "alias da subestação não detectado — informe manualmente")
    cfg = _llm_cfg(provider, model, api_key, base_url)
    mapped = ai_mapper.map_signals(raw_signals, protocol=protocol, llm_cfg=cfg)
    lista = ai_mapper.to_lista_resumida(mapped, eff_alias, min_confidence=min_confidence)
    try:
        xlsx, _ = generate_tdt_from_list(lista, protocol)
    except Exception as e:
        raise HTTPException(500, f"falha na geração da TDT: {e}")
    stamp = dt.datetime.now().strftime("%Y%m%d_%H%M")
    fname = f"TDT_{eff_alias}_{protocol}_raw_{stamp}.xlsx"
    return StreamingResponse(
        io.BytesIO(xlsx),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{fname}"'},
    )


@app.get("/api/health")
def health():
    try:
        import excel_native
        excel_ok = excel_native.excel_available()
    except Exception:
        excel_ok = False
    return {
        "status": "ok",
        "excelNative": excel_ok,
        "note": None if excel_ok else
            "MS Excel/pywin32 ausente — a TDT pode ser recusada pelo ADMS "
            "('Invalid TDI file format'). Instale o MS Excel nesta máquina.",
    }
