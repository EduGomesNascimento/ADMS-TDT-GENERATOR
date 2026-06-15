"""
excel_native.py
===============
Pós-processador que converte um .xlsx gerado pelo openpyxl em um arquivo
**100% nativo do MS Excel**, eliminando as particularidades que fazem o ADMS
rejeitar a importação com:

    "Invalid TDI file format. Please use official Telemetry Data Template
     document and MS Excel or save file in MS Excel prior to attempting
     telemetry data import."

O openpyxl grava o pacote OOXML de um jeito que o parser TDI do ADMS recusa:
  - comentários em `xl/comments/comment1.xml` (Excel usa `xl/comments1.xml`)
  - sem `xl/sharedStrings.xml` (strings inline)
  - worksheets sem `codeName`, `xr:uid`, `mc:Ignorable`, declaração XML, etc.

A correção robusta — exatamente o que a mensagem do ADMS pede ("save file in
MS Excel prior to import") — é abrir o arquivo no Excel instalado (via COM) e
re-salvar. O Excel reescreve o pacote no formato canônico e, de quebra,
redimensiona as Tabelas (ListObjects) para cobrir todas as linhas de dados.

Requisitos: Windows com MS Excel instalado + pywin32.
Se o Excel não estiver disponível, `resave_native()` levanta `ExcelUnavailable`;
o chamador decide se faz fallback para os bytes do openpyxl.
"""
from __future__ import annotations

import os
import tempfile
import threading

# xlOpenXMLWorkbook (.xlsx sem macros)
_XL_XLSX = 51

# Serializa o acesso ao Excel (um app COM por vez) — app local mono-usuário.
_EXCEL_LOCK = threading.Lock()


class ExcelUnavailable(RuntimeError):
    """Excel/pywin32 não disponível para o pós-processamento nativo."""


def excel_available() -> bool:
    """True se pywin32 e o Excel estiverem instalados."""
    try:
        import win32com.client  # noqa: F401
    except ImportError:
        return False
    try:
        import winreg
        winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"Excel.Application\CurVer").Close()
        return True
    except OSError:
        return False


def _resave_worker(src_path: str, dst_path: str, resize_tables: bool, result: dict):
    """Roda numa thread própria com COM inicializado."""
    import pythoncom
    import win32com.client as win32

    pythoncom.CoInitialize()
    excel = None
    try:
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        # AutomationSecurity = 3 (msoAutomationSecurityForceDisable) — bloqueia macros
        try:
            excel.AutomationSecurity = 3
        except Exception:
            pass

        wb = excel.Workbooks.Open(src_path, UpdateLinks=0, ReadOnly=False)

        if resize_tables:
            _resize_all_tables(wb)

        wb.SaveAs(dst_path, FileFormat=_XL_XLSX)
        wb.Close(SaveChanges=False)
        result["ok"] = True
    except Exception as e:  # pragma: no cover - depende do ambiente
        result["error"] = f"{type(e).__name__}: {e}"
    finally:
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()


def _resize_all_tables(wb):
    """Redimensiona cada ListObject para cobrir exatamente sua região de dados.

    O openpyxl não consegue manipular ListObjects; o ref pode ficar curto e a
    formatação 'banded' (faixas) para no fim do template. Aqui usamos a API
    nativa do Excel para esticar a tabela até a última linha preenchida.
    """
    for sheet in wb.Sheets:
        try:
            list_objects = sheet.ListObjects
        except Exception:
            continue
        for tbl in list_objects:
            try:
                rng = tbl.Range
                first_row = rng.Row
                first_col = rng.Column
                n_cols = rng.Columns.Count
                # última linha com dados na 1ª coluna da tabela (a partir do header)
                col = first_col
                last_row = sheet.Cells(sheet.Rows.Count, col).End(-4162).Row  # xlUp
                if last_row < first_row + 1:
                    last_row = first_row + 1  # ao menos header + 1
                new_rng = sheet.Range(
                    sheet.Cells(first_row, first_col),
                    sheet.Cells(last_row, first_col + n_cols - 1),
                )
                tbl.Resize(new_rng)
            except Exception:
                continue


def resave_native(xlsx_bytes: bytes, resize_tables: bool = True) -> bytes:
    """
    Recebe os bytes de um .xlsx (gerado pelo openpyxl) e devolve os bytes do
    mesmo arquivo re-salvo pelo MS Excel (formato canônico, aceito pelo ADMS).

    Levanta `ExcelUnavailable` se o Excel/pywin32 não estiverem disponíveis.
    """
    if not excel_available():
        raise ExcelUnavailable("MS Excel ou pywin32 não disponível")

    tmpdir = tempfile.mkdtemp(prefix="tdt_native_")
    src_path = os.path.join(tmpdir, "in.xlsx")
    dst_path = os.path.join(tmpdir, "out.xlsx")
    with open(src_path, "wb") as f:
        f.write(xlsx_bytes)

    result: dict = {"ok": False}
    with _EXCEL_LOCK:
        th = threading.Thread(
            target=_resave_worker,
            args=(src_path, dst_path, resize_tables, result),
            daemon=True,
        )
        th.start()
        th.join(timeout=120)

    try:
        if not result.get("ok"):
            raise ExcelUnavailable(result.get("error") or "falha ao re-salvar no Excel")
        with open(dst_path, "rb") as f:
            return f.read()
    finally:
        for p in (src_path, dst_path):
            try:
                os.remove(p)
            except OSError:
                pass
        try:
            os.rmdir(tmpdir)
        except OSError:
            pass
