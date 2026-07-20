"""
slim_tdt.py — Remove as abas de nível de UTR de uma TDT para NÃO tocar em
sinais de outros vãos ao importar.

Problema: o esqueleto v3 traz DNP3_RTUs (UTR_LVA_2) e DNP3_TCPLinks. Ao importar,
o ADMS re-processa a UTR — e todos os vãos pendurados nela (ex.: AL21/AL22 já
aplicados) aparecem como "mudados", mesmo sem nenhum sinal deles ser alterado.

Solução: importar SÓ as abas de sinais. A UTR já existe no modelo (foi criada por
uma importação anterior), então os sinais mapeiam nela normalmente — igual a TDT
de referência da URA, que não tem DNP3_RTUs e importa OK.

Mantém: DMSMatchingTemplateInfo (validação) + DNP3_DiscreteSignals + DNP3_AnalogSignals.
Remove: DNP3_RTUs, DNP3_TCPLinks/UDPLinks, DNP3_ScanGroups, Info, AlarmCatalog,
        MessageMappings, Manual*, DiscreteAnalog (vazia), Calculation*.

Uso: python slim_tdt.py arquivo1.xlsx [arquivo2.xlsx ...]
"""
from __future__ import annotations
import io, sys
from pathlib import Path
import openpyxl
import excel_native

DIR = Path("C:/Users/egnpo/Downloads")
KEEP = {"DMSMatchingTemplateInfo", "DNP3_DiscreteSignals", "DNP3_AnalogSignals"}


def slim(name: str):
    p = DIR / name
    wb = openpyxl.load_workbook(p)
    removed = [s for s in wb.sheetnames if s not in KEEP]
    for s in removed:
        del wb[s]
    buf = io.BytesIO()
    wb.save(buf)
    data = excel_native.resave_native(buf.getvalue())
    out = p.with_name(p.stem + "_SEM_UTR.xlsx")
    out.write_bytes(data)
    print(f"{out.name}: removidas {len(removed)} abas de config; mantidas {sorted(KEEP)}")


if __name__ == "__main__":
    files = sys.argv[1:] or [
        "TDT_LVA_COMPLETA_SEM_AL21_AL22.xlsx",
        "TDT_LVA_LIMPEZA.xlsx",
    ]
    for f in files:
        slim(f)
