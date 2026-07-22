"""
zip_casca.py — Empacota TUDO da SE CASCA para a próxima pessoa continuar.

Junta num único .zip: os arquivos gerados (TDT, lista corrigida, relatório),
as entradas que os produziram (lista original, XML do modelo, TDT atual,
esqueleto, erros do ADMS), o código que gera e confere, e o handoff em MD.

Uso: python zip_casca.py
"""
from __future__ import annotations
import zipfile
from pathlib import Path

DOWN = Path("C:/Users/egnpo/Downloads")
RAIZ = Path(__file__).resolve().parent.parent
OUT = DOWN / "CASCA_UTR_CAS_3_PACOTE.zip"

# (caminho, pasta no zip, obrigatório?)
ITENS = [
    # ── o que se entrega ──
    (DOWN / "TDT_CASCA_UTR_CAS_3.xlsx",                      "1-ENTREGA", True),
    (DOWN / "RGE ADMS_Lista Pontos Casca_CORRIGIDA.xlsx",     "1-ENTREGA", False),
    (DOWN / "RGE ADMS_Lista Pontos Casca_CORRIGIDA_NOVA.xlsx", "1-ENTREGA", False),
    (DOWN / "CASCA_RELATORIO.xlsx",                          "1-ENTREGA", True),
    (RAIZ / "CASCA_HANDOFF.md",                              ".",         True),
    # ── de onde saiu ──
    (DOWN / "RGE ADMS_Lista Pontos Casca.xlsx",  "2-ENTRADAS", True),
    (DOWN / "PT-MOD-SE-CASCA.xml",               "2-ENTRADAS", True),
    (DOWN / "CASCA.xlsx",                        "2-ENTRADAS", True),
    (DOWN / "TDT_LVA_AL24.xlsx",                 "2-ENTRADAS", False),
    (DOWN / "erros.csv",                         "2-ENTRADAS", False),
    (DOWN / "Changesets.csv",                    "2-ENTRADAS", False),
    # ── o que gera ──
    (RAIZ / "backend/make_casca.py",             "3-CODIGO", True),
    (RAIZ / "backend/casca_devmap.py",           "3-CODIGO", True),
    (RAIZ / "backend/check_casca.py",            "3-CODIGO", True),
    (RAIZ / "backend/excel_native.py",           "3-CODIGO", True),
    (RAIZ / "backend/tdt_engine.py",             "3-CODIGO", True),
    (RAIZ / "backend/data/sigla_index.json",     "3-CODIGO/data", True),
    # ── contexto do projeto inteiro ──
    (RAIZ / "HANDOFF_COMPLETO.md",               "4-CONTEXTO", False),
    (RAIZ / "TECNICO_COMO_FOI_FEITO.md",         "4-CONTEXTO", False),
]

LEIAME = """SE CASCA — UTR_CAS_3 (DNP3)
============================

Comece por CASCA_HANDOFF.md — ele explica tudo: a regra do projeto, as decisoes
tomadas, as armadilhas ja pagas e o que fazer em seguida.

1-ENTREGA/   o que se importa no ADMS e o relatorio que explica
             - TDT_CASCA_UTR_CAS_3.xlsx ........ A TDT (1282 sinais)
             - RGE ADMS_..._CORRIGIDA.xlsx ..... lista com o INDEX DNP3 arrumado
             - CASCA_RELATORIO.xlsx ............ 16 abas; comece pela "0-LEIA-ME"

2-ENTRADAS/  os arquivos que produziram a entrega
             - RGE ADMS_Lista Pontos Casca.xlsx  fonte dos SINAIS
             - PT-MOD-SE-CASCA.xml ............. fonte dos DISPOSITIVOS
             - CASCA.xlsx ...................... TDT atual (convencao de Device Mapping)
             - TDT_LVA_AL24.xlsx ............... esqueleto de TDT valida
             - erros.csv ....................... retorno do validador do ADMS

3-CODIGO/    para regerar:
                 python make_casca.py     (gera os 3 arquivos)
                 python check_casca.py    (confere; tem que dar OK)
             Precisa de Windows com Excel instalado (pywin32) — a TDT so e
             aceita se for regravada pelo proprio Excel.

4-CONTEXTO/  documentacao do gerador de TDT como um todo


REGRA QUE REGE O MAPEAMENTO
---------------------------
A LISTA manda nos SINAIS (nome, tipo, index, comando, escala).
O UNIFILAR/ADMS manda nos DISPOSITIVOS (Device Mapping).

Onde mexer se uma equivalencia de vao estiver errada:
    3-CODIGO/casca_devmap.py  ->  tabela MODULO_EQUIV
Esta tudo num lugar so; corrigir ali e rodar de novo.


ESTADO NA HORA DO EMPACOTAMENTO
-------------------------------
LISTA:  2535 linhas de sinal, 0 sem index
        D 0..1904 · A 1..361 · C 1..300 — contiguos, sem duplicata
TDT:    1282 sinais, 0 nome duplicado, 0 coordenada duplicada
        858 mapeiam no unifilar (331 com dispositivo rebaixado)
        424 sem dispositivo -> ver aba 13 do relatorio
COMANDOS: 176/176 resolvidos
"""


def main():
    OUT.unlink(missing_ok=True)
    faltando = []
    with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED, compresslevel=6) as z:
        z.writestr("LEIA-ME.txt", LEIAME)
        for src, pasta, obrig in ITENS:
            if not src.exists():
                if obrig:
                    faltando.append(src.name)
                continue
            destino = f"{pasta}/{src.name}" if pasta != "." else src.name
            z.write(src, destino)
            print(f"  + {destino}  ({src.stat().st_size:,} b)")
    print(f"\n{OUT.name}: {OUT.stat().st_size:,} bytes")
    if faltando:
        print(f"ATENCAO — obrigatorios ausentes: {faltando}")


if __name__ == "__main__":
    main()
