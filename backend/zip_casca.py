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

# Se a TDT/lista estava aberta no Excel na geracao, a versao atual e a _NOVA.
# Empacota sempre a MAIS RECENTE das duas, com o nome final.
def _atual(base: str):
    o = DOWN / f"{base}.xlsx"; n = DOWN / f"{base}_NOVA.xlsx"
    if n.exists() and (not o.exists() or n.stat().st_mtime > o.stat().st_mtime):
        return n, f"{base}.xlsx"          # (arquivo real, nome no zip)
    return o, None


# (caminho, pasta no zip, obrigatório?[, nome no zip])
ITENS = [
    # ── o que se entrega ──
    (*_atual("TDT_CASCA_UTR_CAS_3"),                          "1-ENTREGA", True),
    (DOWN / "TDT_CASCA_ATUAL.xlsx",                          "1-ENTREGA", False),
    (DOWN / "TDT_CASCA_FUTURA.xlsx",                         "1-ENTREGA", False),
    (DOWN / "TDT_CASCA_LT2.xlsx",                            "1-ENTREGA", False),
    (*_atual("RGE ADMS_Lista Pontos Casca_CORRIGIDA"),       "1-ENTREGA", False),
    (DOWN / "CASCA_RELATORIO.xlsx",                          "1-ENTREGA", True),
    (RAIZ / "CASCA_HANDOFF.md",                              ".",         True),
    # ── de onde saiu ──
    (DOWN / "RGE ADMS_Lista Pontos Casca.xlsx",  "2-ENTRADAS", True),
    (DOWN / "PT-MOD-SE-CASCA.xml",               "2-ENTRADAS", False),
    (DOWN / "PT-MOD-SE-CAS.xml",                 "2-ENTRADAS", False),
    (DOWN / "CASCA.xlsx",                        "2-ENTRADAS", True),
    (DOWN / "TDT_LVA_AL24.xlsx",                 "2-ENTRADAS", False),
    (DOWN / "erros.csv",                         "2-ENTRADAS", False),
    (DOWN / "Changesets.csv",                    "2-ENTRADAS", False),
    # ── o que gera ──
    (RAIZ / "backend/make_casca.py",             "3-CODIGO", True),
    (RAIZ / "backend/casca_devmap.py",           "3-CODIGO", True),
    (RAIZ / "backend/casca_status.py",           "3-CODIGO", False),
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
             - TDT_CASCA_ATUAL.xlsx ............ so os sinais cujo dispositivo
               JA EXISTE no Cas_Obra. Importa limpo. COMECE POR ELA.
             - TDT_CASCA_FUTURA.xlsx ........... os que esperam o dispositivo
               ser criado (ver aba 23 do relatorio)
             - TDT_CASCA_UTR_CAS_3.xlsx ........ as duas juntas (completa)
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
        654 com dispositivo que JA EXISTE  -> TDT_CASCA_ATUAL.xlsx
        628 esperando o dispositivo        -> TDT_CASCA_FUTURA.xlsx
        0 Device Mapping com tipo errado, 0 "Found multiple"
COMANDOS: 176/176 resolvidos

NAO EXISTE sinal sem dispositivo: apontar para a UTR (UTR_CAS_3) tambem e
recusado pelo ADMS. Por isso a TDT FUTURA carrega o nome do dispositivo que
PRECISA ser criado — o erro do import vira a lista do que fazer (aba 23).
"""


def main():
    OUT.unlink(missing_ok=True)
    faltando = []
    with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED, compresslevel=6) as z:
        z.writestr("LEIA-ME.txt", LEIAME)
        for item in ITENS:
            # itens de _atual() tem nome-no-zip extra: (src, zipnome, pasta, obrig)
            if len(item) == 4:
                src, zipnome, pasta, obrig = item
            else:
                src, pasta, obrig = item; zipnome = None
            if not src.exists():
                if obrig:
                    faltando.append(src.name)
                continue
            nome = zipnome or src.name
            destino = f"{pasta}/{nome}" if pasta != "." else nome
            z.write(src, destino)
            print(f"  + {destino}  ({src.stat().st_size:,} b)")
    print(f"\n{OUT.name}: {OUT.stat().st_size:,} bytes")
    if faltando:
        print(f"ATENCAO — obrigatorios ausentes: {faltando}")


if __name__ == "__main__":
    main()
