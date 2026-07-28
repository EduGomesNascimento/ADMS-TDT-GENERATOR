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
    (DOWN / "Changesets.csv",                    "2-ENTRADAS", False),
    # retornos do ADMS, em ordem cronologica (o ultimo e o que vale)
    (DOWN / "erros.csv",                         "2-ENTRADAS/retorno-adms", False),
    (DOWN / "errpsmapping.csv",                  "2-ENTRADAS/retorno-adms", False),
    (DOWN / "eros3.csv",                         "2-ENTRADAS/retorno-adms", False),
    (DOWN / "erros6_UNIDO.csv",                  "2-ENTRADAS/retorno-adms", False),
    (DOWN / "ERROS10.csv",                       "2-ENTRADAS/retorno-adms", False),
    (DOWN / "ERROS 100.csv",                     "2-ENTRADAS/retorno-adms", False),
    (DOWN / "ERROS CS.csv",                      "2-ENTRADAS/retorno-adms", False),
    # ── o que gera ──
    (RAIZ / "backend/make_casca.py",             "3-CODIGO", True),
    (RAIZ / "backend/casca_devmap.py",           "3-CODIGO", True),
    (RAIZ / "backend/casca_status.py",           "3-CODIGO", False),
    (RAIZ / "backend/check_casca.py",            "3-CODIGO", True),
    (RAIZ / "backend/excel_native.py",           "3-CODIGO", True),
    (RAIZ / "backend/tdt_engine.py",             "3-CODIGO", True),
    (RAIZ / "backend/_extrai_base.py",           "3-CODIGO", False),
    (RAIZ / "backend/data/sigla_index.json",     "3-CODIGO/data", True),
    # convencao das 27 SEs: sufixo de dispositivo, quota e config de comando
    (RAIZ / "backend/data/convencao_dm.json",    "3-CODIGO/data", False),
    (RAIZ / "backend/data/casca_cmd_cfg.json",   "3-CODIGO/data", False),
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
             - TDT_CASCA_LT2.xlsx .............. so o vao LT PRI, p/ teste
             - RGE ADMS_..._CORRIGIDA.xlsx ..... lista com o INDEX DNP3 arrumado
             - CASCA_RELATORIO.xlsx ............ 24 abas; comece pela
               "0-O QUE FALTA" e depois a "20-Historico de erros"

2-ENTRADAS/  os arquivos que produziram a entrega
             - RGE ADMS_Lista Pontos Casca.xlsx  fonte dos SINAIS
             - PT-MOD-SE-CAS.xml ............... fonte dos DISPOSITIVOS
               (ja com os _NEW; o PT-MOD-SE-CASCA.xml e o anterior)
             - CASCA.xlsx ...................... TDT atual (convencao de
               Device Mapping e de configuracao de comando)
             - TDT_LVA_AL24.xlsx ............... esqueleto de TDT valida
             - retorno-adms/ ................... os CSV do validador, em ordem;
               o ultimo (ERROS CS.csv) e o que vale

3-CODIGO/    para regerar:
                 python make_casca.py     (gera os 3 arquivos)
                 python make_casca.py --rapido    (nao regrava a lista
                     corrigida — so precisa quando a lista de pontos muda;
                     e o passo mais lento, ~5300 formulas congeladas)
                 python check_casca.py    (confere; tem que dar OK)
             Precisa de Windows com Excel instalado (pywin32) — a TDT so e
             aceita se for regravada pelo proprio Excel.

             data/convencao_dm.json vem da base completa das 27 SEs
             (_extrai_base.py). Traz sufixo de dispositivo por sigla, quota
             por tipo e a configuracao de COMANDO por sigla.

4-CONTEXTO/  documentacao do gerador de TDT como um todo


REGRA QUE REGE O MAPEAMENTO
---------------------------
A LISTA manda nos SINAIS (nome, tipo, index, comando, escala).
O UNIFILAR/ADMS manda nos DISPOSITIVOS (Device Mapping).

Onde mexer se uma equivalencia de vao estiver errada:
    3-CODIGO/casca_devmap.py  ->  tabela MODULO_EQUIV
Esta tudo num lugar so; corrigir ali e rodar de novo.


ABA OCULTA NAO EXISTE
---------------------
A lista tem 36 abas, mas 20 estao OCULTAS no Excel — voce nao as ve ao abrir.
Elas somam 1455 sinais e sao, em sua maioria, COPIAS dos vaos vivos:
'LT 1 (FUTURO)' repete o 'LT 1', 'BC 1' repete a 'TRANSFERENCIA 24-01'. Outras
sao vaos que simplesmente nao existem no Cas_Obra (AL24, AL25, AL26, TRF29,
AL28, IB20).

Ate 28/07/2026 o gerador lia TODAS as abas, porque o openpyxl nao liga para o
estado da aba. Isso inflava a lista de 1078 para 2535 pontos, criava as 704
coordenadas repetidas e os 1145 #REF! que eram "corrigidos", e enchia a TDT
FUTURA de vaos fantasma. Agora read_lista() pula tudo que nao esteja
'visible' e imprime quais ignorou.

REGRA: o que esta oculto nao existe e nao deve ser considerado.


ESTADO NA HORA DO EMPACOTAMENTO
-------------------------------
LISTA:  1078 linhas de sinal (so as 16 abas VISIVEIS), 0 sem index
        1 unico #REF! remanescente (eram 1145 com as ocultas)
TDT:    888 sinais, 0 nome duplicado, 0 coordenada duplicada
        582 com dispositivo que JA EXISTE  -> TDT_CASCA_ATUAL.xlsx
        306 esperando o dispositivo        -> TDT_CASCA_FUTURA.xlsx
        0 Device Mapping com tipo errado, 0 "Found multiple"
COMANDOS: 176/176 resolvidos, 176/176 com configuracao de saida completa
LINK:    UTR_CAS_3_Link1 presente nas TRES TDTs (a UTR exige >=1)
check_casca.py: OK — nenhum problema encontrado

NAO EXISTE sinal sem dispositivo: apontar para a UTR (UTR_CAS_3) tambem e
recusado pelo ADMS. Por isso a TDT FUTURA carrega o nome do dispositivo que
PRECISA ser criado — o erro do import vira a lista do que fazer (aba 23).


UNICA PENDENCIA QUE NAO DEPENDE DO MODELO
-----------------------------------------
O IP do link esta com o PLACEHOLDER 0.0.0.0. A aba "Informacoes" da lista traz
o IP como "X" — ainda nao definido — e o ADMS nao aceita a celula vazia
("IP Address cannot be left empty"). O 0.0.0.0 nao e roteavel, entao nao ha
risco de conflito com equipamento real, mas TEM que ser trocado antes de a UTR
entrar em operacao. O time de comunicacao informa o definitivo; trocar em
LINK_IP no 3-CODIGO/make_casca.py e rodar de novo. (A LVA usa 10.7.124.99.)

O resto das pendencias e trabalho NO MODELO do ADMS, listado na aba
"0-O QUE FALTA" do relatorio.
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
