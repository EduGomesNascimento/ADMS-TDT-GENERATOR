"""
make_casca.py — TDT DNP3 da SE CASCA (UTR nova UTR_CAS_3, do zero).

Fonte da verdade: "RGE ADMS_Lista Pontos Casca.xlsx". Cada aba de módulo traz
Utilizado?/TIPO/SIGLA/NOME/Escala/Níveis Lógicos/Control Code/INDEX DNP3.
Um sinal comandável aparece em 2 linhas com o MESMO NOME: a `D` dá o índice de
ENTRADA, a `C` o de COMANDO. `A` = analógico, `A/D` = TAP (DiscreteAnalog).

As UTRs atuais (IEC104 UTR_CAS_1 / religadores RAD_*) são só referência — a nova
não interage com elas.

REALOCAÇÃO: a lista tem índices duplicados (duas cadeias de alocação que se
sobrepõem + duplicatas dentro do módulo). Mantemos o 1º ocupante de cada
coordenada e realocamos os demais para a 1ª coordenada livre acima do topo em
uso. Toda realocação é registrada no relatório.

Saídas:
  TDT_CASCA_UTR_CAS_3.xlsx       — a TDT
  CASCA_RELATORIO.xlsx           — inconsistências + de-para das realocações

Uso: python make_casca.py
"""
from __future__ import annotations
import io, json, re
from collections import defaultdict, OrderedDict
from copy import copy
from pathlib import Path
import openpyxl
from openpyxl.styles import Font, PatternFill
import excel_native
import tdt_engine as E

LISTA = Path("C:/Users/egnpo/Downloads/RGE ADMS_Lista Pontos Casca.xlsx")
SKEL = Path("C:/Users/egnpo/Downloads/TDT_LVA_AL24.xlsx")   # esqueleto c/ DNP3_RTUs
OUT_TDT = Path("C:/Users/egnpo/Downloads/TDT_CASCA_UTR_CAS_3.xlsx")
OUT_REL = Path("C:/Users/egnpo/Downloads/CASCA_RELATORIO.xlsx")
DATA = Path(__file__).parent / "data"

# DE-PARA das siglas da lista sem template na base — usa a linha-molde da
# variante de PERFIL TÉCNICO equivalente (o NOME e a SIGLA da lista são
# preservados, e o Signal Alias vem da DESCRIÇÃO da lista; do molde só vem a
# configuração: Signal Type / Input Data Type / Message Mapping).
#   FGOO 'FALHA GOOSE'                   -> FCOM 'FALHA COMUNICACAO'  (Custom, NORMAL@FALHA)
#   TOC  '26 - ALARME TEMP OLEO CDC'     -> TOA  '26 ALARME'          (Custom, NORMAL@ATUADO)
#   81A  '81 - FUNCAO RECOMPOSICAO ERAC' -> 81   '81 FUNCAO'          (Enabled, comandável)
#   81P/81F/81C (estados do ERAC)        -> FALH 'FALHA GERAL'        (Custom status)
# 81E1 NÃO serve p/ 81P/81F/81C: é RelayTrip, e RelayTrip apontando p/ device
# físico é rejeitado pelo validador do ADMS (visto na SND).
FALLBACK_SIGLA = {
    "FGOO": "FCOM", "TOC": "TOA",
    "81A": "81", "81P": "FALH", "81F": "FALH", "81C": "FALH",
}

RU = "UTR_CAS_3"
AOR = "CAS Trans"
FABRICANTE = "ELIPSE"
CONTAINER = "CASCA"
HEADER_ROWS = 4
SKIP_SHEETS = {"Informações", "RELACAO RELES", "MAPA DE REDE", "Lista"}


# ─── leitura da lista ────────────────────────────────────────────────────────
def read_lista():
    wb = openpyxl.load_workbook(LISTA, read_only=True, data_only=True)
    pts = []
    for sn in wb.sheetnames:
        if sn in SKIP_SHEETS:
            continue
        rows = list(wb[sn].iter_rows(values_only=True))
        hi = next((i for i, r in enumerate(rows[:14])
                   if r and any(str(c or "").strip() == "SIGLA SINAL" for c in r)), None)
        if hi is None:
            continue
        hdr = [str(c or "").strip() for c in rows[hi]]
        ix = {n: i for i, n in enumerate(hdr) if n}
        def g(r, k):
            i = ix.get(k)
            return r[i] if i is not None and i < len(r) else None
        for ri, r in enumerate(rows[hi + 1:], hi + 2):
            if not r or str(g(r, "Utilizado?") or "").strip().upper() != "SIM":
                continue
            nome = str(g(r, "NOME") or "").strip()
            if not nome:
                continue
            pts.append({
                "sheet": sn, "linha": ri,
                "tipo": str(g(r, "TIPO") or "").strip(),
                "sigla": str(g(r, "SIGLA SINAL") or "").strip(),
                "nome": nome,
                "idx": str(g(r, "INDEX DNP3") or "").strip(),
                "tipoPt": str(g(r, "Tipo") or "").strip(),
                "n0": str(g(r, "Nível Lógico 0") or "").strip(),
                "n1": str(g(r, "Nível Lógico 1") or "").strip(),
                "escala": g(r, "Escala"),
                "cc": str(g(r, "Control Code / Qualificador") or "").strip(),
                "desc": str(g(r, "DESCRIÇÃO DO PONTO") or "").strip(),
            })
    wb.close()
    return pts


def _nums(idx: str):
    """Coordenadas de um INDEX DNP3 ('97', '97;98', e floats do Excel '97.0')."""
    out = []
    for x in str(idx).split(";"):
        x = x.strip()
        if not x:
            continue
        try:
            out.append(int(float(x)))
        except (ValueError, TypeError):
            pass
    return out


# ─── diagnóstico das duplicatas da lista ─────────────────────────────────────
def diagnosticar(pts):
    """Evidencia as coordenadas repetidas na lista ORIGINAL (célula a célula)."""
    ocup = defaultdict(list)
    for p in pts:
        for n in _nums(p["idx"]):
            ocup[(p["tipo"], n)].append(p)
    dups = []
    for (g, n), lst in sorted(ocup.items(), key=lambda x: (x[0][0], x[0][1])):
        if len(lst) > 1:
            abas = {p["sheet"] for p in lst}
            for p in lst:
                dups.append({**p, "grupo": g, "coord": n,
                             "escopo": "ENTRE MODULOS" if len(abas) > 1 else "DENTRO DO MODULO",
                             "conflita_com": ", ".join(
                                 f'{q["sheet"]}!N{q["linha"]} {q["nome"]}'
                                 for q in lst if q is not p)})
    semidx = [{**p, "grupo": p["tipo"]} for p in pts if not _nums(p["idx"])]
    return dups, semidx


# ─── re-sequenciamento contínuo das coordenadas ──────────────────────────────
def sequenciar(pts):
    """Re-sequencia TODAS as coordenadas em ordem, sem buracos nem repetição.

    Ordena pelo índice ORIGINAL (preserva o arranjo pretendido) e distribui
    sequencialmente: onde havia 490,490,491 passa a 490,491,492. Cada módulo
    continua de onde o anterior parou. Pontos de posição (Multi Coord) consomem
    2 coordenadas; comando usa a mesma repetida (n;n). Pontos com INDEX inválido
    (#REF!) entram no fim da fila do seu grupo.
    """
    final, mapa = {}, []
    for g in ("D", "A", "C", "A/D"):
        lst = [p for p in pts if p["tipo"] == g]
        if not lst:
            continue
        # ordena pelo índice original; sem índice (#REF!) vai pro fim
        def chave(p):
            ns = _nums(p["idx"])
            return (0, ns[0], p["linha"]) if ns else (1, 0, p["linha"])
        lst = sorted(lst, key=chave)
        prox = min((n for p in lst for n in _nums(p["idx"])), default=0)
        for p in lst:
            multi = str(p.get("tipoPt", "")).upper().startswith("MULTI")
            if g == "C":
                val = f"{prox};{prox}"; prox += 1
            elif multi:
                val = f"{prox};{prox + 1}"; prox += 2
            else:
                val = str(prox); prox += 1
            final[(p["sheet"], p["linha"])] = val
            mapa.append({**p, "grupo": g, "de": p["idx"], "para": val,
                         "mudou": "SIM" if str(p["idx"]).strip() != val else "nao"})
    return final, mapa


# ─── relatório ───────────────────────────────────────────────────────────────
def gerar_relatorio(pts, mapa, dups, semidx, sem_tpl, nomes_dup, descartados=(), fallback_rows=()):
    wb = openpyxl.Workbook(); wb.remove(wb.active)
    bold = Font(bold=True); hdrfill = PatternFill("solid", fgColor="DDEBF7")
    warn = PatternFill("solid", fgColor="FFF2CC")

    def sheet(title, cols, rows, fills=None):
        ws = wb.create_sheet(title[:31])
        ws.append(cols)
        for c in range(1, len(cols) + 1):
            ws.cell(1, c).font = bold; ws.cell(1, c).fill = hdrfill
        for r in rows:
            ws.append(r)
        for i, w in enumerate([14, 34, 10, 12, 14, 14, 40], 1):
            if i <= len(cols):
                ws.column_dimensions[chr(64 + i)].width = w
        return ws

    mudou = [m for m in mapa if m["mudou"] == "SIM"]
    sheet("0-LEIA-ME",
          ["O PROBLEMA E A SOLUCAO"],
          [["PROBLEMA ENCONTRADO NA LISTA DE PONTOS"],
           ["1) Coordenadas REPETIDAS: o mesmo INDEX DNP3 aparece em 2+ sinais."],
           ["   - ENTRE MODULOS: duas cadeias de alocacao se sobrepoem."],
           ["     ex.: AL 13!N66 (CAS_AL13_52-13_SGFT) e INTERBARRAS BT!N41"],
           ["          (CAS_IB20_52-20_FA) usam os DOIS a coordenada 556."],
           ["   - DENTRO DO MODULO: no TR 2, a partir da linha 79, o indice sobe"],
           ["     de 2 em 2 (L79 MOLA=416 e L80 SF6A=416) — formula arrastada."],
           ["2) Coordenadas INVALIDAS: 172 pontos com #REF! (formula quebrada)."],
           [""],
           ["POR QUE ISSO QUEBRA: no DNP3 cada ponto de um grupo (entrada binaria,"],
           ["entrada analogica, saida) precisa de um indice UNICO na UTR. Coordenada"],
           ["repetida = o ADMS rejeita ('does not have unique input coordinates')."],
           [""],
           ["COMO FOI CORRIGIDO"],
           ["Re-sequenciamento CONTINUO: as coordenadas foram redistribuidas em"],
           ["ordem, sem buracos e sem repeticao, preservando o arranjo original"],
           ["(ordenado pelo indice antigo). Onde havia 490,490,491 passou a ser"],
           ["490,491,492. Cada modulo continua de onde o anterior parou."],
           ["Ponto de posicao (Multi Coord) consome 2 coordenadas; comando usa a"],
           ["mesma repetida (n;n). Os #REF! entraram no fim da fila do seu grupo."],
           [""],
           ["O QUE VOCE RECEBE"],
           ["- TDT_CASCA_UTR_CAS_3.xlsx ....... a TDT ja com as coordenadas novas"],
           ["- RGE ADMS_Lista Pontos Casca_CORRIGIDA.xlsx ... a lista com a coluna"],
           ["  INDEX DNP3 ja arrumada (mesma estrutura da original)"],
           ["- Aba 2 deste relatorio .......... de-para completo (antes -> depois)"],
           ["- Aba 3 .......................... evidencia das coordenadas repetidas"],
           [""],
           ["ATENCAO: as coordenadas tem que bater com o que for configurado na UTR"],
           ["ELIPSE. Use a lista CORRIGIDA como referencia para parametrizar a UTR."]])

    sheet("1-Resumo",
          ["Item", "Qtde", "Observação"],
          [["Pontos utilizados (SIM)", len(pts), "da lista de pontos"],
           ["Coordenadas RE-SEQUENCIADAS", len(mapa), "todas — ver aba 2 (de-para)"],
           ["  .. que MUDARAM de valor", len(mudou), "as demais ficaram iguais"],
           ["Coordenadas REPETIDAS na lista", len(dups), "ver aba 3 — evidencia"],
           ["Pontos com INDEX invalido (#REF!)", len(semidx), "ver aba 4"],
           ["Siglas SEM template na base", len(sem_tpl), "0 = todas resolvidas"],
           ["NOMES duplicados (descartados)", len(descartados), "ver aba 6"],
           ["UTR", RU, f"nova, DNP3, {FABRICANTE}, AOR {AOR}"]])

    sheet("2-DePara coordenadas",
          ["Aba", "Linha", "Tipo", "SIGLA", "NOME", "Index ORIGINAL", "Index NOVO", "Mudou?"],
          [[m["sheet"], m["linha"], m["grupo"], m["sigla"], m["nome"], m["de"], m["para"],
            m["mudou"]] for m in mapa])

    sheet("3-Coords REPETIDAS (evidencia)",
          ["Escopo", "Tipo", "Coordenada", "Aba", "Linha", "NOME", "Conflita com"],
          [[d["escopo"], d["grupo"], d["coord"], d["sheet"], d["linha"], d["nome"],
            d["conflita_com"]] for d in dups])

    sheet("4-INDEX invalido (#REF!)",
          ["Aba", "Linha", "Tipo", "SIGLA", "NOME", "Valor na lista"],
          [[x["sheet"], x["linha"], x["grupo"], x["sigla"], x["nome"], x["idx"]]
           for x in semidx])

    sheet("5-Siglas sem template",
          ["SIGLA", "Qtde pontos", "Abas", "NOMEs (exemplos)"],
          [[s, d["n"], ", ".join(sorted(d["sheets"])[:5]), ", ".join(d["nomes"][:3])]
           for s, d in sorted(sem_tpl.items(), key=lambda x: -x[1]["n"])])

    sheet("7-Nomes duplicados",
          ["NOME", "Tipo", "Ocorrencias", "Abas"],
          nomes_dup)

    sheet("6-Descartados (nome dup)",
          ["Aba", "Linha", "Tipo", "SIGLA", "NOME", "Index na lista"],
          [[d["sheet"], d["linha"], d["tipo"], d["sigla"], d["nome"], d["idx"]]
           for d in descartados])

    sheet("8-Siglas por equivalencia",
          ["SIGLA da lista", "Molde usado", "Qtde", "Descricao na lista"],
          fallback_rows)

    buf = io.BytesIO(); wb.save(buf)
    OUT_REL.write_bytes(buf.getvalue())


# ─── lista de pontos corrigida ───────────────────────────────────────────────
def gerar_lista_corrigida(mapa):
    """Copia a lista original e reescreve a coluna INDEX DNP3 com as coords novas.
    Casa por POSIÇÃO (aba, linha) — a coluna NOME é fórmula, então casar por nome
    falharia ao abrir o arquivo preservando fórmulas."""
    porpos = {(m["sheet"], m["linha"]): m["para"] for m in mapa}
    wb = openpyxl.load_workbook(LISTA)
    n = 0
    for sn in wb.sheetnames:
        if sn in SKIP_SHEETS:
            continue
        ws = wb[sn]
        hi = None
        for r in range(1, min(15, ws.max_row + 1)):
            if any(str(ws.cell(r, c).value or "").strip() == "SIGLA SINAL"
                   for c in range(1, ws.max_column + 1)):
                hi = r; break
        if hi is None:
            continue
        col = {str(ws.cell(hi, c).value or "").strip(): c
               for c in range(1, ws.max_column + 1)}
        cI = col.get("INDEX DNP3")
        if not cI:
            continue
        for r in range(hi + 1, ws.max_row + 1):
            v = porpos.get((sn, r))
            if v is not None:
                ws.cell(r, cI).value = v
                n += 1
    out = LISTA.with_name(LISTA.stem + "_CORRIGIDA.xlsx")
    wb.save(out)
    print(f"lista corrigida: {out.name} ({n} coordenadas escritas)")


# ─── TDT ─────────────────────────────────────────────────────────────────────
def main():
    pts = read_lista()
    print(f"lista: {len(pts)} pontos utilizados")

    idx = json.loads((DATA / "sigla_index.json").read_text(encoding="utf-8"))
    TPL = {"D": idx["DNP3_DiscreteSignals"], "A": idx["DNP3_AnalogSignals"],
           "A/D": idx.get("DNP3_DiscreteAnalog", {}), "C": idx["DNP3_DiscreteSignals"]}

    # NOMES duplicados: o ADMS exige nome único, então a TDT fica com o 1º.
    # A LISTA recebe TODOS (todo ponto SIM precisa de coordenada) — por isso o
    # descarte marca apenas quais linhas a TDT deve pular.
    vistos = set(); descartados = []; pular_tdt = set()
    for p in pts:
        k = (p["tipo"], p["nome"])
        if p["tipo"] in ("D", "A", "A/D") and k in vistos:
            descartados.append(p); pular_tdt.add((p["sheet"], p["linha"])); continue
        vistos.add(k)
    if descartados:
        print(f"nome duplicado (fora da TDT, mas COM coordenada na lista): {len(descartados)}")

    dups, semidx = diagnosticar(pts)
    print(f"diagnostico: {len(dups)} ocorrencias de coord repetida, {len(semidx)} com #REF!")
    final, mapa = sequenciar(pts)
    print(f"re-sequenciadas: {len(mapa)} coords ({sum(1 for m in mapa if m['mudou']=='SIM')} mudaram)")

    # lista de pontos CORRIGIDA (mesma estrutura, INDEX DNP3 arrumado)
    gerar_lista_corrigida(mapa)

    # siglas sem template + nomes duplicados
    sem_tpl = defaultdict(lambda: {"n": 0, "sheets": set(), "nomes": []})
    for p in pts:
        if p["tipo"] == "C":
            continue
        if p["sigla"] not in TPL.get(p["tipo"], {}) and FALLBACK_SIGLA.get(p["sigla"]) not in TPL.get(p["tipo"], {}):
            d = sem_tpl[p["sigla"]]; d["n"] += 1; d["sheets"].add(p["sheet"])
            if len(d["nomes"]) < 3: d["nomes"].append(p["nome"])
    cnt = defaultdict(list)
    for p in pts:
        cnt[(p["nome"], p["tipo"])].append(p["sheet"])
    nomes_dup = [[n, t, len(s), ", ".join(sorted(set(s)))] for (n, t), s in cnt.items() if len(s) > 1]

    _rel = lambda fb: gerar_relatorio(pts, mapa, dups, semidx, sem_tpl, nomes_dup, descartados, fb)

    # comandos por NOME
    cmd = {}
    for p in pts:
        if p["tipo"] == "C" and p["nome"] not in cmd:
            cmd[p["nome"]] = final[(p["sheet"], p["linha"])]

    # monta as linhas da TDT
    wb = openpyxl.load_workbook(SKEL)
    plano = [("DNP3_DiscreteSignals", "D"), ("DNP3_AnalogSignals", "A"),
             ("DNP3_DiscreteAnalog", "A/D")]
    gerados = defaultdict(int); pulados = 0; usou_fallback = []
    for sheet, tipo in plano:
        ws = wb[sheet]
        lab = {ws.cell(HEADER_ROWS, c).value: c for c in range(1, ws.max_column + 1)
               if ws.cell(HEADER_ROWS, c).value}
        ncol = ws.max_column
        styles = [copy(ws.cell(HEADER_ROWS + 1, c)._style) for c in range(1, ncol + 1)]
        if ws.max_row > HEADER_ROWS:
            ws.delete_rows(HEADER_ROWS + 1, ws.max_row - HEADER_ROWS)
        L = lambda n: lab.get(n)
        rows = []
        for p in pts:
            if p["tipo"] != tipo or (p["sheet"], p["linha"]) in pular_tdt:
                continue
            tpl = TPL[tipo].get(p["sigla"])
            if not tpl:
                alt = FALLBACK_SIGLA.get(p["sigla"])
                tpl = TPL[tipo].get(alt) if alt else None
                if tpl:
                    usou_fallback.append({**p, "molde": alt})
            if not tpl:
                pulados += 1
                continue
            parts = p["nome"].split("_")
            alias = parts[0]; mod = parts[1] if len(parts) > 1 else ""
            dev = parts[2] if len(parts) > 2 else ""
            mapping = {"<<PREFIX>>": f"{alias}_{mod}_{dev}", "<<ALIAS>>": alias,
                       "<<MODULE>>": mod, "<<DEVICE>>": dev, "<<N>>": "1"}
            row = [E._subst(v, mapping) for v in tpl]
            while len(row) < ncol:
                row.append(None)
            row = row[:ncol]
            def put(name, val):
                c = L(name)
                if c: row[c - 1] = val
            put("Signal Name", p["nome"]); put("Remote Point Name", p["nome"])
            # descrição AUTORITATIVA da lista (o molde pode ser de sigla equivalente)
            if p.get("desc"):
                put("Signal Alias", p["desc"])
            put("Signal Custom ID", None)
            put("Remote Point Custom ID", f"{p['nome']}_{RU}")
            put("Remote Unit", RU); put("Signal AOR Group", AOR)
            put("Input Coordinates", final[(p["sheet"], p["linha"])])
            if p["escala"] not in (None, "", "-") and tipo in ("A", "A/D"):
                put("Scaling Factor", p["escala"])
            if tipo == "D":
                out = cmd.get(p["nome"])
                if out:
                    put("Output Coordinates", out)
                else:  # sem comando: vira status puro (senao o validador exige Output)
                    for k in ("Output Coordinates", "Output Data Type", "Control Codes",
                              "Command Times [s]", "Commanding Mode"):
                        put(k, None)
                    ref = TPL["D"].get("MOLA") or []
                    for k in ("Direction", "Command Timeout [s]", "Scan After Command"):
                        c = L(k)
                        if c and c - 1 < len(ref):
                            row[c - 1] = ref[c - 1]
            rows.append(row)
        for i, row in enumerate(rows):
            for c in range(ncol):
                cell = ws.cell(HEADER_ROWS + 1 + i, c + 1, value=row[c])
                cell._style = copy(styles[c])
        gerados[sheet] = len(rows)

    # UTR nova
    ws = wb["DNP3_RTUs"]
    lab = {ws.cell(HEADER_ROWS, c).value: c for c in range(1, ws.max_column + 1)
           if ws.cell(HEADER_ROWS, c).value}
    r = HEADER_ROWS + 1
    def putr(name, val):
        c = lab.get(name)
        if c: ws.cell(r, c).value = val
    putr("Remote Unit (Terminal Server) Name", RU)
    putr("Remote Unit Alias", f"{RU}__")
    putr("Remote Unit Custom ID", None)          # ADMS gera
    putr("Remote Unit AOR Group", AOR)
    putr("Remote Unit Description", FABRICANTE)
    putr("Container Name", CONTAINER)
    putr("Container Custom ID", None)

    from collections import Counter as _C
    fbc = _C((f["sigla"], f["molde"]) for f in usou_fallback)
    fbd = {f["sigla"]: f["desc"] for f in usou_fallback}
    _rel([[sg, mo, n, fbd.get(sg, "")] for (sg, mo), n in sorted(fbc.items(), key=lambda x: -x[1])])
    print(f"relatorio: {OUT_REL.name} ({len(mapa)} coords, {len(dups)} repetidas, "
          f"{len(usou_fallback)} por equivalencia)")

    buf = io.BytesIO(); wb.save(buf)
    OUT_TDT.write_bytes(excel_native.resave_native(buf.getvalue()))
    print(f"TDT: {OUT_TDT.name} | {dict(gerados)} | pulados (sem template): {pulados}")


if __name__ == "__main__":
    main()
