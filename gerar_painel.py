"""Gera o painel do Edifício Urbano a partir da planilha de prestação de contas.

Uso:
    python gerar_painel.py --excel "caminho/da/planilha.xlsx" [--saida site/index.html] [--titulo "Nome do prédio"]

O script lê as tabelas nomeadas da aba "Extrato", recalcula tudo em Python
(não depende das fórmulas do Excel) e injeta os dados no template.html.
Se alguma tabela ou coluna esperada não existir, ele para com erro e não gera
nada, para o site nunca ser publicado com número errado.
"""
from __future__ import annotations

import argparse
import datetime as dt
import json
import os
import sys
from pathlib import Path

try:
    import openpyxl
except ImportError:
    sys.exit("Falta a biblioteca openpyxl. Rode: pip install openpyxl")

AQUI = Path(__file__).resolve().parent
ABA = "Extrato"
MESES = ["jan", "fev", "mar", "abr", "mai", "jun", "jul", "ago", "set", "out", "nov", "dez"]
FORN_TRANSFERENCIA = "recomposição de caixinha"
FORN_CONTRAPARTIDA = "contrapartida"


class ErroPlanilha(Exception):
    pass


# ── leitura ──────────────────────────────────────────────────────────────────
def ler_tabela(ws, nome: str, colunas: list[str]) -> list[dict]:
    """Lê uma tabela nomeada do Excel e devolve uma lista de dicionários.

    As colunas são localizadas pelo nome do cabeçalho (sem diferenciar
    maiúsculas), então mudar a ordem das colunas não quebra nada.
    """
    if nome not in ws.tables:
        raise ErroPlanilha(f"Tabela '{nome}' não encontrada na aba {ABA}. Tabelas existentes: {list(ws.tables)}")
    linhas = list(ws[ws.tables[nome].ref])
    cab = [str(c.value or "").strip().lower() for c in linhas[0]]
    idx = {}
    for col in colunas:
        if col.lower() not in cab:
            raise ErroPlanilha(f"Coluna '{col}' não encontrada na tabela '{nome}'. Cabeçalho atual: {cab}")
        idx[col] = cab.index(col.lower())
    saida = []
    for linha in linhas[1:]:
        valores = {col: linha[i].value for col, i in idx.items()}
        saida.append(valores)
    return saida


def num(v) -> float:
    if v is None or v == "":
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    if isinstance(v, str) and v.startswith("="):
        raise ErroPlanilha(f"Encontrei uma fórmula onde esperava um número: {v}. Abra a planilha no Excel e salve de novo.")
    try:
        return float(str(v).replace(".", "").replace(",", "."))
    except ValueError as e:
        raise ErroPlanilha(f"Valor não numérico: {v!r}") from e


def txt(v) -> str:
    return "" if v is None else str(v).strip()


def fmt_data(v) -> str:
    if isinstance(v, dt.datetime):
        return v.strftime("%d/%m/%Y")
    if isinstance(v, dt.date):
        return v.strftime("%d/%m/%Y")
    return txt(v)


def fmt_mes(v) -> str:
    """Converte a célula de mês (data ou texto 'mai/26') para 'mai/26'."""
    if isinstance(v, (dt.datetime, dt.date)):
        return f"{MESES[v.month - 1]}/{v.strftime('%y')}"
    return txt(v).lower()


def chave_mes(m: str) -> tuple[int, int]:
    nome, ano = m.split("/")
    return (int(ano), MESES.index(nome))


def r2(v: float) -> float:
    return round(v + 0.0, 2)


# ── cálculo ──────────────────────────────────────────────────────────────────
def montar_dados(caminho_excel: Path) -> dict:
    wb = openpyxl.load_workbook(caminho_excel, data_only=True)
    if ABA not in wb.sheetnames:
        raise ErroPlanilha(f"Aba '{ABA}' não encontrada. Abas: {wb.sheetnames}")
    ws = wb[ABA]

    props_xl = ler_tabela(ws, "Tabela1", ["Proprietarios", "Apartamentos", "Percentual"])
    banco_xl = ler_tabela(ws, "Tabela3", ["Data", "Mês", "Proprietário", "Fornecedor", "receita", "despesa", "Descrição"])
    caix_xl = ler_tabela(ws, "Tabela5", ["Data", "Mês", "Fornecedor", "Descrição", "Despesa", "Receita"])
    rateio_xl = ler_tabela(ws, "RateioContrapartida", ["Apartamento", "Proprietário", "Contrapartida"])
    futuros_xl = ler_tabela(ws, "debitosFuturos", ["Data", "Mês/ano", "Propretário", "Fornecedor", "receita", "despesa", "Descrição"])

    # Saldo inicial da caixinha (célula ao lado de "Caixinha inicial G+P")
    saldo_ini_caix = 0.0
    for row in ws.iter_rows(min_row=1, max_row=40, min_col=1, max_col=8):
        for c in row:
            if isinstance(c.value, str) and c.value.strip().lower().startswith("caixinha inicial"):
                saldo_ini_caix = num(ws.cell(c.row, c.column + 1).value)

    # Proprietários
    proprietarios = []
    for p in props_xl:
        if not txt(p["Proprietarios"]):
            continue
        proprietarios.append({"nome": txt(p["Proprietarios"]), "apto": txt(p["Apartamentos"]), "pct": num(p["Percentual"])})
    if not proprietarios:
        raise ErroPlanilha("Nenhum proprietário na Tabela1.")
    soma_pct = sum(p["pct"] for p in proprietarios)
    if abs(soma_pct - 1) > 0.001:
        raise ErroPlanilha(f"Os percentuais da Tabela1 somam {soma_pct:.4f}, deveriam somar 1 (100%).")
    nomes = {p["nome"] for p in proprietarios}

    # Extrato Bradesco
    extrato, saldo = [], 0.0
    for linha in banco_xl:
        if linha["Data"] is None and not txt(linha["Descrição"]):
            continue
        if linha["Data"] is None or linha["Mês"] is None:
            raise ErroPlanilha(f"Lançamento sem Data ou Mês na Tabela3: {linha}")
        receita, despesa = num(linha["receita"]), num(linha["despesa"])
        forn, prop = txt(linha["Fornecedor"]), txt(linha["Proprietário"])
        if prop and prop not in nomes:
            raise ErroPlanilha(f"Proprietário '{prop}' na Tabela3 não existe na Tabela1 (o nome tem que ser idêntico).")
        saldo += receita - despesa
        extrato.append({
            "serial": len(extrato) + 1,
            "data": fmt_data(linha["Data"]),
            "mes": fmt_mes(linha["Mês"]),
            "proprietario": prop,
            "fornecedor": forn,
            "descricao": txt(linha["Descrição"]),
            "receita": r2(receita),
            "despesa": r2(despesa),
            "saldo": r2(saldo),
            "transf": forn.lower() == FORN_TRANSFERENCIA,
            "rendim": despesa < 0,
        })

    # Caixinha G+P
    caix_extrato, saldo_c = [], saldo_ini_caix
    for linha in caix_xl:
        if linha["Data"] is None and not txt(linha["Fornecedor"]):
            continue
        receita, despesa = num(linha["Receita"]), num(linha["Despesa"])
        saldo_c += receita - despesa
        caix_extrato.append({
            "data": fmt_data(linha["Data"]),
            "mes": fmt_mes(linha["Mês"]),
            "fornecedor": txt(linha["Fornecedor"]),
            "descricao": txt(linha["Descrição"]),
            "receita": r2(receita),
            "despesa": r2(despesa),
            "saldo": r2(saldo_c),
        })

    # Rateio especial da Contrapartida (por área acrescida)
    contrap = {}
    for linha in rateio_xl:
        if linha["Apartamento"] is None:
            continue
        contrap[txt(linha["Apartamento"])] = r2(num(linha["Contrapartida"]))

    # Parte de cada proprietário = mesma lógica da coluna "Valor devido" do Excel
    base = sum(e["despesa"] for e in extrato if e["fornecedor"].lower() not in (FORN_TRANSFERENCIA, FORN_CONTRAPARTIDA))
    base += sum(c["despesa"] for c in caix_extrato)
    for p in proprietarios:
        p["pago"] = r2(sum(e["receita"] for e in extrato if e["proprietario"] == p["nome"]))
        p["parte"] = r2(base * p["pct"] + contrap.get(p["apto"], 0.0))

    # Projeções
    projecoes = []
    for linha in futuros_xl:
        if linha["Data"] is None and not txt(linha["Descrição"]):
            continue
        projecoes.append({
            "data": fmt_data(linha["Data"]),
            "mes": fmt_mes(linha["Mês/ano"]),
            "proprietario": txt(linha["Propretário"]),
            "fornecedor": txt(linha["Fornecedor"]),
            "descricao": txt(linha["Descrição"]),
            "receita": r2(num(linha["receita"])),
            "despesa": r2(num(linha["despesa"])),
        })

    meses = sorted({e["mes"] for e in extrato}, key=chave_mes)
    gerado = dt.datetime.fromtimestamp(os.path.getmtime(caminho_excel)).strftime("%d/%m/%Y %H:%M")

    return {
        "proprietarios": proprietarios,
        "extrato": extrato,
        "geradoEm": gerado,
        "mesesOrdem": meses,
        "rateioEspecial": {"Contrapartida": contrap},
        "caixinha": {"saldoInicial": r2(saldo_ini_caix), "extrato": caix_extrato},
        "projecoes": projecoes,
    }


def conferir_com_excel(caminho_excel: Path, dados: dict) -> list[str]:
    """Compara o que o Python calculou com os valores que o próprio Excel
    calculou (se a planilha foi salva pelo Excel). Devolve avisos."""
    avisos = []
    ws = openpyxl.load_workbook(caminho_excel, data_only=True)[ABA]
    try:
        cache = ler_tabela(ws, "Tabela1", ["Proprietarios", "Valor devido", "Valor Pago"])
    except ErroPlanilha:
        return avisos
    por_nome = {p["nome"]: p for p in dados["proprietarios"]}
    for linha in cache:
        nome = txt(linha["Proprietarios"])
        if nome not in por_nome or not isinstance(linha["Valor devido"], (int, float)):
            continue
        p = por_nome[nome]
        if abs(p["parte"] - linha["Valor devido"]) > 0.05:
            avisos.append(f"{nome}: parte calculada {p['parte']:.2f} diferente do Excel {linha['Valor devido']:.2f}")
        if isinstance(linha["Valor Pago"], (int, float)) and abs(p["pago"] - linha["Valor Pago"]) > 0.05:
            avisos.append(f"{nome}: pago calculado {p['pago']:.2f} diferente do Excel {linha['Valor Pago']:.2f}")
    return avisos


def gerar_html(dados: dict, template: Path, titulo: str) -> str:
    html = template.read_text(encoding="utf-8").replace("{{TITULO}}", titulo)
    for chave, valor in dados.items():
        marcador = f"/*DADOS:{chave}*/null"
        if marcador not in html:
            raise ErroPlanilha(f"Marcador {marcador} não encontrado no template.")
        html = html.replace(marcador, json.dumps(valor, ensure_ascii=False, separators=(",", ":")))
    if "/*DADOS:" in html:
        raise ErroPlanilha("Sobrou marcador sem dado no template.")
    return html


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--excel", required=True, type=Path)
    ap.add_argument("--saida", type=Path, default=AQUI / "site" / "index.html")
    ap.add_argument("--template", type=Path, default=AQUI / "template.html")
    ap.add_argument("--titulo", default="Edifício Exemplo — Centro", help="Nome exibido no topo do painel")
    args = ap.parse_args()

    if not args.excel.exists():
        print(f"ERRO: planilha não encontrada em {args.excel}")
        return 1
    try:
        dados = montar_dados(args.excel)
        avisos = conferir_com_excel(args.excel, dados)
        if avisos:
            print("ERRO: os números calculados não batem com os da planilha. Nada foi gerado.")
            for a in avisos:
                print("  -", a)
            return 1
        html = gerar_html(dados, args.template, args.titulo)
    except ErroPlanilha as e:
        print(f"ERRO: {e}")
        print("Nada foi gerado. O site continua na última versão publicada.")
        return 1

    args.saida.parent.mkdir(parents=True, exist_ok=True)
    args.saida.write_text(html, encoding="utf-8")
    saldo = dados["extrato"][-1]["saldo"] if dados["extrato"] else 0
    print(f"Painel gerado em {args.saida}")
    print(f"  {len(dados['extrato'])} lançamentos | saldo Bradesco R$ {saldo:,.2f} | dados de {dados['geradoEm']}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
