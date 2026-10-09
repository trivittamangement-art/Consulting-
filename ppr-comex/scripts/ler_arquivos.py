"""Leitores das planilhas do gerente, no formato em que ele já trabalha.

Cada função recebe o caminho do arquivo e devolve um DataFrame limpo. Os
leitores procuram os cabeçalhos pelo texto (não por posição fixa), então
aguentam linhas ou colunas a mais, abas novas por mês e pequenas variações
de nome. Nenhum nome de cliente fica fixo no código.
"""
import re
import unicodedata

import pandas as pd
from openpyxl import load_workbook

MESES = {"jan": 1, "fev": 2, "mar": 3, "abr": 4, "mai": 5, "jun": 6,
         "jul": 7, "ago": 8, "set": 9, "out": 10, "nov": 11, "dez": 12}


def norm(txt):
    """minúsculo, sem acento e sem espaços extras."""
    s = unicodedata.normalize("NFKD", str(txt or "")).encode("ascii", "ignore").decode()
    return re.sub(r"\s+", " ", s).strip().lower()


def mes_do_texto(txt):
    t = norm(txt)[:3]
    return MESES.get(t)


def numero(v):
    """Converte '1.234,56', 'R$ 1234,5', 1234.5 em float. Vazio ou '-' vira None."""
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).replace("R$", "").replace("\xa0", " ").strip()
    if s in ("", "-"):
        return None
    if "," in s:
        s = s.replace(".", "").replace(",", ".")
    s = re.sub(r"[^0-9.\-]", "", s)
    try:
        return float(s)
    except ValueError:
        return None


def partes(v, espacos=False):
    """Separa célula com vários valores ('a / b', 'a\\nb') em lista.

    Rótulos entre parênteses na linha de baixo ficam junto do valor de cima.
    espacos=True também separa por 3 ou mais espaços ('25267      25268')."""
    if v is None:
        return []
    if isinstance(v, (int, float)):
        return [v]
    txt = re.sub(r"\s*\n\s*\(", " (", str(v).strip())
    itens = re.split(r"\s*/\s*|\n+" + (r"|\s{3,}" if espacos else ""), txt)
    return [i.strip() for i in itens if i.strip() not in ("", "-")]


def data(v):
    if v is None or (isinstance(v, str) and v.strip() in ("", "-")):
        return pd.NaT
    return pd.to_datetime(v, errors="coerce", dayfirst=True)


def linhas(ws):
    return [list(r) for r in ws.iter_rows(values_only=True)]


def acha_cabecalho(rows, *palavras):
    """Índice da primeira linha em que cada palavra começa alguma palavra de uma célula.

    Palavra com '=' na frente exige célula exatamente igual (ex.: '=cliente')."""
    for i, r in enumerate(rows):
        cel = [norm(c) for c in r if c is not None]
        def casa(p):
            if p.startswith("="):
                return any(c == norm(p[1:]) for c in cel)
            p = norm(p)
            return any(c.startswith(p) or (" " + p) in (" " + c) for c in cel)
        if all(casa(p) for p in palavras):
            return i
    return None


def col(cab, *opcoes):
    """Posição da coluna cujo cabeçalho começa com uma das opções."""
    cn = [norm(c) for c in cab]
    for o in opcoes:
        o = norm(o)
        for j, c in enumerate(cn):
            if c.startswith(o):
                return j
    return None


# ------------------------------------------------------------------ meta 2

def ler_periodicidade(caminho):
    """Matriz cliente x mês. Devolve (pedidos, cadastro, totais).

    pedidos: cliente, ano, mes, qtd (uma linha por cliente e mês)
    cadastro: cliente, ano, frequencia_txt, compras_ano, periodicidade_meses, cumprido, observacao
    totais: linhas 'Total' e 'Meta' da aba, por mês
    """
    wb = load_workbook(caminho, data_only=True)
    ped, cad, tot = [], [], []
    for ws in wb.worksheets:
        m = re.search(r"(20\d\d)", ws.title)
        if not m:
            continue
        ano = int(m.group(1))
        rows = linhas(ws)
        h = acha_cabecalho(rows, "=cliente")
        if h is None:
            continue
        cab = rows[h]
        cols_mes = {j: mes_do_texto(c) for j, c in enumerate(cab) if isinstance(c, str) and mes_do_texto(c)}
        j_freq = col(cab, "frequencia")
        j_cump = col(cab, "cumprido", "previsao cumprida")
        j_obs = col(cab, "observacao", "justificativa")
        j_prev = col(cab, "previsao")
        for r in rows[h + 1:]:
            nome = r[0]
            if nome is None or str(nome).strip() == "":
                continue
            rotulo = norm(nome)
            if rotulo.startswith("total") or rotulo.startswith("meta") or isinstance(nome, (int, float)):
                tipo = "meta" if rotulo.startswith("meta") else "total"
                achou = re.search(r"20\d\d", str(nome))
                ref = int(achou.group(0)) if achou else ano
                for j, mes in cols_mes.items():
                    tot.append({"aba": ws.title, "tipo": tipo, "ano": ref, "mes": mes, "valor": numero(r[j]) or 0})
                continue
            cliente = re.sub(r"\s+", " ", str(nome)).strip()
            for j, mes in cols_mes.items():
                v = r[j] if j < len(r) else None
                q = numero(v) if not (isinstance(v, str) and v.strip().upper() == "P") else 0
                ped.append({"cliente": cliente, "ano": ano, "mes": mes, "qtd": q or 0,
                            "previsto_marcado": isinstance(v, str) and v.strip().upper() == "P"})
            freq = r[j_freq] if j_freq is not None and j_freq < len(r) else None
            fm = re.search(r"(\d+)", str(freq or ""))
            vezes = int(fm.group(1)) if fm else None
            cad.append({
                "cliente": cliente, "ano": ano, "frequencia_txt": freq,
                "compras_ano": vezes,
                "periodicidade_meses": round(12 / vezes, 1) if vezes else None,
                "cumprido": r[j_cump] if j_cump is not None and j_cump < len(r) else None,
                "observacao": r[j_obs] if j_obs is not None and j_obs < len(r) else None,
                "previsao": r[j_prev] if j_prev is not None and j_prev < len(r) and j_prev != j_cump else None,
            })
    return pd.DataFrame(ped), pd.DataFrame(cad), pd.DataFrame(tot)


# ------------------------------------------------------------------ meta 3a

def ler_sa(caminho):
    """Exportação de S.A. do sistema (uma aba por setor). Remove duplicadas."""
    wb = load_workbook(caminho, data_only=True)
    out, usuarios = [], []
    for ws in wb.worksheets:
        rows = linhas(ws)
        h = acha_cabecalho(rows, "codigo", "setor")
        if h is None:
            hu = acha_cabecalho(rows, "=usuarios")
            if hu is not None:
                for r in rows[hu + 1:]:
                    vals = [c for c in r if c is not None]
                    if len(vals) >= 2 and norm(vals[0]) != "total":
                        usuarios.append({"usuario": vals[0], "solicitacoes": numero(vals[1])})
            continue
        cab = rows[h]
        jc, jd, js = col(cab, "codigo"), col(cab, "data"), col(cab, "setor")
        jt, jx, ju = col(cab, "tipo"), col(cab, "descri"), col(cab, "usuario")
        jconc, jst = col(cab, "data conclus", "conclus", "data fim", "encerr"), col(cab, "status", "situa")
        for r in rows[h + 1:]:
            if r[jc] is None:
                continue
            out.append({
                "codigo": int(numero(r[jc])) if numero(r[jc]) else r[jc],
                "data_criacao": data(r[jd]), "setor": str(r[js] or "").strip(),
                "tipo": str(r[jt] or "").strip().upper(), "descricao": str(r[jx] or "").strip(),
                "usuario": str(r[ju] or "").strip().upper(), "aba": ws.title.strip(),
                "data_conclusao": data(r[jconc]) if jconc is not None else pd.NaT,
                "status": r[jst] if jst is not None else None,
            })
    df = pd.DataFrame(out)
    if len(df):
        df["duplicada"] = df.duplicated(subset=["codigo", "descricao"], keep="first")
        unicas = ~df["duplicada"]
        df["mesmo_codigo_outra_aba"] = unicas & df[unicas].duplicated(subset=["codigo"], keep=False).reindex(
            df.index, fill_value=False)
    return df, pd.DataFrame(usuarios)


def ler_fluxo(caminho, aba="Controle de Fluxo"):
    """Controle de fluxo do processo de exportação, recalculando os tempos pelas datas.

    A planilha original tem fórmulas que dão números como -45.694 quando uma data
    está vazia; aqui os tempos são refeitos só quando as duas datas existem.
    """
    wb = load_workbook(caminho, data_only=True)
    ws = next((w for w in wb.worksheets if norm(w.title) == norm(aba)), None)
    if ws is None:
        return pd.DataFrame()
    rows = linhas(ws)
    h = acha_cabecalho(rows, "invoice", "cliente")
    cab = rows[h]
    df = pd.DataFrame([r[:len(cab)] for r in rows[h + 1:] if r[col(cab, "cliente")]], columns=cab)
    etapas = [
        ("Exportação → Almoxarifado (pedir embalagem)", "Data de Emissão do Pedido", "Data de solicitação de compra de embalagem"),
        ("Almoxarifado (abrir solicitação de compra)", "Data de solicitação de compra de embalagem", "Abertura da solicitação de compra"),
        ("Compras (pedido ao fornecedor)", "Abertura da solicitação de compra", "Emissão do pedido pelo Compras"),
        ("Fornecedor/Marketing (arte para aprovação)", "Emissão do pedido pelo Compras", "Recebimento Marketing de arte"),
        ("Marketing → Exportação (arte)", "Recebimento Marketing de arte", "Recebimento Exportação de arte"),
        ("Exportação → Laboratório (arte)", "Recebimento Exportação de arte", "Envio da arte para o Laboratório"),
        ("Laboratório/Qualidade (aprovação)", "Envio da arte para o Laboratório", "Retorno de aprovação do Laboratório"),
        ("Marketing → Fornecedor (arte final)", "Retorno de aprovação do Laboratório", "Envio de aprovação de arte ao fornecedor"),
        ("Fornecedor (entrega da embalagem)", "Envio de aprovação de arte ao fornecedor", "Data de recebimento da embalagem"),
        ("PCP (programar produção)", "Data de recebimento da embalagem", "Início de produção"),
        ("Produção", "Início de produção", "Término de produção"),
        ("Logística (carregamento)", "Término de produção", "Data de carregamento"),
        ("Transporte até o terminal", "Data de carregamento", "Data de chegada no terminal"),
        ("Terminal → porto", "Data de chegada no terminal", "Data de chegada no porto"),
        ("Porto → saída do navio", "Data de chegada no porto", "Data de saída do navio"),
    ]
    nomes = {norm(c): c for c in df.columns if c}

    def achar(prefixo):
        p = norm(prefixo)
        return next((nomes[k] for k in nomes if k.startswith(p)), None)

    for c in df.columns:
        if c and ("Data" in str(c) or "Previsão" in str(c) or "Abertura" in str(c) or "Emissão" in str(c)
                  or "Recebimento" in str(c) or "Envio" in str(c) or "Retorno" in str(c) or "Início" in str(c)
                  or "Término" in str(c)):
            df[c] = df[c].map(data)
    tempos = {}
    for nome, a, b in etapas:
        ca, cb = achar(a), achar(b)
        if ca and cb:
            d = (df[cb] - df[ca]).dt.days
            tempos[nome] = d.where((d >= 0) & (d < 365))
    t = pd.DataFrame(tempos)
    ini, fim = achar("Data de Emissão do Pedido"), achar("Data de saída do navio")
    prev = achar("Previsão de entrega inicial")
    out = pd.DataFrame({
        "invoice": df[achar("Nº Invoice")], "pedido": df[achar("Nº Pedido")], "cliente": df[achar("Cliente")],
        "pais": df[achar("País")], "data_pedido": df[ini], "previsao_saida": df[prev], "saida_navio": df[fim],
        "tinha_estoque": df[achar("Possui estoque")], "status": df[achar("Status do Processo")],
        "motivo": df[achar("Motivo do Atraso")],
    })
    total = (out["saida_navio"] - out["data_pedido"]).dt.days
    out["dias_total"] = total.where((total >= 0) & (total < 365))
    atraso = (out["saida_navio"] - out["previsao_saida"]).dt.days
    out["dias_vs_previsao"] = atraso.where(atraso.abs() < 365)
    out["data_pedido_suspeita"] = (out["data_pedido"] > out["saida_navio"]) | \
        (out["data_pedido"] > out["previsao_saida"] + pd.Timedelta(days=60))
    return pd.concat([out, t], axis=1)


# ------------------------------------------------------------------ meta 3b

def ler_gastos(caminho, aba="Controle de Gastos"):
    wb = load_workbook(caminho, data_only=True)
    ws = next((w for w in wb.worksheets if norm(w.title) == norm(aba)), None)
    if ws is None:
        return pd.DataFrame()
    rows = linhas(ws)
    h = acha_cabecalho(rows, "invoice", "cliente")
    cab = [c for c in rows[h]]
    df = pd.DataFrame([r[:len(cab)] for r in rows[h + 1:]], columns=cab)
    df = df.loc[:, [c for c in df.columns if c]]
    df = df[df[cab[col(cab, "cliente")]].notna()]
    for c in df.columns:
        if norm(c).startswith("data") or norm(c).startswith("mes"):
            df[c] = df[c].map(data)
        elif not norm(c) in ("cliente", "pais", "continente"):
            df[c] = df[c].map(numero)
    return df.reset_index(drop=True)


def ler_faturado_cambio(caminho):
    wb = load_workbook(caminho, data_only=True)
    ws = wb.worksheets[0]
    rows = linhas(ws)
    h = acha_cabecalho(rows, "invoice", "cliente")
    cab = rows[h]
    ji, jc, jv, jd, jr = (col(cab, "invoice"), col(cab, "cliente"), col(cab, "valor nf"),
                          col(cab, "data nf"), col(cab, "total recebido"))
    out = []
    for r in rows[h + 1:]:
        if r[jc] is None:
            continue
        out.append({"invoice": str(r[ji]).strip(), "cliente": str(r[jc]).strip(),
                    "valor_nf": numero(r[jv]), "data_nf": data(r[jd]) if not isinstance(r[jd], str) else data(str(r[jd])[:10]),
                    "data_nf_texto": r[jd] if isinstance(r[jd], str) else None,
                    "recebido_cambio": numero(r[jr])})
    df = pd.DataFrame(out)
    df["duplicada"] = df.duplicated(subset=["invoice", "valor_nf"], keep="first")
    df["variacao"] = df["recebido_cambio"] - df["valor_nf"]
    return df


# ------------------------------------------------------------------ meta 3c

def ler_uso_embalagens(caminho, ano=None):
    """Abas mensais 'Uso de Embalagens Obsoletas'. Uma linha por código de embalagem.

    Células com vários códigos ('22469 / 22468') e vários valores ('R$ 430,83\\nR$ 114')
    são separadas; quando há um valor só para vários códigos, o valor fica na
    primeira linha e as demais ficam com valor zero, para não duplicar soma.
    """
    wb = load_workbook(caminho, data_only=True)
    out = []
    for ws in wb.worksheets:
        mes = mes_do_texto(ws.title)
        if not mes:
            continue
        rows = linhas(ws)
        h = acha_cabecalho(rows, "cliente", "invoice")
        if h is None:
            continue
        cab = rows[h]
        jcli, jinv, jdesc, jcod, jtipo = (col(cab, "cliente"), col(cab, "invoice"), col(cab, "descri"),
                                          col(cab, "codigo"), col(cab, "tipo"))
        qtds = [j for j, c in enumerate(cab) if norm(c).startswith("qtde") and "dias" not in norm(c)]
        jqprod = qtds[0] if len(qtds) > 1 else None
        jqemb = qtds[-1] if qtds else None
        jval, jdias = col(cab, "valor"), col(cab, "qtde dias", "dias")
        for r in rows[h + 1:]:
            if r[jcli] is None:
                continue
            cods = partes(r[jcod])
            vals = [numero(v) for v in partes(r[jval])]
            qs = [numero(v) for v in partes(r[jqemb])] if jqemb is not None else []
            ds = [numero(v) for v in partes(r[jdias])] if jdias is not None else []
            tipos = partes(r[jtipo]) if jtipo is not None else []
            n = max(1, len(cods))
            for k in range(n):
                cod = cods[k] if k < len(cods) else None
                obs = None
                if cod:
                    m = re.search(r"\(([^)]*)\)", str(cod))
                    obs = m.group(1) if m else None
                    cod = re.sub(r"\s*\(.*?\)\s*", "", str(cod)).strip()
                valor = vals[k] if len(vals) == n else (vals[0] if k == 0 and vals else 0)
                out.append({
                    "ano": ano, "mes": mes, "cliente": str(r[jcli]).strip(),
                    "invoice": str(r[jinv]).strip() if r[jinv] is not None else None,
                    "qtd_caixas_produto": numero(r[jqprod]) if jqprod is not None else None,
                    "descricao": str(r[jdesc] or "").strip(),
                    "codigo_embalagem": cod, "referencia": obs,
                    "tipo": (tipos[k] if k < len(tipos) else (tipos[0] if tipos else None)),
                    "qtd_embalagem": qs[k] if len(qs) == n else (qs[0] if k == 0 and qs else None),
                    "valor": valor or 0,
                    "dias_parada": ds[k] if len(ds) == n else (ds[0] if ds else None),
                    "celula_com_varios_valores": n > 1,
                })
    return pd.DataFrame(out)


def ler_estoque_parado(caminho):
    """Abas 'Valor do Estoque' e 'Quantidade em Estoque' (faixas de dias sem movimentação)."""
    wb = load_workbook(caminho, data_only=True)
    res = {}
    for chave, aba in (("valor", "valor do estoque"), ("quantidade", "quantidade em estoque")):
        ws = next((w for w in wb.worksheets if norm(w.title).startswith(aba)), None)
        if ws is None:
            continue
        rows = [[c for c in r if c is not None] for r in linhas(ws)]
        out = []
        bloco, datas = None, []
        for r in rows:
            if not r:
                continue
            if len(r) > 1 and all(isinstance(c, (pd.Timestamp,)) or hasattr(c, "year") for c in r[1:]):
                bloco, datas = str(r[0]).strip(), [pd.Timestamp(c) for c in r[1:]]
                continue
            if bloco and len(r) > 1:
                for d, v in zip(datas, r[1:]):
                    out.append({"bloco": bloco, "faixa": str(r[0]).strip(), "data": d, "valor": numero(v)})
        res[chave] = pd.DataFrame(out)
    return res


def ler_estoque_cliente(caminho, ano=None):
    """Abas mensais 'Jan - 26' etc. com estoque de embalagem por código."""
    wb = load_workbook(caminho, data_only=True)
    out = []
    for ws in wb.worksheets:
        mes = mes_do_texto(ws.title)
        if not mes:
            continue
        m = re.search(r"(\d{2})\s*$", ws.title.strip())
        a = 2000 + int(m.group(1)) if m else ano
        rows = linhas(ws)
        grupo = None
        for r in rows:
            vals = [c for c in r if c is not None]
            if not vals:
                continue
            txt = norm(vals[0])
            if txt.startswith("estoque"):
                grupo = str(vals[0]).strip()
                continue
            if txt.startswith("cod"):
                continue
            # bloco da esquerda: código, cód. fabricação, descrição, estoque
            cods, fab, desc, est = r[1], r[2], r[3], r[4]
            for k, cod in enumerate(partes(cods, espacos=True)):
                e = partes(est, espacos=True)
                out.append({"ano": a, "mes": mes, "grupo": grupo, "codigo": str(cod).strip(),
                            "cod_fabricacao": str(fab or "").strip(), "descricao": str(desc or "").strip(),
                            "estoque": numero(e[k]) if k < len(e) else None, "lado": "embalagem"})
            # bloco da direita: caixa externa
            if len(r) > 7 and r[6] is not None and norm(r[6])[:3] != "cod":
                out.append({"ano": a, "mes": mes, "grupo": "Caixa externa", "codigo": str(r[6]).strip(),
                            "cod_fabricacao": "", "descricao": "Caixa externa", "estoque": numero(r[7]),
                            "lado": "caixa externa"})
    df = pd.DataFrame(out)
    return df[df["codigo"].notna() & (df["codigo"] != "None")] if len(df) else df
