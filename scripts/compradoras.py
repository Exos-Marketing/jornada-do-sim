"""Compradoras do produto na TMB (API oficial) e na Voomp (coletor do navegador), com contato.

O contato (nome, e-mail, telefone) só vai para a lista de ação. Para cruzar com a planilha da Anhanguera usa-se
o e-mail embaralhado, com a mesma regra da skill acompanhar-vendas (SHA-256 do texto aparado e em minúsculas,
12 hex), e o telefone como reserva.
"""
import datetime as dt
import hashlib
import json
import re
import urllib.error
import urllib.parse
import urllib.request

TMB_API = "https://api.tmbeducacao.com.br/"
PAGO = "pago"
CANCELADO = "cancelado"
PENDENTE = "pendente"


def hash_id(texto):
    """Mesmo hash da acompanhar-vendas e dos coletores do navegador."""
    texto = str(texto).strip().lower() if texto is not None else ""
    return hashlib.sha256(texto.encode("utf-8")).hexdigest()[:12] if texto else None


def hash_doc(texto):
    return hash_id(re.sub(r"\D", "", str(texto or "")))


def fone(texto):
    """Telefone só com dígitos, sem o 55 do país: os últimos 11 (DDD + número). Vazio se curto demais."""
    dig = re.sub(r"\D", "", str(texto or ""))
    return dig[-11:] if len(dig) >= 10 else ""


BRT = dt.timezone(dt.timedelta(hours=-3))  # Brasília, sem horário de verão desde 2019


def quando(texto):
    """Data e hora em Brasília (sem fuso no objeto). A Voomp grava em UTC ('...Z'); a TMB, no horário local ou com -03:00."""
    texto = str(texto or "").strip()
    if not texto:
        return None
    texto = texto.replace("Z", "+00:00").replace(" ", "T", 1)
    if "." in texto:  # o fromisoformat do 3.9 quer 3 ou 6 casas nos segundos fracionários
        base, resto = texto.split(".", 1)
        fusos = [i for i in (resto.find("+"), resto.find("-")) if i >= 0]
        corte = min(fusos) if fusos else len(resto)
        texto = base + "." + resto[:corte][:6].ljust(6, "0") + resto[corte:]
    try:
        d = dt.datetime.fromisoformat(texto)
    except ValueError:
        try:
            return dt.datetime.fromisoformat(texto[:10])
        except ValueError:
            return None
    if d.tzinfo is not None:
        d = d.astimezone(BRT).replace(tzinfo=None)
    return d


def _dia(texto):
    d = quando(texto)
    return d.date() if d else None


def coletar_tmb(produtos, token):
    """Pedidos dos produtos na API da TMB, com contato. Devolve o dicionário gravado em leituras/tmb-*.json."""
    def get(caminho, **parametros):
        url = TMB_API + caminho + "?" + urllib.parse.urlencode(parametros)
        pedido = urllib.request.Request(url, headers={"Authorization": "Bearer " + token.strip(),
                                                      "Accept": "application/json",
                                                      "User-Agent": "exos-jornada-do-sim"})
        with urllib.request.urlopen(pedido, timeout=60) as resposta:
            return json.loads(resposta.read().decode("utf-8"))

    saida = {"plataforma": "tmb", "coletado_em": dt.datetime.now(dt.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")}
    linhas = []
    try:
        for produto in sorted(int(p) for p in produtos):
            pagina = 1
            while True:
                corpo = get("api/pedidos", produto_id=produto, pageNumber=pagina, pageSize=100)
                for p in corpo.get("data") or []:
                    linhas.append({
                        "pedido": p.get("pedido_id"), "produto_id": p.get("produto_id"),
                        "status": p.get("status_pedido"), "status_financeiro": p.get("status_financeiro"),
                        "criado": p.get("criado_em"), "pago": p.get("data_efetivado"),
                        "nome": p.get("cliente"), "email": p.get("email"), "telefone": p.get("telefone"),
                        "h_email": hash_id(p.get("email")), "h_doc": hash_doc(p.get("documento")),
                    })
                if not corpo.get("data") or pagina >= (corpo.get("totalPages") or 1):
                    break
                pagina += 1
    except urllib.error.HTTPError as erro:
        return dict(saida, erro="token_invalido" if erro.code in (401, 403) else "http_%d" % erro.code)
    except urllib.error.URLError as erro:
        return dict(saida, erro="rede: %s" % erro.reason)
    return dict(saida, linhas=linhas)


def _situacao_tmb(status):
    s = str(status or "").lower()
    if s == "efetivado":
        return PAGO
    if "cancel" in s:
        return CANCELADO
    return PENDENTE


def _situacao_voomp(linha):
    s = str(linha.get("status") or "").lower()
    if linha.get("reembolsado") or s in ("refunded", "chargedback", "chargeback", "canceled", "cancelled", "refund"):
        return CANCELADO
    if s == "paid":
        return PAGO
    return PENDENTE


def normalizar(leitura):
    """Linhas de uma leitura (TMB da jornada, Voomp da jornada ou Voomp da acompanhar-vendas) num formato só."""
    plataforma = leitura.get("plataforma")
    saida = []
    for l in leitura.get("linhas") or []:
        if plataforma == "tmb":
            situacao, pago_em = _situacao_tmb(l.get("status")), _dia(l.get("pago"))
            pedido = l.get("pedido")
        else:
            situacao, pago_em = _situacao_voomp(l), _dia(l.get("pago"))
            pedido = l.get("id")
        saida.append({
            "plataforma": plataforma, "pedido": pedido, "produto_id": l.get("produto_id"),
            "situacao": situacao, "pago_em": pago_em if situacao == PAGO else None,
            "comprou_quando": quando(l.get("criado")) or quando(l.get("pago")),
            "status_financeiro": l.get("status_financeiro"),
            "nome": l.get("nome"), "email": l.get("email"), "telefone": l.get("telefone"),
            "h_email": l.get("h_email") or hash_id(l.get("email")), "h_doc": l.get("h_doc"),
            "fone": fone(l.get("telefone")),
        })
    return saida


def pessoas(linhas):
    """Uma entrada por pessoa (e-mail embaralhado, senão documento): fica quem tem ao menos um pagamento válido.

    Cobrança em dobro e recompra viram uma pessoa só, com a data do primeiro pagamento. Pedido pago e depois
    reembolsado, estornado ou cancelado deixa de valer (a Voomp troca o status da própria venda).
    """
    grupos = {}
    for l in linhas:
        chave = l["h_email"] or (("d:" + l["h_doc"]) if l.get("h_doc") else None) or (("f:" + l["fone"]) if l["fone"] else None)
        if chave:
            grupos.setdefault(chave, []).append(l)
    saida = []
    for chave, itens in grupos.items():
        pagos = sorted((i for i in itens if i["situacao"] == PAGO and i["pago_em"]), key=lambda i: i["pago_em"])
        if not pagos:
            continue
        com_contato = [i for i in pagos if i.get("nome") or i.get("telefone") or i.get("email")]
        contato = com_contato[0] if com_contato else {}
        compras = [i["comprou_quando"] for i in pagos if i.get("comprou_quando")]
        saida.append({
            "chave": "h:" + chave if not chave.startswith(("d:", "f:")) else chave,
            "h_email": pagos[0]["h_email"], "h_doc": pagos[0].get("h_doc"),
            "pago_em": pagos[0]["pago_em"], "comprou_quando": min(compras) if compras else None,
            "plataformas": sorted({i["plataforma"] for i in pagos}),
            "pedidos": len(pagos), "status_financeiro": next((i["status_financeiro"] for i in reversed(pagos)
                                                              if i.get("status_financeiro")), None),
            "nome": contato.get("nome"), "email": contato.get("email"), "telefone": contato.get("telefone"),
            "fone": next((i["fone"] for i in pagos if i["fone"]), ""),
        })
    return sorted(saida, key=lambda p: (p["pago_em"], p["chave"]))
