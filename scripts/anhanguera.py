"""Leitor da planilha de engajamento da Anhanguera (Pós): Base de Alunos e Desempenho Acadêmico.

Porta para Python puro (só openpyxl) as regras de `Exos/Engajamento Pós/analisar_engajamento_pos.py`:
data de corte, disciplinas concluídas, ritmo esperado e os grupos de ação. Acrescenta a cor do semáforo
da jornada do sim.
"""
import datetime as dt
import re
import unicodedata
from collections import Counter

# Ordem = prioridade: cada aluno cai no PRIMEIRO grupo cuja regra casar (igual ao analisador antigo).
SEGMENTOS = [
    ("nunca", "Nunca acessou"),
    ("contrato", "Contrato pendente"),
    ("adormecido", "Adormecido (30+ dias sem acesso)"),
    ("novato", "Novato (menos de 45 dias)"),
    ("travado", "Travado na largada (0 disciplina)"),
    ("atrasado", "Atrasado (2+ disciplinas)"),
    ("emdia", "No ritmo"),
    ("destaque", "Destaque (adiantado)"),
]
SEGMENTO_NOME = dict(SEGMENTOS)
COLUNAS_BASE = ["curso", "id", "nome", "email", "celular", "status", "contrato",
                "ultimo", "inicio", "fim", "dias_integr", "turma"]
COLUNAS_DESEMPENHO = ["id", "aluno", "seq", "disc", "tent", "nota", "turma"]


def sem_acento(txt):
    return "".join(c for c in unicodedata.normalize("NFD", str(txt or "").lower()) if unicodedata.category(c) != "Mn")


def data(valor):
    """Célula de data da planilha (datetime, date ou texto dd/mm/aaaa ou aaaa-mm-dd) em date; vazio vira None."""
    if valor is None or valor == "":
        return None
    if isinstance(valor, dt.datetime):
        return valor.date()
    if isinstance(valor, dt.date):
        return valor
    texto = str(valor).strip()
    for formato in ("%d/%m/%Y", "%d/%m/%Y %H:%M", "%d/%m/%Y %H:%M:%S", "%Y-%m-%d", "%Y-%m-%d %H:%M:%S"):
        try:
            return dt.datetime.strptime(texto, formato).date()
        except ValueError:
            pass
    return None


def numero(valor):
    try:
        return float(valor)
    except (TypeError, ValueError):
        return None


def _linhas(aba, colunas):
    saida = []
    for valores in aba.iter_rows(min_row=2, values_only=True):
        valores = list(valores)[:len(colunas)]
        if all(v is None or str(v).strip() == "" for v in valores):
            continue
        saida.append(dict(zip(colunas, valores + [None] * (len(colunas) - len(valores)))))
    return saida


def cor_pos(segmento, aprovadas):
    """Cor do semáforo para a Pós, a partir do grupo do analisador."""
    if segmento in ("nunca", "contrato"):
        return "nao_entrou"
    if segmento in ("emdia", "destaque"):
        return "verde"
    if segmento == "adormecido":
        return "laranja" if aprovadas >= 1 else "vermelho"
    return "amarelo"  # novato, travado e atrasado: entram, mas ainda sem ritmo


def data_de_corte(alunos):
    """Data do retrato: Data Final − Dias que faltam para integralização (a mais comum); senão, o último acesso mais recente."""
    implicitas = [a["fim"] - dt.timedelta(days=int(a["dias_integr"])) for a in alunos
                  if a["fim"] and a["dias_integr"] is not None]
    if implicitas:
        return Counter(implicitas).most_common(1)[0][0]
    acessos = [a["ultimo"] for a in alunos if a["ultimo"]]
    return max(acessos) if acessos else None


def segmento(a, n_disc):
    if a["ultimo"] is None:
        return "nunca"
    if str(a["contrato"] or "").strip().lower() == "pendente":
        return "contrato"
    if a["dias_sem_acesso"] is not None and a["dias_sem_acesso"] > 30:
        return "adormecido"
    if a["dias_curso"] is not None and a["dias_curso"] < 45:
        return "novato"
    if a["aprovadas"] == 0:
        return "travado"
    if a["atraso"] >= 2:
        return "atrasado"
    if a["atraso"] <= -1 or a["aprovadas"] >= 5:
        return "destaque"
    return "emdia"


def ler_planilha(caminho, cfg):
    """Lê a exportação e devolve data de corte, alunos ativos (com grupo e cor), canceladas e possíveis contas internas."""
    import openpyxl

    pos = cfg.get("anhanguera", cfg)
    livro = openpyxl.load_workbook(caminho, read_only=True, data_only=True)
    base = _linhas(livro["Base de Alunos"], COLUNAS_BASE)
    desempenho = _linhas(livro["Desempenho Acadêmico"], COLUNAS_DESEMPENHO)
    livro.close()

    n_disc = int(pos.get("n_disciplinas", 9))
    duracao = int(pos.get("duracao_dias", 364))
    nota_min = float(pos.get("nota_minima", 70))
    passo = duracao / n_disc
    excluir_ids = {str(i) for i in pos.get("excluir_ids", [])}
    excluir_emails = {e.lower() for e in pos.get("excluir_emails", [])}

    alunos = []
    for b in base:
        a = dict(b)
        a["id"] = str(int(b["id"])) if isinstance(b["id"], float) and b["id"].is_integer() else str(b["id"] or "").strip()
        a["email"] = str(b["email"] or "").strip()
        a["celular"] = str(b["celular"] or "").strip()
        a["status"] = str(b["status"] or "").strip()
        a["contrato"] = str(b["contrato"] or "").strip()
        for campo in ("ultimo", "inicio", "fim"):
            a[campo] = data(b[campo])
        a["dias_integr"] = numero(b["dias_integr"])
        if a["id"] in excluir_ids or a["email"].lower() in excluir_emails:
            continue
        alunos.append(a)

    ref = data_de_corte(alunos)
    aprovadas = {}
    for d in desempenho:
        nota = numero(d["nota"])
        if nota is not None and nota >= nota_min:
            chave = str(int(d["id"])) if isinstance(d["id"], float) and d["id"].is_integer() else str(d["id"] or "").strip()
            aprovadas.setdefault(chave, set()).add(d["seq"])

    partes_expert = [sem_acento(p) for p in str(pos.get("expert") or cfg.get("expert") or "").split() if len(p) > 2]
    dominios = [d.lower() for d in pos.get("dominios_internos", [])]
    ativos, cancelados, suspeitas = [], [], []
    for a in alunos:
        a["aprovadas"] = len(aprovadas.get(a["id"], ()))
        a["dias_curso"] = (ref - a["inicio"]).days if ref and a["inicio"] else None
        a["dias_sem_acesso"] = (ref - a["ultimo"]).days if ref and a["ultimo"] else None
        esperado = min(max((a["dias_curso"] or 0) / passo, 0), n_disc)
        a["atraso"] = esperado - a["aprovadas"]
        if partes_expert and all(p in sem_acento(a["nome"]) for p in partes_expert):
            suspeitas.append(a)
        elif any(a["email"].lower().endswith("@" + d) for d in dominios):
            suspeitas.append(a)
        if a["status"].lower() == "cancelada":
            cancelados.append(a)
            continue
        a["segmento"] = segmento(a, n_disc)
        a["cor"] = cor_pos(a["segmento"], a["aprovadas"])
        a["alerta"] = "matrícula pré-confirmada" if sem_acento(a["status"]).startswith("pre-confirmada") else ""
        ativos.append(a)
    return {"data_corte": ref, "alunos": ativos, "cancelados": cancelados, "suspeitas": suspeitas}


def whatsapp(celular):
    dig = re.sub(r"\D", "", str(celular or ""))
    if len(dig) in (10, 11):
        dig = "55" + dig
    return "https://wa.me/" + dig if len(dig) >= 12 else ""
