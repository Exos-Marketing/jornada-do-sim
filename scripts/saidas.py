"""Saídas da revisão: lista de ação (xlsx), mensagens (md), resumo do canal, relatório do chat e registros."""
import csv
import datetime as dt
import json
import pathlib
import re

AQUI = pathlib.Path(__file__).resolve().parent
MODELOS = AQUI.parent / "modelos" / "mensagens.json"

COLUNAS = [
    ("Nome", 30), ("WhatsApp", 16), ("E-mail", 30), ("Matrícula", 11), ("Fase", 7), ("Acesso", 15),
    ("Cor", 11), ("Cor anterior", 12), ("Alerta", 28), ("Disciplinas", 11), ("Último acesso", 13),
    ("Compra", 11), ("Mensagem sugerida", 60), ("Responsável", 16), ("Contato em", 12), ("Resultado", 16),
    ("Anotações", 40), ("chave", 14),
]
COLUNAS_DO_TIME = ["Responsável", "Contato em", "Resultado", "Anotações"]
RESULTADOS = ["entrou", "vai entrar", "respondeu", "sem resposta", "problema", "não quer seguir"]
NOME_COR = {"verde": "Verde", "amarelo": "Amarelo", "laranja": "Laranja", "vermelho": "Vermelho",
            "nao_entrou": "Não entrou", None: ""}
FUNDO_COR = {"verde": "D7F2D7", "amarelo": "FDF1C7", "laranja": "FBDDC6", "vermelho": "F7D0D6", "nao_entrou": "E3E7EC"}
# Explicação de cada aba, na aba Resumo: (grupo, quem está nela, o que fazer, quando). {janela} = janela_chegada_dias.
EXPLICACAO_DAS_ABAS = [
    (None, "Os números da semana e esta explicação de cada aba.", "Ler antes de começar a semana.", "Segunda"),
    ("chegada",
     "Comprou nos últimos {janela} dias (conta a data da compra) e ainda não aparece estudando na planilha da Anhanguera. "
     "A coluna Acesso diz o motivo: não entrou; contrato pendente (fez login, mas não aceitou o contrato); ainda sem dado "
     "(comprou depois da data da planilha); sem matrícula (pagou e não aparece na Anhanguera).",
     "Ligar seguindo o roteiro: ajudar a aceitar o contrato e a fazer o primeiro acesso e combinar o primeiro resultado "
     "da semana. Não atendeu: mandar a mensagem sugerida no WhatsApp e ligar de novo na quinta. Sem matrícula: avisar a "
     "coordenação da Pós.", "Até quarta"),
    ("nunca",
     "Comprou há mais de {janela} dias e ainda não começou. Último acesso vazio: nunca entrou no portal. Último acesso com "
     "data: entrou, mas está com o contrato pendente (a Pós só abre depois que ela aceita o contrato).",
     "Mandar a mensagem sugerida, convidando para o primeiro passo. Com contrato pendente, ajudar a aceitar o contrato.",
     "Quinta"),
    ("mudou",
     "Piorou de cor desde a revisão da semana anterior (por exemplo, de verde para laranja). A coluna Cor anterior mostra "
     "de onde veio. Quem mudou de cor fica só nesta aba, não na aba da cor nova.",
     "Mandar mensagem na mesma semana: perguntar se está tudo bem e oferecer ajuda para retomar o ritmo.", "Quinta"),
    ("vermelho", "Está há 30 dias ou mais sem acessar e não concluiu nenhuma disciplina.",
     "Contato um a um, de preferência por ligação: entender o que aconteceu e montar um plano de retomada.", "Sexta"),
    ("laranja", "Está há 30 dias ou mais sem acessar, mas já concluiu pelo menos uma disciplina.",
     "Trazer de volta antes que esfrie de vez: lembrar do que ela já conquistou e sugerir por onde retomar.", "Sexta"),
    ("amarelo",
     "Está estudando, mas sem ritmo: tem menos de 45 dias de curso, não concluiu nenhuma disciplina depois de 45 dias ou "
     "está 2 ou mais disciplinas atrás do ritmo esperado.",
     "Ajudar a criar a rotina: sugerir um horário fixo na semana e uma meta curta.", "Sexta"),
    ("verde", "Está no ritmo ou adiantada (ou já concluiu 5 ou mais disciplinas).",
     "Reconhecer o avanço e convidar para gravar um depoimento curto.", "Sexta"),
]


class _Faltando(dict):
    def __missing__(self, chave):
        return "[%s]" % chave.replace("_", " ")


def modelos():
    with open(MODELOS, encoding="utf-8") as f:
        dados = json.load(f)
    return {k: v for k, v in dados.items() if not k.startswith("_")}


def campos(cfg, nome=None):
    m = dict(cfg.get("mensagens", {}))
    m.setdefault("expert_curto", (cfg.get("expert") or "").split()[0] if cfg.get("expert") else "[expert]")
    m.setdefault("produto_curto", cfg.get("produto_curto") or cfg.get("produto") or "[produto]")
    m["primeiro_nome"] = (str(nome).strip().split()[0].title() if nome and str(nome).strip() else "[nome]")
    return _Faltando({k: v for k, v in m.items() if v not in (None, "")})


def preencher(texto, cfg, nome=None):
    return texto.format_map(campos(cfg, nome)).replace("  ", " ").strip()


def _d(valor):
    if isinstance(valor, dt.date):
        return valor.strftime("%d/%m/%Y")
    return valor if valor not in (None, "") else ""


def nome_lista(cfg, hoje):
    return "Lista de ação %s %s.xlsx" % (cfg.get("projeto", ""), hoje)


def gravar_lista(resultado, cfg, destino):
    import openpyxl
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.worksheet.datavalidation import DataValidation

    from jornada import GRUPOS, GRUPO_NOME

    m = resultado["metricas"]
    textos = modelos()
    livro = openpyxl.Workbook()
    resumo = livro.active
    resumo.title = "Resumo"
    _resumo(resumo, m, cfg)

    cabecalho = PatternFill("solid", fgColor="0B1828")
    for _, grupo, _ in GRUPOS:
        linhas = [l for l in resultado["linhas"] if l["grupo"] == grupo]
        if not linhas:
            continue
        aba = livro.create_sheet(re.sub(r"[\\/*?:\[\]]", "", GRUPO_NOME[grupo])[:31])
        aba.append([c for c, _ in COLUNAS])
        for i, (_, largura) in enumerate(COLUNAS, start=1):
            aba.column_dimensions[openpyxl.utils.get_column_letter(i)].width = largura
            celula = aba.cell(row=1, column=i)
            celula.font = Font(bold=True, color="FFFFFF")
            celula.fill = cabecalho
            celula.alignment = Alignment(vertical="center", wrap_text=True)
        for l in linhas:
            texto = textos[l["mensagem"]]["mensagem"]
            aba.append([
                l["nome"] or "", "Abrir WhatsApp" if l["whatsapp"] else ("sem contato: ver na plataforma" if l.get("sem_contato") else ""),
                l["email"] or "", l["matricula"] or "", "D%d" % l["fase_d"] if l["fase_d"] is not None else "",
                l["acesso"], NOME_COR.get(l["cor"], l["cor"] or ""), NOME_COR.get(l["cor_anterior"], l["cor_anterior"] or ""),
                l["alerta"] or "", l["aprovadas"] if l["aprovadas"] is not None else "", _d(l["ultimo"]),
                _d(l["pago_em"]), preencher(texto, cfg, l["nome"]), "", "", "", "", l["chave"],
            ])
            r = aba.max_row
            if l["whatsapp"]:
                aba.cell(row=r, column=2).hyperlink = l["whatsapp"]
                aba.cell(row=r, column=2).font = Font(color="1161BB", underline="single")
            if l["cor"] in FUNDO_COR:
                aba.cell(row=r, column=7).fill = PatternFill("solid", fgColor=FUNDO_COR[l["cor"]])
            aba.cell(row=r, column=13).alignment = Alignment(wrap_text=True, vertical="top")
        aba.freeze_panes = "B2"
        aba.auto_filter.ref = aba.dimensions
        aba.column_dimensions["R"].hidden = True
        validacao = DataValidation(type="list", formula1='"%s"' % ",".join(RESULTADOS), allow_blank=True)
        aba.add_data_validation(validacao)
        validacao.add("P2:P%d" % max(aba.max_row, 2))
    caminho = pathlib.Path(destino) / nome_lista(cfg, m["data_revisao"])
    livro.save(caminho)
    return caminho


def _resumo(aba, m, cfg):
    from openpyxl.styles import Font

    ch, base = m["chegada"], m["base"]
    linhas = [
        ("Jornada do Sim · lista de ação", ""),
        ("Produto", cfg.get("titulo") or cfg.get("produto")),
        ("Revisão", _d(dt.date.fromisoformat(m["data_revisao"]))),
        ("Planilha da plataforma", "base de %s (%s dias)" % (_d(dt.date.fromisoformat(m["data_corte"])), m["dias_desde_corte"])
         if m["data_corte"] else "sem planilha"),
        ("Atenção", "planilha velha: o acesso de quem comprou depois da data da base ainda não aparece; confirme na ligação"
         if m["planilha_velha"] else ""),
        ("", ""),
        ("Chegada (comprou nos últimos %s dias)" % cfg.get("janela_chegada_dias", 30), ch["compraram"]),
        ("  entraram", ch["entrou"]), ("  não entraram", ch["nao_entrou"]), ("  contrato pendente", ch["contrato_pendente"]),
        ("  ainda sem dado", ch["ainda_sem_dado"]), ("  sem matrícula", ch["sem_matricula"]),
        ("  provavelmente com outro e-mail", ch["outro_email"]),
        ("", ""),
        ("Base ativa", base["ativos"]),
        ("  verdes", base["por_cor"]["verde"]), ("  amarelos", base["por_cor"]["amarelo"]),
        ("  laranjas", base["por_cor"]["laranja"]), ("  vermelhos", base["por_cor"]["vermelho"]),
        ("  nunca entraram", base["por_cor"]["nao_entrou"]), ("  mudaram de cor desde a semana anterior", base["mudaram_de_cor"]),
        ("", ""),
        ("Como usar", "uma aba por grupo, na ordem de prioridade. Preencha Responsável, Contato em, Resultado e Anotações: "
                      "elas continuam na planilha na próxima segunda."),
        ("Dado pessoal", "uso interno do suporte. Não baixe, não encaminhe e não compartilhe esta planilha."),
    ]
    for a, b in linhas:
        aba.append([a, b])
    aba["A1"].font = Font(bold=True, size=14)
    for celula in ("A5", "B5"):
        aba[celula].font = Font(bold=True, color="B8283C")
    _explicacao_das_abas(aba, m, cfg)
    aba.column_dimensions["A"].width = 42
    aba.column_dimensions["B"].width = 90
    aba.column_dimensions["C"].width = 60
    aba.column_dimensions["D"].width = 12
    aba.column_dimensions["E"].width = 13


def _explicacao_das_abas(aba, m, cfg):
    from openpyxl.styles import Alignment, Font, PatternFill

    from jornada import GRUPO_NOME, GRUPO_ORDEM

    janela = cfg.get("janela_chegada_dias", 30)
    aba.append([])
    aba.append(["O que tem em cada aba"])
    aba.cell(row=aba.max_row, column=1).font = Font(bold=True, size=13)
    aba.append(["Aba", "Quem está nela", "O que fazer", "Quando", "Nesta semana"])
    for celula in aba[aba.max_row]:
        celula.font = Font(bold=True, color="FFFFFF")
        celula.fill = PatternFill("solid", fgColor="0B1828")
    for grupo, quem, fazer, quando in EXPLICACAO_DAS_ABAS:
        if grupo is None:
            nome, n = "Resumo", ""
        else:
            nome = "%d · %s" % (GRUPO_ORDEM[grupo], GRUPO_NOME[grupo])
            n = m.get("grupos", {}).get(grupo, 0)
            n = n if n else "0 (sem aba nesta semana)"
        aba.append([nome, quem.format(janela=janela), fazer, quando, n])
        for celula in aba[aba.max_row]:
            celula.alignment = Alignment(wrap_text=True, vertical="top")
        aba.cell(row=aba.max_row, column=1).font = Font(bold=True)


def gravar_mensagens(cfg, destino):
    partes = ["# Mensagens e roteiros · %s" % (cfg.get("titulo") or cfg.get("produto")), "",
              "Na voz de %s. O que está entre colchetes falta no config do projeto: completar antes de enviar." % cfg.get("expert", "[expert]"), ""]
    for chave, t in modelos().items():
        partes += ["## %s" % t["titulo"], "", "*Quando:* %s" % t["quando"], "", "> " + preencher(t["mensagem"], cfg).replace("\n", "\n> "), ""]
        if t.get("roteiro_ligacao"):
            partes += ["**Roteiro da ligação**", ""] + ["%d. %s" % (i, preencher(p, cfg)) for i, p in enumerate(t["roteiro_ligacao"], 1)] + [""]
    caminho = pathlib.Path(destino) / "mensagens.md"
    caminho.write_text("\n".join(partes), encoding="utf-8")
    return caminho


def _var(atual, anterior, *chaves):
    if not anterior:
        return ""
    a, b = atual, anterior
    for c in chaves:
        a, b = a.get(c, {}) if isinstance(a, dict) else a, b.get(c, {}) if isinstance(b, dict) else b
    if not isinstance(a, int) or not isinstance(b, int) or a == b:
        return ""
    return " (%+d)" % (a - b)


def _marcacoes(cfg):
    """Quem faz o contato, marcado no ClickUp ([@Nome](#user_mention#id)); vazio se o projeto não marca ninguém."""
    nomes = ["[@%s](#user_mention#%s)" % (p["nome"], p["id"]) for p in cfg.get("contato", {}).get("clickup_marcar", [])]
    return " e ".join([", ".join(nomes[:-1]), nomes[-1]]) if len(nomes) > 1 else "".join(nomes)


def resumo_canal(m, cfg, anterior=None, link=None):
    """Bloco para o canal do ClickUp: só números, sem nome, e-mail ou telefone de aluno. Na linha do link, marca quem
    faz o contato: é assim que o suporte fica sabendo que a lista saiu (o Drive não manda e-mail)."""
    ch, base = m["chegada"], m["base"]
    rev = dt.date.fromisoformat(m["data_revisao"]).strftime("%d/%m")
    linhas = ["**Jornada do Sim · %s · revisão de %s**" % (cfg.get("titulo") or cfg.get("produto"), rev)]
    if m["data_corte"]:
        corte = dt.date.fromisoformat(m["data_corte"]).strftime("%d/%m")
        aviso = " ⚠️ planilha velha: o acesso de quem comprou depois disso ainda não aparece" if m["planilha_velha"] else ""
        linhas.append("Planilha da plataforma: base de %s.%s" % (corte, aviso))
    else:
        linhas.append("Planilha da plataforma: não chegou. A lista traz só a chegada pelas vendas.")
    linhas.append("**Chegada** (comprou nos últimos %s dias): %d · entraram %d · não entraram %d · contrato pendente %d · ainda sem dado %d · sem matrícula %d"
                  % (cfg.get("janela_chegada_dias", 30), ch["compraram"], ch["entrou"], ch["nao_entrou"], ch["contrato_pendente"],
                     ch["ainda_sem_dado"], ch["sem_matricula"]))
    pc = base["por_cor"]
    linhas.append("**Base** (%d ativos): 🟢 %d%s · 🟡 %d%s · 🟠 %d%s · 🔴 %d%s · nunca entraram %d · mudaram de cor %d"
                  % (base["ativos"], pc["verde"], _var(m, anterior, "base", "por_cor", "verde"),
                     pc["amarelo"], _var(m, anterior, "base", "por_cor", "amarelo"),
                     pc["laranja"], _var(m, anterior, "base", "por_cor", "laranja"),
                     pc["vermelho"], _var(m, anterior, "base", "por_cor", "vermelho"),
                     pc["nao_entrou"], base["mudaram_de_cor"]))
    for c in m.get("coortes", []):
        linhas.append("Turma %s: %d compradoras (%s)" % (c["nome"], c["compraram"], c["d_hoje"]))
    linhas.append("**Semana do suporte:** ligar para a chegada até quarta; depois, mudou de cor, vermelhos e laranjas.")
    if link:
        marcar = _marcacoes(cfg)
        linhas.append("%s, a lista desta semana já está no Drive: %s" % (marcar, link) if marcar else "Lista da semana: %s" % link)
    return "\n".join(linhas) + "\n"


def relatorio_chat(m, cfg, anterior=None):
    ch, base = m["chegada"], m["base"]
    partes = ["Revisão de %s · %s" % (m["data_revisao"], cfg.get("titulo") or cfg.get("produto"))]
    if m["data_corte"]:
        partes.append("Planilha da plataforma: base de %s (%s dias)%s." % (
            m["data_corte"], m["dias_desde_corte"], " · VELHA" if m["planilha_velha"] else ""))
    else:
        partes.append("Planilha da plataforma: nenhuma em leituras/.")
    partes.append("Chegada: %(compraram)d compraram · %(entrou)d entraram · %(nao_entrou)d não entraram · "
                  "%(contrato_pendente)d contrato pendente · %(ainda_sem_dado)d ainda sem dado · %(sem_matricula)d sem matrícula · "
                  "%(outro_email)d provavelmente com outro e-mail" % ch)
    partes.append("Base: %d ativos · por cor %s · mudaram de cor %d" % (base["ativos"], json.dumps(base["por_cor"], ensure_ascii=False), base["mudaram_de_cor"]))
    partes.append("Grupos da lista: %s" % json.dumps(m["grupos"], ensure_ascii=False))
    if m.get("fontes"):
        partes.append("Fontes: %s" % json.dumps(m["fontes"], ensure_ascii=False))
    return "\n".join(partes)


def ultima_metrica(pasta, hoje):
    """Métricas da revisão anterior a hoje (para a variação da semana)."""
    anteriores = sorted(p for p in (pathlib.Path(pasta) / "saidas").glob("*/metricas.json") if p.parent.name < str(hoje))
    if not anteriores:
        return None
    with open(anteriores[-1], encoding="utf-8") as f:
        return json.load(f)


def atualizar_registros(resultado, pasta, hoje):
    pasta = pathlib.Path(pasta)
    m = resultado["metricas"]
    linha = {"data_revisao": m["data_revisao"], "data_corte": m["data_corte"] or "",
             "chegada_compraram": m["chegada"]["compraram"], "chegada_entrou": m["chegada"]["entrou"],
             "chegada_sem_acesso": m["grupos"]["chegada"], "ativos": m["base"]["ativos"],
             **{"cor_" + c: n for c, n in m["base"]["por_cor"].items()}, "mudaram_de_cor": m["base"]["mudaram_de_cor"]}
    historico = pasta / "historico.csv"
    linhas = []
    if historico.exists():
        with open(historico, encoding="utf-8") as f:
            linhas = [r for r in csv.DictReader(f) if r["data_revisao"] != m["data_revisao"]]
    linhas.append({k: str(v) for k, v in linha.items()})
    with open(historico, "w", encoding="utf-8", newline="") as f:
        w = csv.DictWriter(f, fieldnames=list(linha.keys()))
        w.writeheader()
        for r in sorted(linhas, key=lambda r: r["data_revisao"]):
            w.writerow({k: r.get(k, "") for k in linha.keys()})
    # estado.csv guarda o histórico de cores por revisão; refazer a revisão do mesmo dia substitui só aquele dia
    estado = pasta / "estado.csv"
    antigas = []
    if estado.exists():
        with open(estado, encoding="utf-8") as f:
            antigas = [r for r in csv.DictReader(f) if r.get("data_revisao") != m["data_revisao"]]
    novas = [{"chave": l["chave"], "cor": l["cor"], "data_revisao": m["data_revisao"]} for l in resultado["linhas"]
             if l["cor"] in ("verde", "amarelo", "laranja", "vermelho")]
    with open(estado, "w", encoding="utf-8", newline="") as f:
        w = csv.DictWriter(f, fieldnames=["chave", "cor", "data_revisao"])
        w.writeheader()
        for r in antigas + novas:
            w.writerow({k: r.get(k, "") for k in ("chave", "cor", "data_revisao")})
