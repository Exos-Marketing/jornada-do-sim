#!/usr/bin/env python3
"""jornada-do-sim: revisão semanal do aluno depois da compra (skill do Claude Code).

Cruza as compradoras do produto (TMB e Voomp) com a planilha de engajamento da plataforma do curso, põe cada
aluno numa fase e numa cor do semáforo e entrega a lista de ação, as mensagens e o resumo do canal.

Subcomandos:
  listar                                    projetos configurados
  coletar-tmb --config CFG                  compradoras da TMB pela API (token em TMB_API_TOKEN)
  coletor-voomp --config CFG                imprime o coletor da Voomp para o browser_evaluate
  anhanguera --config CFG [--arquivo X]     guarda a planilha mais recente em leituras/
  revisar --config CFG [--hoje AAAA-MM-DD] [--sem-gravar]
  publicar --config CFG [--compartilhar]    sobe a lista da última revisão para o Google Sheets
"""
import argparse
import csv
import datetime as dt
import glob
import json
import os
import pathlib
import shutil
import sys

AQUI = pathlib.Path(__file__).resolve().parent
sys.path.insert(0, str(AQUI))

import anhanguera  # noqa: E402
import compradoras as cmp  # noqa: E402

EXOS = pathlib.Path.home() / "Documents" / "Exos"
CORES = ["verde", "amarelo", "laranja", "vermelho"]  # da melhor para a pior
RANK = {c: i for i, c in enumerate(CORES)}
GRUPOS = [
    (1, "chegada", "Chegada sem acesso"),
    (2, "nunca", "Nunca entrou"),
    (3, "mudou", "Mudou de cor"),
    (4, "vermelho", "Vermelhos · resgate"),
    (5, "laranja", "Laranjas · esfriando"),
    (6, "amarelo", "Amarelos · criar hábito"),
    (7, "verde", "Verdes · reconhecer"),
]
GRUPO_NOME = {g: n for _, g, n in GRUPOS}
GRUPO_ORDEM = {g: o for o, g, _ in GRUPOS}
MENSAGEM_DO_GRUPO = {"chegada": "chegada", "nunca": "nunca_entrou", "mudou": "mudou_de_cor",
                     "vermelho": "resgate", "laranja": "estimulo", "amarelo": "habito", "verde": "reconhecimento"}


# ------------------------------------------------------------------------------------------------- análise

def _rank(cor):
    return RANK.get(cor, -1)


def analisar(cfg, base, pessoas, estado_anterior, hoje):
    """Junta base (ler_planilha ou None) e compradoras (cmp.pessoas) numa lista priorizada e nas métricas.

    estado_anterior: {chave: cor} da revisão anterior. hoje: date da revisão.
    """
    prazo = int(cfg.get("prazo_acesso_dias", 3))
    janela = int(cfg.get("janela_chegada_dias", 30))
    data_corte = base["data_corte"] if base else None
    alunos = base["alunos"] if base else []

    por_email, por_fone = {}, {}
    for a in alunos:
        a["h_email"] = cmp.hash_id(a["email"])
        a["fone"] = cmp.fone(a["celular"])
        if a["h_email"]:
            por_email.setdefault(a["h_email"], a)
        if a["fone"]:
            por_fone.setdefault(a["fone"], a)

    linhas, usados = [], set()
    chegada = {"compraram": 0, "entrou": 0, "nao_entrou": 0, "contrato_pendente": 0, "ainda_sem_dado": 0,
               "sem_matricula": 0, "outro_email": 0}
    for p in pessoas:
        d = (hoje - p["pago_em"]).days
        if d < 0 or d > janela:
            continue
        chegada["compraram"] += 1
        aluno, alerta = por_email.get(p["h_email"]), ""
        if aluno is None and p.get("fone"):
            aluno = por_fone.get(p["fone"])
            if aluno is not None:
                alerta = "provavelmente estudando com outro e-mail"
                chegada["outro_email"] += 1
        if aluno is not None:
            usados.add(id(aluno))
            if aluno["segmento"] not in ("nunca", "contrato"):
                chegada["entrou"] += 1
                continue  # entrou: aparece com a cor dela no semáforo
            acesso = "contrato pendente" if aluno["segmento"] == "contrato" else "não entrou"
            chegada["contrato_pendente" if aluno["segmento"] == "contrato" else "nao_entrou"] += 1
        elif data_corte is None or data_corte < p["pago_em"] + dt.timedelta(days=prazo):
            acesso = "ainda sem dado"
            chegada["ainda_sem_dado"] += 1
        else:
            acesso = "sem matrícula"
            chegada["sem_matricula"] += 1
        if p.get("status_financeiro") and str(p["status_financeiro"]).lower() not in ("em dia", "adimplente", "quitado"):
            alerta = "; ".join(x for x in (alerta, "financeiro: %s" % p["status_financeiro"]) if x)
        linhas.append(_linha("chegada", p, aluno, d, acesso, None, estado_anterior, alerta, cfg))

    contagem = {c: 0 for c in CORES + ["nao_entrou"]}
    mudaram = 0
    for a in alunos:
        if id(a) in usados and a["segmento"] in ("nunca", "contrato"):
            continue  # já está na chegada
        contagem[a["cor"]] += 1
        chave = "h:" + a["h_email"] if a["h_email"] else "m:" + a["id"]
        anterior = estado_anterior.get(chave)
        if a["cor"] == "nao_entrou":
            grupo = "nunca"
        elif anterior in RANK and _rank(a["cor"]) > _rank(anterior):
            grupo = "mudou"
            mudaram += 1
        else:
            grupo = a["cor"]
        d = (hoje - a["inicio"]).days if a["inicio"] else None
        # contrato pendente tem cor "não entrou" (não estuda), mas fez login: a coluna Acesso diz o motivo
        acesso = "contrato pendente" if a["segmento"] == "contrato" else "não entrou" if a["cor"] == "nao_entrou" else "entrou"
        linhas.append(_linha(grupo, None, a, d, acesso, anterior, estado_anterior, a.get("alerta", ""), cfg))

    linhas.sort(key=lambda l: (GRUPO_ORDEM[l["grupo"]], -(l["fase_d"] or 0), l["nome"] or ""))
    metricas = {
        "projeto": cfg.get("projeto"), "data_revisao": hoje.isoformat(),
        "data_corte": data_corte.isoformat() if data_corte else None,
        "dias_desde_corte": (hoje - data_corte).days if data_corte else None,
        "planilha_velha": data_corte is None or (hoje - data_corte).days > int(cfg.get("planilha_velha_dias", 8)),
        "chegada": chegada, "base": {"ativos": len(alunos), "por_cor": contagem, "mudaram_de_cor": mudaram},
        "grupos": {g: sum(1 for l in linhas if l["grupo"] == g) for _, g, _ in GRUPOS},
        "coortes": _coortes(cfg, pessoas, hoje),
    }
    return {"linhas": linhas, "metricas": metricas}


def _linha(grupo, pessoa, aluno, fase_d, acesso, anterior, estado_anterior, alerta, cfg):
    nome = (aluno or {}).get("nome") or (pessoa or {}).get("nome")
    email = (aluno or {}).get("email") or (pessoa or {}).get("email")
    celular = (aluno or {}).get("celular") or (pessoa or {}).get("telefone")
    if pessoa is not None:
        chave = pessoa["chave"]
        anterior = estado_anterior.get(chave)
    else:
        chave = "h:" + aluno["h_email"] if aluno.get("h_email") else "m:" + aluno["id"]
    return {
        "grupo": grupo, "prioridade": GRUPO_ORDEM[grupo], "chave": chave, "nome": nome, "email": email,
        "celular": celular, "whatsapp": anhanguera.whatsapp(celular), "matricula": (aluno or {}).get("id"),
        "fase_d": fase_d, "acesso": acesso, "cor": (aluno or {}).get("cor") if aluno else None,
        "cor_anterior": anterior, "alerta": alerta, "aprovadas": (aluno or {}).get("aprovadas"),
        "ultimo": (aluno or {}).get("ultimo"), "pago_em": (pessoa or {}).get("pago_em"),
        "plataformas": ", ".join((pessoa or {}).get("plataformas") or []),
        "mensagem": MENSAGEM_DO_GRUPO[grupo],
        "sem_contato": pessoa is not None and aluno is None and not (pessoa.get("telefone") or pessoa.get("email")),
    }


def _limite(texto, fim=False):
    """Limite da turma: 'AAAA-MM-DD' (dia inteiro) ou 'AAAA-MM-DDTHH:MM', no horário de Brasília."""
    if "T" in texto:
        return dt.datetime.fromisoformat(texto)
    d = dt.datetime.fromisoformat(texto)
    return d.replace(hour=23, minute=59, second=59) if fim else d


def _coortes(cfg, pessoas, hoje):
    """Quantas pessoas compraram em cada turma, pela data da compra (no boleto, o pagamento pode vir depois)."""
    saida = []
    for c in cfg.get("coortes", []):
        inicio, fim = _limite(c["inicio"]), _limite(c["fim"], fim=True)
        n = sum(1 for p in pessoas if inicio <= (p.get("comprou_quando") or dt.datetime.combine(p["pago_em"], dt.time())) <= fim)
        saida.append({"id": c["id"], "nome": c["nome"], "inicio": c["inicio"], "fim": c["fim"], "compraram": n,
                      "d_hoje": "D%d a D%d" % ((hoje - fim.date()).days, (hoje - inicio.date()).days)})
    return saida


# ------------------------------------------------------------------------------------------------- arquivos

def pasta_do(cfg_path):
    return pathlib.Path(cfg_path).resolve().parent


def ler_config(cfg_path):
    with open(cfg_path, encoding="utf-8") as f:
        return json.load(f)


def gravar_config(cfg_path, cfg):
    with open(cfg_path, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)
        f.write("\n")


def ultima(pasta, padrao):
    """Arquivo mais recente (pela ordem do nome) que casa com o padrão. A pasta tem colchetes ([LL][POS]): escapar."""
    achados = sorted(glob.glob(os.path.join(glob.escape(str(pasta)), padrao)))
    return pathlib.Path(achados[-1]) if achados else None


def ler_json(arq):
    """Leitura gravada pelo browser_evaluate: JSON puro, JSON dentro de texto ou JSON codificado duas vezes."""
    bruto = pathlib.Path(arq).read_text(encoding="utf-8")
    try:
        dados = json.loads(bruto)
    except ValueError:
        dados = json.loads(bruto[bruto.find("{"):bruto.rfind("}") + 1])
    return json.loads(dados) if isinstance(dados, str) else dados


def ler_estado(pasta, hoje):
    """{chave: cor} da revisão mais recente ANTERIOR a hoje. Rodar de novo no mesmo dia não apaga a comparação."""
    caminho = pathlib.Path(pasta) / "estado.csv"
    if not caminho.exists():
        return {}
    ultimo = {}
    with open(caminho, encoding="utf-8") as f:
        for r in csv.DictReader(f):
            data = r.get("data_revisao") or ""
            if r.get("cor") and data < hoje.isoformat() and (r["chave"] not in ultimo or data >= ultimo[r["chave"]][0]):
                ultimo[r["chave"]] = (data, r["cor"])
    return {k: v[1] for k, v in ultimo.items()}


def carregar_compradoras(cfg, pasta):
    """Última leitura da TMB e da Voomp da jornada; sem Voomp da jornada, a última da acompanhar-vendas (sem contato)."""
    leituras, origens = [], {}
    for plataforma in ("tmb", "voomp"):
        arq = ultima(pasta / "leituras", "%s-*.json" % plataforma)
        if arq is None and plataforma == "voomp" and cfg.get("vendas", {}).get("leituras_acompanhar_vendas"):
            reserva = (pasta / cfg["vendas"]["leituras_acompanhar_vendas"]).resolve()
            arq = ultima(reserva, "voomp-*.json")
        if arq is None:
            continue
        dados = ler_json(arq)
        if dados.get("erro"):
            raise SystemExit("A leitura %s veio com erro (%s). Refaça a coleta." % (arq.name, dados["erro"]))
        produtos = {str(p) for p in cfg.get("vendas", {}).get(plataforma, {}).get("produtos", [])}
        if produtos:
            dados = dict(dados, linhas=[l for l in dados.get("linhas", []) if str(l.get("produto_id")) in produtos])
        normalizadas = cmp.normalizar(dict(dados, plataforma=plataforma))
        leituras += normalizadas
        origens[plataforma] = {"arquivo": arq.name, "coletado_em": dados.get("coletado_em"),
                               "pagas": sum(1 for l in normalizadas if l["situacao"] == cmp.PAGO),
                               "com_contato": sum(1 for l in normalizadas if l["situacao"] == cmp.PAGO and (l["telefone"] or l["email"]))}
    return cmp.pessoas(leituras), origens


def avisos_das_fontes(cfg, origens):
    """Fonte configurada que não veio, ou veio sem nenhuma venda paga, ou sem contato: avisar antes de entregar."""
    avisos = []
    for plataforma in ("tmb", "voomp"):
        if not cfg.get("vendas", {}).get(plataforma, {}).get("produtos"):
            continue
        o = origens.get(plataforma)
        if o is None:
            avisos.append("AVISO: nenhuma leitura da %s em leituras/." % plataforma.upper())
        elif o["pagas"] == 0:
            avisos.append("AVISO: a leitura da %s (%s) não tem nenhuma venda paga dos produtos do config. "
                          "Confira a conta logada e os ids de produto." % (plataforma.upper(), o["arquivo"]))
        elif o["com_contato"] == 0:
            avisos.append("AVISO: a leitura da %s (%s) não tem contato: quem não casar com a planilha sai sem WhatsApp."
                          % (plataforma.upper(), o["arquivo"]))
    vazios = [k for k in ("passo_a_passo", "primeiro_resultado", "rotina") if not cfg.get("mensagens", {}).get(k)]
    if vazios:
        avisos.append("AVISO: mensagens com lacuna entre colchetes (%s). O suporte completa antes de enviar."
                      % ", ".join(v.replace("_", " ") for v in vazios))
    if not cfg.get("mensagens", {}).get("aprovado_em"):
        avisos.append("AVISO: os textos das mensagens ainda não foram aprovados pelo Cleber (mensagens.aprovado_em).")
    return avisos


# ------------------------------------------------------------------------------------------------- comandos

def cmd_listar(_args):
    for cfg_path in sorted(glob.glob(str(EXOS / "*" / "08-Resultados-Relatorios" / "Jornada do Sim *" / "config.json"))):
        cfg = ler_config(cfg_path)
        print("%s · %s\n  %s" % (cfg.get("projeto"), cfg.get("titulo"), cfg_path))


def cmd_coletar_tmb(args):
    cfg, pasta = ler_config(args.config), pasta_do(args.config)
    token = os.environ.get("TMB_API_TOKEN")
    if not token:
        raise SystemExit("Falta TMB_API_TOKEN (Keychain: security find-generic-password -s tmb-exos-api -w).")
    leitura = cmp.coletar_tmb(cfg["vendas"]["tmb"]["produtos"], token)
    if leitura.get("erro"):
        raise SystemExit("TMB: %s" % leitura["erro"])
    destino = pasta / "leituras" / ("tmb-%s.json" % dt.datetime.now().strftime("%Y-%m-%d_%H%M"))
    destino.parent.mkdir(parents=True, exist_ok=True)
    destino.write_text(json.dumps(leitura, ensure_ascii=False, default=str), encoding="utf-8")
    print("TMB: %d pedidos em %s" % (len(leitura["linhas"]), destino))


def cmd_coletor_voomp(args):
    cfg = ler_config(args.config)
    js = (AQUI / "coletores" / "voomp_compradoras.js").read_text(encoding="utf-8")
    v = cfg["vendas"]["voomp"]
    print(js.replace("__CFG__", json.dumps({"desde": v.get("desde", "2026-01-01"), "produtos": [int(p) for p in v["produtos"]]})))


def cmd_anhanguera(args):
    cfg, pasta = ler_config(args.config), pasta_do(args.config)
    origem = pathlib.Path(args.arquivo) if args.arquivo else None
    if origem is None:
        padrao = cfg["anhanguera"].get("arquivo_padrao", "*.xlsx")
        achados = sorted(glob.glob(str(pathlib.Path.home() / "Downloads" / padrao)), key=os.path.getmtime)
        if not achados:
            raise SystemExit("Nenhuma planilha %s em Downloads." % padrao)
        origem = pathlib.Path(achados[-1])
    base = anhanguera.ler_planilha(str(origem), cfg)
    destino = pasta / "leituras" / ("anhanguera-%s.xlsx" % base["data_corte"].isoformat())
    destino.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(origem, destino)
    print("Planilha %s · data de corte %s · %d ativos · guardada em %s"
          % (origem.name, base["data_corte"].strftime("%d/%m/%Y"), len(base["alunos"]), destino.name))
    if base["suspeitas"]:
        print("Atenção: %d possível(is) conta(s) interna(s) fora da exclusão (matrículas %s)."
              % (len(base["suspeitas"]), ", ".join(a["id"] for a in base["suspeitas"])))


def cmd_revisar(args):
    import saidas

    cfg, pasta = ler_config(args.config), pasta_do(args.config)
    hoje = dt.date.fromisoformat(args.hoje) if args.hoje else dt.date.today()
    planilha = ultima(pasta / "leituras", "anhanguera-*.xlsx")
    base = anhanguera.ler_planilha(str(planilha), cfg) if planilha else None
    pessoas, origens = carregar_compradoras(cfg, pasta)
    resultado = analisar(cfg, base, pessoas, ler_estado(pasta, hoje), hoje)
    for aviso in avisos_das_fontes(cfg, origens):
        print(aviso)
    resultado["metricas"]["fontes"] = dict(origens, anhanguera=planilha.name if planilha else None)
    anterior = saidas.ultima_metrica(pasta, hoje)
    destino = pasta / "saidas" / hoje.isoformat()
    planilha_id = cfg.get("entrega", {}).get("drive_planilha_id")
    link = "https://docs.google.com/spreadsheets/d/%s" % planilha_id if planilha_id else None
    if not args.sem_gravar:
        destino.mkdir(parents=True, exist_ok=True)
        lista = saidas.gravar_lista(resultado, cfg, destino)
        mensagens = saidas.gravar_mensagens(cfg, destino)
        (destino / "metricas.json").write_text(json.dumps(resultado["metricas"], ensure_ascii=False, indent=2), encoding="utf-8")
        canal = saidas.resumo_canal(resultado["metricas"], cfg, anterior, link)
        (destino / "canal.md").write_text(canal, encoding="utf-8")
        saidas.atualizar_registros(resultado, pasta, hoje)
        print(saidas.relatorio_chat(resultado["metricas"], cfg, anterior))
        print("\nLista: %s\nMensagens: %s" % (lista, mensagens))
        print("\n----- RESUMO PARA O CANAL -----\n" + canal)
    else:
        print(saidas.relatorio_chat(resultado["metricas"], cfg, anterior))


def cmd_publicar(args):
    import drive

    cfg, pasta = ler_config(args.config), pasta_do(args.config)
    lista = ultima(pasta / "saidas", "*/Lista de ação*.xlsx")
    if lista is None:
        raise SystemExit("Nenhuma lista em saidas/: rode o revisar antes.")
    import saidas

    info = drive.publicar(lista, cfg, compartilhar=args.compartilhar)
    if info.get("novo_id"):
        cfg.setdefault("entrega", {})["drive_planilha_id"] = info["id"]
        gravar_config(args.config, cfg)
    print("Planilha: %s" % info["link"])
    for linha in info.get("avisos", []):
        print(linha)
    if any(l.startswith("ATENÇÃO") for l in info.get("avisos", [])):
        print("PARE: não poste o link no canal antes de tirar o acesso aberto (peça o OK do usuário).")
        return
    # o bloco do canal sai daqui, já com o link (na primeira publicação o id ainda não existia no revisar)
    pasta_saida = lista.parent
    with open(pasta_saida / "metricas.json", encoding="utf-8") as f:
        metricas = json.load(f)
    anterior = saidas.ultima_metrica(pasta, dt.date.fromisoformat(metricas["data_revisao"]))
    canal = saidas.resumo_canal(metricas, cfg, anterior, info["link"])
    (pasta_saida / "canal.md").write_text(canal, encoding="utf-8")
    aprovado = cfg.get("entrega", {}).get("canal_aprovado")
    print("\nPOSTAR NO CANAL: %s" % ("SIM" if aprovado else "SÓ COM O OK DO CLEBER (primeira revisão deste projeto)"))
    print("----- RESUMO PARA O CANAL -----\n" + canal)


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="cmd", required=True)
    sub.add_parser("listar").set_defaults(f=cmd_listar)
    for nome, f in (("coletar-tmb", cmd_coletar_tmb), ("coletor-voomp", cmd_coletor_voomp)):
        p = sub.add_parser(nome)
        p.add_argument("--config", required=True)
        p.set_defaults(f=f)
    p = sub.add_parser("anhanguera")
    p.add_argument("--config", required=True)
    p.add_argument("--arquivo")
    p.set_defaults(f=cmd_anhanguera)
    p = sub.add_parser("revisar")
    p.add_argument("--config", required=True)
    p.add_argument("--hoje")
    p.add_argument("--sem-gravar", action="store_true")
    p.set_defaults(f=cmd_revisar)
    p = sub.add_parser("publicar")
    p.add_argument("--config", required=True)
    p.add_argument("--compartilhar", action="store_true")
    p.set_defaults(f=cmd_publicar)
    args = ap.parse_args(argv)
    args.f(args)


if __name__ == "__main__":
    main()
