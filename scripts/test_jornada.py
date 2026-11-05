"""Testes da jornada-do-sim. Rodar: cd scripts && python3 -m unittest test_jornada"""
import datetime as dt
import importlib.util
import pathlib
import tempfile
import unittest
from unittest import mock

import anhanguera
import compradoras as cmp
import drive
import jornada
import saidas

D = dt.date
CORTE = D(2026, 9, 8)
CFG = {
    "projeto": "[TT][POS]", "titulo": "Pós de Teste", "expert": "Fulana Teste", "produto": "Pós de Teste",
    "prazo_acesso_dias": 3, "janela_chegada_dias": 30,
    "anhanguera": {"n_disciplinas": 9, "duracao_dias": 364, "nota_minima": 70, "excluir_ids": [999]},
    "mensagens": {"passo_a_passo": "acesse portal.exemplo"},
}


def aluno(id_, nome, inicio, ultimo, aprovadas=0, contrato="Aceito", status="Ativa", email=None, celular=None):
    fim = inicio + dt.timedelta(days=364)
    return {"id": id_, "nome": nome, "email": email or "%s@teste.com" % nome.lower(), "celular": celular or "",
            "status": status, "contrato": contrato, "ultimo": ultimo, "inicio": inicio, "fim": fim,
            "dias_integr": (fim - CORTE).days, "aprovadas": aprovadas}


ALUNOS = [
    aluno(1, "Nunca", D(2026, 6, 1), None),
    aluno(2, "Contrato", D(2026, 6, 1), D(2026, 9, 1), contrato="Pendente"),
    aluno(3, "Esfriou", D(2026, 4, 1), D(2026, 7, 20), aprovadas=1),
    aluno(4, "Sumiu", D(2026, 4, 1), D(2026, 7, 1)),
    aluno(5, "Novata", D(2026, 8, 15), D(2026, 9, 5), celular="(11) 98888-7777"),
    aluno(6, "Travada", D(2026, 6, 15), D(2026, 9, 1)),
    aluno(7, "Atrasada", D(2026, 3, 23), D(2026, 9, 5), aprovadas=2),
    aluno(8, "Ritmo", D(2026, 5, 1), D(2026, 9, 6), aprovadas=3, status="Pré-confirmada"),
    aluno(9, "Destaque", D(2026, 5, 1), D(2026, 9, 6), aprovadas=5),
    aluno(10, "Cancelada", D(2026, 5, 1), D(2026, 9, 6), status="Cancelada"),
    aluno(999, "Teste Interna", D(2026, 5, 1), D(2026, 9, 6)),
]


def planilha(destino):
    import openpyxl

    livro = openpyxl.Workbook()
    base = livro.active
    base.title = "Base de Alunos"
    base.append(anhanguera.COLUNAS_BASE)
    for a in ALUNOS:
        base.append(["Pós de Teste", a["id"], a["nome"], a["email"], a["celular"], a["status"], a["contrato"],
                     a["ultimo"], a["inicio"], a["fim"], a["dias_integr"], "T1"])
    des = livro.create_sheet("Desempenho Acadêmico")
    des.append(anhanguera.COLUNAS_DESEMPENHO)
    for a in ALUNOS:
        for seq in range(1, a["aprovadas"] + 1):
            des.append([a["id"], a["nome"], seq, "Disciplina %d" % seq, 1, 100, "T1"])
    des.append([6, "Travada", 1, "Disciplina 1", 5, 60, "T1"])  # reprovada não conta
    caminho = pathlib.Path(destino) / "planilha.xlsx"
    livro.save(caminho)
    return caminho


def compradora(email, pago, situacao="pago", telefone="", plataforma="tmb", pedido=1):
    return {"plataforma": plataforma, "pedido": pedido, "produto_id": 1, "situacao": situacao,
            "pago_em": pago if situacao == "pago" else None, "status_financeiro": None, "nome": "Nome " + email,
            "email": email, "telefone": telefone, "h_email": cmp.hash_id(email), "h_doc": None, "fone": cmp.fone(telefone)}


class TestAnhanguera(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.base = anhanguera.ler_planilha(str(planilha(self.tmp.name)), CFG)
        self.por_nome = {a["nome"]: a for a in self.base["alunos"]}

    def tearDown(self):
        self.tmp.cleanup()

    def test_data_de_corte(self):
        self.assertEqual(self.base["data_corte"], CORTE)

    def test_segmento_e_cor_de_cada_aluno(self):
        esperado = {
            "Nunca": ("nunca", "nao_entrou"), "Contrato": ("contrato", "nao_entrou"),
            "Esfriou": ("adormecido", "laranja"), "Sumiu": ("adormecido", "vermelho"),
            "Novata": ("novato", "amarelo"), "Travada": ("travado", "amarelo"),
            "Atrasada": ("atrasado", "amarelo"), "Ritmo": ("emdia", "verde"), "Destaque": ("destaque", "verde"),
        }
        for nome, (seg, cor) in esperado.items():
            self.assertEqual((self.por_nome[nome]["segmento"], self.por_nome[nome]["cor"]), (seg, cor), nome)

    def test_exclusao_cancelada_e_alerta(self):
        self.assertNotIn("Teste Interna", self.por_nome)
        self.assertEqual([a["nome"] for a in self.base["cancelados"]], ["Cancelada"])
        self.assertEqual(self.por_nome["Ritmo"]["alerta"], "matrícula pré-confirmada")
        self.assertEqual(self.por_nome["Travada"]["aprovadas"], 0)


class TestCompradoras(unittest.TestCase):
    def test_hash_igual_ao_da_acompanhar_vendas(self):
        caminho = pathlib.Path.home() / ".claude/skills/acompanhar-vendas/scripts/vendas.py"
        if not caminho.exists():
            self.skipTest("acompanhar-vendas não instalada")
        spec = importlib.util.spec_from_file_location("vendas_ref", caminho)
        vendas = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(vendas)
        self.assertEqual(cmp.hash_id(" Maria@Teste.com "), vendas.hash_id("maria@teste.com"))

    def test_compra_em_dobro_vira_uma_pessoa(self):
        linhas = [compradora("a@x.com", D(2026, 10, 8), pedido=1), compradora("a@x.com", D(2026, 10, 8), pedido=2, plataforma="voomp")]
        p = cmp.pessoas(linhas)
        self.assertEqual(len(p), 1)
        self.assertEqual(p[0]["pedidos"], 2)
        self.assertEqual(p[0]["plataformas"], ["tmb", "voomp"])

    def test_reembolso_sai(self):
        linhas = [compradora("b@x.com", D(2026, 10, 8), situacao="cancelado"), compradora("c@x.com", D(2026, 10, 9))]
        self.assertEqual([p["email"] for p in cmp.pessoas(linhas)], ["c@x.com"])

    def test_voomp_em_utc_vira_horario_de_brasilia(self):
        leitura = {"plataforma": "voomp", "linhas": [{"id": 1, "produto_id": 1, "status": "paid", "h_email": "abc",
                                                       "criado": "2026-10-17T00:50:00.000000Z", "pago": "2026-10-17T01:11:30.000000Z"}]}
        linha = cmp.normalizar(leitura)[0]
        self.assertEqual(linha["pago_em"], D(2026, 10, 16))
        self.assertEqual(linha["comprou_quando"], dt.datetime(2026, 10, 16, 21, 50))

    def test_tmb_com_fuso_e_sem_fuso(self):
        self.assertEqual(cmp.quando("2026-10-22T09:41:55.214484-03:00"), dt.datetime(2026, 10, 22, 9, 41, 55, 214484))
        self.assertEqual(cmp.quando("2026-10-22T09:45:43.959419"), dt.datetime(2026, 10, 22, 9, 45, 43, 959419))

    def test_turma_conta_pela_data_da_compra(self):
        p = cmp.pessoas([dict(compradora("boleto@x.com", D(2026, 10, 23)), comprou_quando=dt.datetime(2026, 10, 15, 10, 0))])
        cfg = dict(CFG, coortes=[{"id": "b", "nome": "Black", "inicio": "2026-10-07T20:00", "fim": "2026-10-16"}])
        self.assertEqual(jornada._coortes(cfg, p, D(2026, 10, 23))[0]["compraram"], 1)

    def test_normalizar_voomp_reembolsada(self):
        leitura = {"plataforma": "voomp", "linhas": [{"id": 1, "produto_id": 1, "status": "paid", "pago": "2026-10-08 10:00:00",
                                                       "reembolsado": "2026-10-09 10:00:00", "h_email": "abc"}]}
        self.assertEqual(cmp.normalizar(leitura)[0]["situacao"], "cancelado")


class TestAnalisar(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.base = anhanguera.ler_planilha(str(planilha(self.tmp.name)), CFG)
        hoje = D(2026, 9, 10)
        pessoas = cmp.pessoas([
            compradora("novata@teste.com", D(2026, 8, 15)),                        # entrou (casa por e-mail)
            compradora("nova@x.com", D(2026, 9, 7)),                                # planilha anterior ao prazo
            compradora("outro@x.com", D(2026, 8, 20), telefone="+55 11 98888-7777"),  # casa pelo telefone
            compradora("sumida@x.com", D(2026, 8, 20)),                             # sem matrícula
            compradora("contrato@teste.com", D(2026, 8, 25)),                       # contrato pendente
            compradora("antiga@x.com", D(2026, 6, 1)),                              # fora da janela de chegada
        ])
        self.r = jornada.analisar(CFG, self.base, pessoas, {"h:" + cmp.hash_id("esfriou@teste.com"): "verde"}, hoje)
        self.por_email = {}
        for l in self.r["linhas"]:
            self.por_email.setdefault(l["email"], l)

    def tearDown(self):
        self.tmp.cleanup()

    def test_planilha_anterior_a_compra_nao_vira_nao_entrou(self):
        self.assertEqual(self.por_email["nova@x.com"]["acesso"], "ainda sem dado")

    def test_sem_matricula_quando_a_planilha_cobre_o_prazo(self):
        self.assertEqual(self.por_email["sumida@x.com"]["acesso"], "sem matrícula")

    def test_outro_email_pelo_telefone(self):
        novata = [l for l in self.r["linhas"] if l["nome"] == "Novata"]
        self.assertEqual(len(novata), 1)  # casou: aparece uma vez, com a cor dela
        self.assertEqual(self.r["metricas"]["chegada"]["outro_email"], 1)

    def test_contrato_pendente_fica_na_chegada(self):
        self.assertEqual(self.por_email["contrato@teste.com"]["acesso"], "contrato pendente")
        self.assertEqual(self.por_email["contrato@teste.com"]["grupo"], "chegada")

    def test_mudou_de_cor(self):
        self.assertEqual(self.por_email["esfriou@teste.com"]["grupo"], "mudou")

    def test_fora_da_janela_nao_entra_na_chegada(self):
        self.assertNotIn("antiga@x.com", self.por_email)

    def test_ordem_pela_prioridade(self):
        prioridades = [l["prioridade"] for l in self.r["linhas"]]
        self.assertEqual(prioridades, sorted(prioridades))
        self.assertEqual(self.r["linhas"][0]["grupo"], "chegada")

    def test_saidas_sem_dado_pessoal_no_canal(self):
        texto = saidas.resumo_canal(self.r["metricas"], CFG)
        for l in self.r["linhas"]:
            for campo in ("nome", "email"):
                if l[campo]:
                    self.assertNotIn(l[campo], texto)

    def test_contrato_pendente_da_base_diz_o_motivo_no_acesso(self):
        r = jornada.analisar(CFG, self.base, [], {}, D(2026, 9, 10))
        contrato = next(l for l in r["linhas"] if l["nome"] == "Contrato")
        nunca = next(l for l in r["linhas"] if l["nome"] == "Nunca")
        self.assertEqual((contrato["grupo"], contrato["acesso"]), ("nunca", "contrato pendente"))
        self.assertEqual((nunca["grupo"], nunca["acesso"]), ("nunca", "não entrou"))

    def test_resumo_explica_todas_as_abas(self):
        import openpyxl

        with tempfile.TemporaryDirectory() as tmp:
            livro = openpyxl.load_workbook(saidas.gravar_lista(self.r, CFG, tmp))
        textos = [str(c.value) for linha in livro["Resumo"].iter_rows() for c in linha if c.value is not None]
        for o, _, nome in jornada.GRUPOS:
            self.assertIn("%d · %s" % (o, nome), textos)
        self.assertIn("O que tem em cada aba", textos)
        self.assertFalse(any("{janela}" in t for t in textos))

    def test_canal_marca_quem_faz_o_contato_na_linha_do_link(self):
        m = self.r["metricas"]
        uma = dict(CFG, contato={"clickup_marcar": [{"id": "82103125", "nome": "Lindsy Wegner"}]})
        self.assertTrue(saidas.resumo_canal(m, uma, link="https://docs.google.com/spreadsheets/d/x").endswith(
            "[@Lindsy Wegner](#user_mention#82103125), a lista desta semana já está no Drive: "
            "https://docs.google.com/spreadsheets/d/x\n"))
        duas = dict(CFG, contato={"clickup_marcar": [{"id": "1", "nome": "Ana"}, {"id": "2", "nome": "Bia"}]})
        self.assertIn("[@Ana](#user_mention#1) e [@Bia](#user_mention#2), a lista", saidas.resumo_canal(m, duas, link="L"))
        self.assertNotIn("user_mention", saidas.resumo_canal(m, uma))
        self.assertIn("Lista da semana: L", saidas.resumo_canal(m, CFG, link="L"))

    def test_lista_tem_abas_na_ordem_e_chave(self):
        import openpyxl

        caminho = saidas.gravar_lista(self.r, CFG, self.tmp.name)
        livro = openpyxl.load_workbook(caminho)
        nomes = livro.sheetnames
        self.assertEqual(nomes[0], "Resumo")
        self.assertEqual(nomes[1], "Chegada sem acesso")
        self.assertIn("chave", [c.value for c in livro[nomes[1]][1]])


class TestArquivos(unittest.TestCase):
    def test_ultima_com_colchetes_no_nome_da_pasta(self):
        with tempfile.TemporaryDirectory() as tmp:
            pasta = pathlib.Path(tmp) / "Jornada do Sim [TT][POS]" / "leituras"
            pasta.mkdir(parents=True)
            for nome in ("tmb-2026-10-20_0900.json", "tmb-2026-10-23_0900.json"):
                (pasta / nome).write_text("{}", encoding="utf-8")
            self.assertEqual(jornada.ultima(pasta, "tmb-*.json").name, "tmb-2026-10-23_0900.json")

    def test_rodar_duas_vezes_no_mesmo_dia_nao_apaga_a_comparacao(self):
        with tempfile.TemporaryDirectory() as tmp:
            pasta = pathlib.Path(tmp)
            def revisao(data, cor):
                linha = {"chave": "h:abc", "cor": cor, "grupo": cor, "prioridade": 4}
                m = {"data_revisao": data, "data_corte": None, "chegada": {"compraram": 0, "entrou": 0}, "grupos": {"chegada": 0},
                     "base": {"ativos": 1, "por_cor": {"verde": 0, "amarelo": 0, "laranja": 0, "vermelho": 0, "nao_entrou": 0}, "mudaram_de_cor": 0}}
                saidas.atualizar_registros({"linhas": [linha], "metricas": m}, pasta, D.fromisoformat(data))
            revisao("2026-10-19", "verde")
            revisao("2026-10-26", "laranja")
            revisao("2026-10-26", "laranja")  # de novo no mesmo dia
            self.assertEqual(jornada.ler_estado(pasta, D(2026, 10, 26)), {"h:abc": "verde"})
            self.assertEqual(jornada.ler_estado(pasta, D(2026, 11, 2)), {"h:abc": "laranja"})

    def test_ler_json_em_texto(self):
        with tempfile.TemporaryDirectory() as tmp:
            arq = pathlib.Path(tmp) / "voomp.json"
            arq.write_text('### Result\n{"plataforma": "voomp", "linhas": []}\n', encoding="utf-8")
            self.assertEqual(jornada.ler_json(arq)["plataforma"], "voomp")


class TestDrive(unittest.TestCase):
    def test_mesclar_anotacoes(self):
        antigo = {"h:1": {"Responsável": "Lindsy", "Resultado": "sem resposta"}, "h:2": {"Anotações": "saiu"}}
        novo = drive.mesclar_anotacoes(antigo, [{"chave": "h:1"}, {"chave": "h:3"}])
        self.assertEqual(novo[0]["Responsável"], "Lindsy")
        self.assertEqual(novo[0]["Resultado"], "sem resposta")
        self.assertNotIn("Anotações", novo[1])

    def test_compartilhar_por_nome_sem_email_do_drive(self):
        def falso(metodo, url, tok, dados=None):
            if dados["emailAddress"].startswith("semgoogle@"):
                raise RuntimeError("Drive POST permissions: HTTP 400 invalidSharingRequest")
            return {}

        with mock.patch.object(drive, "_json", side_effect=falso) as chamada:
            avisos = drive.compartilhar_por_nome("ID", ["Ja@x.com", "fulana@gmail.com", "semgoogle@x.com"], {"ja@x.com"}, "tok")
        urls = [c.args[1] for c in chamada.call_args_list]
        self.assertEqual(len(urls), 2)
        self.assertTrue(all(u.endswith("sendNotificationEmail=false") for u in urls))
        self.assertIn("fulana@gmail.com", avisos[0])
        self.assertTrue(avisos[1].startswith("NÃO COMPARTILHOU com semgoogle@x.com"))


if __name__ == "__main__":
    unittest.main()
