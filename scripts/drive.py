"""Lista de ação no Google Sheets, sempre no mesmo link, guardando as anotações do time.

Usa a API do Drive com o token do próprio rclone (remote `gdrive`, Meu Drive da conta). A planilha fica fora do
Drive compartilhado do cliente, onde todo membro enxerga tudo, e é compartilhada por nome.
"""
import io
import json
import pathlib
import subprocess
import time
import urllib.error
import urllib.parse
import urllib.request
import uuid

API = "https://www.googleapis.com/drive/v3/"
UPLOAD = "https://www.googleapis.com/upload/drive/v3/files"
XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
SHEET = "application/vnd.google-apps.spreadsheet"
COLUNAS_DO_TIME = ["Responsável", "Contato em", "Resultado", "Anotações"]


def token():
    """Access token do remote gdrive do rclone; um comando rclone antes renova o token, se vencido."""
    subprocess.run(["rclone", "lsf", "gdrive:", "--max-depth", "1", "--dirs-only"], capture_output=True, timeout=120)
    dump = json.loads(subprocess.run(["rclone", "config", "dump"], capture_output=True, text=True, timeout=60).stdout)
    return json.loads(dump["gdrive"]["token"])["access_token"]


def _pedir(metodo, url, tok, corpo=None, tipo=None, tentativas=6):
    """Requisição com nova tentativa e espera em limite de uso (403 rateLimitExceeded, 429) e erro 5xx."""
    espera = 4
    for tentativa in range(tentativas):
        cab = {"Authorization": "Bearer " + tok}
        if tipo:
            cab["Content-Type"] = tipo
        req = urllib.request.Request(url, data=corpo, method=metodo, headers=cab)
        try:
            with urllib.request.urlopen(req, timeout=120) as r:
                return r.read()
        except urllib.error.HTTPError as erro:
            texto = erro.read().decode("utf-8", "replace")
            limite = erro.code == 429 or erro.code >= 500 or (erro.code == 403 and "ateLimitExceeded" in texto)
            if limite and tentativa < tentativas - 1:
                time.sleep(espera)
                espera *= 2
                continue
            raise RuntimeError("Drive %s %s: HTTP %d %s" % (metodo, url.split("?")[0], erro.code, texto[:300]))


def _json(metodo, url, tok, dados=None):
    corpo = json.dumps(dados).encode("utf-8") if dados is not None else None
    bruto = _pedir(metodo, url, tok, corpo, "application/json; charset=UTF-8" if corpo else None)
    return json.loads(bruto.decode("utf-8")) if bruto else {}


def pasta(nome, tok):
    """Id da pasta com esse nome na raiz do Meu Drive; cria se não existir."""
    q = "name = '%s' and mimeType = 'application/vnd.google-apps.folder' and 'root' in parents and trashed = false" % nome.replace("'", "\\'")
    achadas = _json("GET", API + "files?" + urllib.parse.urlencode({"q": q, "fields": "files(id,name)"}), tok).get("files", [])
    if achadas:
        return achadas[0]["id"]
    return _json("POST", API + "files", tok, {"name": nome, "mimeType": "application/vnd.google-apps.folder"})["id"]


def criar(xlsx_bytes, nome, pasta_id, tok):
    fronteira = "jornada" + uuid.uuid4().hex
    meta = {"name": nome, "mimeType": SHEET}
    if pasta_id:
        meta["parents"] = [pasta_id]
    corpo = ("--%s\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n%s\r\n--%s\r\nContent-Type: %s\r\n\r\n"
             % (fronteira, json.dumps(meta), fronteira, XLSX)).encode("utf-8") + xlsx_bytes + ("\r\n--%s--" % fronteira).encode("utf-8")
    bruto = _pedir("POST", UPLOAD + "?uploadType=multipart&fields=id,name,webViewLink", tok, corpo,
                   "multipart/related; boundary=" + fronteira)
    return json.loads(bruto.decode("utf-8"))


def substituir(planilha_id, xlsx_bytes, tok):
    bruto = _pedir("PATCH", UPLOAD + "/" + planilha_id + "?uploadType=media&fields=id,name,webViewLink", tok, xlsx_bytes, XLSX)
    return json.loads(bruto.decode("utf-8"))


def exportar(planilha_id, tok):
    return _pedir("GET", API + "files/" + planilha_id + "/export?" + urllib.parse.urlencode({"mimeType": XLSX}), tok)


def anotacoes_do_xlsx(xlsx_bytes):
    """{chave: {coluna do time: valor}} de todas as abas de grupo da planilha atual."""
    import openpyxl

    livro = openpyxl.load_workbook(io.BytesIO(xlsx_bytes), data_only=True)
    saida = {}
    for aba in livro.worksheets:
        cab = [c.value for c in aba[1]] if aba.max_row else []
        if "chave" not in cab:
            continue
        idx = {nome: cab.index(nome) for nome in COLUNAS_DO_TIME + ["chave"] if nome in cab}
        for linha in aba.iter_rows(min_row=2, values_only=True):
            chave = linha[idx["chave"]]
            if not chave:
                continue
            valores = {n: linha[i] for n, i in idx.items() if n != "chave" and linha[i] not in (None, "")}
            if valores:
                saida[str(chave)] = valores
    return saida


def mesclar_anotacoes(antigo, novo):
    """Devolve as linhas novas com as colunas do time que já estavam preenchidas para a mesma chave.

    antigo: {chave: {coluna: valor}}; novo: lista de dicts com 'chave'. Quem saiu da lista sai com as anotações.
    """
    saida = []
    for linha in novo:
        linha = dict(linha)
        for coluna, valor in (antigo.get(str(linha.get("chave"))) or {}).items():
            if linha.get(coluna) in (None, ""):
                linha[coluna] = valor
        saida.append(linha)
    return saida


def aplicar_anotacoes(xlsx_path, antigo):
    """Escreve no xlsx novo (em disco) as anotações antigas, linha a linha pela chave. Devolve quantas linhas herdaram."""
    import openpyxl

    livro = openpyxl.load_workbook(xlsx_path)
    herdadas = 0
    for aba in livro.worksheets:
        cab = [c.value for c in aba[1]] if aba.max_row else []
        if "chave" not in cab:
            continue
        ic = cab.index("chave")
        linhas = [{"chave": r[ic].value, "_r": r[ic].row} for r in aba.iter_rows(min_row=2)]
        for linha in mesclar_anotacoes(antigo, linhas):
            if any(linha.get(c) not in (None, "") for c in COLUNAS_DO_TIME):
                herdadas += 1
            for coluna in COLUNAS_DO_TIME:
                if coluna in cab and linha.get(coluna) not in (None, ""):
                    aba.cell(row=linha["_r"], column=cab.index(coluna) + 1).value = linha[coluna]
    livro.save(xlsx_path)
    return herdadas


def compartilhar_por_nome(planilha_id, emails, ja, tok):
    """Acesso de edição por nome, sem o e-mail do Drive: o aviso é a marcação no canal do ClickUp. Devolve os avisos.

    ja: e-mails (minúsculos) que já têm acesso. Conta sem Google não aceita compartilhamento sem e-mail: vira aviso.
    """
    avisos = []
    for email in emails:
        if email.lower() in ja:
            continue
        try:
            _json("POST", API + "files/" + planilha_id + "/permissions?sendNotificationEmail=false", tok,
                  {"role": "writer", "type": "user", "emailAddress": email})
            avisos.append("Compartilhada com %s (pode editar; o Drive não mandou e-mail)." % email)
        except RuntimeError as erro:
            avisos.append("NÃO COMPARTILHOU com %s: %s" % (email, erro))
    return avisos


def publicar(xlsx_path, cfg, compartilhar=False):
    tok = token()
    entrega = cfg.setdefault("entrega", {})
    nome = "Jornada do Sim · Lista de ação · %s" % (cfg.get("titulo") or cfg.get("projeto"))
    avisos = []
    planilha_id = entrega.get("drive_planilha_id")
    if planilha_id:
        atual = exportar(planilha_id, tok)
        # cópia do que está no Drive antes de sobrescrever (tem dado pessoal: fica na pasta local do projeto)
        backup = pathlib.Path(xlsx_path).parent / ("backup-sheets-%s.xlsx" % time.strftime("%H%M%S"))
        backup.write_bytes(atual)
        antigo = anotacoes_do_xlsx(atual)
        herdadas = aplicar_anotacoes(xlsx_path, antigo)
        info = substituir(planilha_id, open(xlsx_path, "rb").read(), tok)
        avisos.append("Anotações do time levadas para %d linha(s)." % herdadas)
        novo = False
    else:
        pasta_id = entrega.get("drive_pasta_id") or pasta(entrega.get("drive_pasta_nome", "Jornada do Sim"), tok)
        info = criar(open(xlsx_path, "rb").read(), nome, pasta_id, tok)
        novo = True
    planilha_id = info["id"]
    permissoes = _json("GET", API + "files/" + planilha_id + "/permissions?fields=permissions(id,type,role,emailAddress)", tok)
    for p in permissoes.get("permissions", []):
        if p.get("type") in ("anyone", "domain"):
            avisos.append("ATENÇÃO: a planilha está aberta para %s; o certo é só por nome." % p["type"])
    if compartilhar:
        ja = {p.get("emailAddress", "").lower() for p in permissoes.get("permissions", [])}
        avisos += compartilhar_por_nome(planilha_id, cfg.get("contato", {}).get("emails_drive", []), ja, tok)
    return {"id": planilha_id, "link": "https://docs.google.com/spreadsheets/d/%s" % planilha_id,
            "novo_id": novo, "avisos": avisos}
