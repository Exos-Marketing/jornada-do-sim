---
name: jornada-do-sim
description: Use quando o usuário pedir para acompanhar alunos depois da compra numa Pós da Anhanguera da Exos (inclusive a turma de uma Black), ou disser "revisão da jornada", "jornada do sim", "quem comprou e não entrou", "alunos que não acessaram", "semáforo do consumo", "lista de ação do suporte", "acompanhamento dos alunos da Pós", "chegou a planilha de engajamento da Anhanguera", "configura a jornada da Pós nova". Também numa segunda-feira com projeto configurado sem a pasta saidas/ do dia.
---

# Jornada do sim (revisão semanal do aluno)

Toda segunda, cruza quem comprou (TMB e Voomp) com a planilha de engajamento da Anhanguera, põe cada aluno numa fase
e numa cor do semáforo e entrega ao suporte a lista de ação no Google Sheets, as mensagens e um resumo só com números
no canal do ClickUp. **A skill não envia mensagem a aluno.** Hoje cobre Pós da Anhanguera; clubes, cursos da Voomp e
comunidades da GrowAI são as próximas etapas (desenho em `Exos/Jornada do Sim/`). Método, config e fontes: [REFERENCE.md](REFERENCE.md).

Todo comando passa pelo `jornada.sh` (isola o Python e traz o openpyxl; funciona no zsh). O shell não guarda
variáveis entre chamadas: defina `CFG` em cada uma.

```bash
SK=~/.claude/skills/jornada-do-sim/scripts
bash $SK/jornada.sh listar      # projetos e o caminho de cada config
CFG="$HOME/Documents/Exos/<Expert>/08-Resultados-Relatorios/Jornada do Sim [SIGLA][POS]/config.json"
```

## Revisão de segunda

1. **Planilha da Anhanguera.** Chega toda segunda em contato@agenciaeu.com.br, de "Anhanguera Experts"
   <academicoexperts@anhanguera.com>, com o assunto "Relatório Semanal de Engajamento e Inteligência de Negócios - <Pós>".
   A caixa é da Agência Eu, fora da cerca da Exos: só ler os e-mails da planilha, nunca listar ou mostrar outros.
   - No Composio, a caixa é a conta Gmail cujo `GMAIL_GET_PROFILE` devolve contato@agenciaeu.com.br (`composio
     connections list` dá os `word_id`; em 23/10/2026 era `gmail_mar-remint`). Use `--account <word_id>`.
   - Busque com `GMAIL_FETCH_EMAILS`: `from:academicoexperts@anhanguera.com has:attachment filename:xlsx newer_than:8d`
     + o nome da Pós. Sem resultado, tire o `from:` e confirme o remetente novo com o usuário. A busca não traz o
     anexo: `GMAIL_FETCH_MESSAGE_BY_MESSAGE_ID` (`format: "full"`) dá o `attachmentId`, e `GMAIL_GET_ATTACHMENT`
     devolve um `s3url`, que se baixa com `curl` numa pasta nova e vazia.
   - Sem a conta, peça que o usuário salve o anexo em Downloads.
   - Depois: `bash $SK/jornada.sh anhanguera --config "$CFG" --arquivo "<xlsx>"`. Se avisar "conta interna fora da
     exclusão", pergunte antes de mexer em `excluir_ids`.
2. **TMB:** `TMB_API_TOKEN=$(security find-generic-password -s tmb-exos-api -w) bash $SK/jornada.sh coletar-tmb --config "$CFG"`.
3. **Voomp.** Navegue a aba do Playwright para `https://adm.voompcreators.com.br/dashboard` e confira o login só com
   `browser_evaluate` devolvendo `{logado: !document.querySelector('input[type=password]')}`. Nunca use
   `browser_snapshot` nem abra `/vendas` (mostram compradoras). Sem login: pare e peça ao usuário; não digite senha.
   Logada: `bash $SK/jornada.sh coletor-voomp --config "$CFG"` e passe a função inteira ao `browser_evaluate`
   **sempre com** `filename` = `<Expert>/08-Resultados-Relatorios/Jornada do Sim [..]/leituras/voomp-AAAA-MM-DD_HHMM.json`
   (relativo à raiz do workspace Exos). Confira com `ls -l` que o arquivo caiu lá; se não caiu, pare.
4. `bash $SK/jornada.sh revisar --config "$CFG"`. Rode uma vez; para testar sem gravar nada, `--sem-gravar`.
   Leia os AVISOs: fonte sem venda paga ou sem contato, lacunas nas mensagens, textos não aprovados.
5. `bash $SK/jornada.sh publicar --config "$CFG"`: atualiza o Sheets no mesmo link, devolve as anotações do time e
   imprime o bloco final do canal, com o link. Se sair "PARE" (planilha aberta a qualquer pessoa): não poste e peça
   o OK para tirar o acesso. `--compartilhar` dá acesso de edição por nome a `contato.emails_drive`, sem o e-mail do
   Drive (o aviso é a marcação no canal), e só depois que o usuário confirmar os e-mails. "NÃO COMPARTILHOU" é conta
   sem Google: pergunte ao usuário.
6. **Entregar.** Nesta conversa: o relatório (só números) e o link. No canal `entrega.clickup_canal_id`
   (`clickup_send_chat_message`, `content_format: text/md`): o bloco "RESUMO PARA O CANAL" do `publicar`.
   O bloco já marca quem está em `contato.clickup_marcar` na linha do link (`[@Nome](#user_mention#id)`): é assim que
   o suporte fica sabendo da lista, nunca por e-mail.
   Com `entrega.canal_aprovado: false`, mostre o bloco e poste só com o OK do Cleber; depois do OK, poste e grave
   `true` no config. Com `true`, poste direto toda segunda.

Planilha velha (mais de 8 dias): a revisão roda, quem comprou depois da base sai como "ainda sem dado" e o resumo
avisa. Diga isso nesta conversa: o relatório da semana não chegou ou não foi achado na caixa.

## Configurar uma Pós nova

1. `bash $SK/jornada.sh listar`; crie `<Expert>/08-Resultados-Relatorios/Jornada do Sim [SIGLA][POS]/` e copie
   **só** o [scripts/config_exemplo.json](scripts/config_exemplo.json) (nunca o config de outro expert).
2. Preencha:
   - regras da plataforma a partir de `Exos/Engajamento Pós/config_*.json` (disciplinas, duração, nota mínima,
     matrícula de teste e domínios do expert);
   - `arquivo_padrao` específico desta Pós (as planilhas de várias Pós chegam na mesma caixa);
   - ids de **todas** as ofertas do produto na TMB e na Voomp, tirados dos configs da `acompanhar-vendas` do expert
     (`<Expert>/08-Resultados-Relatorios/Vendas */config.json`, chaves de `plataformas.*.produtos`);
   - turmas e canal, com `canal_aprovado: false`;
   - quem contata: o texto, a pessoa em `clickup_marcar` (id e nome por `clickup_find_member_by_name`) e o e-mail
     Google dela em `emails_drive`.
3. Peça os textos de `mensagens` (passo a passo do contrato e do primeiro acesso, primeiro resultado da primeira
   semana, rotina). Rode `revisar --sem-gravar`, mostre o `mensagens.md` de uma revisão de teste e grave
   `mensagens.aprovado_em` quando o Cleber aprovar. Lacuna fica entre colchetes; nunca invente link ou passo a passo.
4. Aceite: grupos iguais aos da última análise da mesma planilha (`Engajamento Pós/`) e turma batendo com o fechamento
   da `acompanhar-vendas`, com as diferenças explicadas (fuso, boleto efetivado depois, reembolso, oferta de outra fase).

## Regras

- **Só leitura** nas plataformas. Nada de mensagem, ligação ou task em nome do usuário.
- **Dado pessoal** só na lista (xlsx local, Sheets compartilhado por nome e `leituras/`). No canal e nesta conversa,
  só números. Compartilhe o arquivo, nunca a pasta "Jornada do Sim" (ela junta os experts).
- O Sheets fica no Meu Drive de drive@exosmkt.com, não no Drive compartilhado do cliente (lá todo membro vê).
- Não mude config, turmas, textos aprovados nem `canal_aprovado` sem confirmar.
- Mexeu no motor: `cd $SK && python3 -m unittest test_jornada`.
