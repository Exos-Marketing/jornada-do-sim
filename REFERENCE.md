# jornada-do-sim · referência

Desenho aprovado em 23/10/2026: `Exos/Jornada do Sim/2026-10-23-desenho-da-skill.md`. Origem do método: o deck
"Depois do sim" (`Exos/presentations/depois-do-sim/`).

## Método

**Fases** (D = dias desde o pagamento): chegada D0 a D7 · primeiro valor D8 a D30 · hábito D31 a D90 ·
continuidade D91 a D180 · renovação depois do D180.

**Acesso de quem comprou** (janela de `janela_chegada_dias`, padrão 30):

| Situação | Quando |
|---|---|
| entrou | casou com aluno da planilha que não está em "nunca" nem em "contrato" (sai da chegada e ganha cor) |
| não entrou | casou e o aluno nunca acessou |
| contrato pendente | casou e o contrato está pendente (barreira: ninguém com contrato pendente concluiu disciplina) |
| ainda sem dado | não casou e a planilha é anterior a `pago_em + prazo_acesso_dias` |
| sem matrícula | não casou e a planilha já cobre o prazo |

Casamento: e-mail embaralhado (SHA-256 aparado e minúsculo, 12 hex, igual à `acompanhar-vendas`); senão telefone
(últimos 11 dígitos) com o alerta "provavelmente estudando com outro e-mail". Nunca pelo nome.

**Grupos da Pós** (porta do analisador antigo, mesma ordem e mesmas regras) e **cor**:

| Grupo | Regra (na data de corte) | Cor |
|---|---|---|
| nunca | sem último acesso | não entrou |
| contrato | contrato "Pendente" | não entrou |
| adormecido | 30+ dias sem acesso | laranja se concluiu ≥ 1 disciplina; vermelho se nenhuma |
| novato | menos de 45 dias de curso | amarelo |
| travado | 0 disciplina com 45+ dias | amarelo |
| atrasado | 2+ disciplinas atrás do ritmo | amarelo |
| emdia | no ritmo | verde |
| destaque | adiantado (−1) ou 5+ disciplinas | verde |

Ritmo esperado = dias de curso ÷ (duração ÷ disciplinas). Disciplina concluída = nota ≥ `nota_minima`.
Data de corte = Data Final − Dias que faltam para integralização (a mais comum). "Pré-confirmada" não é barreira:
vira o alerta "matrícula pré-confirmada". Canceladas ficam fora.

**Prioridade da lista:** 1 Chegada sem acesso · 2 Nunca entrou · 3 Mudou de cor (piorou desde a revisão anterior,
pelo `estado.csv`) · 4 Vermelhos · 5 Laranjas · 6 Amarelos · 7 Verdes. Dentro do grupo, maior D primeiro.

**Turmas** (`coortes`): contadas pela **data da compra** (o boleto pode pagar depois). Limites `AAAA-MM-DD` (dia
inteiro) ou `AAAA-MM-DDTHH:MM`, no horário de Brasília. A Voomp grava em UTC: o motor converte (UTC−3).

## Config

O molde é `scripts/config_exemplo.json` (campos vazios). Nunca copie o config de outro expert: produtos, turmas e
`arquivo_padrao` dele colocariam alunas de outra Pós na lista. O que cada campo faz:

| Campo | Para quê |
|---|---|
| `prazo_acesso_dias` (3) | dias depois do pagamento até a planilha poder dizer "sem matrícula" |
| `janela_chegada_dias` (30) | quem comprou há até tantos dias entra na chegada |
| `planilha_velha_dias` (8) | a partir daí o resumo avisa que a base está velha |
| `anhanguera.*` | regras da plataforma; `arquivo_padrao` específico da Pós (várias chegam na mesma caixa) |
| `vendas.tmb.produtos` / `vendas.voomp.produtos` | ids de todas as ofertas do produto (não só as da Black) |
| `vendas.leituras_acompanhar_vendas` | reserva sem contato, relativa à pasta do config; só se o usuário aceitar seguir sem a Voomp |
| `coortes` | turmas, contadas pela data da compra |
| `contato` | quem contata (texto), responsáveis, `clickup_marcar` (`[{id, nome}]`, marcados na linha do link do resumo) e `emails_drive` (acesso por nome à planilha, sem e-mail do Drive) |
| `entrega.canal_aprovado` | `false` até o Cleber aprovar o primeiro resumo; depois `true` |
| `entrega.drive_planilha_id` | gravado pelo primeiro `publicar` |
| `mensagens.*` | campos dos textos; vazio aparece entre colchetes. `aprovado_em` = data da aprovação dos textos |

## Anotações do time e histórico de cores

- Cada linha da lista tem a coluna oculta `chave` (e-mail embaralhado, ou matrícula). O `publicar` exporta o Sheets,
  guarda uma cópia em `saidas/<dia>/backup-sheets-HHMMSS.xlsx` e devolve Responsável, Contato em, Resultado e Anotações
  à linha de mesma `chave`, em qualquer aba. Ordenar, filtrar, inserir linha e renomear aba não atrapalham; apagar a
  coluna `chave` sim. Quem sai da lista sai com as anotações. Rodar duas vezes é seguro.
- A chave muda se a aluna aparecer com outro e-mail na planilha da plataforma: a anotação antiga fica no backup.
- `estado.csv` guarda a cor de cada chave em cada revisão. "Mudou de cor" compara com a revisão anterior a hoje,
  então refazer a revisão no mesmo dia não apaga a comparação.

## Fontes

- **Planilha da Anhanguera** (abas "Base de Alunos" e "Desempenho Acadêmico"). Colunas da base, nesta ordem: curso,
  matrícula, nome, e-mail, celular, status, contrato, último acesso, início, fim, dias para integralização, turma.
  Não tem CPF nem a data do primeiro acesso. Chega toda segunda por e-mail em contato@agenciaeu.com.br (remetente
  academicoexperts@anhanguera.com), com o retrato do dia do envio: a de 19/10/2026 tem data de corte 19/10.
- **TMB** `api/pedidos?produto_id=..&pageNumber=..&pageSize=100` com Bearer do Keychain `tmb-exos-api`. Campos usados:
  `pedido_id`, `produto_id`, `status_pedido` (Efetivado = pago), `status_financeiro`, `criado_em` (com −03:00),
  `data_efetivado` (horário local), `cliente`, `email`, `telefone`, `documento`.
- **Voomp** `apiadm.voompcreators.com.br/api/sale` pela sessão do navegador (`coletores/voomp_compradoras.js`). Pagas
  guardam nome, e-mail e celular. `refused`/`waiting_payment` = pendente; `refunded`/`chargedback` = cancelado.
- **Drive**: token do rclone (`gdrive`, conta drive@exosmkt.com). O cliente OAuth do rclone é compartilhado e às vezes
  estoura o limite por minuto (403 rateLimitExceeded): o `drive.py` tenta de novo com espera.

## Saídas

- `Lista de ação [PROJETO] AAAA-MM-DD.xlsx`: aba Resumo (números da semana e a tabela "O que tem em cada aba", de
  `EXPLICACAO_DAS_ABAS` em `saidas.py`) + uma aba por grupo. Na coluna Acesso, contrato pendente aparece como
  "contrato pendente" também na aba Nunca entrou (a cor é "não entrou", mas o Último acesso tem data). Colunas: Nome, WhatsApp, E-mail,
  Matrícula, Fase, Acesso, Cor, Cor anterior, Alerta, Disciplinas, Último acesso, Compra, Mensagem sugerida,
  Responsável, Contato em, Resultado (lista: entrou, vai entrar, respondeu, sem resposta, problema, não quer seguir),
  Anotações e `chave` (oculta: é por ela que as anotações voltam na semana seguinte).
- `mensagens.md`: modelos de `modelos/mensagens.json` preenchidos pelo config.
- `canal.md` e `metricas.json`: só números. O `publicar` regrava o `canal.md` com o link da planilha e é esse bloco
  que vai ao canal; na linha do link vai a marcação de `contato.clickup_marcar`. `historico.csv`: uma linha por revisão. `estado.csv`: chave, cor e data, de todas as revisões.

## Primeira execução (23/10/2026, Pós Onco da Letícia)

- Planilha de 08/09 (45 dias): 270 ativos; grupos idênticos ao analisador antigo (21 · 17 · 53 · 41 · 36 · 30 · 55 · 17).
  Semáforo: 72 verdes, 107 amarelos, 14 laranjas, 39 vermelhos e 38 que nunca entraram.
- Chegada: 110 compraram nos últimos 30 dias, todas "ainda sem dado" (planilha anterior à compra).
- Refeita no mesmo dia com a planilha de 19/10, baixada da caixa pelo Composio (e-mail de academicoexperts@anhanguera.com):
  390 ativos; 76 verdes, 128 amarelos, 23 laranjas, 62 vermelhos e 38 que nunca entraram. Chegada: 42 entraram,
  48 não entraram, 15 com contrato pendente, 5 ainda sem dado e 1 provavelmente com outro e-mail (aba Chegada com 68).
  A revisão de 08/09 ficou superada; a de 19/10 é a base do "mudou de cor" da segunda seguinte.
- Black ONCO+ (Fase 2): 94 compradoras contra 92 do fechamento da `acompanhar-vendas`: +2 vendas da oferta da Fase 1
  depois do prazo, +1 boleto efetivado em 23/10, −1 reembolso em 23/10.
- Sheets: `1QRtF5ZOF_EkHoh3ZX_0iH71JMmndeYHkPrArqB10usc` (Meu Drive de drive@exosmkt.com, pasta "Jornada do Sim").
