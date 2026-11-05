# jornada-do-sim

Skill do Claude Code que faz a revisão semanal dos alunos depois da compra numa Pós da Anhanguera da Exos. Toda
segunda cruza quem comprou (TMB e Voomp) com a planilha de engajamento da Anhanguera, põe cada aluno numa fase e
numa cor do semáforo e entrega ao suporte a lista de ação no Google Sheets, as mensagens sugeridas e um resumo só
com números no canal do ClickUp. **A skill não manda mensagem a aluno.**

```
planilha da Anhanguera (Gmail)  +  compradoras (TMB pela API, Voomp no navegador)  →  revisar  →  lista de ação (Sheets), mensagens e resumo do canal
```

## Instalação

```bash
git clone https://github.com/Exos-Marketing/jornada-do-sim.git \
  ~/.claude/skills/jornada-do-sim
```

Requisitos:

- [uv](https://docs.astral.sh/uv/): o `jornada.sh` roda o motor isolado e traz o `openpyxl`.
- Token da API da TMB no Keychain, no serviço `tmb-exos-api`.
- O MCP do Playwright com a Voomp logada.
- rclone com o remote `gdrive` (Meu Drive de `drive@exosmkt.com`): a lista de ação é gravada com o token dele.
- Composio com o Gmail que recebe a planilha da Anhanguera (ou o anexo salvo à mão em Downloads).
- A pasta de trabalho da Exos em `~/Documents/Exos`: cada Pós tem o config em
  `<Expert>/08-Resultados-Relatorios/Jornada do Sim [SIGLA][POS]/config.json`.

**Nenhuma credencial mora neste repo.**

## Uso

```bash
SK=~/.claude/skills/jornada-do-sim/scripts
CFG="$HOME/Documents/Exos/<Expert>/08-Resultados-Relatorios/Jornada do Sim [SIGLA][POS]/config.json"
bash $SK/jornada.sh listar
bash $SK/jornada.sh anhanguera --config "$CFG" --arquivo "<planilha.xlsx>"
TMB_API_TOKEN=$(security find-generic-password -s tmb-exos-api -w) bash $SK/jornada.sh coletar-tmb --config "$CFG"
bash $SK/jornada.sh coletor-voomp --config "$CFG"      # função para o browser_evaluate
bash $SK/jornada.sh revisar --config "$CFG"
bash $SK/jornada.sh publicar --config "$CFG"
```

No Claude Code basta pedir "revisão da jornada", ou abrir a sessão numa segunda-feira com projeto configurado.

## Regras

- Só leitura nas plataformas; nada de mensagem, ligação ou task em nome de ninguém.
- Dado pessoal só na lista de ação (xlsx local e Sheets compartilhado por nome) e em `leituras/`; no canal e no
  chat, só números.
- Config, turmas, textos e canal só mudam com confirmação; os textos das mensagens precisam da aprovação do Cleber.
- Mexeu no motor: `cd $SK && python3 -m unittest test_jornada`.

## O que tem aqui

| arquivo | o que faz |
|---|---|
| `SKILL.md` | revisão de segunda, configuração de uma Pós nova e regras |
| `REFERENCE.md` | método (fases e semáforo), schema do config e fontes |
| `modelos/mensagens.json` | modelos das mensagens por momento (chegada, nunca entrou, resgate, hábito…) |
| `scripts/jornada.sh` | porta de entrada: roda o motor isolado com o openpyxl do uv |
| `scripts/jornada.py` | o motor: listar, anhanguera, coletar-tmb, coletor-voomp, revisar e publicar |
| `scripts/anhanguera.py` | lê a planilha de engajamento da Anhanguera |
| `scripts/compradoras.py` | compradoras na TMB (API) e na Voomp (coletor), com contato |
| `scripts/saidas.py` | lista de ação (xlsx), mensagens, resumo do canal e registros |
| `scripts/drive.py` | grava a lista no Google Sheets sempre no mesmo link, guardando as anotações do time |
| `scripts/coletores/voomp_compradoras.js` | coletor de compradoras da Voomp para o `browser_evaluate` (só GET) |
| `scripts/config_exemplo.json` | modelo de config de uma Pós |
| `scripts/test_jornada.py` | testes do motor |

---

Interno · Exos Marketing Digital e Internacionalização Ltda.
