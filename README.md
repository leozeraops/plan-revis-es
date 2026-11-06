# Planilha de controle de revisões

Este projeto contém o código da aplicação da planilha de controle de revisões, executado no Google Apps Script V8. O GitHub é usado para versionamento do código e acompanhamento de issues e tickets; a aplicação não foi migrada para outra plataforma.

## Aplicação

- `appsscript.json` configura o runtime V8 e o fuso horário do Apps Script.
- `.clasp.json` associa o diretório ao projeto Apps Script usado pelo `clasp`.
- Os arquivos JavaScript na raiz implementam as integrações e os fluxos da planilha.

O deploy continua sendo feito para Google Apps Script. Revise as alterações antes de publicar com as ferramentas do `clasp`.

## Skills de engenharia

As skills estão instaladas em `.agents/skills/`. As instruções do projeto ficam em `AGENTS.md`; configuração do rastreador, vocabulário de triagem e regras de documentação de domínio ficam em `docs/agents/`.

Para consultar as issues e tickets, autentique o GitHub CLI com `gh auth login` e use o repositório `leozeraops/plan-revis-es`.

## Decisões e trabalho registrado

- O rastreador de trabalho é GitHub Issues; os rótulos padrão de triagem estão configurados no repositório.
- A análise inicial de arquitetura está em [`architecture-review-report.html`](./architecture-review-report.html).
- A spec de melhorias está na [issue #1](https://github.com/leozeraops/plan-revis-es/issues/1).
- Os tickets filhos da spec cobrem testes locais, importações Salesforce e fluxos Lovable; veja a lista de sub-issues da issue #1.

## Setup concluído

1. As skills de Matt Pocock foram copiadas para este projeto, com o inventário registrado em `skills-lock.json`.
2. O Git local foi inicializado na branch `main`, com o remoto `origin` apontando para `https://github.com/leozeraops/plan-revis-es.git`.
3. Foram criadas as instruções `AGENTS.md` e a documentação de apoio em `docs/agents/`.
4. Os rótulos padrão de triagem foram verificados e os que faltavam foram criados no GitHub.
5. A análise de arquitetura foi registrada em HTML e sua spec e tickets foram publicados em GitHub Issues.

## Testes

Um mecanismo de testes locais para o Apps Script V8 está planejado no ticket de testes da issue #1. Até que seja implementado, não há comando de teste automatizado documentado para a aplicação.
