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
- A spec de melhorias está na [issue #1](https://github.com/leozeraops/plan-revis-es/issues/1).
- Os tickets filhos da spec cobrem testes locais, importações Salesforce e fluxos Lovable; veja a lista de sub-issues da issue #1.

## Setup concluído

1. As skills de Matt Pocock foram copiadas para este projeto, com o inventário registrado em `skills-lock.json`.
2. O Git local foi inicializado na branch `main`, com o remoto `origin` apontando para `https://github.com/leozeraops/plan-revis-es.git`.
3. Foram criadas as instruções `AGENTS.md` e a documentação de apoio em `docs/agents/`.
4. Os rótulos padrão de triagem foram verificados e os que faltavam foram criados no GitHub.
5. A análise de arquitetura foi registrada em HTML e sua spec e tickets foram publicados em GitHub Issues.

## Testes

Os testes locais focados nas integrações podem ser executados com `node --test tests/*.test.js`. Eles simulam os serviços do Apps Script e não precisam de credenciais reais nem fazem chamadas externas.

## Credenciais das integrações

As credenciais e chaves usadas pelas integrações Salesforce e Lovable devem ser configuradas como **Script Properties** no projeto Apps Script. Não as inclua no código nem em arquivos versionados.

Propriedades necessárias:

- Salesforce: `SALESFORCE_CLIENT_ID`, `SALESFORCE_CLIENT_SECRET`, `SALESFORCE_USERNAME`, `SALESFORCE_PASSWORD` e `SALESFORCE_SECURITY_TOKEN`. O código mantém o contrato atual e concatena a senha com o security token no payload OAuth.
- Lovable: `LOVABLE_EMAIL`, `LOVABLE_PASSWORD` e `LOVABLE_API_KEY_ANON`.

No editor do Apps Script, abra **Project Settings** (ícone de engrenagem), localize **Script Properties**, adicione cada nome acima e informe seu valor no campo correspondente. Salve as propriedades antes de executar as integrações. Se alguma propriedade estiver ausente ou vazia, a integração falha antes da chamada remota e informa qual nome precisa ser configurado.
