# Planilha de controle de revisões

Este projeto contém o código da aplicação da planilha de controle de revisões, executado no Google Apps Script V8. O GitHub é usado para versionamento do código e acompanhamento de issues e tickets; a aplicação não foi migrada para outra plataforma.

## Aplicação

- `appsscript.json` configura o runtime V8 e o fuso horário do Apps Script.
- `.clasp.json` associa o diretório ao projeto Apps Script usado pelo `clasp`.
- Os arquivos JavaScript na raiz implementam as integrações e os fluxos da planilha.

O deploy continua sendo feito para Google Apps Script. Revise as alterações antes de publicar com as ferramentas do `clasp`.

## Tour da aplicação

Para entender o caminho de execução, comece pela aula [Tour da aplicação Apps Script e issues](lessons/0002-tour-aplicacao-apps-script-e-issues.html) ou consulte a [referência rápida](reference/apps-script-repo-issues.html).

Em resumo, ao abrir a planilha, `onOpen()` cria os menus. Cada opção chama uma função de entrada que pode ler propriedades, fazer chamadas HTTP para Salesforce ou Supabase/Lovable, transformar os dados e atualizar abas:

```text
Planilha → onOpen() → menu → função de entrada → API externa → transformação → abas
```

Arquivos principais por fluxo:

- [ParametrizaçãoAmbiente.js](Parametriza%C3%A7%C3%A3oAmbiente.js): menus, orquestração e autenticação.
- [GetFieldOficinaAgenda.js](GetFieldOficinaAgenda.js): importação da agenda Salesforce.
- [GetFieldCustosRevisão.js](GetFieldCustosRevis%C3%A3o.js): importação paginada de serviços Salesforce.
- [GetLovableOrdensService.js](GetLovableOrdensService.js): leitura de ordens e itens Lovable.
- [OrdensFinalizadas.js](OrdensFinalizadas.js): leitura de ordens finalizadas e valores.
- [Uploads.js](Uploads.js): envio de atualizações para Salesforce e Lovable.

## Melhorias acompanhadas por issues

A [issue #1](https://github.com/leozeraops/plan-revis-es/issues/1) é a spec arquitetural. Estado das issues consultado em 6 de novembro de 2026:

| Issue | Objetivo | Estado |
| --- | --- | --- |
| [#2](https://github.com/leozeraops/plan-revis-es/issues/2) | Preparar testes locais para Apps Script V8 | Aberta |
| [#3](https://github.com/leozeraops/plan-revis-es/issues/3) | Concentrar acesso Salesforce na importação de agenda | Aberta |
| [#4](https://github.com/leozeraops/plan-revis-es/issues/4) | Reusar cliente Salesforce na importação de custos; depende da #3 | Aberta |
| [#5](https://github.com/leozeraops/plan-revis-es/issues/5) | Isolar e testar atualização de custos Lovable | Aberta |
| [#6](https://github.com/leozeraops/plan-revis-es/issues/6) | Isolar e testar leitura de ordens Lovable | Aberta |
| [#7](https://github.com/leozeraops/plan-revis-es/issues/7) | Isolar e testar leitura de ordens finalizadas | Aberta |
| [#8](https://github.com/leozeraops/plan-revis-es/issues/8) | Mover credenciais para Script Properties | Fechada |

A #2 estabelece a base de testes para as demais melhorias; a #3 precede a #4. A #8 foi uma entrega separada, fora do escopo arquitetural da #1. Embora já haja testes de propriedades, a #2 continua aberta até que todos os seus critérios — inclusive isolamento dos serviços Google e exclusão dos testes do deploy — sejam verificados.

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
