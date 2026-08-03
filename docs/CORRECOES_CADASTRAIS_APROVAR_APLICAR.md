# Aprovação e aplicação de correções cadastrais

## Contrato

`geapaCoreAprovarEAplicarSolicitacaoCadastralPortal(payload, contexto)` executa a decisão e a aplicação sob o mesmo lock. O `ID_PESSOA`, o ator e o ambiente são resolvidos no backend; o navegador envia apenas `idSolicitacao`, `confirmacao` e a chave de idempotência.

Novas decisões transitam de `PENDENTE`, `EM_ANALISE` ou `COMPLEMENTO_SOLICITADO` diretamente para `APLICADA`. Solicitações antigas em `APROVADA` também são aceitas. `APLICADA` retorna sucesso idempotente. Falhas posteriores a alguma escrita são compensadas e registradas como `ERRO_APLICACAO`, permitindo reprocessamento.

O contrato antigo `geapaCoreAplicarSolicitacaoCadastralAprovadaPortal` permanece somente como alias para solicitações legadas já em `APROVADA`.

## Sincronização por campo

| Campo | Fonte principal | Sincronização adicional |
| --- | --- | --- |
| `EMAIL_PRINCIPAL` | `PESSOAS_BASE` | `PESSOAS_IDENTIFICADORES`, resumo operacional e caches de identidade |
| `RGA` | `MEMBROS_DETALHES` | `PESSOAS_IDENTIFICADORES`, resumo operacional e caches de identidade |
| `NOME_COMPLETO`, `NOME_CIVIL`, `CPF`, `DATA_NASCIMENTO` | `PESSOAS_BASE` | resumo operacional |
| `CURSO_ID` | `MEMBROS_DETALHES` | resumo operacional |

CPF não é sincronizado em `PESSOAS_IDENTIFICADORES`: o modelo oficial atualmente resolve identidade por EMAIL e RGA, e não há evidência de CPF como identificador oficial. Essa regra deve ser revista antes de incluir CPF nessa fonte.

Para EMAIL e RGA, a regra canônica é exatamente um identificador ativo e principal por tipo e pessoa. O novo valor é criado ou reativado; registros anteriores são preservados, desativados e deixam de ser principais. Um valor ativo pertencente a outra pessoa bloqueia a operação.

## Consistência e compensação

Antes de escrever, o Core valida ambiente, permissão, schemas, conflito de identidade, valor atual e hash da solicitação. As linhas afetadas são copiadas para memória. Em falha, as mutações são revertidas em ordem inversa e o resumo operacional é restaurado ou recalculado.

Google Sheets não oferece transação multiaba. Portanto, uma falha de compensação é reportada de forma explícita e segura; valores pessoais não são incluídos em logs.

## Cache e sessão

Após EMAIL ou RGA, são invalidados caches pelo valor antigo, novo e `ID_PESSOA`, sempre com separação de ambiente. O Portal também invalida os caches locais de sessão e “Minha situação” e remove os dados internos de invalidação antes de responder ao navegador.

Uma sessão já aberta pode continuar apenas até a próxima validação oficial. Depois da troca de e-mail, a orientação operacional é encerrar a sessão e entrar novamente com o novo e-mail. O identificador antigo inativo não deve resolver um novo login.

## Diagnóstico e reparação PROD

O diagnóstico é somente leitura por padrão:

```javascript
geapaCoreDiagnosticarReparacaoSolicitacaoCadastralProd({
  idSolicitacao: 'SAC-F1BD0911-69B4-42EE-99CE-7F5FD0081CB2',
  dryRun: true
});
```

Ele informa estado mascarado, identificadores, divergências, plano, caches e o token exato. A execução real não faz parte desta entrega e exige:

`REPARAR_SOLICITACAO_CADASTRAL_PROD_SAC-F1BD0911-69B4-42EE-99CE-7F5FD0081CB2`

O token é validado antes de abrir as fontes para escrita.
