# Perfil editavel e correcoes cadastrais - preparacao PROD

Esta branch prepara a promocao do pacote cadastral para producao sem executar
setup, publicar Library ou alterar deployments.

## Isolamento de ambiente

- `HOMOLOG` e `DEV` resolvem exclusivamente entradas `DEV` do Registry;
- `PROD` resolve exclusivamente entradas `PROD`;
- entradas `ALL` nao sao aceitas como fallback para fontes cadastrais;
- o contexto de ambiente deve vir do backend do Portal;
- em PROD, Secretaria/Diretoria nao recebem permissao apenas pelo perfil:
  `membros:analisar_correcoes` deve estar na sessao oficial.

Fontes obrigatorias em PROD:

- `PESSOAS_V2_BASE`;
- `PESSOAS_V2_SOLICITACOES_ATUALIZACAO_CADASTRAL`;
- `PORTAL_PERMISSOES`.

## Setup seguro

Diagnostico sem escrita:

```javascript
geapaCoreSetupSolicitacoesAtualizacaoCadastralProd({
  dryRun: true,
  environment: 'PROD'
});
```

Depois de revisar planilha, Registry, cabecalhos e permissoes:

```javascript
geapaCoreSetupSolicitacoesAtualizacaoCadastralProd({
  dryRun: false,
  environment: 'PROD',
  confirmacao: 'PREPARAR_SOLICITACOES_CADASTRAIS_PROD'
});
```

A entrada manual equivalente e
`geapaCoreSetupSolicitacoesAtualizacaoCadastralProdReal()`. Ela nao deve ser
executada antes da aprovacao desta branch.

O setup usa a planilha ja registrada em `PESSOAS_V2_BASE`, cria ou completa a
aba de solicitacoes, cadastra a key PROD e garante a permissao para Secretaria
e Diretoria. Nao cria planilha nova e nao altera a Script Property `GEAPA_ENV`.

## Testes e publicacao futura

`geapaCoreRunTestesAtualizacaoCadastral()` deve retornar `ok=true` e
`total=22`. A publicacao futura deve criar a Library v12 somente depois da
revisao e dos testes de setup em dry-run. Consumidores PROD devem fixar v12 com
`developmentMode=false`; HOMOLOG pode continuar no snapshot/HEAD ja aprovado.

## Rollback

O Core v11 permanece imutavel. Se a promocao for interrompida, o Portal PROD
continua apontando para v11 e nenhuma rotina nova fica acessivel. Linhas de
solicitacao eventualmente criadas sao trilha auditavel e nao devem ser apagadas
automaticamente.
