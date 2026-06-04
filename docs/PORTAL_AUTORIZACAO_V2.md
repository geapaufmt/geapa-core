# Portal Autorizacao v2

Este documento descreve a camada do GEAPA-CORE para o Portal GEAPA consumir Pessoas v2 e Vigencias v2 sem ler planilhas diretamente para decidir acesso.

## Fontes

- `PESSOAS_BASE`, `PESSOAS_IDENTIFICADORES`, `VINCULOS_GEAPA`, `PESSOAS_RESUMO_OPERACIONAL` e `PORTAL_ACESSOS_EXCECOES` definem identidade, vinculo e excecoes.
- `CARGOS_CONFIG`, `VIGENCIAS_FUNCOES` e `VIGENCIAS_RESUMO_ATUAL` definem cargos/funcoes vigentes e o perfil gerado por cargo.
- `PORTAL_PERFIS` e o catalogo oficial de perfis.
- `PORTAL_PERMISSOES` e a fonte oficial das permissoes efetivas.

## Perfil, Permissao, Cargo e Vinculo

- Perfil e a categoria de acesso do portal, como `MEMBRO`, `DIRETORIA`, `EGRESSO` ou `CONSELHO`.
- Permissao e uma capacidade granular no formato `modulo:acao`, como `portal:acessar` ou `apresentacoes:ver_ate_saida`.
- Cargo/função pertence a Vigencias e pode gerar um perfil por `PERFIL_PORTAL_PADRAO`.
- Vinculo pertence a Pessoas e indica a relacao institucional, como `MEMBRO_EFETIVO` ou `EGRESSO`.
- Acesso efetivo e calculado pelo CORE, cruzando perfil, permissoes e excecoes.

## Regras

- `PORTAL_PERMISSOES` prevalece como fonte final de autorizacao.
- `CARGOS_CONFIG` nao e fonte final de permissoes; colunas `PODE_*` sao tratadas como transitorias/depreciadas para autorizacao final.
- `ADMIN` so deve vir de excecao explicita ativa em `PORTAL_ACESSOS_EXCECOES`.
- Diretoria, presidente ou vice nao viram `ADMIN` automaticamente.
- Egressos podem receber perfil `EGRESSO` e a permissao `apresentacoes:ver_ate_saida`.
- Para egresso, apresentacoes permitidas sao limitadas a `DATA_ATIVIDADE <= DATA_FIM` do vinculo encerrado.
- Se a data de saida do egresso estiver ausente, a listagem retorna bloqueio seguro.

## Funcao oficial de sessao

`corePortalResolverUsuarioAtual(entrada, opts)` e a fonte oficial para o Portal resolver a sessao autorizavel do usuario atual.

A entrada pode ser:

- string com e-mail;
- string com RGA;
- string com `ID_PESSOA`;
- objeto com `email`, `rga`, `idPessoa`, `identificador` ou `emailOuRga`.

Saida segura esperada:

```js
{
  ok: true,
  autenticado: true,
  idPessoa: "PES-000001",
  nomeExibicao: "Membro GEAPA",
  email: "membro@example.org",
  rga: "202300000000",
  perfilPortalEfetivo: "MEMBRO",
  perfisPortal: ["MEMBRO"],
  permissoes: ["portal:acessar", "situacao:ver_propria"],
  portalAtivo: true,
  tipoVinculoAtual: "MEMBRO_EFETIVO",
  statusVinculoAtual: "ATIVO",
  cargoFuncaoAtual: "",
  cargosAtuais: []
}
```

O retorno nao inclui linhas brutas, IDs de planilhas, tokens, chaves internas ou listas de terceiros.

As funcoes legadas `geapaCoreBuscarMembroParaPortal(emailOuRga)` e `geapaCoreBuscarMinhaSituacaoParaPortal(emailOuRga)` continuam disponiveis. A segunda pode incluir `sessao` no retorno para facilitar a transicao do backend do Portal.

## Funcoes Publicas

- `corePortalResolverUsuarioAtual(entrada, opts)`
- `corePortalCalcularPerfilEfetivo(idPessoa)`
- `corePortalListarPermissoesEfetivas(idPessoa)`
- `corePortalValidarAcesso(idPessoa, permissaoOuPerfil)`
- `corePortalGetMeuResumo(email)`
- `corePortalListarApresentacoesPermitidas(email, options)`
- `corePortalListarApresentacoesParaEgresso(idPessoa)`
- `corePortalDiagnosticarPerfisEPermissoes()`
- `corePrepararPortalParaV2()`

## Diagnostico

Use `corePrepararPortalParaV2()` antes de integrar o `geapa-portal`. A funcao retorna `PRONTO`, `PARCIAL` ou `BLOQUEADO`, alem de bloqueios, avisos, funcoes disponiveis, perfis invalidos, permissoes invalidas e recomendacoes.

Testes manuais disponiveis:

- `test_core_portalV2_diagnosticarPerfisEPermissoes()`
- `test_core_portalV2_prepararPortalParaV2()`
- `test_core_portalV2_resolverUsuarioAtual_emailExemplo()`
- `test_core_portalV2_resolverUsuarioAtual_entradaObjeto()`
- `test_core_portalV2_diagnosticarSessoesPortal()`

Para testar varios perfis, configure a Script Property `GEAPA_CORE_PORTAL_TESTE_SESSOES` com um JSON:

```json
{
  "MEMBRO": "membro@example.org",
  "DIRETORIA": "diretoria@example.org",
  "SECRETARIA": "secretaria@example.org",
  "COMUNICACAO": "comunicacao@example.org",
  "CONSELHO": "conselho@example.org",
  "EGRESSO": "egresso@example.org",
  "COLABORADOR": "colaborador@example.org",
  "EXTERNO": "externo@example.org",
  "VISITANTE": "visitante@example.org",
  "ADMIN": "admin@example.org"
}
```
