# Portal Autorizacao v2

Este documento descreve a camada do GEAPA-CORE para o Portal GEAPA consumir Pessoas v2 e Vigencias v2 sem ler planilhas diretamente para decidir acesso.

## Fontes

- `PESSOAS_RESUMO_OPERACIONAL` e a fonte principal para identidade operacional, vinculo atual, perfil portal calculado e `PORTAL_ATIVO`.
- `PESSOAS_BASE`, `PESSOAS_IDENTIFICADORES`, `MEMBROS_DETALHES`, `VINCULOS_GEAPA` e `PORTAL_ACESSOS_EXCECOES` complementam identidade, RGA, vinculo e excecoes.
- `VIGENCIAS_RESUMO_ATUAL` e a fonte principal para cargos/funcoes atuais no login do Portal.
- `CARGOS_CONFIG` e `VIGENCIAS_FUNCOES` continuam como fontes normativas de Vigencias, mas nao devem ser cruzadas diretamente no login quando o resumo estiver atualizado.
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
- `corePortalResolverUsuarioAtual` prioriza `PESSOAS_RESUMO_OPERACIONAL` e `VIGENCIAS_RESUMO_ATUAL` para reduzir cruzamentos pesados no login.
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

As funcoes legadas `geapaCoreBuscarMembroParaPortal(emailOuRga)`, `geapaCoreBuscarUsuarioPortal(emailOuRga)` e `geapaCoreBuscarMinhaSituacaoParaPortal(emailOuRga)` continuam disponiveis. Elas tentam Pessoas v2 como fonte principal e so caem para `MEMBERS_ATUAIS` como fallback de compatibilidade quando a v2 nao localizar a pessoa ou estiver indisponivel. Quando Pessoas v2 localiza uma pessoa sem `portalAtivo`, o contrato legado nao contorna o bloqueio pelo fallback.

## Configuracao operacional do Portal

`corePortalGetOperationalConfig(opts)` le a aba `PORTAL_CONFIG` via Registry e retorna um objeto simples por chave.

Colunas esperadas:

- `CHAVE`
- `VALOR`
- `ATIVO`
- `DESCRICAO`

Regras:

- somente linhas com `ATIVO = SIM` entram no retorno;
- chaves sao normalizadas para caixa alta sem acentos;
- `SIM`/`NAO` viram boolean;
- numeros viram `Number`;
- demais valores permanecem string;
- chaves com indicios de segredo, token, senha, API key ou credencial nao sao expostas;
- o resultado usa `CacheService` por 10 minutos.

Exemplo de retorno:

```js
{
  ATIVIDADES_CHAMADA_ANTECEDENCIA_MINUTOS: 60,
  ATIVIDADES_CHAMADA_TOLERANCIA_POS_MINUTOS: 240,
  ATIVIDADES_PRELOAD_DETALHES: true
}
```

Use `corePortalClearConfigCache()` apenas em diagnostico/manual quando for necessario invalidar o cache antes da expiracao.

## Funcoes Publicas

- `corePortalGetOperationalConfig(opts)`
- `corePortalClearConfigCache()`
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
- `test_core_portalConfig_normalizacao_fakeRecords()`
- `test_core_portalConfig_lerOperacional()`
- `test_core_portalV2_resolverUsuarioAtual_emailExemplo()`
- `test_core_portalV2_resolverUsuarioAtual_entradaObjeto()`
- `test_core_portalV2_resolverUsuarioAtual_rgaConfigurado()`
- `test_core_portalV2_resolverUsuarioAtual_idPessoaConfigurado()`
- `test_core_portalV2_diagnosticarSessoesPortal()`
- `test_core_portalV2_buscarMembroLegado_identificadorConfigurado()`
- `test_core_portalV2_buscarUsuarioLegado_identificadorConfigurado()`
- `test_core_portalV2_minhaSituacaoV2_identificadorConfigurado()`

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
