# Perfil editavel e correcoes cadastrais - HOMOLOG

Esta entrega prepara contratos do Core para o Portal HOMOLOG. Ela nao autoriza
publicacao em PROD, nao altera `PESSOAS_RESUMO_OPERACIONAL` diretamente e nao
inclui upload de foto.

## Ambiente e fontes

O projeto Apps Script HOMOLOG deve ter a Script Property `GEAPA_ENV=DEV`. O
Core atual aceita somente `DEV` ou `PROD`; por isso HOMOLOG usa as entradas DEV
do Registry e as bases Pessoas V2 DEV.

Fontes oficiais de escrita:

- `PESSOAS_BASE` para telefone, Instagram, cidade/UF e campos sensiveis que
  pertencem ao cadastro central;
- `PESSOAS_V2_LINKS_PERFIS` para Lattes, ORCID, LinkedIn, Google Scholar,
  ResearchGate, site pessoal e `OUTRO`;
- `MEMBROS_DETALHES` para historico/resumo academico e RGA aprovado;
- `SOLICITACOES_ATUALIZACAO_CADASTRAL` para solicitacoes, decisoes,
  idempotencia e trilha de alteracao.

## Setup idempotente

Primeiro execute somente o diagnostico no editor do projeto HOMOLOG:

```javascript
geapaCoreSetupSolicitacoesAtualizacaoCadastralDev({ dryRun: true });
```

Depois de conferir `environment=DEV`, a planilha Pessoas V2 e as operacoes
planejadas, execute:

```javascript
geapaCoreSetupSolicitacoesAtualizacaoCadastralDev({
  dryRun: false,
  confirmacao: 'PREPARAR_SOLICITACOES_CADASTRAIS_DEV'
});
```

Alternativamente, a entrada manual sem argumentos e:

```javascript
geapaCoreSetupSolicitacoesAtualizacaoCadastralDevReal();
```

O setup:

- recusa qualquer ambiente diferente de `DEV`, especialmente `PROD`;
- nao cria planilha e resolve a planilha Pessoas V2 por `PESSOAS_V2_BASE`;
- cria ou completa a aba sem apagar registros;
- garante a key DEV `PESSOAS_V2_SOLICITACOES_ATUALIZACAO_CADASTRAL`;
- nunca cria ou atualiza uma linha PROD;
- garante `membros:analisar_correcoes` para `SECRETARIA` e `DIRETORIA` em
  `PORTAL_PERMISSOES` DEV;
- garante `ATUALIZADO_EM` em `MEMBROS_DETALHES`;
- aplica validacao de dados aos campos `STATUS` e `ATIVO`;
- registra um diagnostico seguro no Logger.

## Contratos para o Portal HOMOLOG

```javascript
GEAPA_CORE.geapaCoreAtualizarMeuPerfilParaPortal(payload, contexto);
GEAPA_CORE.geapaCoreSolicitarCorrecaoMeuPerfilParaPortal(payload, contexto);
GEAPA_CORE.geapaCoreListarMinhasSolicitacoesCadastraisPortal(contexto);
GEAPA_CORE.geapaCoreListarSolicitacoesCadastraisAdministracaoPortal(filtros, contexto);
GEAPA_CORE.geapaCoreAnalisarSolicitacaoCadastralPortal(payload, contexto);
GEAPA_CORE.geapaCoreAplicarSolicitacaoCadastralAprovadaPortal(payload, contexto);
```

Toda mutacao exige `chaveIdempotencia` com 8 a 120 caracteres. O Core resolve
novamente a pessoa pelo e-mail da sessao oficial e rejeita `ID_PESSOA` ou RGA
como identidade-alvo no payload.

Exemplo de edicao direta:

```javascript
{
  chaveIdempotencia: 'perfil-20260712-0001',
  telefone: '(65) 99999-9999',
  ufOrigem: 'MT',
  resumoAcademico: 'Resumo atualizado.',
  links: [
    { tipo: 'LATTES', url: 'https://lattes.cnpq.br/...' }
  ]
}
```

URL vazia remove voluntariamente o link. Texto vazio remove voluntariamente um
campo direto quando o campo admite remocao.

Exemplo de solicitacao sensivel:

```javascript
{
  chaveIdempotencia: 'cpf-20260712-0001',
  campo: 'CPF',
  valorSolicitado: '00000000000',
  justificativa: 'Solicito correcao conforme documento oficial apresentado.'
}
```

Campos sensiveis aceitos: `NOME_COMPLETO`, `NOME_CIVIL` quando o cabecalho
existir, `CPF`, `RGA`, `DATA_NASCIMENTO` e `EMAIL_PRINCIPAL`. Vinculo, perfil,
cargos, funcoes e permissoes nao fazem parte deste fluxo.

A aprovacao administrativa apenas marca `APROVADA`. A escrita ocorre somente
na chamada explicita de aplicacao. A aplicacao revalida o valor, confere o hash
do valor anterior, atualiza a fonte e recalcula a view operacional para a pessoa.

## Checklist manual de HOMOLOG

1. Confirmar branch `codex/perfil-editavel-homolog` e worktree limpo.
2. Confirmar que `.clasp.json` aponta para o projeto Apps Script do Core
   HOMOLOG, nunca para PROD.
3. Confirmar `GEAPA_ENV=DEV` no projeto HOMOLOG.
4. Executar o setup em `dryRun: true` e conferir que nenhuma entrada PROD e
   mencionada.
5. Executar o setup real com a confirmacao explicita.
6. Reexecutar o setup real e confirmar idempotencia: nenhuma linha, aba ou
   cabecalho duplicado.
7. Executar `geapaCoreRunTestesAtualizacaoCadastral()`; esperado: `ok=true` e
   `total=14`.
8. Em uma conta de membro HOMOLOG, alterar telefone e resumo academico.
9. Adicionar e remover um link Lattes.
10. Solicitar correcao de CPF e confirmar que `PESSOAS_BASE.CPF` nao mudou.
11. Confirmar que uma solicitacao duplicada para o mesmo campo e rejeitada.
12. Confirmar que membro comum nao acessa a fila administrativa.
13. Com conta Secretaria/Diretoria, marcar `EM_ANALISE`, aprovar e verificar
    que a fonte ainda nao mudou.
14. Aplicar explicitamente, conferir a fonte oficial e a view derivada.
15. Repetir a aplicacao e confirmar resposta idempotente.
16. Conferir que listagens e Logger nao exibem CPF completo, ID de planilha,
    token ou log tecnico.
17. Revalidar `Minha situacao`, `Meu perfil` atual e `Admin > Membros`.

## Publicacao exclusiva em HOMOLOG

A versao sugerida da Library e **12**, por ser a sucessora da versao 11 hoje
referenciada pelo manifesto do Portal local. Antes de criar a versao, execute
`clasp versions`; se 12 ja existir, use a proxima versao sequencial livre.

No checkout do Core configurado para o projeto Apps Script HOMOLOG:

```powershell
git switch codex/perfil-editavel-homolog
git pull --ff-only
npx.cmd @google/clasp versions
npx.cmd @google/clasp push --force
npx.cmd @google/clasp version "Perfil editavel e correcoes cadastrais - HOMOLOG"
```

No projeto Apps Script do Portal HOMOLOG, atualizar somente a dependencia
`GEAPA_CORE` para a versao retornada, fazer `clasp push` e atualizar somente o
deployment HOMOLOG. Nao modificar o manifesto, deployment ou versao da Library
do Portal PROD.
