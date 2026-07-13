# Perfil editavel e correcoes cadastrais - HOMOLOG

Esta entrega prepara contratos do Core para o Portal HOMOLOG. Ela nao autoriza
publicacao em PROD, nao altera `PESSOAS_RESUMO_OPERACIONAL` diretamente e nao
inclui upload de foto.

## Ambiente e fontes

Core, Portal DEV/HOMOLOG e Portal PROD usam o mesmo projeto Apps Script. Nao
altere a Script Property global `GEAPA_ENV` para executar este fluxo. O modulo
cadastral resolve explicitamente as entradas `DEV` do Registry, enquanto os
consumidores PROD continuam presos a sua versao publicada anterior da Library.

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

- aceita somente `environment=DEV` como alvo e recusa explicitamente `PROD`;
- nao le, grava ou troca `GEAPA_ENV` para selecionar a base alvo;
- nao cria planilha e resolve a planilha Pessoas V2 por `PESSOAS_V2_BASE`;
- cria ou completa a aba sem apagar registros;
- garante a key DEV `PESSOAS_V2_SOLICITACOES_ATUALIZACAO_CADASTRAL`;
- nunca cria ou atualiza uma linha PROD;
- quando existe `PORTAL_PERMISSOES` DEV, garante
  `membros:analisar_correcoes` para `SECRETARIA` e `DIRETORIA`;
- quando essa fonte DEV nao existe, nao altera a fonte compartilhada/PROD e usa
  os perfis oficiais `SECRETARIA`/`DIRETORIA` como concessao interna somente no
  contexto `ambientePortal: 'HOMOLOG'`;
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

Toda mutacao exige `chaveIdempotencia` com 8 a 120 caracteres. O backend do
Portal HOMOLOG deve montar o contexto oficial com `ambientePortal: 'HOMOLOG'`;
esse marcador nao deve ser copiado de um campo enviado pelo navegador. O Core
resolve novamente a pessoa pelo e-mail da sessao oficial e rejeita `ID_PESSOA`
ou RGA como identidade-alvo no payload.

```javascript
var contexto = {
  ambientePortal: 'HOMOLOG',
  sessaoOficial: {
    email: sessao.email
  }
};
```

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
2. Confirmar que `.clasp.json` aponta para o projeto Apps Script unico do Core.
3. Nao alterar a Script Property `GEAPA_ENV`.
4. Executar o setup em `dryRun: true` e conferir `environment=DEV`; o campo
   `scriptEnvironmentObserved` pode continuar `PROD` sem risco.
5. Conferir que nenhuma entrada PROD e
   mencionada.
6. Executar o setup real com a confirmacao explicita.
7. Reexecutar o setup real e confirmar idempotencia: nenhuma linha, aba ou
   cabecalho duplicado.
8. Executar `geapaCoreRunTestesAtualizacaoCadastral()`; esperado: `ok=true` e
   `total=16`.
9. Em uma conta de membro HOMOLOG, alterar telefone e resumo academico.
10. Adicionar e remover um link Lattes.
11. Solicitar correcao de CPF e confirmar que `PESSOAS_BASE.CPF` nao mudou.
12. Confirmar que uma solicitacao duplicada para o mesmo campo e rejeitada.
13. Confirmar que membro comum nao acessa a fila administrativa.
14. Com conta Secretaria/Diretoria, marcar `EM_ANALISE`, aprovar e verificar
    que a fonte ainda nao mudou.
15. Aplicar explicitamente, conferir a fonte oficial e a view derivada.
16. Repetir a aplicacao e confirmar resposta idempotente.
17. Conferir que listagens e Logger nao exibem CPF completo, ID de planilha,
    token ou log tecnico.
18. Revalidar `Minha situacao`, `Meu perfil` atual e `Admin > Membros`.

## Publicacao exclusiva em HOMOLOG

A versao sugerida da Library e **12**, por ser a sucessora da versao 11 hoje
referenciada pelo manifesto do Portal local. Antes de criar a versao, execute
`clasp versions`; se 12 ja existir, use a proxima versao sequencial livre.

No checkout do Core configurado para o projeto Apps Script unico:

```powershell
git switch codex/perfil-editavel-homolog
git pull --ff-only
npx.cmd @google/clasp versions
npx.cmd @google/clasp push --force
npx.cmd @google/clasp version "Perfil editavel e correcoes cadastrais - HOMOLOG"
```

No manifesto usado pela branch/deployment do Portal HOMOLOG, atualizar somente
a dependencia `GEAPA_CORE` para a versao retornada. O manifesto/deployment do
Portal PROD deve continuar na versao anterior. A separacao ocorre pela versao
fixada no consumidor, nao por outro projeto Apps Script nem pela propriedade
`GEAPA_ENV`.
