# Rotinas V2 de Pessoas e Vigencias

Este documento descreve a primeira camada viva de rotinas manuais para `PESSOAS v2 - DEV` e `VIGENCIAS v2 - DEV` no `geapa-core`.

Esta etapa prepara os caches e diagnosticos que serao usados pelo Portal e por outros modulos, mas o Portal ainda nao passa a consumir diretamente estas rotinas neste PR.

Para a homologacao do job coordenado, veja tambem [`docs/jobs-v2-manutencao.md`](jobs-v2-manutencao.md).

## Principios

- As rotinas resolvem abas exclusivamente pelo Registry.
- Nao ha triggers automaticos nesta etapa.
- `dryRun` e o padrao nas rotinas de atualizacao.
- Escritas reais usam cabecalho, nao indices fixos de coluna.
- Escritas reais nao apagam, renomeiam ou reordenam abas.
- Colunas faltantes dos caches podem ser adicionadas ao final, de forma nao destrutiva.
- Linhas existentes sao atualizadas por chave tecnica e linhas ausentes sao anexadas.
- Colunas extras ja existentes sao preservadas.
- Durante a homologacao, `options.ambiente` usa `DEV` por padrao para resolver as keys V2 no Registry bruto, mesmo que `GEAPA_ENV` do script esteja em outro ambiente.

## Implantacao e consumo por versao

O Core deve ser consumido pelos outros projetos por uma versao/implantacao controlada do Apps Script, nao por codigo ainda nao publicado.

Fluxo recomendado:

1. Implementar e validar no repositorio `geapa-core`.
2. Publicar uma nova versao do Apps Script Core.
3. Atualizar consumidores apenas depois de confirmar a versao publicada.
4. Registrar no changelog ou nota operacional quais funcoes publicas entraram na versao.

Neste PR nao ha alteracao em `geapa-portal` nem em `geapa-atividades`.

## MODULOS_CONFIG

As rotinas consultam `MODULOS_CONFIG` quando a camada esta disponivel. Os fluxos sugeridos sao:

- `PESSOAS / ATUALIZACAO_V2`
- `PESSOAS / CONFERENCIA_V2`
- `VIGENCIAS / ATUALIZACAO_V2`
- `VIGENCIAS / CONFERENCIA_V2`

Para atualizacoes, a capability usada e `SYNC`. Conferencias sao manuais e somente leitura.

Se a linha de `MODULOS_CONFIG` ainda nao existir, a rotina segue com aviso controlado no retorno. Se existir e bloquear a execucao, a rotina retorna bloqueio controlado e tenta registrar em `MODULOS_STATUS`.

## MODULOS_STATUS

Quando `MODULOS_STATUS` esta disponivel, as rotinas registram:

- execucao iniciada;
- sucesso;
- erro;
- bloqueio por configuracao.

Falhas ao registrar status nao impedem o diagnostico ou o `dryRun`; elas aparecem no envelope de retorno.

## Funcoes publicas

Diagnostico:

- `corePessoasV2Diagnostico(options)`
- `coreVigenciasV2Diagnostico(options)`
- `coreV2DiagnosticoGeral(options)`

Conferencia:

- `corePessoasV2ConferirConsistencia(options)`
- `coreVigenciasV2ConferirConsistencia(options)`

Atualizacao segura:

- `corePessoasV2AtualizarResumoOperacional(options)`
- `coreVigenciasV2AtualizarResumoAtual(options)`

Testes manuais:

- `coreV2RunTesteDiagnosticoGeral()`
- `coreV2RunTestePessoasResumo()`
- `coreV2RunTesteVigenciasResumo()`
- `coreV2_runTesteJobDiarioDryRun()`

Orquestracao:

- `coreV2_jobDiarioManutencao(options)`
- `coreV2InstalarTriggerJobDiario(options)`
- `coreV2RemoverTriggerJobDiario()`
- `coreV2ListarTriggerJobDiario()`

## Envelope de conferencia

As conferencias retornam um envelope padronizado:

```javascript
{
  ok: true,
  modulo: 'PESSOAS',
  fluxo: 'CONFERENCIA_V2',
  totalVerificado: 0,
  totalInconsistencias: 0,
  inconsistencias: [
    {
      gravidade: 'ERRO',
      entidade: 'PESSOA',
      idEntidade: 'PES-000001',
      campo: 'EMAIL_PRINCIPAL',
      valorAtual: 'email-invalido',
      regra: 'EMAIL_VALIDO',
      mensagem: 'EMAIL_PRINCIPAL invalido.',
      acaoRecomendada: 'Corrigir e-mail principal ou mover para observacao.'
    }
  ]
}
```

## Conferencias de Pessoas V2

A rotina `corePessoasV2ConferirConsistencia(options)` verifica, no minimo:

- `ID_PESSOA` vazio em `PESSOAS_BASE`;
- `ID_PESSOA` duplicado;
- `EMAIL_PRINCIPAL` invalido;
- `EMAIL_PRINCIPAL` duplicado em pessoas ativas diferentes;
- RGA duplicado em `PESSOAS_IDENTIFICADORES`;
- CPF fora do padrao de 11 digitos, quando preenchido;
- vinculo ativo sem pessoa valida;
- pessoa inativa com vinculo ativo;
- membro efetivo ativo sem RGA;
- evento em `MEMBROS_EVENTOS_VINCULO` sem pessoa ou vinculo correspondente quando a relacao existir.

## Conferencias de Vigencias V2

A rotina `coreVigenciasV2ConferirConsistencia(options)` verifica, no minimo:

- semestre ativo ausente;
- mais de um semestre ativo simultaneo;
- periodo ativo ausente;
- funcao vigente sem `ID_PESSOA`;
- funcao vigente com `ID_PESSOA` inexistente;
- data de fim anterior a data de inicio;
- `CARGO_KEY` sem correspondencia em `CARGOS_CONFIG`;
- cargo exclusivo duplicado no mesmo intervalo;
- diretoria vigente sem presidente ou sem vice quando aplicavel.

## Atualizacao de Pessoas

Use primeiro:

```javascript
corePessoasV2AtualizarResumoOperacional({ dryRun: true })
```

Para escrever:

```javascript
corePessoasV2AtualizarResumoOperacional({ dryRun: false })
```

Campos minimos tratados no cache:

- `ID_PESSOA`
- `NOME_EXIBICAO`
- `EMAIL_PRINCIPAL`
- `EMAIL`
- `RGA`
- `CPF`
- `TIPO_VINCULO_ATUAL`
- `STATUS_VINCULO_ATUAL`
- `DATA_INICIO_VINCULO`
- `DATA_FIM_VINCULO`
- `PORTAL_ATIVO`
- `PERFIL_PORTAL_BASE`
- `PERFIL_PORTAL_CALCULADO`
- `CARGO_FUNCAO_ATUAL`
- `ULTIMA_ATUALIZACAO`
- `OBS_RESUMO`

Se a conferencia encontrar erros, a escrita real e bloqueada por padrao. A opcao `allowWriteWithErrors: true` existe apenas para decisao operacional explicita.

## Atualizacao de Vigencias

Use primeiro:

```javascript
coreVigenciasV2AtualizarResumoAtual({ dryRun: true })
```

Para escrever:

```javascript
coreVigenciasV2AtualizarResumoAtual({ dryRun: false })
```

Campos minimos tratados no cache:

- `ID_VIGENCIA`
- `ID_PESSOA`
- `NOME_EXIBICAO`
- `RGA`
- `OCUPACAO`
- `GRUPO_FUNCAO`
- `DATA_INICIO`
- `DATA_FIM_PREVISTA`
- `DATA_FIM_REAL`
- `STATUS_VIGENCIA`
- `PERFIL_PORTAL_GERADO`
- `PERMISSOES_GERADAS`
- `CARGO_ATUAL_VISIVEL`
- `APARECE_DIRETORIA_PUBLICA`
- `ULTIMA_ATUALIZACAO`

Tambem sao preenchidos aliases de compatibilidade ja usados por consumidores existentes, como `CARGO_FUNCAO_ATUAL`, `PERFIS_PORTAL_CALCULADOS` e `PERMISSOES_CALCULADAS`.

## Limites desta etapa

- Nao cria triggers.
- Nao migra dados reais automaticamente.
- Nao altera Portal.
- Nao altera `geapa-atividades`.
- Nao remove views ou abas existentes.
- Nao limpa caches inteiros.
- Nao substitui contratos legados de consumidores existentes.

## Job diario V2

`coreV2_jobDiarioManutencao(options)` orquestra a manutencao em ordem segura:

1. conferir Registry;
2. atualizar Pessoas V2;
3. conferir Pessoas V2;
4. atualizar Vigencias V2;
5. conferir Vigencias V2;
6. atualizar views de Atividades V2, quando um runner estiver disponivel;
7. conferir Atividades V2, quando uma rotina de conferencia estiver disponivel;
8. registrar resumo final em `MODULOS_STATUS`.

Exemplo manual recomendado:

```javascript
coreV2_jobDiarioManutencao({ dryRun: true })
```

Escrita manual controlada:

```javascript
coreV2_jobDiarioManutencao({ dryRun: false })
```

Para coordenar Atividades a partir de um projeto que tenha o modulo carregado, passe os callbacks:

```javascript
GEAPA_CORE.coreV2_jobDiarioManutencao({
  dryRun: true,
  atividadesJob: atividadesV2_jobPortal,
  atividadesConferir: atividadesV2_conferirPortal
})
```

Quando os callbacks nao existem no runtime, as etapas de Atividades retornam `SKIPPED` e o job Core nao quebra.

### MODULOS_CONFIG do job Core

Linha sugerida:

- `CORE / JOB_DIARIO_V2`

Comportamento:

- `ON`: permite execucao manual e por trigger;
- `MANUAL`: permite execucao manual e bloqueia trigger;
- `DRY_RUN`: executa sem escrita real;
- `OFF`: bloqueia o job.

### Trigger do job Core

Instalar manualmente:

```javascript
coreV2InstalarTriggerJobDiario({ hour: 4 })
```

Listar:

```javascript
coreV2ListarTriggerJobDiario()
```

Remover:

```javascript
coreV2RemoverTriggerJobDiario()
```

O instalador nao e chamado automaticamente por nenhuma rotina deste PR.
