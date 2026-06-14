# Jobs V2 de Manutencao

Este documento descreve a homologacao manual do job coordenado de manutencao V2 do Core.

O objetivo desta etapa e validar Pessoas V2, Vigencias V2 e a orquestracao geral antes de avancar para consumo direto pelo Portal.

## Funcoes publicas

Rotinas base:

- `corePessoasV2Diagnostico(options)`
- `corePessoasV2ConferirConsistencia(options)`
- `corePessoasV2AtualizarResumoOperacional(options)`
- `coreVigenciasV2Diagnostico(options)`
- `coreVigenciasV2ConferirConsistencia(options)`
- `coreVigenciasV2AtualizarResumoAtual(options)`
- `coreV2DiagnosticoGeral(options)`

Job:

- `coreV2_jobDiarioManutencao(options)`
- `coreV2_runTesteJobDiarioDryRun()`
- `coreV2InstalarTriggerJobDiario(options)`
- `coreV2RemoverTriggerJobDiario()`
- `coreV2ListarTriggerJobDiario()`

Testes manuais:

- `coreV2RunTesteDiagnosticoGeral()`
- `coreV2RunTestePessoasResumo()`
- `coreV2RunTesteVigenciasResumo()`
- `coreV2_runTesteJobDiarioDryRun()`

## Comportamento seguro esperado

- O Core resolve bases V2 pelo Registry.
- As escritas usam cabecalho, nao posicoes fixas.
- Atualizacoes manuais rodam em `dryRun` por padrao quando chamadas pelos testes.
- O job geral retorna relatorio consolidado e resumido.
- A ausencia do modulo Atividades nao quebra o Core; as etapas de Atividades retornam `SKIPPED`.
- `MODULOS_CONFIG` pode bloquear execucao ou forcar `DRY_RUN`.
- `MODULOS_STATUS` registra execucao, sucesso, erro ou bloqueio em best effort.
- Nenhuma funcao instala trigger sem chamada explicita do instalador.
- Nenhuma funcao apaga, renomeia ou reordena abas.
- Nenhuma rotina deve depender de ordem fixa de colunas.

## MODULOS_CONFIG

Linhas sugeridas para homologacao DEV:

| MODULO | FLUXO | MODO recomendado no inicio | Observacao |
| --- | --- | --- | --- |
| CORE | JOB_DIARIO_V2 | DRY_RUN | Homologar o job sem escrita real. |
| PESSOAS | ATUALIZACAO_V2 | DRY_RUN | Validar resumo operacional. |
| PESSOAS | CONFERENCIA_V2 | ON | Somente leitura. |
| VIGENCIAS | ATUALIZACAO_V2 | DRY_RUN | Validar resumo atual. |
| VIGENCIAS | CONFERENCIA_V2 | ON | Somente leitura. |

Para trigger:

- `ON`: permite execucao por trigger;
- `MANUAL`: bloqueia trigger e permite chamada manual;
- `DRY_RUN`: executa sem escrita real;
- `OFF`: bloqueia execucao.

## Ordem de teste

1. Rode `coreV2RunTesteDiagnosticoGeral()`.
2. Rode `corePessoasV2ConferirConsistencia({ limit: 5 })`.
3. Rode `coreVigenciasV2ConferirConsistencia({ limit: 5 })`.
4. Rode `coreV2RunTestePessoasResumo()`.
5. Rode `coreV2RunTesteVigenciasResumo()`.
6. Rode `coreV2_runTesteJobDiarioDryRun()`.
7. Confira `MODULOS_STATUS`.
8. Somente depois, se aprovado, avalie `coreV2_jobDiarioManutencao({ dryRun: false })` em DEV.

## O que verificar no Google Drive

Nas bases DEV:

- `PESSOAS v2 - DEV`
- `VIGENCIAS v2 - DEV`
- `ATIVIDADES INTERNAS GEAPA v2 - DEV`, quando o modulo Atividades estiver conectado

Verifique:

- as abas existem e possuem cabecalhos;
- `PESSOAS_RESUMO_OPERACIONAL` nao perdeu colunas extras;
- `VIGENCIAS_RESUMO_ATUAL` nao perdeu colunas extras;
- colunas faltantes, quando adicionadas, entraram ao final;
- nao houve criacao, renomeacao ou remocao de abas;
- `MODULOS_STATUS` recebeu resumo sem nomes, e-mails ou documentos pessoais;
- em `dryRun`, nenhuma linha de cache foi escrita.

## Erros que bloqueiam avanco

Bloqueiam avanco para Portal:

- Registry sem key V2 obrigatoria;
- aba V2 ausente;
- cabecalho obrigatorio ausente em fonte principal;
- `ID_PESSOA` vazio ou duplicado em Pessoas;
- e-mail principal invalido em pessoa ativa;
- e-mail principal duplicado em pessoas ativas diferentes;
- RGA duplicado em identificadores;
- vinculo ativo sem pessoa valida;
- membro efetivo ativo sem RGA;
- semestre ativo ausente ou duplicado;
- periodo ativo ausente;
- funcao vigente sem `ID_PESSOA`;
- funcao vigente com pessoa inexistente;
- `CARGO_KEY` vigente sem correspondencia em `CARGOS_CONFIG`;
- cargo exclusivo duplicado em intervalo sobreposto;
- job geral com etapa `ERROR`.

## Alertas aceitaveis temporariamente

Podem ser aceitos temporariamente, se documentados:

- CPF preenchido fora do padrao, desde que nao seja usado para login ou autorizacao;
- pessoa sem resumo antes da primeira atualizacao;
- campo de cargo/perfil vazio para pessoa sem funcao vigente;
- diretoria sem vice quando houver justificativa normativa temporaria;
- Atividades V2 `SKIPPED` no Core quando o modulo Atividades ainda nao estiver carregado no mesmo runtime;
- pendencias de portal relacionadas a conteudo editorial ainda nao curado.

## Checklist de homologação manual

Antes de iniciar:

- Confirmar que `GEAPA_ENV` esta em `DEV` ou que as keys DEV estao acessiveis no Registry.
- Confirmar que `MODULOS_CONFIG` possui as linhas de `PESSOAS`, `VIGENCIAS` e `CORE` listadas acima.
- Confirmar que nenhum trigger V2 foi instalado sem decisao explicita.

Sequencia minima:

1. `coreV2RunTesteDiagnosticoGeral()`
   - Esperado: todas as keys principais disponiveis.
   - Bloqueia: qualquer key/aba essencial ausente.

2. `corePessoasV2ConferirConsistencia({ limit: 5 })`
   - Esperado: envelope padronizado com `totalVerificado`.
   - Bloqueia: inconsistencias `ERRO` de identidade, e-mail duplicado, vinculo sem pessoa ou membro efetivo sem RGA.

3. `coreVigenciasV2ConferirConsistencia({ limit: 5 })`
   - Esperado: envelope padronizado com `totalVerificado`.
   - Bloqueia: semestre/periodo ativo ausente, funcao vigente sem pessoa, cargo inexistente ou cargo exclusivo duplicado.

4. `coreV2RunTestePessoasResumo()`
   - Esperado: `dryRun: true`, linhas calculadas e `totalEscrito` igual a zero.
   - Verificar no Drive: `PESSOAS_RESUMO_OPERACIONAL` nao mudou.

5. `coreV2RunTesteVigenciasResumo()`
   - Esperado: `dryRun: true`, linhas calculadas e `totalEscrito` igual a zero.
   - Verificar no Drive: `VIGENCIAS_RESUMO_ATUAL` nao mudou.

6. `coreV2_runTesteJobDiarioDryRun()`
   - Esperado: relatorio consolidado com etapas de Pessoas e Vigencias.
   - Aceitavel: Atividades como `SKIPPED` se o modulo nao estiver disponivel.
   - Bloqueia: qualquer etapa principal de Core como `ERROR`.

7. Conferir `MODULOS_STATUS`
   - Esperado: registros de execucao/sucesso/erro/bloqueio sem dados pessoais.
   - Bloqueia: erro recorrente de permissao, Registry ou escrita em cache.

Liberacao para proxima fase:

- Todas as etapas Core sem `ERROR`.
- Erros bloqueantes zerados ou formalmente saneados.
- Alertas temporarios documentados.
- Trigger V2 ainda nao instalado, salvo decisao explicita de homologacao agendada.
