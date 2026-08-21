# Dominios Centrais v2

Este documento formaliza as planilhas `PESSOAS v2 - DEV` e `VIGENCIAS v2 - DEV` como contratos operacionais em ambiente de desenvolvimento. A migracao inicial e a conferencia manual ja foram concluidas; a v2 nao deve ser tratada como rascunho nem receber nova migracao legado -> v2 sem decisao explicita.

## Principios

- `PESSOAS` concentra identidade, identificadores, vinculos e estado consolidado.
- `VIGENCIAS` concentra cargos, funcoes, gestoes, permissoes e responsabilidades no tempo.
- `ATIVIDADES` continua sendo fonte de atividades, frequencia, presencas, faltas, justificativas, apresentacoes e carga horaria.
- `CORE` fornece contratos, leitura segura, auditoria, diagnostico legado/v2, recalculo de caches e composicao.
- `PESSOAS_RESUMO_OPERACIONAL` e `VIGENCIAS_RESUMO_ATUAL` sao caches calculados, nao fontes normativas.
- Perfil de portal deve ser consequencia de vinculo, vigencia, `CARGOS_CONFIG` e excecoes, nao o contrario.

## Pessoas v2

### PESSOAS_BASE

Fonte de identidade comum da pessoa no ecossistema. Guarda dados cadastrais basicos como nome, e-mail principal, CPF, telefone e estado cadastral. Nao decide se alguem e membro, egresso, professor ou externo.

### PESSOAS_IDENTIFICADORES

Fonte de identificadores vinculados a `ID_PESSOA`, como RGA, CPF, e-mail, `ID_PROFESSOR` e `ID_PARTICIPANTE_EXTERNO`. Deve ser usada para resolucao segura de identidade.

### MEMBROS_DETALHES

Fonte complementar do dominio de membros. Guarda RGA, semestre, data de integracao original e historico academico. Nao decide frequencia, apresentacoes ou cargos.

### COLABORADORES_ACADEMICOS

Fonte cadastral de professores, tecnicos e colaboradores academicos. Guarda instituicao, titulacao, area de atuacao e eixo associado. Nao e base de destinatarios de comunicacao nem de links de perfil.

### PESSOAS_LINKS_PERFIS

Fonte geral de links academicos e perfis por `ID_PESSOA`. Ela atende membros,
egressos, conselheiros, colaboradores academicos, participantes externos e
outros perfis sem repetir colunas de link em cada detalhe cadastral.

Campos do contrato: `ID_LINK`, `ID_PESSOA`, `TIPO_LINK`, `URL`, `ROTULO`,
`PUBLICAVEL`, `VISIVEL_PORTAL`, `FONTE`, `VALIDADO_EM`, `ATIVO`, `CRIADO_EM`,
`ATUALIZADO_EM` e `OBS`. Os tipos iniciais sao `LATTES`, `LINKEDIN`, `ORCID`,
`INSTAGRAM`, `SITE_PESSOAL`, `GOOGLE_SCHOLAR`, `RESEARCHGATE` e `OUTRO`.

Durante o rollout a aba e opcional para leitura. Depois de provisionada, ela e
a fonte exclusiva de links para o contrato autenticado de Meu Perfil. A coluna
legada em `COLABORADORES_ACADEMICOS` so pode ser removida apos a verificacao
integral por pessoa, tipo e URL.

Visibilidade:

- o proprio usuario autenticado pode receber seus links ativos no contrato de
  `Meu perfil`;
- leituras para superficies publicas usam somente registros ativos com
  `PUBLICAVEL = SIM` e `VISIVEL_PORTAL = SIM`;
- `OBS` e `FONTE` nao fazem parte do payload exibido pelo Portal.

Provisionamento e migracao manual em DEV:

1. Rode `corePessoasV2PrepareLinksPerfis({ dryRun: true })`.
2. Rode `corePessoasV2PrepareLinksPerfis({ dryRun: false, confirmacao: 'PREPARAR_PESSOAS_V2_LINKS_PERFIS' })`.
3. Garanta `PESSOAS_V2_DB` no Registry DEV apontando para a planilha que contem
   a aba canonica `PESSOAS_LINKS_PERFIS`. A key especifica
   `PESSOAS_V2_LINKS_PERFIS` e apenas fallback temporario de leitura.
4. Rode `corePessoasV2MigrarLinksPerfisLegados({ dryRun: true, ambiente: 'DEV' })`.
5. Depois de conferir as linhas, rode a mesma funcao com `dryRun: false` e
   `confirmacao: 'MIGRAR_LATTES_LEGADO_PESSOAS_V2'`.
6. Rode `corePessoasV2VerificarRemocaoLinkLattesLegado({ ambiente: 'DEV' })`.
7. Somente com `prontoParaRemocao: true`, execute
   `corePessoasV2RemoverColunaLinkLattesLegadoReal()`.

A migracao e idempotente por `ID_PESSOA + TIPO_LINK + URL`, cria apenas
`LATTES` ainda ausentes e inicia `PUBLICAVEL`/`VISIVEL_PORTAL` como `NAO`.
Edicao pelo membro permanece fora do escopo desta etapa. A remocao nao atinge
o campo-fonte `CURRICULO_LATTES` dos formularios de docentes/tecnicos nem os
campos de links das abas editoriais publicas.

### PARTICIPANTES_EXTERNOS_DETALHES

Fonte cadastral de participantes externos. Guarda categoria de publico, instituicao, cargo/profissao, curso, cidade e autorizacao de contato.

### VINCULOS_GEAPA

Fonte normativa do vinculo institucional com o GEAPA. Define tipo de vinculo, status, datas e origem do processo. E o local correto para diferenciar membro efetivo, egresso e membro em espera. A nomenclatura `EX_MEMBRO` fica como alias legado; o contrato novo deve preferir `EGRESSO`.

Quando uma pessoa deixa de ser membro, o vinculo `MEMBRO_EFETIVO` deve ser encerrado e um vinculo atual `EGRESSO` pode representar o estado institucional posterior. Assim, um egresso pode ter `TIPO_VINCULO = EGRESSO`, `STATUS_VINCULO = ATIVO` e `ATIVO = SIM`, enquanto o vinculo antigo de membro fica encerrado.

### MEMBROS_EVENTOS_VINCULO

Fonte de eventos de ciclo de vida de membros. Apenas eventos homologados/processaveis devem produzir efeitos em vinculos e resumos.

### PESSOAS_COMUNICACAO_CONSENTIMENTOS

Fonte de consentimento e segmentacao de comunicacao. Deve ser usada por comunicacoes abertas, incluindo egressos, professores/tecnicos e externos quando houver consentimento ativo.

### PORTAL_ACESSOS_EXCECOES

Fonte de excecoes operacionais do portal. Nao substitui vinculos, cargos ou autenticacao.

### PESSOAS_RESUMO_OPERACIONAL

Cache/visao calculada para leitura rapida. Consolida vinculo atual, status, cargo atual, perfil calculado e indicadores vindos de Vigencias e Atividades.

Recalculo operacional:

- usar `coreRecalcularPessoasResumoOperacionalV2({ dryRun: true })` para validar a amostra;
- usar `coreRecalcularPessoasResumoOperacionalV2({ dryRun: false, confirmacao: 'RECALCULAR_PESSOAS_RESUMO_V2' })` somente depois de conferir a amostra;
- a funcao atualiza por `ID_PESSOA` e adiciona linhas ausentes, preservando cabecalhos, colunas extras e valores manuais fora do contrato derivado;
- identidade, RGA, vinculo, portal, tempo efetivo, semestres, apresentacoes, frequencia, pendencias, suspensao e elegibilidade basica sao recalculados a partir das fontes oficiais disponiveis;
- `CICLO_ULTIMA_APRESENTACAO` e o cabecalho canonico; `PERIODO_ULTIMA_APRESENTACAO` e aceito apenas durante a transicao;
- Atividades v2 e lida por Registry no mesmo ambiente DEV de Pessoas/Vigencias durante a homologacao;
- o relatorio inclui contagens de campos preenchidos/sem valor e motivos de fontes indisponiveis;
- campos sem regra segura, como `DATA_LIMITE_ESTIMADA_DIRETORIA`, ficam vazios e aparecem em `camposNaoCalculaveis`.
- o relatorio inclui `divergenciasLegado` como resumo diagnostico somente leitura quando a comparacao com legado for possivel.

Recalculo de semestre atual do curso:

- usar `coreRecalcularMembrosDetalhesSemestreAtualV2({ dryRun: true })` para conferir a amostra;
- usar `coreRecalcularMembrosDetalhesSemestreAtualV2({ dryRun: false, confirmacao: 'RECALCULAR_MEMBROS_DETALHES_SEMESTRE_ATUAL_V2' })` para escrever;
- a funcao altera apenas `MEMBROS_DETALHES.SEMESTRE_ATUAL`;
- `SEMESTRE_ATUAL` representa o semestre atual do aluno no curso, calculado pelo RGA;
- `SEMESTRE_ENTRADA` representa o semestre de entrada do membro no GEAPA e nao deve ser recalculado pelo RGA.

Diagnostico operacional:

- usar `coreDiagnosticarPessoasResumoOperacionalV2()` para verificar lacunas e divergencias sem alterar dados.

## Vigencias v2

### SEMESTRES

Fonte de semestres, janelas academicas e parametros operacionais relacionados ao calendario, como reunioes previstas e periodos de matricula/ajuste.

### CICLOS

Fonte de ciclos de gestao/operacao e parametros congelados do ciclo, incluindo membros previstos, limites de faltas, planejamento e dados SEI. O contrato usa `VIGENCIAS_V2_CICLOS`, `ID_CICLO`, `NOME_CICLO` e `TIPO_CICLO`.

Durante a transicao, `ID_PERIODO`, `NOME_PERIODO` e `TIPO_PERIODO` sao aceitos
somente como aliases de leitura. Novas escritas e configuracoes devem usar os
nomes de ciclo.

### DIRETORIAS

Fonte de gestoes/diretorias, datas, status, ata de posse e lema institucional da gestao.

### SEMESTRES_DIRETORIA

Fonte de janelas de diretoria. Guarda datas, ordem, total de dias e peso operacional para limites da diretoria.

### CARGOS_CONFIG

Fonte de configuracao de cargos/funcoes, grupos, hierarquia, permissao de portal, obrigatoriedade, e-mails de grupo e regras de composicao.

As colunas transitorias `PODE_*` podem existir para compatibilidade com rotinas antigas, mas nao sao obrigatorias para recalculos operacionais. A fonte final de autorizacao do Portal e `PORTAL_PERMISSOES`.

### VIGENCIAS_FUNCOES

Fonte normativa temporal de funcoes/cargos exercidos por pessoas. Deve apontar para `ID_PESSOA`, `ID_VINCULO`, `CARGO_KEY`, diretoria/janela e datas de vigencia.

### VIGENCIAS_RESUMO_ATUAL

Cache/visao calculada com funcoes vigentes, perfis e permissoes calculadas. Nao deve ser editada como fonte normativa.

Recalculo operacional:

- usar `coreRecalcularVigenciasResumoAtualV2({ dryRun: true })` para validar a amostra;
- usar `coreRecalcularVigenciasResumoAtualV2({ dryRun: false, confirmacao: 'RECALCULAR_RESUMO_ATUAL_V2' })` somente depois de conferir a amostra;
- a funcao limpa e reescreve apenas as linhas de dados do cache `VIGENCIAS_RESUMO_ATUAL`.
- `PERFIS_PORTAL_CALCULADOS` e `PERMISSOES_CALCULADAS` sao derivados de `CARGOS_CONFIG`.

Rotina segura nova:

- usar `coreVigenciasV2AtualizarResumoAtual({ dryRun: true })` para validar a previa;
- usar `coreVigenciasV2AtualizarResumoAtual({ dryRun: false })` para escrita manual controlada;
- a funcao nova usa Registry, adiciona colunas faltantes ao final quando necessario, atualiza/anexa por cabecalho e nao limpa a aba inteira.

## APIs Operacionais Publicas

Pessoas v2:

- `corePessoasV2Diagnostico(options)`
- `corePessoasV2ConferirConsistencia(options)`
- `corePessoasV2AtualizarResumoOperacional(options)`
- `corePessoasGetById(idPessoa)`
- `corePessoasFindByEmail(email)`
- `corePessoasFindByRga(rga)`
- `corePessoasGetOperationalSummary(idPessoa)`
- `corePessoasListCurrentMembers(opts)`
- `corePessoasListEffectiveMembers(opts)`
- `corePessoasListExMembers(opts)`
- `corePessoasListWaitingMembers(opts)`
- `corePessoasListAcademicCollaborators(opts)`
- `corePessoasListExternalParticipants(opts)`
- `coreRecalcularMembrosDetalhesSemestreAtualV2(options)`
- `coreDiagnosticarPessoasResumoOperacionalV2(options)`

Vigencias v2:

- `coreVigenciasV2Diagnostico(options)`
- `coreVigenciasV2ConferirConsistencia(options)`
- `coreVigenciasV2AtualizarResumoAtual(options)`
- `coreVigenciasGetCurrentFunctionByPessoa(idPessoa)`
- `coreVigenciasListCurrentFunctions(opts)`
- `coreVigenciasGetPortalPermissionsByPessoa(idPessoa)`

Diagnostico somente leitura:

- `coreV2DiagnosticoGeral(options)`
- `coreAuditarPessoasV2()`
- `coreAuditarVigenciasV2()`
- `coreCompararLegadoComV2(opts)`

## Relacao Com Atividades v2

Atividades v2 continuara sendo a fonte para frequencia, presenca, falta, justificativas, apresentacoes e carga horaria. Pessoas v2 e Vigencias v2 podem exibir resumos desses dados, mas nao devem recalcula-los como fonte primaria.

## O Que Nao Fazer Diretamente

- Nao limpar ou recriar as planilhas v2 sem confirmacao explicita.
- Nao alterar planilhas legadas durante auditoria v2.
- Nao editar caches como fonte normativa.
- Nao cadastrar cargo novo apenas por edicao manual em resumo.
- Nao usar `COLABORADORES_ACADEMICOS` como lista de disparo de comunicacao.
- Nao substituir o portal atual antes do diagnostico de prontidao.

## Integracao Futura

O `geapa-membros` devera operar o dominio Pessoas usando APIs do CORE. O portal devera consumir leituras normalizadas e caches recalculados pelo CORE. As substituicoes de consumidores legados devem ser graduais, com auditoria e diagnostico de prontidao antes da troca de contrato.
