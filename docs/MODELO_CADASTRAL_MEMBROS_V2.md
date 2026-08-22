# Modelo cadastral de membros V2

Esta entrega prepara, somente para DEV/HOMOLOG, a evolucao aditiva dos dados de origem, curso e periodo academico. Ela nao administra processo seletivo e nao cria vinculo por si mesma.

## Localidades

O catalogo versionado em `39a_core_locality_catalog_data.js` e gerado por `scripts/update_locality_catalog.mjs` a partir da [API de Localidades do IBGE](https://servicodados.ibge.gov.br/api/docs/localidades), versao 1.0.0. O envio do formulario nao depende de rede externa: o backend valida pais, UF, municipio, codigo e nome contra o snapshot local.

Para o Brasil, `PAIS_ORIGEM_CODIGO=BR`, a UF deve ser uma das 27 unidades federativas e o codigo do municipio deve pertencer a essa UF. Para outro pais, a cidade e a regiao podem ser texto, enquanto UF e codigo de municipio ficam vazios. Origem nao representa residencia atual.

## Cursos e dados academicos

`CURSOS_CATALOGO` e uma fonte propria. O setup recomenda somente Agronomia - UFMT - Sinop e `OUTRO`; a inclusao de outros cursos exige decisao institucional. `OUTRO` exige nome complementar.

- `SEMESTRE_ENTRADA`: ingresso no GEAPA.
- `PERIODO_INGRESSO_CURSO`: inicio da graduacao, derivado do RGA.
- `SEMESTRE_ATUAL_CURSO_CALCULADO`: estimativa ordinal obtida pela sequencia real de `VIGENCIAS_V2_SEMESTRES` iniciados.

A estimativa nao considera trancamentos, reprovacoes, aproveitamentos ou alteracoes individuais. Intervalos sem semestre letivo iniciado nao incrementam o valor. O membro nao edita curso, RGA, periodo de ingresso ou semestre calculado; correcoes seguem o fluxo cadastral sensivel.

`STATUS_COMPLETUDE_CADASTRAL` usa regra deterministica (`PENDENTE`, `PARCIAL` ou `COMPLETO`) e nao bloqueia o Portal. Nao e exibido percentual enquanto nao houver regra institucional documentada.

## Identificadores tecnicos

Novas pessoas usam `ID_PESSOA` sequencial no formato `PES-000001`. A alocacao
considera o maior ID numerico existente e IDs ja reservados na fila de ingresso,
sempre dentro do mesmo lock da admissao. A faixa `PES-900000+` e tratada como
reservada e nao avanca a sequencia regular. Identificadores antigos fora desse
padrao nao participam do calculo da sequencia e nao devem ser renomeados em uma
unica aba: qualquer saneamento precisa atualizar todas as referencias V2 de
forma coordenada e previamente auditada.

## Setup DEV/HOMOLOG

Funcao publica:

```javascript
geapaCoreSetupEvolucaoMembrosV2Dev({ ambiente: 'DEV', dryRun: true })
```

O dry-run e o padrao, nao escreve e retorna ambiente, planilha, aba, cabecalhos existentes e ausentes, linhas recomendadas, alteracoes planejadas, token e indicador de idempotencia. A escrita real, que nao faz parte desta entrega, aceita somente DEV e exige o token exato `PREPARAR_EVOLUCAO_MEMBROS_V2_DEV`.

O setup adiciona cabecalhos somente ao final, nao apaga, nao reordena e nao sobrescreve dados. Abas ausentes sao planejadas para `CURSOS_CATALOGO`, `INGRESSOS_MEMBROS`, `CONVITES_AVALIACAO_EGRESSOS` e `RESPOSTAS_AVALIACAO_EGRESSOS`. As keys especificas DEV sao apenas compatibilidade; escritas operacionais devem resolver `PESSOAS_V2_DB`.

## Implantacao manual futura

1. Executar e revisar o dry-run em DEV.
2. Confirmar que `PESSOAS_V2_DB` resolve exclusivamente a planilha DEV.
3. Revisar as linhas iniciais do catalogo de cursos.
4. Executar o setup real somente com aprovacao explicita e token exato.
5. Repetir o dry-run e confirmar que nao ha cabecalhos ou linhas pendentes.
6. Cadastrar as keys DEV planejadas no Registry apenas pela rotina aprovada.
7. Validar que ADMIN e DIRETORIA receberam `membros:cadastrar_novos_membros`; SECRETARIA nao recebe automaticamente.

PROD e recusado pela rotina desta branch. Nenhum ID de planilha e codificado no codigo operacional.
