# Bootstrap de configuracao V2

Este documento descreve a rotina segura de conferencia e preparacao das linhas
operacionais V2 na Planilha Geral do GEAPA.

A rotina atua somente sobre:

- `MODULOS_CONFIG`;
- `MODULOS_STATUS`.

Ela nao cria triggers, nao altera Registry, nao apaga linhas, nao renomeia abas
e nao sobrescreve configuracoes existentes.

## Funcoes publicas

### `coreV2_conferirConfiguracao(options)`

Conferencia somente leitura das configuracoes V2 esperadas.

Uso recomendado:

```javascript
coreV2_conferirConfiguracao({
  ambiente: 'DEV'
});
```

Essa funcao nunca escreve, mesmo que `options.criarAusentes` seja informado.

### `coreV2_bootstrapConfiguracao(options)`

Conferencia com possibilidade de criar linhas ausentes, desde que todas as
condicoes de seguranca sejam atendidas.

Opcoes:

```javascript
{
  dryRun: true,
  criarAusentes: false,
  ambiente: 'DEV',
  origem: 'MANUAL'
}
```

Regras:

- `dryRun` e `true` por padrao;
- `criarAusentes` e `false` por padrao;
- nenhuma escrita ocorre quando `dryRun: true`;
- nenhuma escrita ocorre quando `criarAusentes: false`;
- escrita real so e permitida em `ambiente: 'DEV'`;
- linhas existentes nunca sao sobrescritas;
- linhas duplicadas ou invalidas sao reportadas e bloqueiam a criacao automatica;
- escrita usa `LockService`;
- escrita usa cabecalhos, nao posicoes fixas.

Para planejar sem escrever:

```javascript
coreV2_bootstrapConfiguracao({
  dryRun: true,
  criarAusentes: true,
  ambiente: 'DEV',
  origem: 'HOMOLOGACAO_MANUAL'
});
```

Para criar somente as linhas ausentes em DEV:

```javascript
coreV2_bootstrapConfiguracao({
  dryRun: false,
  criarAusentes: true,
  ambiente: 'DEV',
  origem: 'HOMOLOGACAO_MANUAL'
});
```

### `coreV2_runTesteBootstrapDryRun()`

Atalho manual seguro para o editor do Apps Script.

Equivale a:

```javascript
coreV2_bootstrapConfiguracao({
  dryRun: true,
  criarAusentes: true,
  ambiente: 'DEV',
  origem: 'TESTE_MANUAL'
});
```

## Configuracoes conferidas

A rotina espera as seguintes linhas em `MODULOS_CONFIG` para o ambiente
informado:

| MODULO | FLUXO | Padrao criado |
|---|---|---|
| CORE | JOB_DIARIO_V2 | `MODO=DRY_RUN`, `PERMITE_TRIGGER=NAO`, `PERMITE_SYNC=SIM` |
| PESSOAS | ATUALIZACAO_V2 | `MODO=DRY_RUN`, `PERMITE_SYNC=SIM` |
| PESSOAS | CONFERENCIA_V2 | `MODO=DRY_RUN`, `PERMITE_SYNC=NAO` |
| VIGENCIAS | ATUALIZACAO_V2 | `MODO=DRY_RUN`, `PERMITE_SYNC=SIM` |
| VIGENCIAS | CONFERENCIA_V2 | `MODO=DRY_RUN`, `PERMITE_SYNC=NAO` |
| ATIVIDADES | ATUALIZACAO_PORTAL_V2 | `MODO=DRY_RUN`, `PERMITE_SYNC=SIM` |
| ATIVIDADES | CONFERENCIA_V2 | `MODO=DRY_RUN`, `PERMITE_SYNC=NAO` |
| ATIVIDADES | FREQUENCIA_V2 | `MODO=DRY_RUN`, `PERMITE_SYNC=SIM` |

Todas as linhas criadas recebem:

- `ATIVO=SIM`;
- `PERMITE_TRIGGER=NAO`;
- `PERMITE_EMAIL=NAO`;
- `PERMITE_INBOX=NAO`;
- `PERMITE_DRIVE=NAO`;
- `ALTERADO_POR` com `options.origem`;
- `OBS` indicando que a linha foi criada por bootstrap seguro.

`MODO=DRY_RUN` e intencional: a homologacao pode validar a existencia da
configuracao sem liberar escrita real nem execucao automatica.

## Status operacionais conferidos

A rotina espera uma linha correspondente em `MODULOS_STATUS` para cada par
`MODULO + FLUXO` listado acima.

Quando a criacao e autorizada, as linhas ausentes sao criadas com contadores
zerados:

- `EXECUCOES_24H=0`;
- `BLOQUEIOS_24H=0`;
- `SUCESSOS_24H=0`;
- `ERROS_24H=0`.

Campos de data/status existentes nunca sao alterados.

## Retorno

O envelope retornado inclui:

- `ok`;
- `prontoHomologacao`;
- `dryRun`;
- `criarAusentes`;
- `ambiente`;
- `config.existente`;
- `config.ausente`;
- `config.duplicado`;
- `config.invalido`;
- `config.seriaCriado`;
- `status.existente`;
- `status.ausente`;
- `status.duplicado`;
- `status.invalido`;
- `status.seriaCriado`;
- `criados.config`;
- `criados.status`;
- `totais`;
- `avisos`;
- `erros`.

`prontoHomologacao=true` significa que todas as linhas esperadas existem, nao
ha duplicidades nas linhas esperadas e nao ha invalidos detectados nas abas
operacionais.

## Ordem recomendada para homologacao

1. Rodar `coreV2_runTesteBootstrapDryRun()`.
2. Conferir `config.seriaCriado` e `status.seriaCriado`.
3. Corrigir manualmente qualquer duplicidade ou linha invalida.
4. Se o relatorio estiver limpo exceto por ausentes, rodar:

```javascript
coreV2_bootstrapConfiguracao({
  dryRun: false,
  criarAusentes: true,
  ambiente: 'DEV',
  origem: 'HOMOLOGACAO_MANUAL'
});
```

5. Rodar novamente `coreV2_conferirConfiguracao({ ambiente: 'DEV' })`.
6. Avancar para os testes V2 somente se `prontoHomologacao=true`.

## Bloqueios

Bloqueiam a homologacao:

- aba `MODULOS_CONFIG` ausente;
- aba `MODULOS_STATUS` ausente;
- cabecalhos obrigatorios ausentes;
- linhas duplicadas para o mesmo `MODULO + FLUXO + AMBIENTE` em `MODULOS_CONFIG`;
- linhas duplicadas para o mesmo `MODULO + FLUXO` em `MODULOS_STATUS`;
- valores invalidos em `ATIVO`, `MODO`, `AMBIENTE` ou `PERMITE_*`;
- tentativa de escrita fora de `DEV`.

Alertas aceitaveis temporariamente:

- linhas ausentes reportadas em `dryRun`;
- `MODO=DRY_RUN` nas linhas recem-criadas;
- `PERMITE_TRIGGER=NAO` durante homologacao manual.
