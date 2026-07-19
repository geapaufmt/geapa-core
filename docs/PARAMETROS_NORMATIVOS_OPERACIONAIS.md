# Parametros normativos operacionais

O Core resolve a key `NORMAS_PARAMETROS_OPERACIONAIS` com ambiente explicito.
Nao existe fallback entre DEV e PROD. A leitura valida `PARAMETRO_ID`,
`TIPO_VALOR`, `VALOR`, `UNIDADE`, `VIGENTE=SIM`, `BASE_LEGAL` e compatibilidade
de `MODULO_SISTEMA`.

`TIPO_VALOR` aceita `NUMERO` e `BOOLEANO`. Linhas numericas legadas sem tipo
continuam inferidas como `NUMERO`; os parametros de prazo conhecidos continuam
exigindo numero positivo e `UNIDADE=DIAS`. Booleanos aceitam somente
`SIM/NAO`, `TRUE/FALSE` ou `1/0`, retornam `true/false` e usam unidade vazia ou
`NAO_APLICAVEL`.

Os valores variaveis nao pertencem ao codigo. O modulo consumidor deve salvar
um snapshot no pedido e reler o parametro antes da decisao. Alteracao de valor
ou unidade sem nova `BASE_LEGAL` e inconsistencia normativa bloqueante.

O cache e somente leitura, dura no maximo cinco minutos e pode ser invalidado
explicitamente por `coreInvalidarCacheParametrosNormativosOperacionais` depois
de uma atualizacao normativa. A invalidacao nao altera Registry ou planilhas.

## Ambientes

Cada consumidor informa `DEV` ou `PROD`. Se atualmente existir somente uma
linha PROD no Registry, HOMOLOG deve permanecer bloqueado ate uma destas opcoes
ser aprovada institucionalmente:

1. cadastrar entrada DEV explicita, apontando para uma fonte DEV auditavel; ou
2. cadastrar a mesma fonte institucional somente leitura duas vezes, uma como
   DEV e outra como PROD, deixando o compartilhamento explicito no Registry.

Nunca copiar valores para o codigo nem consultar PROD como fallback de DEV.

## Ata na decisao final de desligamento

O parametro `DESLIGAMENTO_VOLUNTARIO_DECISAO_FINAL_EXIGE_ATA` e booleano. A
configuracao vigente recomendada usa `VALOR=SIM`, `UNIDADE=NAO_APLICAVEL`,
`BASE_LEGAL=NC01-2025-ART16-IV` e `MODULO_SISTEMA=GEAPA_MEMBROS`.

Uma mudanca futura para `NAO` precisa de nova base legal. O Core rejeita `NAO`
quando a base ainda for `NC01-2025-ART16-IV`, e o consumidor tambem deve
comparar snapshot e regra vigente antes da decisao.

O planejamento aditivo, sempre sem escrita, e retornado por:

```javascript
corePrepararParametrosNormativosTipados({
  ambiente: 'DEV',
  headers: cabecalhosLidos,
  records: registrosLidos
});
```
