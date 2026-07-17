# Parametros normativos operacionais

O Core resolve a key `NORMAS_PARAMETROS_OPERACIONAIS` com ambiente explicito.
Nao existe fallback entre DEV e PROD. A leitura valida `PARAMETRO_ID`, `VALOR`
positivo, `UNIDADE=DIAS`, `VIGENTE=SIM`, `BASE_LEGAL` e compatibilidade de
`MODULO_SISTEMA`.

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
