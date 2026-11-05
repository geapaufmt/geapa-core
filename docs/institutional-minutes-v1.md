# Atas da Diretoria — V1

Estado: ciclo documental validado em DEV com identidades sintéticas. Nenhuma ativação produtiva está implícita nesta implementação.

## Autoridade e contrato

`institutional/207_institutional_minutes.gs` amplia o contrato existente de `206_institutional_governance.gs`. O adaptador Node `src/institutional-governance.cjs` conserva a execução transacional; não há uma API paralela de importação ou assinatura.

O Portal utiliza `/admin/gestao-institucional/governanca/editar`, com fases `preview`, `execute` e `status`, e o executor `institutional.governance.draft.save`. COMPOSITION é uma prévia somente leitura. A criação mantém `INSTITUTIONAL_GOVERNANCE_DRAFT`; edição/submissão, manifestação e arquivo usam respectivamente `INSTITUTIONAL_MINUTES_EDIT`, `INSTITUTIONAL_MINUTES_MANIFEST` e `INSTITUTIONAL_MINUTES_ARCHIVE`.

`institutionalGovernanceControl/current` exige `minutesV1Enabled=true`, ações explicitamente admitidas, atores, policies e vigência técnica. Publicar o código não habilita as atas. Nenhuma conta técnica adquire competência institucional.

## Composição, versões e manifestações

- A composição vem de funções, vínculo e gestão canônicos na data da reunião, com snapshot e revisões. Todos os diretores encontrados devem constar como presentes ou ausentes.
- O ator deve continuar com função e policy vigentes na operação. A composição histórica não estende autorização atual.
- Estados: `DRAFT`, `IN_REVIEW`, `CORRECTIONS_REQUESTED`, `ARCHIVED`.
- Cada revisão de conteúdo preserva a versão anterior e zera as manifestações da nova versão. Aprovações não são transportadas silenciosamente.
- Presente: `APPROVE_SIGN`. Ausente: `ACKNOWLEDGE`. Ambos podem registrar observação ou solicitar correção. Ausente não assina presença.
- As manifestações vinculam pessoa autenticada, funções, timestamp real, revisão e hash exatos. A assinatura é confirmação eletrônica interna, não certificação ICP-Brasil nem voto favorável às deliberações.
- Arquivar requer todas as manifestações obrigatórias, ausência de correção pendente e composição íntegra. O documento arquivado é imutável. Retificação é outro registro, com referência validada ao original.

## Persistência e recuperação

Reutiliza agregado e histórico de governança, controle de taxa, ownership, receipt e auditoria: seis escritas atômicas. Prévia determinística, revisão esperada, controle de recurso e autorização são conferidos novamente no commit. Resposta perdida é reconciliada pelo mesmo operationId; replay concluído não gera efeito adicional.

O leitor aplica escopo BOARD ou COMPOSITION e retorna apenas ações calculadas pelo Core. Não inclui aprovações fictícias em documentos históricos. Uma alteração documental não gera atividade, frequência, notificação, sanção, alteração de parâmetro ou publicação pública.

## Limites e teste

Limites herdados do agregado: 100 registros, 500 kB serializados, 500 eventos. Até 20 versões anteriores por ata, 20 pessoas na composição e 80 manifestações por versão. Controle temporário de homologação: 20 admissões/hora por ator. O sistema rejeita excesso; não trunca o histórico. Escala superior exige evolução específica de armazenamento.

Teste dirigido: `backend/functions-node/test/institutional-minutes-emulator.test.cjs`, com `GEAPA_INSTITUTIONAL_EMULATOR_TEST=1` e emulador Firestore. Executar sequencialmente com o teste de governança: as fixtures compartilham o namespace do emulador. Abrange correção, nova versão, assinatura, ciência, arquivo, retificação, duplicidade, ausência de diretor, revisão obsoleta, revogação e recuperação após resposta perdida.

A homologação física DEV percorreu oito operações pelo Portal normal e duas contas sintéticas legítimas, com Auth/App Check ativos. Controles temporários e função sintética adicional foram encerrados, preservando evidências e recibos. As evidências nominais ficam fora do Git.

## Publicação futura

PROD permanece fechado para esta V1. Exige manifesto próprio com atores/cargos, policies, recursos, controles, versões e aprovação explícita. Não restaurar agregados antigos sobre operações posteriores. Em contingência, fechar novas admissões e preservar o executor para status/recuperação; qualquer reversão documental depende das revisões atuais.
