# Leitura pessoal de conta vinculada em DEV

O leitor canônico revalida `coreDevLinkedReaderGrants/{uid}` para a homologação de contas vinculadas às fixtures `DEVADM_*_1` ou `DEVADM_*_2`. A janela dura no máximo quatro horas, fixa pessoa e versão do binding e exige origem `DEV_ADMIN_SYNTHETIC_V1`. A troca ou revogação do vínculo interrompe o acesso.

O emissor administrativo `backend/functions-node/scripts/administer-dev-linked-reader.cjs` exporta `administer`: recebe Admin SDK autorizado, projeto DEV, UID, pessoa, referência de autorização, operador e revisão esperada. O padrão é dry-run; `apply` grava concessão e auditoria na mesma transação. `revoke` encerra a concessão com CAS. Dados nominais e credenciais permanecem fora do Git.

A concessão admite apenas o transporte das consultas pessoais enumeradas em `dev-linked-reader-scope.cjs`. Não concede capability, cargo, acesso administrativo, mutação ou recuperação genérica de projeções. O controle do serviço, Auth, App Check e produtores canônicos continuam obrigatórios. PROD ignora a concessão sem sequer lê-la. Expiração não autoriza renovação automática.

As quatro consultas da agenda de membros também exigem `ACTIVITY_AGENDA_READ_SELF` e preservam os filtros canônicos de publicação e visibilidade. A concessão não admite a agenda administrativa, chamada ou modelos de atividades.
