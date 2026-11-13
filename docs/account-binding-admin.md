# Associação administrativa de Conta Google — candidata local

## Limite desta implementação

A candidata implementa somente uma primeira associação. Não cria usuários Auth, não substitui vínculos ativos ou revogados, não altera dados cadastrais ou competências e não habilita ambiente algum. `PEOPLE_DIRECTORY_READ` continua independente de `MEMBERS_BINDING_MANAGE`. As entradas genéricas `issueBrowser`, endpoint Apps Script e contrato de comandos do Portal continuam bloqueando `members.binding.associate`.

## Caminho de execução

`member-binding-admin.cjs` autentica com `verifyIdToken(token, true)`, exige App Check e verifica os controles do executor. O Core (`113_member_binding_admin.gs`) identifica o administrador pelas fontes canônicas e exige concessão `EXPLICIT` vigente de `MEMBERS_BINDING_MANAGE`, com pessoa e escopo `ADMINISTRACAO` expressamente autorizados. Não há fallback para DEV_ADMIN ou autorização por menu.

Somente depois dessa autorização o adaptador consulta `auth.getUserByEmail`. Exige conta ativa, e-mail verificado, provedor `google.com` com o mesmo e-mail e ausência de tenant. O UID é derivado pelo servidor e não integra nenhum contrato de entrada ou resposta do navegador. A consulta do e-mail de uma conta já vinculada também reconcilia o UID retornado pelo Auth.

O administrador informa CPF ou RGA conferido, uma referência documental de até 48 caracteres (`A-Z`, `a-z`, números, `_`, `-`) e atesta expressamente que verificou o documento e a titularidade da conta com a pessoa. O Core compara o valor com um identificador canônico ativo. Não se trata de prova automática de titularidade civil pelo Google: a diligência documental é uma responsabilidade humana explícita do administrador autorizado. Sem identificador canônico correspondente ou sem essa conferência, o fluxo permanece fechado.

`memberBindingVerifications/{operationId}` é uma proposta técnica privada e imutável, válida por até cinco minutos. Guarda conta verificada, pessoa/revisão, administrador e referência/revisão do identificador, mas não copia CPF/RGA. Não é uma segunda autoridade de vínculo. `memberAuthBindings` permanece canônico.

Depois da confirmação administrativa separada, o Core produz as evidências existentes `memberEvidenceContents` / `memberEvidenceAttestations`, `memberPreparedIntents` e `devProtectedInputs` (nome histórico do port protegido também usado operacionalmente). A chamada interna `admission.issue` recebe o UID e as revisões exclusivamente desses registros. O executor original conserva ownership, ticket, CAS, lease/fencing, ledger, recibo e `memberAudit`; a família `79_member_commands_binding.gs` continua responsável pela mudança. `portalUsers` e `portalAccessStatus` são produzidos pelo planejador canônico na mesma transação.

O novo escopo fechado `ACCOUNT_BINDING` só funciona com capability `MEMBERS_BINDING_MANAGE` e concessão `EXPLICIT`. O alvo precisa constar em `allowedPersons`; os três recursos do comando precisam coincidir com o UID da proposta protegida do mesmo administrador e caso. Ele não transforma diretório, ROLE ou prefixos arbitrários em permissão de escrita.

## Conflitos e recuperação

- O UID não pode ter registro anterior, nem inativo, para esta primeira associação.
- Uma consulta transacional limitada a duas linhas verifica qualquer vínculo anterior da pessoa, incluindo revogados. Usa igualdade em `idPessoa` e ordenação por ID de documento; não requer nova coleção de acervo nem varredura completa. O índice de campo único de `idPessoa` precisa estar disponível em DEV.
- Ingresso ativo em `memberEnrollmentIndexes` bloqueia a associação paralela.
- Pessoa elegível: identidade ativa, não técnica, não bloqueada e estado canônico `MEMBRO_EFETIVO`, `MEMBRO_INGRESSANTE` ou `EGRESSO`.
- Confirmação/admissão/finalização revalidam fontes, revisão, concessão, conflitos, validade da prova e conta Google. Nada sobrescreve um vínculo existente.
- Após resposta desconhecida, `memberBindingStatus` consulta a operação. `memberBindingRecover` só retoma uma operação já admitida, no executor original. Recibos terminais não dependem de nova prova da conta-alvo para serem recuperados; a identidade e autorização atuais do administrador continuam exigidas.
- O navegador persiste somente `operationId`, em chave derivada de projeto + administrador + pessoa. Não guarda e-mail, CPF/RGA, UID do alvo, credenciais ou tokens nesse journal. Mudança de sessão invalida respostas antigas.

## Habilitação futura em DEV — não executada nesta campanha

1. Revalidar versões publicadas e incorporar somente a candidata aprovada; não sobrescrever mudanças concorrentes.
2. Revisar o perfil de runtime DEV e a permissão de leitura de usuários Auth do service account. Não conceder escrita em Auth. Se faltar permissão de leitura, tratar isso como impedimento de homologação.
3. Manter `GEAPA_MEMBERS_ENABLED=true`, App Check `REQUIRED`, origem e app IDs válidos; os controles do executor de Membros já utilizados continuam obrigatórios.
4. Habilitar, mediante aprovação específica, `memberBindingAdminControl/current` (`schemaVersion: member-binding-admin-control-v1`, `projectId`, `revision>0`, `enabled`, `grantIds`). Não há configuração automática ou concessão incluída no pacote.
5. Selecionar uma concessão canônica EXPLICIT de administrador legítimo, com capability `MEMBERS_BINDING_MANAGE`, escopo `ADMINISTRACAO`, `allowedPersons` restrito às pessoas sintéticas, `commandScopes: [ACCOUNT_BINDING]` e máscara vazia. Conservar validade, issuer/approval, bindingVersion, subjectEpoch e demais contratos de concessão. Não derivar isso da permissão de visualizar Pessoas.
6. Configurar a rota `MEMBERS_BINDING_ASSOCIATE`, escopo sintético e controles/budgets do executor original. Se `commandTypes` estiver definido no controle, incluir explicitamente `members.binding.associate`. Preservar todas as outras rotas, IDs e operações recuperáveis.
7. Usar Conta Google sintética já autenticada, e-mail verificado e pessoa sintética com CPF/RGA canônico. Homologar pela interface, com Auth/App Check reais; depois conferir vínculo, projeções, recibo, auditoria, reload e recuperação. Não usar a conta ou o cadastro real mencionado pelo usuário.

Nenhum desses passos autoriza PROD. Uma futura promoção depende de autorização específica, integridade, controle de versões e prova DEV real. Para interromper novas associações, fechar admissão/rota conforme protocolo do executor, preservando recuperação; não apagar vínculos, evidências, ledger ou recibos como rollback de código.

## Testes

`backend/functions-node/test/binding-admin-emulator.test.cjs` inclui validação do port Admin SDK e matriz no Firestore Emulator. Execução integrada requer `GEAPA_BINDING_EMULATOR_TEST=1`, Node 22 e emulador dedicado `127.0.0.1:8187`, projeto exclusivamente `demo-geapa-functions-local`. O teste limpa apenas esse banco emulado entre cenários. Auth e App Check são ports sintéticos nesse teste; isto não equivale a homologação de Google em DEV.
