# Acesso administrativo técnico DEV

`coreDevAdministrativeGrants/{uid}` é uma concessão explícita do Core, independente de `coreDevAdmins`, cargos e mandatos. Ela vincula UID, pessoa, versão do binding e epoch do sujeito. É ignorada antes da leitura em PROD, expira em até 30 dias e tem revisão própria. Nenhuma concessão surge de um perfil ADMIN ou de hints do Portal.

As capabilities são independentes. `PEOPLE_DIRECTORY_READ` expõe somente as pessoas sintéticas nominadas e suas máscaras. `MEMBERS_PROFILE_CORRECT` consulta e opera somente essas fixtures, pelo pipeline canônico de preparação, admissão, execução, CAS, evidências, receipts e auditoria. A revelação do dossiê também gera auditoria. O proof `DEV_TECHNICAL` mantém a revisão e a referência da concessão.

`INSTITUTIONAL_DIRECTORY_READ` permite consultar as seções expressamente listadas. `INSTITUTIONAL_GOVERNANCE_DRAFT` permite criar e editar somente rascunhos sintéticos da própria identidade. O servidor marca sua autoria técnica. Não admite assinaturas, aprovação, publicação, procedimentos de mandato, anexos, referências institucionais ou transição para revisão. Contatos, parâmetros, normas e demais operações continuam com suas políticas independentes. O controle de cada módulo continua obrigatório.

O Portal existente consome `canReadPeople`, `canReviewCorrections` e `admin-navigation-hints-v1` do Core. `director` continua derivado exclusivamente das funções canônicas. Renovar o acesso relê os contratos; hints vencem em no máximo 60 segundos. Cada consulta e operação revalida a concessão; o leitor confere alterações concorrentes antes de devolver dados.

## Administração oficial

O cadastro de pessoa aprovada pode ser incluído explicitamente com `registration: true` e `expectedIngressRevision` no JSON administrativo. Emite uma aprovação sintética nominal e uma terceira pessoa sem vínculo. A capability `MEMBERS_INGRESS_EXECUTE`, com máscara `preRegister`, aceita somente `PRE_REGISTER` para essa pessoa e os dados exatos da aprovação. Não autoriza aprovação, efetivação, envio de convite, associação de conta ou registro de ingresso real. A listagem de cadastros usa apenas os índices das fixtures autorizadas.

Quando a rota de ingresso já está compartilhada, seu executor, epoch, atores e recibos são preservados. `devActorWindow` mantém o prazo anterior dos demais atores e estende somente o UID solicitado. Preparação, admissão e recuperação revalidam esse prazo no Core; a extensão é recusada em PROD.

Executar `backend/functions-node/scripts/administer-dev-access.cjs` com identidade IAM autorizada no projeto DEV, usando o Admin SDK oficial. O comando não aceita outro projeto. O JSON nominal deve ficar fora do Git, com `uid`, `personId`, `operator`, `approvalRef`, `days`, `expectedGrantRevision`, `expectedRouteRevision` e `expectedGovernanceRevision`. O token OAuth da sessão administrativa entra somente pela variável `GEAPA_ADMIN_ACCESS_TOKEN`, nunca por arquivo versionado, argumento ou log.

Sem `--apply`, o comando produz apenas o plano. Com `--apply`, uma transação verifica as revisões, o binding, as fixtures, a ausência de operações pendentes na janela compartilhada vencida e a exclusividade da janela reaberta. Cria duas pessoas sintéticas dedicadas e solicitações de correção somente na primeira emissão. Renova apenas o UID solicitado, mantém os dados das fixtures nas renovações, incrementa a geração das projeções e grava `coreDevAdministrativeAudit`. Não reativa o antigo `FUNCTIONAL_ALL`, não altera policies institucionais, Auth, cargos ou mandatos. O plano recusa um controle compartilhado ainda ativo para evitar retirar acesso concorrente.

`--revoke --apply` desativa a concessão com CAS, auditoria e invalidação de projeções; as fixtures e receipts ficam preservados. Uma concessão vencida exige nova emissão administrativa explícita; a renovação da sessão não prolonga privilégios. Não há renovação automática indefinida.

Validar com Node 22, `npm ci`, build, testes e empacotamento local. `GEAPA_DEV_ACCESS_EMULATOR_TEST=1` executa consultas, recusa de outro UID, revogação concorrente, revelação auditada, análise e aplicação de correção, criação e edição de rascunho no Firestore Emulator. Esse teste usa Auth/App Check simulados exclusivamente no emulador; o aceite DEV exige sessão real no Hosting publicado. A publicação deve preservar as mudanças live e o executor proprietário dos receipts. PROD exige autorização própria e não recebe estas concessões.
