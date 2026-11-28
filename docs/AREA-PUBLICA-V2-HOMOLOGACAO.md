# Área Pública V2 — preparação da homologação

Status atualizado: o solicitante autorizou a homologação controlada exclusivamente em DEV, com políticas e identidades sintéticas temporárias. A composição live, o impedimento atual de App Check e a sequência de retomada estão no [checkpoint de composição DEV](AREA-PUBLICA-V2-COMPOSICAO-DEV.md). Não há autorização de merge na main, PROD, competências permanentes ou migração definitiva. Este documento não concede competência por si só; as referências à autorização posterior abaixo descrevem os gates do plano original, agora sujeitos aos limites expressos dessa autorização DEV.

## Reconciliação e configuração candidata

Branch nos dois repositórios: `feature/area-publica-v1`. As candidatas auditadas `b086d98` (Core) e `4e1f287` (Portal) permanecem ancestrais. A main oficial do Core `64449c5fac329b91c29a00175fbebb13d3d79800` foi incorporada por `8bc4c23`; a main do Portal `a87e5a57ad3b8f4d1469682a0704644548e9583f`, por `08f8db697103e5a0ef9113c1b9cdf29f52a1f169`. Fetch autenticado reconfirmou essas pontas em 27/11/2026 (Cuiabá). Não houve novo merge nesta retomada nem integração de branches paralelas. O relatório local em `backend/functions-node/evidence/reconciliacao-area-publica/` registra os SHAs finais e recibos de validação.

Antes de uma publicação futura, repetir fetch e comparar a candidata à revisão realmente servida em DEV, incluindo overlays, assets, rotas e manifestos. Main e live são referências distintas. Notificações e Eixos em frentes paralelas não foram incorporados: preservar a combinação aprovada que estiver em DEV, reconciliar diferenças legítimas e repetir contratos afetados. Preservar Admissão e as subabas de Governança incorporadas pela main.

## Competências editoriais: proposta sujeita a decisão

Atualmente `PUBLIC_CONTENT_MANAGE` agrega consulta, rascunho, publicação, retirada e reordenação, limitada pelas máscaras `pages`/`timeline`, recursos, política institucional, função vigente, identidade e controle temporal. Elegibilidade de cargo não concede competência por si só. A política aceita os cargos previstos em `members/66_member_contracts.gs`; `editorialControl/current` não substitui essa autoridade.

| Competência proposta | Regra de aplicação futura |
|---|---|
| Consulta administrativa | Leitor sem escrita, com recurso e alcance de rascunhos/histórico explícitos. |
| Preparação de rascunhos | Ações próprias de criação/edição; CAS e política institucional; não modifica a projeção pública. |
| Publicação e despublicação | Ação própria revalidada na prévia e no commit; inclui reordenação que afeta conteúdo já público. Decidir se requer pessoa aprovadora distinta. |
| Delegação | Reutilizar função/política e prova institucional, apenas se o mecanismo suportar o fundamento, concedente, delegado, ações, recursos, prazo, revogação e limite da competência concedente. Não presumir subdelegação. |

Proposta técnica: ampliar o contrato canônico com ações explícitas e negação por padrão, preservando os mecanismos institucionais de autorização. Política antiga sem novo campo não deve virar autorização integral. Não implementar nem conceder a separação antes de decisão específica. O piloto com o código atual exige aprovação do ciclo agregado para identidade exclusivamente sintética; não demonstra segregação redator/publicador.

## Fonte e composição das regras

A autoridade operacional identificada é `projects/geapa-convergencia-dev/releases/cloud.firestore`, cujo ponteiro foi lido antes e depois da obtenção integral de `source.files`: ruleset `263b1e02-e880-4ba9-b7d4-a2f57a4dfda6`, atualizado em `2026-10-26T22:58:12.823937Z`. SHA-256 da fonte: `f7333a2f341c4e5116ebe216cbb57b871eede89c936c7929eaa8188c037ac309`. Comentários históricos dentro da fonte não substituem essa comprovação. `legacy/firestore.rules` do Portal não é a fonte da release.

O snapshot integral fica fora do Git. `scripts/compose-editorial-rules.cjs` em `backend/functions-node` valida sua procedência e hash e gera a composição exclusivamente local. Se o hash mudar, o script interrompe: revisar a nova fonte integral antes de adaptar o compositor. Não há rotina de deploy.

Comando local: `node scripts/compose-editorial-rules.cjs evidence/reconciliacao-area-publica/source.json evidence/reconciliacao-area-publica/firestore.composed.rules`. SHA-256 composto esperado para este snapshot: `2e5bad3b24f4e80a3c605ba3a1432674aa7fc64a96066e2d1c129a62c1572fe1`. O manifesto adjacente registra os hashes dos trechos não alterados. Configurações e logs dos dois ensaios integrais ficam na pasta de evidências; a configuração versionada do emulador continua apontando para o contrato local isolado.

Alterações limitadas aos três matches públicos na raiz:

- `publicPages`: somente get, IDs `home`, `about`, `eixos-tematicos` e `EIXOS_[0-9A-Za-z_-]{1,32}_EIXO_[A-Z0-9_]{1,27}`; base padrão; ausência permitida para fallback; documento existente deve conservar identidade, schema, publicação, saneamento, allowlist de campos e 1–30 seções.
- `publicTimeline`: preservar formato de get e filtros da lista; limitar consulta pública a 101 resultados, conforme contrato de detecção de excesso do Portal.
- `publicDocuments`: preservar formato de get, versão e estado `VIGENTE`; lista filtrada, até 101 resultados. Nenhuma escrita de cliente.

Todo o prefixo e sufixo fora desses blocos é preservado byte a byte, inclusive regras de sessão, titularidade, concessões, admissões, recibos e projeções nos namespaces existentes. Não acrescentar um segundo match permissivo: permissões coincidentes se somam. O namespace histórico `coreDevValidation/ACT_FUNCTIONAL_20261008` mantém seus caminhos públicos já existentes; isso não autoriza leitura de canônicos editoriais nem do namespace `CF_DEV_EDIT_PILOT_V1`. As queries dependem da sanitização do produtor: como na fonte atual, a validação estrutural completa de get não é aplicada como filtro de listagem.

Executar o teste `editorial-composition.test.cjs` com a fonte original e a composta, e `editorial-rules.test.cjs`/`editorial-emulator.test.cjs` com a composição. Evidências necessárias: acessos positivos do titular, negativos do terceiro/anônimo, bloqueio de escrita/eliminação/listagem, permissões públicas preservadas e sucesso do ciclo editorial. Admin SDK do backend não comprova as regras de acesso direto; os testes de cliente usam REST sem privilégios administrativos.

A gramática por geração aceita os identificadores candidatos de Eixos sem conceder ao CMS autoridade sobre eles. O leitor atual da main ainda valida oito IDs históricos; a futura gestão dinâmica precisa atualizar seu próprio produtor/leitor e repetir seus testes quando sua integração for aprovada.

Antes de publicar as regras, inventariar IDs e consumidores efetivamente ativos de `publicPages` em DEV. A fonte anterior aceita IDs arbitrários publicados; a composição restringe aos contratos conhecidos. Esta etapa local não leu os documentos reais para inferir que outros IDs não existem. Se houver outro produtor legítimo, incorporar seu contrato aprovado à composição e testar sua preservação antes do deploy.

## Contrato com Gestão Normativa

`publicDocuments/{id}` é projeção do domínio institucional, nunca cadastro paralelo do CMS. `DOCUMENT_LIST` apenas apresenta documentos. O produtor normativo definitivo pode ser entregue separadamente; páginas e cronologia podem ser homologadas com catálogo vazio ou sintético.

| Responsabilidade | Requisito |
|---|---|
| Identificação e DTO | ID estável ligado à norma canônica; campos públicos `id, schemaVersion, published, sanitized, type, title, summary, status, year, publicUrl, order`; versão `geapa-public-document-v1`; URL HTTPS autorizada. |
| Fonte e publicação | Provar aprovação, vigência, publicidade, custódia e versão liberada. Guardar revisão, fundamento, hash e vínculo ao ato na origem. Não expor campos internos ou dados pessoais dispensáveis. |
| Atualização | Projeção por revisão, idempotente; rejeitar eventos obsoletos e duplicatas. Conferir consistência entre fonte, projeção e arquivo. |
| Retirada | Revogação, substituição, fim de vigência ou retirada de publicidade devem retirar/tornar inacessível a projeção; tratar também o acesso ao arquivo quando exigido. Não basta ocultar um link no CMS. |
| Verificação | Publicação inicial, atualização, replay, concorrência, eventos fora de ordem, retirada e reconciliação; leitura anônima e exclusão de revogadas. Catálogo acima de 100 exige paginação aprovada antes da migração. |

## Plano operacional DEV, após autorização posterior

1. Congelar SHAs e manifestos da combinação aprovada. Registrar versões live de Core, Portal e regras, configuração pública, hashes de overlays privados e rollback. Corrigir divergências legítimas antes de publicar; jamais substituir uma frente ativa por checkpoint anterior.
2. Aprovar piloto sintético: titular elegível, identidade sem competência, terceiro fora da allowlist, caso de função expirada/sessão revogada e cliente anônimo. Fixar responsável, janela, recursos e encerramento. Auth e App Check deverão ser reais em DEV; os emuladores locais não comprovam essa etapa.
3. Confirmar que `canonicalEditorial/home`, `about`, índice/itens de cronologia e respectivas projeções estão disponíveis para ensaio sintético. Os IDs são compartilhados e o leitor usa a raiz da base padrão: preview Hosting não isola os dados. Se houver conteúdo real, escolher estratégia segura antes de escrever; preservar revisões e hashes. Não apagar/recriar pessoas ou políticas reais.
4. Preparar política em `memberAuthzPolicies/{id}` usando contrato canônico, seletor ROLE para a função sintética vigente, `PUBLIC_CONTENT_MANAGE`, máscaras apenas necessárias e recursos `canonicalEditorial/pages` e/ou `canonicalEditorial/timeline`. Registrar fundamento e expiração. Não criar cargos ou permissões reais para viabilizar teste.
5. Preparar `editorialControl/current` fechado: `schemaVersion=editorial-control-v1`, `projectId=geapa-convergencia-dev`, `commandType=editorial.public.manage`, `executor=CF_GEN2_DEV_EDIT_V1`, `enabled=false`, `epoch` e `revision` positivos, `actorUids` e `policyIds` explícitos, `availabilityMode=TIMEBOXED`, `startsAt`, `expiresAt`, `browserEnabled=true`, `appCheck.mode=REQUIRED` e app IDs autorizados. Quota proposta do piloto: `windowMs=3600000`, `maxAdmissions=20`; ajustar ao roteiro aprovado sem exceder os limites do contrato.
6. Preservar `projectionRecoveryControl/current` e sua revisão. Acrescentar exclusivamente UID sintético e rotas `/act/context`, `/admin/conteudo-publico` em `specializedReaders`, com janela e App Check válidos. Examinar o produto entre listas globais de UIDs e rotas para não ampliar acesso indevido. O leitor ainda deve conferir sua política específica. Os comandos `editorialPreview`, `editorialExecute`, `editorialStatus` pertencem ao executor, não à allowlist de rotas de leitura. Usar endpoint DEV e namespace compatíveis com modo `SHARED_DEV`.
7. Ordem após autorização: regras compostas validadas e compatíveis; Core com admissões editoriais fechadas; saúde das rotas existentes; Portal; leitura positiva/negativa; somente então abrir a janela sintética. Identificar versões efetivamente servidas, conferir hashes e disponibilidade dos índices. Reconferir se a release de regras mudou desde o snapshot.
8. Fixtures: Home/Sobre com duas versões, seção pública e seção privada, URLs rejeitáveis, dois marcos de cronologia, norma sintética vigente/revogada ou catálogo vazio. Notificações simuladas. Não enviar comunicação institucional.
9. Aceite: rascunho invisível, prévia sem escrita, publicação exata, atualização/retirada, cronologia e ordem, histórico privado, CAS concorrente, resposta perdida/recibo original, replay sem duplicata, sessão/competência revogada, expiração e limpeza do DOM. Desktop 1280 e mobile 390 px sem overflow. Testar anônimo/titular/terceiro contra regras e preservar navegação de Admissão, Governança, Eixos e Notificações presentes na combinação real.
10. Encerrar: bloquear novas admissões, resolver operações em andamento e recuperar recibos com titular original; expirar/remover apenas concessões e entradas sintéticas próprias por revisão; conferir negação posterior, projeções retiradas e ausência de comunicação real. Manter recibos e auditoria.

## Rollback

Registrar antes do piloto a revisão do Core, versão do Hosting, fonte integral das regras e estados dos alvos sintéticos. Fechar admissões e resolver recibos pendentes. Restaurar primeiro Portal compatível, depois tráfego do Core, preservando recuperação pelo executor original. Reverter regras somente se o ponteiro/conteúdo ainda corresponder ao esperado e nenhuma mudança legítima concorrente for perdida; caso contrário compor reversão mínima sobre a fonte atual. Retirar/restaurar projeções por comando autorizado e revisão esperada. Nunca aplicar limpeza global de emulador em DEV. Conferir acessos, contratos e fechamento de concessões após rollback.

## Obrigatório para piloto versus publicação institucional

Obrigatório para piloto: autorização posterior de PR/DEV, combinação aprovada, comparação live, regras integrais compostas, testes locais, identidades/política/janela sintéticas aprovadas, alvos seguros, Auth/App Check reais, índices, logs/recibos e encerramento verificado.

Somente para publicação institucional: decisão permanente sobre competências e delegação; produtor normativo definitivo; conteúdo revisado por responsáveis; comprovação de direitos, versão, vigência e publicidade; migração e aceite institucional. A ausência do produtor normativo não bloqueia o editor isolado. A segregação redator/publicador não pode ser alegada com o contrato agregado atual.

Migração futura: inventariar fontes canônicas, classificar vigentes/minutas/modelos/histórico, identificar responsável e versão, resolver duplicatas e links, simular diferenças e prévias, migrar lotes idempotentes com CAS, reconciliar contagens/hashes e obter aceite editorial. Normas são migradas exclusivamente pelo domínio institucional. Publicação em PROD exige autorização específica independente.
