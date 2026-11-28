# Área Pública V2 — composição para homologação DEV

Checkpoint de preparação em 28/11/2026 UTC (27/11 em Cuiabá). Este documento complementa [o plano operacional](AREA-PUBLICA-V2-HOMOLOGACAO.md). A autorização atual permite publicação e piloto sintético exclusivamente em `geapa-convergencia-dev`; não permite merge na main nem publicação em PROD.

## Estado efetivamente publicado, anterior ao piloto

| Item | Referência conferida |
|---|---|
| Core | `75b192bd2cbcc3a28b04e339b456355773801e57`, revisão `geapacommanddevpilot-00185-kap`, 100% do tráfego |
| Fonte imutável Core | `geapaCommandDevPilot/function-source.zip`, geração `1795799616181906`; 54 arquivos de pacote e 262 fontes canônicas conferidos por hash |
| Source digest Core | `2279fa20e09ab591249dd7bada374c60edb874b9a65d3532a8ddfc9228f036bc` |
| Worker de consequências | `geapacommanddevpilotboardconsequences-00005-jem`, preservado |
| Portal | `c15f6622f2382d20148b10829fe1f76077bc232f` |
| Hosting | versão `6d16acf0fbd0670c`, release `1795803784367000`; 149 arquivos da aplicação conferidos por HTTPS contra o manifesto do checkpoint |
| Arquivos reservados | `/__/firebase/init.js` e `/__/firebase/init.json` são respostas geridas pelo Firebase; inventário total de 151 entradas. Não confundir seus hashes de resposta com os blobs do upload. |
| Regras integrais | release `projects/geapa-convergencia-dev/releases/cloud.firestore`, ruleset `263b1e02-e880-4ba9-b7d4-a2f57a4dfda6` |
| SHA-256 das regras | `f7333a2f341c4e5116ebe216cbb57b871eede89c936c7929eaa8188c037ac309` |
| Índices | 64 definições inventariadas; consultas públicas de cronologia/documentos aceitas, sem criação de índices |
| Overlay privado | 27 parâmetros; cinco templates reproduzem os bytes servidos. SHA-256 `0b2d6591ef49fdef81672ac9b392be40928b3be45255242183c83950207b0e88`; valores fora do Git |

O snapshot de configuração completa do Core e os inventários privados estão em evidências locais ignoradas pelo Git. Hash de configurações do Core: `5CA18CC92958C595610113198568BC99D4DB4D2645F22548DA88EF3D07AFDFC0`; worker: `347E0AFFFD36273960097772F729FCD5A162A433F2A1B31578BE3CA5E83CB2F3`. Preservar variáveis, identidade, limites, origem, Auth/App Check e controles das demais frentes. Revalidar a fotografia imediatamente antes de qualquer escrita remota.

## Reconciliação identificada

As candidatas autorizadas `14a0619` (Core) e `08f8db6` (Portal) não continham Notificações já publicada. Substituí-las diretamente causaria regressão. As mains oficiais conferidas são `64449c5fac329b91c29a00175fbebb13d3d79800` e `a87e5a57ad3b8f4d1469682a0704644548e9583f`.

Foram compostos somente os checkpoints imutáveis comprovadamente publicados acima, preservando os dois pais e a ancestralidade. Isso não constitui aprovação funcional ou merge das demais frentes. Não foram incorporadas pontas posteriores ou outras branches em desenvolvimento; suas worktrees não foram modificadas.

- Core de composição: `f85de0fcaa70f4abd3fde73feef83a23c9ea426a`. Conflito em `act-readers.cjs`: soma limitada das rotas editoriais e das duas rotas de Notificações; preservada a transação somente de leitura do leitor de notificações. Teste cobre coexistência no limite, excesso, duplicatas, terceiro e controle fechado.
- Portal de composição: `8df45874081009f996798de7f7ec1540800fdca0`. Conflitos nas configurações especializadas DEV e padrão: união das rotas do CMS e Notificações, preservando endpoints, bloqueios e demais rotas. Sininho, contagem e correção mobile publicada mantidos.
- As próximas frentes devem partir dessa composição após sua publicação e reconferir versões efetivas. Em particular, preservar a capacidade conjunta dos leitores e as duas listas de rotas. Não usar a main isoladamente como descrição do ambiente live.

## Regras e segurança dos alvos

A fonte integral foi relida com identidade DEV e seu ponteiro reconferido. O compositor voltou a produzir SHA-256 `2e5bad3b24f4e80a3c605ba3a1432674aa7fc64a96066e2d1c129a62c1572fe1`. A alteração continua limitada aos três matches públicos da raiz; prefixo/sufixo das demais regras são preservados byte a byte. Nenhuma regra foi publicada nesta preparação.

O inventário encontrou somente os oito IDs históricos de Eixos e `eixos-tematicos` em `publicPages`, todos admitidos pela composição. A gramática futura por geração continua contemplada, sem incorporar automaticamente o produtor/leitor de Eixos ainda em desenvolvimento.

`editorialControl/current`, `canonicalEditorial/home`, `canonicalEditorial/about`, `canonicalEditorial/timeline`, `publicPages/home` e `publicPages/about` estavam ausentes; `publicTimeline` estava vazia. As consultas públicas de `publicDocuments` também retornaram catálogo vazio. Repetir a conferência transacional antes das fixtures; se houver conteúdo institucional concorrente, não sobrescrevê-lo. Um preview do Hosting não isola esses documentos.

`projectionRecoveryControl/current` e a configuração de Notificações foram apenas fotografados. Nenhuma identidade, política, concessão ou fixture remota foi criada. `publicDocuments` permanece projeção institucional, sem cadastro normativo no CMS.

## Impedimento para o aceite publicado

Existe diagnóstico anterior de recusas HTTP 403 na troca do App Check real. O SDK servido `12.18.0` aplica espera de 24 horas após 403/404. O último recibo conhecido encerrou em `2026-11-27T18:28:24.341Z`; o limite conservador registrado é **`2026-11-28T18:28:24.341Z` (28/11 às 14h28 em Cuiabá)**. O `allowRequestsAfter` exato não foi capturado. Não se trata de prova de bloqueio do servidor por duração determinada; o término da espera não garante atestação válida.

Não criar contexto/provedor para contornar a espera, usar debug token, reduzir limiar ou desativar controles. Nesta preparação não houve nova tentativa de atestação. As evidências autenticadas de outras frentes não homologam a Área Pública nem a nova composição.

A publicação coordenada fica retida antes da primeira alteração remota, para que a saúde autenticada da combinação possa ser verificada na mesma janela e o piloto possa ser encerrado. Até lá, permanecem live as referências da tabela, sem necessidade de rollback. Nenhuma concessão da Área Pública foi aberta ou ficou pendente de encerramento.

## Retomada dentro da autorização DEV existente

1. Após o limite conhecido, verificar se há recusa posterior e recalcular a janela quando necessário. Reconferir GitHub, Core, Hosting, regras, índices, configurações e ausência de dados reais nos alvos.
2. Obter Auth e App Check reais em sessão convencional no domínio DEV, registrando somente códigos/horários/resultados sanitizados. Se ocorrer novo 403/404, registrar a nova espera e interromper tentativas. Não pedir senha ou token em chat.
3. Com a sessão utilizável, seguir a sequência do plano: regras compatíveis; Core com admissões fechadas; saúde de Admissão/Governança/Eixos/Notificações; Portal e hashes; leituras; janela sintética limitada, com expiração absoluta registrada.
4. Usar identidades dedicadas e `PUBLIC_CONTENT_MANAGE` somente no piloto. Registrar os deltas mínimos de políticas/leitores por revisão; conferir o produto UID × rota. Não alterar permissões de pessoas reais.
5. Executar todo o aceite publicado, inclusive CAS, recuperação, idempotência, titular/terceiro, revogação, expiração, descarte privado e desktop/mobile. Emuladores e App Check simulado não substituem essa etapa.
6. Fechar admissões, reconciliar recibos, retirar projeções sintéticas e encerrar concessões por revisão, preservando histórico/auditoria. Confirmar negações e saúde final.

Rollback deve usar o Core `00185-kap` e o Hosting `6d16acf0fbd0670c` como baseline desta preparação, após nova conferência de concorrência. Nunca retornar a uma main anterior que elimine Notificações. Reverter apenas deltas próprios e apenas sobre a revisão esperada; preservar mudanças legítimas intervenientes.

O parecer de merge permanece condicionado ao aceite funcional publicado e ao encerramento verificado. Competências permanentes, delegação, produtor normativo, conteúdo institucional e migração definitiva continuam decisões posteriores; não impedem o piloto protegido.
