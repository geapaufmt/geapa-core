# Contratos recuperados da suíte raiz

Execute `npm ci`, `npm run build` e `node --test ../../tests/*.test.cjs` em `backend/functions-node` com Node 22. As quatro suítes exercitam a VM do executor e o bundle construído de `core-sources.json`. A compatibilidade instrumentada de membros usa o controle sintético da fixture existente e a barreira real. O executor Node continua com os serviços externos bloqueados.

O índice abaixo usa a numeração dos casos na [base 7019032](https://github.com/geapaufmt/geapa-core/tree/701903224bc97951cd02a13eab10310542193d92/tests). Para recuperar uma suíte anterior, use `git show 7019032:tests/<arquivo>.test.cjs`.

| Suíte original e assertivas | Garantia exercida e destino |
| --- | --- |
| core_domain_resolver, 1–2 e 9–10 | Na mesma suíte, resolve explicitamente DEV/PROD, bloqueia duplicidade ativa, ambiente inválido e ausência da chave no ambiente pedido. |
| core_domain_resolver, 3–7 | A origem e o aviso de fallback são testados em leitura. A escrita recusa DB ausente, aba canônica ausente e divergência, inclusive ao abrir o domínio inteiro. |
| core_domain_resolver, 8 e 11–12 | Mantém abertura única, nomes e chave de solicitações de vínculo. Verifica os dois IDs e os três identificadores de resolvedor obsoleto em todas as fontes do manifesto. Exige a presença do resolvedor antes da varredura. |
| core_member_profile_v2, 1–5 | Mesmas assertivas contra o catálogo do bundle, origem nacional e estrangeira, inconsistência de UF/nome e cursos inexistentes, inativos ou OUTRO sem descrição. |
| core_member_profile_v2, 6–8 e 12 | Conserva os valores do período/RGA, semestre ordinal, intervalos e rejeição de RGA inválido. Mantém completude parcial sem percentual e os três nomes de schema/Registry ainda expostos. |
| core_member_profile_v2, 9–11 | Simular abas para conferir token DEV, idempotência de cabeçalhos/linhas e erro específico do instalador PROD era aceite da preparação de uma planilha. A entrada atual, sem autoridade LEGACY instalada, deve recusar até o dry-run, antes de consultar uma planilha, em DEV e PROD. O teste exige o código dessa barreira. Para reproduzir a ferramenta histórica em modo LEGACY, recuperar os casos 9–11 na base. |
| core_normative_parameters, 1–7 | Mesmos tipos, unidades e valores. Abrange número implícito/expresso, seis representações de booleano, booleano ambíguo e tipo desconhecido. |
| core_normative_parameters, 8–14 | Preserva os códigos de erro para parâmetro ausente, vigência duplicada/inativa, Registry duplicado, base legal ausente/incompatível, módulo divergente e tentativa de cruzar ambientes. |
| core_normative_parameters, 15–16 | Mantém leitura em cache antes de invalidar, contagem de três chaves, nova leitura após invalidação e planejamento puro de campos/linha do tipo booleano sem escrita. |
| core_portal_session_environment, seis assertivas de contagem e ambiente | Mesmas contagens e rótulos. Usa o mesmo ID físico sintético nos dois ambientes e exige conteúdo distinto, mesmo após aquecer o cache DEV. Verifica recusa sem a barreira configurada. |
| core_portal_session_environment, inspeções de sessão, fallback e assinatura do export | O teste chama o export real, passa ambas as grafias de sessão, exige a identidade devolvida e o erro de fallback desabilitado. A segunda resolução da sessão provoca falha. |
| core_portal_session_environment, inspeções do Registry e da chave de cache | Compara as chaves reais e exige recusa nos dois sentidos se a configuração existir apenas no outro ambiente. |
| core_portal_session_environment, inspeções de paginação e de RESUMO | A leitura de situação solicita exclusivamente RESUMO. A expressão textual de slice da listagem antiga foi substituída pela paginação do leitor atual em `backend/functions-node/test/people-directory.test.cjs`, que verifica limite, ordenação, páginas disjuntas e mudança de snapshot. |

A CI executa a suíte raiz após gerar o domínio. As suítes transacionais usam o projeto descartável do emulador, em série, com limpeza do banco padrão do emulador local entre arquivos para evitar contaminação de coleções. A suíte do pacote preserva seu resultado independente.
