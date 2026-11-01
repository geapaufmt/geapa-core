# Core GEAPA: instruções atuais

- Fonte única de regras institucionais, contratos, autorização, domínio e operações sobre dados.
- Frontend, TypeScript, Design System, navegação e Hosting pertencem a geapaufmt/geapa-portal.
- Preservar Auth, App Check, vínculo/função vigente, capability/policy, máscaras, CAS, ownership, idempotência, receipts, leases, auditoria, fences e recuperação.
- Não criar segunda autoridade em Sheets, cache de navegador ou adaptador Node.
- Não inferir vigência de norma, decisão humana, aprovação, vínculo ou permissão.
- Fontes .gs canônicas são reutilizadas pelo bundle Node: editar a fonte listada em core-sources.json, nunca generated/.
- Compilar a partir de checkpoint versionado; testar com Node 22 e fixtures sintéticas. Não criar fatos institucionais em PROD para testes.
- Não versionar overlays nominais, evidências privadas, tokens, credenciais ou dumps de Firestore/Auth.
- Push/PR/merge não autorizam deploy. Exigir autorização de ambiente, manifesto e rollback antes de efeitos remotos.
- Preservar worktrees e checkpoints. Não fazer force-push ou downgrade de executor com operações recuperáveis.
- docs/history e arquivos antigos são referência de compatibilidade; suas descrições Apps Script/Sheets não definem a arquitetura operacional atual.
