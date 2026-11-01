# GEAPA Core canônico

Regras institucionais, contratos, autorização e domínio do GEAPA, executados em Cloud Functions Gen2 com Firestore e Cloud Storage. Frontend oficial: https://github.com/geapaufmt/geapa-portal. Sheets/Apps Script são compatibilidade e recuperação histórica onde ainda exigidas; não são autoridade operacional concorrente.

## Fontes e build

backend/functions-node/core-sources.json relaciona 230 fontes canônicas .gs. São JavaScript de domínio reutilizado, não uma segunda implementação em Node. O adaptador SDK está em backend/functions-node/src. O bundle gerado nunca deve ser editado manualmente.

Usar Node 22. Dentro de backend/functions-node, executar em sequência:

    npm ci
    npm run build
    npm test
    npm run verify:published
    npm run release -- --prod

A geração exige fontes canônicas idênticas ao commit. Release é empacotamento local e exige checkpoint limpo; não faz deploy. O CI não possui credenciais ou publicação. Os testes usam identidades e fixtures sintéticas, sem operações remotas.

O frontend antes mantido em diretórios de integração locais deve ser alterado exclusivamente no repositório Portal. O backend não depende de arquivos externos ao checkout.

## Publicação e recuperação

Os registros releases/published-*-20261101.json contêm hashes dos pacotes realmente publicados. verify:published reconstrói o código e, após verificar todas as fontes e hashes, reaplica metadados históricos para comparação exata. Isso não autoriza reaplicar release antigo.

Nova publicação exige PR revisado, testes, pacote com hashes, configuração privada revisada, autorização explícita e rollback compatível com operações admitidas. Nunca substituir executor por versão incapaz de recuperar seus receipts/ownership.

Arquivos raiz antigos e docs/history preservam a biblioteca Apps Script e sua documentação histórica. Não executar jobs, migrações, e-mails ou operações institucionais para comprovar build.

Leia [convergência e limites](docs/SOURCE-CONVERGENCE.md).
