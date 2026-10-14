# Piloto DEV: disponibilidade HTTP

Somente Node Function de healthcheck, sem SDK Firebase, segredo, comando ou
acesso a dados. Não é o executor definitivo e não usa a pasta do protótipo B0
como backend. O Portal publicado permanece no Apps Script.

- Root Directory: `pilots/fast-node-dev`.
- Destino: projeto separado **DEV**, plano Hobby, deployment **Preview**.
- Região solicitada: `gru1`; duração máxima de 10 segundos.
- Rota: `GET /api/health`; `HEAD` permitido; demais métodos negados.
- Deploy no ambiente `production` falha fechado na função.
- `firstInvocation` identifica a primeira chamada observada naquela instância;
  sozinho não mede cold start da plataforma. Medir o tempo HTTP no cliente.
- Sem build, dependências, cron, configuração de domínio ou IAM.

Teste: `node --test health.test.cjs` nesta pasta.

A conexão Vercel → Firebase depende de investigação e autorização separadas
para OIDC/WIF/IAM. Nunca colocar chave JSON ou grant OAuth local neste piloto.

O código está preparado; a publicação depende do acesso ao workspace Vercel
correto. A criação não pode mudar plano ou habilitar cobrança.
