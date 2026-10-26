# Diretrizes Permanentes do Sistema GEAPA

Este repositorio faz parte do Sistema GEAPA. Ao atuar aqui, priorize estabilidade, desempenho e seguranca antes de ampliar funcionalidades.

## Prioridades

- Priorize performance antes de novas funcionalidades.
- Prefira melhorias que reduzam tempo de carregamento, chamadas ao backend e leituras repetidas.
- Nao altere producao, publique versoes ou execute pushes/deploys sem instrucao explicita.
- Documente mudancas relevantes em README, comentarios operacionais ou notas de migracao quando afetarem contratos, dados, permissoes ou performance.

## Arquitetura

- GitHub Pages e o front-end.
- Apps Script e o backend/API.
- Google Sheets e banco interno do sistema, nao interface principal para usuarios finais.
- O front-end pode orientar navegacao e experiencia, mas permissoes e autorizacoes criticas devem ser validadas no backend.

## Performance

Metas de resposta percebida:

- Pagina inicial: 1,5 a 3 s.
- Aba Atividades: 1 a 2,5 s.
- Detalhes de atividade: 0,5 a 1,5 s.
- Troca entre telas ja carregadas: instantanea ou quase instantanea.

Diretrizes praticas:

- Use views `PORTAL_*` para leitura rapida sempre que existirem.
- Evite cruzamentos pesados em tempo real no Apps Script.
- Nao faca leituras amplas de multiplas abas para cada interacao do usuario se uma view consolidada ou cache resolver.
- Reduza chamadas ao backend agrupando dados necessarios por tela.
- Use cache quando fizer sentido, especialmente para dados de leitura frequente e baixa volatilidade.
- Prefira respostas enxutas, com apenas os campos necessarios para o front.

## Apps Script

- Use leitura e escrita em lote (`getValues`, `setValues`, ranges agregados) em vez de loops com chamadas por celula.
- Em escritas sensiveis, use `LockService` para evitar concorrencia e duplicidade.
- Separe leitura de escrita; nao misture consultas de tela com mutacoes.
- Nao implemente autorizacao critica somente no front-end.
- Retorne erros controlados, sem expor IDs internos de planilhas, tokens, chaves ou dados de terceiros.

## Views e Dados

- Nao remova views existentes sem migracao documentada.
- Antes de substituir uma view, registre o contrato antigo, o contrato novo e os consumidores impactados.
- Prefira criar ou evoluir views `PORTAL_*` para telas do portal em vez de fazer cruzamentos pesados no momento da requisicao.
- Evite retornar listas completas quando a tela precisa apenas de um registro, resumo ou subconjunto paginado.

## Front-end do Portal

- Trate dados vindos do Apps Script como contratos de API.
- Evite chamadas redundantes ao backend ao trocar entre telas ja carregadas.
- Reaproveite dados em memoria quando forem validos para a sessao/tela.
- Use permissoes retornadas pelo backend para montar navegacao, mas mantenha validacao real no backend.

## Seguranca Operacional

- Nao exponha e-mails ou dados de terceiros sem necessidade do caso de uso.
- Nao use dados mockados para logica oficial.
- Nao escreva em planilhas oficiais em testes sem instrucao explicita.
- Prefira funcoes internas com sufixo `_` e wrappers publicos claros para contratos consumidos por outros modulos.
