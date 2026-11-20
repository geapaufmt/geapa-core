# Reconciliação Gen2 de 20/11/2026

A base oficial do Core foi `54d4eb71927fa2eb5c0de621d304e16b7266142c`. Em 20/11, os cartões de revisões do Cloud Run mostraram 100% de tráfego nas quatro revisões da auditoria. O console do Scheduler mostrou os dois jobs na região southamerica-east1, com expressão `* * * * *` e fuso America/Cuiaba. PROD estava habilitado e DEV pausado.

| Revisão | Domínio original | Candidato original | Commit local de reprodução | Fontes |
|---|---|---|---|---:|
| geapacommand-00053-puc | b2c0f183 | b2c0f183 | e40fce7420b264ccef3b28d8306b9347e05c94ec | 260 |
| geapacommanddevpilot-00167-joc | 9f6ce06e | 7ea23623 | fe3e04b5ec3b455f83b3b3cc86c092d24e6351fd | 262 |
| geapacommandboardconsequences-00003-xox | b73b4de0 | 0a028d18 | c1d792e99625e2e343fffae7675ceefa5b12ecf6 | 258 |
| geapacommanddevpilotboardconsequences-00005-jem | b68eee39 | b68eee39 | bd67b8611a4eddb037dbdbf2dd31bbb11197bf57 | 259 |

| Revisão abreviada | Geração da fonte nos recibos | Digest da imagem observado no Cloud Run (SHA-256) |
|---|---|---|
| 00053-puc | 1795198109168034 | ff01fc224dff3227163ce8b8dbbdcc6658f75720a186b4da5cc8209d949d9d3a |
| 00167-joc | 1795199031542609 | 03ebc297442276467896988877499d4219ec907360b2d7296d1a3aeef4df389b |
| 00003-xox | 1794599456657213 | c1e45c8c078d2de4885f4e7be7d5d3f3758261a3a162f56a9a8ed9943962613a |
| 00005-jem | 1794598735627000 | e6d6c15868a1f1523057d439c00441dbb5141f9cd0f498b7125e977473888f8a |

Os quatro recibos indicam Node.js 22, 512 MiB, 1 CPU e timeout de 120 segundos. Os serviços principais permitem até duas instâncias e concorrência dois. Os workers permitem uma instância e concorrência um. O entry point configurado pelo provedor seleciona o worker apropriado. O comando start do pacote pode citar o executor principal. A geração declarada e a geração resolvida coincidem em cada recibo.

## Reprodução em clone completo

Use Node 22 e Git. Execute as linhas em PowerShell na raiz de um clone completo da candidata, com um destino novo de worktree para cada revisão. O verificador lê os manifestos públicos de `backend/functions-node/releases/20261120` e reconstrói o domínio usando o compilador do próprio repositório. A ordem das entradas, o timestamp DOS, o modo de arquivo e a compressão DEFLATE de nível 9 reproduzem o contêiner ZIP. O último argumento opcional grava o ZIP em um arquivo novo. Os arquivos gerados pertencem ao worktree de ensaio.

```powershell
npm --prefix backend/functions-node ci
foreach ($file in Get-ChildItem backend/functions-node/releases/20261120/*.json) {
    $record = Get-Content -Raw $file | ConvertFrom-Json
    $checkout = Join-Path (Split-Path $PWD) $record.revision
    git worktree add --detach $checkout $record.treeCommit
    if ($LASTEXITCODE) { throw 'Falha no checkout de release' }
    node backend/functions-node/scripts/replay-gen2.cjs $record.revision $checkout
    if ($LASTEXITCODE) { throw 'Falha na reprodução' }
}
```

Cada reprodução verifica as fontes e a ordem de core-sources.json, o hash do domínio, as 53 entradas inventariadas e o SHA-256 do ZIP de 54 entradas, que inclui release-manifest.json. A identidade histórica do commit no manifesto de fontes e os metadados da publicação são reaplicados após compilar. Os registros JSON fornecem os SHAs históricos completos, dependências e hashes por caminho. A imagem OCI foi identificada no console. A reconstrução da imagem também exigiria o ambiente do Cloud Build, a imagem base e a cadeia de buildpacks da publicação.

## Desenvolvimento

Quatro commits lineares preservam as composições recuperadas na ancestralidade local. Um quinto commit repõe a árvore integrada da main. Os commits originais de origem e de empacotamento ficam referenciados pelos SHAs completos nos registros de release e preservados no acervo anterior. A reprodução usa os quatro novos commits autossuficientes. Um futuro PR precisa conservar estes commits no histórico ou em referências duráveis.

Sobre a base integrada, foram recuperadas as quatro operações extraordinárias de tema, as verificações de escopo, o uso do executor compartilhado com recibo, o registro de evidência documental e as regras para conservar aprovação e realização anteriores. O código de governança PIA, os rascunhos, ciclo, frequência, RH e procedimentos permanecem na linha integrada. As diferenças de membros e as concessões sintéticas específicas da publicação DEV ficam reproduzíveis no commit histórico DEV. O escopo desta candidata não transfere essas concessões para o fluxo comum.

A revisão dos contratos do Portal usou sua principal em `8780de7f249552c964269fcb95dcbc3b16029f9f`. Os contratos e catálogos de rotas dos dois ambientes contêm autorizar, revogar, regularizar e registrar título/eixos. O adaptador Node associa os quatro caminhos aos tipos canônicos correspondentes. Os testes recuperados exercitam o domínio, o pipeline e o contrato com dados sintéticos.

A execução de referência passou nos testes unitários e de pipeline. Os 22 testes condicionados à integração com o Firestore Emulator ficaram pendentes. O processo Java recebeu `Permission denied: connect` ao criar sua conexão de loopback. O envio para revisão de integração fica pendente da execução desses testes em ambiente local com loopback habilitado.
