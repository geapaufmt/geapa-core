# Registro administrativo de governança

O contrato `institutional.governance.draft.save` representa rascunhos administrativos. Não aprova deliberações, não estabelece vigência e não publica conteúdo. Os nove registros documentais históricos continuam na consulta existente; o contrato não os sobrescreve.

## Autoridade e invariantes

- Fonte: `institutional/206_institutional_governance.gs`, incorporada ao bundle Node pelo manifesto de fontes.
- Agregado: `institutionalGovernanceDrafts/current`; histórico e auditoria registram revisões e valores anteriores/posteriores.
- Admissão: `institutionalGovernanceControl/current`, inicialmente fechada na ausência de configuração. Exige Auth, App Check, vínculo/função vigente, policy, recurso e máscara `draft` para `INSTITUTIONAL_GOVERNANCE_DRAFT`.
- A capacidade de registro administrativo não confere aprovação formal. A elegibilidade estrutural da Administração não cria grants produtivos; a liberação nominal exige deliberação e manifesto próprios.
- Prévia curta vinculada à revisão, hash do agregado, referências e controle. Execução atômica com ownership, recibo, auditoria, limite por hora e comparação de revisões. Nenhum efeito externo é disparado.
- Repetição de operação confirmada retorna o recibo. Resposta perdida deve ser recuperada por status, nunca por nova intenção automática. Recuperação continua exigindo identidade proprietária e autorização vigente.
- Data desconhecida permanece nula. São aceitos somente `DRAFT`, `UNDER_REVIEW` e `EVIDENCE_PENDING`. Aprovação, publicação e vigência não são campos editáveis.

## Referências documentais

Uma referência declarada fica `PENDING_VERIFICATION`. Somente relação explícita no registro documental canônico permite `DOCUMENTED_SOURCE_LINK`. Isso comprova o vínculo documental, não aprovação ou vigência. O hash informado identifica a cópia conferida pelo responsável; o contrato não declara ter baixado e certificado o conteúdo do Drive.

O leitor exige as máscaras de resumo e campos da seção de governança. Fontes, referências e revisão participam dos controles transacionais. Registros públicos e fontes documentais originais nunca integram o conjunto de escrita.

## Parâmetro piloto

Para `ATIVIDADES_DESTACAR_PROXIMA`, ciência exige pessoa diferente e a outra função entre Presidência e Administração. Ciência não ativa a mudança. A ativação explícita revalida revisões, controles e as autorizações dos participantes. Propostas vencidas ou de outra revisão do controle permanecem no histórico, sem admissão de novas etapas.

## Validação e publicação

Os testes de emulador cobrem concorrência, resposta perdida, repetição, revogação, função encerrada, App Check, referências, publicação bloqueada e leitura canônica após alteração. Evidências autenticadas, policies de fixtures e recibos permanecem no workspace privado.

Esta implementação foi publicada somente em DEV. Publicação PROD e grants produtivos requerem autorização específica. Rollback deve primeiro fechar admissões; preservar recibos e agregados que tenham recebido operações posteriores, sem restaurar snapshots sobre esses fatos.
