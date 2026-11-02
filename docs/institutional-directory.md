# Gestão Institucional

O leitor `/admin/gestao-institucional` consulta quatro documentos limitados (`identity`, `parameters`, `norms`, `governance`). Fontes de Sheets são evidência de reconciliação offline; nunca são consultadas pela Function.

A capability `INSTITUTIONAL_DIRECTORY_READ` exige policy ROLE explícita, titular nominal, função vigente, recurso e máscaras próprios. Não utiliza DEV_ADMIN nem empresta `ACTIVITY_BOARD_READ` para documentos institucionais. Novo controle e policies ausentes significam acesso fechado. Auth e App Check continuam no adaptador compartilhado; controle, policy, vínculo e documento são revalidados antes da resposta.

Todos os registros importáveis desta campanha possuem natureza `DOCUMENTARY_ONLY`. Estado na fonte não é ato de aprovação; `canonical` apenas aponta para autoridade já existente. Convites reservados não integram o DTO. Pesquisa e paginação são no Core, limitadas a 150 registros por seção e 10 por página.

`prepareContact` é uma preparação pura, sem admission/endpoint de escrita. Restringe seis campos, exige motivo, revisão e hash. Não altera dados, identidade oficial, cores ou normas. A futura execução exige autorização própria, receipt, CAS, auditoria e revalidação; não pode executar diretamente este retorno.

`publicFields` é uma projeção pura com allowlist e aprovação vinculada à revisão/hash exatos. Não publica nada por efeito da importação e não inclui convites. A integração editorial com Sobre/Contato/rodapé permanece fechada enquanto não houver campos aprovados.

O compilador `backend/functions-node/preparation/institutional/compile-documentary.cjs <diretório privado de evidências>` consome extração OOXML e snapshot canônico, prepara quatro documentos e detecta duplicatas. Preserva revisões das fontes e não executa escritas. Manifestos nominais, fontes reais e capturas ficam fora do repositório.

A importação PROD exige aprovação do manifesto fechado. Execução futura deve verificar ausência dos documentos novos e updateTime dos controles; rollback fecha só o leitor novo e reverte apenas documentos que ainda coincidam com o hash/revisão aplicados. Nunca restaurar snapshot sobre alterações posteriores. Não alterar normas canônicas, policies existentes, vínculos ou fatos de ACT/Membros.

## Edição interna de contatos (candidata, depende de ativação)

O comando `institutional.contact.update` reaproveita `prepareContact`, a identificação institucional, o contrato de policies ROLE, o controle canônico de disponibilidade e os planos transacionais compartilhados. Atua somente no campo Valor de seis registros editoriais de identity. Normas, parâmetros, atas, proveniência, classificação e visibilidade não são editáveis por esse comando.

A capability distinta `INSTITUTIONAL_CONTACT_EDIT` é admitida apenas para Diretor Administrativo/ADMINISTRACAO, com policy nominal, recurso e máscaras explícitos. O leitor não herda edição. Controle ausente/desabilitado impede novas alterações. A candidata não concede direitos por si; a abertura PROD exige manifesto aprovado.

Cada alteração é uma única transação Firestore: identidade documental com revisão + contador horário por ator + ownership + recibo + auditoria. Não há efeito externo nem trabalho assíncrono; por isso não se cria lease ou ticket fictício. O ownership usa commandExecutorOwners, com schema e tipo próprios; não é reinterpretado pelos executores ACT/Membros. O adaptador Node é o único executor desse comando. Status apenas lê um recibo vinculado ao usuário, digest e executor e não reproduz efeitos. A ausência de recibo significa não admitido atomicamente. Depois de fechar a admissão, status/replay terminal permanecem possíveis enquanto a autorização institucional continuar válida.

A confirmação exige hash e revisão apresentados, epoch/revisão do controle e prazo curto. Auth revogado, App Check inválido, função encerrada, policy revogada e máscara incompatível impedem o commit. Uma atualização concorrente invalida a revisão; não existe merge silencioso. A política limita até 20 alterações por hora/ator, configurável para menos. A consulta usa o documento canônico, sem rebuild global; o adaptador de leitura verifica seus fences antes de devolver a resposta.

A publicação pública continua fechada e não faz parte dos efeitos. A proveniência original permanece intacta; o evento auditado registra valor anterior, novo valor, motivo, pessoa, policy/revisão e hashes antes/depois. Aprovações públicas antigas vinculadas ao hash/revisão não autorizam o novo conteúdo.

Rollback operacional: desabilitar apenas institutionalEditControl/current com CAS; preservar executor, ownership, recibos e auditoria. Nunca restaurar todo identity sobre alterações legítimas posteriores. Reversão de conteúdo exige nova intenção humana explícita, valor anterior documentado e revisão atual. Não apagar operações para desfazer uma edição.
