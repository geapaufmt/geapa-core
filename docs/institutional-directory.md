# Gestão Institucional

O leitor `/admin/gestao-institucional` consulta quatro documentos limitados (`identity`, `parameters`, `norms`, `governance`). Fontes de Sheets são evidência de reconciliação offline; nunca são consultadas pela Function.

A capability `INSTITUTIONAL_DIRECTORY_READ` exige policy ROLE explícita, titular nominal, função vigente, recurso e máscaras próprios. Não utiliza DEV_ADMIN nem empresta `ACTIVITY_BOARD_READ` para documentos institucionais. Novo controle e policies ausentes significam acesso fechado. Auth e App Check continuam no adaptador compartilhado; controle, policy, vínculo e documento são revalidados antes da resposta.

Todos os registros importáveis desta campanha possuem natureza `DOCUMENTARY_ONLY`. Estado na fonte não é ato de aprovação; `canonical` apenas aponta para autoridade já existente. Convites reservados não integram o DTO. Pesquisa e paginação são no Core, limitadas a 150 registros por seção e 10 por página.

`prepareContact` é uma preparação pura, sem admission/endpoint de escrita. Restringe seis campos, exige motivo, revisão e hash. Não altera dados, identidade oficial, cores ou normas. A futura execução exige autorização própria, receipt, CAS, auditoria e revalidação; não pode executar diretamente este retorno.

`publicFields` é uma projeção pura com allowlist e aprovação vinculada à revisão/hash exatos. Não publica nada por efeito da importação e não inclui convites. A integração editorial com Sobre/Contato/rodapé permanece fechada enquanto não houver campos aprovados.

O compilador `backend/functions-node/preparation/institutional/compile-documentary.cjs <diretório privado de evidências>` consome extração OOXML e snapshot canônico, prepara quatro documentos e detecta duplicatas. Preserva revisões das fontes e não executa escritas. Manifestos nominais, fontes reais e capturas ficam fora do repositório.

A importação PROD exige aprovação do manifesto fechado. Execução futura deve verificar ausência dos documentos novos e updateTime dos controles; rollback fecha só o leitor novo e reverte apenas documentos que ainda coincidam com o hash/revisão aplicados. Nunca restaurar snapshot sobre alterações posteriores. Não alterar normas canônicas, policies existentes, vínculos ou fatos de ACT/Membros.
