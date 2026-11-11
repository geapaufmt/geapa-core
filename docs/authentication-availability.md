# Disponibilidade do verificador de autenticação

O adaptador Node distingue falhas conhecidas de credencial (`AUTH_BINDING_INVALID`, HTTP 401) de indisponibilidade do SDK/verificador (`SERVICE_UNAVAILABLE`, HTTP 503). Ambos recusam a requisição; nenhuma operação protegida é admitida sem verificação válida.

`verifyIdToken(token, true)` continua obrigatório. Permanecem as verificações de assinatura, projeto, emissor, UID, expiração, data de autenticação e tenant. Auth, App Check, autorizações institucionais, policies, domínio e Rules não foram alterados.

O SDK instalado pode encapsular falha na obtenção dos certificados como `auth/argument-error`. A mensagem interna conhecida é utilizada apenas para classificação e nunca devolvida ou registrada com dados privados. Erros desconhecidos falham de forma fechada como indisponibilidade; não são aceitos como autenticação.

Regressões: `backend/functions-node/test/auth-availability.test.cjs`, suíte de segurança existente e emuladores institucionais. As fixtures de revogação usam o código real do Admin SDK, em lugar de um `Error('revoked')` genérico. A fixture de atas contém a projeção de vínculo sintético necessária para testar a negativa de assinatura do ausente, sem eliminar a validação canônica do vínculo.

Executar com Node 22. Para os emuladores locais, usar projeto `demo-geapa-functions-local`; os testes não comprovam reCAPTCHA Enterprise real. Caso a CLI tente gravar seu arquivo de preferências fora do workspace ao encerrar, executar com `XDG_CONFIG_HOME` apontando para um diretório local de teste, sem copiar credenciais ou alterar a configuração do usuário.

A alteração não foi implantada nesta campanha. A homologação autenticada e a publicação precisam de autorização específica. Recuperação produtiva deve preservar receipts, journals, configurações da Function e executores atuais; não restaurar dados institucionais.
