/**
 * ============================================================
 * 17_core_mail_hub.js
 * ============================================================
 *
 * Mail Hub V1
 *
 * Objetivo desta camada:
 * - ler mensagens do Gmail;
 * - registrar eventos em planilhas centrais;
 * - deduplicar por Id Mensagem Gmail;
 * - manter indice por chave de correlacao;
 * - registrar anexos recebidos;
 * - expor consultas minimas para posterior consumo por modulos.
 *
 * Fora de escopo nesta rodada:
 * - migracao completa dos modulos consumidores;
 * - retry avancado da fila de envio;
 * - salvamento de anexos no Drive;
 * - roteamento complexo por regras.
 */

var CORE_MAIL_HUB_KEYS = Object.freeze({
  EVENTOS: 'MAIL_EVENTOS',
  INDICE: 'MAIL_INDICE',
  SAIDA: 'MAIL_SAIDA',
  ANEXOS: 'MAIL_ANEXOS',
  REGRAS: 'MAIL_REGRAS',
  CONFIG: 'MAIL_CONFIG'
});

var CORE_MAIL_HUB_STATUS = Object.freeze({
  PENDENTE: 'PENDENTE',
  PROCESSADO: 'PROCESSADO',
  IGNORADO: 'IGNORADO'
});

var CORE_MAIL_ATTACHMENT_STATUS = Object.freeze({
  PENDENTE: 'PENDENTE',
  PROCESSADO: 'PROCESSADO',
  SALVO_DRIVE: 'SALVO_DRIVE',
  IGNORADO: 'IGNORADO',
  ERRO: 'ERRO'
});

var CORE_MAIL_HUB_DEFAULTS = Object.freeze({
  query: 'in:inbox',
  maxThreads: 25,
  maxMessagesPerThread: 100,
  start: 0
});

var CORE_MAIL_HUB_UX = Object.freeze({
  defaultDataRowHeight: 28,
  cacheKey: 'CORE_MAIL_HUB_OPERATIONAL_SHEET_UX_V2',
  cacheTtlSeconds: 21600,
  colors: Object.freeze({
    identity: '#d9ead3',
    routing: '#fff2cc',
    body: '#fce5cd',
    status: '#ead1dc',
    operational: '#d0e0e3',
    neutralText: '#202124'
  }),
  sheetRules: Object.freeze({
    MAIL_EVENTOS: Object.freeze({
      notes: Object.freeze({
        'Id Evento': 'Identificador unico do evento registrado no Mail Hub. Ex.: MEV-19d53c96101e99bd.',
        'Data Hora Evento': 'Data e hora do evento no Gmail ou na fila central.',
        'Direcao': 'Direcao da mensagem no ecossistema. Valores usuais: ENTRADA ou SAIDA.',
        'Tipo Evento': 'Tipo tecnico do evento. Ex.: EMAIL_RECEBIDO ou EMAIL_ENVIADO.',
        'Modulo Dono': 'Modulo de negocio responsavel pela conversa. Ex.: MEMBROS.',
        'Tipo Entidade': 'Tipo da entidade de negocio associada. Ex.: MEMBRO.',
        'Id Entidade': 'Identificador da entidade relacionada. Ex.: RGA ou ID interno.',
        'Chave de Correlacao': 'Chave tecnica usada para ligar eventos, indice e assunto final. Ex.: MEM-2026-202311801013-CNV-CAND.',
        'Etapa Fluxo': 'Etapa funcional da conversa quando conhecida. Ex.: CAND.',
        'Id Thread Gmail': 'ID da thread do Gmail associada ao evento.',
        'Id Mensagem Gmail': 'ID unico da mensagem no Gmail.',
        'Id Mensagem Pai': 'ID da mensagem anterior, quando o fluxo conseguir inferir essa relacao.',
        'Assunto': 'Assunto registrado no evento. Em saidas, ja vem com [GEAPA][CHAVE].',
        'Email Remetente': 'Endereco principal de quem enviou a mensagem.',
        'Nome Remetente': 'Nome legivel do remetente, quando disponivel.',
        'Emails Destinatarios': 'Lista de destinatarios principais separados por virgula.',
        'Emails Cc': 'Lista de emails em copia.',
        'Emails Cco': 'Lista de emails em copia oculta.',
        'Trecho Corpo': 'Trecho curto para leitura rapida no log. O conteudo completo pode estar em Corpo Texto.',
        'Corpo Texto': 'Corpo textual completo salvo para auditoria. O valor continua inteiro mesmo quando a celula fica em modo compacto.',
        'Possui Anexos': 'Indica se o evento registrou anexos relevantes. Valores usuais: SIM ou NAO.',
        'Quantidade Anexos': 'Quantidade de anexos considerados no evento.',
        'Nomes Anexos': 'Lista resumida dos nomes dos anexos associados ao evento.',
        'Status Roteamento': 'Resultado do roteamento tecnico. Ex.: ROTEADO, NAO_IDENTIFICADO, IGNORADO ou ERRO.',
        'Status Processamento': 'Estado do tratamento funcional do evento. Ex.: PENDENTE, PROCESSADO, IGNORADO ou ERRO.',
        'Processado Por': 'Funcao ou modulo que marcou o evento como processado.',
        'Data Hora Processamento': 'Momento em que o evento foi processado ou ignorado.',
        'Observacoes': 'Campo tecnico para contexto adicional. Ex.: NOISE_REASON=ASSUNTO_TECNICO.',
        'Json Bruto': 'Resumo serializado do contexto tecnico do evento para auditoria e debug.',
        'Criado Em': 'Momento em que a linha foi criada na central.',
        'Atualizado Em': 'Momento da ultima atualizacao da linha.'
      }),
      colorGroups: Object.freeze([
        Object.freeze({ color: '#d9ead3', headers: ['Id Evento', 'Modulo Dono', 'Tipo Entidade', 'Id Entidade', 'Chave de Correlacao', 'Etapa Fluxo'] }),
        Object.freeze({ color: '#d0e0e3', headers: ['Data Hora Evento', 'Id Thread Gmail', 'Id Mensagem Gmail', 'Id Mensagem Pai', 'Criado Em', 'Atualizado Em'] }),
        Object.freeze({ color: '#fff2cc', headers: ['Direcao', 'Tipo Evento', 'Status Roteamento', 'Status Processamento', 'Processado Por', 'Data Hora Processamento'] }),
        Object.freeze({ color: '#fce5cd', headers: ['Assunto', 'Email Remetente', 'Nome Remetente', 'Emails Destinatarios', 'Emails Cc', 'Emails Cco', 'Trecho Corpo', 'Corpo Texto', 'Observacoes', 'Json Bruto'] }),
        Object.freeze({ color: '#ead1dc', headers: ['Possui Anexos', 'Quantidade Anexos', 'Nomes Anexos'] })
      ]),
      dropdownRules: Object.freeze({
        'Direcao': Object.freeze({
          values: Object.freeze(['ENTRADA', 'SAIDA']),
          allowInvalid: true,
          helpText: 'Use ENTRADA para mensagens recebidas e SAIDA para envios registrados.'
        }),
        'Tipo Evento': Object.freeze({
          values: Object.freeze(['EMAIL_RECEBIDO', 'EMAIL_ENVIADO']),
          allowInvalid: true,
          helpText: 'Tipos principais de evento na central de e-mails.'
        }),
        'Possui Anexos': Object.freeze({
          values: Object.freeze(['SIM', 'NAO']),
          allowInvalid: true,
          helpText: 'Indique se o evento possui anexos relevantes.'
        }),
        'Status Roteamento': Object.freeze({
          values: Object.freeze(['PENDENTE', 'ROTEADO', 'NAO_IDENTIFICADO', 'IGNORADO', 'ERRO']),
          allowInvalid: true,
          helpText: 'Resultado do roteamento tecnico do evento.'
        }),
        'Status Processamento': Object.freeze({
          values: Object.freeze(['PENDENTE', 'PROCESSADO', 'IGNORADO', 'ERRO']),
          allowInvalid: true,
          helpText: 'Estado do processamento funcional do evento.'
        })
      }),
      clipHeaders: Object.freeze([
        'Trecho Corpo',
        'Corpo Texto',
        'Json Bruto',
        'Observacoes'
      ]),
      dataRowHeight: 28
    }),
    MAIL_INDICE: Object.freeze({
      notes: Object.freeze({
        'Chave de Correlacao': 'Chave tecnica que consolida a conversa inteira no indice.',
        'Prefixo Correlacao': 'Prefixo ou familia da chave para agrupamentos operacionais. Ex.: MEM.',
        'Modulo Dono': 'Modulo principal responsavel pela conversa.',
        'Tipo Entidade': 'Tipo da entidade associada ao historico. Ex.: MEMBRO.',
        'Id Entidade': 'Identificador da entidade resolvida para a conversa.',
        'Etapa Atual': 'Etapa funcional mais recente inferida pelo Mail Hub.',
        'Id Thread Gmail': 'Thread Gmail mais recente associada a esta correlacao.',
        'Id Ultima Mensagem': 'Mensagem mais recente associada a esta correlacao.',
        'Ultima Direcao': 'Direcao do ultimo evento da conversa. Ex.: ENTRADA ou SAIDA.',
        'Ultimo Tipo Evento': 'Tipo do ultimo evento registrado.',
        'Ultimo Email Remetente': 'Email remetente do ultimo evento da conversa.',
        'Ultimo Assunto': 'Assunto mais recente conhecido para a conversacao.',
        'Data Hora Ultimo Evento': 'Momento do evento mais recente desta chave.',
        'Ha Entrada Pendente': 'Indica se ainda existe entrada aguardando tratamento. Valores usuais: SIM ou NAO.',
        'Ha Anexo Pendente': 'Indica se ainda existe anexo pendente associado a essa chave.',
        'Quantidade Eventos': 'Total de eventos registrados para a chave.',
        'Quantidade Entradas': 'Quantidade de eventos de ENTRADA.',
        'Quantidade Saidas': 'Quantidade de eventos de SAIDA.',
        'Quantidade Anexos': 'Quantidade total de anexos associados a essa conversa.',
        'Status Conversa': 'Resumo semantico da conversa. Ex.: PENDENTE, CONCLUIDA, IGNORADO ou ERRO.',
        'Criado Em': 'Momento da primeira ocorrencia conhecida da chave.',
        'Atualizado Em': 'Momento da ultima recomposicao do indice.'
      }),
      colorGroups: Object.freeze([
        Object.freeze({ color: '#d9ead3', headers: ['Chave de Correlacao', 'Prefixo Correlacao', 'Modulo Dono', 'Tipo Entidade', 'Id Entidade', 'Etapa Atual'] }),
        Object.freeze({ color: '#d0e0e3', headers: ['Id Thread Gmail', 'Id Ultima Mensagem', 'Ultima Direcao', 'Ultimo Tipo Evento', 'Ultimo Email Remetente', 'Ultimo Assunto', 'Data Hora Ultimo Evento'] }),
        Object.freeze({ color: '#fff2cc', headers: ['Ha Entrada Pendente', 'Ha Anexo Pendente', 'Status Conversa'] }),
        Object.freeze({ color: '#ead1dc', headers: ['Quantidade Eventos', 'Quantidade Entradas', 'Quantidade Saidas', 'Quantidade Anexos'] }),
        Object.freeze({ color: '#fce5cd', headers: ['Criado Em', 'Atualizado Em'] })
      ]),
      dropdownRules: Object.freeze({
        'Ha Entrada Pendente': Object.freeze({
          values: Object.freeze(['SIM', 'NAO']),
          allowInvalid: true,
          helpText: 'Mostra se ainda existe mensagem de entrada pendente.'
        }),
        'Ha Anexo Pendente': Object.freeze({
          values: Object.freeze(['SIM', 'NAO']),
          allowInvalid: true,
          helpText: 'Mostra se ainda existe anexo aguardando tratamento.'
        }),
        'Status Conversa': Object.freeze({
          values: Object.freeze(['PENDENTE', 'CONCLUIDA', 'IGNORADO', 'ERRO']),
          allowInvalid: true,
          helpText: 'Resumo funcional da situacao da conversa.'
        })
      }),
      clipHeaders: Object.freeze(['Ultimo Assunto']),
      dataRowHeight: 28
    }),
    MAIL_ANEXOS: Object.freeze({
      notes: Object.freeze({
        'Id Anexo': 'Identificador unico do anexo registrado no Mail Hub.',
        'Id Evento': 'Evento de e-mail ao qual este anexo pertence.',
        'Modulo Dono': 'Modulo dono da conversa associada.',
        'Tipo Entidade': 'Tipo da entidade relacionada ao anexo. Ex.: MEMBRO.',
        'Id Entidade': 'Identificador da entidade associada ao anexo.',
        'Chave de Correlacao': 'Chave tecnica da conversa que originou o anexo.',
        'Etapa Fluxo': 'Etapa funcional associada ao anexo, quando conhecida.',
        'Id Mensagem Gmail': 'Mensagem Gmail que carregou o anexo.',
        'Id Thread Gmail': 'Thread Gmail relacionada.',
        'Indice Anexo Mensagem': 'Posicao do anexo dentro da mensagem original, usada para reabrir o blob real sob demanda.',
        'Nome Arquivo': 'Nome original do arquivo.',
        'Tipo Mime': 'Tipo MIME do arquivo. Ex.: application/pdf.',
        'Tamanho Bytes': 'Tamanho aproximado do anexo em bytes.',
        'Foi Salvo No Drive': 'Indica se o anexo ja foi persistido no Drive. Valores usuais: SIM ou NAO.',
        'Id Arquivo Drive': 'ID do arquivo salvo no Drive, quando aplicavel.',
        'Link Arquivo Drive': 'URL do arquivo salvo no Drive, quando aplicavel.',
        'Pasta Destino Drive': 'Pasta de destino usada no Drive, quando aplicavel.',
        'Status Anexo': 'Estado operacional do anexo. Ex.: PENDENTE, PROCESSADO, SALVO_DRIVE, IGNORADO ou ERRO.',
        'Processado Por': 'Funcao ou modulo que marcou o estado atual do anexo.',
        'Data Hora Processamento': 'Momento em que o anexo foi tratado, quando aplicavel.',
        'Observacoes': 'Campo tecnico para notas de processamento do anexo.',
        'Criado Em': 'Momento do registro do anexo na central.',
        'Atualizado Em': 'Momento da ultima atualizacao da linha.'
      }),
      colorGroups: Object.freeze([
        Object.freeze({ color: '#d9ead3', headers: ['Id Anexo', 'Id Evento', 'Modulo Dono', 'Tipo Entidade', 'Id Entidade', 'Chave de Correlacao', 'Etapa Fluxo'] }),
        Object.freeze({ color: '#d0e0e3', headers: ['Id Mensagem Gmail', 'Id Thread Gmail', 'Indice Anexo Mensagem', 'Nome Arquivo', 'Tipo Mime', 'Tamanho Bytes'] }),
        Object.freeze({ color: '#fff2cc', headers: ['Status Anexo', 'Foi Salvo No Drive', 'Processado Por', 'Data Hora Processamento'] }),
        Object.freeze({ color: '#ead1dc', headers: ['Id Arquivo Drive', 'Link Arquivo Drive', 'Pasta Destino Drive'] }),
        Object.freeze({ color: '#fce5cd', headers: ['Observacoes', 'Criado Em', 'Atualizado Em'] })
      ]),
      dropdownRules: Object.freeze({
        'Status Anexo': Object.freeze({
          values: Object.freeze(['PENDENTE', 'PROCESSADO', 'SALVO_DRIVE', 'IGNORADO', 'ERRO']),
          allowInvalid: true,
          helpText: 'Estado atual do anexo dentro do fluxo operacional.'
        }),
        'Foi Salvo No Drive': Object.freeze({
          values: Object.freeze(['SIM', 'NAO']),
          allowInvalid: true,
          helpText: 'Indique se o anexo ja foi salvo no Drive pelo modulo consumidor.'
        })
      }),
      clipHeaders: Object.freeze(['Link Arquivo Drive', 'Pasta Destino Drive', 'Observacoes']),
      dataRowHeight: 28
    }),
    MAIL_CONFIG: Object.freeze({
      notes: Object.freeze({
        'Chave': 'Nome da configuracao operacional lida pelo Mail Hub. Ex.: USAR_SOMENTE_ASSUNTOS_GEAPA.',
        'Valor': 'Valor bruto da configuracao. Pode ser texto, numero, regex ou lista separada por linha, virgula ou ponto e virgula.',
        'Ativo': 'Se SIM, a configuracao entra em vigor. Se NAO, a linha e ignorada.'
      }),
      colorGroups: Object.freeze([
        Object.freeze({ color: '#d9ead3', headers: ['Chave'] }),
        Object.freeze({ color: '#fce5cd', headers: ['Valor'] }),
        Object.freeze({ color: '#fff2cc', headers: ['Ativo'] })
      ]),
      dropdownRules: Object.freeze({
        'Ativo': Object.freeze({
          values: Object.freeze(['SIM', 'NAO']),
          allowInvalid: true,
          helpText: 'Ative ou desative a linha de configuracao.'
        })
      }),
      clipHeaders: Object.freeze(['Valor']),
      dataRowHeight: 28
    }),
    MAIL_REGRAS: Object.freeze({
      notes: Object.freeze({
        'Id Regra': 'Identificador unico da regra de roteamento. Ex.: REG-0001.',
        'Ativa': 'Se SIM, a regra pode ser aplicada. Se NAO, a linha e ignorada.',
        'Ordem': 'Prioridade numerica. Menor numero e avaliado primeiro.',
        'Campo Analise': 'Campo da mensagem analisado. V1: ASSUNTO, REMETENTE, DESTINATARIO, CORPO ou TUDO.',
        'Tipo Comparacao': 'Comparacao aplicada. V1: CONTEM, IGUAL, COMECA_COM, TERMINA_COM ou REGEX.',
        'Valor Comparacao': 'Texto ou regex comparado contra o campo escolhido.',
        'Modulo Dono': 'Modulo que deve receber a mensagem quando a regra bater.',
        'Tipo Entidade': 'Tipo de entidade associada ao roteamento.',
        'Etapa Fluxo': 'Etapa funcional opcional registrada em MAIL_EVENTOS e MAIL_INDICE.',
        'Acao Quando Bater': 'Acao executada quando a regra casar. V1 operacional: ROTEAR.',
        'Observacoes': 'Contexto humano da regra.',
        'Criado Em': 'Data/hora de criacao da regra.',
        'Atualizado Em': 'Data/hora da ultima atualizacao da regra.'
      }),
      colorGroups: Object.freeze([
        Object.freeze({ color: '#d9ead3', headers: ['Id Regra', 'Ativa', 'Ordem'] }),
        Object.freeze({ color: '#fff2cc', headers: ['Campo Analise', 'Tipo Comparacao', 'Valor Comparacao', 'Acao Quando Bater'] }),
        Object.freeze({ color: '#d0e0e3', headers: ['Modulo Dono', 'Tipo Entidade', 'Etapa Fluxo'] }),
        Object.freeze({ color: '#fce5cd', headers: ['Observacoes', 'Criado Em', 'Atualizado Em'] })
      ]),
      dropdownRules: Object.freeze({
        'Ativa': Object.freeze({
          values: Object.freeze(['SIM', 'NAO']),
          allowInvalid: true,
          helpText: 'Se SIM, a regra entra em vigor.'
        }),
        'Campo Analise': Object.freeze({
          values: Object.freeze(['ASSUNTO', 'REMETENTE', 'DESTINATARIO', 'CORPO', 'TUDO']),
          allowInvalid: true,
          helpText: 'Campo da mensagem analisado.'
        }),
        'Tipo Comparacao': Object.freeze({
          values: Object.freeze(['CONTEM', 'IGUAL', 'COMECA_COM', 'TERMINA_COM', 'REGEX']),
          allowInvalid: true,
          helpText: 'Tipo de comparacao aplicada.'
        }),
        'Modulo Dono': Object.freeze({
          values: Object.freeze(['APRESENTACOES', 'ATIVIDADES', 'SELETIVO', 'MEMBROS', 'COMUNICACOES', 'DESLIGAMENTOS']),
          allowInvalid: true,
          helpText: 'Modulo dono quando a regra bater.'
        }),
        'Acao Quando Bater': Object.freeze({
          values: Object.freeze(['ROTEAR']),
          allowInvalid: true,
          helpText: 'V1: use ROTEAR.'
        })
      }),
      clipHeaders: Object.freeze(['Valor Comparacao', 'Observacoes']),
      dataRowHeight: 28
    }),
    MAIL_SAIDA: Object.freeze({
      notes: Object.freeze({
        'Id Saida': 'Identificador unico da linha de outbox na central.',
        'Modulo Dono': 'Modulo de negocio que solicitou o envio.',
        'Tipo Entidade': 'Tipo da entidade associada ao envio. Ex.: MEMBRO.',
        'Id Entidade': 'Identificador da entidade relacionada ao envio.',
        'Chave de Correlacao': 'Chave tecnica usada no assunto final e na rastreabilidade.',
        'Etapa Fluxo': 'Etapa funcional do envio. Ex.: CAND.',
        'Email Destinatario Principal': 'Envelope principal usado no envio tecnico.',
        'Emails Destinatarios': 'Lista de destinatarios principais separados por virgula.',
        'Emails Cc': 'Lista de destinatarios em copia.',
        'Emails Cco': 'Lista de destinatarios em copia oculta.',
        'Nome Destinatario': 'Nome legivel do destinatario principal, quando houver.',
        'Assunto': 'Assunto final ja montado, incluindo [GEAPA][CHAVE].',
        'Corpo Texto': 'Versao textual do email enviada ou preparada para envio.',
        'Corpo Html': 'HTML final institucional enviado ou pronto para envio.',
        'Data Hora Agendada': 'Momento minimo para processamento da saida.',
        'Prioridade': 'Prioridade operacional da fila. Ex.: BAIXA, MEDIA, NORMAL ou ALTA.',
        'Status Envio': 'Estado tecnico da saida. Ex.: PENDENTE, ENVIADO ou ERRO.',
        'Tentativas': 'Quantidade de tentativas de processamento da linha.',
        'Ultimo Erro': 'Ultimo erro registrado no processamento da outbox.',
        'Id Thread Gmail': 'Thread gerada ou reutilizada pelo envio.',
        'Id Mensagem Gmail': 'Mensagem Gmail gerada pelo envio.',
        'Enviado Em': 'Momento em que o envio tecnico foi concluido.',
        'Criado Em': 'Momento em que a linha entrou na fila central.',
        'Atualizado Em': 'Momento da ultima atualizacao da linha.',
        'Observacoes': 'Contrato serializado e metadata operacional da saida.'
      }),
      colorGroups: Object.freeze([
        Object.freeze({ color: '#d9ead3', headers: ['Id Saida', 'Modulo Dono', 'Tipo Entidade', 'Id Entidade', 'Chave de Correlacao', 'Etapa Fluxo'] }),
        Object.freeze({ color: '#d0e0e3', headers: ['Email Destinatario Principal', 'Emails Destinatarios', 'Emails Cc', 'Emails Cco', 'Nome Destinatario'] }),
        Object.freeze({ color: '#fce5cd', headers: ['Assunto', 'Corpo Texto', 'Corpo Html', 'Observacoes'] }),
        Object.freeze({ color: '#fff2cc', headers: ['Data Hora Agendada', 'Prioridade', 'Status Envio', 'Tentativas', 'Ultimo Erro'] }),
        Object.freeze({ color: '#ead1dc', headers: ['Id Thread Gmail', 'Id Mensagem Gmail', 'Enviado Em', 'Criado Em', 'Atualizado Em'] })
      ]),
      dropdownRules: Object.freeze({
        'Prioridade': Object.freeze({
          values: Object.freeze(['BAIXA', 'MEDIA', 'NORMAL', 'ALTA']),
          allowInvalid: true,
          helpText: 'Prioridade usada para organizacao operacional da fila.'
        }),
        'Status Envio': Object.freeze({
          values: Object.freeze(['PENDENTE', 'ENVIADO', 'ERRO', 'CANCELADO']),
          allowInvalid: true,
          helpText: 'Estado tecnico atual da saida.'
        })
      }),
      clipHeaders: Object.freeze([
        'Corpo Texto',
        'Corpo Html',
        'Observacoes',
        'Ultimo Erro'
      ]),
      dataRowHeight: 28
    })
  })
});

var CORE_MAIL_HUB_SCHEMA = Object.freeze({
  MAIL_EVENTOS: Object.freeze([
    'Id Evento',
    'Data Hora Evento',
    'Direcao',
    'Tipo Evento',
    'Modulo Dono',
    'Chave de Correlacao',
    'Id Thread Gmail',
    'Id Mensagem Gmail',
    'Assunto',
    'Email Remetente',
    'Emails Destinatarios',
    'Status Processamento',
    'Processado Por',
    'Data Hora Processamento',
    'Possui Anexos',
    'Quantidade Anexos',
    'Criado Em',
    'Atualizado Em'
  ]),
  MAIL_INDICE: Object.freeze([
    'Chave de Correlacao',
    'Modulo Dono',
    'Tipo Entidade',
    'Id Entidade',
    'Etapa Atual',
    'Id Thread Gmail',
    'Id Ultima Mensagem',
    'Ultima Direcao',
    'Ultimo Tipo Evento',
    'Ultimo Email Remetente',
    'Ultimo Assunto',
    'Data Hora Ultimo Evento',
    'Ha Entrada Pendente',
    'Ha Anexo Pendente',
    'Quantidade Eventos',
    'Quantidade Entradas',
    'Quantidade Saidas',
    'Quantidade Anexos',
    'Criado Em',
    'Atualizado Em'
  ]),
  MAIL_ANEXOS: Object.freeze([
    'Id Anexo',
    'Id Evento',
    'Modulo Dono',
    'Tipo Entidade',
    'Id Entidade',
    'Chave de Correlacao',
    'Etapa Fluxo',
    'Id Mensagem Gmail',
    'Id Thread Gmail',
    'Indice Anexo Mensagem',
    'Nome Arquivo',
    'Tipo Mime',
    'Tamanho Bytes',
    'Foi Salvo No Drive',
    'Id Arquivo Drive',
    'Link Arquivo Drive',
    'Pasta Destino Drive',
    'Status Anexo',
    'Processado Por',
    'Data Hora Processamento',
    'Observacoes',
    'Criado Em',
    'Atualizado Em'
  ]),
  MAIL_CONFIG: Object.freeze([
    'Chave',
    'Valor',
    'Ativo'
  ]),
  MAIL_REGRAS: Object.freeze([
    'Id Regra',
    'Ativa',
    'Ordem',
    'Campo Analise',
    'Tipo Comparacao',
    'Valor Comparacao',
    'Modulo Dono',
    'Tipo Entidade',
    'Etapa Fluxo',
    'Acao Quando Bater',
    'Observacoes',
    'Criado Em',
    'Atualizado Em'
  ]),
  MAIL_SAIDA: Object.freeze([
    'Id Saida',
    'Modulo Dono',
    'Tipo Entidade',
    'Id Entidade',
    'Chave de Correlacao',
    'Etapa Fluxo',
    'Email Destinatario Principal',
    'Emails Destinatarios',
    'Emails Cc',
    'Emails Cco',
    'Nome Destinatario',
    'Assunto',
    'Corpo Texto',
    'Corpo Html',
    'Data Hora Agendada',
    'Prioridade',
    'Status Envio',
    'Tentativas',
    'Ultimo Erro',
    'Id Thread Gmail',
    'Id Mensagem Gmail',
    'Enviado Em',
    'Criado Em',
    'Atualizado Em',
    'Observacoes'
  ])
});

function coreMailHubGetEventosSheet_() {
  return core_getSheetByKey_(CORE_MAIL_HUB_KEYS.EVENTOS);
}

function coreMailHubGetIndiceSheet_() {
  return core_getSheetByKey_(CORE_MAIL_HUB_KEYS.INDICE);
}

function coreMailHubGetSaidaSheet_() {
  return core_getSheetByKey_(CORE_MAIL_HUB_KEYS.SAIDA);
}

function coreMailHubGetAnexosSheet_() {
  return core_getSheetByKey_(CORE_MAIL_HUB_KEYS.ANEXOS);
}

function coreMailHubGetRegrasSheet_() {
  return core_getSheetByKey_(CORE_MAIL_HUB_KEYS.REGRAS);
}

function coreMailHubGetConfigSheet_() {
  return core_getSheetByKey_(CORE_MAIL_HUB_KEYS.CONFIG);
}

function coreMailHubGetSheetContext_(sheet, opts) {
  core_assertRequired_(sheet, 'sheet');
  opts = opts || {};

  var headerRow = Number(opts.headerRow || 1);
  var startRow = headerRow + 1;
  var lastCol = sheet.getLastColumn();
  var lastRow = sheet.getLastRow();
  var headers = lastCol > 0
    ? sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0].map(function(value) {
        return String(value || '').trim();
      })
    : [];

  var headerMapZero = core_buildHeaderIndexMap_(headers, {
    normalize: true,
    oneBased: false,
    keepFirst: true
  });

  var headerMapOne = core_buildHeaderIndexMap_(headers, {
    normalize: true,
    oneBased: true,
    keepFirst: true
  });

  var rows = opts.includeRows === true && lastCol > 0 && lastRow >= startRow
    ? sheet.getRange(startRow, 1, lastRow - startRow + 1, lastCol).getValues()
    : [];

  return {
    sheet: sheet,
    headerRow: headerRow,
    startRow: startRow,
    lastRow: lastRow,
    lastCol: lastCol,
    headers: headers,
    headerMapZero: headerMapZero,
    headerMapOne: headerMapOne,
    rows: rows
  };
}

function coreMailHubFindHeaderIndexZero_(ctx, headerName) {
  return core_findHeaderIndex_(ctx.headerMapZero, headerName, {
    normalize: true,
    notFoundValue: -1
  });
}

function coreMailHubFindHeaderIndexOne_(ctx, headerName) {
  return core_findHeaderIndex_(ctx.headerMapOne, headerName, {
    normalize: true,
    notFoundValue: 0
  });
}

function coreMailHubGetRowValue_(row, ctx, headerName, defaultValue) {
  var idx = coreMailHubFindHeaderIndexZero_(ctx, headerName);
  if (idx < 0) {
    return typeof defaultValue === 'undefined' ? '' : defaultValue;
  }
  return row[idx];
}

function coreMailHubSetRowValue_(row, ctx, headerName, value) {
  var idx = coreMailHubFindHeaderIndexZero_(ctx, headerName);
  if (idx < 0) return false;
  row[idx] = value;
  return true;
}

function coreMailHubAppendRow_(sheet, ctx, payload) {
  var row = new Array(ctx.headers.length);
  for (var i = 0; i < row.length; i++) row[i] = '';

  Object.keys(payload || {}).forEach(function(headerName) {
    coreMailHubSetRowValue_(row, ctx, headerName, payload[headerName]);
  });

  sheet.appendRow(row);
  return row;
}

function coreMailHubWriteCell_(sheet, rowNumber, ctx, headerName, value) {
  return core_writeCellByHeader_(sheet, rowNumber, ctx.headerMapOne, headerName, value, {
    normalize: true,
    oneBased: true
  });
}

function coreMailHubEnsureAnexosSchemaExtensions_() {
  var sheet = coreMailHubGetAnexosSheet_();
  var ctx = coreMailHubGetSheetContext_(sheet);
  var missing = CORE_MAIL_HUB_SCHEMA.MAIL_ANEXOS.filter(function(headerName) {
    return coreMailHubFindHeaderIndexZero_(ctx, headerName) < 0;
  });

  if (!missing.length) {
    return Object.freeze({
      ok: true,
      sheetName: sheet.getName(),
      addedHeaders: []
    });
  }

  var startColumn = ctx.lastCol + 1;
  sheet.getRange(1, startColumn, 1, missing.length).setValues([missing]);

  return Object.freeze({
    ok: true,
    sheetName: sheet.getName(),
    addedHeaders: missing
  });
}

function coreMailHubAssertSheetSchema_(sheetKey, sheet, requiredHeaders) {
  var ctx = coreMailHubGetSheetContext_(sheet);
  if (!ctx.headers.length) {
    throw new Error(
      'Mail Hub schema invalido em "' + sheetKey + '": a aba precisa ter cabecalhos na linha 1.'
    );
  }

  var missing = requiredHeaders.filter(function(headerName) {
    return coreMailHubFindHeaderIndexZero_(ctx, headerName) < 0;
  });

  if (missing.length) {
    throw new Error(
      'Mail Hub schema invalido em "' + sheetKey + '" (' + sheet.getName() + '). ' +
      'Cabecalhos obrigatorios ausentes: ' + missing.join(', ')
    );
  }

  return Object.freeze({
    key: sheetKey,
    sheetName: sheet.getName(),
    requiredHeaders: requiredHeaders.slice(),
    headerCount: ctx.headers.length
  });
}

function coreMailHubAssertSchema_() {
  var anexosExtensions = coreMailHubEnsureAnexosSchemaExtensions_();
  var results = [
    coreMailHubAssertSheetSchema_(
      CORE_MAIL_HUB_KEYS.EVENTOS,
      coreMailHubGetEventosSheet_(),
      CORE_MAIL_HUB_SCHEMA.MAIL_EVENTOS
    ),
    coreMailHubAssertSheetSchema_(
      CORE_MAIL_HUB_KEYS.INDICE,
      coreMailHubGetIndiceSheet_(),
      CORE_MAIL_HUB_SCHEMA.MAIL_INDICE
    ),
    coreMailHubAssertSheetSchema_(
      CORE_MAIL_HUB_KEYS.ANEXOS,
      coreMailHubGetAnexosSheet_(),
      CORE_MAIL_HUB_SCHEMA.MAIL_ANEXOS
    ),
    coreMailHubAssertSheetSchema_(
      CORE_MAIL_HUB_KEYS.CONFIG,
      coreMailHubGetConfigSheet_(),
      CORE_MAIL_HUB_SCHEMA.MAIL_CONFIG
    ),
    coreMailHubAssertSheetSchema_(
      CORE_MAIL_HUB_KEYS.REGRAS,
      coreMailHubGetRegrasSheet_(),
      CORE_MAIL_HUB_SCHEMA.MAIL_REGRAS
    ),
    coreMailHubAssertSheetSchema_(
      CORE_MAIL_HUB_KEYS.SAIDA,
      coreMailHubGetSaidaSheet_(),
      CORE_MAIL_HUB_SCHEMA.MAIL_SAIDA
    )
  ];

  coreMailHubMaybeApplyOperationalSheetUx_();

  return Object.freeze({
    ok: true,
    validatedAt: new Date(),
    sheets: results,
    autoExtended: Object.freeze({
      mailAnexos: anexosExtensions
    })
  });
}

function coreMailHubIsTypedColumnsRestrictionError_(err) {
  var message = err && err.message ? String(err.message) : String(err || '');
  return message.toLowerCase().indexOf('colunas com tipo') >= 0 ||
    message.toLowerCase().indexOf('columns with type') >= 0;
}

function coreMailHubTryUxOperation_(sheet, operation, fn) {
  try {
    var result = fn();
    return Object.freeze({
      operation: operation,
      status: 'APPLIED',
      result: typeof result === 'undefined' ? null : result
    });
  } catch (err) {
    if (!coreMailHubIsTypedColumnsRestrictionError_(err)) {
      throw err;
    }

    return Object.freeze({
      operation: operation,
      status: 'SKIPPED_TYPED_COLUMNS',
      reason: err && err.message ? err.message : String(err || 'Operacao nao permitida em colunas com tipo.')
    });
  }
}

function coreMailHubApplyHeaderNotesBySheet_(sheet, notesMap, headerRow) {
  var row = Math.max(1, Number(headerRow || 1));
  var lastColumn = Math.max(sheet.getLastColumn(), 1);
  var headerValues = sheet.getRange(row, 1, 1, lastColumn).getValues()[0];
  var resolvedNotes = {};

  for (var i = 0; i < headerValues.length; i++) {
    var header = String(headerValues[i] || '').trim();
    if (!header) continue;
    resolvedNotes[header] = String((notesMap || {})[header] || '').trim();
  }

  return core_applyHeaderNotes_(sheet, resolvedNotes, row);
}

function coreMailHubApplyHeaderStyles_(sheet, groups, headerRow) {
  return core_applyHeaderColors_(sheet, groups || [], headerRow || 1, {
    defaultColor: '#f3f3f3',
    fontColor: CORE_MAIL_HUB_UX.colors.neutralText,
    fontWeight: 'bold',
    wrap: true
  });
}

function coreMailHubApplyDropdowns_(sheet, rules, headerRow) {
  return core_applyDropdownValidationByHeader_(sheet, rules || {}, headerRow || 1, {});
}

function coreMailHubApplyClipWrapByHeaders_(sheet, headers, headerRow) {
  var ctx = coreMailHubGetSheetContext_(sheet, { headerRow: headerRow || 1 });
  var applied = 0;
  var wrapStrategy = SpreadsheetApp.WrapStrategy.CLIP;

  for (var i = 0; i < headers.length; i++) {
    var colNumber = coreMailHubFindHeaderIndexOne_(ctx, headers[i]);
    if (!colNumber) continue;

    sheet.getRange(1, colNumber, Math.max(sheet.getMaxRows(), 1), 1)
      .setWrapStrategy(wrapStrategy)
      .setVerticalAlignment('middle');
    applied++;
  }

  return applied;
}

function coreMailHubApplyFixedDataRowHeights_(sheet, headerRow, rowHeight) {
  var startRow = Math.max(2, Number(headerRow || 1) + 1);
  var maxRows = Math.max(sheet.getMaxRows(), startRow);
  var rowCount = Math.max(maxRows - startRow + 1, 0);

  if (!rowCount) return 0;

  if (typeof sheet.setRowHeightsForced === 'function') {
    sheet.setRowHeightsForced(startRow, rowCount, rowHeight);
  } else {
    sheet.setRowHeights(startRow, rowCount, rowHeight);
  }

  return rowCount;
}

function coreMailHubApplySheetUx_(sheet, schemaKey, opts) {
  opts = opts || {};
  var rules = CORE_MAIL_HUB_UX.sheetRules[schemaKey] || {};
  var headerRow = Number(opts.headerRow || 1);
  var dataRowHeight = Number(opts.dataRowHeight || rules.dataRowHeight || CORE_MAIL_HUB_UX.defaultDataRowHeight);
  var operations = [];

  operations.push(coreMailHubTryUxOperation_(sheet, 'freezeHeaderRow', function() {
    return core_freezeHeaderRow_(sheet, headerRow);
  }));

  operations.push(coreMailHubTryUxOperation_(sheet, 'headerNotes', function() {
    return coreMailHubApplyHeaderNotesBySheet_(sheet, rules.notes || {}, headerRow);
  }));

  operations.push(coreMailHubTryUxOperation_(sheet, 'headerStyles', function() {
    return coreMailHubApplyHeaderStyles_(sheet, rules.colorGroups || [], headerRow);
  }));

  operations.push(coreMailHubTryUxOperation_(sheet, 'filter', function() {
    return core_ensureFilter_(sheet, headerRow, { recreate: false });
  }));

  operations.push(coreMailHubTryUxOperation_(sheet, 'clipColumns', function() {
    return coreMailHubApplyClipWrapByHeaders_(sheet, rules.clipHeaders || [], headerRow);
  }));

  operations.push(coreMailHubTryUxOperation_(sheet, 'fixedRowHeights', function() {
    return coreMailHubApplyFixedDataRowHeights_(sheet, headerRow, dataRowHeight);
  }));

  operations.push(coreMailHubTryUxOperation_(sheet, 'dropdownValidation', function() {
    return coreMailHubApplyDropdowns_(sheet, rules.dropdownRules || {}, headerRow);
  }));

  return Object.freeze({
    sheetName: sheet.getName(),
    schemaKey: schemaKey,
    dataRowHeight: dataRowHeight,
    operations: Object.freeze(operations)
  });
}

function coreMailApplyOperationalSheetUx_(opts) {
  opts = opts || {};
  coreMailHubAssertSchema_();

  return Object.freeze({
    ok: true,
    eventos: coreMailHubApplySheetUx_(coreMailHubGetEventosSheet_(), 'MAIL_EVENTOS', opts),
    indice: coreMailHubApplySheetUx_(coreMailHubGetIndiceSheet_(), 'MAIL_INDICE', opts),
    anexos: coreMailHubApplySheetUx_(coreMailHubGetAnexosSheet_(), 'MAIL_ANEXOS', opts),
    regras: coreMailHubApplySheetUx_(coreMailHubGetRegrasSheet_(), 'MAIL_REGRAS', opts),
    config: coreMailHubApplySheetUx_(coreMailHubGetConfigSheet_(), 'MAIL_CONFIG', opts),
    saida: coreMailHubApplySheetUx_(coreMailHubGetSaidaSheet_(), 'MAIL_SAIDA', opts)
  });
}

function coreMailHubMaybeApplyOperationalSheetUx_() {
  var cache = null;

  try {
    cache = CacheService.getScriptCache();
    if (cache && cache.get(CORE_MAIL_HUB_UX.cacheKey)) {
      return Object.freeze({
        ok: true,
        applied: false,
        reason: 'CACHE_HIT'
      });
    }
  } catch (cacheErr) {
    cache = null;
  }

  try {
    var result = Object.freeze({
      ok: true,
      applied: true,
      eventos: coreMailHubApplySheetUx_(coreMailHubGetEventosSheet_(), 'MAIL_EVENTOS', {}),
      indice: coreMailHubApplySheetUx_(coreMailHubGetIndiceSheet_(), 'MAIL_INDICE', {}),
      anexos: coreMailHubApplySheetUx_(coreMailHubGetAnexosSheet_(), 'MAIL_ANEXOS', {}),
      regras: coreMailHubApplySheetUx_(coreMailHubGetRegrasSheet_(), 'MAIL_REGRAS', {}),
      config: coreMailHubApplySheetUx_(coreMailHubGetConfigSheet_(), 'MAIL_CONFIG', {}),
      saida: coreMailHubApplySheetUx_(coreMailHubGetSaidaSheet_(), 'MAIL_SAIDA', {})
    });

    try {
      if (cache) {
        cache.put(CORE_MAIL_HUB_UX.cacheKey, '1', CORE_MAIL_HUB_UX.cacheTtlSeconds);
      }
    } catch (cachePutErr) {}

    return result;
  } catch (err) {
    return Object.freeze({
      ok: false,
      applied: false,
      reason: err && err.message ? err.message : String(err || 'UX_ERROR')
    });
  }
}

function coreMailHubNormalizeFlag_(value) {
  return core_normalizeText_(value, {
    removeAccents: true,
    collapseWhitespace: true,
    caseMode: 'upper'
  });
}

function coreMailHubGetConfigMap_() {
  var sheet = coreMailHubGetConfigSheet_();
  var ctx = coreMailHubGetSheetContext_(sheet, { includeRows: true });
  var out = Object.create(null);

  for (var i = 0; i < ctx.rows.length; i++) {
    var row = ctx.rows[i];
    var key = core_normalizeText_(coreMailHubGetRowValue_(row, ctx, 'Chave', ''), {
      collapseWhitespace: true,
      caseMode: 'upper'
    });

    if (!key) continue;

    var ativo = coreMailHubNormalizeFlag_(coreMailHubGetRowValue_(row, ctx, 'Ativo', 'SIM'));
    if (ativo && ativo !== 'SIM' && ativo !== 'TRUE' && ativo !== '1') continue;

    out[key] = coreMailHubGetRowValue_(row, ctx, 'Valor', '');
  }

  return out;
}

function coreMailHubGetConfigValue_(configMap, key, defaultValue) {
  var normalizedKey = core_normalizeText_(key, {
    collapseWhitespace: true,
    caseMode: 'upper'
  });

  if (!Object.prototype.hasOwnProperty.call(configMap, normalizedKey)) {
    return defaultValue;
  }

  var value = configMap[normalizedKey];
  if (value === '' || value === null || typeof value === 'undefined') {
    return defaultValue;
  }

  return value;
}

function coreMailHubGetConfigNumber_(configMap, key, defaultValue) {
  var raw = coreMailHubGetConfigValue_(configMap, key, defaultValue);
  var num = Number(raw);
  return isNaN(num) ? Number(defaultValue) : num;
}

function coreMailHubGetConfigBoolean_(configMap, key, defaultValue) {
  var raw = coreMailHubNormalizeFlag_(coreMailHubGetConfigValue_(configMap, key, defaultValue ? 'SIM' : 'NAO'));
  if (raw === 'SIM' || raw === 'TRUE' || raw === '1') return true;
  if (raw === 'NAO' || raw === 'FALSE' || raw === '0') return false;
  return defaultValue === true;
}

function coreMailHubParseConfigList_(value) {
  return String(value || '')
    .split(/[\r\n,;]+/)
    .map(function(item) { return String(item || '').trim(); })
    .filter(function(item) { return !!item; });
}

function coreMailHubGetConfigList_(configMap, key) {
  return coreMailHubParseConfigList_(coreMailHubGetConfigValue_(configMap, key, ''));
}

function coreMailHubGetConfigRegexList_(configMap, key) {
  return coreMailHubGetConfigList_(configMap, key).map(function(pattern) {
    try {
      return new RegExp(pattern, 'i');
    } catch (err) {
      throw new Error('Regex invalida em MAIL_CONFIG para "' + key + '": ' + pattern);
    }
  });
}

function coreMailHubGetConfig_(key, defaultValue) {
  return coreMailHubGetConfigValue_(coreMailHubGetConfigMap_(), key, defaultValue);
}

function coreMailHubGetConfigBooleanByKey_(key, defaultValue) {
  return coreMailHubGetConfigBoolean_(coreMailHubGetConfigMap_(), key, defaultValue === true);
}

function coreMailHubGetConfigListByKey_(key) {
  return coreMailHubGetConfigList_(coreMailHubGetConfigMap_(), key);
}

function coreMailHubBuildIngestConfig_(configMap, opts) {
  opts = opts || {};

  var useOnlyTaggedSubjects = Object.prototype.hasOwnProperty.call(opts, 'useOnlyTaggedSubjects')
    ? opts.useOnlyTaggedSubjects === true
    : coreMailHubGetConfigBoolean_(configMap, 'USAR_SOMENTE_ASSUNTOS_GEAPA', false);
  var requiredSubjectPrefix = String(coreMailHubGetConfigValue_(configMap, 'ASSUNTO_PREFIXO_OBRIGATORIO', '[GEAPA]') || '').trim();

  return {
    requiredSubjectPrefix: requiredSubjectPrefix || '[GEAPA]',
    useOnlyTaggedSubjects: useOnlyTaggedSubjects,
    ignoredSenders: coreMailHubGetConfigList_(configMap, 'IGNORAR_REMETENTES').map(function(item) {
      return core_extractEmailAddress_(item);
    }).filter(function(item) {
      return !!item;
    }),
    ignoredDomains: coreMailHubGetConfigList_(configMap, 'IGNORAR_DOMINIOS').map(function(item) {
      return core_normalizeText_(item, {
        collapseWhitespace: true,
        caseMode: 'lower'
      });
    }).filter(function(item) {
      return !!item;
    }),
    ignoredSubjectRegexes: coreMailHubGetConfigRegexList_(configMap, 'IGNORAR_ASSUNTOS_REGEX'),
    maxEventsPerExecution: Math.max(
      0,
      Number(
        Object.prototype.hasOwnProperty.call(opts, 'maxEventsPerExecution')
          ? opts.maxEventsPerExecution
          : coreMailHubGetConfigNumber_(configMap, 'MAX_EVENTOS_POR_EXECUCAO', 0)
      ) || 0
    ),
    saveFullBody: Object.prototype.hasOwnProperty.call(opts, 'saveFullBody')
      ? opts.saveFullBody === true
      : coreMailHubGetConfigBoolean_(configMap, 'SALVAR_CORPO_COMPLETO', false),
    markNoiseAsIgnored: Object.prototype.hasOwnProperty.call(opts, 'markNoiseAsIgnored')
      ? opts.markNoiseAsIgnored === true
      : coreMailHubGetConfigBoolean_(configMap, 'MARCAR_RUIDO_COMO_IGNORADO', false)
  };
}

function coreMailHubBuildEventosState_(sheet) {
  var ctx = coreMailHubGetSheetContext_(sheet, { includeRows: true });
  var messageIds = Object.create(null);
  var rowByEventId = Object.create(null);

  for (var i = 0; i < ctx.rows.length; i++) {
    var row = ctx.rows[i];
    var messageId = String(coreMailHubGetRowValue_(row, ctx, 'Id Mensagem Gmail', '') || '').trim();
    var eventId = String(coreMailHubGetRowValue_(row, ctx, 'Id Evento', '') || '').trim();
    var rowNumber = ctx.startRow + i;

    if (messageId) messageIds[messageId] = true;
    if (eventId) rowByEventId[eventId] = rowNumber;
  }

  return {
    ctx: ctx,
    messageIds: messageIds,
    rowByEventId: rowByEventId
  };
}

function coreMailHubBuildIndiceState_(sheet) {
  var ctx = coreMailHubGetSheetContext_(sheet, { includeRows: true });
  var rowByCorrelationKey = Object.create(null);
  var recordByCorrelationKey = Object.create(null);

  for (var i = 0; i < ctx.rows.length; i++) {
    var row = ctx.rows[i];
    var correlationKey = core_normalizeText_(
      coreMailHubGetRowValue_(row, ctx, 'Chave de Correlacao', ''),
      { collapseWhitespace: true, caseMode: 'upper' }
    );

    if (!correlationKey) continue;

    rowByCorrelationKey[correlationKey] = ctx.startRow + i;
    recordByCorrelationKey[correlationKey] = row;
  }

  return {
    ctx: ctx,
    rowByCorrelationKey: rowByCorrelationKey,
    recordByCorrelationKey: recordByCorrelationKey
  };
}

function coreMailHubBuildSaidaState_(sheet) {
  var ctx = coreMailHubGetSheetContext_(sheet, { includeRows: true });
  var rowBySaidaId = Object.create(null);

  for (var i = 0; i < ctx.rows.length; i++) {
    var row = ctx.rows[i];
    var saidaId = String(coreMailHubGetRowValue_(row, ctx, 'Id Saida', '') || '').trim();
    if (!saidaId) continue;
    rowBySaidaId[saidaId] = ctx.startRow + i;
  }

  return {
    ctx: ctx,
    rowBySaidaId: rowBySaidaId
  };
}

function coreMailHubNormalizeEmailList_(value) {
  return core_uniqueEmails_(value).slice();
}

function coreMailHubJoinEmails_(value) {
  return coreMailHubNormalizeEmailList_(value).join(', ');
}

function coreMailHubGenerateOutgoingId_() {
  return 'MSA-' + new Date().getTime().toString(36).toUpperCase();
}

function coreMailHubNormalizePriority_(value) {
  var normalized = coreMailHubNormalizeFlag_(value || 'NORMAL');
  return normalized || 'NORMAL';
}

function coreMailHubParseOutboxMetadata_(value) {
  if (!value) return {};
  if (typeof value === 'object') return value;

  var text = String(value || '').trim();
  if (!text) return {};

  try {
    return JSON.parse(text);
  } catch (err) {
    return { raw: text };
  }
}

function coreMailHubEnvelopeOfficialEmail_() {
  var officialData = typeof coreMailRendererGetOfficialGroupData_ === 'function'
    ? coreMailRendererGetOfficialGroupData_()
    : null;
  var officialEmail = officialData && officialData.email
    ? core_extractEmailAddress_(officialData.email)
    : '';
  if (officialEmail) return officialEmail;

  try {
    var activeUserEmail = core_extractEmailAddress_(Session.getActiveUser().getEmail());
    if (activeUserEmail) return activeUserEmail;
  } catch (err) {}

  return '';
}

function coreMailHubNormalizeOutgoingContract_(contract) {
  contract = contract || {};

  var moduleName = String(contract.moduleName || '').trim();
  var templateKey = String(contract.templateKey || 'GEAPA_OPERACIONAL').trim();
  var correlationKey = String(contract.correlationKey || '').trim().toUpperCase();
  var subjectHuman = String(contract.subjectHuman || '').trim();
  var toList = coreMailHubNormalizeEmailList_(contract.to);
  var ccList = coreMailHubNormalizeEmailList_(contract.cc);
  var bccList = coreMailHubNormalizeEmailList_(contract.bcc);
  var finalSubject = coreMailBuildFinalSubject_(subjectHuman, correlationKey);
  var metadata = coreMailHubParseOutboxMetadata_(contract.metadata);
  var sendAfter = contract.sendAfter ? coreMailHubCoerceDate_(contract.sendAfter) : null;
  var payload = contract.payload || {};
  var entityType = String(contract.entityType || '').trim();
  var entityId = String(contract.entityId || '').trim();
  var flowCode = String(contract.flowCode || '').trim();
  var stage = String(contract.stage || contract.flowCode || '').trim();
  var name = String(contract.name || contract.recipientName || '').trim();
  var replyTo = String(contract.replyTo || '').trim();
  var attachmentRefs = Array.isArray(contract.attachments)
    ? contract.attachments.slice()
    : (Array.isArray(metadata.attachmentRefs) ? metadata.attachmentRefs.slice() : []);
  var forceQueueDuplicate = contract.forceQueueDuplicate === true || metadata.forceQueueDuplicate === true;

  core_assertRequired_(moduleName, 'contract.moduleName');
  core_assertRequired_(correlationKey, 'contract.correlationKey');
  core_assertRequired_(subjectHuman, 'contract.subjectHuman');

  if (!toList.length && !bccList.length) {
    throw new Error('coreMailQueueOutgoing_: informe ao menos um destinatario em "to" ou "bcc".');
  }

  if (!toList.length && bccList.length) {
    toList = [coreMailHubEnvelopeOfficialEmail_()].filter(function(item) { return !!item; });
    if (!toList.length) {
      throw new Error(
        'coreMailQueueOutgoing_: envio em BCC requer EMAIL_OFICIAL em DADOS_OFICIAIS_GEAPA ' +
        'ou uma conta ativa disponivel para servir como envelope principal.'
      );
    }
  }

  if (!entityType || !entityId || !flowCode || !stage) {
    var parsed = coreMailParseCorrelationKey_(correlationKey);
    if (parsed && parsed.isValid) {
      entityType = entityType || String(parsed.entityType || '').trim();
      entityId = entityId || String(parsed.entityId || parsed.businessId || '').trim();
      flowCode = flowCode || String(parsed.flowCode || '').trim();
      stage = stage || String(parsed.stage || parsed.flowCode || '').trim();
    }
  }

  return {
    moduleName: moduleName,
    templateKey: templateKey || 'GEAPA_OPERACIONAL',
    correlationKey: correlationKey,
    entityType: entityType,
    entityId: entityId,
    flowCode: flowCode,
    stage: stage || flowCode,
    to: toList,
    cc: ccList,
    bcc: bccList,
    recipientName: name,
    subjectHuman: subjectHuman,
    finalSubject: finalSubject,
    payload: payload,
    priority: coreMailHubNormalizePriority_(contract.priority),
    sendAfter: sendAfter,
    replyTo: replyTo,
    attachments: attachmentRefs,
    forceQueueDuplicate: forceQueueDuplicate,
    metadata: metadata
  };
}

function coreMailHubSerializeOutboxContract_(normalizedContract) {
  return JSON.stringify({
    templateKey: normalizedContract.templateKey,
    subjectHuman: normalizedContract.subjectHuman,
    payload: normalizedContract.payload || {},
    metadata: normalizedContract.metadata || {},
    entityType: normalizedContract.entityType || '',
    entityId: normalizedContract.entityId || '',
    flowCode: normalizedContract.flowCode || '',
    stage: normalizedContract.stage || '',
    recipientName: normalizedContract.recipientName || '',
    replyTo: normalizedContract.replyTo || '',
    attachmentRefs: normalizedContract.attachments || [],
    forceQueueDuplicate: normalizedContract.forceQueueDuplicate === true
  });
}

function coreMailHubParseOutboxObservacoes_(value) {
  var text = String(value || '').trim();
  if (!text) return {};

  try {
    return JSON.parse(text);
  } catch (err) {
    return { raw: text };
  }
}

function coreMailHubResolveOutgoingAttachments_(attachmentRefs) {
  var refs = Array.isArray(attachmentRefs) ? attachmentRefs : [];
  var blobs = [];

  for (var i = 0; i < refs.length; i++) {
    var ref = refs[i];
    if (!ref) continue;
    blobs.push(coreGetAssetBlob_(ref));
  }

  return blobs;
}

function coreMailHubNormalizeTemplateKey_(value) {
  var normalized = coreMailHubNormalizeFlag_(value);
  if (
    normalized === 'GEAPA_COMEMORATIVO' ||
    normalized === 'GEAPA_OPERACIONAL' ||
    normalized === 'GEAPA_CONVITE' ||
    normalized === 'GEAPA_CLASSICO'
  ) {
    return normalized;
  }
  return 'GEAPA_OPERACIONAL';
}

function coreMailEventExistsByMessageId_(messageId, eventosState) {
  var key = String(messageId || '').trim();
  if (!key) return false;
  return eventosState.messageIds[key] === true;
}

function coreMailExtractCorrelationKey_(subject) {
  var text = String(subject || '').trim();
  if (!text) return '';

  var match = text.match(/\[GEAPA\]\[([^\]]+)\]/i);
  if (!match) return '';

  return String(match[1] || '').trim().toUpperCase();
}

function coreMailResolveModule_(correlationKey) {
  var parsed = coreMailParseCorrelationKey_(correlationKey);
  return parsed.isValid ? String(parsed.moduleName || '').trim() : 'NAO_IDENTIFICADO';
}

function coreMailHubBuildSnippet_(message) {
  try {
    var plainBody = String(message.getPlainBody() || '').replace(/\s+/g, ' ').trim();
    return plainBody.slice(0, 500);
  } catch (err) {
    return '';
  }
}

function coreMailHubGetMessageAttachments_(message) {
  var attachments = [];

  try {
    attachments = message.getAttachments({
      includeInlineImages: false,
      includeAttachments: true
    }) || [];
  } catch (err) {
    try {
      attachments = message.getAttachments() || [];
    } catch (fallbackErr) {
      attachments = [];
    }
  }

  return attachments;
}

function coreMailHubGetThreadLabels_(thread) {
  var labels = thread.getLabels();
  var out = [];

  for (var i = 0; i < labels.length; i++) {
    out.push(String(labels[i].getName() || '').trim());
  }

  return out.join(' | ');
}

function coreMailHubGetEmailDomain_(email) {
  var normalized = core_extractEmailAddress_(email);
  if (!normalized || normalized.indexOf('@') === -1) return '';
  return normalized.split('@')[1];
}

function coreMailHubSubjectHasRequiredPrefix_(subject, ingestConfig) {
  var prefix = String((ingestConfig && ingestConfig.requiredSubjectPrefix) || '[GEAPA]').trim();
  if (!prefix) return true;
  var normalizedSubject = String(subject || '').trim();
  var replyPrefixPattern = /^((re|fw|fwd)\s*:\s*)+/i;

  while (replyPrefixPattern.test(normalizedSubject)) {
    normalizedSubject = normalizedSubject.replace(replyPrefixPattern, '').trim();
  }

  return normalizedSubject.toUpperCase().indexOf(prefix.toUpperCase()) === 0;
}

function coreMailHubDetectNoise_(msgCtx, ingestConfig) {
  var fromEmail = core_extractEmailAddress_(msgCtx.fromEmail || msgCtx.fromRaw || '');
  var fromDomain = coreMailHubGetEmailDomain_(fromEmail);
  var subject = String(msgCtx.subject || '').trim();
  var extractedCorrelationKey = String(coreMailExtractCorrelationKey_(subject) || '').trim();
  var lowerSubject = subject.toLowerCase();
  var lowerFrom = String(fromEmail || '').toLowerCase();
  var technicalPatterns = [
    /github/i,
    /\bcodex\b/i,
    /\b(alerta|alert|incident|workflow|pull request|issue|build failed|monitor|sentry)\b/i
  ];

  if (ingestConfig.ignoredSenders.indexOf(fromEmail) !== -1) {
    return { isNoise: true, reason: 'IGNORAR_REMETENTES' };
  }

  if (fromDomain && ingestConfig.ignoredDomains.indexOf(String(fromDomain || '').toLowerCase()) !== -1) {
    return { isNoise: true, reason: 'IGNORAR_DOMINIOS' };
  }

  for (var i = 0; i < ingestConfig.ignoredSubjectRegexes.length; i++) {
    if (ingestConfig.ignoredSubjectRegexes[i].test(subject)) {
      return { isNoise: true, reason: 'IGNORAR_ASSUNTOS_REGEX' };
    }
  }

  if (
    ingestConfig.useOnlyTaggedSubjects &&
    !coreMailHubSubjectHasRequiredPrefix_(subject, ingestConfig) &&
    !extractedCorrelationKey
  ) {
    return { isNoise: true, reason: 'ASSUNTO_SEM_PREFIXO_OBRIGATORIO' };
  }

  if (
    fromDomain === 'github.com' ||
    fromDomain === 'githubusercontent.com' ||
    lowerFrom.indexOf('github') !== -1 ||
    lowerFrom.indexOf('codex') !== -1
  ) {
    return { isNoise: true, reason: 'REMETENTE_TECNICO' };
  }

  for (var j = 0; j < technicalPatterns.length; j++) {
    if (technicalPatterns[j].test(lowerSubject)) {
      return { isNoise: true, reason: 'ASSUNTO_TECNICO' };
    }
  }

  return { isNoise: false, reason: '' };
}

function coreMailHubBuildMessageContext_(thread, message) {
  var subject = String(message.getSubject() || '').trim();
  var fromRaw = String(message.getFrom() || '').trim();
  var attachments = coreMailHubGetMessageAttachments_(message);

  var msgCtx = {
    subject: subject,
    fromRaw: fromRaw,
    fromEmail: core_extractEmailAddress_(fromRaw),
    fromName: core_extractDisplayName_(fromRaw),
    to: String(message.getTo() || '').trim(),
    cc: String(message.getCc() || '').trim(),
    bcc: String(message.getBcc() || '').trim(),
    replyTo: String(message.getReplyTo() || '').trim(),
    messageId: String(message.getId() || '').trim(),
    threadId: String(thread.getId() || '').trim(),
    messageDate: message.getDate() || '',
    snippet: coreMailHubBuildSnippet_(message),
    plainBody: '',
    labels: coreMailHubGetThreadLabels_(thread),
    attachments: attachments
  };

  try {
    msgCtx.plainBody = String(message.getPlainBody() || '');
  } catch (err) {
    msgCtx.plainBody = '';
  }

  return msgCtx;
}

function coreMailHubBuildEventPayload_(thread, message, ingestedAt, ingestConfig) {
  var msgCtx = coreMailHubBuildMessageContext_(thread, message);
  var noise = coreMailHubDetectNoise_(msgCtx, ingestConfig || coreMailHubBuildIngestConfig_(coreMailHubGetConfigMap_(), {}));

  var routing = coreMailResolveRouting_(msgCtx);
  var correlationKey = String(routing.correlationKey || coreMailExtractCorrelationKey_(msgCtx.subject) || '').trim().toUpperCase();
  var parsedCorrelation = correlationKey ? coreMailParseCorrelationKey_(correlationKey) : { isValid: false };
  var attachments = msgCtx.attachments || [];
  var shouldIgnore = noise.isNoise && ingestConfig && ingestConfig.markNoiseAsIgnored === true;

  return {
    eventId: 'MEV-' + msgCtx.messageId,
    messageId: msgCtx.messageId,
    threadId: msgCtx.threadId,
    messageDate: msgCtx.messageDate,
    ingestedAt: ingestedAt || new Date(),
    direction: 'ENTRADA',
    eventType: 'EMAIL_RECEBIDO',
    flowStep: String(routing.stage || parsedCorrelation.stage || routing.flowCode || parsedCorrelation.flowCode || 'INGESTAO').trim(),
    routingStatus: shouldIgnore ? CORE_MAIL_HUB_STATUS.IGNORADO : (routing.matched ? 'ROTEADO' : 'NAO_IDENTIFICADO'),
    subject: msgCtx.subject,
    fromRaw: msgCtx.fromRaw,
    fromEmail: msgCtx.fromEmail,
    fromName: msgCtx.fromName,
    to: msgCtx.to,
    cc: msgCtx.cc,
    bcc: msgCtx.bcc,
    replyTo: msgCtx.replyTo,
    correlationKey: correlationKey,
    correlationPrefix: coreMailHubExtractCorrelationPrefix_(correlationKey),
    moduleName: routing.matched
      ? String(routing.moduleName || '').trim()
      : (parsedCorrelation.isValid ? String(parsedCorrelation.moduleName || '').trim() : coreMailResolveModule_(correlationKey)),
    moduleCode: String(routing.moduleCode || parsedCorrelation.moduleCode || '').trim(),
    entityType: String(routing.entityType || parsedCorrelation.entityType || '').trim(),
    entityId: String(routing.entityId || parsedCorrelation.entityId || parsedCorrelation.businessId || '').trim(),
    routingReason: String(routing.reason || '').trim(),
    routingConfidence: Number(routing.confidence || 0),
    processingStatus: shouldIgnore ? CORE_MAIL_HUB_STATUS.IGNORADO : CORE_MAIL_HUB_STATUS.PENDENTE,
    processorName: shouldIgnore ? 'coreMailIngestInbox' : '',
    processedAt: shouldIgnore ? (ingestedAt || new Date()) : '',
    hasAttachments: attachments.length ? 'SIM' : 'NAO',
    attachmentCount: attachments.length,
    snippet: msgCtx.snippet,
    plainBody: ingestConfig && ingestConfig.saveFullBody === true ? msgCtx.plainBody : '',
    labels: msgCtx.labels,
    attachments: attachments,
    parsedCorrelation: parsedCorrelation,
    isNoise: noise.isNoise,
    noiseReason: noise.reason,
    shouldIgnore: shouldIgnore
  };
}

function coreMailRegisterEvent_(eventosSheet, eventPayload, eventosState) {
  var ctx = eventosState.ctx;

  coreMailHubAppendRow_(eventosSheet, ctx, {
    'Id Evento': eventPayload.eventId,
    'Data Hora Evento': eventPayload.messageDate,
    'Direcao': eventPayload.direction,
    'Tipo Evento': eventPayload.eventType,
    'Modulo Dono': eventPayload.moduleName,
    'Tipo Entidade': eventPayload.entityType,
    'Id Entidade': eventPayload.entityId,
    'Chave de Correlacao': eventPayload.correlationKey,
    'Etapa Fluxo': eventPayload.flowStep,
    'Id Mensagem Gmail': eventPayload.messageId,
    'Id Thread Gmail': eventPayload.threadId,
    'Id Mensagem Pai': '',
    'Assunto': eventPayload.subject,
    'Email Remetente': eventPayload.fromEmail,
    'Nome Remetente': eventPayload.fromName,
    'Emails Destinatarios': eventPayload.to,
    'Emails Cc': eventPayload.cc,
    'Emails Cco': eventPayload.bcc,
    'Trecho Corpo': eventPayload.snippet,
    'Corpo Texto': eventPayload.plainBody || '',
    'Possui Anexos': eventPayload.hasAttachments,
    'Quantidade Anexos': eventPayload.attachmentCount,
    'Nomes Anexos': (eventPayload.attachments || []).map(function(item) {
      return String(item.getName() || '').trim();
    }).join(' | '),
    'Status Roteamento': eventPayload.routingStatus,
    'Status Processamento': eventPayload.processingStatus,
    'Processado Por': eventPayload.processorName,
    'Data Hora Processamento': eventPayload.processedAt,
    'Observacoes': eventPayload.noiseReason ? ('NOISE_REASON=' + eventPayload.noiseReason) : '',
    'Json Bruto': JSON.stringify({
      messageId: eventPayload.messageId,
      threadId: eventPayload.threadId,
      labels: eventPayload.labels,
      replyTo: eventPayload.replyTo,
      fromRaw: eventPayload.fromRaw,
      moduleCode: eventPayload.moduleCode,
      entityType: eventPayload.entityType,
      entityId: eventPayload.entityId,
      routingReason: eventPayload.routingReason,
      routingConfidence: eventPayload.routingConfidence
    }),
    'Criado Em': eventPayload.ingestedAt,
    'Atualizado Em': eventPayload.ingestedAt
  });

  var rowNumber = eventosSheet.getLastRow();
  eventosState.messageIds[eventPayload.messageId] = true;
  eventosState.rowByEventId[eventPayload.eventId] = rowNumber;
  return Object.freeze({
    eventId: eventPayload.eventId,
    rowNumber: rowNumber
  });
}

function coreMailRegisterAttachments_(anexosSheet, eventPayload, anexosCtx) {
  var attachments = eventPayload.attachments || [];
  var inserted = 0;

  for (var i = 0; i < attachments.length; i++) {
    var attachment = attachments[i];
    var bytesLength = 0;

    try {
      bytesLength = attachment.getBytes().length;
    } catch (err) {
      bytesLength = 0;
    }

    coreMailHubAppendRow_(anexosSheet, anexosCtx, {
      'Id Anexo': 'MAN-' + eventPayload.messageId + '-' + String(i + 1),
      'Id Evento': eventPayload.eventId,
      'Modulo Dono': eventPayload.moduleName,
      'Tipo Entidade': eventPayload.entityType,
      'Id Entidade': eventPayload.entityId,
      'Chave de Correlacao': eventPayload.correlationKey,
      'Etapa Fluxo': eventPayload.flowStep,
      'Id Mensagem Gmail': eventPayload.messageId,
      'Id Thread Gmail': eventPayload.threadId,
      'Indice Anexo Mensagem': i + 1,
      'Nome Arquivo': String(attachment.getName() || '').trim(),
      'Tipo Mime': String(attachment.getContentType() || '').trim(),
      'Tamanho Bytes': bytesLength,
      'Foi Salvo No Drive': 'NAO',
      'Id Arquivo Drive': '',
      'Link Arquivo Drive': '',
      'Pasta Destino Drive': '',
      'Status Anexo': 'PENDENTE',
      'Processado Por': '',
      'Data Hora Processamento': '',
      'Observacoes': 'Attachment Index=' + String(i + 1),
      'Criado Em': eventPayload.ingestedAt,
      'Atualizado Em': eventPayload.ingestedAt
    });

    inserted++;
  }

  return inserted;
}

function coreMailHubCoerceDate_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) return value;
  if (value === '' || value === null || typeof value === 'undefined') return null;

  var parsed = new Date(value);
  return isNaN(parsed.getTime()) ? null : parsed;
}

function coreMailHubExtractCorrelationPrefix_(correlationKey) {
  var normalized = String(correlationKey || '').trim().toUpperCase();
  if (!normalized) return '';
  var firstDash = normalized.indexOf('-');
  return firstDash > 0 ? normalized.substring(0, firstDash) : normalized;
}

function coreMailHubCollectAttachmentStatsByCorrelationKey_(correlationKey) {
  var normalizedKey = core_normalizeText_(correlationKey, {
    collapseWhitespace: true,
    caseMode: 'upper'
  });
  if (!normalizedKey) {
    return {
      totalAttachments: 0,
      hasPendingAttachment: false
    };
  }

  var anexosSheet = coreMailHubGetAnexosSheet_();
  var ctx = coreMailHubGetSheetContext_(anexosSheet, { includeRows: true });
  var stats = {
    totalAttachments: 0,
    hasPendingAttachment: false
  };

  for (var i = 0; i < ctx.rows.length; i++) {
    var row = ctx.rows[i];
    var rowKey = core_normalizeText_(coreMailHubGetRowValue_(row, ctx, 'Chave de Correlacao', ''), {
      collapseWhitespace: true,
      caseMode: 'upper'
    });
    if (rowKey !== normalizedKey) continue;

    stats.totalAttachments++;

    var statusAnexo = coreMailHubNormalizeAttachmentStatus_(coreMailHubGetRowValue_(row, ctx, 'Status Anexo', 'PENDENTE'));
    if (coreMailHubIsPendingAttachmentStatus_(statusAnexo)) {
      stats.hasPendingAttachment = true;
    }
  }

  return stats;
}

function coreMailHubCollectCorrelationStats_(correlationKey) {
  var normalizedKey = core_normalizeText_(correlationKey, {
    collapseWhitespace: true,
    caseMode: 'upper'
  });
  if (!normalizedKey) return null;

  var eventosSheet = coreMailHubGetEventosSheet_();
  var ctx = coreMailHubGetSheetContext_(eventosSheet, { includeRows: true });
  var stats = {
    correlationKey: normalizedKey,
    correlationPrefix: coreMailHubExtractCorrelationPrefix_(normalizedKey),
    moduleName: '',
    entityType: '',
    entityId: '',
    currentStage: '',
    latestThreadId: '',
    latestMessageId: '',
    latestDirection: '',
    latestEventType: '',
    latestFromEmail: '',
    latestSubject: '',
    latestEventAt: '',
    latestReplyAt: '',
    totalEvents: 0,
    totalEntries: 0,
    totalOutputs: 0,
    totalAttachments: 0,
    hasPendingEntry: false,
    hasPendingAttachment: false,
    statusConversa: 'PENDENTE',
    createdAt: '',
    updatedAt: new Date(),
    allIgnored: true
  };
  var latestTime = 0;
  var parsedCorrelation = coreMailParseCorrelationKey_(normalizedKey);

  for (var i = 0; i < ctx.rows.length; i++) {
    var row = ctx.rows[i];
    var rowKey = core_normalizeText_(coreMailHubGetRowValue_(row, ctx, 'Chave de Correlacao', ''), {
      collapseWhitespace: true,
      caseMode: 'upper'
    });
    if (rowKey !== normalizedKey) continue;

    stats.totalEvents++;

    var direction = coreMailHubNormalizeFlag_(coreMailHubGetRowValue_(row, ctx, 'Direcao', ''));
    var processingStatus = coreMailHubNormalizeFlag_(coreMailHubGetRowValue_(row, ctx, 'Status Processamento', ''));
    var quantityAttachments = Number(coreMailHubGetRowValue_(row, ctx, 'Quantidade Anexos', 0) || 0);

    if (direction === 'SAIDA') {
      stats.totalOutputs++;
    } else {
      stats.totalEntries++;
    }

    stats.totalAttachments += quantityAttachments;

    if (direction !== 'SAIDA' && processingStatus === CORE_MAIL_HUB_STATUS.PENDENTE) {
      stats.hasPendingEntry = true;
    }
    if (processingStatus !== CORE_MAIL_HUB_STATUS.IGNORADO) {
      stats.allIgnored = false;
    }

    var createdAt = coreMailHubGetRowValue_(row, ctx, 'Criado Em', '');
    if (!stats.createdAt || (coreMailHubCoerceDate_(createdAt) && coreMailHubCoerceDate_(createdAt).getTime() < coreMailHubCoerceDate_(stats.createdAt).getTime())) {
      stats.createdAt = createdAt;
    }

    var eventAt = coreMailHubCoerceDate_(coreMailHubGetRowValue_(row, ctx, 'Data Hora Evento', '')) ||
      coreMailHubCoerceDate_(createdAt);
    var eventTime = eventAt ? eventAt.getTime() : 0;

    if (eventTime >= latestTime) {
      latestTime = eventTime;
      stats.moduleName = String(coreMailHubGetRowValue_(row, ctx, 'Modulo Dono', '') || '').trim();
      stats.entityType = String(coreMailHubGetRowValue_(row, ctx, 'Tipo Entidade', '') || '').trim();
      stats.entityId = String(coreMailHubGetRowValue_(row, ctx, 'Id Entidade', '') || '').trim();
      stats.currentStage = String(coreMailHubGetRowValue_(row, ctx, 'Etapa Fluxo', '') || '').trim();
      stats.latestThreadId = String(coreMailHubGetRowValue_(row, ctx, 'Id Thread Gmail', '') || '').trim();
      stats.latestMessageId = String(coreMailHubGetRowValue_(row, ctx, 'Id Mensagem Gmail', '') || '').trim();
      stats.latestDirection = direction;
      stats.latestEventType = String(coreMailHubGetRowValue_(row, ctx, 'Tipo Evento', '') || '').trim();
      stats.latestFromEmail = String(coreMailHubGetRowValue_(row, ctx, 'Email Remetente', '') || '').trim();
      stats.latestSubject = String(coreMailHubGetRowValue_(row, ctx, 'Assunto', '') || '').trim();
      stats.latestEventAt = coreMailHubGetRowValue_(row, ctx, 'Data Hora Evento', '');
      if (direction !== 'SAIDA') {
        stats.latestReplyAt = stats.latestEventAt;
      }
    }
  }

  if (!stats.totalEvents) return null;

  if (parsedCorrelation.isValid) {
    if (!stats.moduleName) stats.moduleName = String(parsedCorrelation.moduleName || '').trim();
    if (!stats.entityType) stats.entityType = String(parsedCorrelation.entityType || '').trim();
    if (!stats.entityId) stats.entityId = String(parsedCorrelation.entityId || parsedCorrelation.businessId || '').trim();
    if (!stats.currentStage) stats.currentStage = String(parsedCorrelation.stage || parsedCorrelation.flowCode || '').trim();
  }

  var attachmentStats = coreMailHubCollectAttachmentStatsByCorrelationKey_(normalizedKey);
  if (attachmentStats.totalAttachments > 0) {
    stats.totalAttachments = attachmentStats.totalAttachments;
  }
  stats.hasPendingAttachment = attachmentStats.hasPendingAttachment;

  if (stats.allIgnored) {
    stats.statusConversa = CORE_MAIL_HUB_STATUS.IGNORADO;
  } else if (stats.hasPendingEntry || stats.hasPendingAttachment) {
    stats.statusConversa = CORE_MAIL_HUB_STATUS.PENDENTE;
  } else {
    stats.statusConversa = 'CONCLUIDA';
  }

  return stats;
}

function coreMailHubRefreshIndexSummaryByCorrelationKey_(correlationKey, indiceSheet, indiceState) {
  var stats = coreMailHubCollectCorrelationStats_(correlationKey);
  if (!stats) {
    return Object.freeze({
      updated: false,
      skipped: true,
      reason: 'SEM_EVENTOS'
    });
  }

  indiceSheet = indiceSheet || coreMailHubGetIndiceSheet_();
  indiceState = indiceState || coreMailHubBuildIndiceState_(indiceSheet);

  var rowNumber = indiceState.rowByCorrelationKey[stats.correlationKey];
  var ctx = indiceState.ctx;
  var inserted = false;

  if (!rowNumber) {
    var row = coreMailHubAppendRow_(indiceSheet, ctx, {
      'Chave de Correlacao': stats.correlationKey,
      'Criado Em': stats.createdAt || stats.updatedAt,
      'Atualizado Em': stats.updatedAt
    });
    rowNumber = indiceSheet.getLastRow();
    indiceState.rowByCorrelationKey[stats.correlationKey] = rowNumber;
    indiceState.recordByCorrelationKey[stats.correlationKey] = row;
    inserted = true;
  }

  var currentRow = indiceState.recordByCorrelationKey[stats.correlationKey];
  var payload = {
    'Chave de Correlacao': stats.correlationKey,
    'Prefixo Correlacao': stats.correlationPrefix,
    'Modulo Dono': stats.moduleName,
    'Tipo Entidade': stats.entityType,
    'Id Entidade': stats.entityId,
    'Etapa Atual': stats.currentStage,
    'Id Thread Gmail': stats.latestThreadId,
    'Id Ultima Mensagem': stats.latestMessageId,
    'Ultima Direcao': stats.latestDirection,
    'Ultimo Tipo Evento': stats.latestEventType,
    'Ultimo Email Remetente': stats.latestFromEmail,
    'Ultimo Assunto': stats.latestSubject,
    'Data Hora Ultimo Evento': stats.latestEventAt,
    'Ha Entrada Pendente': stats.hasPendingEntry ? 'SIM' : 'NAO',
    'Ha Anexo Pendente': stats.hasPendingAttachment ? 'SIM' : 'NAO',
    'Quantidade Eventos': stats.totalEvents,
    'Quantidade Entradas': stats.totalEntries,
    'Quantidade Saidas': stats.totalOutputs,
    'Quantidade Anexos': stats.totalAttachments,
    'Ultima Resposta Em': stats.latestReplyAt,
    'Status Conversa': stats.statusConversa,
    'Atualizado Em': stats.updatedAt
  };

  Object.keys(payload).forEach(function(headerName) {
    coreMailHubWriteCell_(indiceSheet, rowNumber, ctx, headerName, payload[headerName]);
    coreMailHubSetRowValue_(currentRow, ctx, headerName, payload[headerName]);
  });

  if (inserted) {
    coreMailHubWriteCell_(indiceSheet, rowNumber, ctx, 'Criado Em', stats.createdAt || stats.updatedAt);
    coreMailHubSetRowValue_(currentRow, ctx, 'Criado Em', stats.createdAt || stats.updatedAt);
  }

  return Object.freeze({
    updated: true,
    inserted: inserted,
    rowNumber: rowNumber
  });
}

function coreMailUpsertIndex_(indiceSheet, eventPayload, indiceState) {
  var correlationKey = String(eventPayload.correlationKey || '').trim().toUpperCase();
  if (!correlationKey) {
    return Object.freeze({
      updated: false,
      skipped: true,
      reason: 'SEM_CHAVE_CORRELACAO'
    });
  }

  return coreMailHubRefreshIndexSummaryByCorrelationKey_(correlationKey, indiceSheet, indiceState);
}

function coreMailHubSendOutgoingMessage_(mailPayload) {
  core_assertRequired_(mailPayload, 'mailPayload');

  var subject = String(mailPayload.subject || '').trim();
  var toList = coreMailHubNormalizeEmailList_(mailPayload.to);
  var ccList = coreMailHubNormalizeEmailList_(mailPayload.cc);
  var bccList = coreMailHubNormalizeEmailList_(mailPayload.bcc);
  var newerThanDays = Number(mailPayload.newerThanDays || 7);
  var maxThreads = Number(mailPayload.maxThreads || 10);
  var sleepMs = Number(mailPayload.sleepMs || 1500);
  var replyTo = String(mailPayload.replyTo || '').trim();
  var attachments = Array.isArray(mailPayload.attachments) ? mailPayload.attachments : [];

  if (!toList.length) {
    throw new Error('coreMailHubSendOutgoingMessage_: destinatario principal ausente.');
  }
  if (!subject) {
    throw new Error('coreMailHubSendOutgoingMessage_: assunto ausente.');
  }

  if (mailPayload.htmlBody) {
    core_sendEmailHtmlWithDefaultInline_(Object.assign({}, mailPayload, {
      to: toList.join(','),
      cc: ccList.join(','),
      bcc: bccList.join(','),
      subject: subject,
      replyTo: replyTo || undefined,
      attachments: attachments.length ? attachments : undefined
    }));
  } else {
    core_sendEmailText_(Object.assign({}, mailPayload, {
      to: toList.join(','),
      cc: ccList.join(','),
      bcc: bccList.join(','),
      subject: subject,
      replyTo: replyTo || undefined,
      attachments: attachments.length ? attachments : undefined
    }));
  }

  if (sleepMs > 0) {
    Utilities.sleep(sleepMs);
  }

  var threadId = '';
  var messageId = '';
  var postSendWarning = '';

  try {
    var safeSubject = subject.replace(/"/g, '\\"');
    var query = 'subject:"' + safeSubject + '" newer_than:' + newerThanDays + 'd';
    if (toList.length === 1) {
      query = 'to:' + toList[0] + ' ' + query;
    }

    var threads = GmailApp.search(query, 0, maxThreads);

    if (threads && threads.length) {
      var thread = threads[0];
      threadId = String(thread.getId() || '').trim();

      var messages = thread.getMessages();
      if (messages && messages.length) {
        messageId = String(messages[messages.length - 1].getId() || '').trim();
      }
    }
  } catch (err) {
    postSendWarning = err && err.message ? err.message : String(err || '');
  }

  return Object.freeze({
    threadId: threadId,
    messageId: messageId,
    sent: true,
    postSendWarning: postSendWarning
  });
}

function coreMailHubBuildOutboxRecord_(row, ctx, rowNumber) {
  return {
    saidaId: String(coreMailHubGetRowValue_(row, ctx, 'Id Saida', '') || '').trim(),
    moduleName: String(coreMailHubGetRowValue_(row, ctx, 'Modulo Dono', '') || '').trim(),
    entityType: String(coreMailHubGetRowValue_(row, ctx, 'Tipo Entidade', '') || '').trim(),
    entityId: String(coreMailHubGetRowValue_(row, ctx, 'Id Entidade', '') || '').trim(),
    correlationKey: String(coreMailHubGetRowValue_(row, ctx, 'Chave de Correlacao', '') || '').trim(),
    flowStep: String(coreMailHubGetRowValue_(row, ctx, 'Etapa Fluxo', '') || '').trim(),
    to: coreMailHubNormalizeEmailList_(coreMailHubGetRowValue_(row, ctx, 'Emails Destinatarios', '')),
    cc: coreMailHubNormalizeEmailList_(coreMailHubGetRowValue_(row, ctx, 'Emails Cc', '')),
    bcc: coreMailHubNormalizeEmailList_(coreMailHubGetRowValue_(row, ctx, 'Emails Cco', '')),
    recipientName: String(coreMailHubGetRowValue_(row, ctx, 'Nome Destinatario', '') || '').trim(),
    subject: String(coreMailHubGetRowValue_(row, ctx, 'Assunto', '') || '').trim(),
    bodyText: String(coreMailHubGetRowValue_(row, ctx, 'Corpo Texto', '') || '').trim(),
    htmlBody: String(coreMailHubGetRowValue_(row, ctx, 'Corpo Html', '') || '').trim(),
    scheduledAt: coreMailHubGetRowValue_(row, ctx, 'Data Hora Agendada', ''),
    priority: String(coreMailHubGetRowValue_(row, ctx, 'Prioridade', '') || '').trim(),
    status: coreMailHubNormalizeFlag_(coreMailHubGetRowValue_(row, ctx, 'Status Envio', '')),
    attempts: Number(coreMailHubGetRowValue_(row, ctx, 'Tentativas', 0) || 0),
    lastError: String(coreMailHubGetRowValue_(row, ctx, 'Ultimo Erro', '') || '').trim(),
    threadId: String(coreMailHubGetRowValue_(row, ctx, 'Id Thread Gmail', '') || '').trim(),
    messageId: String(coreMailHubGetRowValue_(row, ctx, 'Id Mensagem Gmail', '') || '').trim(),
    sentAt: coreMailHubGetRowValue_(row, ctx, 'Enviado Em', ''),
    createdAt: coreMailHubGetRowValue_(row, ctx, 'Criado Em', ''),
    updatedAt: coreMailHubGetRowValue_(row, ctx, 'Atualizado Em', ''),
    observacoes: coreMailHubParseOutboxObservacoes_(coreMailHubGetRowValue_(row, ctx, 'Observacoes', '')),
    rowNumber: rowNumber
  };
}

function coreMailHubFindLatestOutboxByCorrelationKey_(correlationKey, saidaState) {
  var target = core_normalizeText_(correlationKey, {
    collapseWhitespace: true,
    caseMode: 'upper'
  });
  if (!target) return null;

  var ctx = saidaState.ctx;
  for (var i = ctx.rows.length - 1; i >= 0; i--) {
    var row = ctx.rows[i];
    var rowKey = core_normalizeText_(coreMailHubGetRowValue_(row, ctx, 'Chave de Correlacao', ''), {
      collapseWhitespace: true,
      caseMode: 'upper'
    });
    if (rowKey !== target) continue;
    return coreMailHubBuildOutboxRecord_(row, ctx, ctx.startRow + i);
  }

  return null;
}

function coreMailHubBuildOutgoingEventPayload_(outboxRecord, draft, sendResult, sentAt) {
  var brand = typeof coreMailRendererGetBrandProfile_ === 'function'
    ? coreMailRendererGetBrandProfile_(draft.templateKey || 'GEAPA_OPERACIONAL', sentAt)
    : null;

  return {
    eventId: sendResult.messageId ? ('MEV-' + sendResult.messageId) : ('MEV-OUT-' + outboxRecord.saidaId),
    messageId: sendResult.messageId || ('OUT-' + outboxRecord.saidaId),
    threadId: sendResult.threadId || '',
    messageDate: sentAt,
    ingestedAt: sentAt,
    direction: 'SAIDA',
    eventType: 'EMAIL_ENVIADO',
    flowStep: outboxRecord.flowStep || '',
    routingStatus: 'ROTEADO',
    subject: draft.subject,
    fromRaw: brand && brand.officialEmail ? brand.officialEmail : '',
    fromEmail: brand && brand.officialEmail ? brand.officialEmail : '',
    fromName: brand && brand.orgName ? brand.orgName : '',
    to: outboxRecord.to.join(', '),
    cc: outboxRecord.cc.join(', '),
    bcc: outboxRecord.bcc.join(', '),
    replyTo: outboxRecord.observacoes && outboxRecord.observacoes.replyTo ? String(outboxRecord.observacoes.replyTo || '').trim() : '',
    correlationKey: outboxRecord.correlationKey,
    correlationPrefix: coreMailHubExtractCorrelationPrefix_(outboxRecord.correlationKey),
    moduleName: outboxRecord.moduleName,
    moduleCode: '',
    entityType: outboxRecord.entityType,
    entityId: outboxRecord.entityId,
    routingReason: 'MAIL_SAIDA',
    routingConfidence: 1,
    processingStatus: CORE_MAIL_HUB_STATUS.PROCESSADO,
    processorName: 'coreMailProcessOutbox',
    processedAt: sentAt,
    hasAttachments: outboxRecord.observacoes && outboxRecord.observacoes.attachmentRefs && outboxRecord.observacoes.attachmentRefs.length ? 'SIM' : 'NAO',
    attachmentCount: outboxRecord.observacoes && outboxRecord.observacoes.attachmentRefs ? outboxRecord.observacoes.attachmentRefs.length : 0,
    snippet: '',
    plainBody: draft.bodyText || '',
    labels: '',
    attachments: []
  };
}

function coreMailQueueOutgoing_(contract) {
  return core_withLock_('CORE_MAIL_HUB_QUEUE_OUTGOING', function() {
    coreMailHubAssertSchema_();

    var normalizedContract = coreMailHubNormalizeOutgoingContract_(contract);
    var saidaSheet = coreMailHubGetSaidaSheet_();
    var saidaState = coreMailHubBuildSaidaState_(saidaSheet);
    var now = new Date();
    var existingOutbox = coreMailHubFindLatestOutboxByCorrelationKey_(normalizedContract.correlationKey, saidaState);

    if (
      !normalizedContract.forceQueueDuplicate &&
      existingOutbox &&
      (existingOutbox.status === 'PENDENTE' || existingOutbox.status === 'ENVIADO')
    ) {
      return Object.freeze({
        ok: true,
        queued: false,
        duplicate: true,
        saidaId: existingOutbox.saidaId,
        status: existingOutbox.status,
        subject: existingOutbox.subject,
        correlationKey: normalizedContract.correlationKey
      });
    }

    if (!normalizedContract.forceQueueDuplicate && existingOutbox && existingOutbox.status === 'ERRO') {
      coreMailHubWriteCell_(saidaSheet, existingOutbox.rowNumber, saidaState.ctx, 'Modulo Dono', normalizedContract.moduleName);
      coreMailHubWriteCell_(saidaSheet, existingOutbox.rowNumber, saidaState.ctx, 'Tipo Entidade', normalizedContract.entityType);
      coreMailHubWriteCell_(saidaSheet, existingOutbox.rowNumber, saidaState.ctx, 'Id Entidade', normalizedContract.entityId);
      coreMailHubWriteCell_(saidaSheet, existingOutbox.rowNumber, saidaState.ctx, 'Etapa Fluxo', normalizedContract.stage);
      coreMailHubWriteCell_(saidaSheet, existingOutbox.rowNumber, saidaState.ctx, 'Email Destinatario Principal', normalizedContract.to[0] || '');
      coreMailHubWriteCell_(saidaSheet, existingOutbox.rowNumber, saidaState.ctx, 'Emails Destinatarios', coreMailHubJoinEmails_(normalizedContract.to));
      coreMailHubWriteCell_(saidaSheet, existingOutbox.rowNumber, saidaState.ctx, 'Emails Cc', coreMailHubJoinEmails_(normalizedContract.cc));
      coreMailHubWriteCell_(saidaSheet, existingOutbox.rowNumber, saidaState.ctx, 'Emails Cco', coreMailHubJoinEmails_(normalizedContract.bcc));
      coreMailHubWriteCell_(saidaSheet, existingOutbox.rowNumber, saidaState.ctx, 'Nome Destinatario', normalizedContract.recipientName);
      coreMailHubWriteCell_(saidaSheet, existingOutbox.rowNumber, saidaState.ctx, 'Assunto', normalizedContract.finalSubject);
      coreMailHubWriteCell_(saidaSheet, existingOutbox.rowNumber, saidaState.ctx, 'Corpo Texto', '');
      coreMailHubWriteCell_(saidaSheet, existingOutbox.rowNumber, saidaState.ctx, 'Corpo Html', '');
      coreMailHubWriteCell_(saidaSheet, existingOutbox.rowNumber, saidaState.ctx, 'Data Hora Agendada', normalizedContract.sendAfter || '');
      coreMailHubWriteCell_(saidaSheet, existingOutbox.rowNumber, saidaState.ctx, 'Prioridade', normalizedContract.priority);
      coreMailHubWriteCell_(saidaSheet, existingOutbox.rowNumber, saidaState.ctx, 'Status Envio', 'PENDENTE');
      coreMailHubWriteCell_(saidaSheet, existingOutbox.rowNumber, saidaState.ctx, 'Ultimo Erro', '');
      coreMailHubWriteCell_(saidaSheet, existingOutbox.rowNumber, saidaState.ctx, 'Id Thread Gmail', '');
      coreMailHubWriteCell_(saidaSheet, existingOutbox.rowNumber, saidaState.ctx, 'Id Mensagem Gmail', '');
      coreMailHubWriteCell_(saidaSheet, existingOutbox.rowNumber, saidaState.ctx, 'Enviado Em', '');
      coreMailHubWriteCell_(saidaSheet, existingOutbox.rowNumber, saidaState.ctx, 'Atualizado Em', now);
      coreMailHubWriteCell_(saidaSheet, existingOutbox.rowNumber, saidaState.ctx, 'Observacoes', coreMailHubSerializeOutboxContract_(normalizedContract));

      return Object.freeze({
        ok: true,
        queued: true,
        requeued: true,
        saidaId: existingOutbox.saidaId,
        status: 'PENDENTE',
        subject: normalizedContract.finalSubject,
        correlationKey: normalizedContract.correlationKey,
        scheduledAt: normalizedContract.sendAfter || null
      });
    }

    var saidaId = coreMailHubGenerateOutgoingId_();

    coreMailHubAppendRow_(saidaSheet, saidaState.ctx, {
      'Id Saida': saidaId,
      'Modulo Dono': normalizedContract.moduleName,
      'Tipo Entidade': normalizedContract.entityType,
      'Id Entidade': normalizedContract.entityId,
      'Chave de Correlacao': normalizedContract.correlationKey,
      'Etapa Fluxo': normalizedContract.stage,
      'Email Destinatario Principal': normalizedContract.to[0] || '',
      'Emails Destinatarios': coreMailHubJoinEmails_(normalizedContract.to),
      'Emails Cc': coreMailHubJoinEmails_(normalizedContract.cc),
      'Emails Cco': coreMailHubJoinEmails_(normalizedContract.bcc),
      'Nome Destinatario': normalizedContract.recipientName,
      'Assunto': normalizedContract.finalSubject,
      'Corpo Texto': '',
      'Corpo Html': '',
      'Data Hora Agendada': normalizedContract.sendAfter || '',
      'Prioridade': normalizedContract.priority,
      'Status Envio': 'PENDENTE',
      'Tentativas': 0,
      'Ultimo Erro': '',
      'Id Thread Gmail': '',
      'Id Mensagem Gmail': '',
      'Enviado Em': '',
      'Criado Em': now,
      'Atualizado Em': now,
      'Observacoes': coreMailHubSerializeOutboxContract_(normalizedContract)
    });

    return Object.freeze({
      ok: true,
      queued: true,
      saidaId: saidaId,
      status: 'PENDENTE',
      subject: normalizedContract.finalSubject,
      correlationKey: normalizedContract.correlationKey,
      scheduledAt: normalizedContract.sendAfter || null
    });
  });
}

function coreMailProcessOutbox_() {
  return core_withLock_('CORE_MAIL_HUB_PROCESS_OUTBOX', function() {
    coreMailHubAssertSchema_();

    var startedAt = new Date();
    var runId = core_runId_();
    var now = new Date();
    var saidaSheet = coreMailHubGetSaidaSheet_();
    var eventosSheet = coreMailHubGetEventosSheet_();
    var indiceSheet = coreMailHubGetIndiceSheet_();
    var saidaState = coreMailHubBuildSaidaState_(saidaSheet);
    var eventosState = coreMailHubBuildEventosState_(eventosSheet);
    var indiceState = coreMailHubBuildIndiceState_(indiceSheet);
    var counters = {
      scanned: 0,
      pending: 0,
      sent: 0,
      skippedScheduled: 0,
      skippedStatus: 0,
      errors: 0,
      eventLogs: 0,
      indexUpserts: 0
    };

    for (var i = 0; i < saidaState.ctx.rows.length; i++) {
      var row = saidaState.ctx.rows[i];
      var rowNumber = saidaState.ctx.startRow + i;
      var record = coreMailHubBuildOutboxRecord_(row, saidaState.ctx, rowNumber);
      counters.scanned++;

      if (record.status !== 'PENDENTE') {
        counters.skippedStatus++;
        continue;
      }

      counters.pending++;

      var scheduledAt = coreMailHubCoerceDate_(record.scheduledAt);
      if (scheduledAt && scheduledAt.getTime() > now.getTime()) {
        counters.skippedScheduled++;
        continue;
      }

      try {
        var safeTemplateKey = coreMailHubNormalizeTemplateKey_(
          record.observacoes && record.observacoes.templateKey
            ? record.observacoes.templateKey
            : ''
        );
        var draftContract = Object.assign({}, record.observacoes || {}, {
          moduleName: record.moduleName,
          correlationKey: record.correlationKey,
          entityType: record.entityType,
          entityId: record.entityId,
          flowCode: record.observacoes && record.observacoes.flowCode ? record.observacoes.flowCode : '',
          stage: record.flowStep || (record.observacoes && record.observacoes.stage) || '',
          to: record.to,
          cc: record.cc,
          bcc: record.bcc,
          subjectHuman: record.observacoes && record.observacoes.subjectHuman ? record.observacoes.subjectHuman : record.subject,
          templateKey: safeTemplateKey,
          payload: record.observacoes && record.observacoes.payload ? record.observacoes.payload : {},
          metadata: record.observacoes && record.observacoes.metadata ? record.observacoes.metadata : {},
          priority: record.priority,
          name: record.recipientName,
          replyTo: record.observacoes && record.observacoes.replyTo ? record.observacoes.replyTo : '',
          attachments: record.observacoes && record.observacoes.attachmentRefs ? record.observacoes.attachmentRefs : []
        });

        var draft = coreMailBuildOutgoingDraft_(draftContract);
        var resolvedAttachments = coreMailHubResolveOutgoingAttachments_(draftContract.attachments);
        var sentAt = new Date();
        var sendResult = coreMailHubSendOutgoingMessage_({
          to: draft.toList,
          cc: draft.ccList,
          bcc: draft.bccList,
          subject: draft.subject,
          body: draft.bodyText,
          htmlBody: draft.htmlBody,
          replyTo: draftContract.replyTo || '',
          attachments: resolvedAttachments
        });

        coreMailHubWriteCell_(saidaSheet, rowNumber, saidaState.ctx, 'Email Destinatario Principal', draft.toList[0] || '');
        coreMailHubWriteCell_(saidaSheet, rowNumber, saidaState.ctx, 'Emails Destinatarios', coreMailHubJoinEmails_(draft.toList));
        coreMailHubWriteCell_(saidaSheet, rowNumber, saidaState.ctx, 'Emails Cc', coreMailHubJoinEmails_(draft.ccList));
        coreMailHubWriteCell_(saidaSheet, rowNumber, saidaState.ctx, 'Emails Cco', coreMailHubJoinEmails_(draft.bccList));
        coreMailHubWriteCell_(saidaSheet, rowNumber, saidaState.ctx, 'Assunto', draft.subject);
        coreMailHubWriteCell_(saidaSheet, rowNumber, saidaState.ctx, 'Corpo Texto', draft.bodyText || '');
        coreMailHubWriteCell_(saidaSheet, rowNumber, saidaState.ctx, 'Corpo Html', draft.htmlBody || '');
        coreMailHubWriteCell_(saidaSheet, rowNumber, saidaState.ctx, 'Status Envio', 'ENVIADO');
        coreMailHubWriteCell_(saidaSheet, rowNumber, saidaState.ctx, 'Tentativas', record.attempts + 1);
        coreMailHubWriteCell_(saidaSheet, rowNumber, saidaState.ctx, 'Ultimo Erro', '');
        coreMailHubWriteCell_(saidaSheet, rowNumber, saidaState.ctx, 'Id Thread Gmail', sendResult.threadId || '');
        coreMailHubWriteCell_(saidaSheet, rowNumber, saidaState.ctx, 'Id Mensagem Gmail', sendResult.messageId || '');
        coreMailHubWriteCell_(saidaSheet, rowNumber, saidaState.ctx, 'Enviado Em', sentAt);
        coreMailHubWriteCell_(saidaSheet, rowNumber, saidaState.ctx, 'Atualizado Em', sentAt);

        if (sendResult.postSendWarning) {
          core_logWarn_(runId, 'coreMailProcessOutbox: envio concluido com falha no enriquecimento Gmail', {
            saidaId: record.saidaId,
            correlationKey: record.correlationKey,
            warning: sendResult.postSendWarning
          });
        }

        var eventPayload = coreMailHubBuildOutgoingEventPayload_(record, draft, sendResult, sentAt);
        coreMailRegisterEvent_(eventosSheet, eventPayload, eventosState);
        counters.eventLogs++;

        var upsertInfo = coreMailUpsertIndex_(indiceSheet, eventPayload, indiceState);
        if (upsertInfo.updated) counters.indexUpserts++;

        counters.sent++;
      } catch (err) {
        var failedAt = new Date();
        coreMailHubWriteCell_(saidaSheet, rowNumber, saidaState.ctx, 'Status Envio', 'ERRO');
        coreMailHubWriteCell_(saidaSheet, rowNumber, saidaState.ctx, 'Tentativas', record.attempts + 1);
        coreMailHubWriteCell_(saidaSheet, rowNumber, saidaState.ctx, 'Ultimo Erro', err && err.message ? err.message : String(err || ''));
        coreMailHubWriteCell_(saidaSheet, rowNumber, saidaState.ctx, 'Atualizado Em', failedAt);
        counters.errors++;
        core_logError_(runId, 'coreMailProcessOutbox: erro ao processar saida', {
          saidaId: record.saidaId,
          correlationKey: record.correlationKey,
          error: err && err.message ? err.message : String(err || '')
        });
      }
    }

    core_logSummarize_(runId, 'coreMailProcessOutbox_', startedAt, counters);

    return Object.freeze({
      ok: true,
      processedAt: now,
      counters: Object.freeze(counters)
    });
  });
}

function core_mailIngestInbox_(opts) {
  opts = opts || {};

  return core_withLock_('CORE_MAIL_HUB_INGEST_INBOX', function() {
    var runId = core_runId_();
    var startedAt = new Date();
    var configMap = coreMailHubGetConfigMap_();
    var query = String(
      Object.prototype.hasOwnProperty.call(opts, 'query')
        ? opts.query
        : coreMailHubGetConfigValue_(configMap, 'GMAIL_QUERY_INGEST', CORE_MAIL_HUB_DEFAULTS.query)
    );
    var start = Number(
      Object.prototype.hasOwnProperty.call(opts, 'start')
        ? opts.start
        : coreMailHubGetConfigNumber_(configMap, 'GMAIL_START', CORE_MAIL_HUB_DEFAULTS.start)
    );
    var maxThreads = Number(
      Object.prototype.hasOwnProperty.call(opts, 'maxThreads')
        ? opts.maxThreads
        : coreMailHubGetConfigNumber_(configMap, 'GMAIL_MAX_THREADS', CORE_MAIL_HUB_DEFAULTS.maxThreads)
    );
    var maxMessagesPerThread = Number(
      Object.prototype.hasOwnProperty.call(opts, 'maxMessagesPerThread')
        ? opts.maxMessagesPerThread
        : coreMailHubGetConfigNumber_(configMap, 'GMAIL_MAX_MESSAGES_PER_THREAD', CORE_MAIL_HUB_DEFAULTS.maxMessagesPerThread)
    );
    var dryRun = opts.dryRun === true;
    var ingestConfig = coreMailHubBuildIngestConfig_(configMap, opts);

    query = String(query || '').trim() || CORE_MAIL_HUB_DEFAULTS.query;
    start = isNaN(start) || start < 0 ? CORE_MAIL_HUB_DEFAULTS.start : start;
    maxThreads = isNaN(maxThreads) || maxThreads < 1 ? CORE_MAIL_HUB_DEFAULTS.maxThreads : maxThreads;
    maxMessagesPerThread =
      isNaN(maxMessagesPerThread) || maxMessagesPerThread < 1
        ? CORE_MAIL_HUB_DEFAULTS.maxMessagesPerThread
        : maxMessagesPerThread;

    coreMailHubAssertSchema_();

    var eventosSheet = coreMailHubGetEventosSheet_();
    var indiceSheet = coreMailHubGetIndiceSheet_();
    var anexosSheet = coreMailHubGetAnexosSheet_();

    var eventosState = coreMailHubBuildEventosState_(eventosSheet);
    var indiceState = coreMailHubBuildIndiceState_(indiceSheet);
    var anexosCtx = coreMailHubGetSheetContext_(anexosSheet);

    var threads = GmailApp.search(query, start, maxThreads);
    var counters = {
      threadsScanned: threads.length,
      messagesScanned: 0,
      newEvents: 0,
      duplicatesSkipped: 0,
      noiseSkipped: 0,
      ignoredNoiseEvents: 0,
      attachmentsRegistered: 0,
      indexUpserts: 0,
      withoutCorrelationKey: 0,
      executionLimitReached: false,
      errors: 0
    };

    core_logInfo_(runId, 'Mail Hub ingest iniciado', {
      query: query,
      start: start,
      maxThreads: maxThreads,
      maxMessagesPerThread: maxMessagesPerThread,
      dryRun: dryRun,
      ingestConfig: {
        requiredSubjectPrefix: ingestConfig.requiredSubjectPrefix,
        useOnlyTaggedSubjects: ingestConfig.useOnlyTaggedSubjects,
        ignoredSendersCount: ingestConfig.ignoredSenders.length,
        ignoredDomainsCount: ingestConfig.ignoredDomains.length,
        ignoredSubjectRegexCount: ingestConfig.ignoredSubjectRegexes.length,
        maxEventsPerExecution: ingestConfig.maxEventsPerExecution,
        saveFullBody: ingestConfig.saveFullBody,
        markNoiseAsIgnored: ingestConfig.markNoiseAsIgnored
      }
    });

    outerLoop:
    for (var i = 0; i < threads.length; i++) {
      var thread = threads[i];
      var messages = thread.getMessages();
      var startIndex = Math.max(0, messages.length - maxMessagesPerThread);

      for (var j = startIndex; j < messages.length; j++) {
        if (ingestConfig.maxEventsPerExecution > 0 && counters.newEvents >= ingestConfig.maxEventsPerExecution) {
          counters.executionLimitReached = true;
          break outerLoop;
        }

        var message = messages[j];
        counters.messagesScanned++;

        try {
          var eventPayload = coreMailHubBuildEventPayload_(thread, message, new Date(), ingestConfig);
          if (!eventPayload.messageId) {
            counters.errors++;
            core_logWarn_(runId, 'Mail Hub pulou mensagem sem messageId', {
              threadId: String(thread.getId() || '').trim(),
              subject: String(message.getSubject() || '').trim()
            });
            continue;
          }

          if (coreMailEventExistsByMessageId_(eventPayload.messageId, eventosState)) {
            counters.duplicatesSkipped++;
            continue;
          }

          if (eventPayload.isNoise && !eventPayload.shouldIgnore) {
            counters.noiseSkipped++;
            continue;
          }

          if (!eventPayload.correlationKey) {
            counters.withoutCorrelationKey++;
          }

          if (dryRun) {
            counters.newEvents++;
            if (eventPayload.shouldIgnore) {
              counters.ignoredNoiseEvents++;
              if (eventPayload.correlationKey) counters.indexUpserts++;
            } else {
              counters.attachmentsRegistered += eventPayload.attachmentCount;
              if (eventPayload.correlationKey) counters.indexUpserts++;
            }
            continue;
          }

          var registeredEvent = coreMailRegisterEvent_(eventosSheet, eventPayload, eventosState);
          eventPayload.eventId = registeredEvent.eventId;

          counters.newEvents++;
          if (eventPayload.shouldIgnore) {
            counters.ignoredNoiseEvents++;
            if (eventPayload.correlationKey) {
              var ignoredUpsertInfo = coreMailUpsertIndex_(indiceSheet, eventPayload, indiceState);
              if (ignoredUpsertInfo.updated) counters.indexUpserts++;
            }
          } else {
            counters.attachmentsRegistered += coreMailRegisterAttachments_(anexosSheet, eventPayload, anexosCtx);

            var upsertInfo = coreMailUpsertIndex_(indiceSheet, eventPayload, indiceState);
            if (upsertInfo.updated) counters.indexUpserts++;
          }
        } catch (err) {
          counters.errors++;
          core_logError_(runId, 'Mail Hub erro ao ingerir mensagem', {
            error: err && err.message ? err.message : String(err || ''),
            threadId: String(thread.getId() || '').trim(),
            messageIndex: j
          });
        }
      }
    }

    core_logSummarize_(runId, 'core_mailIngestInbox_', startedAt, counters);

    return Object.freeze({
      ok: true,
      dryRun: dryRun,
      query: query,
      start: start,
      maxThreads: maxThreads,
      maxMessagesPerThread: maxMessagesPerThread,
      ingestConfig: Object.freeze({
        requiredSubjectPrefix: ingestConfig.requiredSubjectPrefix,
        useOnlyTaggedSubjects: ingestConfig.useOnlyTaggedSubjects,
        maxEventsPerExecution: ingestConfig.maxEventsPerExecution,
        saveFullBody: ingestConfig.saveFullBody,
        markNoiseAsIgnored: ingestConfig.markNoiseAsIgnored
      }),
      counters: Object.freeze(counters)
    });
  });
}

function core_mailListPendingByModule_(moduleName) {
  core_assertRequired_(moduleName, 'moduleName');
  return coreMailHubListEvents_({
    moduleName: moduleName,
    processingStatus: CORE_MAIL_HUB_STATUS.PENDENTE
  });
}

function coreMailHubNormalizeOptionalFilter_(value) {
  if (value === '' || value === null || typeof value === 'undefined') return '';
  return core_normalizeText_(value, {
    collapseWhitespace: true,
    caseMode: 'upper'
  });
}

function coreMailHubBuildEventRecord_(row, ctx, rowNumber) {
  return {
    eventId: String(coreMailHubGetRowValue_(row, ctx, 'Id Evento', '') || '').trim(),
    messageId: String(coreMailHubGetRowValue_(row, ctx, 'Id Mensagem Gmail', '') || '').trim(),
    threadId: String(coreMailHubGetRowValue_(row, ctx, 'Id Thread Gmail', '') || '').trim(),
    correlationKey: String(coreMailHubGetRowValue_(row, ctx, 'Chave de Correlacao', '') || '').trim(),
    direction: String(coreMailHubGetRowValue_(row, ctx, 'Direcao', '') || '').trim(),
    moduleName: coreMailHubNormalizeOptionalFilter_(coreMailHubGetRowValue_(row, ctx, 'Modulo Dono', '')),
    entityType: String(coreMailHubGetRowValue_(row, ctx, 'Tipo Entidade', '') || '').trim(),
    entityId: String(coreMailHubGetRowValue_(row, ctx, 'Id Entidade', '') || '').trim(),
    flowStep: String(coreMailHubGetRowValue_(row, ctx, 'Etapa Fluxo', '') || '').trim(),
    eventType: String(coreMailHubGetRowValue_(row, ctx, 'Tipo Evento', '') || '').trim(),
    subject: String(coreMailHubGetRowValue_(row, ctx, 'Assunto', '') || '').trim(),
    fromEmail: String(coreMailHubGetRowValue_(row, ctx, 'Email Remetente', '') || '').trim(),
    fromName: String(coreMailHubGetRowValue_(row, ctx, 'Nome Remetente', '') || '').trim(),
    to: String(coreMailHubGetRowValue_(row, ctx, 'Emails Destinatarios', '') || '').trim(),
    cc: String(coreMailHubGetRowValue_(row, ctx, 'Emails Cc', '') || '').trim(),
    bcc: String(coreMailHubGetRowValue_(row, ctx, 'Emails Cco', '') || '').trim(),
    snippet: String(coreMailHubGetRowValue_(row, ctx, 'Trecho Corpo', '') || '').trim(),
    plainBody: String(coreMailHubGetRowValue_(row, ctx, 'Corpo Texto', '') || '').trim(),
    processingStatus: coreMailHubNormalizeFlag_(coreMailHubGetRowValue_(row, ctx, 'Status Processamento', '')),
    routingStatus: coreMailHubNormalizeFlag_(coreMailHubGetRowValue_(row, ctx, 'Status Roteamento', '')),
    receivedAt: coreMailHubGetRowValue_(row, ctx, 'Data Hora Evento', ''),
    ingestedAt: coreMailHubGetRowValue_(row, ctx, 'Criado Em', ''),
    updatedAt: coreMailHubGetRowValue_(row, ctx, 'Atualizado Em', ''),
    hasAttachments: String(coreMailHubGetRowValue_(row, ctx, 'Possui Anexos', '') || '').trim(),
    attachmentCount: Number(coreMailHubGetRowValue_(row, ctx, 'Quantidade Anexos', 0) || 0),
    processedBy: String(coreMailHubGetRowValue_(row, ctx, 'Processado Por', '') || '').trim(),
    processedAt: coreMailHubGetRowValue_(row, ctx, 'Data Hora Processamento', ''),
    observations: String(coreMailHubGetRowValue_(row, ctx, 'Observacoes', '') || '').trim(),
    rawJson: String(coreMailHubGetRowValue_(row, ctx, 'Json Bruto', '') || '').trim(),
    rowNumber: rowNumber
  };
}

function coreMailHubNormalizeAttachmentStatus_(value) {
  var normalized = coreMailHubNormalizeFlag_(value);
  if (
    normalized === CORE_MAIL_ATTACHMENT_STATUS.PENDENTE ||
    normalized === CORE_MAIL_ATTACHMENT_STATUS.PROCESSADO ||
    normalized === CORE_MAIL_ATTACHMENT_STATUS.SALVO_DRIVE ||
    normalized === CORE_MAIL_ATTACHMENT_STATUS.IGNORADO ||
    normalized === CORE_MAIL_ATTACHMENT_STATUS.ERRO
  ) {
    return normalized;
  }
  return CORE_MAIL_ATTACHMENT_STATUS.PENDENTE;
}

function coreMailHubIsPendingAttachmentStatus_(status) {
  var normalized = coreMailHubNormalizeAttachmentStatus_(status);
  return normalized !== CORE_MAIL_ATTACHMENT_STATUS.PROCESSADO &&
    normalized !== CORE_MAIL_ATTACHMENT_STATUS.SALVO_DRIVE &&
    normalized !== CORE_MAIL_ATTACHMENT_STATUS.IGNORADO;
}

function coreMailHubExtractAttachmentIndex_(record) {
  var direct = Number(record.messageAttachmentIndex || 0);
  if (!isNaN(direct) && direct > 0) return direct;

  var match = String(record.observations || '').match(/Attachment Index\s*=\s*(\d+)/i);
  return match ? (parseInt(match[1], 10) || 0) : 0;
}

function coreMailHubBuildAttachmentRecord_(row, ctx, rowNumber) {
  return {
    attachmentId: String(coreMailHubGetRowValue_(row, ctx, 'Id Anexo', '') || '').trim(),
    eventId: String(coreMailHubGetRowValue_(row, ctx, 'Id Evento', '') || '').trim(),
    moduleName: coreMailHubNormalizeOptionalFilter_(coreMailHubGetRowValue_(row, ctx, 'Modulo Dono', '')),
    entityType: String(coreMailHubGetRowValue_(row, ctx, 'Tipo Entidade', '') || '').trim(),
    entityId: String(coreMailHubGetRowValue_(row, ctx, 'Id Entidade', '') || '').trim(),
    correlationKey: String(coreMailHubGetRowValue_(row, ctx, 'Chave de Correlacao', '') || '').trim(),
    flowStep: String(coreMailHubGetRowValue_(row, ctx, 'Etapa Fluxo', '') || '').trim(),
    messageId: String(coreMailHubGetRowValue_(row, ctx, 'Id Mensagem Gmail', '') || '').trim(),
    threadId: String(coreMailHubGetRowValue_(row, ctx, 'Id Thread Gmail', '') || '').trim(),
    messageAttachmentIndex: Number(coreMailHubGetRowValue_(row, ctx, 'Indice Anexo Mensagem', 0) || 0),
    fileName: String(coreMailHubGetRowValue_(row, ctx, 'Nome Arquivo', '') || '').trim(),
    mimeType: String(coreMailHubGetRowValue_(row, ctx, 'Tipo Mime', '') || '').trim(),
    sizeBytes: Number(coreMailHubGetRowValue_(row, ctx, 'Tamanho Bytes', 0) || 0),
    savedToDrive: coreMailHubNormalizeFlag_(coreMailHubGetRowValue_(row, ctx, 'Foi Salvo No Drive', 'NAO')),
    driveFileId: String(coreMailHubGetRowValue_(row, ctx, 'Id Arquivo Drive', '') || '').trim(),
    driveFileUrl: String(coreMailHubGetRowValue_(row, ctx, 'Link Arquivo Drive', '') || '').trim(),
    driveFolder: String(coreMailHubGetRowValue_(row, ctx, 'Pasta Destino Drive', '') || '').trim(),
    attachmentStatus: coreMailHubNormalizeAttachmentStatus_(coreMailHubGetRowValue_(row, ctx, 'Status Anexo', 'PENDENTE')),
    processedBy: String(coreMailHubGetRowValue_(row, ctx, 'Processado Por', '') || '').trim(),
    processedAt: coreMailHubGetRowValue_(row, ctx, 'Data Hora Processamento', ''),
    observations: String(coreMailHubGetRowValue_(row, ctx, 'Observacoes', '') || '').trim(),
    createdAt: coreMailHubGetRowValue_(row, ctx, 'Criado Em', ''),
    updatedAt: coreMailHubGetRowValue_(row, ctx, 'Atualizado Em', ''),
    rowNumber: rowNumber
  };
}

function coreMailHubListAttachments_(opts) {
  opts = opts || {};
  coreMailHubAssertSchema_();

  var sheet = coreMailHubGetAnexosSheet_();
  var ctx = coreMailHubGetSheetContext_(sheet, { includeRows: true });
  var out = [];
  var targetModule = coreMailHubNormalizeOptionalFilter_(opts.moduleName);
  var targetCorrelationKey = coreMailHubNormalizeOptionalFilter_(opts.correlationKey);
  var targetEntityType = coreMailHubNormalizeOptionalFilter_(opts.entityType);
  var targetEntityId = String(opts.entityId || '').trim();
  var targetFlowStep = coreMailHubNormalizeOptionalFilter_(opts.flowStep);
  var targetStatus = coreMailHubNormalizeOptionalFilter_(opts.statusAnexo || opts.attachmentStatus);
  var targetEventId = String(opts.eventId || '').trim();
  var targetMessageId = String(opts.messageId || '').trim();
  var targetThreadId = String(opts.threadId || '').trim();
  var targetAttachmentId = String(opts.attachmentId || '').trim();
  var hasPendingAttachment = Object.prototype.hasOwnProperty.call(opts, 'hasPendingAttachment')
    ? opts.hasPendingAttachment === true
    : null;
  var limit = Number(opts.limit || 0);

  for (var i = 0; i < ctx.rows.length; i++) {
    var record = coreMailHubBuildAttachmentRecord_(ctx.rows[i], ctx, ctx.startRow + i);

    if (targetModule && record.moduleName !== targetModule) continue;
    if (targetCorrelationKey && coreMailHubNormalizeOptionalFilter_(record.correlationKey) !== targetCorrelationKey) continue;
    if (targetEntityType && coreMailHubNormalizeOptionalFilter_(record.entityType) !== targetEntityType) continue;
    if (targetEntityId && record.entityId !== targetEntityId) continue;
    if (targetFlowStep && coreMailHubNormalizeOptionalFilter_(record.flowStep) !== targetFlowStep) continue;
    if (targetStatus && coreMailHubNormalizeOptionalFilter_(record.attachmentStatus) !== targetStatus) continue;
    if (targetEventId && record.eventId !== targetEventId) continue;
    if (targetMessageId && record.messageId !== targetMessageId) continue;
    if (targetThreadId && record.threadId !== targetThreadId) continue;
    if (targetAttachmentId && record.attachmentId !== targetAttachmentId) continue;
    if (hasPendingAttachment === true && !coreMailHubIsPendingAttachmentStatus_(record.attachmentStatus)) continue;
    if (hasPendingAttachment === false && coreMailHubIsPendingAttachmentStatus_(record.attachmentStatus)) continue;

    out.push(record);
  }

  out.sort(function(a, b) {
    var aDate = coreMailHubCoerceDate_(a.updatedAt) || coreMailHubCoerceDate_(a.createdAt) || coreMailHubCoerceDate_(a.processedAt);
    var bDate = coreMailHubCoerceDate_(b.updatedAt) || coreMailHubCoerceDate_(b.createdAt) || coreMailHubCoerceDate_(b.processedAt);
    var aTime = aDate ? aDate.getTime() : 0;
    var bTime = bDate ? bDate.getTime() : 0;
    return bTime - aTime;
  });

  if (limit > 0) {
    return out.slice(0, limit);
  }

  return out;
}

function coreMailHubGetLatestPendingEventWithAttachment_(opts) {
  var attachments = coreMailHubListAttachments_(Object.assign({}, opts || {}, {
    hasPendingAttachment: true,
    limit: 50
  }));

  if (!attachments.length) return null;

  var seenEvent = Object.create(null);
  for (var i = 0; i < attachments.length; i++) {
    var attachment = attachments[i];
    if (!attachment.eventId || seenEvent[attachment.eventId]) continue;
    seenEvent[attachment.eventId] = true;

    var eventRecord = core_mailGetLatestEvent_({
      messageId: attachment.messageId
    });
    if (!eventRecord) continue;

    return Object.freeze({
      event: eventRecord,
      attachments: attachments.filter(function(item) {
        return item.eventId === attachment.eventId;
      })
    });
  }

  return null;
}

function coreMailHubGetThreadById_(threadId) {
  var normalized = String(threadId || '').trim();
  if (!normalized) throw new Error('threadId obrigatorio para recuperar anexos do Gmail.');
  return GmailApp.getThreadById(normalized);
}

function coreMailHubGetMessageByIds_(threadId, messageId) {
  var normalizedMessageId = String(messageId || '').trim();
  var thread = coreMailHubGetThreadById_(threadId);
  var messages = thread.getMessages() || [];

  for (var i = 0; i < messages.length; i++) {
    var candidate = messages[i];
    if (String(candidate.getId() || '').trim() === normalizedMessageId) {
      return candidate;
    }
  }

  throw new Error('Mensagem nao encontrada na thread informada: ' + normalizedMessageId);
}

function coreMailHubFindRealAttachmentByRecord_(record) {
  if (!record || !record.threadId || !record.messageId) {
    throw new Error('Registro de anexo sem threadId/messageId suficiente para reabrir o Gmail.');
  }

  var message = coreMailHubGetMessageByIds_(record.threadId, record.messageId);
  var attachments = coreMailHubGetMessageAttachments_(message);
  var attachmentIndex = coreMailHubExtractAttachmentIndex_(record);
  var selected = null;

  if (attachmentIndex > 0 && attachmentIndex <= attachments.length) {
    selected = attachments[attachmentIndex - 1];
  }

  if (!selected) {
    for (var i = 0; i < attachments.length; i++) {
      var candidate = attachments[i];
      var sameName = String(candidate.getName() || '').trim() === record.fileName;
      var sameType = String(candidate.getContentType() || '').trim() === record.mimeType;
      if (!sameName) continue;
      if (record.mimeType && !sameType) continue;
      selected = candidate;
      attachmentIndex = i + 1;
      break;
    }
  }

  if (!selected) {
    throw new Error('Anexo real nao encontrado no Gmail para o registro ' + record.attachmentId + '.');
  }

  return Object.freeze({
    attachmentId: record.attachmentId,
    eventId: record.eventId,
    moduleName: record.moduleName,
    entityType: record.entityType,
    entityId: record.entityId,
    correlationKey: record.correlationKey,
    flowStep: record.flowStep,
    messageId: record.messageId,
    threadId: record.threadId,
    fileName: String(selected.getName() || '').trim(),
    mimeType: String(selected.getContentType() || '').trim(),
    sizeBytes: record.sizeBytes,
    messageAttachmentIndex: attachmentIndex,
    attachment: selected,
    blob: selected.copyBlob()
  });
}

function coreMailHubUpdateAttachmentStatus_(attachmentId, processorName, patch) {
  core_assertRequired_(attachmentId, 'attachmentId');
  core_assertRequired_(processorName, 'processorName');
  patch = patch || {};

  return core_withLock_('CORE_MAIL_HUB_MARK_ATTACHMENT', function() {
    coreMailHubAssertSchema_();

    var sheet = coreMailHubGetAnexosSheet_();
    var ctx = coreMailHubGetSheetContext_(sheet, { includeRows: true });
    var normalizedAttachmentId = String(attachmentId || '').trim();
    var rowNumber = 0;
    var record = null;

    for (var i = 0; i < ctx.rows.length; i++) {
      var candidate = coreMailHubBuildAttachmentRecord_(ctx.rows[i], ctx, ctx.startRow + i);
      if (candidate.attachmentId === normalizedAttachmentId) {
        rowNumber = candidate.rowNumber;
        record = candidate;
        break;
      }
    }

    if (!rowNumber || !record) {
      throw new Error('Anexo nao encontrado em MAIL_ANEXOS: ' + attachmentId);
    }

    var processedAt = patch.processedAt || new Date();
    var nextStatus = coreMailHubNormalizeAttachmentStatus_(patch.statusAnexo || patch.attachmentStatus || record.attachmentStatus);
    var nextSavedFlag = Object.prototype.hasOwnProperty.call(patch, 'savedToDrive')
      ? (patch.savedToDrive === true ? 'SIM' : 'NAO')
      : (nextStatus === CORE_MAIL_ATTACHMENT_STATUS.SALVO_DRIVE ? 'SIM' : record.savedToDrive || 'NAO');
    var nextObservations = Object.prototype.hasOwnProperty.call(patch, 'observations')
      ? String(patch.observations || '').trim()
      : record.observations;

    coreMailHubWriteCell_(sheet, rowNumber, ctx, 'Status Anexo', nextStatus);
    coreMailHubWriteCell_(sheet, rowNumber, ctx, 'Foi Salvo No Drive', nextSavedFlag);
    coreMailHubWriteCell_(sheet, rowNumber, ctx, 'Processado Por', String(processorName || '').trim());
    coreMailHubWriteCell_(sheet, rowNumber, ctx, 'Data Hora Processamento', processedAt);
    coreMailHubWriteCell_(sheet, rowNumber, ctx, 'Observacoes', nextObservations);
    coreMailHubWriteCell_(sheet, rowNumber, ctx, 'Atualizado Em', processedAt);

    if (Object.prototype.hasOwnProperty.call(patch, 'driveFileId')) {
      coreMailHubWriteCell_(sheet, rowNumber, ctx, 'Id Arquivo Drive', String(patch.driveFileId || '').trim());
    }
    if (Object.prototype.hasOwnProperty.call(patch, 'driveFileUrl')) {
      coreMailHubWriteCell_(sheet, rowNumber, ctx, 'Link Arquivo Drive', String(patch.driveFileUrl || '').trim());
    }
    if (Object.prototype.hasOwnProperty.call(patch, 'driveFolder')) {
      coreMailHubWriteCell_(sheet, rowNumber, ctx, 'Pasta Destino Drive', String(patch.driveFolder || '').trim());
    }

    if (record.correlationKey) {
      coreMailHubRefreshIndexSummaryByCorrelationKey_(record.correlationKey);
    }

    return Object.freeze({
      ok: true,
      attachmentId: normalizedAttachmentId,
      rowNumber: rowNumber,
      statusAnexo: nextStatus,
      savedToDrive: nextSavedFlag,
      processedBy: String(processorName || '').trim(),
      processedAt: processedAt,
      correlationKey: record.correlationKey
    });
  });
}

function coreMailHubListEvents_(opts) {
  opts = opts || {};
  coreMailHubAssertSchema_();

  var sheet = coreMailHubGetEventosSheet_();
  var ctx = coreMailHubGetSheetContext_(sheet, { includeRows: true });
  var out = [];
  var targetModule = coreMailHubNormalizeOptionalFilter_(opts.moduleName);
  var targetStatus = coreMailHubNormalizeOptionalFilter_(opts.processingStatus);
  var targetCorrelationKey = coreMailHubNormalizeOptionalFilter_(opts.correlationKey);
  var targetMessageId = String(opts.messageId || '').trim();
  var targetThreadId = String(opts.threadId || '').trim();
  var limit = Number(opts.limit || 0);

  for (var i = 0; i < ctx.rows.length; i++) {
    var row = ctx.rows[i];
    var record = coreMailHubBuildEventRecord_(row, ctx, ctx.startRow + i);

    if (targetModule && record.moduleName !== targetModule) continue;
    if (targetStatus && record.processingStatus !== targetStatus) continue;
    if (targetCorrelationKey && coreMailHubNormalizeOptionalFilter_(record.correlationKey) !== targetCorrelationKey) continue;
    if (targetMessageId && record.messageId !== targetMessageId) continue;
    if (targetThreadId && record.threadId !== targetThreadId) continue;

    out.push(record);
  }

  out.sort(function(a, b) {
    var aDate = coreMailHubCoerceDate_(a.receivedAt) || coreMailHubCoerceDate_(a.ingestedAt);
    var bDate = coreMailHubCoerceDate_(b.receivedAt) || coreMailHubCoerceDate_(b.ingestedAt);
    var aTime = aDate ? aDate.getTime() : 0;
    var bTime = bDate ? bDate.getTime() : 0;
    return bTime - aTime;
  });

  if (limit > 0) {
    return out.slice(0, limit);
  }

  return out;
}

function core_mailGetLatestEvent_(opts) {
  var events = coreMailHubListEvents_(Object.assign({}, opts || {}, { limit: 1 }));
  return events.length ? events[0] : null;
}

function core_mailListAttachments_(opts) {
  return coreMailHubListAttachments_(opts || {});
}

function core_mailListPendingAttachments_(opts) {
  return coreMailHubListAttachments_(Object.assign({}, opts || {}, {
    hasPendingAttachment: true
  }));
}

function core_mailGetLatestPendingEventWithAttachment_(opts) {
  return coreMailHubGetLatestPendingEventWithAttachment_(opts || {});
}

function core_mailListAttachmentsByEvent_(eventId, opts) {
  core_assertRequired_(eventId, 'eventId');
  return coreMailHubListAttachments_(Object.assign({}, opts || {}, {
    eventId: eventId
  }));
}

function core_mailGetAttachmentById_(attachmentId, opts) {
  core_assertRequired_(attachmentId, 'attachmentId');
  opts = opts || {};

  var records = coreMailHubListAttachments_({
    attachmentId: attachmentId,
    limit: 1
  });
  if (!records.length) return null;

  var record = records[0];
  if (opts.includeAttachment !== true && opts.includeBlob !== true) {
    return record;
  }

  return coreMailHubFindRealAttachmentByRecord_(record);
}

function core_mailGetAttachmentsByEvent_(eventId, opts) {
  core_assertRequired_(eventId, 'eventId');
  opts = opts || {};

  var records = coreMailHubListAttachments_({
    eventId: eventId
  });
  if (opts.includeAttachment !== true && opts.includeBlob !== true) {
    return records;
  }

  return records.map(function(record) {
    return coreMailHubFindRealAttachmentByRecord_(record);
  });
}

function core_mailMarkAttachmentProcessed_(attachmentId, processorName, observations) {
  return coreMailHubUpdateAttachmentStatus_(attachmentId, processorName, {
    statusAnexo: CORE_MAIL_ATTACHMENT_STATUS.PROCESSADO,
    observations: observations
  });
}

function core_mailMarkAttachmentSavedToDrive_(attachmentId, processorName, driveInfo) {
  driveInfo = driveInfo || {};
  return coreMailHubUpdateAttachmentStatus_(attachmentId, processorName, {
    statusAnexo: CORE_MAIL_ATTACHMENT_STATUS.SALVO_DRIVE,
    savedToDrive: true,
    driveFileId: driveInfo.driveFileId || driveInfo.fileId || '',
    driveFileUrl: driveInfo.driveFileUrl || driveInfo.fileUrl || '',
    driveFolder: driveInfo.driveFolder || driveInfo.folder || '',
    observations: driveInfo.observations
  });
}

function core_mailMarkAttachmentIgnored_(attachmentId, processorName, observations) {
  return coreMailHubUpdateAttachmentStatus_(attachmentId, processorName, {
    statusAnexo: CORE_MAIL_ATTACHMENT_STATUS.IGNORADO,
    observations: observations
  });
}

function core_mailMarkAttachmentError_(attachmentId, processorName, observations) {
  return coreMailHubUpdateAttachmentStatus_(attachmentId, processorName, {
    statusAnexo: CORE_MAIL_ATTACHMENT_STATUS.ERRO,
    observations: observations
  });
}

function core_mailMarkLatestPendingByModule_(moduleName, processorName) {
  core_assertRequired_(moduleName, 'moduleName');
  core_assertRequired_(processorName, 'processorName');

  var latestPending = core_mailGetLatestEvent_({
    moduleName: moduleName,
    processingStatus: CORE_MAIL_HUB_STATUS.PENDENTE
  });

  if (!latestPending) {
    return Object.freeze({
      ok: true,
      found: false,
      moduleName: coreMailHubNormalizeOptionalFilter_(moduleName),
      processorName: String(processorName || '').trim()
    });
  }

  var result = core_mailMarkEventProcessed_(latestPending.eventId, processorName);
  return Object.freeze({
    ok: true,
    found: true,
    latestEvent: latestPending,
    result: result
  });
}

function core_mailMarkEventProcessed_(eventId, processorName) {
  core_assertRequired_(eventId, 'eventId');
  core_assertRequired_(processorName, 'processorName');

  return core_withLock_('CORE_MAIL_HUB_MARK_EVENT_PROCESSED', function() {
    coreMailHubAssertSchema_();

    var sheet = coreMailHubGetEventosSheet_();
    var eventosState = coreMailHubBuildEventosState_(sheet);
    var normalizedEventId = String(eventId || '').trim();
    var rowNumber = eventosState.rowByEventId[normalizedEventId];

    if (!rowNumber) {
      throw new Error('Evento nao encontrado em MAIL_EVENTOS: ' + eventId);
    }

    var processedAt = new Date();
    coreMailHubWriteCell_(sheet, rowNumber, eventosState.ctx, 'Status Processamento', CORE_MAIL_HUB_STATUS.PROCESSADO);
    coreMailHubWriteCell_(sheet, rowNumber, eventosState.ctx, 'Processado Por', String(processorName || '').trim());
    coreMailHubWriteCell_(sheet, rowNumber, eventosState.ctx, 'Data Hora Processamento', processedAt);
    coreMailHubWriteCell_(sheet, rowNumber, eventosState.ctx, 'Atualizado Em', processedAt);

    var row = sheet.getRange(rowNumber, 1, 1, eventosState.ctx.lastCol).getValues()[0];
    var correlationKey = String(coreMailHubGetRowValue_(row, eventosState.ctx, 'Chave de Correlacao', '') || '').trim();
    if (correlationKey) {
      coreMailHubRefreshIndexSummaryByCorrelationKey_(correlationKey);
    }

    return Object.freeze({
      ok: true,
      eventId: normalizedEventId,
      processorName: String(processorName || '').trim(),
      status: CORE_MAIL_HUB_STATUS.PROCESSADO,
      processedAt: processedAt,
      rowNumber: rowNumber
    });
  });
}

function coreMailCleanupNoiseEvents_() {
  return core_withLock_('CORE_MAIL_HUB_CLEANUP_NOISE_EVENTS', function() {
    coreMailHubAssertSchema_();

    var configMap = coreMailHubGetConfigMap_();
    var ingestConfig = coreMailHubBuildIngestConfig_(configMap, {
      markNoiseAsIgnored: true
    });
    var sheet = coreMailHubGetEventosSheet_();
    var ctx = coreMailHubGetSheetContext_(sheet, { includeRows: true });
    var touchedCorrelationKeys = Object.create(null);
    var counters = {
      scanned: 0,
      updated: 0,
      alreadyIgnored: 0,
      skipped: 0
    };
    var processedAt = new Date();

    for (var i = 0; i < ctx.rows.length; i++) {
      var row = ctx.rows[i];
      var rowNumber = ctx.startRow + i;
      counters.scanned++;

      var msgCtx = {
        subject: String(coreMailHubGetRowValue_(row, ctx, 'Assunto', '') || '').trim(),
        fromEmail: String(coreMailHubGetRowValue_(row, ctx, 'Email Remetente', '') || '').trim(),
        fromRaw: String(coreMailHubGetRowValue_(row, ctx, 'Email Remetente', '') || '').trim(),
        snippet: String(coreMailHubGetRowValue_(row, ctx, 'Trecho Corpo', '') || '').trim(),
        plainBody: String(coreMailHubGetRowValue_(row, ctx, 'Corpo Texto', '') || '').trim()
      };
      var noise = coreMailHubDetectNoise_(msgCtx, ingestConfig);
      if (!noise.isNoise) {
        counters.skipped++;
        continue;
      }

      var currentStatus = coreMailHubNormalizeFlag_(coreMailHubGetRowValue_(row, ctx, 'Status Processamento', ''));
      if (currentStatus === CORE_MAIL_HUB_STATUS.IGNORADO) {
        counters.alreadyIgnored++;
        continue;
      }

      var currentObservacoes = String(coreMailHubGetRowValue_(row, ctx, 'Observacoes', '') || '').trim();
      var noiseNote = 'NOISE_REASON=' + noise.reason;
      var nextObservacoes = currentObservacoes
        ? (currentObservacoes.indexOf(noiseNote) >= 0 ? currentObservacoes : currentObservacoes + ' | ' + noiseNote)
        : noiseNote;

      coreMailHubWriteCell_(sheet, rowNumber, ctx, 'Status Roteamento', CORE_MAIL_HUB_STATUS.IGNORADO);
      coreMailHubWriteCell_(sheet, rowNumber, ctx, 'Status Processamento', CORE_MAIL_HUB_STATUS.IGNORADO);
      coreMailHubWriteCell_(sheet, rowNumber, ctx, 'Processado Por', 'coreMailCleanupNoiseEvents');
      coreMailHubWriteCell_(sheet, rowNumber, ctx, 'Data Hora Processamento', processedAt);
      coreMailHubWriteCell_(sheet, rowNumber, ctx, 'Observacoes', nextObservacoes);
      coreMailHubWriteCell_(sheet, rowNumber, ctx, 'Atualizado Em', processedAt);
      counters.updated++;

      var correlationKey = String(coreMailHubGetRowValue_(row, ctx, 'Chave de Correlacao', '') || '').trim();
      if (correlationKey) {
        touchedCorrelationKeys[core_normalizeText_(correlationKey, {
          collapseWhitespace: true,
          caseMode: 'upper'
        })] = true;
      }
    }

    Object.keys(touchedCorrelationKeys).forEach(function(correlationKey) {
      coreMailHubRefreshIndexSummaryByCorrelationKey_(correlationKey);
    });

    return Object.freeze({
      ok: true,
      processedAt: processedAt,
      counters: Object.freeze(counters),
      refreshedIndexKeys: Object.keys(touchedCorrelationKeys)
    });
  });
}
