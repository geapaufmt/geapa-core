/**
 * Navegacao centralizada e politica visual de rotas do Portal GEAPA.
 *
 * Esta camada apenas consome perfil/permissoes vindos da sessao resolvida pelo
 * GEAPA-CORE. O Portal nao calcula cargo, vinculo, perfil institucional ou
 * permissao administrativa.
 */
(function configurarNavegacaoPortal(global) {
  var MODELO = global.PortalGeapaModel || {};
  var PERFIS = {
    ADMIN: 'ADMIN',
    ADMIN_TECNICO: 'ADMIN_TECNICO',
    DIRETORIA: 'DIRETORIA',
    SECRETARIA: 'SECRETARIA',
    COMUNICACAO: 'COMUNICACAO',
    CONSELHO: 'CONSELHO',
    MEMBRO_INGRESSANTE: 'MEMBRO_INGRESSANTE',
    MEMBRO: 'MEMBRO',
    EGRESSO: 'EGRESSO',
    COLABORADOR: 'COLABORADOR',
    EXTERNO: 'EXTERNO',
    VISITANTE: 'VISITANTE'
  };
  var TODOS_PERFIS = [
    PERFIS.ADMIN,
    PERFIS.ADMIN_TECNICO,
    PERFIS.DIRETORIA,
    PERFIS.SECRETARIA,
    PERFIS.COMUNICACAO,
    PERFIS.CONSELHO,
    PERFIS.MEMBRO_INGRESSANTE,
    PERFIS.MEMBRO,
    PERFIS.EGRESSO,
    PERFIS.COLABORADOR,
    PERFIS.EXTERNO,
    PERFIS.VISITANTE
  ];
  var PUBLICO = [];
  var OPERACIONAIS = [PERFIS.ADMIN, PERFIS.ADMIN_TECNICO, PERFIS.DIRETORIA, PERFIS.SECRETARIA, PERFIS.COMUNICACAO];
  var LOGADOS = TODOS_PERFIS;
  var GRUPOS_MENU = [
    { id: 'publico', label: 'Área Pública', ordem: 10, abertoPadrao: false },
    { id: 'agenda', label: 'Atividades', ordem: 20, abertoPadrao: false },
    { id: 'meu-vinculo', label: 'Meu Vínculo', ordem: 30, abertoPadrao: false },
    { id: 'gestao-geapa', label: 'Gestão do GEAPA', ordem: 40, abertoPadrao: false }
  ];
  var MOTIVOS_ACESSO = {
    NOT_AUTHENTICATED: 'NOT_AUTHENTICATED',
    PORTAL_INATIVO: 'PORTAL_INATIVO',
    PERFIL_NAO_AUTORIZADO: 'PERFIL_NAO_AUTORIZADO',
    PERMISSAO_INSUFICIENTE: 'PERMISSAO_INSUFICIENTE',
    FEATURE_DISABLED: 'FEATURE_DISABLED',
    ROTA_INEXISTENTE: 'ROTA_INEXISTENTE'
  };
  var ROTAS_PORTAL = [
    rota('admin-conteudo-publico', 'Conteúdo Público', 'admin/conteudo-publico', 'gestao-geapa', 72,
      'tela-public-admin', 'view-public-admin', true, [PERFIS.ADMIN, PERFIS.ADMIN_TECNICO, PERFIS.DIRETORIA],
      ['PUBLIC_CONTENT_MANAGE'], 'Gestão editorial de Início, Sobre e História.', 'implementado'),
    rota('inicio', 'Início', 'inicio', 'publico', 10, 'tela-inicio', 'view-inicio', false, PUBLICO, [], 'Home pública do Portal GEAPA.', 'implementado'),
    rota('sobre', 'Sobre o GEAPA', 'sobre', 'publico', 20, 'tela-placeholder', 'view-placeholder', false, PUBLICO, [], 'Apresentação institucional pública do GEAPA.', 'implementado'),
    rota('historia', 'História', 'historia', 'publico', 30, 'tela-placeholder', 'view-placeholder', false, PUBLICO, [], 'Histórico público do grupo e marcos institucionais.', 'implementado'),
    rota('diretoria-publica', 'Diretoria', 'diretoria', 'publico', 40, 'tela-placeholder', 'view-placeholder', false, PUBLICO, [], 'Composição pública da diretoria, sem dados privados.', 'implementado'),
    rota('orientadores', 'Orientadores', 'orientadores', 'publico', 50, 'tela-placeholder', 'view-placeholder', false, PUBLICO, [], 'Orientadores e professores colaboradores publicados pelo CMS editorial.', 'implementado'),
    rota('membros-publicos', 'Membros', 'membros', 'publico', 60, 'tela-placeholder', 'view-placeholder', false, PUBLICO, [], 'Membros, egressos e destaques publicados com autorização editorial.', 'implementado'),
    rota('normas', 'Documentos e normas', 'documentos', 'publico', 70, 'tela-placeholder', 'view-placeholder', false, PUBLICO, [], 'Normas públicas e documentos orientativos do GEAPA.', 'implementado'),
    rota('parceiros', 'Parceiros', 'parceiros', 'publico', 80, 'tela-placeholder', 'view-placeholder', false, PUBLICO, [], 'Instituições e parceiros públicos do GEAPA.', 'implementado'),
    rota('processo-seletivo', 'Processo seletivo', 'processo-seletivo', 'publico', 90, 'tela-placeholder', 'view-placeholder', false, PUBLICO, [], 'Informações públicas sobre ingresso e processo seletivo.', 'placeholder'),
    rota('contato', 'Contato', 'contato', 'publico', 100, 'tela-placeholder', 'view-placeholder', false, PUBLICO, [], 'Canais públicos de contato do GEAPA.', 'placeholder'),
    rota('avaliacao-egresso', 'Avaliação de egresso', 'avaliacao-egresso', 'publico', 105, 'tela-avaliacao-egresso', 'view-avaliacao-egresso', false, PUBLICO, [], 'Formulário voluntário e protegido por token.', 'implementado', false),
    rota('login', 'Entrar', 'login', 'publico', 110, 'tela-acesso', 'view-login', false, PUBLICO, [], 'Entrada com e-mail e senha.', 'implementado', false),
    rota('interesse', 'Tenho interesse', 'interesse', 'publico', 120, 'tela-placeholder', 'view-placeholder', false, PUBLICO, [], 'Formulário futuro para cadastro de interesse sem conceder acesso interno automaticamente.', 'placeholder'),

    rota('meu-perfil', 'Meu perfil', 'app/meu-perfil', 'meu-vinculo', 10, 'tela-meu-perfil', 'view-meu-perfil', true, LOGADOS, ['portal:acessar'], 'Dados cadastrais proprios em modo somente leitura.', 'implementado'),
    rota('minha-situacao', 'Minha situação', 'app/minha-situacao', 'meu-vinculo', 20, 'tela-situacao', 'view-situacao', true, [PERFIS.MEMBRO_INGRESSANTE, PERFIS.MEMBRO, PERFIS.DIRETORIA, PERFIS.SECRETARIA, PERFIS.COMUNICACAO, PERFIS.CONSELHO, PERFIS.ADMIN], ['situacao:ver_propria'], 'Painel operacional resumido do usuario logado.', 'implementado'),
    rota('frequencia', 'Minha frequência', 'app/frequencia', 'meu-vinculo', 40, 'tela-placeholder', 'view-placeholder', true, [PERFIS.MEMBRO, PERFIS.DIRETORIA, PERFIS.SECRETARIA, PERFIS.ADMIN], ['situacao:ver_propria', 'presencas:ler'], 'Frequência própria carregada pela view V2 filtrada no backend.', 'implementado'),
    rota('minhas-apresentacoes', 'Minhas apresentações', 'app/minhas-apresentacoes', 'meu-vinculo', 60, 'tela-placeholder', 'view-placeholder', true, [PERFIS.MEMBRO, PERFIS.DIRETORIA, PERFIS.SECRETARIA, PERFIS.COMUNICACAO, PERFIS.CONSELHO, PERFIS.EGRESSO, PERFIS.ADMIN], ['situacao:ver_propria', 'apresentacoes:ver_propria', 'apresentacoes:ver_ate_saida'], 'Apresentações próprias ou histórico permitido pelo backend.', 'implementado'),
    rota('certificados', 'Meus certificados', 'app/certificados', 'meu-vinculo', 70, 'tela-placeholder', 'view-placeholder', true, LOGADOS, ['certificados:ver_proprios'], 'Certificados próprios emitidos ou liberados.', 'placeholder'),
    rota('inscricoes', 'Minhas inscrições', 'app/inscricoes', 'meu-vinculo', 80, 'tela-placeholder', 'view-placeholder', true, LOGADOS, ['inscricoes:ver_proprias'], 'Inscrições próprias em atividades e processo seletivo.', 'placeholder'),
    rota('participacoes', 'Minhas participações', 'app/participacoes', 'meu-vinculo', 90, 'tela-placeholder', 'view-placeholder', true, LOGADOS, ['participacoes:ver_proprias'], 'Participações próprias em atividades permitidas.', 'placeholder'),
    rota('preferencias', 'Preferências de comunicação', 'app/preferencias', 'meu-vinculo', 100, 'tela-placeholder', 'view-placeholder', true, LOGADOS, ['comunicacoes:preferencias_proprias'], 'Preferências de comunicação e consentimentos.', 'placeholder'),

    rota('atividades', 'Próximas atividades', 'atividades', 'agenda', 10, 'tela-atividades', 'view-atividades', false, PUBLICO, [], 'Atividades futuras ou em andamento, filtradas conforme vínculo e permissões.', 'implementado'),
    rota('historico-atividades', 'Histórico de atividades', 'agenda/historico-atividades', 'agenda', 20, 'tela-atividades', 'view-atividades', true, [PERFIS.MEMBRO, PERFIS.DIRETORIA, PERFIS.SECRETARIA, PERFIS.COMUNICACAO, PERFIS.CONSELHO, PERFIS.ADMIN], [], 'Histórico interno de atividades realizadas, com filtros por ciclo, semestre, tipo e eixo. Participações externas ficam em Meu Vínculo.', 'implementado'),
    rota('calendario', 'Calendário', 'agenda/calendario', 'agenda', 30, 'tela-placeholder', 'view-placeholder', false, PUBLICO, [], 'Visão futura para agenda mensal, eventos, reuniões, simpósios, confraternizações e apresentações.', 'placeholder'),

    rota('admin', 'Painel administrativo', 'admin', 'gestao-geapa', 10, 'tela-placeholder', 'view-placeholder', true, OPERACIONAIS, ['gestao:acessar'], 'Painel operacional conforme perfil e permissões.', 'placeholder'),
    rota('admin-membros', 'Membros', 'admin/membros', 'gestao-geapa', 20, 'tela-admin-membros', 'view-admin-membros', true, [PERFIS.SECRETARIA, PERFIS.DIRETORIA, PERFIS.ADMIN], ['membros:ler'], 'Consulta operacional de membros autorizada pelo backend.', 'implementado'),
    rota('admin-correcoes-cadastrais', 'Correções cadastrais', 'admin/correcoes-cadastrais', 'gestao-geapa', 25, 'tela-admin-correcoes-cadastrais', 'view-admin-correcoes-cadastrais', true, [PERFIS.SECRETARIA, PERFIS.DIRETORIA, PERFIS.ADMIN], ['membros:analisar_correcoes'], 'Análise e aplicação de solicitações cadastrais autorizadas pelo Core.', 'implementado'),
    rota('admin-solicitacoes-vinculo', 'Solicitações de vínculo', 'admin/membros/solicitacoes-vinculo', 'gestao-geapa', 27, 'tela-admin-solicitacoes-vinculo', 'view-admin-solicitacoes-vinculo', true, [PERFIS.SECRETARIA, PERFIS.DIRETORIA, PERFIS.ADMIN], ['membros:analisar_solicitacoes_vinculo'], 'Análise humana de suspensões e desligamentos voluntários.', 'implementado'),
    rota('admin-atividades', 'Atividades', 'admin/atividades', 'gestao-geapa', 30, 'tela-placeholder', 'view-placeholder', true, OPERACIONAIS, ['atividades:gerir'], 'Criação, edição e operação de atividades.', 'implementado'),
    rota('admin-chamadas', 'Chamadas', 'admin/chamadas', 'gestao-geapa', 40, 'tela-placeholder', 'view-placeholder', true, [PERFIS.SECRETARIA, PERFIS.DIRETORIA, PERFIS.ADMIN], ['presencas:gerir'], 'Registro e acompanhamento de chamadas/presenças.', 'placeholder'),
    rota('admin-justificativas', 'Justificativas', 'admin/justificativas', 'gestao-geapa', 50, 'tela-placeholder', 'view-placeholder', true, [PERFIS.SECRETARIA, PERFIS.DIRETORIA, PERFIS.ADMIN], ['justificativas:analisar'], 'Análise operacional e decisão de justificativas conforme permissão.', 'implementado'),
    rota('admin-apresentacoes', 'Apresentações', 'admin/apresentacoes', 'gestao-geapa', 60, 'tela-placeholder', 'view-placeholder', true, [PERFIS.SECRETARIA, PERFIS.DIRETORIA, PERFIS.ADMIN], ['apresentacoes:gerir'], 'Gestão operacional de apresentações e materiais.', 'implementado'),
    rota('admin-comunicacao', 'Comunicação', 'admin/comunicacao', 'gestao-geapa', 70, 'tela-placeholder', 'view-placeholder', true, [PERFIS.COMUNICACAO, PERFIS.DIRETORIA, PERFIS.ADMIN], ['comunicacoes:gerir', 'mensageria:ler'], 'Conteúdo público, divulgação e materiais de comunicação.', 'placeholder'),
    rota('admin-processo-seletivo', 'Processo seletivo', 'admin/processo-seletivo', 'gestao-geapa', 80, 'tela-placeholder', 'view-placeholder', true, [PERFIS.DIRETORIA, PERFIS.SECRETARIA, PERFIS.ADMIN], ['processo_seletivo:gerir'], 'Gestão de processo seletivo e classificações autorizadas.', 'placeholder'),
    rota('painel-diretoria-v2', 'Painel da Diretoria', 'admin/painel-diretoria-v2', 'gestao-geapa', 85, 'tela-placeholder', 'view-placeholder', true, [PERFIS.DIRETORIA, PERFIS.SECRETARIA, PERFIS.ADMIN, PERFIS.ADMIN_TECNICO], ['diretoria:painel_v2', 'diretoria:pendencias', 'sistema:status_v2', 'sistema:admin', 'atividades:gerir', 'membros:ler', 'justificativas:analisar'], 'Painel agregado de pendências, inconsistências e status das rotinas V2.', 'implementado'),
    rota('pendencias-diretoria', 'Pendências da Diretoria', 'admin/pendencias-diretoria', 'gestao-geapa', 90, 'tela-placeholder', 'view-placeholder', true, [PERFIS.DIRETORIA, PERFIS.SECRETARIA, PERFIS.ADMIN], ['diretoria:pendencias', 'membros:ler', 'atividades:gerir', 'justificativas:analisar'], 'Pendências consolidadas em views V2, conforme permissão do backend.', 'implementado'),
    rota('status-v2', 'Status do Sistema V2', 'admin/status-v2', 'gestao-geapa', 100, 'tela-placeholder', 'view-placeholder', true, [PERFIS.DIRETORIA, PERFIS.SECRETARIA, PERFIS.ADMIN], ['sistema:status_v2', 'sistema:admin', 'atividades:gerir', 'membros:ler'], 'Status de atualização das views V2 consumidas pelo Portal.', 'implementado'),
    rota('admin-normas-parametros', 'Normas e parâmetros', 'admin/normas-parametros', 'gestao-geapa', 110, 'tela-placeholder', 'view-placeholder', true, [PERFIS.DIRETORIA, PERFIS.ADMIN], ['parametros:gerir'], 'Normas e parâmetros operacionais sujeitos à autorização.', 'placeholder'),
    rota('admin-configuracoes', 'Configurações', 'admin/configuracoes', 'gestao-geapa', 120, 'tela-placeholder', 'view-placeholder', true, [PERFIS.ADMIN], ['sistema:admin'], 'Configurações técnicas, integrações, usuários, segurança e logs.', 'placeholder')
  ];
  var ALIASES_ROTAS = {
    situacao: 'minha-situacao',
    'minha_situacao': 'minha-situacao',
    app: 'minha-situacao',
    'resumo-vinculo': 'minha-situacao',
    'resumo_do_vinculo': 'minha-situacao',
    'meu-cadastro': 'meu-perfil',
    'app/meu-cadastro': 'meu-perfil',
    cadastro: 'meu-perfil',
    perfil: 'meu-perfil',
    justificativas: 'frequencia',
    'app/justificativas': 'frequencia',
    atividades: 'atividades',
    agenda: 'atividades',
    'atividades-abertas': 'atividades',
    'atividades-internas': 'historico-atividades',
    documentos: 'normas',
    materiais: 'atividades',
    membros: 'membros-publicos',
    gestao: 'admin',
    diretoria: 'diretoria-publica',
    administracao: 'admin-configuracoes',
    logs: 'admin-configuracoes'
  };
  var rotaAtual = 'inicio';
  var inicializado = false;
  var ignorarProximoHash = false;

  ROTAS_PORTAL.push(rota('minha-atuacao','Minha atuação','app/minha-atuacao','meu-vinculo',25,'tela-minha-atuacao','view-minha-atuacao',true,[],[],'Vínculo e funções próprias.','implementado'));
  function publicV2Enabled(page) {
    return global.PortalPublicData && global.PortalPublicData.isEnabled(page);
  }

  if (publicV2Enabled()) {
    ROTAS_PORTAL.push(rota('producao-cientifica', 'Produção Científica', 'producao-cientifica', 'publico', 30,
      'tela-placeholder', 'view-placeholder', false, PUBLICO, [], 'Em preparação.', 'placeholder'));
    ROTAS_PORTAL.push(rota('equipe', 'Equipe', 'equipe', 'publico', 40,
      'tela-placeholder', 'view-placeholder', false, PUBLICO, [], 'Em preparação.', 'placeholder'));
    ROTAS_PORTAL.forEach(function(item) {
      if (item.id === 'parceiros') { item.label = 'Apoio e Parcerias'; item.ordem = 50; }
      if (item.id === 'processo-seletivo') { item.label = 'Processo Seletivo'; item.ordem = 60; }
      if (item.id === 'contato') item.ordem = 80;
    });
  }

  function publicV2MenuVisible(item) {
    if (!publicV2Enabled()) return true;
    if (publicV2Enabled('about') && ['historia', 'normas'].indexOf(item.id) >= 0) return false;
    if (['diretoria-publica', 'orientadores', 'membros-publicos', 'interesse'].indexOf(item.id) >= 0) return false;
    var futureFlags = { equipe: 'team', parceiros: 'partners', 'producao-cientifica': 'science',
      'processo-seletivo': 'selection', contato: 'contact' };
    return !futureFlags[item.id] || publicV2Enabled(futureFlags[item.id]);
  }

  function iniciar() {
    if (inicializado || typeof document === 'undefined') {
      return;
    }

    inicializado = true;
    renderizarMenu();
    registrarEventos();
    sincronizarRotaComHash();
    if (!obterRotaPorHash()) {
      irPara('inicio', { atualizarHash: false, origem: 'inicial' });
    }
    atualizarMenu();
  }

  function rota(id, label, path, grupoMenu, ordem, sectionId, viewClass, requerLogin, perfisPermitidos, permissoesNecessarias, descricao, status, mostrarNoMenu) {
    return {
      id: id,
      label: label,
      path: path,
      viewClass: viewClass,
      sectionId: sectionId,
      grupoMenu: grupoMenu,
      perfisPermitidos: perfisPermitidos || [],
      permissoesNecessarias: permissoesNecessarias || [],
      requerLogin: requerLogin === true,
      mostrarNoMenu: mostrarNoMenu !== false,
      ordem: ordem,
      descricao: descricao || '',
      status: status || 'placeholder'
    };
  }

  function registrarEventos() {
    document.addEventListener('click', function navegarPorAtributo(evento) {
      var alvo = evento.target.closest('[data-route-target]');

      if (!alvo || alvo.disabled) {
        return;
      }

      evento.preventDefault();
      irPara(alvo.getAttribute('data-route-target'));
    });

    global.addEventListener('hashchange', function navegarPorHash() {
      if (ignorarProximoHash) {
        ignorarProximoHash = false;
        return;
      }

      sincronizarRotaComHash();
    });
  }

  function sincronizarRotaComHash() {
    var hashBruto = obterHashBruto_();
    var idRota = obterRotaPorHash();

    if (['app/minhas-funcoes','minhas-funcoes','app/minha-diretoria','minha-diretoria'].includes(hashBruto)) { irPara('minha-atuacao',{atualizarHash:true,substituirHash:true}); return; }
    if (hashBruto === 'app/justificativas' || hashBruto === 'justificativas') {
      irPara('frequencia', {
        atualizarHash: true,
        substituirHash: true,
        motivo: 'justificativas-redirecionadas'
      });
      return;
    }

    if (idRota) {
      irPara(idRota, { atualizarHash: false, origem: 'hash' });
    }
  }

  function restoreCurrentRoute(options) {
    var routeId = obterRotaPorHash();
    if (!routeId) return irPara('inicio', { atualizarHash: false, motivo: 'sem-rota-restauravel' });
    return irPara(routeId, Object.assign({}, options || {}, { atualizarHash: false }));
  }

  function irPara(idRota, opcoes) {
    var rota = resolverRota(idRota);
    var anterior = rotaAtual;
    var detalhes = opcoes || {};
    var sessao = getSessaoAtual();
    var acesso;

    if (!rota) {
      renderAccessDenied({ label: String(idRota || 'Rota') }, MOTIVOS_ACESSO.ROTA_INEXISTENTE);
      aplicarViewAcessoNegado();
      fecharMenu();
      return false;
    }

    acesso = resolveRouteAccess(sessao, rota);

    if (!acesso.ok && acesso.reason === MOTIVOS_ACESSO.PORTAL_INATIVO && rota.id !== 'inicio') {
      rota = resolverRota('inicio');
      acesso = resolveRouteAccess(sessao, rota);
      detalhes = Object.assign({}, detalhes, {
        atualizarHash: true,
        motivo: 'portal-inativo-redirecionado',
        substituirHash: true
      });
    }

    if (!acesso.ok) {
      renderAccessDenied(rota, acesso.reason);
      aplicarViewAcessoNegado();
      atualizarMenu();
      atualizarHashDaRota(rota, detalhes);
      fecharMenu();
      emitirEvento('portal:navigationdenied', {
        rota: rota,
        anterior: anterior,
        motivo: acesso.reason,
        sessao: sessao
      });
      return false;
    }

    aplicarRota(rota);
    rotaAtual = rota.id;
    atualizarMenu();
    atualizarHashDaRota(rota, detalhes);
    fecharMenu();
    emitirEvento('portal:navigationchange', {
      rota: rota,
      anterior: anterior,
      motivo: detalhes.motivo || '',
      sessao: sessao
    });

    return true;
  }

  function aplicarRota(rota) {
    var app = document.getElementById('portal-app');

    if (!app) {
      return;
    }

    limparViews(app);

    if (rota.viewClass) {
      app.classList.add(rota.viewClass);
    }

    mostrarSecao(rota.sectionId);

    if (rota.status === 'placeholder') {
      renderPlaceholderRoute(rota, getSessaoAtual());
    }
  }

  function aplicarViewAcessoNegado() {
    var app = document.getElementById('portal-app');

    if (!app) {
      return;
    }

    limparViews(app);
    app.classList.add('view-access-denied');
    mostrarSecao('tela-acesso-negado');
  }

  function limparViews(app) {
    getRotasInternas().forEach(function limpar(rota) {
      if (rota.viewClass) {
        app.classList.remove(rota.viewClass);
      }

      if (rota.sectionId) {
        var secaoRota = document.getElementById(rota.sectionId);

        if (secaoRota) {
          secaoRota.hidden = true;
        }
      }
    });

    app.classList.remove('view-access-denied');
  }

  function mostrarSecao(sectionId) {
    var secao = sectionId ? document.getElementById(sectionId) : null;

    if (secao) {
      secao.hidden = false;
    }
  }

  function resolveRouteAccess(sessao, rota) {
 if(!rota)return {ok:false,reason:MOTIVOS_ACESSO.ROTA_INEXISTENTE};
 if(['inicio','login'].includes(rota.id))return {ok:true,reason:''};
 if(['atividades','historico-atividades','frequencia','minhas-apresentacoes','admin-apresentacoes','admin-justificativas','admin-atividades'].includes(rota.id))return {ok:!!global.PortalDevContext?.act?.available&&(!global.PortalDevContext.act.materialOnly||rota.id==='minhas-apresentacoes')&&(!['admin-atividades','admin-justificativas','admin-apresentacoes'].includes(rota.id)||global.PortalDevContext.act.board),reason:'FEATURE_DISABLED'};
 if(!['meu-perfil','minha-situacao','minha-atuacao'].includes(rota.id))return {ok:false,reason:'PENDENTE_DE_MIGRACAO_PRIVILEGIADA'};
 return {ok:!!global.PortalDevContext?.authenticated,reason:global.PortalDevContext?.authenticated?'':'NOT_AUTHENTICATED'};
 }

  function hasPermission(sessao, permissao) {
    var mapa = sessao && sessao.permissoesMapa ? sessao.permissoesMapa : {};

    return mapa[permissao] === true;
  }

  function hasAnyProfile(sessao, perfis) {
    var perfisSessao = sessao && Array.isArray(sessao.perfisPortal)
      ? sessao.perfisPortal
      : [PERFIS.VISITANTE];

    if (!perfis || !perfis.length) {
      return true;
    }

    return perfis.some(function verificarPerfil(perfil) {
      return perfisSessao.indexOf(perfil) >= 0;
    });
  }

  function hasRoutePermission(sessao, rota) {
    var necessarias = rota.permissoesNecessarias || [];

    if (!necessarias.length) {
      return true;
    }

    if (!sessao || !Array.isArray(sessao.permissoes) || !sessao.permissoes.length) {
      return !rota.requerLogin ||
        !ehVisitanteTecnico(sessao) ||
        permiteFallbackVisitanteTecnico(rota);
    }

    if (!sessao.permissoes.some(ehPermissaoCanonica)) {
      return true;
    }

    return necessarias.some(function verificarPermissao(permissao) {
      return hasPermission(sessao, permissao);
    });
  }

  function ehVisitanteTecnico(sessao) {
    var perfis = sessao && Array.isArray(sessao.perfisPortal)
      ? sessao.perfisPortal
      : [];

    return sessao &&
      sessao.autenticado === true &&
      perfis.indexOf(PERFIS.VISITANTE) >= 0 &&
      perfis.length === 1;
  }

  function permiteFallbackVisitanteTecnico(rota) {
    return [
      'app',
      'meu-perfil',
      'inscricoes',
      'preferencias'
    ].indexOf(rota.id) >= 0;
  }

  function ehPermissaoCanonica(permissao) {
    return String(permissao || '').indexOf(':') > 0;
  }

  function getAllowedRoutes(sessao) {
    return ROTAS_PORTAL.filter(function filtrarRota(rota) {
      var sessaoAtual = sessao || getSessaoAtual();
      return rota.mostrarNoMenu &&
        publicV2MenuVisible(rota) &&
        routeFeatureEnabled(rota) &&
        resolveRouteAccess(sessaoAtual, rota).ok;
    }).sort(ordenarRotas);
  }

  function routeFeatureEnabled(rota) {
    var config = global.PortalGeapaConfig || {};
    var readOnly = config.READ_ONLY_MODE === true;
    var gestaoPermitidaNoHomolog = ['admin-membros', 'admin-correcoes-cadastrais', 'admin-solicitacoes-vinculo', 'admin-justificativas'];
    if (readOnly && rota.grupoMenu === 'gestao-geapa' && gestaoPermitidaNoHomolog.indexOf(rota.id) < 0) return false;
    if (rota.id === 'admin-atividades' && config.ENABLE_ACTIVITY_MANAGEMENT === false) return false;
    if (rota.id === 'admin-correcoes-cadastrais' && config.ENABLE_PROFILE_UPDATES !== true) return false;
    if (rota.id === 'admin-solicitacoes-vinculo' && config.ENABLE_VINCULO_REQUESTS !== true) return false;
    if (rota.id === 'admin-justificativas' && config.ENABLE_JUSTIFICATIVAS === false) return false;
    return true;
  }

  function getMenuRoutes(sessao) {
    var routes = getAllowedRoutes(sessao), editorial = global.PortalPublicAdmin;
    var route = resolverRota('admin-conteudo-publico');
    // O perfil apenas mantem o atalho para consultar acesso. O guard e o backend
    // continuam exigindo a permissao editorial confirmada para liberar o editor.
    if (editorial && editorial.isEnabled() && editorial.getAccessStatus && sessao.autenticado &&
        sessao.portalAtivo !== false && hasAnyProfile(sessao, route.perfisPermitidos) && routeFeatureEnabled(route) &&
        ['UNCHECKED', 'CHECKING', 'UNAVAILABLE', 'SESSION_EXPIRED'].indexOf(editorial.getAccessStatus()) >= 0 &&
        !routes.some(function(item) { return item.id === route.id; })) routes.push(route);
    return routes.sort(ordenarRotas);
  }

  function renderizarMenu() {
    var nav = document.querySelector('[data-portal-nav]');
    var botaoSair = document.getElementById('sair');
    var sessao = getSessaoAtual();
    var rotasPermitidas = getMenuRoutes(sessao);
    var gruposHtml = GRUPOS_MENU.sort(ordenarGrupos).map(function montarGrupo(grupo) {
      var rotasDoGrupo = rotasPermitidas.filter(function filtrarPorGrupo(rota) {
        return rota.grupoMenu === grupo.id;
      });

      if (!rotasDoGrupo.length) {
        return '';
      }

      return [
        '<details class="nav-group" ',
        grupo.abertoPadrao === false ? '' : 'open',
        '>',
        '<summary class="nav-group-title">',
        '<span>' + escaparHtml(grupo.label) + '</span>',
        '<span class="nav-group-arrow" aria-hidden="true">&#9662;</span>',
        '</summary>',
        '<div class="nav-group-items">',
        rotasDoGrupo.map(montarBotaoMenu).join(''),
        '</div>',
        '</details>'
      ].join('');
    }).join('');

    if (!nav) {
      return;
    }

    nav.innerHTML = gruposHtml || '<p class="nav-empty">Nenhuma área disponível.</p>';
    if (botaoSair) {
      botaoSair.hidden = !sessao.autenticado;
    }
    atualizarAcoesCabecalho(sessao);
    destacarRotaAtual();
  }

  function atualizarAcoesCabecalho(sessao) {
    var acoes = document.querySelector('.header-session');

    if (!acoes) {
      return;
    }

    acoes.hidden = Boolean(sessao && sessao.autenticado);
  }

  function montarBotaoMenu(rota) {
    return [
      '<button class="nav-button" type="button" data-route-target="',
      escaparHtml(rota.id),
      '">',
      '<span>',
      escaparHtml(rota.label),
      '</span>',
      '</button>'
    ].join('');
  }

  function atualizarMenu() {
    renderizarMenu();
    destacarRotaAtual();
  }

  function destacarRotaAtual() {
    Array.prototype.forEach.call(document.querySelectorAll('[data-route-target]'), function destacar(botao) {
      var ativo = normalizarIdRota(botao.getAttribute('data-route-target')) === rotaAtual;

      botao.classList.toggle('active', ativo);
      botao.setAttribute('aria-current', ativo ? 'page' : 'false');
    });
  }

  function renderPlaceholderRoute(rota, sessao) {
    var container = document.getElementById('placeholder-content');
    var conteudo;

    if (!container) {
      return;
    }

    conteudo = [
      '<p class="eyebrow">',
      escaparHtml(obterLabelGrupo(rota.grupoMenu)),
      '</p>',
      '<h2>',
      escaparHtml(rota.label),
      '</h2>',
      '<p class="intro">',
      escaparHtml(rota.descricao),
      '</p>'
    ];

    if (deveMostrarPoliticaVisual()) {
      conteudo = conteudo.concat([
        '<div class="route-policy-card">',
        '<h3>Política visual desta rota</h3>',
        '<dl class="route-policy-list">',
        montarLinhaPolitica('Status', rota.status === 'placeholder' ? 'Tela futura / placeholder' : rota.status),
        montarLinhaPolitica('Perfil atual', sessao.perfilPortalEfetivo || 'VISITANTE'),
        montarLinhaPolitica('Perfis permitidos', (rota.perfisPermitidos || []).join(', ')),
        montarLinhaPolitica('Permissões da rota', (rota.permissoesNecessarias || []).join(', ') || 'Nenhuma'),
        '</dl>',
        '</div>'
      ]);
    }

    container.innerHTML = conteudo.join('');
  }

  function deveMostrarPoliticaVisual() {
    var config = global.PortalGeapaConfig || {};

    return config.SHOW_ROUTE_POLICY === true || config.ROUTE_DEBUG === true;
  }

  function renderAccessDenied(rota, reason) {
    var container = document.getElementById('access-denied-content');
    var mensagem = obterMensagemAcessoNegado(reason, getSessaoAtual());
    var destino = reason === MOTIVOS_ACESSO.NOT_AUTHENTICATED ? 'login' : 'inicio';
    var rotuloBotao = reason === MOTIVOS_ACESSO.NOT_AUTHENTICATED ? 'Voltar para o login' : 'Ir para o início';

    if (!container) {
      return;
    }

    var editorial = global.PortalPublicAdmin;
    if (rota.id === 'admin-conteudo-publico' && editorial && editorial.isEnabled() && editorial.renderAccessGate) {
      editorial.renderAccessGate(container);
      return;
    }

    container.innerHTML = [
      '<p class="eyebrow">Acesso restrito</p>',
      '<h2>',
      escaparHtml(rota.label || 'Rota indisponível'),
      '</h2>',
      '<p class="intro">',
      escaparHtml(mensagem),
      '</p>',
      '<button class="secondary-button" type="button" data-route-target="',
      destino,
      '">',
      rotuloBotao,
      '</button>'
    ].join('');
  }

  function obterMensagemAcessoNegado(reason, sessao) {
    if (reason === MOTIVOS_ACESSO.NOT_AUTHENTICATED) {
      return 'Faça login para acessar esta área.';
    }

    if (reason === MOTIVOS_ACESSO.PORTAL_INATIVO) {
      return sessao && sessao.mensagemBloqueio
        ? sessao.mensagemBloqueio
        : 'Seu e-mail nao esta liberado para acessar o Portal GEAPA ou nao possui vinculo ativo no grupo. Entre com o mesmo e-mail cadastrado junto ao GEAPA.';
    }

    if (reason === MOTIVOS_ACESSO.PERMISSAO_INSUFICIENTE) {
      return 'Seu perfil atual não possui permissão para acessar esta área.';
    }

    if (reason === MOTIVOS_ACESSO.ROTA_INEXISTENTE) {
      return 'Área indisponível no Portal GEAPA.';
    }

    if (reason === MOTIVOS_ACESSO.FEATURE_DISABLED) {
      return 'Esta area esta desativada neste ambiente.';
    }

    return 'Seu perfil atual não possui acesso a esta área.';
  }

  function getSessaoAtual() {return {autenticado:!!global.PortalDevContext?.authenticated,portalAtivo:true,usuarioResolvido:true,perfisPortal:[],permissoesMapa:{}};}

  function resolverRota(idRota) {
    var idNormalizado = normalizarIdRota(idRota);

    return ROTAS_PORTAL.find(function encontrarRota(rota) {
      return rota.id === idNormalizado || rota.path === idNormalizado;
    }) || null;
  }

  function normalizarIdRota(idRota) {
    var id = String(idRota || '').trim();
    if (['app/minhas-funcoes','minhas-funcoes','app/minha-diretoria','minha-diretoria','app/minha-atuacao'].includes(id)) return 'minha-atuacao';

    if (publicV2Enabled('about') && ['historia', 'normas', 'documentos'].indexOf(id) >= 0) return 'sobre';
    if (publicV2Enabled() && ['diretoria', 'diretoria-publica', 'membros', 'membros-publicos'].indexOf(id) >= 0) return 'equipe';

    return ALIASES_ROTAS[id] || id;
  }

  function obterRotaPorHash() {
    var hash = obterHashBruto_();

    if (!hash) {
      return '';
    }

    return normalizarIdRota(hash);
  }

  function obterHashBruto_() {
    return String(global.location.hash || '').replace(/^#\/?/, '').trim().split('?')[0];
  }

  function atualizarHashDaRota(rota, opcoes) {
    if (opcoes && opcoes.atualizarHash === false) {
      return;
    }

    if (!rota.path || rota.id === 'login') {
      return;
    }

    var novoHash = '#/' + rota.path;

    if (global.location.hash !== novoHash) {
      if (opcoes && opcoes.substituirHash === true && global.history && typeof global.history.replaceState === 'function') {
        global.history.replaceState(null, '', global.location.pathname + global.location.search + novoHash);
        return;
      }

      ignorarProximoHash = true;
      global.location.hash = novoHash;
    }
  }

  function getRotasInternas() {
    return ROTAS_PORTAL.concat([{
      id: 'acesso-negado',
      viewClass: 'view-access-denied',
      sectionId: 'tela-acesso-negado'
    }]);
  }

  function fecharMenu() {
    var frame = document.querySelector('.portal-frame');
    var botaoMenu = document.getElementById('menu-toggle');
    var backdrop = document.querySelector('[data-close-sidebar]');

    if (frame) {
      frame.classList.remove('sidebar-open');
    }

    if (botaoMenu) {
      botaoMenu.setAttribute('aria-expanded', 'false');
      botaoMenu.setAttribute('aria-label', 'Abrir menu');
    }

    if (backdrop) {
      backdrop.hidden = true;
    }
  }

  function montarLinhaPolitica(titulo, valor) {
    return [
      '<div>',
      '<dt>',
      escaparHtml(titulo),
      '</dt>',
      '<dd>',
      escaparHtml(valor || 'Não informado'),
      '</dd>',
      '</div>'
    ].join('');
  }

  function obterLabelGrupo(idGrupo) {
    var grupo = GRUPOS_MENU.find(function encontrarGrupo(item) {
      return item.id === idGrupo;
    });

    return grupo ? grupo.label : 'Portal GEAPA';
  }

  function ordenarRotas(a, b) {
    return a.ordem - b.ordem;
  }

  function ordenarGrupos(a, b) {
    return a.ordem - b.ordem;
  }

  function emitirEvento(nome, detalhe) {
    document.dispatchEvent(new CustomEvent(nome, {
      detail: detalhe
    }));
  }

  function escaparHtml(valor) {
    return String(valor)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#039;');
  }

  global.PortalGeapaNavigation = {
    PERFIS: PERFIS,
    GRUPOS_MENU: GRUPOS_MENU.slice(),
    MOTIVOS_ACESSO: MOTIVOS_ACESSO,
    ROTAS_PORTAL: ROTAS_PORTAL.slice(),
    iniciar: iniciar,
    irPara: irPara,
    restoreCurrentRoute: restoreCurrentRoute,
    atualizarMenu: atualizarMenu,
    canAccessRoute: function canAccessRoute(sessao, rota) {
      return resolveRouteAccess(sessao || getSessaoAtual(), rota).ok;
    },
    resolveRouteAccess: resolveRouteAccess,
    hasPermission: hasPermission,
    hasAnyProfile: hasAnyProfile,
    getAllowedRoutes: getAllowedRoutes,
    fecharMenu: fecharMenu,
    getRotaAtual: function getRotaAtual() {
      return rotaAtual;
    },
    getRotas: function getRotas() {
      return ROTAS_PORTAL.slice();
    }
  };

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', iniciar);
  } else {
    iniciar();
  }
})(window);
