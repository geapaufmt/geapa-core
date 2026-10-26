/**
 * Gestao administrativa de Atividades V2.
 *
 * O front-end renderiza contratos sanitizados. Permissao, modelo e campos
 * editaveis sao sempre revalidados no backend antes de qualquer escrita.
 */
(function configurarGestaoAtividades(global) {
  var api = global.PortalGeapaApi;
  var ui = global.PortalGeapaUi;
  var navigation = global.PortalGeapaNavigation;
  var state = {
    active: false,
    loading: false,
    activities: [],
    filterOptions: {},
    details: {},
    filters: {}
  };

  document.addEventListener('portal:act-session-cleared',function(){state.activities=[];state.details={};state.filterOptions={};if(state.active){var el=getContainer();if(el)el.textContent='';}state.active=false;});
  function start() {
    if (!api || !ui || !navigation) return;
    document.addEventListener('portal:navigationchange', function onNavigation(event) {
      var route = event.detail && event.detail.rota;
      state.active = !!(route && route.id === 'admin-atividades');
      if (state.active) loadActivities(false);
    });
    document.addEventListener('click', handleClick);
    document.addEventListener('submit', handleSubmit);
    document.addEventListener('change', handleFilterChange);
    document.addEventListener('input', handleFilterInput);
    document.addEventListener('portal:atividade-criada', function onActivityCreated(event) {
      state.details = {};
      var activity = event.detail && event.detail.atividade || {};
      toast('success', 'Atividade ' + (activity.idAtividade || '') + ' criada como rascunho.');
      if (state.active) loadActivities(true);
    });
    document.addEventListener('portal:apresentacoes-atualizadas', function onPresentationsUpdated() {
      state.details = {};
      if (state.active) loadActivities(true);
    });
    if (typeof navigation.getRotaAtual === 'function' && navigation.getRotaAtual() === 'admin-atividades') {
      state.active = true;
      loadActivities(false);
    }
  }

  function loadActivities(background) {
    var container = getContainer();
    if (!container || state.loading) return;
    state.loading = true;
    if (!background) renderShell('<p class="empty-state">Carregando atividades administrativas...</p>');
    api.apiGet('/admin/atividades', { filtros: JSON.stringify({}) })
      .then(function onResponse(response) {
        if (!response.ok) throw new Error(response.message || 'Nao foi possivel carregar as atividades.');
        var data = response.data || {};
        state.activities = Array.isArray(data.atividades) ? data.atividades : [];
        state.filterOptions = data.opcoesFiltros || {};
        renderContent();
      })
      .catch(function onError(error) {
        renderShell('<p class="empty-state readonly-error">' + escape(adminErrorMessage(error, 'Erro ao carregar atividades.')) + '</p>');
      })
      .then(function done() { state.loading = false; });
  }

  function renderShell(content) {
    var container = getContainer();
    if (!container || !state.active) return;
    container.innerHTML = [
      '<p class="eyebrow">Gestao do GEAPA</p>',
      '<div class="admin-activities-heading">',
      '<div><h2>Atividades</h2><p class="intro">Revise rascunhos e gerencie atividades criadas por modelo.</p></div>',
      '<button class="secondary-button compact-button" type="button" data-create-activity'+(global.PortalGeapaAuth?.canCreateActivity?.()?'':' disabled title="Criação indisponível para este acesso"')+'>Criar atividade</button><button class="secondary-button compact-button" type="button" data-admin-config-open'+(global.PortalGeapaAuth?.canConfigureActivity?.()?'':' disabled title="Configuração indisponível para este acesso"')+'>Modelos e regras</button>',
      '</div>',
      '<div class="portal-feedback-slot" data-admin-activities-feedback hidden></div>',
      content
    ].join('');
  }

  function renderContent() {
    var filtered = filterActivities(state.activities, state.filters);
    renderShell([
      renderFilters(),
      '<div class="admin-activities-summary"><strong>' + filtered.length + '</strong><span>atividade(s) encontrada(s)</span></div>',
      renderTable(filtered)
    ].join(''));
  }

  function renderFilters() {
    var options = state.filterOptions || {};
    return [
      '<form class="admin-activities-filters" data-admin-activity-filters>',
      filterText('texto', 'Buscar', 'ID, titulo, modelo ou pessoa'),
      filterSelect('statusOperacional', 'Status operacional', options.statusOperacionais),
      filterSelect('statusPublicacao', 'Publicacao', options.statusPublicacao),
      filterSelect('visibilidade', 'Visibilidade', options.visibilidades),
      filterSelect('ciclo', 'Ciclo', options.ciclos),
      filterSelect('modelo', 'Modelo', options.modelos),
      filterSelect('tipo', 'Tipo', options.tipos),
      filterSelect('subtipo', 'Subtipo', options.subtipos),
      filterSelect('pendencia', 'Pendencia', ['TITULO_EIXO', 'MATERIAL', 'FOTO_REUNIAO']),
      '<button class="secondary-button compact-button" type="button" data-admin-clear-filters>Limpar filtros</button>',
      '</form>'
    ].join('');
  }

  function filterText(name, label, placeholder) {
    return '<label><span>' + escape(label) + '</span><input type="search" name="' + name + '" value="' + escape(state.filters[name] || '') + '" placeholder="' + escape(placeholder) + '"></label>';
  }

  function filterSelect(name, label, values) {
    var selected = String(state.filters[name] || '');
    var list = Array.isArray(values) ? values : [];
    return [
      '<label><span>' + escape(label) + '</span><select name="' + name + '">',
      '<option value="">Todos</option>',
      list.map(function option(value) {
        return '<option value="' + escape(value) + '"' + (String(value) === selected ? ' selected' : '') + '>' + escape(formatLabel(value)) + '</option>';
      }).join(''),
      '</select></label>'
    ].join('');
  }

  function renderTable(items) {
    if (!items.length) return '<p class="empty-state">Nenhuma atividade corresponde aos filtros.</p>';
    return [
      '<div class="readonly-table-wrap admin-activities-table-wrap">',
      '<table class="readonly-table admin-activities-table">',
      '<thead><tr><th>Atividade</th><th>Data</th><th>Modelo</th><th>Status</th><th>Pendencias</th><th><span class="sr-only">Acoes</span></th></tr></thead>',
      '<tbody>',
      items.map(function row(item) {
        return [
          '<tr>',
          '<td><strong>' + escape(item.tituloPublico || 'Atividade') + '</strong><small>' + escape(item.idAtividade) + '</small>',
          item.nomePessoaPrincipalPublico ? '<small>' + escape(item.nomePessoaPrincipalPublico) + '</small>' : '',
          '</td>',
          '<td>' + escape(ui.formatarData(item.dataAtividade)) + '<small>' + escape([item.horarioInicio, item.horarioFim].filter(Boolean).join(' - ')) + '</small></td>',
          '<td>' + escape(item.nomeModeloPortal || formatLabel(item.subtipoAtividade)) + '<small>' + escape(item.idConfigModelo || '') + '</small></td>',
          '<td><span class="admin-status-chip">' + escape(formatLabel(item.statusOperacional)) + (item.estadoConferencia?' · '+escape(item.estadoConferencia):'') + '</span><small>' + escape(formatLabel(item.statusPublicacaoPortal)) + ' / ' + escape(formatLabel(item.visibilidadePortal)) + '</small></td>',
          '<td>' + renderPending(getActivePendencies(item)) + '</td>',
          '<td><div class="admin-row-actions">',
          '<button class="secondary-button compact-button" type="button" data-admin-open="' + escape(item.idAtividade) + '">Abrir</button>',
          renderQuickActions(item),
          '</div></td>',
          '</tr>'
        ].join('');
      }).join(''),
      '</tbody></table></div>'
    ].join('');
  }

  function renderPending(pending) {
    var items = Array.isArray(pending) ? pending : [];
    return items.length ? items.map(function(item) {
      var kind = pendingKind(item);
      var status = item && typeof item === 'object' && typeof item.status === 'string' ? item.status : '';
      return '<span class="admin-pending-chip">' + escape(formatLabel(kind || 'PENDENCIA')) + (status ? ': ' + escape(formatLabel(status)) : '') + '</span>';
    }).join('') : '<span class="section-note">Sem pendencias</span>';
  }

  function renderQuickActions(item) {
    if (item.somenteConsulta || !item.podeEditar) return '';
    if (item.somenteConsulta) return '';
    var operational = String(item.statusOperacional || '').toUpperCase();
    var publication = String(item.statusPublicacaoPortal || '').toUpperCase();
    if (operational === 'CANCELADA' || operational === 'ARQUIVADA') {
      return '<button class="secondary-button compact-button" type="button" data-admin-state-action="reabrir" data-id-atividade="' + escape(item.idAtividade) + '">Reabrir</button>';
    }
    if (publication === 'PUBLICADA') {
      return '<button class="secondary-button compact-button" type="button" data-admin-state-action="ocultar" data-id-atividade="' + escape(item.idAtividade) + '">Ocultar</button>';
    }
    return '<button class="secondary-button compact-button" type="button" data-admin-state-action="publicar" data-id-atividade="' + escape(item.idAtividade) + '">Publicar</button>';
  }

  function handleClick(event) {
    var configTarget=event.target.closest('[data-admin-config-open],[data-admin-config-kind],[data-admin-config-edit]');
    if(configTarget){event.preventDefault();if(configTarget.hasAttribute('data-admin-config-edit'))editActivityConfig(configTarget.getAttribute('data-admin-config-edit'));else openActivityConfig(configTarget.getAttribute('data-admin-config-kind')||'models');return;}
    if (!state.active) return;
    var openButton = event.target.closest('[data-admin-open]');
    var actionButton = event.target.closest('[data-admin-state-action]');
    var clearButton = event.target.closest('[data-admin-clear-filters]');
    var closeButton = event.target.closest('[data-admin-modal-close]');
    if (openButton) return openDetail(openButton.getAttribute('data-admin-open'));
    if (actionButton) return executeStateAction(actionButton.getAttribute('data-admin-state-action'), actionButton.getAttribute('data-id-atividade'), actionButton);
    if (clearButton) { state.filters = {}; renderContent(); return; }
    if (closeButton) closeModal();
  }

  function handleFilterChange(event) {
    if (!state.active || !event.target.closest('[data-admin-activity-filters]')) return;
    state.filters[event.target.name] = event.target.value;
    renderContent();
  }

  function handleFilterInput(event) {
    if (!state.active || event.target.name !== 'texto' || !event.target.closest('[data-admin-activity-filters]')) return;
    state.filters.texto = event.target.value;
    renderContent();
    var input = document.querySelector('[data-admin-activity-filters] [name="texto"]');
    if (input) { input.focus(); input.setSelectionRange(input.value.length, input.value.length); }
  }

  function handleSubmit(event) {
    var review=event.target.closest('[data-admin-review-form]');if(review){event.preventDefault();saveReview(review,event.submitter&&event.submitter.value);return;}
    var configForm=event.target.closest('[data-admin-config-form]');if(configForm){event.preventDefault();saveActivityConfig(configForm);return;}
    var realization=event.target.closest('[data-admin-realization-form]');
    if(realization){event.preventDefault();saveRealization(realization);return;}
    var form = event.target.closest('[data-admin-activity-edit-form]');
    if (!form) return;
    event.preventDefault();
    saveEdit(form);
  }

  function openDetail(idAtividade, force) {
    if (!idAtividade) return;
    if (!force && state.details[idAtividade]) return renderDetail(state.details[idAtividade]);
    openModal('Gestao da atividade', '<p class="empty-state">Carregando detalhe administrativo...</p>');
    api.apiGet('/admin/atividades/detalhe', { idAtividade: idAtividade })
      .then(function onResponse(response) {
        if (!response.ok) throw new Error(response.message || 'Nao foi possivel carregar o detalhe.');
        state.details[idAtividade] = response.data || {};
        renderDetail(state.details[idAtividade]);
      })
      .catch(showModalError);
  }

  function renderDetail(detail) {
    var activity = detail.atividade || {};
    var editable = Array.isArray(detail.camposEditaveis) ? detail.camposEditaveis : [];
    var content = [
      '<p class="eyebrow">' + escape(activity.idAtividade || 'Atividade') + '</p>',
      '<div class="admin-detail-status"><strong>' + escape(activity.tituloPublico || 'Atividade') + '</strong><span>' + escape(formatLabel(activity.statusOperacional)) + ' / ' + escape(formatLabel(activity.statusPublicacaoPortal)) + '</span></div>',
      '<form class="admin-activity-edit-form" data-admin-activity-edit-form>',
      '<input type="hidden" name="idAtividade" value="' + escape(activity.idAtividade || '') + '"><input type="hidden" name="revision" value="' + detail.canonicalActivityRevision + '">',
      renderEditableFields(activity, editable),
      '<div class="admin-form-actions"><button class="secondary-button" type="button" data-admin-modal-close>Fechar</button>' + (editable.length?'<button type="submit">Salvar edicao</button>':'') + '</div>',
      '</form>',
      renderReview(detail),
      renderProtectedFields(activity, detail.modeloAplicado),
      renderRelations(detail),
      renderRealization(detail),
      '<div class="admin-detail-actions">',
      detail.podePublicar ? '<button type="button" data-admin-state-action="publicar" data-id-atividade="' + escape(activity.idAtividade) + '">Publicar</button>' : '',
      !activity.somenteConsulta ? '<button class="secondary-button" type="button" data-admin-state-action="ocultar" data-id-atividade="' + escape(activity.idAtividade) + '">Ocultar</button>' : '',
      !activity.somenteConsulta && (!detail.historicoRealizacao || !detail.historicoRealizacao.length) ? '<button class="secondary-button admin-danger-button" type="button" data-admin-state-action="cancelar" data-id-atividade="' + escape(activity.idAtividade) + '">Cancelar atividade</button>' : '',
      detail.podeReabrir ? '<button class="secondary-button" type="button" data-admin-state-action="reabrir" data-id-atividade="' + escape(activity.idAtividade) + '">Reabrir como rascunho</button>' : '',
      '</div>'
    ].join('');
    openModal('Gestao da atividade', content);
  }

  function renderReview(detail) {
    var v=detail.atividade||{};if(!v.estadoConferencia)return '';
    var head='<section class="admin-detail-relations"><h3>Conferência administrativa</h3><p><strong>'+escape(v.estadoConferencia)+'</strong></p>';
    if(v.estadoConferencia==='CONFERIDA')return head+'<p>Conferida por '+escape(v.conferenciaResponsavel||'')+' em '+escape(v.conferenciaEm||'')+'. Novas edições comuns estão bloqueadas.</p></section>';
    if(!v.podeConferir)return head+'<p>Conferência pendente. Operação indisponível para este acesso.</p></section>';
    return head+'<p>A conferência é independente da realização, das apresentações e da frequência. As correções abaixo preservam os fatos históricos.</p><form data-admin-review-form><input type="hidden" name="idAtividade" value="'+escape(v.idAtividade)+'"><input type="hidden" name="revision" value="'+detail.canonicalActivityRevision+'">'+renderEditableFields(v,v.camposConferencia||[])+'<label>Justificativa da conferência ou correção<textarea name="motivo" maxlength="2000" required></textarea></label><label><input type="checkbox" name="confirmado"> Conferi os dados e eventuais correções preenchidas e estou ciente de que marcar como conferida bloqueia novas edições comuns, sem reabertura nesta etapa.</label><p data-review-feedback role="status"></p><div class="admin-form-actions"><button type="submit" value="EDIT">Salvar correções da conferência</button><button type="submit" value="CONFIRM">Marcar como conferida</button></div><p>Salvar é opcional. Marcar como conferida registra eventuais correções e confirma a atividade em uma única operação.</p></form></section>';
  }
  function saveReview(form,action) {
    var data=new FormData(form),feedback=form.querySelector('[data-review-feedback]'),id=String(data.get('idAtividade')),fields={};
    if(!['EDIT','CONFIRM'].includes(action))return;
    if(action==='CONFIRM'&&!data.has('confirmado')){feedback.textContent='Confirme explicitamente a conferência antes de continuar.';return;}
    Array.from(form.querySelectorAll('input,textarea')).forEach(function(el){if(el.name&&!['idAtividade','revision','motivo','confirmado'].includes(el.name)&&String(el.value).trim()!==String(el.defaultValue).trim())fields[el.name]=String(el.value);});
    if(action==='EDIT'&&!Object.keys(fields).length){feedback.textContent='Não há correções para salvar. Você pode marcar a atividade como conferida diretamente.';return;}
    var payload={idAtividade:id,revision:Number(data.get('revision')),action:action,campos:fields,motivo:String(data.get('motivo')),confirmado:action==='CONFIRM'};
    setFormBusy(form,true);feedback.textContent='Registrando conferência...';
    api.apiPost('/admin/atividades/conferencia',{payload:JSON.stringify(payload)}).then(function(response){if(!response.ok)throw response;delete state.details[id];notificarAtividadesAtualizadas(id,'CONFERENCIA');toast('success',action==='CONFIRM'?'Atividade conferida.':'Correções registradas.');closeModal();loadActivities(true);openDetail(id,true);}).catch(function(error){feedback.textContent=adminErrorMessage(error,'Não foi possível registrar a conferência. Atualize o detalhe antes de uma nova tentativa.');}).finally(function(){setFormBusy(form,false);});
  }
  function renderEditableFields(activity, editable) {
    function has(name) { return editable.indexOf(name) >= 0; }
    var fields = [];
    if (has('tituloPublico')) fields.push(textField('tituloPublico', 'Titulo publico', activity.tituloPublico));
    if (has('descricaoPublica')) fields.push(textareaField('descricaoPublica', 'Descricao publica', activity.descricaoPublica));
    if (has('descricaoInterna')) fields.push(textareaField('descricaoInterna', 'Descricao interna', activity.descricaoInterna));
    if (has('dataAtividade')) fields.push(textField('dataAtividade', 'Data', activity.dataAtividade, 'date'));
    if (has('horarioInicio')) fields.push(textField('horarioInicio', 'Horario inicial', normalizeTimeInput(activity.horarioInicio), 'time'));
    if (has('horarioFim')) fields.push(textField('horarioFim', 'Horario final', normalizeTimeInput(activity.horarioFim), 'time'));
    if (has('local')) fields.push(textField('local', 'Local', activity.local));
    if (has('formato')) fields.push(selectField('formato', 'Formato', ['PRESENCIAL', 'REMOTO', 'HIBRIDO'], activity.formato));
    if (has('responsavelEmail')) fields.push(textField('responsavelEmail', 'E-mail do responsável', activity.responsavelEmail));
    if (has('responsavelInterno')) fields.push(textField('responsavelInterno', 'Responsavel interno', activity.responsavelInterno));
    if (has('publicoAlvo')) fields.push(textField('publicoAlvo', 'Publico-alvo', activity.publicoAlvo));
    if (has('observacoes')) fields.push(textareaField('observacoes', 'Observacoes internas', activity.observacoes));
    return '<div class="admin-activity-form-grid">' + fields.join('') + '</div>';
  }

  function renderRealization(detail) {
    if(!detail.canonicalActivityRevision)return '';
    var activity=detail.atividade||{},history=detail.historicoRealizacao||[];
    return '<section class="admin-related-data"><h3>Realização</h3>' +
      (history.length?'<ol>'+history.map(function(v){return '<li>'+escape(ui.formatarData(v.dataOcorrencia))+' — '+(v.ocorreu?'Ocorreu':'Não ocorreu')+'<p>'+escape(v.motivo)+'</p></li>';}).join('')+'</ol>':'<p>A realização ainda não foi reconhecida.</p>') +
      (detail.podeRegistrarRealizacao?'<form data-admin-realization-form><input type="hidden" name="idAtividade" value="'+escape(activity.idAtividade)+'"><input type="hidden" name="revision" value="'+detail.canonicalActivityRevision+'"><div class="admin-activity-form-grid"><label>Resultado<select name="ocorreu" required><option value="">Selecionar</option><option value="SIM">Ocorreu</option><option value="NAO">Não ocorreu</option></select></label><label>Data da ocorrência<input name="dataOcorrencia" type="date" required value="'+escape(activity.dataAtividade||'')+'"></label><label>Fundamento do reconhecimento ou revisão<textarea name="motivo" maxlength="2000" required></textarea></label></div><p data-realization-feedback role="status"></p><button type="submit">'+(history.length?'Registrar revisão da realização':'Reconhecer realização')+'</button></form>':'') + '</section>';
  }
  function saveRealization(form) {
    var data=new FormData(form),payload={idAtividade:String(data.get('idAtividade')),revision:Number(data.get('revision')),ocorreu:data.get('ocorreu')==='SIM',dataOcorrencia:String(data.get('dataOcorrencia')),motivo:String(data.get('motivo'))};
    if(!['SIM','NAO'].includes(data.get('ocorreu'))||!payload.motivo.trim())return;
    setFormBusy(form,true);var feedback=form.querySelector('[data-realization-feedback]');feedback.textContent='Registrando realização...';
    api.apiPost('/admin/atividades/realizacao',{payload:JSON.stringify(payload)}).then(function(answer){if(!answer.ok)throw Error(answer.message||'Não foi possível registrar a realização.');delete state.details[payload.idAtividade];notificarAtividadesAtualizadas(payload.idAtividade,'REALIZACAO');toast('success','Realização registrada.');openDetail(payload.idAtividade,true);loadActivities(true);}).catch(function(error){feedback.textContent=error.message||'Não foi possível registrar a realização. Atualize os dados antes de tentar novamente.';}).finally(function(){setFormBusy(form,false);});
  }
  var activityConfig={kind:'models',items:[],related:[],proofDestinations:[],generation:0};
  var configKinds={models:'Modelos de atividade',windows:'Prazos de chamada'};
  var configFields={models:[['name','Nome do modelo','text'],['type','Tipo','text'],['subtype','Subtipo','text'],['durationMinutes','Carga horária em minutos','number'],['visibility','Visibilidade permitida','visibility'],['countsAttendance','Contabiliza presença?','boolean'],['countsAbsence','Contabiliza falta?','boolean'],['allowsJustification','Permite justificativa?','boolean'],['status','Status','status'],['attendancePolicyId','Prazos de chamada','related'],...[["group","Grupo do modelo","text"],["description","Descrição do modelo","text"],["principalRole","Papel da pessoa principal","text"],["requiresPrincipal","Exige pessoa principal?","boolean"],["allowedPrincipalTypes","Tipos de pessoa: MEMBRO, PROFESSOR, CONVIDADO, EXTERNO, EGRESSO","text"],["requiresPrincipalEmail","Exige e-mail da pessoa principal?","boolean"],["requiresPrincipalInstitution","Exige instituição da pessoa principal?","boolean"],["requiresAxis","Exige eixo temático?","boolean"],["allowsSecondaryAxis","Permite eixo secundário?","boolean"],["memberPresentation","Apresentação de membro com título e material posteriores?","boolean"],["requiresMaterial","Exige material?","boolean"],["materialExtensions","Extensões do material, separadas por vírgula","text"],["maxMaterialMb","Limite do material em MB","number"],["materialInstructions","Orientação sobre o material","text"],["showCalendar","Exibir no calendário?","boolean"],["showUpcoming","Exibir em próximas atividades?","boolean"]]],periods:[['label','Nome do período','text'],['civilStart','Primeiro dia','date'],['civilEnd','Último dia','date'],['unit','Unidade de frequência','unit'],['justifiedMode','Falta justificada','mode'],['excusedMode','Falta abonada','mode']],windows:[['cycleId','Ciclo de atividades','related'],['opensBeforeMinutes','Abrir chamada: minutos antes do início','number'],['closesAfterMinutes','Fechar chamada: minutos depois do término','number'],['previousJustificationDays','Justificativa prévia: dias antes','number'],['justificationAfterDays','Justificativa: dias depois','number'],['acceptLateJustification','Aceitar envio de justificativa fora do prazo para análise?','boolean'],['proofUploadPolicyId','Destino privado dos comprovantes','proofDestinations']]};
  var configOptions={visibility:[['BOARD','Diretoria'],['MEMBERS','Membros'],['PUBLIC','Pública']],status:[['ACTIVE','Ativo'],['INACTIVE','Inativo']],unit:[['ACTIVITY','Por atividade'],['MINUTE','Por minuto']],mode:[['EXCLUDE','Excluir do cálculo'],['COUNT_PRESENT','Contar como presença'],['COUNT_ABSENT','Contar como falta']],boolean:[['true','Sim'],['false','Não']]};
  document.addEventListener('portal:act-session-cleared',function(){activityConfig.generation++;activityConfig.items=[];activityConfig.related=[];activityConfig.proofDestinations=[];closeModal();});
  function configLabel(v){return v.name||v.label||'Abre '+v.opensBeforeMinutes+' min antes / fecha '+v.closesAfterMinutes+' min depois';}
  async function configRows(kind){var rows=[],cursor;do{var answer=await api.apiGet('/admin/atividades/configuracoes',Object.assign({kind:kind},cursor?{cursor:cursor}:{}));if(!answer.ok)throw Error(answer.message||'Não foi possível consultar a configuração.');rows=rows.concat(answer.data.items);cursor=answer.data.nextCursor;if(rows.length>=200&&cursor)throw Error('Há muitas configurações para esta consulta.');}while(cursor);return rows;}
  async function openActivityConfig(kind){
    if(!configKinds[kind])return;var gen=++activityConfig.generation;activityConfig.kind=kind;openModal('Modelos e regras','<p>Carregando configurações...</p>');
    try{var rows=await configRows(kind),related=kind==='models'?await configRows('windows'):kind==='windows'?await configRows('cycles'):[];var proofDestinations=kind==='windows'?await configRows('proofDestinations'):[];if(gen!==activityConfig.generation)return;activityConfig.proofDestinations=proofDestinations;activityConfig.items=rows;activityConfig.related=related;openModal('Modelos e regras','<div class="admin-form-actions">'+Object.keys(configKinds).map(function(k){return '<button type="button" class="secondary-button" data-admin-config-kind="'+k+'">'+configKinds[k]+'</button>';}).join('')+'</div><h3>'+configKinds[kind]+'</h3><p>As alterações valem para novos registros. Atividades e chamadas existentes preservam sua base.</p><button type="button" data-admin-config-edit="">Adicionar</button><ul>'+rows.map(function(v){return '<li>'+escape(configLabel(v))+' <button type="button" class="secondary-button" data-admin-config-edit="'+escape(v.id)+'">'+(v.inUse?'Consultar':'Editar')+'</button></li>';}).join('')+'</ul>');}catch(e){if(gen===activityConfig.generation)showModalError(e);}
  }
  function editActivityConfig(id){var kind=activityConfig.kind,row=id?activityConfig.items.find(function(v){return v.id===id}):{id:'',revision:0};if(!row)return;var fields=configFields[kind].map(function(field){var key=field[0],label=field[1],type=field[2],value=row[key]==null?'':String(row[key]),choices=type==='proofDestinations'?activityConfig.proofDestinations.map(function(v){return [v.id,v.label||v.id]}):type==='related'?activityConfig.related.map(function(v){return [v.id,configLabel(v)]}):configOptions[type],required=['durationMinutes','attendancePolicyId','proofUploadPolicyId','group','description','principalRole','allowedPrincipalTypes','materialExtensions','materialInstructions'].includes(key)?'':' required',control;
      if(choices)control='<select name="'+key+'"'+required+(row.inUse?' disabled':'')+'><option value="">Selecionar</option>'+choices.map(function(v){return '<option value="'+escape(v[0])+'"'+(value===String(v[0])?' selected':'')+'>'+escape(v[1])+'</option>';}).join('')+'</select>';else control='<input name="'+key+'" type="'+type+'" value="'+escape(value)+'"'+required+(type==='number'?' min="0" step="1"':'')+(row.inUse?' disabled':'')+'>';return '<label>'+label+control+'</label>';}).join('');
    openModal(configKinds[kind],'<form data-admin-config-form><input type="hidden" name="kind" value="'+kind+'"><input type="hidden" name="id" value="'+escape(row.id)+'"><input type="hidden" name="revision" value="'+row.revision+'"><div class="admin-activity-form-grid">'+fields+'</div>'+(row.inUse?'<p>Este período já é utilizado. Seus critérios foram preservados; crie outro período para uma nova base.</p>':'<label>Motivo da configuração<textarea name="reason" required maxlength="2000"></textarea></label>')+'<p data-config-feedback role="status"></p><div class="admin-form-actions"><button type="button" class="secondary-button" data-admin-config-kind="'+kind+'">Voltar</button>'+(row.inUse?'':'<button type="submit">Salvar configuração</button>')+'</div></form>');
  }
  function saveActivityConfig(form){var data=new FormData(form),kind=String(data.get('kind')),payload={kind:kind,id:String(data.get('id')),revision:Number(data.get('revision')),reason:String(data.get('reason'))},gen=activityConfig.generation;configFields[kind].forEach(function(f){var v=String(data.get(f[0])||'');payload[f[0]]=f[0]==='durationMinutes'&&v===''?null:f[2]==='number'?Number(v):f[2]==='boolean'?v==='true':v;});setFormBusy(form,true);var feedback=form.querySelector('[data-config-feedback]');feedback.textContent='Salvando configuração...';api.apiPost('/admin/atividades/configuracoes/salvar',{payload:JSON.stringify(payload)}).then(function(answer){if(gen!==activityConfig.generation)return;if(!answer.ok)throw Error(answer.message||'Não foi possível salvar.');document.dispatchEvent(new CustomEvent('portal:atividade-configurada'));toast('success','Configuração salva.');openActivityConfig(kind);}).catch(function(e){if(gen===activityConfig.generation)feedback.textContent=e.message;}).finally(function(){setFormBusy(form,false)});}
  function renderProtectedFields(activity, model) {
    var values = [
      ['Modelo', activity.nomeModeloPortal || (model && model.nomeModeloPortal)],
      ['ID do modelo', activity.idConfigModelo],
      ['Tipo', activity.tipoAtividade], ['Subtipo', activity.subtipoAtividade],
      ['Acesso', activity.classificacaoAcesso], ['Conta presenca', activity.contaPresenca],
      ['Conta falta', activity.contaFalta], ['Gera certificado', activity.geraCertificado],
      ['Carga horaria', activity.cargaHoraria]
    ];
    return '<section class="admin-protected-fields"><h3>Regras herdadas do modelo</h3><dl>' + values.map(function(item) {
      return '<div><dt>' + escape(item[0]) + '</dt><dd>' + escape(formatLabel(item[1]) || '-') + '</dd></div>';
    }).join('') + '</dl><p class="section-note">Esses campos sao protegidos e exigem fluxo proprio para alteracao.</p></section>';
  }

  function renderRelations(detail) {
    var activity = detail.atividade || {};
    var presentations = Array.isArray(detail.apresentacoes) ? detail.apresentacoes : [];
    var involved = Array.isArray(detail.envolvidos) ? detail.envolvidos : [];
    var pending = getActivePendencies({
      statusOperacional: activity.statusOperacional,
      pendencias: Array.isArray(detail.pendencias) ? detail.pendencias : []
    });
    return [
      '<section class="admin-related-data"><h3>Vinculos e pendencias</h3>',
      '<p><strong>Apresentacoes:</strong> ' + presentations.length + ' &nbsp; <strong>Envolvidos:</strong> ' + involved.length + '</p>',
      presentations.map(function(presentation) {
        return '<article class="admin-related-file"><strong>' + escape(presentation.idApresentacao || 'Apresentacao') + '</strong>' +
          (presentation.nomeApresentador ? '<p><strong>Apresentador:</strong> ' + escape(presentation.nomeApresentador) + '</p>' : '') +
          (presentation.tituloApresentacao || presentation.tema ? '<p><strong>Tema:</strong> ' + escape(presentation.tituloApresentacao || presentation.tema) + '</p>' : '') +
          '<div><span>Slide: ' + escape(formatLabel(presentation.statusMaterial || 'PENDENTE')) + '</span>' +
          (presentation.canonicalMaterialFile ? '<a class="secondary-button compact-button" href="#" data-portal-v2-action="abrir-material-canonico" data-id-apresentacao="' + escape(presentation.idApresentacao) + '" target="_blank" rel="noopener noreferrer">Abrir slide</a>' : presentation.linkMaterial ? '<a class="secondary-button compact-button" href="' + escape(presentation.linkMaterial) + '" target="_blank" rel="noopener noreferrer">Abrir slide</a>' : '') + '</div>' +
          '<div><span>Foto da reuniao: ' + escape(formatLabel(presentation.statusFotoReuniao || 'PENDENTE')) + '</span>' +
          (presentation.canonicalPhotoFile ? '<a class="secondary-button compact-button" href="#" data-portal-v2-action="abrir-foto-canonica" data-id-apresentacao="' + escape(presentation.idApresentacao) + '" target="_blank" rel="noopener noreferrer">Abrir foto</a>' : presentation.linkFotoReuniao ? '<a class="secondary-button compact-button" href="' + escape(presentation.linkFotoReuniao) + '" target="_blank" rel="noopener noreferrer">Abrir foto</a>' : '') + '</div>' +
          '</article>';
      }).join(''),
      pending.length ? '<div>' + pending.map(function(item) { return '<span class="admin-pending-chip">' + escape(formatLabel(item.tipo)) + ': ' + escape(formatLabel(item.status)) + '</span>'; }).join('') + '</div>' : '<p class="section-note">Sem pendencias vinculadas.</p>',
      '</section>'
    ].join('');
  }

  function saveEdit(form) {
    var data = new FormData(form);
    var id = String(data.get('idAtividade') || '');
    var fields = {};
    data.forEach(function(value, key) { if (key !== 'idAtividade' && key !== 'revision') fields[key] = value; });
    setFormBusy(form, true);
    api.apiPost('/admin/atividades/salvar', { payload: JSON.stringify({ idAtividade: id, revision: Number(data.get('revision')), campos: fields }) })
      .then(function onResponse(response) {
        if (!response.ok) throw response;
        delete state.details[id];
        notificarAtividadesAtualizadas(id, 'SALVAR');
        toast('success', response.message || 'Atividade atualizada.');
        closeModal();
        loadActivities(true);
      })
      .catch(function onError(error) {
        var message = error && error.message ? error.message : 'Nao foi possivel salvar a atividade.';
        var fieldErrors = error && error.data && error.data.fieldErrors ? error.data.fieldErrors : {};
        if (Object.keys(fieldErrors).length && ui.aplicarErrosCampos) {
          ui.aplicarErrosCampos(form, fieldErrors);
        }
        toast('error', message);
      })
      .finally(function done() { setFormBusy(form, false); });
  }

  function executeStateAction(action, idAtividade, button) {
    if (!idAtividade) return;
    var messages = {
      publicar: 'Publicar esta atividade no Portal?',
      ocultar: 'Ocultar esta atividade do Portal?',
      cancelar: 'Cancelar esta atividade sem apagar seus dados?',
      reabrir: 'Reabrir esta atividade como rascunho?'
    };
    var shown = state.details[idAtividade] || state.activities.find(function(item) { return item.idAtividade === idAtividade; });
    var revision = shown && shown.canonicalActivityRevision;
    if (!Number.isSafeInteger(revision)) { toast('error', 'Atualize os dados da atividade antes de confirmar.'); return; }
    var shown = state.details[idAtividade] || state.activities.find(function(item) { return item.idAtividade === idAtividade; });
    var revision = shown && shown.canonicalActivityRevision;
    if (!Number.isSafeInteger(revision)) { toast('error', 'Atualize os dados da atividade antes de confirmar.'); return; }
    if (!global.confirm(messages[action] || 'Confirmar esta acao?')) return;
    var originalText = button && button.textContent || '';
    if (button) {
      button.disabled = true;
      button.textContent = 'Enviando...';
    }
    ui.mostrarLoading('Atualizando atividade...');
    api.apiPost('/admin/atividades/' + action, { payload: JSON.stringify({ idAtividade: idAtividade, revision: revision }) })
      .then(function onResponse(response) {
        if (!response.ok) throw response;
        delete state.details[idAtividade];
        notificarAtividadesAtualizadas(idAtividade, action.toUpperCase());
        var warnings = response.warnings || response.avisos || response.data && response.data.warnings || [];
        toast(warnings.length ? 'warning' : 'success', warnings.length
          ? (warnings[0].message || warnings[0].mensagem || response.message || 'Acao registrada com pendencias secundarias.')
          : (response.message || 'Acao concluida.'));
        closeModal();
        loadActivities(true);
      })
      .catch(function onError(error) { toast('error', error && error.message ? error.message : 'Nao foi possivel concluir a acao.'); })
      .finally(function done() {
        if (button) {
          button.disabled = false;
          button.textContent = originalText;
        }
        ui.ocultarLoading();
      });
  }

  function filterActivities(items, filters) {
    return (items || []).filter(function(item) {
      if (filters.statusOperacional && item.statusOperacional !== filters.statusOperacional) return false;
      if (filters.statusPublicacao && item.statusPublicacaoPortal !== filters.statusPublicacao) return false;
      if (filters.visibilidade && item.visibilidadePortal !== filters.visibilidade) return false;
      if (filters.ciclo && (item.ciclo || item.rotuloSemestre) !== filters.ciclo) return false;
      if (filters.modelo && item.idConfigModelo !== filters.modelo) return false;
      if (filters.tipo && item.tipoAtividade !== filters.tipo) return false;
      if (filters.subtipo && item.subtipoAtividade !== filters.subtipo) return false;
      if (filters.pendencia && !getActivePendencies(item).some(function(pending) { return pendingKind(pending) === filters.pendencia; })) return false;
      if (filters.texto) {
        var search = String(filters.texto).toLowerCase();
        var text = [item.idAtividade, item.tituloPublico, item.nomeModeloPortal, item.nomePessoaPrincipalPublico].join(' ').toLowerCase();
        if (text.indexOf(search) < 0) return false;
      }
      return true;
    });
  }

  function notificarAtividadesAtualizadas(idAtividade, acao) {
    document.dispatchEvent(new CustomEvent('portal:atividades-atualizadas', {
      detail: {
        idAtividade: String(idAtividade || ''),
        acao: String(acao || '')
      }
    }));
  }

  function getActivePendencies(item) {
    var status = String((item || {}).statusOperacional || '').toUpperCase();
    if (['CANCELADA', 'CANCELADO', 'ARQUIVADA', 'ARQUIVADO'].indexOf(status) >= 0) return [];
    return Array.isArray((item || {}).pendencias) ? item.pendencias : [];
  }

  function pendingKind(value) {
    return typeof value === 'string' ? value : value && typeof value.tipo === 'string' ? value.tipo : '';
  }

  function textField(name, label, value, type) {
    return '<label><span>' + escape(label) + '</span><input type="' + (type || 'text') + '" name="' + name + '" value="' + escape(value || '') + '"></label>';
  }

  function textareaField(name, label, value) {
    return '<label class="admin-field-wide"><span>' + escape(label) + '</span><textarea name="' + name + '" rows="3">' + escape(value || '') + '</textarea></label>';
  }

  function selectField(name, label, values, selected) {
    return '<label><span>' + escape(label) + '</span><select name="' + name + '">' + values.map(function(value) {
      return '<option value="' + escape(value) + '"' + (value === selected ? ' selected' : '') + '>' + escape(formatLabel(value)) + '</option>';
    }).join('') + '</select></label>';
  }

  function openModal(title, content) {
    var modal = document.getElementById('atividade-modal');
    var titleNode = document.getElementById('atividade-modal-title');
    var contentNode = document.getElementById('atividade-modal-content');
    if (!modal || !contentNode) return;
    if (titleNode) titleNode.textContent = title;
    contentNode.innerHTML = content;
    modal.hidden = false;
    document.body.classList.add('modal-open');
  }

  function closeModal() {
    var modal = document.getElementById('atividade-modal');
    if (modal) modal.hidden = true;
    document.body.classList.remove('modal-open');
  }

  function showModalError(error) {
    openModal('Gestao da atividade', '<p class="empty-state readonly-error">' + escape(adminErrorMessage(error, 'Erro ao carregar detalhe.')) + '</p><button class="secondary-button" type="button" data-admin-modal-close>Fechar</button>');
  }

  function setFormBusy(form, busy) {
    Array.prototype.forEach.call(form.querySelectorAll('button, input, select, textarea'), function(control) { control.disabled = !!busy; });
  }

  function normalizeTimeInput(value) { return String(value || '').trim().replace('h', ':'); }
  function adminErrorMessage(error, fallback) {
    var code = String(error && (error.code || error.errorCode) || '').toUpperCase();
    var message = String(error && error.message || '');
    if (code === 'ACAO_NAO_RECONHECIDA' || /acao nao reconhecida/i.test(message.normalize ? message.normalize('NFD').replace(/[\u0300-\u036f]/g, '') : message)) {
      return 'O backend publicado ainda nao possui a gestao de atividades. Atualize o Apps Script e publique uma nova versao da API.';
    }
    return message || fallback;
  }
  function formatLabel(value) { if (typeof value === 'boolean') return value ? 'Sim' : 'Nao'; return ui.formatarRotulo ? ui.formatarRotulo(value || '') : String(value || ''); }
  function toast(type, message) { ui.mostrarToast({ type: type, message: message }); }
  function escape(value) { return ui.escaparHtml(String(value === null || value === undefined ? '' : value)); }
  function getContainer() { return document.getElementById('placeholder-content'); }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', start);
  else start();
})(window);
