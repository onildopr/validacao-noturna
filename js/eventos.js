// Eventos da página (cliques, leitor de código de barras) e inicialização.

// =======================
// Eventos / Boot
// =======================
$(document).ready(async () => {
  $(document).on('click', '#db-search-open', () => {
    $('#initial-interface').addClass('d-none');
    $('#db-search-interface').removeClass('d-none');
    $('#db-search-results-wrap').addClass('d-none');
  });

  $(document).on('click', '#db-search-back', () => {
    $('#db-search-interface').addClass('d-none');
    $('#db-search-results-wrap').addClass('d-none');
    $('#initial-interface').removeClass('d-none');
  });

  // Status da busca: no index vai para o painel lateral, no search.html para #db-search-status
  const searchStatus = (txt, kind) => {
    ConferenciaApp.setStatus(txt, kind);
    $('#db-search-status').text(txt);
  };

  const runDbSearch = async () => {
    try {
      const rawIds = ($('#db-ids').val() || '').trim();
      if (!rawIds) { alert('Informe pelo menos um ID.'); return; }

      const dayFrom = ($('#db-day-from').val() || '').trim() || undefined;
      const dayTo = ($('#db-day-to').val() || '').trim() || undefined;
      const op = ($('#db-op-filter').val() || $('#db-op-code').val() || '').trim() || undefined;

      searchStatus('Buscando histórico no banco...', 'info');

      const res = await ConferenciaApp.searchIdsFull(rawIds, {
        operation_code: op,
        day_from: dayFrom,
        day_to: dayTo
      });

      if (!res.ids.length) {
        searchStatus('Nenhum ID válido informado.', 'warning');
        return;
      }

      ConferenciaApp.renderDbSearchSummary(res.summary);
      searchStatus(`Busca concluída • ${res.summary.length} ID(s)`, 'success');
    } catch (e) {
      console.error(e);
      searchStatus('Erro ao buscar histórico.', 'danger');
      alert('Erro na busca: ' + (e?.message || e));
    }
  };

  $(document).on('click', '#db-search-btn', runDbSearch);

  $(document).on('keydown', '#db-ids', (e) => {
    if (e.ctrlKey && e.key === 'Enter') runDbSearch();
  });

  $(document).on('click', '#db-clear-btn', () => {
    $('#db-ids').val('');
    $('#db-search-results').empty();
    $('#db-search-results-wrap').addClass('d-none');
    $('#db-search-status').text('—');
  });

  // search.html: página só de busca, não carrega rotas/realtime
  if ($('#db-search-page').length) return;

  const today = ConferenciaApp.todayLocalISO();
  $('#work-day').val(today);

  await ConferenciaApp.ensureOperationSelected();

  if (ConferenciaApp.getOperationCode()) {
    await ConferenciaApp.applyWorkDay(today);
  }
});

// Confirmar operação escolhida
$(document).on('click', '#btn-op-confirm', async () => {
  const code = String($('#op-select').val() || '').trim().toUpperCase();
  if (!code) return;
  await ConferenciaApp.flushPendingNow();
  ConferenciaApp.setOperationCode(code);
  ConferenciaApp.resetForOperationChange();
  $('#modal-operation').modal('hide');

  const day = $('#work-day').val() || ConferenciaApp.todayLocalISO();
  await ConferenciaApp.applyWorkDay(day);
});

// Trocar operação
$(document).on('click', '#btn-change-op', async () => {
  await ConferenciaApp.ensureOperationSelected();
  $('#modal-operation').modal('show');
});

// Filtro de rotas
$(document).on('focus mousedown keydown input', '#saved-routes-inapp, #saved-routes, #route-search', () => {
  ConferenciaApp.lockRouteUi(3000);
});

$(document).on('focus', '#saved-routes-inapp', () => {
  ConferenciaApp.isRouteDropdownOpen = true;
  ConferenciaApp.lockRouteUi(3000);
});

$(document).on('blur change', '#saved-routes-inapp', () => {
  ConferenciaApp.isRouteDropdownOpen = false;
  ConferenciaApp.lockRouteUi(800);
});

$(document).on('input', '#route-search', (e) => {
  ConferenciaApp.lockRouteUi(3000);
  ConferenciaApp.applyRouteDropdownFilter(e.target.value);
});

// Troca de dia
$(document).on('change', '#work-day', async (e) => {
  const day = e.target.value;
  if (!day) return;
  await ConferenciaApp.applyWorkDay(day);
});

// Relatório noturno
$(document).on('click', '#finish-night-btn', () => {
  try {
    ConferenciaApp.showNightReport();
    const el = document.querySelector('#night-report');
    if (el) el.scrollIntoView({ block: 'start', behavior: 'smooth' });
  } catch (e) {
    console.error(e);
    alert('Falha ao gerar o relatório noturno.');
  }
});

$(document).on('click', '#finish-btn', () => {
  try {
    ConferenciaApp.showNightReport();
    const el = document.querySelector('#night-report');
    if (el) el.scrollIntoView({ block: 'start', behavior: 'smooth' });
  } catch (e) {
    console.error(e);
    alert('Falha ao gerar o relatório noturno.');
  }
});

$(document).on('click', '#night-report-close', () => {
  ConferenciaApp.hideNightReport();
});

// Importar HTML
$('#extract-btn').click(() => {
  const raw = $('#html-input').val();
  if (!raw.trim()) return alert('Cole o HTML antes de importar.');

  const rep = ConferenciaApp.importRoutesFromHtml(raw);
  alert(ConferenciaApp.formatImportReport(rep));

  // Só limpa o campo se importou algo (se deu erro, o usuário pode conferir o que colou)
  if (rep.importadas) $('#html-input').val('');
});

// Carregar rota
$('#load-route').click(() => {
  const id = $('#saved-routes').val();
  if (!id) return alert('Selecione uma rota salva.');

  ConferenciaApp.setCurrentRoute(id);

  $('#initial-interface').addClass('d-none');
  $('#manual-interface').addClass('d-none');
  $('#conference-interface').removeClass('d-none');
  $('#barcode-input').focus();
});

// Excluir rota
$('#delete-route').click(async () => {
  const id = $('#saved-routes').val();
  if (!id) return alert('Selecione uma rota para excluir.');

  const r = ConferenciaApp.routes.get(String(id));
  const nome = r && r.cluster ? `CLUSTER ${r.cluster} (rota ${id})` : `rota ${id}`;
  if (!confirm(`Excluir a ${nome}?

As bipagens dela deixam de contar. Se importar de novo, ela começa zerada.`)) return;
  if (!(await ConferenciaApp.requirePin(`Excluir ${nome}`))) return;

  ConferenciaApp.deleteRoute(id);
});

// Limpar todas
$('#clear-all-routes').click(async () => {
  const ok1 = confirm(
    'ATENÇÃO: isso vai APAGAR TODAS as rotas do DIA selecionado.\n\n' +
    'Quer continuar?'
  );
  if (!ok1) return;

  const day = ConferenciaApp.workDay || $('#work-day').val() || '(dia desconhecido)';
  const typed = prompt(
    `CONFIRMAÇÃO FINAL\n\n` +
    `Para apagar TUDO do dia ${day}, digite exatamente:\n` +
    `APAGAR\n\n` +
    `(Qualquer outra coisa cancela)`
  );

  if (typed !== 'APAGAR') {
    alert('Ação cancelada. Nada foi apagado.');
    return;
  }
  if (!(await ConferenciaApp.requirePin(`Apagar todas as rotas do dia ${day}`))) return;

  ConferenciaApp.clearAllRoutes();
  alert(`Tudo do dia ${day} foi removido.`);
});

// Trocar rota
$('#switch-route').click(() => {
  const id = $('#saved-routes-inapp').val();
  if (!id) return;
  ConferenciaApp.setCurrentRoute(id);
  $('#barcode-input').focus();
});

// Manual
$('#manual-btn').click(() => {
  $('#initial-interface').addClass('d-none');
  $('#manual-interface').removeClass('d-none');
});

$('#submit-manual').click(() => {
  try {
    const routeId = ($('#manual-routeid').val() || '').trim();
    if (!routeId) return alert('Informe o RouteId.');

    const cluster = ($('#manual-cluster').val() || '').trim();
    const brutos = ($('#manual-input').val() || '').split(/[\s,;]+/).map(x => x.trim()).filter(Boolean);
    // Mesmo formato da bipagem (11 dígitos), senão o ID nunca bateria na conferência
    const manualIds = Array.from(new Set(brutos.map(x => ConferenciaApp.normalizarCodigo(x)).filter(Boolean)));
    const ignorados = brutos.filter(x => !ConferenciaApp.normalizarCodigo(x)).length;

    if (!manualIds.length) return alert('Nenhum ID válido inserido (esperado: 11 dígitos começando com 4).');
    if (ignorados && !confirm(`${ignorados} valor(es) não parecem IDs válidos e serão ignorados. Continuar?`)) return;

    const route = ConferenciaApp.routes.get(String(routeId)) || ConferenciaApp.makeEmptyRoute(routeId);
    route.cluster = cluster || route.cluster;

    for (const id of manualIds) {
      route.ids.add(id);
      if (!route.conferidos.has(id)) route.faltantes.add(id);
    }

    route.totalInicial = route.ids.size;
    ConferenciaApp.routes.set(String(routeId), route);

    ConferenciaApp.lastRoutesSignature = '';
    ConferenciaApp.saveToStorage(ConferenciaApp.workDay);
    ConferenciaApp.markDirty('inserção manual');

    ConferenciaApp.renderRoutesSelects();

    alert(`Rota ${routeId} salva com ${route.totalInicial} ID(s).`);

    $('#manual-interface').addClass('d-none');
    $('#initial-interface').removeClass('d-none');
  } catch (e) {
    console.error(e);
    alert('Erro ao processar IDs manuais.');
  }
});

// Leitura do barcode
$('#barcode-input').on('keypress', (e) => {
  if (e.which === 13) {
    ConferenciaApp.viaCsv = false;

    const raw = $('#barcode-input').val();
    const id = ConferenciaApp.normalizarCodigo(raw);

    if (!id) {
      $('#barcode-input').val('').focus();
      return;
    }

    ConferenciaApp.conferirId(id);
  }
});

// Checar CSV
$('#check-csv').click(() => {
  const r = ConferenciaApp.current;
  if (!r) return alert('Selecione uma rota antes.');

  const fileInput = document.getElementById('csv-input');
  if (fileInput.files.length === 0) return alert('Selecione um arquivo CSV.');

  ConferenciaApp.viaCsv = true;

  const file = fileInput.files[0];
  const reader = new FileReader();

  reader.onload = e => {
    const csvText = e.target.result;
    const linhas = csvText.split(/\r?\n/);
    if (!linhas.length) return alert('Arquivo CSV vazio.');

    const header = linhas[0].split(',');
    const textCol = header.findIndex(h => /(text|texto|id)/i.test(h));
    if (textCol === -1) return alert('Coluna apropriada não encontrada (text/texto/id).');

    for (let i = 1; i < linhas.length; i++) {
      if (!linhas[i].trim()) continue;
      const cols = linhas[i].split(',');
      if (cols.length <= textCol) continue;

      let campo = cols[textCol].trim().replace(/^"|"$/g, '').replace(/""/g, '"');
      const id = ConferenciaApp.normalizarCodigo(campo);
      if (id) ConferenciaApp.conferirId(id);
    }

    ConferenciaApp.viaCsv = false;
    $('#barcode-input').focus();
  };

  reader.readAsText(file, 'UTF-8');
});

// Exports
$(document).on('click', '#export-csv-rota-atual', () => {
  ConferenciaApp.exportRotaAtualCsvPadrao();
});

$(document).on('click', '#export-xlsx-todas-rotas', () => {
  ConferenciaApp.exportTodasRotasXlsx();
});

$('#back-btn').click(() => {
  $('#conference-interface').addClass('d-none');
  $('#manual-interface').addClass('d-none');
  $('#initial-interface').removeClass('d-none');

  $('#barcode-input').val('');
  $('#html-input').focus();
});

// Carretas
$(document).on('click', '#carreta-btn', () => {
  $('#initial-interface').addClass('d-none');
  $('#conference-interface').addClass('d-none');
  $('#manual-interface').addClass('d-none');
  $('#carreta-interface').removeClass('d-none');

  ConferenciaApp.renderCarretaUI();
  $('#carreta-input').val('').focus();
});

$(document).on('click', '#carreta-back-btn', () => {
  $('#carreta-interface').addClass('d-none');
  $('#initial-interface').removeClass('d-none');
  $('#carreta-input').val('');
  $('#html-input').focus();
});

$(document).on('click', '#carreta-clear-current', () => {
  ConferenciaApp.carretas.currentPlateKey = null;
  ConferenciaApp.renderCarretaUI();
  $('#carreta-input').val('').focus();
});

$(document).on('click', '#patio-refresh', function() {
  ConferenciaApp.renderPatioGeral();
});

$(document).on('click', '#carreta-refresh-progress', () => {
  ConferenciaApp.renderCarretaProgress();
});

const processCarretaScan = (rawValue) => {
  const raw = String(rawValue || '').trim();
  if (!raw) return;

  const parsed = ConferenciaApp.parseScanPayload(raw);

  if (parsed.kind === 'plate') {
    const key = ConferenciaApp.ensurePlate(parsed.plate);
    ConferenciaApp.carretas.currentPlateKey = key;
    ConferenciaApp.renderCarretaUI();
    return;
  }

  if (parsed.kind === 'routeqr') {
    const pk = ConferenciaApp.carretas.currentPlateKey;
    if (!pk) {
      alert('Bipe uma PLACA primeiro.');
      return;
    }
    ConferenciaApp.vincularRouteQrNaPlaca(
      parsed.routeKey,
      {
        raw: (parsed.route && parsed.route.raw) ? parsed.route.raw : raw,
        jsonText: (parsed.route && parsed.route.jsonText) ? parsed.route.jsonText : ((parsed.route && parsed.route.obj) ? JSON.stringify(parsed.route.obj) : '')
      },
      pk,
      parsed.routeIdCandidate || ''
    );
    ConferenciaApp.renderCarretaUI();
    return;
  }

  if (parsed.kind === 'shipment') {
    alert('Aqui é a tela da CARRETA. Bipe a PLACA e os QRs das ROTAS (assignment/container).');
    return;
  }

  alert('QR não reconhecido. Bipe uma PLACA (JSON com license_plate) ou um QR de ROTA (JSON com assignment/container_id).');
};

$(document).on('keydown', '#carreta-input', (e) => {
  if (e.key === 'Enter' || e.which === 13) {
    e.preventDefault();
    const raw = $('#carreta-input').val();
    $('#carreta-input').val('');
    processCarretaScan(raw);
  }
});

$(document).on('click', '#carreta-check-links', () => {
  ConferenciaApp.checkLinksForCurrentPlate();
});

$(document).on('click', '#carreta-clear-bipagem-plate', async () => {
  const pk = ConferenciaApp.carretas.currentPlateKey;
  if (!pk) return alert('Nenhuma placa ativa.');

  const ok = confirm(`Tem certeza que deseja EXCLUIR a bipagem/vínculos da placa ${pk}?`);
  if (!ok) return;
  if (!(await ConferenciaApp.requirePin(`Excluir bipagem da placa ${pk}`))) return;

  ConferenciaApp.clearBipagemForPlate(pk);
});

$(document).on('paste', '#carreta-input', (e) => {
  const pasted = (e.originalEvent && e.originalEvent.clipboardData)
    ? e.originalEvent.clipboardData.getData('text')
    : '';
  setTimeout(() => {
    const raw = $('#carreta-input').val() || pasted;
    $('#carreta-input').val('');
    processCarretaScan(raw);
  }, 0);
});

// Exports novos
$(document).on('click', '#export-csv-rota-atual-placa', () => {
  ConferenciaApp.exportRotaAtualCsvComPlacaERota();
});

$(document).on('click', '#export-csv-todas-rotas-placa', () => {
  ConferenciaApp.exportTodasRotasCsvComPlacaERota();
});

$(document).on('click', '#export-csv-mapa-carretas', () => {
  ConferenciaApp.exportMapaCarretasCsv();
});

// Admin UI
$(document).on('click', '#btn-admin-open', async () => {
  $('#modal-admin').modal('show');
  await refreshAdminOps();
});

// Atalho para o Admin a partir do modal de operação (útil quando ainda não há operações)
$(document).on('click', '#btn-op-admin', () => {
  $('#modal-operation').one('hidden.bs.modal', () => $('#btn-admin-open').trigger('click'));
  $('#modal-operation').modal('hide');
});

// Ao fechar o Admin, recarrega a lista de operações (reabre a seleção se ainda não houver operação válida)
$(document).on('hidden.bs.modal', '#modal-admin', () => {
  ConferenciaApp.ensureOperationSelected();
});

// Com dois modais abertos (Admin + PIN), fechar o do PIN não pode destravar a rolagem do Admin
$(document).on('hidden.bs.modal', '#modal-pin', () => {
  if ($('.modal.show').length) $('body').addClass('modal-open');
});

async function refreshAdminOps() {
  try {
    const ops = await ConferenciaApp.adminLoadOperations(true);
    const $tbody = $('#admin-ops-tbody');
    if (!$tbody.length) return;
    $tbody.empty();
    ops.forEach(o => {
      const act = o.active ? 'SIM' : 'NÃO';
      const name = o.name || '';
      $tbody.append(`<tr><td>${ConferenciaApp.escHtml(o.code)}</td><td>${ConferenciaApp.escHtml(name)}</td><td>${act}</td></tr>`);
    });
  } catch (e) {
    console.warn(e);
  }
}

$(document).on('click', '#btn-admin-save-op', async () => {
  const code = $('#admin-op-code').val();
  const name = $('#admin-op-name').val();
  const active = $('#admin-op-active').is(':checked');
  try {
    if (!(await ConferenciaApp.requirePin('Cadastrar / alterar operação no Admin'))) return;

    await ConferenciaApp.adminUpsertOperation(code, name, active);
    await refreshAdminOps();
    alert('Operação salva.');
  } catch (e) {
    console.error(e);
    alert('Erro ao salvar operação: ' + (e.message || e));
  }
});

// Acompanhamento geral
$(document).on('click', '#btn-global-acomp', async () => {
  try {
    const day = $('#work-day').val() || ConferenciaApp.todayLocalISO();

    $('#initial-interface, #carreta-interface, #manual-interface, #conference-interface, #db-search-interface').addClass('d-none');
    $('#global-interface').removeClass('d-none');

    ConferenciaApp.setStatus(`Carregando acompanhamento geral • ${day}`, 'info');
    const items = await ConferenciaApp.loadGlobalProgress(day);
    ConferenciaApp.renderGlobalProgress(items, day);
    ConferenciaApp.setStatus(`Acompanhamento geral carregado • ${day}`, 'success');
  } catch (e) {
    console.warn(e);
    ConferenciaApp.setStatus('Falha ao carregar acompanhamento geral (ver console).', 'danger');
  }
});

$(document).on('click', '#global-refresh', async () => {
  try {
    const day = $('#work-day').val() || ConferenciaApp.todayLocalISO();
    const items = await ConferenciaApp.loadGlobalProgress(day);
    ConferenciaApp.renderGlobalProgress(items, day);
  } catch (e) {
    console.warn(e);
    ConferenciaApp.setStatus('Falha ao atualizar acompanhamento geral.', 'danger');
  }
});

$(document).on('click', '#global-back', () => {
  $('#global-interface').addClass('d-none');
  $('#initial-interface').removeClass('d-none');
});

// encerra realtime ao sair
window.addEventListener('beforeunload', () => {
  try { ConferenciaApp.persistEventsNow(); } catch {}
  try { ConferenciaApp.stopRealtimeSync(); } catch {}
});

// Ao voltar para a aba, sincroniza na hora (a checagem periódica pausa com a aba oculta)
document.addEventListener('visibilitychange', () => {
  if (!document.hidden) ConferenciaApp.periodicSyncTick();
});