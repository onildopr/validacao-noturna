// Renderização da interface (listas, selects, relatório noturno, carretas).

Object.assign(ConferenciaApp, {
  // Re-renderiza a tela depois de mesclar dados vindos do banco
  refreshAfterRemoteMerge() {
    if (this._rtUiTimer) clearTimeout(this._rtUiTimer);
    this._rtUiTimer = setTimeout(() => {
      try {
        const keepRouteId = this.currentRouteId;
        if (!this.isRouteUiLocked()) {
          this.renderRoutesSelects();
        } else if (keepRouteId && this.routes.has(String(keepRouteId))) {
          this.currentRouteId = String(keepRouteId);
        }
        this.refreshUIFromCurrent();
        this.renderAcompanhamento();
        if (!$('#carreta-interface').hasClass('d-none')) this.renderCarretaUI();
      } catch (e) {
        console.warn('Falha ao renderizar após sync:', e);
      }
    }, 250);
  },

  // Mostra o que ainda não chegou no banco (e avisa com destaque quando está sem conexão)
  updatePendingFlag() {
    const nEventos = this.eventQueue ? this.eventQueue.size : 0;
    this.dirty = !!(this.defsDirty || nEventos);

    const partes = [];
    if (nEventos) partes.push(`${nEventos} bipagem(ns)`);
    if (this.defsDirty) partes.push('rotas');
    const texto = partes.length ? `${partes.join(' + ')} aguardando envio` : '';

    const $flag = $('#dirty-flag');
    $flag.text(texto || 'pendente');
    if (this.dirty) $flag.removeClass('d-none');
    else $flag.addClass('d-none');

    const $alert = $('#pending-alert');
    if (this.dirty && this.cloudOffline) {
      $alert.html(
        `<strong>Sem conexão com o banco.</strong><br>${this.escHtml(texto)} — guardado neste aparelho, ` +
        'será enviado quando a conexão voltar. Não limpe os dados do navegador.'
      ).removeClass('d-none');
    } else {
      $alert.addClass('d-none').empty();
    }
  },

  // Pede o PIN (único, conferido no banco) antes de uma ação destrutiva. Resolve true se liberado.
  // Sem PIN cadastrado => liberado direto. PIN certo vale por 10 minutos neste aparelho.
  // Sem conexão não dá para conferir o PIN => a ação é bloqueada.
  async requirePin(acao) {
    if (Date.now() < this.pinOkUntil) return true;

    let temPin;
    try {
      temPin = await this.pinEnabled();
    } catch (e) {
      console.warn('Falha ao consultar o PIN:', e);
      alert('Sem conexão com o banco: não foi possível conferir o PIN. Tente de novo quando a conexão voltar.');
      return false;
    }
    if (!temPin) return true;

    return new Promise((resolve) => {
      const $m = $('#modal-pin');
      const $in = $('#pin-input');
      const $err = $('#pin-error');
      const $btn = $('#pin-confirm');
      $('#pin-acao').text(acao);
      $in.val('');
      $err.addClass('d-none');

      let done = false;
      const finish = (ok) => {
        if (done) return;
        done = true;
        $btn.off('click.pin').prop('disabled', false);
        $in.off('keydown.pin');
        $m.off('hidden.bs.modal.pin');
        $m.modal('hide');
        resolve(ok);
      };
      const tentar = async () => {
        if ($btn.prop('disabled')) return;
        $btn.prop('disabled', true);
        try {
          if (await this.verifyPin($in.val())) {
            this.pinOkUntil = Date.now() + 10 * 60 * 1000;
            finish(true);
            return;
          }
          $err.text('PIN incorreto.').removeClass('d-none');
        } catch (e) {
          console.warn('Falha ao conferir o PIN:', e);
          $err.text('Sem conexão com o banco. Tente de novo.').removeClass('d-none');
        }
        $btn.prop('disabled', false);
        $in.val('').focus();
      };

      $btn.on('click.pin', tentar);
      $in.on('keydown.pin', (e) => { if (e.key === 'Enter') tentar(); });
      $m.on('hidden.bs.modal.pin', () => finish(false));
      $m.one('shown.bs.modal', () => $in.trigger('focus'));
      $m.modal('show');
    });
  },

  // Tela de bloqueio do primeiro acesso do dia neste aparelho. Resolve quando a senha do dia é aceita.
  ensureDailyUnlock() {
    if (this.isUnlockedToday()) return Promise.resolve();

    return new Promise((resolve) => {
      const lock = document.createElement('div');
      lock.id = 'daily-lock';
      lock.setAttribute('role', 'dialog');
      lock.setAttribute('aria-modal', 'true');
      lock.style.cssText =
        'position:fixed;inset:0;z-index:3000;background:#000;display:flex;align-items:center;justify-content:center;padding:16px;';
      lock.innerHTML = `
        <div style="background:#fff;color:#212529;border-left:4px solid #ff8c00;border-radius:6px;padding:24px;width:100%;max-width:360px;text-align:center;">
          <img src="logorodacoop.png" alt="" style="max-width:160px;margin-bottom:12px;" onerror="this.remove()">
          <h5 style="color:#ff8c00;margin-bottom:4px;">Senha do dia</h5>
          <div class="small text-muted mb-3">Primeiro acesso de hoje neste aparelho</div>
          <input id="daily-lock-input" type="password" inputmode="numeric" autocomplete="off"
                 class="form-control text-center mb-2" placeholder="Senha" aria-label="Senha do dia">
          <div id="daily-lock-error" class="text-danger small mb-2" style="display:none;">Senha incorreta.</div>
          <button id="daily-lock-btn" type="button" class="btn btn-primary btn-block">Entrar</button>
        </div>`;
      document.body.appendChild(lock);

      const input = lock.querySelector('#daily-lock-input');
      const err = lock.querySelector('#daily-lock-error');
      const tentar = () => {
        if (this.checkDailyPassword(input.value)) {
          this.markUnlockedToday();
          lock.remove();
          resolve();
        } else {
          err.style.display = 'block';
          input.value = '';
          input.focus();
        }
      };
      lock.querySelector('#daily-lock-btn').addEventListener('click', tentar);
      input.addEventListener('keydown', (e) => { if (e.key === 'Enter') tentar(); });
      setTimeout(() => input.focus(), 0);
    });
  },

  async ensureOperationSelected() {
    if (!this.getSb()) return;

    let data;
    try {
      data = await this.loadOperationsRows(false);
    } catch (error) {
      console.warn('Falha ao carregar operações:', error);
      return;
    }

    const $sel = $('#op-select');
    if ($sel.length) {
      $sel.empty();
      (data || []).forEach(op => {
        const label = op.name ? `${op.code} — ${op.name}` : op.code;
        $sel.append(`<option value="${this.escHtml(op.code)}">${this.escHtml(label)}</option>`);
      });
    }

    const current = this.getOperationCode();
    const exists = (data || []).some(o => o.code === current);

    if ($sel.length && current && exists) {
      $sel.val(current);
    }

    if (!current || !exists) {
      $('#modal-operation').modal({ backdrop: 'static', keyboard: false });
    } else {
      this.setOperationCode(current);
    }
  },

  renderDbSearchSummary(summaryRows) {
    const $tb = $('#db-search-results');
    const $wrap = $('#db-search-results-wrap');
    if (!$tb.length) return;

    $tb.empty();

    (summaryRows || []).forEach(r => {
      const dt = r.last_seen_at ? new Date(r.last_seen_at) : null;
      const day = r.last_day || (dt ? dt.toISOString().slice(0, 10) : '');
      const hhmm = dt ? String(dt.getHours()).padStart(2, '0') + ':' + String(dt.getMinutes()).padStart(2, '0') : '';

      const ops = (r.operations && r.operations.length) ? r.operations.join(',') : '';
      const status = r.has_db_history
        ? (r.last_result || '')
        : (r.local_status ? `local:${r.local_status}` : 'sem histórico');

      const routeId = r.has_db_history ? (r.last_route_id ?? '') : (r.local_route_id ?? '');
      const cluster = r.has_db_history ? (r.last_cluster ?? '') : (r.local_cluster ?? '');
      const xpt = r.has_db_history ? (r.last_xpt ?? '') : (r.local_xpt ?? '');

      const esc = (x) => this.escHtml(x);
      $tb.append(`
        <tr>
          <td>${esc(r.id)}</td>
          <td>${esc(r.has_db_history ? (r.last_operation ?? '') : '')}</td>
          <td>${esc(day)}</td>
          <td>${hhmm}</td>
          <td>${esc(routeId)}</td>
          <td>${esc(cluster)}</td>
          <td>${esc(xpt)}</td>
          <td>${esc(status)} ${ops ? `<small class="text-muted">(${esc(ops)})</small>` : ''}</td>
        </tr>
      `);
    });

    if ($wrap.length) $wrap.removeClass('d-none');
  },

  // Log local das últimas bipagens (painel "Acompanhamento")
  pushEvent(evt) {
    this.lastEvents.unshift(Object.assign({}, evt));
    if (this.lastEvents.length > 80) this.lastEvents.length = 80;
    this.renderAcompanhamento();
  },

  renderAcompanhamento() {
    this.updatePendingFlag();

    const mapa = new Map();

    for (const r of this.routes.values()) {
      const cluster = (r.cluster && String(r.cluster).trim()) || '(sem cluster)';
      const total = (r.totalInicial || r.ids.size || 0);
      const conf = (r.conferidos ? r.conferidos.size : 0);

      if (!mapa.has(cluster)) {
        mapa.set(cluster, {
          conferidos: 0,
          total: 0,
          precisaRevalidar: false
        });
      }

      const acc = mapa.get(cluster);
      acc.conferidos += conf;
      acc.total += total;

      if (conf < total) {
        acc.precisaRevalidar = true;
      }
    }

    const resumo = Array.from(mapa.entries())
      .sort((a, b) => String(a[0]).localeCompare(String(b[0])))
      .map(([cluster, v]) => {
        const ok = (v.total > 0 && v.conferidos === v.total) ? ' ✅' : '';
        return `CLUSTER ${this.escHtml(cluster)}: ${v.conferidos}/${v.total}${ok}`;
      })
      .join('<br>');

    $('#acompanhamento-resumo')
      .css({
        'max-height': '200px',
        'overflow-y': 'auto',
        'overflow-x': 'hidden'
      })
      .html(resumo || '<span class="text-muted">sem clusters</span>');

    const textoClusters = Array.from(mapa.entries())
      .sort((a, b) => String(a[0]).localeCompare(String(b[0])))
      .map(([cluster, v]) => {
        return `CLUSTER ${cluster}: ${v.precisaRevalidar ? 'REVALIDAR' : 'CONTAR'}`;
      })
      .join('\n');

    $('#clusters-copiavel').val(textoClusters);

    const items = this.lastEvents.slice(0, 30).map(ev => {
      const d = new Date(ev.ts);
      const hh = String(d.getHours()).padStart(2, '0');
      const mm = String(d.getMinutes()).padStart(2, '0');
      const ss = String(d.getSeconds()).padStart(2, '0');
      const base = `${hh}:${mm}:${ss} • ${this.escHtml(ev.code)}`;

      if (ev.type === 'fora') {
        const rr = ev.correctRouteId ? this.routes.get(String(ev.correctRouteId)) : null;
        const cl = rr && rr.cluster ? String(rr.cluster).trim() : '';
        const extra = (cl ? ` <small class="text-muted">(CLUSTER ${this.escHtml(cl)})</small>` : ' <small class="text-muted">(cluster desconhecido)</small>');
        return `<li class="list-group-item list-group-item-warning">${base}${extra}</li>`;
      }
      if (ev.type === 'dup') {
        return `<li class="list-group-item list-group-item-secondary">${base} (duplicado)</li>`;
      }
      return `<li class="list-group-item list-group-item-success">${base} (ok)</li>`;
    });

    $('#acompanhamento-log').html(items.join('') || '<li class="list-group-item text-muted">sem eventos</li>');
  },

  buildNightReportHtml() {
    const esc = (x) => this.escHtml(x);

    const routesSorted = Array.from(this.routes.values()).sort((a, b) => {
      const ca = String(a.cluster || '').trim();
      const cb = String(b.cluster || '').trim();
      const byC = ca.localeCompare(cb);
      if (byC) return byC;
      return String(a.routeId).localeCompare(String(b.routeId));
    });

    const items = routesSorted.map(r => {
      const total = Number(r.totalInicial || r.ids.size || 0);
      const conf = Number(r.conferidos?.size || 0);
      const falt = Number(r.faltantes?.size || 0);
      const ok = total > 0 ? (conf === total) : (falt === 0);
      return {
        routeId: String(r.routeId),
        cluster: String(r.cluster || '').trim(),
        total, conf, falt,
        ok,
        route: r
      };
    });

    const completas = items.filter(x => x.ok);
    const incompletas = items.filter(x => !x.ok);

    const headerLine = (it) => {
      const c = it.cluster ? `CLUSTER ${esc(it.cluster)}` : 'CLUSTER (vazio)';
      const perc = it.total ? Math.floor((it.conf / it.total) * 100) : 0;
      const badge = it.ok ? `<span class="badge badge-success ml-2">100%</span>` : `<span class="badge badge-warning ml-2">${esc(it.falt)} falt.</span>`;
      return `${c} <small class="text-muted">• Rota ${esc(it.routeId)} • ${esc(it.conf)}/${esc(it.total)} (${perc}%)</small>${badge}`;
    };

    const listCompletas = completas.length
      ? completas.map(it => `<li class="list-group-item d-flex justify-content-between align-items-center">${headerLine(it)}</li>`).join('')
      : `<li class="list-group-item text-muted">Nenhuma rota 100%.</li>`;

    const blocosIncompletas = incompletas.length
      ? incompletas.map((it, idx) => {
          const collapseId = `nr_${esc(it.routeId)}_${idx}`;
          const r = it.route;
          const faltantes = Array.from(r.faltantes || []);
          faltantes.sort((a, b) => String(a).localeCompare(String(b)));

          const faltHtml = faltantes.map(id => {
            const loc = this.locateIdInOtherRoutes(id, r.routeId);
            if (!loc) return `<li class="list-group-item list-group-item-danger">${esc(id)}</li>`;
            const whereTxt = (loc.where === 'conferido')
              ? 'conferido'
              : (loc.where === 'fora')
                ? 'bipado (fora de rota)'
                : 'pertence à rota';
            const cl = loc.cluster ? `CLUSTER ${esc(loc.cluster)}` : 'CLUSTER (vazio)';
            return `<li class="list-group-item list-group-item-warning">
                      ${esc(id)}
                      <small class="text-muted ml-2">→ ${whereTxt} em ${cl} (Rota ${esc(loc.routeId)})</small>
                    </li>`;
          }).join('') || `<li class="list-group-item text-muted">Sem faltantes</li>`;

          return `
            <div class="card mb-2">
              <div class="card-header p-2">
                <button class="btn btn-link p-0" type="button" data-toggle="collapse" data-target="#${collapseId}">
                  ${headerLine(it)}
                </button>
              </div>
              <div id="${collapseId}" class="collapse">
                <ul class="list-group list-group-flush">
                  ${faltHtml}
                </ul>
              </div>
            </div>
          `;
        }).join('')
      : `<div class="text-muted">Nenhuma rota com faltantes.</div>`;

    return `
      <div class="mb-2">
        <div class="small text-muted">Relatório do dia ${esc(this.workDay || this.todayLocalISO())}</div>
        <div class="small text-muted">Mostra todas as rotas, rotas 100% e faltantes (indicando onde foram vistos em outras rotas).</div>
      </div>

      <div class="mb-2">
        <span class="badge badge-success">100%: ${completas.length}</span>
        <span class="badge badge-warning ml-1">Com faltantes: ${incompletas.length}</span>
        <span class="badge badge-light ml-1">Total: ${routesSorted.length}</span>
      </div>

      <div class="mt-2">
        <div class="font-weight-bold mb-1">Rotas 100%</div>
        <ul class="list-group mb-3">
          ${listCompletas}
        </ul>

        <div class="font-weight-bold mb-1">Rotas com faltantes</div>
        ${blocosIncompletas}
      </div>
    `;
  },

  showNightReport() {
    const html = this.buildNightReportHtml();
    $('#night-report').html(html).removeClass('d-none');
    $('#acompanhamento-resumo').addClass('d-none');
    $('#acompanhamento-log').closest('div').addClass('d-none');
    $('#finish-night-btn').addClass('d-none');
    $('#night-report-close').removeClass('d-none');
  },

  hideNightReport() {
    $('#night-report').addClass('d-none').html('');
    $('#acompanhamento-resumo').removeClass('d-none');
    $('#acompanhamento-log').closest('div').removeClass('d-none');
    $('#finish-night-btn').removeClass('d-none');
    $('#night-report-close').addClass('d-none');
  },

  resetForOperationChange() {
    this.stopRealtimeSync();
    this.resetDayState();
    this.viaCsv = false;

    try {
      $('#saved-routes').html('<option value="">(Nenhuma selecionada)</option>');
      $('#saved-routes-inapp').empty();
      $('#carreta-routes').empty();
      $('#fora-rota-list').empty();
      $('#route-title').text('');
      $('#cluster-title').text('');
      $('#destination-facility-title').text('');
      $('#destination-facility-name').text('');
      $('#extracted-total').text('0');
      $('#verified-total').text('0');
    } catch {}

    $('#global-interface').addClass('d-none');
    $('#carreta-interface').addClass('d-none');
    $('#manual-interface').addClass('d-none');
    $('#initial-interface').removeClass('d-none');
  },

  renderGlobalProgress(items, dayISO) {
    $('#global-day-label').text(dayISO);
    const $tb = $('#global-ops-tbody');
    const $log = $('#global-log');
    $tb.empty();
    $log.empty();

    const sorted = (items || []).slice().sort((a, b) => (a.code || '').localeCompare(b.code || ''));

    for (const it of sorted) {
      const u = it.updated_at ? new Date(it.updated_at).toLocaleString('pt-BR') : '-';
      $tb.append(`
        <tr>
          <td><strong>${this.escHtml(it.code)}</strong>${it.name ? ` <span class="text-muted small">(${this.escapeHtml(it.name)})</span>` : ''}</td>
          <td>${it.stats.routesCount}</td>
          <td>${it.stats.totalIds}</td>
          <td>${it.stats.conferidos}</td>
          <td>${it.stats.faltantes}</td>
          <td>${it.stats.fora}</td>
          <td>${u}</td>
        </tr>
      `);

      $log.append(`
        <li class="list-group-item d-flex justify-content-between align-items-center">
          <span><strong>${this.escHtml(it.code)}</strong> • ${it.stats.conferidos}/${it.stats.totalIds} conferidos • ${it.stats.faltantes} faltantes</span>
          <span class="badge badge-light">${u}</span>
        </li>
      `);
    }

    if (!sorted.length) {
      $tb.append('<tr><td colspan="7" class="text-muted">Nenhuma operação ativa encontrada.</td></tr>');
    }
  },

  setCurrentRoute(routeId) {
    console.log('[setCurrentRoute] versão', '2026-02-22-1');

    const id = String(routeId);
    if (!this.routes.has(id)) {
      alert('Rota não encontrada.');
      return;
    }

    this.currentRouteId = id;
    this.renderRoutesSelects();
    this.refreshUIFromCurrent();
    this.renderAcompanhamento();
  },

  renderRoutesSelects() {
    const $sel1 = $('#saved-routes');
    const $sel2 = $('#saved-routes-inapp');

    const routesSorted = Array.from(this.routes.values())
      .sort((a, b) => String(a.routeId).localeCompare(String(b.routeId)));

    const makeLabel = (r) => {
      const parts = [];
      if (r.cluster) parts.push(`CLUSTER ${r.cluster}`);
      if (r.destinationFacilityId) parts.push(`XPT ${r.destinationFacilityId}`);
      return parts.join(' • ') || `(sem dados) • Rota ${r.routeId}`;
    };

    const newCache = routesSorted.map(r => ({
      routeId: String(r.routeId),
      label: makeLabel(r),
      clusterKey: this.normalizeCluster(r.cluster || ''),
      labelKey: this.normalizeCluster(makeLabel(r)),
    }));

    const sig = newCache
      .map(x => `${x.routeId}|${x.clusterKey}|${x.label}`)
      .join('||');

    const changed = sig !== this.lastRoutesSignature;

    if (changed) {
      this.lastRoutesSignature = sig;
      this._routesDropdownCache = newCache;

      $sel1.html(
        ['<option value="">(Nenhuma selecionada)</option>']
          .concat(this._routesDropdownCache.map(x => `<option value="${this.escHtml(x.routeId)}">${this.escHtml(x.label)}</option>`))
          .join('')
      );

      this.applyRouteDropdownFilter($('#route-search').val() || '');
    }

    if (this.currentRouteId) {
      $sel1.val(String(this.currentRouteId));

      const filteredText = $('#route-search').val() || '';
      const q = this.normalizeCluster(String(filteredText).trim());

      const filtered = !q
        ? (this._routesDropdownCache || [])
        : (this._routesDropdownCache || []).filter(x =>
            (x.clusterKey && x.clusterKey.includes(q)) ||
            (x.labelKey && x.labelKey.includes(q)) ||
            String(x.routeId).includes(String(filteredText).trim())
          );

      const existsInFiltered = filtered.some(x => x.routeId === String(this.currentRouteId));
      if (existsInFiltered) {
        $sel2.val(String(this.currentRouteId));
      }
    }
  },

  applyRouteDropdownFilter(filterText) {
    const $sel2 = $('#saved-routes-inapp');

    const list = Array.isArray(this._routesDropdownCache) ? this._routesDropdownCache : [];
    const q = this.normalizeCluster(String(filterText || '').trim());

    const filtered = !q
      ? list
      : list.filter(x =>
          (x.clusterKey && x.clusterKey.includes(q)) ||
          (x.labelKey && x.labelKey.includes(q)) ||
          String(x.routeId).includes(String(filterText || '').trim())
        );

    const options = ['<option value="">(Selecione)</option>']
      .concat(filtered.map(x => `<option value="${this.escHtml(x.routeId)}">${this.escHtml(x.label)}</option>`))
      .join('');

    $sel2.html(options);

    if (this.currentRouteId) {
      const exists = filtered.some(x => x.routeId === String(this.currentRouteId));
      if (exists) $sel2.val(String(this.currentRouteId));
    }
  },

  refreshUIFromCurrent() {
    const r = this.current;
    if (!r) {
      $('#route-title').html('');
      $('#cluster-title').html('');
      $('#destination-facility-title').html('');
      $('#destination-facility-name').html('');
      $('#extracted-total').text('0');
      $('#verified-total').text('0');
      $('#progress-bar').css('width', '0%').text('0%');
      $('#conferidos-list, #faltantes-list, #fora-rota-list, #duplicados-list').html('');
      return;
    }

    const esc = (x) => this.escHtml(x);
    $('#route-title').html(`ROTA: <strong>${esc(r.routeId)}</strong>`);
    $('#cluster-title').html(r.cluster ? `CLUSTER: <strong>${esc(r.cluster)}</strong>` : '');
    $('#destination-facility-title').html(r.destinationFacilityId ? `<strong>XPT:</strong> ${esc(r.destinationFacilityId)}` : '');
    $('#destination-facility-name').html(r.destinationFacilityName ? `<strong>DESTINO:</strong> ${esc(r.destinationFacilityName)}` : '');

    $('#extracted-total').text(r.totalInicial || r.ids.size);
    $('#verified-total').text(r.conferidos.size);

    this.atualizarListas();
  },

  atualizarProgresso() {
    const r = this.current;
    if (!r) return;

    const total = r.totalInicial || (r.ids.size || (r.conferidos.size + r.faltantes.size));
    const perc = total ? (r.conferidos.size / total) * 100 : 0;

    $('#progress-bar').css('width', perc + '%').text(Math.floor(perc) + '%');
  },

  atualizarListas() {
    const r = this.current;
    if (!r) return;

    $('#conferidos-list').html(
      `<h6>Conferidos (<span class='badge badge-success'>${r.conferidos.size}</span>)</h6>` +
      Array.from(r.conferidos).map(id => `<li class='list-group-item list-group-item-success'>${this.escHtml(id)}</li>`).join('')
    );

    $('#faltantes-list').html(
      `<h6>Faltantes (<span class='badge badge-danger'>${r.faltantes.size}</span>)</h6>` +
      Array.from(r.faltantes).map(id => `<li class='list-group-item list-group-item-danger'>${this.escHtml(id)}</li>`).join('')
    );

    $('#fora-rota-list').html(
      `<h6>Fora de Rota (<span class='badge badge-warning'>${r.foraDeRota.size}</span>)</h6>` +
      Array.from(r.foraDeRota).map(id => {
        const correct = this.findCorrectRouteForId(id);
        if (correct && String(correct) !== String(this.currentRouteId)) {
          const rr = this.routes.get(String(correct));
          const cl = rr && rr.cluster ? String(rr.cluster).trim() : '';
          const extra = cl
            ? ` <small class="text-muted">(CLUSTER ${this.escHtml(cl)})</small>`
            : ` <small class="text-muted">(cluster desconhecido)</small>`;
          return `<li class='list-group-item list-group-item-warning'>${this.escHtml(id)}${extra}</li>`;
        }
        return `<li class='list-group-item list-group-item-warning'>${this.escHtml(id)}</li>`;
      }).join('')
    );

    $('#duplicados-list').html(
      `<h6>Duplicados (<span class='badge badge-secondary'>${r.duplicados.size}</span>)</h6>` +
      Array.from(r.duplicados.entries())
        .map(([id, count]) => `<li class='list-group-item list-group-item-secondary'>${this.escHtml(id)} <span class="badge badge-dark ml-2">${this.dupCount(count)}x</span></li>`)
        .join('')
    );

    $('#verified-total').text(r.conferidos.size);
    this.atualizarProgresso();
  },

  renderCarretaUI() {
    const plateKey = this.carretas.currentPlateKey;
    const $cur = $('#carreta-current');
    const $list = $('#carreta-routes');
    const $sum = $('#carreta-summary');

    if (!plateKey) {
      $cur.html('<span class="text-muted">Nenhuma placa ativa</span>');
      $list.html('<li class="list-group-item text-muted">bipe uma placa para começar</li>');
      $sum.text('');
      this.renderCarretaProgress();
      return;
    }

    const p = this.carretas.plates.get(plateKey);
    if (!p) return;

    const meta = [];
    const esc = (x) => this.escHtml(x);
    meta.push(`<strong>${esc(p.license_plate)}</strong>`);
    if (p.vehicle_type_description) meta.push(`<span class="text-muted">(${esc(p.vehicle_type_description)})</span>`);
    if (p.carrier_name) meta.push(`<span class="text-muted">• ${esc(p.carrier_name)}</span>`);
    $cur.html(meta.join(' '));

    const routesArr = Array.from(p.routes);
    routesArr.sort((a, b) => String(a).localeCompare(String(b)));

    $list.html(
      routesArr.map(rk => `<li class="list-group-item">${this.escHtml(rk)}</li>`).join('') ||
      '<li class="list-group-item text-muted">sem rotas nessa placa</li>'
    );

    $sum.text(`${routesArr.length} rota(s) vinculada(s)`);
    this.renderCarretaProgress();
  },

  renderCarretaProgress() {
    const $label = $('#carreta-progress-label');
    const $pct = $('#carreta-progress-percent');
    const $bar = $('#carreta-progress-bar');
    const $missing = $('#carreta-missing-list');
    const $extra = $('#carreta-extra-list');

    if (!$label.length || !$pct.length || !$bar.length) return;

    const plateKey = this.carretas.currentPlateKey;

    const setUi = (done, total) => {
      const pctVal = total > 0 ? Math.round((done / total) * 100) : 0;
      $label.text(`${done}/${total}`);
      $pct.text(`${pctVal}%`);
      $bar.css('width', `${pctVal}%`);
      $bar.attr('aria-valuenow', String(pctVal));
    };

    const expected = Array.from(this.routes.keys()).map(String);
    expected.sort((a, b) => (Number(a) - Number(b)) || String(a).localeCompare(String(b)));

    if (!plateKey) {
      setUi(0, expected.length);
      $missing.html('<li class="list-group-item text-muted">bipe uma placa para ver o acompanhamento</li>');
      $extra.html('<li class="list-group-item text-muted">—</li>');
      return;
    }

    const plate = this.carretas.plates.get(plateKey);

    const linked = expected.filter(routeId => {
      const r = this.routes.get(String(routeId));
      return r && String(r.plateKey || '') === String(plateKey);
    });

    const missing = expected.filter(routeId => !linked.includes(routeId));

    const clustersImportados = new Set(
      Array.from(this.routes.values()).map(r => this.normalizeCluster(r.cluster)).filter(Boolean)
    );

    const plateQrRoutes = plate ? Array.from(plate.routes || []) : [];
    plateQrRoutes.sort((a, b) => String(a).localeCompare(String(b)));

    const extra = plateQrRoutes.filter(rk => {
      const m = String(rk).match(/^assignment:(.+)$/);
      const cl = this.normalizeCluster(m?.[1] || '');
      return cl && !clustersImportados.has(cl);
    });

    setUi(linked.length, expected.length);

    const fmtExpected = (routeId) => {
      const r = this.routes.get(String(routeId));
      const cl = r && r.cluster ? String(r.cluster).trim() : '';
      const fac = r && r.destinationFacilityName ? String(r.destinationFacilityName).trim() : '';
      const parts = [`${routeId}`];
      if (cl) parts.push(`CLUSTER ${cl}`);
      if (fac) parts.push(fac);
      return parts.join(' • ');
    };

    const fmtExtra = (rk) => {
      const m = String(rk).match(/^assignment:(.+)$/);
      return m ? m[1] : rk;
    };

    $missing.html(
      missing.map(routeId => `<li class="list-group-item">${this.escHtml(fmtExpected(routeId))}</li>`).join('') ||
      '<li class="list-group-item text-muted">nada faltando 🎉</li>'
    );

    $extra.html(
      extra.map(rk => `<li class="list-group-item">${this.escHtml(fmtExtra(rk))}</li>`).join('') ||
      '<li class="list-group-item text-muted">—</li>'
    );
  },

  renderPatioGeral() {
    // placeholder seguro caso não exista implementação específica
  },
});
