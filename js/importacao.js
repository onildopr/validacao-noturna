// Importação de rotas (HTML) e leitura de QR de placa/rota (carretas).

Object.assign(ConferenciaApp, {
  normalizeCaretKey(k) {
    const key = String(k || '').trim().toLowerCase();
    const noAcc = key.normalize('NFD').replace(/[\u0300-\u036f]/g, '');

    if (noAcc === 'assignment' || noAcc === 'assigment' || noAcc === 'asssignment') return 'assignment';
    if (noAcc === 'license_plate') return 'license_plate';
    if (noAcc === 'carrier_name') return 'carrier_name';
    if (noAcc === 'carrier_id') return 'carrier_id';
    if (noAcc === 'vehicle_type_description') return 'vehicle_type_description';
    if (noAcc === 'container_id') return 'container_id';
    if (noAcc === 'facility_id') return 'facility_id';
    if (noAcc === 'id') return 'id';

    return noAcc;
  },

  parseCaretKV(raw) {
    const cleaned = String(raw || '').replace(/\r/g, '').trim();
    const first = cleaned.split('\n')[0].trim();

    const kv = {};
    const tokens = first.split(',').map(t => t.trim()).filter(Boolean);

    for (const tok of tokens) {
      const m = tok.match(/^(\^?)([^\\^]+?)\^Ç\^?(.+?)\^?$/);
      if (!m) continue;

      const rawKey = m[2];
      let val = m[3];
      const key = this.normalizeCaretKey(rawKey);

      val = String(val)
        .replace(/^\^+|\^+$/g, '')
        .replace(/[{}]/g, '')
        .trim();

      kv[key] = val;
    }

    return kv;
  },

  parseScanPayload(raw) {
    const cleaned = String(raw || '').trim();
    if (!cleaned) return { kind: 'empty' };

    const firstLine = cleaned.split(/\?\n/)[0].trim();
    if (firstLine.includes('^') && firstLine.includes('Ç')) {
      const kv = this.parseCaretKV(firstLine);

      if (kv.license_plate) {
        const plateKey = String(kv.license_plate).trim().toUpperCase();
        const plateObj = {
          id: kv.id || '',
          carrier_id: kv.carrier_id || '',
          carrier_name: kv.carrier_name || '',
          license_plate: plateKey,
          vehicle_type_description: kv.vehicle_type_description || ''
        };
        const jsonText = JSON.stringify(plateObj);
        return {
          kind: 'plate',
          plateKey,
          plate: {
            raw: firstLine,
            jsonText,
            license_plate: plateKey,
            carrier_name: plateObj.carrier_name || '',
            vehicle_type_description: plateObj.vehicle_type_description || '',
            carrier_id: plateObj.carrier_id || '',
            id: plateObj.id || ''
          }
        };
      }

      if (kv.container_id || kv.assignment) {
        let assignment = kv.assignment ? String(kv.assignment).trim() : '';
        assignment = this.normalizeCluster(assignment);

        const obj = {
          container_id: kv.container_id ? Number(kv.container_id) : undefined,
          facility_id: kv.facility_id || '',
          assignment: assignment
        };
        Object.keys(obj).forEach(k => obj[k] === undefined && delete obj[k]);

        const routeKey = assignment
          ? `assignment:${assignment}`
          : (kv.container_id ? `container:${kv.container_id}` : `caret:${firstLine}`);

        const jsonText = JSON.stringify(obj);

        return {
          kind: 'routeqr',
          routeKey,
          routeIdCandidate: '',
          route: { raw: firstLine, obj, jsonText }
        };
      }
    }

    if (cleaned.startsWith('{') && cleaned.endsWith('}')) {
      try {
        const obj = JSON.parse(cleaned);

        if (obj && typeof obj === 'object' && obj.license_plate) {
          const plateKey = String(obj.license_plate || '').trim().toUpperCase();
          return {
            kind: 'plate',
            plateKey,
            plate: {
              raw: cleaned,
              license_plate: plateKey,
              carrier_name: obj.carrier_name || '',
              vehicle_type_description: obj.vehicle_type_description || '',
              carrier_id: obj.carrier_id || '',
              vehicle_type_id: obj.vehicle_type_id || '',
              id: obj.id || ''
            }
          };
        }

        if (obj && typeof obj === 'object' && (obj.container_id || obj.assignment || obj.routeId || obj.route_id)) {
          const candidate = obj.routeId || obj.route_id || obj.container_id || obj.assignment;
          const routeIdCandidate = candidate != null ? String(candidate) : '';
          const routeKey = (obj.container_id != null)
            ? `container:${obj.container_id}`
            : (obj.assignment != null)
              ? `assignment:${obj.assignment}`
              : (routeIdCandidate ? `route:${routeIdCandidate}` : `json:${cleaned}`);

          return {
            kind: 'routeqr',
            routeKey,
            routeIdCandidate,
            route: { raw: cleaned, obj }
          };
        }
      } catch (e) {}
    }

    const shipmentId = this.normalizarCodigo(cleaned);
    if (shipmentId) return { kind: 'shipment', shipmentId };

    const plateLike = cleaned.toUpperCase().replace(/[^A-Z0-9]/g, '');
    if (/^[A-Z]{3}\d[A-Z]\d{2}$/.test(plateLike) || /^[A-Z]{3}\d{4}$/.test(plateLike)) {
      return {
        kind: 'plate',
        plateKey: plateLike,
        plate: { raw: cleaned, license_plate: plateLike, carrier_name: '', vehicle_type_description: '' }
      };
    }

    const digits = cleaned.replace(/\D/g, '');
    if (digits.length >= 4) {
      const routeIdCandidate = digits;
      return {
        kind: 'routeqr',
        routeKey: `route:${routeIdCandidate}`,
        routeIdCandidate,
        route: { raw: cleaned, obj: null }
      };
    }

    return { kind: 'unknown' };
  },

  ensurePlate(plateInfo) {
    const now = Date.now();
    const key = String(plateInfo.license_plate || '').trim().toUpperCase();
    if (!key) return null;

    if (!this.carretas.plates.has(key)) {
      this.carretas.plates.set(key, {
        raw: plateInfo.raw || '',
        jsonText: plateInfo.jsonText || '',
        tsScan: now,
        license_plate: key,
        carrier_name: plateInfo.carrier_name || '',
        vehicle_type_description: plateInfo.vehicle_type_description || '',
        routes: new Set(),
        tsFirst: now,
        tsLast: now
      });
    } else {
      const p = this.carretas.plates.get(key);
      p.tsLast = now;
      if (plateInfo.raw) p.raw = plateInfo.raw;
      if (plateInfo.jsonText) p.jsonText = plateInfo.jsonText;
      p.tsScan = now;
      if (plateInfo.carrier_name) p.carrier_name = plateInfo.carrier_name;
      if (plateInfo.vehicle_type_description) p.vehicle_type_description = plateInfo.vehicle_type_description;
    }
    return key;
  },

  vincularRouteQrNaPlaca(routeKey, routeRaw, plateKey, routeIdCandidate = '') {
    if (!routeKey || !plateKey) return false;

    const plate = this.carretas.plates.get(plateKey);
    if (!plate) return false;

    plate.routes.add(routeKey);
    this.carretas.routeToPlate.set(routeKey, plateKey);

    const rawStr = (typeof routeRaw === 'string')
      ? routeRaw
      : ((routeRaw && routeRaw.raw) ? String(routeRaw.raw) : '');

    const jsonText = (routeRaw && typeof routeRaw === 'object' && routeRaw.jsonText)
      ? String(routeRaw.jsonText)
      : '';

    if (rawStr) this.carretas.routesRaw.set(routeKey, rawStr);
    if (jsonText) this.carretas.routesJson.set(routeKey, jsonText);
    this.carretas.routesTs.set(routeKey, Date.now());

    const assignMatch = String(routeKey).match(/^assignment:(.+)$/);
    const clusterCandidate = this.normalizeCluster(assignMatch?.[1] || '');

    let linked = 0;

    if (clusterCandidate) {
      for (const r of this.routes.values()) {
        const c = this.normalizeCluster(r.cluster);
        if (c && clusterCandidate && c === clusterCandidate) {
          r.plateKey = plateKey;
          r.plateRaw = plate.jsonText || plate.raw || '';
          r.plateLicense = plate.license_plate || plateKey;
          r.routeQrKey = routeKey;
          r.routeQrRaw = (typeof routeRaw === 'string' ? routeRaw : (routeRaw && routeRaw.raw) ? routeRaw.raw : '') || '';
          const _json = (routeRaw && routeRaw.jsonText) ? routeRaw.jsonText : '';
          if (_json) r.routeQrRaw = _json;
          r.plateScanTs = Number(plate.tsLast || Date.now());
          r.routeQrScanTs = Date.now();
          r.plateUpdatedAt = Date.now();
          linked++;
        }
      }
    }

    if (!linked) {
      const candidateIds = [];
      if (routeIdCandidate) candidateIds.push(String(routeIdCandidate));
      const m = String(routeKey).match(/^route:(.+)$/);
      if (m && m[1]) candidateIds.push(String(m[1]));

      for (const cid of candidateIds) {
        if (this.routes.has(String(cid))) {
          const r = this.routes.get(String(cid));
          r.plateKey = plateKey;
          r.plateRaw = plate.jsonText || plate.raw || '';
          r.plateLicense = plate.license_plate || plateKey;
          r.routeQrKey = routeKey;
          r.routeQrRaw = (typeof routeRaw === 'string' ? routeRaw : (routeRaw && routeRaw.raw) ? routeRaw.raw : '') || '';
          const _json = (routeRaw && routeRaw.jsonText) ? routeRaw.jsonText : '';
          if (_json) r.routeQrRaw = _json;
          r.plateScanTs = Number(plate.tsLast || Date.now());
          r.routeQrScanTs = Date.now();
          r.plateUpdatedAt = Date.now();
          linked++;
        }
      }
    }

    this.saveToStorage(this.workDay);
    this.markDirty('carreta');
    return true;
  },

  checkLinksForCurrentPlate() {
    const plateKey = this.carretas.currentPlateKey;
    if (!plateKey) return alert('Nenhuma placa ativa.');

    const p = this.carretas.plates.get(plateKey);
    if (!p) return alert('Placa ativa não encontrada na memória.');

    const routes = Array.from(p.routes || []);
    routes.sort((a, b) => String(a).localeCompare(String(b)));

    const clustersImportados = new Set(
      Array.from(this.routes.values()).map(r => this.normalizeCluster(r.cluster)).filter(Boolean)
    );

    const detalhes = routes.map(rk => {
      const m = String(rk).match(/^assignment:(.+)$/);
      const cl = this.normalizeCluster(m?.[1] || '');
      const ok = cl && clustersImportados.has(cl);
      return `- ${rk}  => cluster: ${cl || '(vazio)'}  ${ok ? '[OK]' : '[NÃO ENCONTRADO NAS ROTAS IMPORTADAS]'}`;
    });

    const msg =
      `PLACA ATIVA: ${plateKey}\n` +
      `ROTAS VINCULADAS: ${routes.length}\n\n` +
      (detalhes.length ? detalhes.join('\n') : '(nenhuma rota vinculada)\n') +
      `\n\nObs: [OK] significa que existe rota importada com cluster igual ao do QR.`;

    alert(msg);
  },

  clearBipagemForPlate(plateKeyRaw) {
    const plateKey = String(plateKeyRaw || '').trim().toUpperCase();
    if (!plateKey) return alert('Informe uma placa válida.');

    const p = this.carretas.plates.get(plateKey);
    if (!p) return alert('Essa placa não está carregada/vinculada.');

    const routeKeys = Array.from(p.routes || []);
    for (const rk of routeKeys) {
      this.carretas.routeToPlate.delete(rk);
      this.carretas.routesRaw.delete(rk);
      this.carretas.routesJson.delete(rk);
      this.carretas.routesTs.delete(rk);
    }

    p.routes = new Set();
    p.clearedAt = Date.now();

    for (const r of this.routes.values()) {
      if ((r.plateKey || '').toUpperCase() === plateKey) {
        r.plateKey = '';
        r.plateRaw = '';
        r.plateLicense = '';
        r.routeQrKey = '';
        r.routeQrRaw = '';
        r.plateScanTs = 0;
        r.routeQrScanTs = 0;
        r.plateUpdatedAt = Date.now();
      }
    }

    this.saveToStorage(this.workDay);
    this.markDirty('excluir bipagem placa');
    this.renderCarretaUI();
    this.renderPatioGeral();
    this.renderAcompanhamento();

    alert(`Bipagem/vínculos removidos para a placa ${plateKey}.`);
  },

  // Lê o HTML colado e extrai as rotas (não altera nada; usado pela importação e pelos testes)
  parseRoutesFromHtml(rawHtml) {
    const html = String(rawHtml || '').replace(/<[^>]+>/g, ' ');

    const idxs = [];
    for (const m of html.matchAll(/"routeId":(\d+)/g)) idxs.push(m.index);

    const porRota = new Map();
    for (let i = 0; i < idxs.length; i++) {
      const block = html.slice(idxs[i], i + 1 < idxs.length ? idxs[i + 1] : html.length);
      const routeId = String(/"routeId":(\d+)/.exec(block)[1]);

      const r = porRota.get(routeId) || { routeId, cluster: '', destinationFacilityId: '', destinationFacilityName: '', ids: new Set() };

      const clusterMatch = /"cluster":"([^"]+)"/.exec(block);
      if (clusterMatch && !r.cluster) r.cluster = this.normalizeCluster(clusterMatch[1]);

      const facMatch = /"destinationFacilityId":"([^"]+)","name":"([^"]+)"/.exec(block);
      if (facMatch && !r.destinationFacilityId) {
        r.destinationFacilityId = facMatch[1];
        r.destinationFacilityName = facMatch[2];
      }

      const regexId = /"id":\s*(\d{11})/g;
      let mId;
      while ((mId = regexId.exec(block)) !== null) {
        if (/^4\d{10}$/.test(mId[1])) r.ids.add(mId[1]);
      }

      porRota.set(routeId, r);
    }

    return Array.from(porRota.values());
  },

  // Importa rotas do HTML e devolve um relatório do que aconteceu (para avisar o usuário)
  importRoutesFromHtml(rawHtml) {
    const parsed = this.parseRoutesFromHtml(rawHtml);
    const report = {
      encontradas: parsed.length,
      importadas: 0,
      novas: [],            // routeIds criados agora
      atualizadas: [],      // {routeId, idsNovos} rotas que já existiam e ganharam IDs
      semIds: [],           // routeIds ignorados por não ter nenhum ID de pacote
      semCluster: [],       // routeIds importados sem cluster
      idsEmOutraRota: [],   // {id, routeId, outraRota} mesmo pacote listado em duas rotas
    };

    const donoAntes = new Map();
    for (const [rid, r] of this.routes.entries()) for (const id of r.ids) if (!donoAntes.has(id)) donoAntes.set(id, rid);

    for (const p of parsed) {
      if (!p.ids.size) {
        report.semIds.push(p.routeId);
        continue;
      }

      const routeId = p.routeId;
      // Só "ressuscita" uma rota excluída se o bloco realmente tiver pacotes
      let revivedAt = 0;
      if (this.deletedRoutes?.has(routeId)) {
        this.deletedRoutes.delete(routeId);
        if (!this.revivedRoutes) this.revivedRoutes = new Map();
        revivedAt = Date.now();
        this.revivedRoutes.set(routeId, revivedAt);
      }

      const existia = this.routes.has(routeId);
      const route = this.routes.get(routeId) || this.makeEmptyRoute(routeId);
      // Rota excluída e importada de novo começa zerada (bipagens antigas não contam)
      if (revivedAt) route.resetAt = revivedAt;

      if (p.cluster) route.cluster = p.cluster;
      if (p.destinationFacilityId) {
        route.destinationFacilityId = p.destinationFacilityId;
        route.destinationFacilityName = p.destinationFacilityName;
      }

      let idsNovos = 0;
      for (const id of p.ids) {
        const outra = donoAntes.get(id);
        if (outra && outra !== routeId) report.idsEmOutraRota.push({ id, routeId, outraRota: outra });
        if (!route.ids.has(id)) idsNovos++;
        route.ids.add(id);
        if (!route.conferidos.has(id)) route.faltantes.add(id);
        if (!donoAntes.has(id)) donoAntes.set(id, routeId);
      }

      route.totalInicial = route.ids.size;
      this.routes.set(routeId, route);
      report.importadas++;

      if (!existia) report.novas.push(routeId);
      else if (idsNovos) report.atualizadas.push({ routeId, idsNovos });
      if (!route.cluster) report.semCluster.push(routeId);
    }

    if (report.importadas) {
      this.lastRoutesSignature = '';
      this.saveToStorage(this.workDay);
      this.markDirty('import HTML');

      this.renderRoutesSelects();
      this.renderAcompanhamento();

      if (!this.currentRouteId || !this.routes.has(String(this.currentRouteId))) {
        this.currentRouteId = String(report.novas[0] || this.routes.keys().next().value);
      }
      this.refreshUIFromCurrent();
      this.atualizarListas();
    }

    return report;
  },

  // Texto do relatório de importação para mostrar ao usuário
  formatImportReport(rep) {
    const linhas = [];
    if (!rep.encontradas) return 'Não encontrei nenhum "routeId" no HTML. Confira se copiou a página inteira.';

    linhas.push(`Rotas encontradas no HTML: ${rep.encontradas}`);
    linhas.push(`Importadas: ${rep.importadas} (${rep.novas.length} nova(s), ${rep.atualizadas.length} atualizada(s))`);

    const lista = (arr, max = 10) => arr.slice(0, max).join(', ') + (arr.length > max ? ` e mais ${arr.length - max}` : '');

    if (rep.atualizadas.length) {
      linhas.push('', 'ℹ️ Rotas que já existiam e ganharam pacotes novos:');
      rep.atualizadas.slice(0, 10).forEach(a => linhas.push(`  • Rota ${a.routeId}: +${a.idsNovos} pacote(s)`));
    }
    if (rep.semIds.length) {
      linhas.push('', `⚠️ ${rep.semIds.length} rota(s) IGNORADA(S) por não ter nenhum ID de pacote (11 dígitos começando com 4):`);
      linhas.push(`  ${lista(rep.semIds)}`);
      linhas.push('  Pode ser HTML incompleto ou mudança no formato da página de origem.');
    }
    if (rep.semCluster.length) {
      linhas.push('', `⚠️ ${rep.semCluster.length} rota(s) sem CLUSTER: ${lista(rep.semCluster)}`);
    }
    if (rep.idsEmOutraRota.length) {
      linhas.push('', `⚠️ ${rep.idsEmOutraRota.length} pacote(s) aparecem em mais de uma rota (vale a primeira):`);
      rep.idsEmOutraRota.slice(0, 5).forEach(x => linhas.push(`  • ${x.id}: rota ${x.outraRota} e rota ${x.routeId}`));
    }
    return linhas.join('\n');
  },
});
