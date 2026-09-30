// Exportações CSV / XLSX.

Object.assign(ConferenciaApp, {
  getIdsForExportByTimestamp(r) {
    if (!r) return [];
    const set = new Set([
      ...Array.from(r.conferidos || []),
      ...Array.from(r.foraDeRota || []),
      ...Array.from((r.duplicados || new Map()).keys())
    ]);
    const ids = Array.from(set);

    ids.sort((a, b) => {
      const ta = r.timestamps?.get(a) ? Number(r.timestamps.get(a)) : 0;
      const tb = r.timestamps?.get(b) ? Number(r.timestamps.get(b)) : 0;
      return (ta - tb) || String(a).localeCompare(String(b));
    });
    return ids;
  },

  csvEscape(v) {
    const s = (v == null) ? '' : String(v);
    return '"' + s.replace(/"/g, '""') + '"';
  },

  buildScannerCsvHeader() {
    return '"date","time","time_zone","format","text","notes","favorite","date_utc","time_utc","metadata"';
  },

  buildScannerCsvRow(dt, format, text, metadata = '') {
    const d = (dt instanceof Date) ? dt : new Date(Number(dt || Date.now()));
    const pad2 = n => String(n).padStart(2, '0');

    const date = `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`;
    const time = `${pad2(d.getHours())}:${pad2(d.getMinutes())}:${pad2(d.getSeconds())}`;

    const iso = d.toISOString();
    const dateUtc = iso.slice(0, 10);
    const timeUtc = iso.split('T')[1].split('.')[0];

    const tzLabel = 'Horário Padrão do Amazonas';

    const esc = (v) => '"' + String(v ?? '').replace(/"/g, '""') + '"';

    return [
      esc(date),
      esc(time),
      esc(tzLabel),
      esc(format || 'QR Code'),
      esc(text || ''),
      esc(''),
      esc('0'),
      esc(dateUtc),
      esc(timeUtc),
      esc(metadata || '')
    ].join(',');
  },

  buildScannerCsvLinesForRoute(r) {
    const lines = [];
    if (!r) return lines;

    const ids = this.getIdsForExportByTimestamp(r);

    const firstIdTs = ids.length ? (r.timestamps?.get(ids[0]) || Date.now()) : Date.now();
    const plateTs = r.plateScanTs || firstIdTs;
    const routeTs = r.routeQrScanTs || (plateTs ? (Number(plateTs) + 1) : firstIdTs);

    let plateText = '';
    if (r.plateKey) {
      const p = this.carretas?.plates?.get(r.plateKey);
      if (p && p.jsonText) {
        plateText = String(p.jsonText);
      } else {
        plateText = JSON.stringify({
          id: (p && p.id) ? Number(p.id) : undefined,
          carrier_id: (p && p.carrier_id) ? Number(p.carrier_id) : undefined,
          carrier_name: (p && p.carrier_name) ? String(p.carrier_name) : undefined,
          license_plate: String((p && p.license_plate) || r.plateLicense || r.plateKey),
          vehicle_type_description: (p && p.vehicle_type_description) ? String(p.vehicle_type_description) : undefined,
          vehicle_type_id: (p && p.vehicle_type_id) ? Number(p.vehicle_type_id) : undefined,
          tracking_provider_ids: []
        }, (k, v) => (v === undefined ? undefined : v));
        plateText = plateText.replace(/,\s*"(?:id|carrier_id|carrier_name|vehicle_type_description|vehicle_type_id)"\s*:\s*null/g, '');
      }
    }

    let routeText = '';
    if (r.routeQrKey) {
      const routeJson = this.carretas?.routesJson?.get(r.routeQrKey);
      if (routeJson) {
        routeText = String(routeJson);
      } else if (r.routeQrRaw && String(r.routeQrRaw).trim().startsWith('{')) {
        routeText = String(r.routeQrRaw).trim();
      } else {
        routeText = JSON.stringify({
          container_id: (r.container_id != null) ? Number(r.container_id) : undefined,
          facility_id: (r.destinationFacilityId || ''),
          assignment: String(r.routeQrKey).replace(/^assignment:/, '')
        }, (k, v) => (v === undefined ? undefined : v));
      }
    }

    if (plateText) lines.push(this.buildScannerCsvRow(plateTs, 'QR Code', plateText, ''));
    if (routeText) lines.push(this.buildScannerCsvRow(routeTs, 'QR Code', routeText, ''));

    for (const id of ids) {
      const ts = r.timestamps?.get(id) || Date.now();
      const payload = JSON.stringify({ id: String(id), t: 'lm' });
      lines.push(this.buildScannerCsvRow(ts, 'QR Code', payload, ''));
    }

    return lines;
  },

  exportRotaAtualCsvComPlacaERota() {
    const r = this.current;
    if (!r) return alert('Nenhuma rota selecionada.');

    if (!r.plateKey || !r.routeQrKey) {
      return alert('Esta rota ainda não está vinculada a uma PLACA e a um QR de ROTA. Use a tela da CARRETA primeiro.');
    }

    const lines = [];
    lines.push(this.buildScannerCsvHeader());

    const body = this.buildScannerCsvLinesForRoute(r);
    if (!body.length) return alert('Nenhum registro para exportar.');

    lines.push(...body);

    const csv = lines.join('\r\n');
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);

    const cluster = (r.cluster || 'semCluster').replace(/[^\w\-]+/g, '_');
    link.download = `RECEBIMENTO_${this.workDay || this.todayLocalISO()}_${cluster}_ROTA_${r.routeId}_PLACA.csv`;
    link.click();
  },

  exportTodasRotasCsvComPlacaERota() {
    if (!this.routes || this.routes.size === 0) return alert('Não há rotas salvas para exportar.');

    const plateGroups = new Map();
    for (const r of this.routes.values()) {
      if (!r.plateKey || !r.routeQrKey) continue;
      if (!plateGroups.has(r.plateKey)) plateGroups.set(r.plateKey, []);
      plateGroups.get(r.plateKey).push(r);
    }

    if (!plateGroups.size) {
      return alert('Nenhuma rota está vinculada a PLACA/QR de rota. Use a tela da CARRETA primeiro.');
    }

    const lines = [];
    lines.push(this.buildScannerCsvHeader());

    const plateKeys = Array.from(plateGroups.keys()).sort((a, b) => String(a).localeCompare(String(b)));

    for (const plateKey of plateKeys) {
      const routesArr = plateGroups.get(plateKey) || [];
      routesArr.sort((a, b) => {
        const ca = String(a.cluster || '').localeCompare(String(b.cluster || ''));
        if (ca !== 0) return ca;
        return String(a.routeId || '').localeCompare(String(b.routeId || ''));
      });

      for (const r of routesArr) {
        const body = this.buildScannerCsvLinesForRoute(r);
        if (body.length) lines.push(...body);
      }
    }

    if (lines.length <= 1) return alert('Nenhum registro para exportar.');

    const csv = lines.join('\r\n');
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);

    const now = new Date();
    const stamp = `${now.getFullYear()}-${this.pad2(now.getMonth() + 1)}-${this.pad2(now.getDate())}_${this.pad2(now.getHours())}${this.pad2(now.getMinutes())}`;
    link.download = `RECEBIMENTO_${this.workDay || this.todayLocalISO()}_PLACAS_ROTAS_${stamp}.csv`;
    link.click();
  },

  exportMapaCarretasCsv() {
    if (!this.carretas.plates || this.carretas.plates.size === 0) {
      alert('Nenhuma placa/rota vinculada ainda.');
      return;
    }

    const header = 'plate,carrier,vehicle_type,route_qr_key,route_qr_raw';
    const linhas = [];

    for (const [plateKey, p] of this.carretas.plates.entries()) {
      const carrier = (p.carrier_name || '').replace(/,/g, ' ');
      const vt = (p.vehicle_type_description || '').replace(/,/g, ' ');
      for (const rk of Array.from(p.routes)) {
        const raw = (this.carretas.routesRaw.get(rk) || '').replace(/\r?\n/g, ' ');
        const rawEsc = `"${String(raw).replace(/"/g, '""')}"`;
        linhas.push(`${plateKey},${carrier},${vt},${rk},${rawEsc}`);
      }
    }

    const conteudo = [header, ...linhas].join('\r\n');
    const blob = new Blob([conteudo], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);

    const now = new Date();
    const stamp = `${now.getFullYear()}-${this.pad2(now.getMonth() + 1)}-${this.pad2(now.getDate())}_${this.pad2(now.getHours())}${this.pad2(now.getMinutes())}`;
    link.download = `mapa_carretas_${this.workDay || this.todayLocalISO()}_${stamp}.csv`;
    link.click();
  },

  exportRotaAtualCsvPadrao() {
    const r = this.current;
    if (!r) {
      alert('Nenhuma rota selecionada.');
      return;
    }

    const all = [
      ...Array.from(r.conferidos),
      ...Array.from(r.foraDeRota),
      ...Array.from(r.duplicados.keys())
    ];

    if (all.length === 0) {
      alert('Nenhum ID para exportar.');
      return;
    }

    const parseDateSafe = (value) => {
      if (!value) return new Date();
      if (value instanceof Date) return value;
      if (typeof value === 'number') return new Date(value);
      if (typeof value === 'string') {
        if (/^\d{4}-\d{2}-\d{2}T/.test(value)) {
          const d = new Date(value);
          if (!isNaN(d.getTime())) return d;
        }
        const m = value.match(/^(\d{2})\/(\d{2})\/(\d{4})[ T](\d{2}):(\d{2})(?::(\d{2}))?$/);
        if (m) {
          const [, dd, mm, yyyy, HH, MM, SS = '00'] = m;
          const iso = `${yyyy}-${mm}-${dd}T${HH}:${MM}:${SS}`;
          const d = new Date(iso);
          if (!isNaN(d.getTime())) return d;
        }
        if (/^\d{13}$/.test(value)) return new Date(Number(value));
        const d = new Date(value);
        if (!isNaN(d.getTime())) return d;
      }
      return new Date();
    };

    const zona = 'Horário Padrão de Brasília';
    const header = 'date,time,time_zone,format,text,notes,favorite,date_utc,time_utc,metadata,duplicates';

    const linhas = all.map(id => {
      const lidaEm = parseDateSafe(r.timestamps.get(id));
      const pad2 = n => String(n).padStart(2, '0');
      const date = `${lidaEm.getFullYear()}-${pad2(lidaEm.getMonth() + 1)}-${pad2(lidaEm.getDate())}`;
      const time = `${pad2(lidaEm.getHours())}:${pad2(lidaEm.getMinutes())}:${pad2(lidaEm.getSeconds())}`;

      const dateUtc = lidaEm.toISOString().slice(0, 10);
      const timeUtc = lidaEm.toISOString().split('T')[1].split('.')[0];
      const dupCount = r.duplicados.get(id) ? (Number(r.duplicados.get(id)) - 1) : 0;

      return `${date},${time},${zona},Code 128,${id},,0,${dateUtc},${timeUtc},,${dupCount}`;
    });

    const conteudo = [header, ...linhas].join('\r\n');
    const blob = new Blob([conteudo], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);

    const cluster = (r.cluster || 'semCluster').replace(/[^\w\-]+/g, '_');
    const rota = (r.routeId || 'semRota').replace(/[^\w\-]+/g, '_');

    link.download = `${cluster}_${rota}_padrao.csv`;
    link.click();
  },

  exportTodasRotasXlsx() {
    if (typeof XLSX === 'undefined') {
      alert('Biblioteca XLSX não carregou. Verifique o script do SheetJS no HTML.');
      return;
    }
    if (!this.routes || this.routes.size === 0) {
      alert('Não há rotas salvas para exportar.');
      return;
    }

    const routesSorted = Array.from(this.routes.values())
      .sort((a, b) => String(a.routeId).localeCompare(String(b.routeId)));

    const cols = routesSorted.map((r) => {
      const routeId = String(r.routeId || '');
      const cluster = String(r.cluster || '').trim();
      const header = cluster ? `${routeId}-${cluster}` : routeId;

      const ids = Array.from(r.conferidos || []);

      ids.sort((x, y) => {
        const tx = r.timestamps?.get(x) ? Number(r.timestamps.get(x)) : 0;
        const ty = r.timestamps?.get(y) ? Number(r.timestamps.get(y)) : 0;
        return (tx - ty) || String(x).localeCompare(String(y));
      });

      return { header, ids };
    });

    const maxLen = cols.reduce((m, c) => Math.max(m, c.ids.length), 0);

    const aoa = [];
    aoa.push(cols.map(c => c.header || 'ROTA'));
    for (let i = 0; i < maxLen; i++) {
      aoa.push(cols.map(c => c.ids[i] || ''));
    }

    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet(aoa);

    ws['!freeze'] = { xSplit: 0, ySplit: 1 };
    ws['!cols'] = cols.map(() => ({ wch: 18 }));

    XLSX.utils.book_append_sheet(wb, ws, 'Bipagens');

    const now = new Date();
    const stamp = `${now.getFullYear()}-${this.pad2(now.getMonth() + 1)}-${this.pad2(now.getDate())}_${this.pad2(now.getHours())}${this.pad2(now.getMinutes())}`;

    XLSX.writeFile(wb, `bipagens_todas_rotas_${this.workDay || this.todayLocalISO()}_${stamp}.xlsx`);
  }
});
