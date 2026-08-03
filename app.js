const DB_NAME = 'PolarComexDB', DB_VERSION = 1, STORE_NAME = 'operacoes';
const app = {
    db: null, records: [], filtered: [], showRealNames: false, companyName: 'POLAR', chart: null, competitorNames: new Map(),
    async init() { await this.initDB(); await this.loadRecords(); this.setupDragDrop(); if (this.records.length > 0) this.showDashboard(); },
    initDB() { return new Promise((resolve, reject) => { const req = indexedDB.open(DB_NAME, DB_VERSION); req.onerror = () => reject(req.error); req.onsuccess = () => { this.db = req.result; resolve(); }; req.onupgradeneeded = (e) => { const db = e.target.result; if (!db.objectStoreNames.contains(STORE_NAME)) { const store = db.createObjectStore(STORE_NAME, { keyPath: 'id', autoIncrement: true }); store.createIndex('dedupe', ['anoMes', 'importador', 'exportador', 'ncm', 'valorCifTotal'], { unique: true }); } }; }); },
    async loadRecords() { return new Promise((resolve) => { const tx = this.db.transaction(STORE_NAME, 'readonly'); const store = tx.objectStore(STORE_NAME); const req = store.getAll(); req.onsuccess = () => { this.records = req.result || []; const el = document.getElementById('recordCount'); if (el) el.textContent = this.records.length + ' registros carregados'; resolve(); }; }); },
    async saveRecord(rec) { return new Promise((resolve, reject) => { const tx = this.db.transaction(STORE_NAME, 'readwrite'); const store = tx.objectStore(STORE_NAME); const idx = store.index('dedupe'); const key = [rec.anoMes, rec.importador, rec.exportador, rec.ncm, rec.valorCifTotal]; const req = idx.get(key); req.onsuccess = () => { if (req.result) { rec.id = req.result.id; store.put(rec); } else { store.add(rec); } resolve(); }; req.onerror = () => reject(req.error); }); },
    async clearDatabase() { if (!confirm('Tem certeza que deseja apagar TODOS os dados importados?')) return; return new Promise((resolve) => { const tx = this.db.transaction(STORE_NAME, 'readwrite'); const store = tx.objectStore(STORE_NAME); store.clear(); tx.oncomplete = () => { this.records = []; this.filtered = []; location.reload(); resolve(); }; }); },
    setupDragDrop() { const zone = document.getElementById('dropZone'); zone.addEventListener('dragover', (e) => { e.preventDefault(); zone.classList.add('dragover'); }); zone.addEventListener('dragleave', () => zone.classList.remove('dragover')); zone.addEventListener('drop', (e) => { e.preventDefault(); zone.classList.remove('dragover'); this.handleFiles(e.dataTransfer.files); }); },
    async handleFiles(files) { const status = document.getElementById('uploadStatus'); const progress = document.getElementById('uploadProgress'); const fill = document.getElementById('uploadProgressFill'); progress.style.display = 'block'; let total = 0, imported = 0, errors = []; for (let i = 0; i < files.length; i++) { const file = files[i]; status.textContent = 'Processando ' + file.name + '... (' + (i + 1) + '/' + files.length + ')'; fill.style.width = ((i / files.length) * 100) + '%'; try { const data = await file.arrayBuffer(); const workbook = XLSX.read(data, { type: 'array', cellDates: true }); for (const sheetName of workbook.SheetNames) { const ws = workbook.Sheets[sheetName]; const json = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' }); if (json.length < 3) { errors.push(file.name + '/' + sheetName + ': poucas linhas'); continue; } const headerRow = json[1]; const dataRows = json.slice(2); const colMap = {}; headerRow.forEach((h, idx) => { if (h) colMap[String(h).trim()] = idx; }); const required = ['ANO/MÊS', 'PROVÁVEL EXPORTADOR', 'PROVÁVEL IMPORTADOR', 'Valor FOB Estimado Unitário']; const missing = required.filter(r => !(r in colMap)); if (missing.length > 0) { errors.push(file.name + '/' + sheetName + ': faltando ' + missing.join(', ')); continue; } for (const row of dataRows) { if (!row[colMap['ANO/MÊS']] && !row[colMap['PROVÁVEL IMPORTADOR']]) continue; const rec = this.parseRow(row, colMap, sheetName); if (rec) { await this.saveRecord(rec); imported++; } } total += dataRows.length; } } catch (e) { errors.push(file.name + ': ' + e.message); } } fill.style.width = '100%'; status.innerHTML = '<strong>' + imported + ' registros importados</strong> de ' + total + ' linhas lidas.' + (errors.length ? '<br><span style="color:var(--danger)">' + errors.length + ' erros: ' + errors.slice(0, 3).join('; ') + '</span>' : ''); await this.loadRecords(); if (this.records.length > 0) setTimeout(() => this.showDashboard(), 800); },
    parseRow(row, colMap, sheetName) {
        const get = (key) => { const idx = colMap[key]; if (idx === undefined) return ''; const val = row[idx]; if (val === undefined || val === null || val === '') return ''; return String(val).trim(); };
        const getNum = (key) => { const v = get(key); if (!v) return 0; const n = parseFloat(v.replace(/,/g, '').replace(/\s/g, '')); return isNaN(n) ? 0 : n; };
        const anoMes = get('ANO/MÊS'); if (!anoMes) return null;
        const importador = get('PROVÁVEL IMPORTADOR'); const exportador = get('PROVÁVEL EXPORTADOR'); if (!importador || !exportador) return null;
        const qtd = getNum('QTD Estatística') || getNum('Quantidade comercializada') || getNum('Peso líquido');
        const valorFobUnit = getNum('Valor FOB Estimado Unitário'); const valorFobTotal = getNum('VALOR FOB ESTIMADO TOTAL');
        const valorCifTotal = getNum('VALOR CIF TOTAL'); const freteUnit = getNum('Valor Frete Unitário'); const seguroUnit = getNum('Valor Seguro Unitário');
        const valorCifUnit = getNum('Valor CIF Unitario'); const freteTotal = getNum('Valor Frete total'); const seguroTotal = getNum('Valor Seguro total');
        return {
            anoMes, ano: anoMes.substring(0, 4), mes: anoMes.substring(4, 6),
            paisOrigem: get('PAIS DE ORIGEM'), ncm: get('NCM'),
            produto: get('Palavras chave de descrição do produto') || sheetName,
            exportador, importador, qtd, valorFobUnit, valorFobTotal, valorCifTotal, valorCifUnit,
            freteUnit, seguroUnit, freteTotal, seguroTotal,
            modal: get('MODAL'), cidadeImportador: get('CIDADE DO IMPORTADOR'), ufImportador: get('UF IMPORTADOR'),
            urf: get('URF de Entrada'), codCgce: get('Cod CGCE'), codCnae: get('Cod CNAE Importador'),
            qtdOperacoes: getNum('Qtd. de operações estimada'), notificado: get('Notificado'), sheetName
        };
    },
    showDashboard() { document.getElementById('uploadScreen').style.display = 'none'; document.getElementById('dashboardScreen').style.display = 'block'; this.populateFilters(); this.applyFilters(); },
    updateCompanyFilter() { this.companyName = document.getElementById('companyFilter').value.trim() || 'POLAR'; this.refreshAll(); },
    isPolar(importador) { if (!importador) return false; return importador.toUpperCase().includes(this.companyName.toUpperCase()); },
    getAnonName(realName) { if (this.showRealNames) return realName; if (!this.competitorNames.has(realName)) { const idx = this.competitorNames.size; this.competitorNames.set(realName, 'Empresa ' + String.fromCharCode(65 + (idx % 26)) + (idx >= 26 ? Math.floor(idx / 26) : '')); } return this.competitorNames.get(realName); },
    toggleNames() { this.showRealNames = !this.showRealNames; document.getElementById('toggleNames').classList.toggle('active', this.showRealNames); this.refreshAll(); },
    populateFilters() { const anos = [...new Set(this.records.map(r => r.ano))].sort(); const meses = [...new Set(this.records.map(r => r.mes))].sort(); const fornecedores = [...new Set(this.records.map(r => r.exportador))].sort(); const ncms = [...new Set(this.records.map(r => r.ncm).filter(Boolean))].sort(); const paises = [...new Set(this.records.map(r => r.paisOrigem).filter(Boolean))].sort(); const fill = (id, vals) => { const sel = document.getElementById(id); const cur = sel.value; sel.innerHTML = '<option value="">Todos</option>' + vals.map(v => '<option value="' + v + '">' + v + '</option>').join(''); sel.value = cur; }; fill('filterAno', anos); fill('filterMesIni', meses); fill('filterMesFim', meses); fill('filterFornecedor', fornecedores); fill('filterNCM', ncms); fill('filterPais', paises); },
    applyFilters() { const ano = document.getElementById('filterAno').value; const mesIni = document.getElementById('filterMesIni').value; const mesFim = document.getElementById('filterMesFim').value; const forn = document.getElementById('filterFornecedor').value; const ncm = document.getElementById('filterNCM').value; const pais = document.getElementById('filterPais').value; this.filtered = this.records.filter(r => { if (ano && r.ano !== ano) return false; if (mesIni && parseInt(r.mes) < parseInt(mesIni)) return false; if (mesFim && parseInt(r.mes) > parseInt(mesFim)) return false; if (forn && r.exportador !== forn) return false; if (ncm && r.ncm !== ncm) return false; if (pais && r.paisOrigem !== pais) return false; return true; }); this.refreshAll(); },
    refreshAll() { this.renderResumo(); this.renderCompras(); this.renderRanking(); this.renderEconomia(); this.renderChart(); },
    renderResumo() { const polar = this.filtered.filter(r => this.isPolar(r.importador)); const totalFob = polar.reduce((s, r) => s + r.valorFobTotal, 0); const totalQtd = polar.reduce((s, r) => s + r.qtd, 0); const totalOps = polar.reduce((s, r) => s + r.qtdOperacoes, 0); const fornecedores = new Set(polar.map(r => r.exportador)).size; const mediaKg = totalQtd > 0 ? totalFob / totalQtd : 0; const mediaTon = mediaKg * 1000; const freteTotal = polar.reduce((s, r) => s + (r.freteUnit * r.qtd), 0); const freteKg = totalQtd > 0 ? freteTotal / totalQtd : 0; const seguroTotal = polar.reduce((s, r) => s + (r.seguroUnit * r.qtd), 0); const seguroKg = totalQtd > 0 ? seguroTotal / totalQtd : 0; const cards = [{ label: 'Valor Total FOB', value: this.fmtMoney(totalFob), sub: 'US$ no período', cls: '' }, { label: 'Quantidade Total', value: this.fmtNum(totalQtd) + ' kg', sub: this.fmtNum(totalQtd / 1000) + ' ton', cls: '' }, { label: 'Preço Médio', value: this.fmtMoney(mediaKg) + '/kg', sub: this.fmtMoney(mediaTon) + '/ton', cls: '' }, { label: 'Frete Médio', value: this.fmtMoney(freteKg) + '/kg', sub: 'US$/kg', cls: '' }, { label: 'Seguro Médio', value: this.fmtMoney(seguroKg) + '/kg', sub: 'US$/kg', cls: '' }, { label: 'Fornecedores', value: fornecedores, sub: 'distintos', cls: '' }, { label: 'Operações', value: totalOps, sub: 'embarques estimados', cls: '' }]; document.getElementById('resumoCards').innerHTML = cards.map(c => '<div class="card ' + c.cls + '"><div class="card-label">' + c.label + '</div><div class="card-value">' + c.value + '</div><div class="card-sub">' + c.sub + '</div></div>').join(''); },

    renderCompras() {
        const polar = this.filtered.filter(r => this.isPolar(r.importador));
        const tbody = document.getElementById('tbodyCompras');
        const tfoot = document.getElementById('tfootCompras');
        if (polar.length === 0) { tbody.innerHTML = '<tr><td colspan="13" class="text-center" style="padding:32px;color:var(--text-light)">Nenhuma compra da POLAR no período filtrado</td></tr>'; tfoot.innerHTML = ''; return; }
        // Agrupar por anoMes+exportador+ncm para deduplicar visualmente
        const grouped = {};
        for (const r of polar) {
            const key = r.anoMes + '|' + r.exportador + '|' + r.ncm;
            if (!grouped[key]) {
                grouped[key] = { records: [], qtd: 0, valorFobTotal: 0, valorCifTotal: 0, freteTotal: 0, seguroTotal: 0 };
                // guardar o primeiro registro para dados completos no modal
                grouped[key].firstRecord = r;
            }
            grouped[key].records.push(r);
            grouped[key].qtd += r.qtd;
            grouped[key].valorFobTotal += r.valorFobTotal;
            grouped[key].valorCifTotal += r.valorCifTotal;
            grouped[key].freteTotal += (r.freteUnit * r.qtd);
            grouped[key].seguroTotal += (r.seguroUnit * r.qtd);
        }
        const rows = Object.values(grouped).sort((a, b) => b.firstRecord.anoMes.localeCompare(a.firstRecord.anoMes));

        let totalQtd = 0, totalFob = 0, totalCif = 0, totalFrete = 0, totalSeguro = 0;

        tbody.innerHTML = rows.map((g) => {
            const r = g.firstRecord;
            const precoKg = g.qtd > 0 ? g.valorFobTotal / g.qtd : 0;
            const precoTon = precoKg * 1000;
            totalQtd += g.qtd; totalFob += g.valorFobTotal; totalCif += g.valorCifTotal; totalFrete += g.freteTotal; totalSeguro += g.seguroTotal;

            // Dados do exportador no mês (todas as vendas dele)
            const vendasExp = this.filtered.filter(x => x.anoMes === r.anoMes && x.exportador === r.exportador);
            const volTotalExp = vendasExp.reduce((s, x) => s + x.qtd, 0);

            // Melhor preço que o exportador vendeu no mês
            const precosExp = vendasExp.map(x => x.qtd > 0 ? x.valorFobTotal / x.qtd : Infinity).filter(p => p > 0 && p !== Infinity);
            const melhorPrecoExp = precosExp.length > 0 ? Math.min(...precosExp) : null;
            const melhorComprador = melhorPrecoExp !== null ? vendasExp.find(x => x.qtd > 0 && Math.abs((x.valorFobTotal / x.qtd) - melhorPrecoExp) < 0.0001) : null;
            const nomeMelhor = melhorComprador ? (this.isPolar(melhorComprador.importador) ? '<strong style="color:var(--accent)">POLAR</strong>' : this.getAnonName(melhorComprador.importador)) : '—';

            // Comparação com média do mesmo fornecedor/mês
            const comp = this.getComparacao(r.anoMes, r.exportador, r.ncm, precoKg);
            const diff = comp.diffPct;
            const signal = diff <= -2 ? '🟢' : diff >= 2 ? '🔴' : '🟡';
            const classif = this.classificar(diff);
            const diffStr = comp.temComparacao ? (diff > 0 ? '+' : '') + diff.toFixed(1) + '%' : '—';

            // Diferença vs melhor preço do exportador
            const diffMelhor = melhorPrecoExp !== null && melhorPrecoExp > 0 ? ((precoKg - melhorPrecoExp) / melhorPrecoExp) * 100 : null;

            return '<tr onclick="app.openModalCompra(' + this.escJS(r.anoMes) + ',' + this.escJS(r.exportador) + ',' + this.escJS(r.ncm) + ')">' +
                '<td>' + r.anoMes.substring(4, 6) + '/' + r.anoMes.substring(0, 4) + '</td>' +
                '<td>' + (r.ncm || r.produto) + '</td>' +
                '<td>' + r.exportador + '</td>' +
                '<td class="text-right font-mono">' + this.fmtNum(g.qtd) + '</td>' +
                '<td class="text-right font-mono">' + this.fmtMoney(g.valorFobTotal) + '</td>' +
                '<td class="text-right font-mono">' + this.fmtMoney(precoKg) + '</td>' +
                '<td class="text-right font-mono">' + this.fmtMoney(precoTon) + '</td>' +
                '<td class="text-right font-mono">' + (melhorPrecoExp !== null ? this.fmtMoney(melhorPrecoExp) : '—') + '</td>' +
                '<td class="text-center">' + nomeMelhor + '</td>' +
                '<td class="text-right font-mono">' + (diffMelhor !== null ? (diffMelhor > 0 ? '+' : '') + diffMelhor.toFixed(1) + '%' : '—') + '</td>' +
                '<td class="text-right font-mono">' + this.fmtNum(volTotalExp) + '</td>' +
                '<td class="text-center signal ' + (diff <= -2 ? 'signal-green' : diff >= 2 ? 'signal-red' : 'signal-yellow') + '">' + signal + ' ' + diffStr + '</td>' +
                '<td class="text-center"><span class="badge badge-' + classif.cls + '">' + classif.label + '</span></td>' +
                '</tr>';
        }).join('');

        // Totais no rodapé
        const mediaGeral = totalQtd > 0 ? totalFob / totalQtd : 0;
        tfoot.innerHTML = '<tr>' +
            '<td colspan="3" style="text-align:right"><strong>TOTAIS</strong></td>' +
            '<td class="text-right font-mono"><strong>' + this.fmtNum(totalQtd) + '</strong></td>' +
            '<td class="text-right font-mono"><strong>' + this.fmtMoney(totalFob) + '</strong></td>' +
            '<td class="text-right font-mono"><strong>' + this.fmtMoney(mediaGeral) + '</strong></td>' +
            '<td class="text-right font-mono"><strong>' + this.fmtMoney(mediaGeral * 1000) + '</strong></td>' +
            '<td colspan="4"></td>' +
            '<td class="text-right font-mono"><strong>FOB: ' + this.fmtMoney(totalFob) + '<br>CIF: ' + this.fmtMoney(totalCif) + '<br>Frete: ' + this.fmtMoney(totalFrete) + '<br>Seguro: ' + this.fmtMoney(totalSeguro) + '</strong></td>' +
            '<td colspan="2"></td>' +
            '</tr>';
    },

    renderRanking() {
        const polar = this.filtered.filter(r => this.isPolar(r.importador));
        const fornecedores = [...new Set(polar.map(r => r.exportador))];
        const tbody = document.getElementById('tbodyRanking');
        if (fornecedores.length === 0) { tbody.innerHTML = '<tr><td colspan="7" class="text-center" style="padding:32px;color:var(--text-light)">Sem dados</td></tr>'; return; }
        const rows = fornecedores.map(forn => {
            const polarCompras = polar.filter(r => r.exportador === forn);
            const polarQtd = polarCompras.reduce((s, r) => s + r.qtd, 0);
            const polarFob = polarCompras.reduce((s, r) => s + r.valorFobTotal, 0);
            const polarPreco = polarQtd > 0 ? polarFob / polarQtd : 0;
            const todos = this.filtered.filter(r => r.exportador === forn);
            const compradores = [...new Set(todos.map(r => r.importador))];
            const precos = compradores.map(imp => { const ops = todos.filter(r => r.importador === imp); const q = ops.reduce((s, r) => s + r.qtd, 0); const v = ops.reduce((s, r) => s + r.valorFobTotal, 0); return { importador: imp, preco: q > 0 ? v / q : Infinity, qtd: q }; }).filter(p => p.preco > 0 && p.preco !== Infinity);
            precos.sort((a, b) => a.preco - b.preco);
            const pos = precos.findIndex(p => this.isPolar(p.importador)) + 1;
            const total = precos.length;
            const melhor = total > 0 ? precos[0].preco : 0;
            const demais = precos.filter(p => !this.isPolar(p.importador));
            const mediaDemais = demais.length > 0 ? demais.reduce((s, p) => s + p.preco * p.qtd, 0) / demais.reduce((s, p) => s + p.qtd, 0) : 0;
            const diff = mediaDemais > 0 ? ((polarPreco - mediaDemais) / mediaDemais) * 100 : 0;
            const classif = this.classificar(diff);
            return { forn, pos, total, polarPreco, melhor, mediaDemais, classif };
        }).sort((a, b) => (a.pos / a.total) - (b.pos / b.total));
        tbody.innerHTML = rows.map(r => '<tr>' +
            '<td>' + r.forn + '</td>' +
            '<td class="text-center font-mono" style="font-weight:700;color:' + (r.pos === 1 ? 'var(--success)' : r.pos === r.total ? 'var(--danger)' : 'var(--text)') + '">' + r.pos + 'º</td>' +
            '<td class="text-center">de ' + r.total + '</td>' +
            '<td class="text-right font-mono">' + this.fmtMoney(r.polarPreco) + '</td>' +
            '<td class="text-right font-mono">' + this.fmtMoney(r.melhor) + '</td>' +
            '<td class="text-right font-mono">' + this.fmtMoney(r.mediaDemais) + '</td>' +
            '<td class="text-center"><span class="badge badge-' + r.classif.cls + '">' + r.classif.label + '</span></td>' +
            '</tr>').join('');
    },

    renderEconomia() {
        const polar = this.filtered.filter(r => this.isPolar(r.importador));
        const tbody = document.getElementById('tbodyEconomia');
        const tfoot = document.getElementById('tfootEconomia');
        if (polar.length === 0) { tbody.innerHTML = '<tr><td colspan="6" class="text-center" style="padding:32px;color:var(--text-light)">Sem dados</td></tr>'; tfoot.innerHTML = ''; return; }
        const grouped = {};
        for (const r of polar) { const key = r.anoMes + '|' + r.exportador; if (!grouped[key]) grouped[key] = { anoMes: r.anoMes, forn: r.exportador, qtd: 0, valorFobTotal: 0 }; grouped[key].qtd += r.qtd; grouped[key].valorFobTotal += r.valorFobTotal; }
        let totalEco = 0;
        const rows = Object.values(grouped).sort((a, b) => b.anoMes.localeCompare(a.anoMes));
        tbody.innerHTML = rows.map(r => {
            const precoPago = r.qtd > 0 ? r.valorFobTotal / r.qtd : 0;
            const mercado = this.filtered.filter(x => x.anoMes === r.anoMes && x.exportador === r.forn && !this.isPolar(x.importador));
            let melhorPreco = Infinity;
            for (const m of mercado) { if (m.qtd > 0) { const p = m.valorFobTotal / m.qtd; if (p < melhorPreco) melhorPreco = p; } }
            let economia = 0;
            if (melhorPreco !== Infinity && precoPago > melhorPreco) economia = (precoPago - melhorPreco) * r.qtd;
            totalEco += economia;
            const ecoStr = economia > 0 ? this.fmtMoney(economia) : '<span style="color:var(--success);font-weight:600">Melhor preço ✓</span>';
            return '<tr>' +
                '<td>' + r.anoMes.substring(4, 6) + '/' + r.anoMes.substring(0, 4) + '</td>' +
                '<td>' + r.forn + '</td>' +
                '<td class="text-right font-mono">' + this.fmtNum(r.qtd) + '</td>' +
                '<td class="text-right font-mono">' + this.fmtMoney(precoPago) + '</td>' +
                '<td class="text-right font-mono">' + (melhorPreco !== Infinity ? this.fmtMoney(melhorPreco) : '—') + '</td>' +
                '<td class="text-right font-mono">' + ecoStr + '</td>' +
                '</tr>';
        }).join('');
        tfoot.innerHTML = '<tr><td colspan="5" style="text-align:right;padding:12px 16px">Economia Potencial Total:</td><td class="text-right font-mono" style="color:var(--danger);font-size:1.1rem">' + this.fmtMoney(totalEco) + '</td></tr>';
    },

    renderChart() {
        const ctx = document.getElementById('chartEvolucao').getContext('2d');
        const meses = [...new Set(this.filtered.map(r => r.anoMes))].sort();
        const polarPrecos = []; const minPrecos = []; const avgPrecos = []; const maxPrecos = []; const labels = [];
        for (const am of meses) {
            labels.push(am.substring(4, 6) + '/' + am.substring(0, 4));
            const polarOps = this.filtered.filter(r => r.anoMes === am && this.isPolar(r.importador));
            const polarQtd = polarOps.reduce((s, r) => s + r.qtd, 0); const polarFob = polarOps.reduce((s, r) => s + r.valorFobTotal, 0);
            polarPrecos.push(polarQtd > 0 ? polarFob / polarQtd : null);
            const mercadoOps = this.filtered.filter(r => r.anoMes === am && !this.isPolar(r.importador));
            if (mercadoOps.length === 0) { minPrecos.push(null); avgPrecos.push(null); maxPrecos.push(null); continue; }
            const precos = mercadoOps.map(r => r.qtd > 0 ? r.valorFobTotal / r.qtd : null).filter(Boolean);
            if (precos.length === 0) { minPrecos.push(null); avgPrecos.push(null); maxPrecos.push(null); continue; }
            minPrecos.push(Math.min(...precos)); maxPrecos.push(Math.max(...precos));
            const totalQtd = mercadoOps.reduce((s, r) => s + r.qtd, 0); const totalFob = mercadoOps.reduce((s, r) => s + r.valorFobTotal, 0);
            avgPrecos.push(totalQtd > 0 ? totalFob / totalQtd : null);
        }
        if (this.chart) this.chart.destroy();
        this.chart = new Chart(ctx, {
            type: 'line',
            data: {
                labels, datasets: [
                    { label: 'POLAR', data: polarPrecos, borderColor: '#3182ce', backgroundColor: '#3182ce', tension: .3, fill: false, pointRadius: 4 },
                    { label: 'Menor Preço Mercado', data: minPrecos, borderColor: '#38a169', backgroundColor: '#38a169', tension: .3, fill: false, borderDash: [5, 5], pointRadius: 3 },
                    { label: 'Média Mercado', data: avgPrecos, borderColor: '#718096', backgroundColor: '#718096', tension: .3, fill: false, pointRadius: 3 },
                    { label: 'Maior Preço Mercado', data: maxPrecos, borderColor: '#e53e3e', backgroundColor: '#e53e3e', tension: .3, fill: false, borderDash: [5, 5], pointRadius: 3 }
                ]
            },
            options: { responsive: true, maintainAspectRatio: false, interaction: { mode: 'index', intersect: false }, plugins: { legend: { position: 'top' }, tooltip: { callbacks: { label: (ctx) => ctx.dataset.label + ': US$ ' + (ctx.parsed.y !== null ? ctx.parsed.y.toFixed(3) : '—') + '/kg' } } }, scales: { y: { beginAtZero: false, title: { display: true, text: 'US$/kg' } } } }
        });
    },

    getComparacao(anoMes, exportador, ncm, polarPreco) {
        const demais = this.filtered.filter(r => r.anoMes === anoMes && r.exportador === exportador && !this.isPolar(r.importador));
        if (demais.length === 0) {
            const demaisNcm = this.filtered.filter(r => r.anoMes === anoMes && r.ncm === ncm && !this.isPolar(r.importador));
            if (demaisNcm.length === 0) return { temComparacao: false, diffPct: 0, posicao: 1, total: 1 };
            return this.calcComparacao(demaisNcm, polarPreco);
        }
        return this.calcComparacao(demais, polarPreco);
    },

    calcComparacao(demais, polarPreco) {
        const precos = demais.map(r => ({ preco: r.qtd > 0 ? r.valorFobTotal / r.qtd : 0, qtd: r.qtd })).filter(p => p.preco > 0);
        const todos = [...precos, { preco: polarPreco, qtd: 1, isPolar: true }];
        todos.sort((a, b) => a.preco - b.preco);
        const pos = todos.findIndex(p => p.isPolar) + 1;
        const total = todos.length;
        const menor = todos[0].preco;
        const totalQtd = precos.reduce((s, p) => s + p.qtd, 0);
        const totalVal = precos.reduce((s, p) => s + p.preco * p.qtd, 0);
        const media = totalQtd > 0 ? totalVal / totalQtd : 0;
        const diffPct = media > 0 ? ((polarPreco - media) / media) * 100 : 0;
        return { temComparacao: true, diffPct, posicao: pos, total, menor, media };
    },

    classificar(diff) {
        if (diff <= -10) return { label: 'Excelente compra', cls: 'excelente' };
        if (diff <= -2) return { label: 'Boa compra', cls: 'boa' };
        if (diff < 2) return { label: 'Na média', cls: 'media' };
        if (diff <= 10) return { label: 'Compra cara', cls: 'cara' };
        return { label: 'Muito acima do mercado', cls: 'muito-cara' };
    },

    // ===== MODAL COMPRA COMPLETA =====
    openModalCompra(anoMes, exportador, ncm) {
        const modal = document.getElementById('modalCompra');
        const subtitle = document.getElementById('modalCompraSubtitle');
        const dadosDiv = document.getElementById('modalCompraDados');
        const summaryDiv = document.getElementById('modalCompraSummary');
        const rankingBody = document.getElementById('modalCompraRanking');
        const mercadoDiv = document.getElementById('modalCompraMercado');

        // Pegar todos os registros da POLAR para esse grupo
        const polarRecs = this.filtered.filter(r => r.anoMes === anoMes && r.exportador === exportador && r.ncm === ncm && this.isPolar(r.importador));
        if (polarRecs.length === 0) return;
        const rec = polarRecs[0]; // usar o primeiro para dados completos

        const polarQtd = polarRecs.reduce((s, r) => s + r.qtd, 0);
        const polarFob = polarRecs.reduce((s, r) => s + r.valorFobTotal, 0);
        const polarCif = polarRecs.reduce((s, r) => s + r.valorCifTotal, 0);
        const polarFrete = polarRecs.reduce((s, r) => s + (r.freteUnit * r.qtd), 0);
        const polarSeguro = polarRecs.reduce((s, r) => s + (r.seguroUnit * r.qtd), 0);
        const polarPreco = polarQtd > 0 ? polarFob / polarQtd : 0;

        subtitle.textContent = anoMes.substring(4, 6) + '/' + anoMes.substring(0, 4) + ' — ' + exportador + ' — ' + (ncm || rec.produto);

        // Dados completos da compra
        const campos = [
            { label: 'ANO/MÊS', value: rec.anoMes },
            { label: 'País de Origem', value: rec.paisOrigem || '—' },
            { label: 'NCM', value: rec.ncm || '—' },
            { label: 'Produto', value: rec.produto || '—' },
            { label: 'Provável Exportador', value: rec.exportador },
            { label: 'Provável Importador', value: rec.importador },
            { label: 'Modal', value: rec.modal || '—' },
            { label: 'Cidade do Importador', value: rec.cidadeImportador || '—' },
            { label: 'UF Importador', value: rec.ufImportador || '—' },
            { label: 'URF de Entrada', value: rec.urf || '—' },
            { label: 'Código CGCE', value: rec.codCgce || '—' },
            { label: 'Código CNAE', value: rec.codCnae || '—' },
            { label: 'Quantidade (kg)', value: this.fmtNum(polarQtd), cls: 'num' },
            { label: 'Qtd. Operações Estimada', value: rec.qtdOperacoes || '—', cls: 'num' },
            { label: 'Valor FOB Total (US$)', value: this.fmtMoney(polarFob), cls: 'money' },
            { label: 'Valor FOB Unitário (US$/kg)', value: this.fmtMoney(polarPreco), cls: 'money' },
            { label: 'Valor CIF Total (US$)', value: this.fmtMoney(polarCif), cls: 'money' },
            { label: 'Valor CIF Unitário (US$/kg)', value: this.fmtMoney(rec.valorCifUnit), cls: 'money' },
            { label: 'Frete Total (US$)', value: this.fmtMoney(polarFrete), cls: 'money' },
            { label: 'Frete Unitário (US$/kg)', value: this.fmtMoney(rec.freteUnit), cls: 'money' },
            { label: 'Seguro Total (US$)', value: this.fmtMoney(polarSeguro), cls: 'money' },
            { label: 'Seguro Unitário (US$/kg)', value: this.fmtMoney(rec.seguroUnit), cls: 'money' },
        ];

        dadosDiv.innerHTML = campos.map(c => '<div class="data-item"><div class="data-item-label">' + c.label + '</div><div class="data-item-value ' + (c.cls || '') + '">' + c.value + '</div></div>').join('');

        // ===== COMPARAÇÃO MESMO FORNECEDOR/MÊS =====
        const todos = this.filtered.filter(r => r.anoMes === anoMes && r.exportador === exportador);
        const compradores = [...new Set(todos.map(r => r.importador))];
        let rows = compradores.map(imp => {
            const ops = todos.filter(r => r.importador === imp);
            const q = ops.reduce((s, r) => s + r.qtd, 0);
            const v = ops.reduce((s, r) => s + r.valorFobTotal, 0);
            return { importador: imp, preco: q > 0 ? v / q : 0, qtd: q, fob: v, isPolar: this.isPolar(imp) };
        }).filter(r => r.preco > 0);

        if (rows.length === 0) {
            const todosNcm = this.filtered.filter(r => r.anoMes === anoMes && r.ncm === ncm);
            const compradoresNcm = [...new Set(todosNcm.map(r => r.importador))];
            rows = compradoresNcm.map(imp => {
                const ops = todosNcm.filter(r => r.importador === imp);
                const q = ops.reduce((s, r) => s + r.qtd, 0);
                const v = ops.reduce((s, r) => s + r.valorFobTotal, 0);
                return { importador: imp, preco: q > 0 ? v / q : 0, qtd: q, fob: v, isPolar: this.isPolar(imp) };
            }).filter(r => r.preco > 0);
        }

        rows.sort((a, b) => a.preco - b.preco);
        const pos = rows.findIndex(r => r.isPolar) + 1;
        const total = rows.length;
        const menor = rows[0] ? rows[0].preco : 0;
        const demais = rows.filter(r => !r.isPolar);
        const mediaDemais = demais.length > 0 ? demais.reduce((s, r) => s + r.preco * r.qtd, 0) / demais.reduce((s, r) => s + r.qtd, 0) : 0;
        const diffMenor = menor > 0 ? ((polarPreco - menor) / menor) * 100 : 0;
        const diffMedia = mediaDemais > 0 ? ((polarPreco - mediaDemais) / mediaDemais) * 100 : 0;

        summaryDiv.innerHTML =
            '<p><strong>Preço POLAR:</strong> US$ ' + this.fmtMoney(polarPreco) + '/kg</p>' +
            '<p><strong>Posição no ranking:</strong> ' + pos + 'ª melhor de ' + total + ' compradores</p>' +
            '<p><strong>Menor preço do grupo:</strong> US$ ' + this.fmtMoney(menor) + '/kg (' + (diffMenor > 0 ? '+' : '') + diffMenor.toFixed(1) + '% vs POLAR)</p>' +
            '<p><strong>Média dos demais compradores:</strong> US$ ' + this.fmtMoney(mediaDemais) + '/kg (' + (diffMedia > 0 ? '+' : '') + diffMedia.toFixed(1) + '% vs POLAR)</p>' +
            '<p><strong>Veredito:</strong> ' + (diffMedia <= -2 ? '✅ Sua compra foi MELHOR que a média do mercado' : diffMedia >= 2 ? '❌ Sua compra foi PIOR que a média do mercado' : '⚠️ Sua compra ficou na média do mercado') + '</p>';

        rankingBody.innerHTML = rows.map((r, idx) => {
            const nome = r.isPolar ? '<strong style="color:var(--accent)">' + r.importador + '</strong>' : this.getAnonName(r.importador);
            const badge = r.isPolar ? '<span class="badge badge-excelente">POLAR</span>' : '';
            const cls = r.isPolar ? 'polar-row' : (idx === 0 ? 'best-row' : '');
            return '<tr class="' + cls + '">' +
                '<td class="text-center">' + (idx + 1) + 'º</td>' +
                '<td>' + nome + ' ' + badge + '</td>' +
                '<td class="text-right font-mono">' + this.fmtMoney(r.preco) + '</td>' +
                '<td class="text-right font-mono">' + this.fmtNum(r.qtd) + '</td>' +
                '<td class="text-right font-mono">' + this.fmtMoney(r.fob) + '</td>' +
                '<td class="text-center">' + (r.isPolar ? 'Você' : (idx === 0 ? 'Melhor preço' : '')) + '</td>' +
                '</tr>';
        }).join('');

        // ===== COMPARAÇÃO TODO O MERCADO DO MÊS (MESMO NCM) =====
        const mercadoMes = this.filtered.filter(r => r.anoMes === anoMes && r.ncm === ncm && !this.isPolar(r.importador));
        if (mercadoMes.length > 0) {
            const precosMercado = mercadoMes.map(r => r.qtd > 0 ? r.valorFobTotal / r.qtd : 0).filter(p => p > 0);
            const minMercado = Math.min(...precosMercado);
            const maxMercado = Math.max(...precosMercado);
            const totalQtdMerc = mercadoMes.reduce((s, r) => s + r.qtd, 0);
            const totalFobMerc = mercadoMes.reduce((s, r) => s + r.valorFobTotal, 0);
            const mediaMercado = totalQtdMerc > 0 ? totalFobMerc / totalQtdMerc : 0;
            const diffMinMerc = minMercado > 0 ? ((polarPreco - minMercado) / minMercado) * 100 : 0;
            const diffAvgMerc = mediaMercado > 0 ? ((polarPreco - mediaMercado) / mediaMercado) * 100 : 0;
            const compradoresMerc = [...new Set(mercadoMes.map(r => r.importador))].length;
            const fornecedoresMerc = [...new Set(mercadoMes.map(r => r.exportador))].length;

            mercadoDiv.innerHTML =
                '<p><strong>Compradores no mercado (mês/NCM):</strong> ' + compradoresMerc + ' empresas</p>' +
                '<p><strong>Fornecedores ativos no mercado:</strong> ' + fornecedoresMerc + ' exportadores</p>' +
                '<p><strong>Menor preço do mercado:</strong> US$ ' + this.fmtMoney(minMercado) + '/kg (' + (diffMinMerc > 0 ? '+' : '') + diffMinMerc.toFixed(1) + '% vs POLAR)</p>' +
                '<p><strong>Maior preço do mercado:</strong> US$ ' + this.fmtMoney(maxMercado) + '/kg</p>' +
                '<p><strong>Média ponderada do mercado:</strong> US$ ' + this.fmtMoney(mediaMercado) + '/kg (' + (diffAvgMerc > 0 ? '+' : '') + diffAvgMerc.toFixed(1) + '% vs POLAR)</p>' +
                '<p><strong>Volume total do mercado:</strong> ' + this.fmtNum(totalQtdMerc) + ' kg / US$ ' + this.fmtMoney(totalFobMerc) + ' FOB</p>';
        } else {
            mercadoDiv.innerHTML = '<p style="color:var(--text-light)">Nenhum outro importador encontrado para este NCM neste mês.</p>';
        }

        modal.classList.add('active');
    },

    closeModalCompra() { document.getElementById('modalCompra').classList.remove('active'); },

    fmtMoney(v) { if (v === undefined || v === null || isNaN(v)) return '—'; return v.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 }); },
    fmtNum(v) { if (v === undefined || v === null || isNaN(v)) return '—'; return v.toLocaleString('pt-BR', { minimumFractionDigits: 0, maximumFractionDigits: 0 }); },
    escJS(s) { if (!s) return '""'; return '"' + s.replace(/"/g, '\\"') + '"'; },
};

document.addEventListener('DOMContentLoaded', () => app.init());
