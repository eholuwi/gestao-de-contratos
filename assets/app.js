/* ============================================================
   app.js — Dashboard de Gestão de Contratos
   Estado, navegação, filtros, busca, gráficos, CRUD e geração
   de snapshot. Roda em dois modos, sem duplicar código:

   - Modo LIVE (arquivo aberto normalmente): usa excel.js para
     abrir/gravar a MATRIZ DE CONTRATOS 2026.xlsx.
   - Modo SNAPSHOT (arquivo gerado por "Salvar", distribuído por
     e-mail/rede): window.DASH_CONTRATOS já vem embutido no HTML;
     todo o CRUD e os botões Abrir/Salvar ficam ocultos.
   ============================================================ */
(function () {
  'use strict';

  var R = window.Regras;
  var IS_SNAPSHOT = !!window.DASH_CONTRATOS;
  var HOJE = new Date();

  var CONTRATOS = [];
  var DELETED_ROWS0 = [];
  var DIRTY = false;
  var NEXT_TEMP_ID = -1;
  var CURRENT_PAGE = 'executiva';
  var CHART_INSTANCES = {};

  var GESTAO_FILTERS = { search: '', departamento: '', qualidade: 'todos' };
  var GESTAO_SORT = { field: 'razaoSocial', dir: 'asc' };
  var SLA_FILTERS = { status: 'todos', departamento: '' };

  if (window.Chart && window.ChartDataLabels) {
    Chart.register(ChartDataLabels);
    Chart.defaults.plugins.datalabels = Chart.defaults.plugins.datalabels || {};
    Chart.defaults.plugins.datalabels.display = false;
  }

  // ==============================================================
  // BOOTSTRAP
  // ==============================================================
  function boot() {
    if (IS_SNAPSHOT) {
      loadSnapshotData();
      document.body.classList.add('snapshot-mode');
      hideEditAffordances();
    } else if (!window.ExcelIO || !ExcelIO.isSupported()) {
      showUnsupportedBanner();
    }
    wireNav();
    wireModalGeneric();
    goToPage(CURRENT_PAGE);
    window.addEventListener('beforeunload', function (e) {
      if (DIRTY) { e.preventDefault(); e.returnValue = ''; }
    });
  }

  function loadSnapshotData() {
    var payload = window.DASH_CONTRATOS;
    CONTRATOS = (payload.contratos || []).map(function (c) {
      return Object.assign({}, c, {
        dataInicio: c.dataInicio ? new Date(c.dataInicio) : null,
        dataFim: c.dataFim ? new Date(c.dataFim) : null
      });
    });
    var info = document.getElementById('snapshot-info');
    if (info) {
      info.textContent = 'Modo visualização — gerado em ' + R.formatDateBR(new Date(payload.generatedAt)) +
        (payload.sourceFile ? ' a partir de ' + payload.sourceFile : '');
      info.classList.remove('hidden');
    }
  }

  function hideEditAffordances() {
    document.querySelectorAll('.live-only').forEach(function (el) { el.classList.add('hidden'); });
  }

  function showUnsupportedBanner() {
    var b = document.getElementById('unsupported-banner');
    if (b) b.classList.remove('hidden');
  }

  // ==============================================================
  // NAVEGAÇÃO
  // ==============================================================
  function wireNav() {
    document.querySelectorAll('.nav button[data-page]').forEach(function (btn) {
      btn.addEventListener('click', function () {
        goToPage(btn.getAttribute('data-page'));
        closeSidebarMobile();
      });
    });
    var hamb = document.getElementById('btn-hamburger');
    if (hamb) hamb.addEventListener('click', openSidebarMobile);
    var backdrop = document.getElementById('sidebar-backdrop');
    if (backdrop) backdrop.addEventListener('click', closeSidebarMobile);

    var btnOpen = document.getElementById('btn-open');
    if (btnOpen) btnOpen.addEventListener('click', handleOpenFolder);
    var btnSave = document.getElementById('btn-save');
    if (btnSave) btnSave.addEventListener('click', handleSave);
  }

  function openSidebarMobile() {
    document.getElementById('sidebar').classList.add('open');
    document.getElementById('sidebar-backdrop').classList.add('show');
  }
  function closeSidebarMobile() {
    var sb = document.getElementById('sidebar');
    if (sb) sb.classList.remove('open');
    var bd = document.getElementById('sidebar-backdrop');
    if (bd) bd.classList.remove('show');
  }

  var PAGE_META = {
    executiva: { title: 'Visão Executiva', subtitle: 'KPIs consolidados e alertas' },
    sla: { title: 'SLA Operacional', subtitle: 'Acompanhamento de vencimentos' },
    financeira: { title: 'Análise Financeira', subtitle: 'Despesa por departamento, fornecedor e tipo' },
    gestao: { title: 'Gestão de Contratos', subtitle: 'Cadastro, edição e qualidade de dados' }
  };

  function goToPage(page) {
    CURRENT_PAGE = page;
    document.querySelectorAll('.nav button[data-page]').forEach(function (btn) {
      btn.classList.toggle('active', btn.getAttribute('data-page') === page);
    });
    document.querySelectorAll('.page').forEach(function (p) {
      p.classList.toggle('active', p.id === 'page-' + page);
    });
    var meta = PAGE_META[page];
    document.getElementById('topbar-title').textContent = meta.title;
    document.getElementById('topbar-subtitle').textContent = meta.subtitle;
    renderCurrentPage();
  }

  function renderCurrentPage() {
    if (!CONTRATOS.length) {
      renderEmptyState(CURRENT_PAGE);
      return;
    }
    if (CURRENT_PAGE === 'executiva') renderExecutiva();
    else if (CURRENT_PAGE === 'sla') renderSLA();
    else if (CURRENT_PAGE === 'financeira') renderFinanceira();
    else if (CURRENT_PAGE === 'gestao') renderGestao();
  }

  function renderEmptyState(page) {
    var root = document.getElementById('page-' + page);
    if (!root) return;
    root.innerHTML =
      '<div class="empty-state">' +
      '<div class="big">📂</div>' +
      '<h3>Nenhuma planilha carregada</h3>' +
      '<p>Abra a pasta com a <b>MATRIZ DE CONTRATOS 2026.xlsx</b> para começar.</p>' +
      (IS_SNAPSHOT ? '' : '<button class="btn btn-primary" id="btn-empty-open">📂 Abrir planilha</button>') +
      '</div>';
    var b = document.getElementById('btn-empty-open');
    if (b) b.addEventListener('click', handleOpenFolder);
  }

  // ==============================================================
  // ABRIR / SALVAR
  // ==============================================================
  function handleOpenFolder() {
    if (IS_SNAPSHOT) return;
    if (DIRTY && !confirm('Existem alterações não salvas que serão perdidas. Continuar?')) return;
    toast('Selecione a pasta onde está a MATRIZ DE CONTRATOS 2026.xlsx…');
    ExcelIO.openFolder().then(function (result) {
      CONTRATOS = result.contratos;
      DELETED_ROWS0 = [];
      DIRTY = false;
      updateFileStatus();
      toast('Planilha carregada: ' + result.fileName + ' (' + CONTRATOS.length + ' contratos)');
      renderCurrentPage();
    }).catch(function (err) {
      console.error(err);
      toast(err.message || 'Não foi possível abrir a planilha.', true);
    });
  }

  function handleSave() {
    if (IS_SNAPSHOT) return;
    if (!ExcelIO.getState().dirHandle) {
      toast('Abra a planilha antes de salvar.', true);
      return;
    }
    var btn = document.getElementById('btn-save');
    if (btn) btn.disabled = true;
    ExcelIO.save(CONTRATOS, DELETED_ROWS0).then(function () {
      DELETED_ROWS0 = [];
      DIRTY = false;
      updateFileStatus();
      toast('Planilha e backup atualizados. Gerando snapshot…');
      return generateAndSaveSnapshot();
    }).then(function (name) {
      toast('Salvo — planilha, backup e ' + name + ' atualizados.');
    }).catch(function (err) {
      console.error(err);
      toast(err.message || 'Erro ao salvar.', true);
    }).finally(function () {
      if (btn) btn.disabled = false;
    });
  }

  function updateFileStatus() {
    var st = ExcelIO.getState();
    var nameEl = document.getElementById('file-status-name');
    var dotEl = document.getElementById('file-status-dot');
    if (nameEl) nameEl.textContent = st.fileName || 'Nenhuma planilha aberta';
    if (dotEl) dotEl.className = 'status-dot ' + (DIRTY ? 'dirty' : (st.fileName ? 'on' : 'off'));
    var btnSave = document.getElementById('btn-save');
    if (btnSave) btnSave.disabled = !st.fileName;
  }

  function markDirty() {
    DIRTY = true;
    updateFileStatus();
  }

  // ==============================================================
  // SNAPSHOT — gera HTML autossuficiente (view-only) para distribuição
  // ==============================================================
  function generateAndSaveSnapshot() {
    var payload = {
      generatedAt: new Date().toISOString(),
      sourceFile: ExcelIO.getState().fileName,
      contratos: CONTRATOS.map(function (c) {
        return {
          id: c.id, razaoSocial: c.razaoSocial, tipoServico: c.tipoServico, departamento: c.departamento,
          prazoPagamentoRaw: c.prazoPagamentoRaw, prazoPagamentoDias: c.prazoPagamentoDias,
          prazoNegociadoRaw: c.prazoNegociadoRaw, valorMensal2026: c.valorMensal2026, valorAnual2026: c.valorAnual2026,
          vigencia: c.vigencia, dataInicio: c.dataInicio ? c.dataInicio.toISOString() : null,
          dataFim: c.dataFim ? c.dataFim.toISOString() : null, comentarios: c.comentarios,
          statusManual: c.statusManual, statusLegadoRaw: c.statusLegadoRaw
        };
      })
    };
    var html = buildSnapshotHtml(payload);
    return ExcelIO.saveSnapshotHtml(html);
  }

  // Monta o HTML do snapshot como um bootstrap "limpo" (sem conteúdo dinâmico
  // inline) que decodifica em runtime, no navegador de quem abrir o arquivo,
  // os blobs base64 do modelo (assets/snapshot-template.js). Todo o conteúdo
  // variável (payload de contratos, incluindo texto livre de "Comentários")
  // é embutido só depois de escapar "<", então não há como um comentário do
  // tipo "</script>" fechar a tag prematuramente ao abrir o arquivo gerado.
  function buildSnapshotHtml(payload) {
    var T = window.SNAPSHOT_TEMPLATE;
    if (!T) throw new Error('Modelo de snapshot não encontrado (assets/snapshot-template.js).');
    var dataJson = JSON.stringify(payload).replace(/</g, '\\u003c');
    var bootstrap =
      'window.DASH_CONTRATOS = ' + dataJson + ';\n' +
      'function _decode(b64){\n' +
      '  var bin = atob(b64), bytes = new Uint8Array(bin.length);\n' +
      '  for (var i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);\n' +
      '  return new TextDecoder("utf-8").decode(bytes);\n' +
      '}\n' +
      'var B = window.__SNAP_ASSETS__;\n' +
      'var styleEl = document.createElement("style");\n' +
      'styleEl.textContent = _decode(B.css);\n' +
      'document.head.appendChild(styleEl);\n' +
      'document.body.innerHTML = _decode(B.bodyHtml);\n' +
      '["chartLib", "datalabelsLib", "regrasJs", "appJs"].forEach(function (k) {\n' +
      '  var s = document.createElement("script");\n' +
      '  s.textContent = _decode(B[k]);\n' +
      '  document.body.appendChild(s);\n' +
      '});\n';

    return '<!DOCTYPE html><html lang="pt-BR"><head><meta charset="UTF-8">' +
      '<meta name="viewport" content="width=device-width, initial-scale=1.0">' +
      '<title>Dashboard de Gestão de Contratos (Snapshot)</title>' +
      '</head><body>' +
      '<script>window.__SNAP_ASSETS__ = ' + JSON.stringify(T) + ';<' + '/script>' +
      '<script>' + bootstrap + '<' + '/script>' +
      '</body></html>';
  }

  // ==============================================================
  // HELPERS DE RENDER
  // ==============================================================
  function el(html) { var d = document.createElement('div'); d.innerHTML = html.trim(); return d.firstChild; }

  function kpiCard(label, value, accent, sub) {
    return '<div class="kpi-card' + (accent ? ' accent-' + accent : '') + '">' +
      '<div class="kpi-label">' + label + '</div>' +
      '<div class="kpi-value">' + value + '</div>' +
      (sub ? '<div class="kpi-sub">' + sub + '</div>' : '') + '</div>';
  }

  function statusBadge(status) {
    return '<span class="badge ' + R.STATUS_CSS_CLASS[status] + '">' + status + '</span>';
  }

  function destroyChart(id) {
    if (CHART_INSTANCES[id]) { CHART_INSTANCES[id].destroy(); delete CHART_INSTANCES[id]; }
  }

  function drawDonut(canvasId, labels, values, colors) {
    destroyChart(canvasId);
    var ctx = document.getElementById(canvasId);
    if (!ctx) return;
    CHART_INSTANCES[canvasId] = new Chart(ctx, {
      type: 'doughnut',
      data: { labels: labels, datasets: [{ data: values, backgroundColor: colors, borderWidth: 2, borderColor: '#fff' }] },
      options: {
        maintainAspectRatio: false,
        cutout: '62%',
        plugins: {
          legend: { position: 'bottom', labels: { boxWidth: 10, padding: 14, font: { size: 11.5 } } },
          datalabels: {
            display: true, color: '#fff', font: { weight: '700', size: 11 },
            formatter: function (v, ctx) {
              var total = ctx.chart.data.datasets[0].data.reduce(function (a, b) { return a + b; }, 0);
              return total ? Math.round(v / total * 100) + '%' : '';
            }
          }
        }
      }
    });
  }

  function drawBar(canvasId, labels, values, opts) {
    opts = opts || {};
    destroyChart(canvasId);
    var ctx = document.getElementById(canvasId);
    if (!ctx) return;
    CHART_INSTANCES[canvasId] = new Chart(ctx, {
      type: 'bar',
      data: { labels: labels, datasets: [{ data: values, backgroundColor: opts.color || '#F58220', borderRadius: 4, maxBarThickness: 26 }] },
      options: {
        indexAxis: opts.horizontal ? 'y' : 'x',
        maintainAspectRatio: false,
        scales: {
          x: { grid: { display: !opts.horizontal }, ticks: { font: { size: 10.5 } } },
          y: { grid: { display: opts.horizontal }, ticks: { font: { size: 10.5 } } }
        },
        plugins: {
          legend: { display: false },
          datalabels: {
            display: !!opts.dataLabels, anchor: 'end', align: opts.horizontal ? 'end' : 'top',
            color: '#555', font: { size: 10 },
            formatter: opts.labelFormatter || function (v) { return v; }
          }
        }
      }
    });
  }

  function toast(msg, isError) {
    var t = document.getElementById('toast');
    if (!t) return;
    t.textContent = msg;
    t.className = 'toast show' + (isError ? ' error' : '');
    clearTimeout(toast._t);
    toast._t = setTimeout(function () { t.classList.remove('show'); }, 3400);
  }

  function filteredContratos() { return CONTRATOS; }

  // ==============================================================
  // PÁGINA: EXECUTIVA
  // ==============================================================
  function renderExecutiva() {
    var root = document.getElementById('page-executiva');
    var contratos = filteredContratos();
    var k = R.computeKPIs(contratos, HOJE);

    var alertasVencidosCriticos = contratos
      .filter(function (c) { return R.statusFinal(c, HOJE) === 'Vencido' && c.valorAnual2026 > R.VENCIDO_CRITICO_LIMIAR; })
      .sort(function (a, b) { return (b.valorAnual2026 || 0) - (a.valorAnual2026 || 0); })
      .slice(0, 5);
    var alertasSemData = contratos.filter(function (c) { return !c.dataFim; }).slice(0, 5);
    var alertasSemValor = contratos.filter(function (c) { return c.valorAnual2026 == null; }).slice(0, 5);

    root.innerHTML =
      '<div class="kpi-grid">' +
      kpiCard('Total de Contratos', R.formatNumberBR(k.totalContratos), 'primary') +
      kpiCard('No Prazo', R.formatNumberBR(k.contratosNoPrazo), 'no-prazo') +
      kpiCard('A Vencer (30d)', R.formatNumberBR(k.contratosAVencer), 'a-vencer') +
      kpiCard('Vencidos', R.formatNumberBR(k.contratosVencidos), 'vencido', R.formatPercent(k.pctContratosVencidos) + ' do total') +
      kpiCard('Em Análise', R.formatNumberBR(k.contratosEmAnalise), 'em-analise') +
      '</div>' +
      '<div class="kpi-grid" style="grid-template-columns:repeat(auto-fit,minmax(220px,1fr));margin-bottom:22px;">' +
      kpiCard('Despesa Total Anual', R.formatBRL(k.despesaTotalAnual, true), 'primary') +
      kpiCard('Contratos Sem Valor Informado', R.formatNumberBR(k.contratosSemValorInformado)) +
      kpiCard('Vencidos Críticos (> R$ 100 mil)', R.formatNumberBR(k.contratosVencidosCriticos), 'vencido') +
      '</div>' +
      '<div class="grid grid-2">' +
      '<div class="card"><h3>Status dos Contratos</h3><div class="card-sub">Distribuição por status calculado</div>' +
      '<div class="chart-wrap"><canvas id="chart-exec-donut"></canvas></div></div>' +
      '<div class="card"><h3>Top 8 Tipo de Serviço</h3><div class="card-sub">Por despesa anual (Valor Anual2)</div>' +
      '<div class="chart-wrap"><canvas id="chart-exec-tipo"></canvas></div></div>' +
      '</div>' +
      '<div class="card" style="margin-top:18px;">' +
      '<h3>⚠ Alertas</h3><div class="card-sub">Contratos que precisam de atenção</div>' +
      '<div class="grid grid-3">' +
      alertBlock('Vencidos críticos (' + k.contratosVencidosCriticos + ')', alertasVencidosCriticos, function (c) { return R.formatBRL(c.valorAnual2026, true); }) +
      alertBlock('Sem data de fim (' + k.contratosEmAnalise + ')', alertasSemData, function () { return 'Em Análise'; }) +
      alertBlock('Sem valor informado (' + k.contratosSemValorInformado + ')', alertasSemValor, function () { return '—'; }) +
      '</div></div>';

    drawDonut('chart-exec-donut', R.STATUS_ORDER,
      [k.contratosNoPrazo, k.contratosAVencer, k.contratosVencidos, k.contratosEmAnalise],
      R.STATUS_ORDER.map(function (s) { return R.STATUS_COLORS[s]; }));

    var tipoPairs = R.topN(R.sumBy(contratos, function (c) { return c.tipoServico; }, function (c) { return c.valorAnual2026; }), 8);
    drawBar('chart-exec-tipo', tipoPairs.map(function (p) { return p.label; }), tipoPairs.map(function (p) { return p.value; }),
      { horizontal: true, dataLabels: true, labelFormatter: function (v) { return R.formatBRL(v, true); } });
  }

  function alertBlock(title, items, valueFn) {
    var body = items.length
      ? '<ul class="alert-list">' + items.map(function (c) {
        return '<li><span>' + escapeHtml(c.razaoSocial) + '</span><span class="alert-tag">' + valueFn(c) + '</span></li>';
      }).join('') + '</ul>'
      : '<p class="cell-muted" style="font-size:12.5px;">Nenhum registro.</p>';
    return '<div><h4 style="font-size:12.5px;margin-bottom:8px;">' + title + '</h4>' + body + '</div>';
  }

  // ==============================================================
  // PÁGINA: SLA OPERACIONAL
  // ==============================================================
  function renderSLA() {
    var root = document.getElementById('page-sla');
    var contratos = filteredContratos();
    var k = R.computeKPIs(contratos, HOJE);
    var deptos = distinctDepartamentos(contratos);

    root.innerHTML =
      '<div class="kpi-grid">' +
      kpiCard('A Vencer (30 dias)', R.formatNumberBR(k.contratosAVencer), 'a-vencer') +
      kpiCard('Vencidos', R.formatNumberBR(k.contratosVencidos), 'vencido') +
      kpiCard('Prazo Médio de Pagamento', R.formatDiasSuffix(k.prazoMedioPagamentoDias)) +
      '</div>' +
      '<div class="filter-bar">' +
      '<div class="pill-toggle" id="sla-status-pills">' +
      ['todos'].concat(R.STATUS_ORDER).map(function (s) {
        return '<button data-status="' + s + '" class="' + (SLA_FILTERS.status === s ? 'active' : '') + '">' + (s === 'todos' ? 'Todos' : s) + '</button>';
      }).join('') + '</div>' +
      '<select id="sla-depto-filter"><option value="">Todos os departamentos</option>' +
      deptos.map(function (d) { return '<option value="' + escapeHtml(d) + '"' + (SLA_FILTERS.departamento === d ? ' selected' : '') + '>' + escapeHtml(d) + '</option>'; }).join('') +
      '</select>' +
      '<div class="spacer"></div>' +
      '<span class="filter-count" id="sla-count"></span>' +
      '</div>' +
      '<div class="table-wrap"><div class="table-scroll">' +
      '<table class="data-table"><thead><tr>' +
      '<th>Fornecedor</th><th>Departamento</th><th>Status</th><th>Data Fim</th><th>Dias até vencer</th><th class="cell-num">Valor Anual</th>' +
      '</tr></thead><tbody id="sla-tbody"></tbody></table></div></div>' +
      '<div class="footer-note">Régua: sem data de fim → Em Análise · vencido se Fim ≤ hoje · a vencer se Fim ≤ hoje + 30 dias.</div>';

    document.querySelectorAll('#sla-status-pills button').forEach(function (btn) {
      btn.addEventListener('click', function () { SLA_FILTERS.status = btn.getAttribute('data-status'); renderSLA(); });
    });
    document.getElementById('sla-depto-filter').addEventListener('change', function (e) {
      SLA_FILTERS.departamento = e.target.value; renderSLA();
    });

    var rows = contratos.filter(function (c) {
      var st = R.statusFinal(c, HOJE);
      if (SLA_FILTERS.status !== 'todos' && st !== SLA_FILTERS.status) return false;
      if (SLA_FILTERS.departamento && c.departamento !== SLA_FILTERS.departamento) return false;
      return true;
    }).sort(function (a, b) {
      if (!a.dataFim && !b.dataFim) return 0;
      if (!a.dataFim) return 1;
      if (!b.dataFim) return -1;
      return a.dataFim - b.dataFim;
    });

    document.getElementById('sla-count').textContent = rows.length + ' de ' + contratos.length + ' contratos';
    document.getElementById('sla-tbody').innerHTML = rows.map(function (c) {
      var st = R.statusFinal(c, HOJE);
      var dias = c.dataFim ? Math.round((R.startOfDay(c.dataFim) - R.startOfDay(HOJE)) / 86400000) : null;
      return '<tr>' +
        '<td>' + escapeHtml(c.razaoSocial) + '</td>' +
        '<td>' + escapeHtml(c.departamento) + '</td>' +
        '<td>' + statusBadge(st) + '</td>' +
        '<td>' + R.formatDateBR(c.dataFim) + '</td>' +
        '<td class="cell-num">' + (dias == null ? '—' : dias) + '</td>' +
        '<td class="cell-num">' + R.formatBRL(c.valorAnual2026) + '</td>' +
        '</tr>';
    }).join('') || '<tr><td colspan="6" class="cell-muted" style="text-align:center;padding:24px;">Nenhum contrato encontrado.</td></tr>';
  }

  // ==============================================================
  // PÁGINA: FINANCEIRA
  // ==============================================================
  function renderFinanceira() {
    var root = document.getElementById('page-financeira');
    var contratos = filteredContratos();
    var k = R.computeKPIs(contratos, HOJE);

    root.innerHTML =
      '<div class="kpi-grid">' +
      kpiCard('Despesa Total Anual', R.formatBRL(k.despesaTotalAnual, true), 'primary') +
      kpiCard('Despesa Média por Contrato', R.formatBRL(k.despesaMediaPorContrato, true)) +
      kpiCard('Despesa Contratos Vencidos', R.formatBRL(k.despesaContratosVencidos, true), 'vencido') +
      kpiCard('Contratos Vigência Longa (>3 anos)', R.formatNumberBR(k.contratosVigenciaLonga)) +
      '</div>' +
      '<div class="grid grid-2">' +
      '<div class="card"><h3>Despesa por Departamento</h3><div class="card-sub">Soma de Valor Anual2</div>' +
      '<div class="chart-wrap"><canvas id="chart-fin-depto"></canvas></div></div>' +
      '<div class="card"><h3>Top 10 Fornecedores por Despesa</h3><div class="card-sub">Soma de Valor Anual2</div>' +
      '<div class="table-wrap"><div class="table-scroll" style="max-height:300px;">' +
      '<table class="data-table"><thead><tr><th>Fornecedor</th><th class="cell-num">Despesa Anual</th><th class="cell-num">Contratos</th></tr></thead>' +
      '<tbody id="fin-top10-tbody"></tbody></table></div></div></div>' +
      '</div>' +
      '<div class="grid grid-2" style="margin-top:18px;">' +
      '<div class="card"><h3>Despesa Mensal 2026</h3><div class="card-sub">Colunas JAN–DEZ da planilha</div>' +
      placeholderBlock('Aguardando preenchimento da planilha (colunas JAN a DEZ estão vazias na fonte hoje).') + '</div>' +
      '<div class="card"><h3>Evolução 2024 / 2025 / 2026</h3><div class="card-sub">Comparativo anual</div>' +
      placeholderBlock('Aguardando preenchimento da planilha (colunas TOTAL ANUAL 2024/2025 estão zeradas na fonte hoje).') + '</div>' +
      '</div>';

    var deptoPairs = R.sumBy(contratos, function (c) { return c.departamento; }, function (c) { return c.valorAnual2026; })
      .sort(function (a, b) { return b.value - a.value; });
    drawBar('chart-fin-depto', deptoPairs.map(function (p) { return p.label; }), deptoPairs.map(function (p) { return p.value; }),
      { color: '#F58220', dataLabels: false });

    var byFornecedor = {};
    contratos.forEach(function (c) {
      var k2 = c.razaoSocial || 'Não informado';
      byFornecedor[k2] = byFornecedor[k2] || { despesa: 0, n: 0 };
      if (c.valorAnual2026 != null) byFornecedor[k2].despesa += c.valorAnual2026;
      byFornecedor[k2].n++;
    });
    var top10 = Object.keys(byFornecedor).map(function (nome) {
      return { nome: nome, despesa: byFornecedor[nome].despesa, n: byFornecedor[nome].n };
    }).sort(function (a, b) { return b.despesa - a.despesa; }).slice(0, 10);

    document.getElementById('fin-top10-tbody').innerHTML = top10.map(function (r) {
      return '<tr><td>' + escapeHtml(r.nome) + '</td><td class="cell-num">' + R.formatBRL(r.despesa) + '</td><td class="cell-num">' + r.n + '</td></tr>';
    }).join('');
  }

  function placeholderBlock(msg) {
    return '<div style="padding:40px 10px;text-align:center;color:var(--muted);font-size:12.5px;">🕒 ' + msg + '</div>';
  }

  // ==============================================================
  // PÁGINA: GESTÃO DE CONTRATOS (tabela + CRUD + qualidade)
  // ==============================================================
  function distinctDepartamentos(contratos) {
    var set = {};
    contratos.forEach(function (c) { if (c.departamento) set[c.departamento] = true; });
    return Object.keys(set).sort();
  }
  function distinctTipos(contratos) {
    var set = {};
    contratos.forEach(function (c) { if (c.tipoServico) set[c.tipoServico] = true; });
    return Object.keys(set).sort();
  }

  function renderGestao() {
    var root = document.getElementById('page-gestao');
    var deptos = distinctDepartamentos(CONTRATOS);

    root.innerHTML =
      '<div class="filter-bar">' +
      (IS_SNAPSHOT ? '' : '<button class="btn btn-primary live-only" id="btn-novo-contrato">+ Novo Contrato</button>') +
      '<input type="text" id="gestao-search" placeholder="Buscar fornecedor, tipo, departamento, comentário…" value="' + escapeHtml(GESTAO_FILTERS.search) + '">' +
      '<select id="gestao-depto-filter"><option value="">Todos os departamentos</option>' +
      deptos.map(function (d) { return '<option value="' + escapeHtml(d) + '"' + (GESTAO_FILTERS.departamento === d ? ' selected' : '') + '>' + escapeHtml(d) + '</option>'; }).join('') +
      '</select>' +
      '<select id="gestao-qualidade-filter">' +
      '<option value="todos"' + (GESTAO_FILTERS.qualidade === 'todos' ? ' selected' : '') + '>Qualidade: Todos</option>' +
      '<option value="semData"' + (GESTAO_FILTERS.qualidade === 'semData' ? ' selected' : '') + '>Sem data de fim</option>' +
      '<option value="semValor"' + (GESTAO_FILTERS.qualidade === 'semValor' ? ' selected' : '') + '>Sem valor informado</option>' +
      '<option value="statusInconsistente"' + (GESTAO_FILTERS.qualidade === 'statusInconsistente' ? ' selected' : '') + '>Status manual divergente</option>' +
      '</select>' +
      '<div class="spacer"></div>' +
      '<span class="filter-count" id="gestao-count"></span>' +
      '</div>' +
      '<div class="table-wrap"><div class="table-scroll">' +
      '<table class="data-table"><thead><tr>' +
      thSort('razaoSocial', 'Fornecedor') + thSort('departamento', 'Depto') + thSort('tipoServico', 'Tipo de Serviço') +
      thSort('status', 'Status') + thSort('dataInicio', 'Início') + thSort('dataFim', 'Fim') +
      thSort('valorMensal2026', 'Valor Mensal') + thSort('valorAnual2026', 'Valor Anual') +
      '<th>Qualidade</th>' + (IS_SNAPSHOT ? '' : '<th class="live-only">Ações</th>') +
      '</tr></thead><tbody id="gestao-tbody"></tbody></table></div></div>';

    document.getElementById('gestao-search').addEventListener('input', function (e) {
      GESTAO_FILTERS.search = e.target.value; renderGestaoTable();
    });
    document.getElementById('gestao-depto-filter').addEventListener('change', function (e) {
      GESTAO_FILTERS.departamento = e.target.value; renderGestaoTable();
    });
    document.getElementById('gestao-qualidade-filter').addEventListener('change', function (e) {
      GESTAO_FILTERS.qualidade = e.target.value; renderGestaoTable();
    });
    document.querySelectorAll('#page-gestao th[data-sort]').forEach(function (th) {
      th.addEventListener('click', function () {
        var f = th.getAttribute('data-sort');
        if (GESTAO_SORT.field === f) GESTAO_SORT.dir = GESTAO_SORT.dir === 'asc' ? 'desc' : 'asc';
        else { GESTAO_SORT.field = f; GESTAO_SORT.dir = 'asc'; }
        renderGestao();
      });
    });
    var btnNovo = document.getElementById('btn-novo-contrato');
    if (btnNovo) btnNovo.addEventListener('click', function () { openContratoModal(null); });

    renderGestaoTable();
  }

  function thSort(field, label) {
    var active = GESTAO_SORT.field === field;
    return '<th data-sort="' + field + '" class="' + (active ? 'sort-active' : '') + '">' + label + (active ? (GESTAO_SORT.dir === 'asc' ? ' ▲' : ' ▼') : '') + '</th>';
  }

  function applyGestaoFilters() {
    var q = R.stripAccents(GESTAO_FILTERS.search || '').toLowerCase().trim();
    return CONTRATOS.filter(function (c) {
      if (GESTAO_FILTERS.departamento && c.departamento !== GESTAO_FILTERS.departamento) return false;
      if (q) {
        var hay = R.stripAccents([c.razaoSocial, c.tipoServico, c.departamento, c.comentarios].join(' ')).toLowerCase();
        if (hay.indexOf(q) === -1) return false;
      }
      if (GESTAO_FILTERS.qualidade !== 'todos') {
        var qf = R.qualityFlags(c, HOJE);
        if (!qf[GESTAO_FILTERS.qualidade]) return false;
      }
      return true;
    });
  }

  function renderGestaoTable() {
    var rows = applyGestaoFilters();
    rows = rows.slice().sort(function (a, b) {
      var f = GESTAO_SORT.field, dir = GESTAO_SORT.dir === 'asc' ? 1 : -1;
      var va, vb;
      if (f === 'status') { va = R.statusFinal(a, HOJE); vb = R.statusFinal(b, HOJE); }
      else { va = a[f]; vb = b[f]; }
      if (va == null && vb == null) return 0;
      if (va == null) return 1;
      if (vb == null) return -1;
      if (va instanceof Date) return (va - vb) * dir;
      if (typeof va === 'number') return (va - vb) * dir;
      return String(va).localeCompare(String(vb), 'pt-BR') * dir;
    });

    document.getElementById('gestao-count').textContent = rows.length + ' de ' + CONTRATOS.length + ' contratos';
    document.getElementById('gestao-tbody').innerHTML = rows.map(function (c) {
      var st = R.statusFinal(c, HOJE);
      var qf = R.qualityFlags(c, HOJE);
      var flags = [];
      if (qf.semData) flags.push('<span class="badge badge-flag" title="Sem data de fim">Sem data</span>');
      if (qf.semValor) flags.push('<span class="badge badge-flag" title="Sem valor informado">Sem valor</span>');
      if (qf.statusInconsistente) flags.push('<span class="badge badge-flag" title="Status manual diferente do calculado">Status manual</span>');
      return '<tr class="' + (IS_SNAPSHOT ? '' : 'row-clickable') + '" data-id="' + c.id + '">' +
        '<td>' + escapeHtml(c.razaoSocial) + '</td>' +
        '<td>' + escapeHtml(c.departamento) + '</td>' +
        '<td>' + escapeHtml(c.tipoServico) + '</td>' +
        '<td>' + statusBadge(st) + '</td>' +
        '<td>' + R.formatDateBR(c.dataInicio) + '</td>' +
        '<td>' + R.formatDateBR(c.dataFim) + '</td>' +
        '<td class="cell-num">' + R.formatBRL(c.valorMensal2026) + '</td>' +
        '<td class="cell-num">' + R.formatBRL(c.valorAnual2026) + '</td>' +
        '<td>' + (flags.join(' ') || '<span class="cell-muted">OK</span>') + '</td>' +
        (IS_SNAPSHOT ? '' :
          '<td class="td-actions live-only">' +
          '<button class="btn btn-sm btn-icon" data-action="edit" title="Editar">✏️</button>' +
          '<button class="btn btn-sm btn-icon btn-danger" data-action="del" title="Excluir">🗑️</button>' +
          '</td>') +
        '</tr>';
    }).join('') || '<tr><td colspan="' + (IS_SNAPSHOT ? 9 : 10) + '" class="cell-muted" style="text-align:center;padding:24px;">Nenhum contrato encontrado com os filtros atuais.</td></tr>';

    if (!IS_SNAPSHOT) {
      document.querySelectorAll('#gestao-tbody tr[data-id]').forEach(function (tr) {
        var id = Number(tr.getAttribute('data-id'));
        tr.querySelector('[data-action="edit"]').addEventListener('click', function (e) { e.stopPropagation(); openContratoModal(id); });
        tr.querySelector('[data-action="del"]').addEventListener('click', function (e) { e.stopPropagation(); deleteContrato(id); });
        tr.addEventListener('click', function () { openContratoModal(id); });
      });
    }
  }

  // ==============================================================
  // MODAL — cadastro/edição de contrato
  // ==============================================================
  function wireModalGeneric() {
    var overlay = document.getElementById('modal-overlay');
    if (!overlay) return;
    overlay.addEventListener('click', function (e) { if (e.target === overlay) closeModal(); });
  }

  function closeModal() {
    var overlay = document.getElementById('modal-overlay');
    if (overlay) overlay.classList.add('hidden');
    document.getElementById('modal-root').innerHTML = '';
  }

  function fieldText(id, label, value, opts) {
    opts = opts || {};
    return '<div class="field' + (opts.full ? ' full' : '') + '" id="fld-' + id + '">' +
      '<label for="in-' + id + '">' + label + (opts.required ? ' *' : '') + '</label>' +
      '<input type="text" id="in-' + id + '" value="' + escapeHtml(value || '') + '"' + (opts.list ? ' list="dl-' + id + '"' : '') + '>' +
      (opts.list ? '<datalist id="dl-' + id + '">' + opts.list.map(function (v) { return '<option value="' + escapeHtml(v) + '">'; }).join('') + '</datalist>' : '') +
      '<div class="field-error">' + (opts.errorMsg || 'Campo obrigatório.') + '</div></div>';
  }
  function fieldDate(id, label, value) {
    return '<div class="field" id="fld-' + id + '"><label for="in-' + id + '">' + label + '</label>' +
      '<input type="date" id="in-' + id + '" value="' + R.toDateInputValue(value) + '">' +
      '<div class="field-error">Data inválida.</div></div>';
  }
  function fieldMoney(id, label, value) {
    return '<div class="field" id="fld-' + id + '"><label for="in-' + id + '">' + label + '</label>' +
      '<input type="text" id="in-' + id + '" value="' + (value == null ? '' : R.formatNumberBR(value, 2)) + '" placeholder="0,00">' +
      '<div class="field-error">Valor numérico inválido.</div></div>';
  }
  function fieldSelect(id, label, value, options) {
    return '<div class="field" id="fld-' + id + '"><label for="in-' + id + '">' + label + '</label>' +
      '<select id="in-' + id + '">' + options.map(function (o) {
        return '<option value="' + escapeHtml(o.value) + '"' + (o.value === (value || '') ? ' selected' : '') + '>' + escapeHtml(o.label) + '</option>';
      }).join('') + '</select></div>';
  }
  function fieldTextarea(id, label, value) {
    return '<div class="field full" id="fld-' + id + '"><label for="in-' + id + '">' + label + '</label>' +
      '<textarea id="in-' + id + '">' + escapeHtml(value || '') + '</textarea></div>';
  }

  function openContratoModal(id) {
    var c = id == null ? null : CONTRATOS.find(function (x) { return x.id === id; });
    var isEdit = !!c;
    var deptos = distinctDepartamentos(CONTRATOS);
    var tipos = distinctTipos(CONTRATOS);
    var statusOptions = [{ value: '', label: 'Automático (por data)' }].concat(R.STATUS_ORDER.map(function (s) { return { value: s, label: s }; }));

    var html =
      '<div class="modal">' +
      '<div class="modal-head"><h2>' + (isEdit ? 'Editar Contrato' : 'Novo Contrato') + '</h2>' +
      '<button class="modal-close" id="modal-close-btn">✕</button></div>' +
      '<div class="modal-body"><div class="form-grid">' +
      fieldText('razaoSocial', 'Fornecedor (Razão Social)', c ? c.razaoSocial : '', { required: true, full: true }) +
      fieldText('departamento', 'Departamento', c ? c.departamento : '', { list: deptos }) +
      fieldText('tipoServico', 'Tipo de Serviço', c ? c.tipoServico : '', { list: tipos }) +
      fieldDate('dataInicio', 'Início', c ? c.dataInicio : null) +
      fieldDate('dataFim', 'Fim', c ? c.dataFim : null) +
      fieldMoney('valorMensal2026', 'Valor Mensal (R$)', c ? c.valorMensal2026 : null) +
      fieldMoney('valorAnual2026', 'Valor Anual (R$)', c ? c.valorAnual2026 : null) +
      fieldText('vigencia', 'Vigência', c ? c.vigencia : '') +
      fieldText('prazoPagamentoRaw', 'Prazo de Pagamento', c ? c.prazoPagamentoRaw : '', { errorMsg: '' }) +
      fieldText('prazoNegociadoRaw', 'Prazo Negociado', c ? c.prazoNegociadoRaw : '') +
      fieldSelect('statusManual', 'Status Manual', c ? (c.statusManual || '') : '', statusOptions) +
      fieldTextarea('comentarios', 'Comentários', c ? c.comentarios : '') +
      '</div></div>' +
      '<div class="modal-foot">' +
      '<button class="btn" id="modal-cancel-btn">Cancelar</button>' +
      '<button class="btn btn-primary" id="modal-save-btn">Salvar</button>' +
      '</div></div>';

    document.getElementById('modal-root').innerHTML = html;
    document.getElementById('modal-overlay').classList.remove('hidden');
    document.getElementById('modal-close-btn').addEventListener('click', closeModal);
    document.getElementById('modal-cancel-btn').addEventListener('click', closeModal);
    document.getElementById('modal-save-btn').addEventListener('click', function () { submitContratoModal(id); });
  }

  function setFieldError(id, hasError) {
    var f = document.getElementById('fld-' + id);
    if (f) f.classList.toggle('invalid', !!hasError);
  }
  function val(id) { var e = document.getElementById('in-' + id); return e ? e.value : ''; }

  function submitContratoModal(id) {
    var razaoSocial = val('razaoSocial').trim();
    var dataInicio = R.parseDateInput(val('dataInicio'));
    var dataFim = R.parseDateInput(val('dataFim'));
    var valorMensalRaw = val('valorMensal2026');
    var valorAnualRaw = val('valorAnual2026');
    var valorMensal = R.parseMoneyInput(valorMensalRaw);
    var valorAnual = R.parseMoneyInput(valorAnualRaw);

    var ok = true;
    setFieldError('razaoSocial', false);
    setFieldError('dataFim', false);
    setFieldError('valorMensal2026', false);
    setFieldError('valorAnual2026', false);

    if (!razaoSocial) { setFieldError('razaoSocial', true); ok = false; }
    if (dataInicio && dataFim && dataFim < dataInicio) { setFieldError('dataFim', true); ok = false; }
    if (valorMensalRaw.trim() !== '' && valorMensal == null) { setFieldError('valorMensal2026', true); ok = false; }
    if (valorAnualRaw.trim() !== '' && valorAnual == null) { setFieldError('valorAnual2026', true); ok = false; }
    if (!ok) return;

    var prazoRaw = val('prazoPagamentoRaw').trim();
    var statusManual = val('statusManual') || null;

    var fields = {
      razaoSocial: R.normalizeRazaoSocial(razaoSocial),
      departamento: R.normalizeDepartamento(val('departamento').trim()),
      tipoServico: R.normalizeTipoServico(val('tipoServico').trim()),
      dataInicio: dataInicio,
      dataFim: dataFim,
      valorMensal2026: valorMensal,
      valorAnual2026: valorAnual,
      vigencia: val('vigencia').trim(),
      prazoPagamentoRaw: prazoRaw,
      prazoPagamentoDias: R.parsePrazoPagamentoDias(prazoRaw),
      prazoNegociadoRaw: val('prazoNegociadoRaw').trim(),
      statusManual: statusManual,
      comentarios: val('comentarios').trim()
    };

    if (id == null) {
      var novo = Object.assign({ id: NEXT_TEMP_ID--, isNew: true, _row0: null, statusLegadoRaw: '' }, fields);
      CONTRATOS.push(novo);
    } else {
      var c = CONTRATOS.find(function (x) { return x.id === id; });
      Object.assign(c, fields);
    }
    markDirty();
    closeModal();
    toast('Contrato ' + (id == null ? 'cadastrado' : 'atualizado') + '. Clique em "Salvar" para gravar na planilha.');
    renderCurrentPage();
  }

  function deleteContrato(id) {
    var c = CONTRATOS.find(function (x) { return x.id === id; });
    if (!c) return;
    if (!confirm('Excluir o contrato de "' + c.razaoSocial + '"? Esta ação só é gravada na planilha ao clicar em "Salvar".')) return;
    if (!c.isNew && c._row0 != null) DELETED_ROWS0.push(c._row0);
    CONTRATOS = CONTRATOS.filter(function (x) { return x.id !== id; });
    markDirty();
    toast('Contrato removido. Clique em "Salvar" para gravar na planilha.');
    renderCurrentPage();
  }

  // ==============================================================
  // UTIL
  // ==============================================================
  function escapeHtml(s) {
    if (s == null) return '';
    return String(s).replace(/[&<>"']/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
    });
  }

  // No snapshot gerado em runtime, este script é injetado DEPOIS que
  // DOMContentLoaded já disparou no documento hospedeiro — nesse caso
  // chamamos boot() direto em vez de esperar um evento que não virá mais.
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', boot);
  } else {
    boot();
  }
})();
