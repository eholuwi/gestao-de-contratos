/* ============================================================
   app.js — Dashboard de Gestão de Contratos
   Estado, navegação, filtros, busca, gráficos, CRUD e geração
   de snapshot. Roda em dois modos, sem duplicar código:

   - Modo LIVE (arquivo aberto normalmente): usa excel.js para
     abrir/gravar a MATRIZ DE CONTRATOS 2026.xlsx.
   - Modo SNAPSHOT (arquivo gerado por "Salvar", distribuído por
     e-mail/rede): window.DASH_CONTRATOS já vem embutido no HTML;
     todo o CRUD e os botões Abrir/Salvar ficam ocultos.
   - V1.7.0, modo BASE (MODO_BASE): a pasta tem as planilhas do Power BI
     (MATRIZ DE POWERBI CONTRATOS <ano>.xlsx). Mesmos números do Power BI:
     status pela data de posição da planilha (refAtiva) e cada exercício
     lido do seu arquivo (contratosDoAno).
   ============================================================ */
(function () {
  'use strict';

  var R = window.Regras;
  var IS_SNAPSHOT = !!window.DASH_CONTRATOS;
  var HOJE = new Date();

  var CONTRATOS = [];
  var DELETED_ROWS0 = [];
  var COTACOES = {}; // {ano: {jan: 4, fev: 5.09, ...}} — cotação do dólar por ano/mês, lida da linha 2 da planilha
  var DIRTY = false;
  var NEXT_TEMP_ID = -1;
  var CURRENT_PAGE = 'home';
  var CHART_INSTANCES = {};

  var GESTAO_FILTERS = { search: '', departamento: '' };
  var GESTAO_SORT = { field: 'razaoSocial', dir: 'asc' };
  var SLA_FILTERS = { status: 'todos', departamento: '' };
  // Anos com colunas mensais: vem da planilha ao abrir (anosMensais), com o
  // padrão do Regras como fallback (ex.: snapshot de uma versão antiga).
  var ANOS_MENSAL = Regras.ANOS_MENSAIS.slice();
  // Exercício ativo — estado ÚNICO de ano do app. Nasce no ano em que estamos
  // (não no mais recente da planilha) e o seletor da barra superior, que é
  // replicado no topo de cada página, escreve aqui: KPIs, rankings, colunas
  // mensais da tabela de Gestão e os gráficos falam sempre do mesmo ano.
  var ANO_ATIVO = R.anoAtual();
  // V1.6.0 — modo "planilha por ano": quando a pasta aberta tem uma MATRIZ por
  // exercício, este mapa lista {ano: {fileName}} e o seletor de ano TROCA a
  // planilha ativa (via ExcelIO.ativarAno). null = modo clássico (uma planilha
  // com blocos multi-ano).
  var ANOS_PLANILHAS = null;
  // V1.7.0 — modo "base do Power BI": a pasta tem as MATRIZ DE POWERBI CONTRATOS <ano>.xlsx
  // (a mesma base do Power BI). Status e dias contam da data de POSIÇÃO de cada planilha
  // (31/12 em ano fechado; dia do último recálculo no Excel no ano corrente) — o mesmo
  // número do Power BI. No snapshot, posição e ano anterior vêm do payload.
  var MODO_BASE = false;
  // V1.9.2 (protótipo B) — "trabalhar conforme hoje". No app ao vivo, status e
  // "dias até vencer" passam a ser calculados com HOJE (a data de hoje) direto
  // sobre o FIM do contrato; a posição da planilha deixa de ser a régua e vira
  // só retrato/nota ("até quando a planilha foi recalculada"). No snapshot, o dia
  // fica congelado (REF_SNAPSHOT) como sempre.
  var STATUS_CONFORME_HOJE = !IS_SNAPSHOT;
  var REF_SNAPSHOT = null;
  var POSICAO_SNAPSHOT = '';
  var OUTROS_ANOS = {};
  // Modo de demonstração ("?mock=1" na URL ou window.DASH_MOCK antes do script):
  // injeta dispêndio mensal sintético em memória para quem quer ver os gráficos
  // cheios sem tocar na planilha real. Nada é gravado no arquivo.
  var SIMULAR = /[?&]mock=1\b/i.test(location.search) || !!window.DASH_MOCK;
  // V1.9.0 — saving do exercício ativo: aba Saving da base (molde v3), um item por
  // BID/negociação, com a sua própria linha na planilha (_row0) — não depende da
  // posição dos contratos. Cada ano tem o seu arquivo: trocar o ano recarrega.
  var SAVING = [];
  var DELETED_SAVING0 = [];
  var TEM_SAVING = false; // a planilha aberta tem a aba Saving: habilita o editor
  var NEXT_TEMP_SAV_ID = -1;

  function cotaDe(ano, mesKey) {
    var porAno = COTACOES[ano];
    return porAno ? porAno[mesKey] : undefined;
  }

  // Posição literal da planilha (retrato): só isso. Não é mais a régua do status.
  function refPlanilha() {
    if (IS_SNAPSHOT) return REF_SNAPSHOT || HOJE;
    var r = (MODO_BASE && window.ExcelIO) ? ExcelIO.referenciaDoAno(ANO_ATIVO) : null;
    return r || HOJE;
  }
// Data de referência do status e dos "dias até vencer" (V1.9.2, protótipo B):
// no ar, o ANO CORRENTE usa HOJE (conforme a data de hoje); anos fechados seguem
// na referência da planilha (31/12 congelado) — retrato do passado. Snapshot usa
// o dia congelado do payload.
  function refAtiva() {
    if (STATUS_CONFORME_HOJE && ANO_ATIVO === R.anoAtual()) return HOJE;
    return refPlanilha();
  }
  function posicaoAtiva() {
    if (IS_SNAPSHOT) return POSICAO_SNAPSHOT;
    return (MODO_BASE && window.ExcelIO) ? ExcelIO.posicaoDoAno(ANO_ATIVO) : '';
  }
  // Contratos de um exercício. No modo base cada ano tem o seu arquivo (substate no
  // ExcelIO); no layout antigo os anos são colunas dos mesmos contratos.
  function contratosDoAno(ano) {
    if (ano === ANO_ATIVO) return filteredContratos();
    if (IS_SNAPSHOT) return OUTROS_ANOS[ano] || filteredContratos();
    if (MODO_BASE && window.ExcelIO) return ExcelIO.contratosDoAno(ano) || [];
    return filteredContratos();
  }
  // Anos cujos meses o formulário mostra: no modo base, só o do arquivo aberto.
  function anosDoFormulario() { return MODO_BASE ? [ANO_ATIVO] : ANOS_MENSAL; }

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
    } else if (SIMULAR) {
      // Demonstração autocontida: gera os contratos direto em memória, sem
      // exigir "Abrir planilha" — a File System Access API (showDirectoryPicker)
      // não funciona em file:// nem em navegadores mais antigos, e o objetivo
      // do ?mock=1 é ver o dashboard cheio sem essa barreira.
      loadPlanilhaResult({ contratos: gerarContratosSimulados(), cotacoes: {}, anosMensais: R.ANOS_MENSAIS.slice() });
    } else if (window.DesktopBridge) {
      autoOpenDesktop();
    } else if (window.ExcelIO && ExcelIO.isSupported()) {
      autoOpenNavegador();
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

  function autoOpenDesktop() {
    if (!window.ExcelIO || typeof ExcelIO.openDesktopDefault !== 'function') return;
    ExcelIO.openDesktopDefault().then(function (result) {
      loadPlanilhaResult(result);
      toast('Planilha carregada: ' + result.fileName + ' (' + result.contratos.length + ' contratos)');
    }).catch(function (err) {
      console.warn('[desktop] auto-open:', err && err.message);
      toast((err && err.message) || 'Clique em "Abrir planilha" para escolher a planilha.', true);
    });
  }

  // V1.9.1 — navegador: reabre automaticamente a ÚLTIMA PASTA escolhida (handle guardado
  // no IndexedDB ao clicar em "Abrir planilha"). Sem permissão persistida, informa que
  // precisa de um clique — nunca abre picker sem gesto do usuário.
  function autoOpenNavegador() {
    if (!window.ExcelIO || typeof ExcelIO.abrirPastaPrincipal !== 'function') return;
    ExcelIO.abrirPastaPrincipal().then(function (result) {
      if (!result) {
        avisarPrimeiraAbertura();
        return;
      }
      if (result.precisaGestor) {
        toast('Clique em "Abrir planilha" para reabrir a última pasta. Depois disso, abre sozinha.');
        return;
      }
      loadPlanilhaResult(result);
      toast('Planilha reaberta automaticamente: ' + result.fileName + ' (' + result.contratos.length + ' contratos)');
    }).catch(function (err) {
      console.warn('[navegador] auto-open:', err && err.message);
      toast('Não foi possível reabrir automaticamente. Clique em "Abrir planilha".', true);
    });
  }

  // Primeira vez sem pasta lembrada: explica o 1º clique (mostrar uma única vez).
  function avisarPrimeiraAbertura() {
    try {
      if (localStorage.getItem('dashboard-reabrir-aviso')) return;
      localStorage.setItem('dashboard-reabrir-aviso', '1');
    } catch (e) { /* origin sem storage — mostra mesmo assim */ }
    toast('Primeira vez: clique em "Abrir planilha" e escolha a pasta (ex.: _arquivo). Depois, o dashboard abre sozinho.');
  }

  function loadSnapshotData() {
    var payload = window.DASH_CONTRATOS;
    CONTRATOS = (payload.contratos || []).map(function (c) {
      return Object.assign({}, c, {
        dataInicio: c.dataInicio ? new Date(c.dataInicio) : null,
        dataFim: c.dataFim ? new Date(c.dataFim) : null
      });
    });
    COTACOES = payload.cotacoes || {};
    ANOS_MENSAL = (payload.anosMensais && payload.anosMensais.length)
      ? payload.anosMensais.slice()
      : Regras.ANOS_MENSAIS.slice();
    MODO_BASE = !!payload.modoBase;
    REF_SNAPSHOT = payload.referencia ? new Date(payload.referencia) : null;
    POSICAO_SNAPSHOT = payload.posicao || '';
    OUTROS_ANOS = payload.outrosAnos || {};
    SAVING = payload.saving || [];
    TEM_SAVING = !!payload.temSaving;
    ANO_ATIVO = payload.anoAtivo || anoPadrao();
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
    var btnExport = document.getElementById('btn-export');
    if (btnExport) btnExport.addEventListener('click', handleExportarHtmlDoDia);

    renderSeletoresAno();
    var selAno = document.getElementById('ano-ativo');
    if (selAno) wireSeletorAno(selAno);

    var btnCollapse = document.getElementById('btn-collapse');
    if (btnCollapse) btnCollapse.addEventListener('click', toggleSidebarCollapse);
    document.querySelectorAll('[data-action="fullscreen"]').forEach(function (btn) {
      btn.addEventListener('click', toggleFullscreen);
    });
    document.addEventListener('keydown', function (e) {
      if (e.key === 'Escape' && document.getElementById('app').classList.contains('fullscreen')) toggleFullscreen();
    });
  }

  // Sidebar só com ícones (o texto some via CSS; o title do botão vira a dica).
  function toggleSidebarCollapse() {
    var collapsed = document.getElementById('app').classList.toggle('sidebar-collapsed');
    var btn = document.getElementById('btn-collapse');
    if (btn) {
      btn.textContent = collapsed ? '»' : '«';
      btn.title = collapsed ? 'Expandir menu' : 'Recolher menu';
    }
    resizeCharts();
  }

  // Tela cheia da página: esconde sidebar e topbar (Esc ou o botão flutuante saem).
  function toggleFullscreen() {
    document.getElementById('app').classList.toggle('fullscreen');
    resizeCharts();
  }

  // O Chart.js só recalcula o canvas quando o container muda de tamanho por
  // resize da janela — trocar a largura via classe precisa deste empurrão.
  function resizeCharts() {
    Object.keys(CHART_INSTANCES).forEach(function (id) { CHART_INSTANCES[id].resize(); });
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
    home: { title: 'Home', subtitle: 'Apresentação do sistema' },
    executiva: { title: 'Visão Executiva', subtitle: 'KPIs consolidados' },
    sla: { title: 'Status do contrato', subtitle: 'Acompanhamento de vencimentos' },
    financeira: { title: 'Financeiro', subtitle: 'Despesa por departamento, fornecedor e tipo' },
    gestao: { title: 'Gestão de Contratos', subtitle: 'Cadastro e edição dos contratos' }
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

  // Exercício que abre o app: o ano em que estamos. Só cai para o ano mais
  // recente da planilha se o app estiver rodando num ano que não existe como
  // número (snapshot de uma base sem colunas, por exemplo).
  function anoPadrao() {
    var atual = R.anoAtual();
    return /^\d{4}$/.test(atual) ? atual : (ANOS_MENSAL[0] || '2026');
  }

  // ---------------------------------------------------------------
  // SELETOR DE EXERCÍCIO
  //
  // Há um seletor no topo de CADA página (além do da barra superior) e todos
  // leem e escrevem o MESMO estado, ANO_ATIVO. O padrão é o ano em que estamos,
  // mesmo que a planilha ainda não tenha as colunas dele — nesse caso a opção
  // entra como "<ano> — em preparação" e os números daquele exercício saem
  // zerados, porque `Regras` não acha os campos de mês.
  // ---------------------------------------------------------------
  // Anos oferecidos: no modo "planilha por ano" (V1.6.0), só os exercícios que
// têm arquivo na pasta — nunca inventa ano sem planilha. No modo clássico, os
// que a planilha tem, mais o ano corrente (que pode não ter bloco ainda) e o
// ano que estiver aberto.
  function opcoesAno() {
    if (ANOS_PLANILHAS) {
      return Object.keys(ANOS_PLANILHAS).sort().reverse();
    }
    var anos = ANOS_MENSAL.slice();
    [R.anoAtual(), ANO_ATIVO].forEach(function (a) {
      if (anos.indexOf(a) === -1) anos.push(a);
    });
    return anos;
  }

  function anoSemColunas(ano) { return !ANOS_PLANILHAS && ANOS_MENSAL.indexOf(ano) === -1; }

  function rotuloAno(ano) {
    return anoSemColunas(ano) ? ano + ' — em preparação' : ano;
  }

  // Repopula TODOS os seletores de ano (barra superior + um por página) e
  // sincroniza o valor exibido. Chamar depois de qualquer innerHTML de página.
  function renderSeletoresAno() {
    var opts = opcoesAno().map(function (a) {
      return '<option value="' + escapeHtml(a) + '">' + escapeHtml(rotuloAno(a)) + '</option>';
    }).join('');
    document.querySelectorAll('select.ano-select').forEach(function (sel) {
      if (sel.innerHTML !== opts) sel.innerHTML = opts;
      if (sel.value !== ANO_ATIVO) sel.value = ANO_ATIVO;
    });
  }

  function trocarAno(ano) {
    if (!ano || ano === ANO_ATIVO) return;
    if (ANOS_PLANILHAS) {
      trocarPlanilhaDeAno(ano);
      return;
    }
    ANO_ATIVO = ano;
    renderSeletoresAno();
    renderCurrentPage();
  }

  // V1.6.0 — trocar de exercício no modo "planilha por ano": cada ano tem o seu
  // arquivo (já carregado em memória), então a troca re-aponta o ExcelIO para o
  // outro substate e a app recarrega o estado daquele exercício. Só pergunta
  // antes se houver alterações não salvas no ano que está sendo abandonado.
  function trocarPlanilhaDeAno(ano) {
    if (DIRTY && !confirm('Existem alterações não salvas em ' + ANO_ATIVO + '. Trocar de exercício vai descartá-las. Continuar?')) {
      renderSeletoresAno();
      return;
    }
    // Modo base: descartar de verdade (relê o ano), senão a edição voltaria no próximo Salvar.
    var descartar = (DIRTY && MODO_BASE) ? ExcelIO.descartarAlteracoes(ANO_ATIVO) : Promise.resolve();
    descartar.then(function () { return ExcelIO.ativarAno(ano); }).then(function (res) {
      ANO_ATIVO = String(ano);
      loadPlanilhaResult(res);
      toast('Exercício ' + ANO_ATIVO + ': ' + res.fileName + ' (' + res.contratos.length + ' contratos)');
    }).catch(function (err) {
      console.error(err);
      renderSeletoresAno();
      toast(err.message || 'Não foi possível abrir a planilha de ' + ano + '.', true);
    });
  }

  function wireSeletorAno(sel) {
    sel.addEventListener('change', function (e) { trocarAno(e.target.value); });
  }

  function barraAnoPagina() {
    return '<div class="ano-bar"><span class="ano-bar-label">Exercício</span>' +
      '<select class="ano-select" aria-label="Exercício"></select>' +
      '<span class="ano-bar-nota">' +
      (anoSemColunas(ANO_ATIVO)
        ? 'A planilha ainda não tem as colunas de ' + ANO_ATIVO + ' — este exercício aparece zerado até você clicar em "+ Inserir ano".'
        : 'Todos os números da página são de ' + ANO_ATIVO + notaPosicao()) +
      '</span></div>';
  }

  // Modo base: a posição da planilha (a mesma do cabeçalho do Power BI) vira RETRATO
  // (V1.9.2): o status em tela já é "conforme hoje"; a nota informa até quando a
  // planilha foi recalculada e avisa quando ela está atrasada em relação a hoje.
  function notaPosicao() {
    var pos = posicaoAtiva();
    if (!MODO_BASE || !pos) return '.';
    var velha = ANO_ATIVO === R.anoAtual() && R.startOfDay(refPlanilha()) < R.startOfDay(HOJE);
    var seHoje = ANO_ATIVO === R.anoAtual();
    return ' · posição ' + pos + (seHoje ? ' (status calculado com hoje)' : ' (exercício fechado, congelado)') +
      (velha ? ' — abra a planilha no Excel e salve para a posição do Power BI chegar a hoje.' : '.');
  }

  function renderCurrentPage() {
    // A Home é apresentação: renderiza com ou sem planilha aberta.
    if (CURRENT_PAGE === 'home') {
      renderHome();
      return;
    }
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
      '<h3>Nenhuma planilha carregada</h3>' +
      '<p>Abra a pasta das planilhas do Power BI (<b>MATRIZ DE POWERBI CONTRATOS &lt;ano&gt;.xlsx</b>) — ' +
      'a pasta "Banco de Dados" ou qualquer pasta acima dela.</p>' +
      (IS_SNAPSHOT ? '' : '<button class="btn btn-primary" id="btn-empty-open">Abrir planilha</button>') +
      '</div>';
    var b = document.getElementById('btn-empty-open');
    if (b) b.addEventListener('click', handleOpenFolder);
  }

  // ---------------------------------------------------------------
  // HOME — página de apresentação (V1.8.0). Fundo da empresa + título
  // + botões de navegação para as telas. Mesmo shell do app; renderiza
  // sempre, com ou sem planilha aberta.
  // ---------------------------------------------------------------
  function renderHome() {
    var root = document.getElementById('page-home');
    if (!root) return;
    root.innerHTML =
      '<div class="home-page">' +
        '<div class="home-page-bg"></div>' +
        '<div class="home-page-scrim"></div>' +
        '<div class="home-page-content">' +
          '<h1 class="home-title">Contratos</h1>' +
          '<p class="home-subtitle">Gestão de Contratos - Inventus Power</p>' +
          '<div class="home-navbtns">' +
            '<button class="btn-home-nav" data-page="executiva">Visão Executiva <span class="arrow">→</span></button>' +
            '<button class="btn-home-nav" data-page="sla">Status do contrato <span class="arrow">→</span></button>' +
            '<button class="btn-home-nav" data-page="financeira">Financeiro <span class="arrow">→</span></button>' +
            '<button class="btn-home-nav" data-page="gestao">Gestão de Contratos <span class="arrow">→</span></button>' +
          '</div>' +
        '</div>' +
      '</div>';
    root.querySelectorAll('.btn-home-nav').forEach(function (btn) {
      btn.addEventListener('click', function () {
        goToPage(btn.getAttribute('data-page'));
      });
    });
  }

  // ==============================================================
  // ABRIR / SALVAR
  // ==============================================================
  function handleOpenFolder() {
    if (IS_SNAPSHOT) return;
    if (DIRTY && !confirm('Existem alterações não salvas que serão perdidas. Continuar?')) return;
    toast(window.DesktopBridge
      ? 'Selecione o arquivo MATRIZ DE CONTRATOS <ano>.xlsx…'
      : 'Selecione a pasta das planilhas (Banco de Dados, ou uma pasta acima dela)…');
    ExcelIO.openFolder().then(function (result) {
      loadPlanilhaResult(result);
      toast('Planilha carregada: ' + result.fileName + ' (' + result.contratos.length + ' contratos)');
    }).catch(function (err) {
      console.error(err);
      toast(err.message || 'Não foi possível abrir a planilha.', true);
    });
  }

  function loadPlanilhaResult(result) {
    CONTRATOS = SIMULAR ? simularDados(result.contratos) : result.contratos;
    COTACOES = SIMULAR ? simularCotacoes(result.cotacoes) : (result.cotacoes || {});
    // Saving: o da planilha; na demonstração sem saving real, itens sintéticos (nunca gravados).
    SAVING = (result.saving && result.saving.length) ? result.saving : (SIMULAR ? simularSavingItens() : []);
    TEM_SAVING = !!result.temSaving || SIMULAR;
    DELETED_SAVING0 = [];
    ANOS_MENSAL = (result.anosMensais && result.anosMensais.length)
      ? result.anosMensais.slice()
      : Regras.ANOS_MENSAIS.slice();
    ANOS_PLANILHAS = result.anosPlanilhas || null;
    MODO_BASE = !!result.modoBase;
    ANO_ATIVO = result.anoAtivo || anoPadrao();
    DELETED_ROWS0 = [];
    DIRTY = false;
    renderSeletoresAno();
    updateFileStatus();
    renderCurrentPage();
  }

  // Dados sintéticos para visualização (SIMULAR): sorteio determinístico por
  // contrato — cada reabertura gera os MESMOS números. Preenche só o que a
  // planilha real deixou em branco (valor, departamento, tipo, vigência,
  // meses de 2025/2026) para a demonstração sair completa nas 5 telas sem
  // sobrescrever o que a Hendrya já preencheu. Nunca grava no arquivo — ver
  // guard em handleSave.
  // Nomes fictícios para o boot autocontido de ?mock=1 (sem planilha real
  // aberta). Todo campo além de id/razaoSocial fica vazio de propósito —
  // simularDados() é quem preenche o resto.
  var SIMULAR_EMPRESAS = [
    'Alfa Facilities Ltda', 'Bravo Segurança Patrimonial S.A.', 'Cedro Manutenção Predial Ltda',
    'Delta Consultoria Jurídica Ltda', 'Épsilon TI Serviços S.A.', 'Fortaleza Telecom Ltda',
    'Gaivota Transportes e Logística Ltda', 'Horizonte Seguros S.A.', 'Ipê Marketing e Publicidade Ltda',
    'Jataí Engenharia Ltda', 'Kalunga Auditoria Externa S.A.', 'Lótus Recursos Humanos Ltda',
    'Meridiano Limpeza e Conservação Ltda', 'Nortel Facilities S.A.', 'Ouro Verde Segurança Ltda',
    'Pampa Manutenção Industrial Ltda', 'Quality Consultoria Ltda', 'Rio Claro Telecom S.A.',
    'Sol Nascente Transportes Ltda', 'Trópico Seguros Ltda', 'União Marketing Ltda',
    'Vetor Engenharia S.A.', 'Wanderley Auditoria Ltda', 'Xingu RH Serviços Ltda',
    'Ypê Limpeza Industrial Ltda', 'Zênite Facilities Ltda', 'Amazônia TI Suporte S.A.',
    'Boreal Jurídico Associados Ltda', 'Cristal Telecom Ltda', 'Diamante Transportes S.A.',
    'Estrela do Sul Seguros Ltda', 'Fênix Marketing Ltda', 'Girassol Engenharia Ltda',
    'Harmonia Auditoria S.A.', 'Íris RH Ltda', 'Jequitibá Limpeza Ltda', 'Kiara Facilities S.A.',
    'Luar Segurança Ltda', 'Montanha Manutenção Ltda', 'Nascente Consultoria Ltda',
    'Oceano TI Serviços Ltda', 'Prisma Telecom S.A.', 'Quero-Quero Transportes Ltda',
    'Raiz Seguros Ltda', 'Safira Marketing S.A.', 'Trevo Engenharia Ltda',
    'Universo Auditoria Ltda', 'Vale Verde RH S.A.', 'Wave Limpeza Ltda', 'Zaffira Facilities Ltda'
  ];

  function gerarContratosSimulados() {
    return SIMULAR_EMPRESAS.map(function (nome, i) {
      return {
        id: i + 1, isNew: false,
        razaoSocial: nome, tipoServico: '', departamento: '',
        prazoPagamentoRaw: '', prazoPagamentoDias: null, prazoNegociadoRaw: '',
        valorMensal2026: null, valorAnual2026: null, vigencia: '',
        dataInicio: null, dataFim: null, comentarios: '',
        statusManual: null, statusLegadoRaw: ''
      };
    });
  }

  // Saving sintético (?mock=1 sem aba Saving real): mesmo formato dos itens da planilha —
  // saving constante em R$ a partir do mês em que o BID fechou, alguns com total sem mês.
  var BID_TIPOS = ['Refeição', 'Conservação e Limpeza', 'Calibração', 'Manutenção preventiva',
    'SGI CCTV', 'Copeira', 'Transporte de funcionários', 'Telecom'];
  var BID_PAGAMENTOS = ['30 dias', '15 dias', 'de 10 dias P/ 45 dias', 'de 30 dias P/ 45 dias',
    'de 30 dias P/ 40 dias', 'de 10 dias P/ 15 dias'];
  function simularSavingItens() {
    return SIMULAR_EMPRESAS.slice(1, 15).map(function (nome, i) {
      var it = {
        id: NEXT_TEMP_SAV_ID--, isNew: true, _row0: null, codigoFornecedor: '', prestador: nome,
        tipoServico: BID_TIPOS[i % BID_TIPOS.length], bidConcluido: i % 4 === 3 ? 'Não' : 'Sim',
        pagamento: BID_PAGAMENTOS[i % BID_PAGAMENTOS.length], savSemMes: i % 5 === 0 ? 5000 + i * 731 : null
      };
      var inicio = (i * 5) % 12, valor = 800 + ((i * 379) % 9500);
      R.MESES.forEach(function (m, k) { it[R.campoSaving(m.key)] = (i % 4 !== 3 && k >= inicio) ? valor : null; });
      return it;
    });
  }

  var SIMULAR_DEPTOS = ['Compras', 'Financeiro', 'Facilities', 'TI', 'RH', 'Manutenção', 'Jurídico', 'Marketing', 'Operações', 'Qualidade'];
  var SIMULAR_TIPOS = ['Manutenção Predial', 'Segurança Patrimonial', 'Limpeza e Conservação', 'Consultoria Jurídica', 'TI - Suporte', 'Telecomunicações', 'Transporte e Logística', 'Seguros', 'Marketing e Publicidade', 'Serviços de Engenharia', 'Auditoria Externa', 'Recursos Humanos'];
  var SIMULAR_PRAZOS = [15, 30, 30, 45, 60];
  var SIMULAR_COMENTARIOS = {
    'No Prazo': ['Renovação automática prevista.', 'Sem pendências no momento.', 'Acompanhamento trimestral com o fornecedor.'],
    'A Vencer': ['Aguardando retorno do fornecedor sobre renovação.', 'Negociação de reajuste em andamento.'],
    'Vencido': ['Pendente de renovação — cobrar fornecedor.', 'Em análise jurídica para renovação.'],
    'Em Análise': ['Aguardando definição de vigência com o fornecedor.', 'Novo contrato, cláusulas em revisão.']
  };

  function simularDados(contratos) {
    function seedHash(s) {
      var h = 2166136261;
      for (var i = 0; i < s.length; i++) {
        h ^= s.charCodeAt(i);
        h = Math.imul(h, 16777619);
      }
      return h >>> 0;
    }
    function rand(h) {
      h = Math.imul(h ^ (h >>> 15), 2246822519);
      h = Math.imul(h ^ (h >>> 13), 3266489917);
      h ^= h >>> 16;
      return (h >>> 0) / 4294967296;
    }
    function pick(list, h) { return list[Math.floor(rand(h) * list.length)]; }

    contratos.forEach(function (c) {
      var base = seedHash(c.razaoSocial || String(c.id) || 'x');

      if (c.valorAnual2026 == null) c.valorAnual2026 = Math.round(20000 + rand(base) * 980000);
      if (c.valorMensal2026 == null) c.valorMensal2026 = Math.round(c.valorAnual2026 / 12);
      if (!c.departamento) c.departamento = pick(SIMULAR_DEPTOS, base + 1);
      if (!c.tipoServico) c.tipoServico = pick(SIMULAR_TIPOS, base + 2);
      if (!c.prazoPagamentoDias) c.prazoPagamentoDias = pick(SIMULAR_PRAZOS, base + 3);

      if (!c.dataFim) {
        var bucket = rand(base + 4);
        var dias = null;
        if (bucket < 0.40) dias = Math.round(60 + rand(base + 5) * 700); // No Prazo
        else if (bucket < 0.42) dias = Math.round(5 + rand(base + 5) * 25); // A Vencer
        else if (bucket < 0.80) dias = -Math.round(5 + rand(base + 5) * 500); // Vencido
        // senão fica sem data (Em Análise) — status legítimo, não lacuna
        if (dias != null) c.dataFim = R.addDays(HOJE, dias);
      }
      if (!c.dataInicio && c.dataFim) {
        c.dataInicio = R.addDays(c.dataFim, -Math.round(365 + rand(base + 6) * 730));
      }
      if (!c.comentarios) {
        var statusAtual = R.statusFinal(c, HOJE);
        var pool = SIMULAR_COMENTARIOS[statusAtual] || SIMULAR_COMENTARIOS['No Prazo'];
        c.comentarios = pick(pool, base + 8);
      }

      var anualBase = c.valorAnual2026;
      ['2025', '2026'].forEach(function (ano) {
        R.MESES.forEach(function (m) {
          var f = R.fieldBrlAno(ano, m.key);
          if (c[f] == null) {
            var s = (base + seedHash(ano + m.key)) >>> 0;
            c[f] = Math.round((anualBase / 12) * (0.6 + rand(s) * 0.9));
          }
        });
      });
    });
    return contratos;
  }

  // Cotação sintética do dólar (SIMULAR): preenche os meses de 2025/2026 que a
  // planilha real deixou vazios, para os valores em US$ aparecerem em todas as
  // telas. Deriva de uma taxa-base com pequena variação mensal determinística.
  function simularCotacoes(cotacoes) {
    cotacoes = cotacoes || {};
    ['2025', '2026'].forEach(function (ano) {
      cotacoes[ano] = cotacoes[ano] || {};
      R.MESES.forEach(function (m, i) {
        if (cotacoes[ano][m.key] == null) {
          var base = (ano === '2026' ? 5.05 : 4.95) + i * 0.015;
          cotacoes[ano][m.key] = Math.round(base * 100) / 100;
        }
      });
    });
    return cotacoes;
  }

  // No modo desktop não existe dirHandle (o app opera por caminho de arquivo);
  // exigir só o dirHandle deixava o botão Salvar do .exe sem efeito.
  function planilhaAberta() {
    var st = ExcelIO.getState();
    return !!(st.dirHandle || st.filePath);
  }

  function handleSave() {
    if (IS_SNAPSHOT) return;
    if (SIMULAR) {
      toast('Modo demonstração (dados simulados) — Salvar fica desativado para não gravar valores fictícios na planilha.', true);
      return;
    }
    if (!planilhaAberta()) {
      toast('Abra a planilha antes de salvar.', true);
      return;
    }
    var btn = document.getElementById('btn-save');
    if (btn) btn.disabled = true;
    var res = null;
    var extra = TEM_SAVING ? { saving: SAVING, savingRemovidos: DELETED_SAVING0 } : null;
    ExcelIO.save(CONTRATOS, DELETED_ROWS0, extra).then(function (r) {
      res = r || null;
      DELETED_ROWS0 = [];
      DELETED_SAVING0 = [];
      DIRTY = false;
      updateFileStatus();
      renderCurrentPage();
      if (res && res.gravado === false) return null;
      toast('Planilha atualizada. Gerando snapshot…');
      return generateAndSaveSnapshot();
    }).then(function (name) {
      if (res && res.gravado === false) {
        toast('Nada mudou desde a última gravação — a planilha não foi alterada.');
      } else if (MODO_BASE) {
        // O Power BI lê o que o Excel calculou: sem abrir e salvar no Excel, ele fica com os números de antes.
        toast('Salvo em ' + res.arquivo + ' (' + name + ' também). Para o Power BI ver: abra a planilha no Excel, salve e atualize o Power BI — até lá ele avisa em vermelho.');
      } else {
        toast('Salvo — planilha e ' + name + ' atualizados.');
      }
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
    var txt;
    if (SIMULAR) {
      txt = 'Demonstração' + (st.fileName ? ' — ' + st.fileName : '') + ' (dados simulados)';
    } else if (!st.fileName) {
      txt = 'Nenhuma planilha aberta';
    } else {
      var anos = ANOS_PLANILHAS ? Object.keys(ANOS_PLANILHAS).sort() : [];
      txt = MODO_BASE
        ? ('Base única — exercício ' + ANO_ATIVO)
        : (anos.length ? ('Planilhas ' + anos.join(' + ')) : ('Planilha ' + ANO_ATIVO));
      if (nameEl) nameEl.title = st.fileName;
    }
    if (nameEl) nameEl.textContent = txt;
    if (dotEl) dotEl.className = 'status-dot ' + (DIRTY ? 'dirty' : ((st.fileName || SIMULAR) ? 'on' : 'off'));
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
  // Modo base: os meses dos outros exercícios (cada um no seu arquivo), para o
  // "Dispêndio Mensal Acumulado" do ano anterior sair também no snapshot.
  function outrosAnosPayload() {
    var out = {};
    Object.keys(ANOS_PLANILHAS || {}).forEach(function (ano) {
      if (ano === ANO_ATIVO) return;
      out[ano] = contratosDoAno(ano).map(function (c) {
        var o = { razaoSocial: c.razaoSocial };
        R.mesFieldsAno(ano).brl.forEach(function (f) { o[f] = c[f]; });
        return o;
      });
    });
    return out;
  }

  function snapshotPayload() {
    var base = MODO_BASE ? {
      modoBase: true, anoAtivo: ANO_ATIVO, posicao: posicaoAtiva(), outrosAnos: outrosAnosPayload(),
      temSaving: TEM_SAVING, saving: SAVING.map(function (it) { return Object.assign({}, it); }),
      referencia: refAtiva() ? refAtiva().toISOString() : null
    } : {};
    return Object.assign(base, {
      generatedAt: new Date().toISOString(),
      sourceFile: ExcelIO.getState().fileName,
      anosMensais: ANOS_MENSAL.slice(),
      cotacoes: COTACOES,
      contratos: CONTRATOS.map(function (c) {
        var out = {
          id: c.id, codigoFornecedor: c.codigoFornecedor, nomeCurto: c.nomeCurto,
          razaoSocial: c.razaoSocial, tipoServico: c.tipoServico, departamento: c.departamento,
          prazoPagamentoRaw: c.prazoPagamentoRaw, prazoPagamentoDias: c.prazoPagamentoDias,
          prazoNegociadoRaw: c.prazoNegociadoRaw, valorMensal2026: c.valorMensal2026, valorAnual2026: c.valorAnual2026,
          vigencia: c.vigencia, dataInicio: c.dataInicio ? c.dataInicio.toISOString() : null,
          dataFim: c.dataFim ? c.dataFim.toISOString() : null, comentarios: c.comentarios,
          statusManual: c.statusManual, statusLegadoRaw: c.statusLegadoRaw
        };
        ANOS_MENSAL.forEach(function (ano) {
          R.mesFieldsAno(ano).brl.concat(R.mesFieldsAno(ano).usd).forEach(function (f) { out[f] = c[f]; });
          out[R.fieldTotalAno(ano)] = c[R.fieldTotalAno(ano)];
          out[R.fieldTotalUsdAno(ano)] = c[R.fieldTotalUsdAno(ano)];
        });
        return out;
      })
    });
  }

  function generateAndSaveSnapshot() {
    return ExcelIO.saveSnapshotHtml(buildSnapshotHtml(snapshotPayload()));
  }

  // "Exportar HTML do dia": mesmo snapshot standalone, com a data no nome, para
  // mandar à diretoria sem sobrescrever o snapshot corrente. Grava ao lado da
  // planilha (desktop e navegador); sem planilha aberta, cai para download.
  function handleExportarHtmlDoDia() {
    if (IS_SNAPSHOT) return;
    if (!CONTRATOS.length) {
      toast('Abra a planilha antes de exportar.', true);
      return;
    }
    var btn = document.getElementById('btn-export');
    if (btn) btn.disabled = true;
    var nome = 'Dashboard_Contratos_' + R.toDateInputValue(new Date()) + '.html';
    var html = buildSnapshotHtml(snapshotPayload());
    var acao = planilhaAberta()
      ? ExcelIO.saveSnapshotHtml(html, nome)
      : Promise.reject(new Error('sem pasta'));
    acao.then(function (name) {
      toast('HTML do dia gerado: ' + name);
    }).catch(function () {
      downloadHtml(nome, html);
      toast('HTML do dia baixado: ' + nome);
    }).finally(function () {
      if (btn) btn.disabled = false;
    });
  }

  function downloadHtml(fileName, html) {
    var url = URL.createObjectURL(new Blob([html], { type: 'text/html;charset=utf-8' }));
    var a = document.createElement('a');
    a.href = url;
    a.download = fileName;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(function () { URL.revokeObjectURL(url); }, 1000);
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

  // Escreve o total no vazio do anel — plugin local (não registrado global) para
  // não afetar nenhum outro gráfico.
  var donutCenterPlugin = {
    id: 'donutCenter',
    afterDatasetsDraw: function (chart, args, opts) {
      var meta = chart.getDatasetMeta(0);
      if (!meta || !meta.data.length) return;
      var c = chart.ctx;
      var x = meta.data[0].x, y = meta.data[0].y;
      c.save();
      c.textAlign = 'center';
      c.textBaseline = 'middle';
      c.fillStyle = '#2B2B2B';
      c.font = '700 30px "Segoe UI", sans-serif';
      c.fillText(opts.total, x, y - 9);
      c.fillStyle = '#7a7f87';
      c.font = '600 11px "Segoe UI", sans-serif';
      c.fillText(opts.label, x, y + 14);
      c.restore();
    }
  };

  function drawDonut(canvasId, labels, values, colors, centro) {
    destroyChart(canvasId);
    var ctx = document.getElementById(canvasId);
    if (!ctx) return;
    CHART_INSTANCES[canvasId] = new Chart(ctx, {
      type: 'doughnut',
      data: { labels: labels, datasets: [{ data: values, backgroundColor: colors, borderWidth: 2, borderColor: '#fff' }] },
      plugins: centro ? [donutCenterPlugin] : [],
      options: {
        maintainAspectRatio: false,
        cutout: '62%',
        plugins: {
          donutCenter: centro || {},
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
        // o datalabel na ponta da barra horizontal fica fora da área do gráfico:
        // sem esta folga à direita ele é cortado pela borda do card.
        layout: { padding: { right: (opts.horizontal && opts.dataLabels) ? 58 : 0 } },
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
  // V1.9.0 — "BIDs e negociações" (aba Saving do exercício): uma linha por item, na ordem em
  // que o saving começou (sem mês por último). Mesma régua da planilha (Regras.saving*).
  function bidsBlockHtml() {
    var ano = ANO_ATIVO;
    var ordem = function (it) { var d = R.savingDesde(it); var i = R.MESES.map(function (m) { return m.label; }).indexOf(d); return i < 0 ? 99 : i; };
    var tu = 0, tr = 0;
    var rows = SAVING.slice().sort(function (a, b) {
      return ordem(a) - ordem(b) || String(a.prestador).localeCompare(String(b.prestador), 'pt-BR');
    }).map(function (it) {
      var u = R.savingUSD(it, ano, COTACOES), r = R.savingBRL(it);
      tu += u; tr += r;
      return '<tr>' +
        '<td>' + escapeHtml(it.prestador || '') + '</td>' +
        '<td>' + escapeHtml(it.tipoServico || '') + '</td>' +
        '<td>' + escapeHtml(it.pagamento || '') + '</td>' +
        '<td>' + escapeHtml(R.savingDesde(it) || (Number(it.savSemMes) ? 'sem mês' : '—')) + '</td>' +
        '<td class="cell-num">' + R.formatUSD(u) + '</td>' +
        '<td class="cell-num">' + R.formatBRL(r) + '</td>' +
        '</tr>';
    }).join('');
    if (!rows) {
      return '<div class="bids-empty">' + (TEM_SAVING
        ? 'Nenhum BID ou negociação lançado em ' + ano + ' (aba Saving da planilha ou tela Gestão de Contratos).'
        : 'A planilha aberta não tem a aba Saving.') + '</div>';
    }
    return '<table class="data-table bids-table"><thead><tr>' +
      '<th>Nome do Prestador</th><th>Tipo de Serviço</th><th>Pagamento</th><th>Desde</th>' +
      '<th class="th-num">Saving US$</th><th class="th-num">Saving R$</th>' +
      '</tr></thead><tbody>' + rows + '</tbody>' +
      '<tfoot><tr class="bids-tfoot"><td colspan="4">Total do ano (' + ano + ')</td>' +
      '<td class="cell-num">' + R.formatUSD(tu) + '</td><td class="cell-num">' + R.formatBRL(tr) + '</td>' +
      '</tr></tfoot>' +
      '</table>';
  }

  // Saving por mês do exercício (R$) + o total sem detalhe mensal, que não entra no gráfico.
  function savingMensal() {
    var vals = R.MESES.map(function (m) {
      return SAVING.reduce(function (t, it) { return t + (R.savingMes(it, m.key) || 0); }, 0);
    });
    var semMes = SAVING.reduce(function (t, it) { return t + (Number(it.savSemMes) || 0); }, 0);
    return { meses: vals, semMes: semMes };
  }

  function renderExecutiva() {
    var root = document.getElementById('page-executiva');
    var contratos = filteredContratos();
    var k = R.computeKPIs(contratos, refAtiva(), COTACOES, ANO_ATIVO);

    var sav = savingMensal();
    var saveMensal = sav.meses;
    var saveTotal = saveMensal.reduce(function (a, b) { return a + b; }, 0) + sav.semMes;
    root.innerHTML =
      barraAnoPagina() +
      '<div class="kpi-grid">' +
      kpiCard('Total de Contratos', R.formatNumberBR(k.totalContratos), 'primary',
        '= No Prazo + A Vencer + Vencidos + Em Análise + Distrato/Expirado') +
      kpiCard('No Prazo', R.formatNumberBR(k.contratosNoPrazo), 'no-prazo') +
      kpiCard('A Vencer (30d)', R.formatNumberBR(k.contratosAVencer), 'a-vencer') +
      kpiCard('Vencidos', R.formatNumberBR(k.contratosVencidos), 'vencido') +
      kpiCard('Em Análise', R.formatNumberBR(k.contratosEmAnalise), 'em-analise') +
      '</div>' +
      '<div class="kpi-grid" style="grid-template-columns:repeat(auto-fit,minmax(220px,1fr));margin-bottom:22px;">' +
      kpiCard('Despesa Total Anual', R.formatBRL(k.despesaTotalAnual, true), 'primary', 'Valor negociado: Anual, ou Mensal × 12') +
      kpiCard('Despesa Total Anual em Dólar', R.formatUSD(k.despesaTotalAnualUSD, true), 'primary', 'Soma do que foi faturado em US$ no ano') +
      kpiCard('Contratos com Prazo de Pagamento Negociado', R.formatNumberBR(k.contratosPrazoPagamentoNegociado), 'no-prazo',
        (k.totalContratos ? R.formatPercent(k.contratosPrazoPagamentoNegociado / k.totalContratos) + ' dos contratos' : '')) +
      kpiCard('Distrato/Expirado', R.formatNumberBR(k.contratosDistrato), 'distrato', 'Encerrados ou vencidos sem renovação') +
      '</div>' +
      '<div class="grid grid-2">' +
      '<div class="card"><h3>Status dos Contratos</h3><div class="card-sub">Distribuição por status calculado</div>' +
      '<div class="chart-wrap"><canvas id="chart-exec-donut"></canvas></div>' +
      '</div>' +
      '<div class="card"><h3>Top 10 Contratos por Dispêndio</h3><div class="card-sub">Soma do dispêndio mensal de ' + ANO_ATIVO + '</div>' +
      '<div class="chart-wrap lg"><canvas id="chart-exec-contratos"></canvas></div>' +
      '</div>' +
      '<div class="card"><h3>BIDs e negociações</h3><div class="card-sub">Aba Saving de ' + ANO_ATIVO +
      ' · US$ pela cotação de cada mês</div>' +
      bidsBlockHtml() +
      '</div>' +
      '<div class="card"><h3>Saving mensal</h3><div class="card-sub">Total ' + ANO_ATIVO + ': ' + R.formatBRL(saveTotal, true) +
      (sav.semMes ? ' (' + R.formatBRL(sav.semMes, true) + ' sem detalhe mensal, fora do gráfico)' : '') + '</div>' +
      '<div class="chart-wrap"><canvas id="chart-exec-saving"></canvas></div>' +
      '</div>' +
      '</div>';

    renderSeletoresAno();
    document.querySelectorAll('#page-executiva select.ano-select').forEach(wireSeletorAno);

    drawDonut('chart-exec-donut', R.STATUS_ORDER,
      [k.contratosNoPrazo, k.contratosAVencer, k.contratosVencidos, k.contratosEmAnalise, k.contratosDistrato],
      R.STATUS_ORDER.map(function (s) { return R.STATUS_COLORS[s]; }),
      { total: R.formatNumberBR(k.totalContratos), label: 'Total de Contratos' });

    var top = topContratos(contratos, 10, 'dispendio');
    drawBar('chart-exec-contratos', top.map(function (c) { return rotuloCurtoContrato(c.razaoSocial); }),
      top.map(function (c) { return c.dispendio; }),
      { horizontal: true, dataLabels: true, labelFormatter: function (v) { return R.formatBRL(v, true); } });

    // V1.8.1 — Saving mensal (colunas horizontais JAN..DEZ), abaixo do Top 10.
    drawBar('chart-exec-saving', R.MESES.map(function (m) { return m.label; }),
      saveMensal,
      { dataLabels: true, labelFormatter: function (v) { return R.formatBRL(v, true); } });
  }

  // ==============================================================
  // PÁGINA: STATUS DO CONTRATO
  // ==============================================================
  function renderSLA() {
    var root = document.getElementById('page-sla');
    var contratos = filteredContratos();
    var k = R.computeKPIs(contratos, refAtiva(), COTACOES, ANO_ATIVO);
    var deptos = distinctDepartamentos(contratos);

    root.innerHTML =
      barraAnoPagina() +
      '<div class="kpi-grid">' +
      kpiCard('A Vencer (30 dias)', R.formatNumberBR(k.contratosAVencer), 'a-vencer') +
      kpiCard('Vencidos', R.formatNumberBR(k.contratosVencidos), 'vencido') +
      kpiCard('Em Análise (Comprador)', R.formatNumberBR(k.contratosEmAnalise), 'em-analise') +
      kpiCard('Distrato/Expirado', R.formatNumberBR(k.contratosDistrato), 'distrato') +
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
      '<th>Fornecedor</th><th>Departamento</th><th>Status</th><th>Data Fim</th><th>Dias até vencer</th><th class="cell-num">Dispêndio ' + ANO_ATIVO + ' R$</th>' +
      '</tr></thead><tbody id="sla-tbody"></tbody></table></div></div>' +
      '<div class="footer-note">Régua: Vencimento Indefinido → Em Análise · vencido se Fim ≤ ' + (MODO_BASE ? 'data de posição' : 'hoje') +
      ' · a vencer se Fim ≤ ' + (MODO_BASE ? 'data de posição' : 'hoje') + ' + 30 dias.</div>';

    renderSeletoresAno();
    document.querySelectorAll('#page-sla select.ano-select').forEach(wireSeletorAno);
    document.querySelectorAll('#sla-status-pills button').forEach(function (btn) {
      btn.addEventListener('click', function () { SLA_FILTERS.status = btn.getAttribute('data-status'); renderSLA(); });
    });
    document.getElementById('sla-depto-filter').addEventListener('change', function (e) {
      SLA_FILTERS.departamento = e.target.value; renderSLA();
    });

    var rows = contratos.filter(function (c) {
      var st = R.statusFinal(c, refAtiva());
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
      var st = R.statusFinal(c, refAtiva());
      var dias = c.dataFim ? Math.round((R.startOfDay(c.dataFim) - R.startOfDay(refAtiva())) / 86400000) : null;
      return '<tr>' +
        '<td>' + escapeHtml(c.razaoSocial) + '</td>' +
        '<td>' + escapeHtml(c.departamento) + '</td>' +
        '<td>' + statusBadge(st) + '</td>' +
        '<td>' + R.formatDateBR(c.dataFim) + '</td>' +
        '<td class="cell-num">' + (dias == null ? '—' : dias) + '</td>' +
        '<td class="cell-num">' + R.formatBRL(R.despesaAno(c, ANO_ATIVO)) + '</td>' +
        '</tr>';
    }).join('') || '<tr><td colspan="6" class="cell-muted" style="text-align:center;padding:24px;">Nenhum contrato encontrado.</td></tr>';
  }

  // ==============================================================
  // PÁGINA: FINANCEIRA
  // ==============================================================
  function renderFinanceira() {
    var root = document.getElementById('page-financeira');
    var contratos = filteredContratos();
    var k = R.computeKPIs(contratos, refAtiva(), COTACOES, ANO_ATIVO);

    var topPorDespesa = topContratos(contratos, 10, 'despesa');
    var topPorDispendio = topContratos(contratos, 10, 'dispendio');
    var anoCorrente = /^\d{4}$/.test(R.anoAtual()) ? R.anoAtual() : (ANOS_MENSAL[0] || '2026');
    var anoPassado = String(Number(anoCorrente) - 1);
    var seriePassado = serieDispendioAcumulado(contratosDoAno(anoPassado), anoPassado);
    var serieAnoCorrente = serieDispendioAcumulado(contratosDoAno(anoCorrente), anoCorrente);

    root.innerHTML =
      barraAnoPagina() +
      '<div class="filter-bar" style="margin-bottom:6px;">' +
      '<span style="font-size:12px;color:var(--muted);">' + (SIMULAR ? '<b style="color:#DC2626;">[DADOS SIMULADOS] </b>' : '') + 'Demonstrativo de ' + ANO_ATIVO + '.</span>' +
      '<div class="spacer"></div>' +
      '<span class="filter-count">' + (MODO_BASE ? 'Cotação lida da aba Cotações da planilha' : 'Cotação lida da linha 2 da planilha') +
      ' — atualize no Excel.</span>' +
      '</div>' +
      cotacoesBar(ANO_ATIVO) +
      '<div class="kpi-grid">' +
      kpiCard('Valor Total dos Contratos', R.formatBRL(k.despesaTotalAnual, true), 'primary', 'Valor negociado: Anual, ou Mensal × 12') +
      kpiCard('Despesa Total Anual em Dólar', R.formatUSD(k.despesaTotalAnualUSD, true), 'primary', 'Faturado em US$ no ano') +
      kpiCard('Total Faturado ' + ANO_ATIVO, R.formatBRL(k.despesaTotalAno, true), 'primary', 'Soma dos 12 meses lançados') +
      '</div>' +
      '<div class="grid grid-2">' +
      '<div class="card"><h3>Top 10 Contratos por Despesa</h3><div class="card-sub">Maior despesa de ' + ANO_ATIVO + ' (total faturado no ano)</div>' +
      '<div class="table-wrap"><div class="table-scroll" style="max-height:320px;">' +
      '<table class="data-table"><thead><tr><th>Fornecedor</th><th class="cell-num">Despesa ' + ANO_ATIVO + '</th>' +
      '<th class="cell-num">Valor do Contrato</th><th>Departamento</th></tr></thead>' +
      '<tbody id="fin-top10-tbody"></tbody></table></div></div></div>' +
      '<div class="card"><h3>Top 10 Contratos por Dispêndio</h3><div class="card-sub">Soma do dispêndio mensal de ' + ANO_ATIVO + '</div>' +
      '<div class="chart-wrap"><canvas id="chart-fin-contratos"></canvas></div></div>' +
      '</div>' +
      '<div class="grid grid-2" style="margin-top:18px;">' +
      '<div class="card"><h3>Dispêndio Mensal Acumulado ' + anoPassado + '</h3><div class="card-sub">Acumulado mês a mês · Total do ano: <b>' + R.formatBRL(seriePassado.total, true) + '</b></div>' +
      '<div class="chart-wrap"><canvas id="chart-fin-disp-passado"></canvas></div></div>' +
      '<div class="card"><h3>Dispêndio Mensal Acumulado ' + anoCorrente + '</h3><div class="card-sub">Acumulado mês a mês · Total do ano: <b>' + R.formatBRL(serieAnoCorrente.total, true) + '</b></div>' +
      '<div class="chart-wrap"><canvas id="chart-fin-disp-atual"></canvas></div></div>' +
      '</div>';

    renderSeletoresAno();
    document.querySelectorAll('#page-financeira select.ano-select').forEach(wireSeletorAno);

    drawBar('chart-fin-contratos', topPorDispendio.map(function (c) { return rotuloCurtoContrato(c.razaoSocial); }),
      topPorDispendio.map(function (c) { return c.dispendio; }),
      { horizontal: true, dataLabels: true, labelFormatter: function (v) { return R.formatBRL(v, true); } });

    drawBar('chart-fin-disp-passado', R.MESES.map(function (m) { return m.label; }), seriePassado.serie,
      { color: '#1F6FB2', dataLabels: false });
    drawBar('chart-fin-disp-atual', R.MESES.map(function (m) { return m.label; }), serieAnoCorrente.serie,
      { color: '#F58220', dataLabels: false });

    document.getElementById('fin-top10-tbody').innerHTML = topPorDespesa.map(function (c) {
      return '<tr><td>' + escapeHtml(c.razaoSocial) + '</td>' +
        '<td class="cell-num">' + R.formatBRL(c.despesaAno) + '</td>' +
        '<td class="cell-num">' + R.formatBRL(c.valor) + '</td>' +
        '<td>' + escapeHtml(c.departamento) + '</td></tr>';
    }).join('') || '<tr><td colspan="4" class="cell-muted" style="text-align:center;padding:24px;">Nenhum contrato com despesa em ' + ANO_ATIVO + '.</td></tr>';
  }

// Top N contratos. `por` escolhe a base da ordenação: 'valor' (valor total do
// contrato), 'dispendio' (soma dos meses do dispêndio no exercício) ou 'despesa'
// (despesa do exercício, com fallbacks). É uma linha por contrato — não
// agrupado por categoria, então não existe "Outros" aqui como existe no R.topN.
function topContratos(contratos, n, por) {
  return contratos.map(function (c) {
    return {
      razaoSocial: c.razaoSocial,
      departamento: c.departamento,
      valor: R.valorTotalContrato(c, ANO_ATIVO),
      despesaAno: R.despesaAno(c, ANO_ATIVO),
      dispendio: R.somaMesesBRL(c, ANO_ATIVO)
    };
  }).filter(function (c) {
    return (por === 'valor' ? c.valor : (por === 'dispendio' ? c.dispendio : c.despesaAno)) != null;
  }).sort(function (a, b) {
    if (por === 'valor') return b.valor - a.valor;
    if (por === 'dispendio') return (b.dispendio || 0) - (a.dispendio || 0);
    return (b.despesaAno || 0) - (a.despesaAno || 0);
  }).slice(0, n);
}

// Série acumulada mês a mês do dispêndio (soma dos meses) de um ano inteiro,
// mais o total do ano. `serie[i]` é o acumulado até o mês i (JAN..DEZ).
function serieDispendioAcumulado(contratos, ano) {
  var total = 0;
  var serie = R.MESES.map(function (m) {
    var soma = 0;
    contratos.forEach(function (c) {
      var v = c[R.fieldBrlAno(ano, m.key)];
      if (v != null) soma += v;
    });
    total += soma;
    return total;
  });
  return { serie: serie, total: total };
}

  // Razão social é longa demais para o eixo do Chart.js: corta no limite de
  // palavra e mantém a sigla entre parênteses, que é o que identifica o
  // prestador ("X Servícios (CNPJ)" -> "X Serviços… (…)").
  function rotuloCurtoContrato(nome) {
    var s = String(nome == null ? '' : nome);
    if (s.length <= 34) return s;
    return s.slice(0, 33).replace(/\s+\S*$/, '') + '…';
  }

  // Cotações do dólar mês a mês de um ano, como estão na planilha (linha 2).
  // Só leitura — quem atualiza é a Hendrya, direto no Excel.
  function cotacoesBar(ano) {
    var itens = R.MESES.map(function (m) {
      var v = cotaDe(ano, m.key);
      return '<span class="cotacao-chip' + (v == null ? ' vazia' : '') + '">' + m.label +
        '<b>' + (v == null ? '—' : R.formatNumberBR(v, 2)) + '</b></span>';
    }).join('');
    return '<div class="cotacoes-bar">' +
      '<span class="cotacoes-titulo">Cotação do dólar por mês em ' + ano + ' (R$/US$)</span>' +
      '<div class="cotacoes-chips">' + itens + '</div>' +
      '</div>';
  }

  // ==============================================================
  // PÁGINA: GESTÃO DE CONTRATOS (gráfico mensal + tabela espelho da planilha + CRUD)
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
      barraAnoPagina() +
      '<div class="filter-bar">' +
      (IS_SNAPSHOT ? '' : '<button class="btn btn-primary live-only" id="btn-novo-contrato">+ Novo Contrato</button>' +
        // Base do Power BI: ano novo é planilha nova ("Salvar como" no Excel), não colunas novas.
        (MODO_BASE ? '' : '<button class="btn btn-sm live-only" id="btn-inserir-ano" title="Adiciona as colunas mensais de um ano novo na planilha">+ Inserir ano</button>')) +
      '<input type="text" id="gestao-search" placeholder="Buscar ' + (MODO_BASE ? 'código, ' : '') + 'fornecedor, tipo, departamento, comentário…" value="' + escapeHtml(GESTAO_FILTERS.search) + '">' +
      '<select id="gestao-depto-filter"><option value="">Todos os departamentos</option>' +
      deptos.map(function (d) { return '<option value="' + escapeHtml(d) + '"' + (GESTAO_FILTERS.departamento === d ? ' selected' : '') + '>' + escapeHtml(d) + '</option>'; }).join('') +
      '</select>' +
      '<div class="spacer"></div>' +
      '<span class="filter-count" id="gestao-count"></span>' +
      '</div>' +
      '<div class="table-wrap"><div class="table-scroll">' +
      '<table class="data-table table-wide"><thead><tr>' +
      (MODO_BASE ? thSort('codigoFornecedor', 'Código') : '') +
      thSort('razaoSocial', 'Razão Social') + thSort('tipoServico', 'Tipo de Serviço') + thSort('departamento', 'Departamento') +
      '<th>Prazo Pagamento</th><th>Prazo Negociado</th>' +
      thSort('valorMensal2026', MODO_BASE ? 'Valor Mensal' : 'Valor do Contrato') +
      '<th class="cell-num">Total Faturado ' + ANO_ATIVO + ' R$</th><th class="cell-num">Total Faturado ' + ANO_ATIVO + ' $</th>' +
      '<th>VIGÊNCIA</th>' + thSort('dataInicio', 'INÍCIO') + thSort('dataFim', 'FIM') + thSort('status', 'STATUS.') +
      '<th>Comentários</th>' +
      R.MESES.map(function (m) {
        return '<th class="cell-num">' + m.label.toUpperCase() + ' ' + ANO_ATIVO + ' R$</th>' +
          '<th class="cell-num">' + m.label.toUpperCase() + ' ' + ANO_ATIVO + ' $</th>';
      }).join('') +
      (IS_SNAPSHOT ? '' : '<th class="live-only">Ações</th>') +
      '</tr></thead><tbody id="gestao-tbody"></tbody></table></div></div>' +
      (TEM_SAVING ? savingGestaoHtml() : '');

    renderSeletoresAno();
    document.querySelectorAll('#page-gestao select.ano-select').forEach(wireSeletorAno);
    if (TEM_SAVING) wireSavingGestao();
    var btnAno = document.getElementById('btn-inserir-ano');
    if (btnAno) btnAno.addEventListener('click', handleInserirAno);
    document.getElementById('gestao-search').addEventListener('input', function (e) {
      GESTAO_FILTERS.search = e.target.value; renderGestaoTable();
    });
    document.getElementById('gestao-depto-filter').addEventListener('change', function (e) {
      GESTAO_FILTERS.departamento = e.target.value; renderGestaoTable();
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

  // ---------------------------------------------------------------
  // V1.9.0 — Saving na Gestão: tabela da aba Saving do exercício + cadastro/edição.
  // Grava pelo mesmo "Salvar" (DIRTY), na 1ª linha livre da tSaving, como os contratos.
  // ---------------------------------------------------------------
  function savingGestaoHtml() {
    var linhas = SAVING.slice().sort(function (a, b) {
      return String(a.prestador).localeCompare(String(b.prestador), 'pt-BR');
    }).map(function (it) {
      return '<tr class="' + (IS_SNAPSHOT ? '' : 'row-clickable') + '" data-sav-id="' + it.id + '">' +
        '<td>' + escapeHtml(it.codigoFornecedor) + '</td>' +
        '<td>' + escapeHtml(it.prestador) + '</td>' +
        '<td>' + escapeHtml(it.tipoServico) + '</td>' +
        '<td>' + escapeHtml(it.bidConcluido) + '</td>' +
        '<td>' + escapeHtml(it.pagamento) + '</td>' +
        '<td>' + escapeHtml(R.savingDesde(it) || (Number(it.savSemMes) ? 'sem mês' : '—')) + '</td>' +
        '<td class="cell-num">' + R.formatBRL(R.savingBRL(it)) + '</td>' +
        '<td class="cell-num">' + R.formatUSD(R.savingUSD(it, ANO_ATIVO, COTACOES)) + '</td>' +
        (IS_SNAPSHOT ? '' :
          '<td class="td-actions live-only">' +
          '<button class="btn btn-sm btn-icon" data-action="edit-sav">Editar</button>' +
          '<button class="btn btn-sm btn-icon btn-danger" data-action="del-sav">Excluir</button></td>') +
        '</tr>';
    }).join('') || '<tr><td colspan="9" class="cell-muted" style="text-align:center;padding:18px;">Nenhum BID ou negociação lançado em ' +
      ANO_ATIVO + '.</td></tr>';
    return '<div class="card saving-gestao"><div class="filter-bar">' +
      '<h3 style="margin:0;">Saving — BIDs e negociações de ' + ANO_ATIVO + '</h3><div class="spacer"></div>' +
      (IS_SNAPSHOT ? '' : '<button class="btn btn-primary live-only" id="btn-novo-saving">+ Novo saving</button>') +
      '</div><div class="table-wrap"><div class="table-scroll"><table class="data-table"><thead><tr>' +
      '<th>Código</th><th>Prestador</th><th>Tipo de Serviço</th><th>BID concluído</th><th>Pagamento</th><th>Desde</th>' +
      '<th class="cell-num">Saving R$</th><th class="cell-num">Saving US$</th>' + (IS_SNAPSHOT ? '' : '<th class="live-only">Ações</th>') +
      '</tr></thead><tbody>' + linhas + '</tbody></table></div></div>' +
      '<div class="footer-note">Saving em R$ por mês; o total sem detalhe mensal fica à parte. US$ pela cotação de cada mês ' +
      '(o sem mês, pela média das cotações do ano) — a mesma conta da planilha.</div></div>';
  }

  function wireSavingGestao() {
    if (IS_SNAPSHOT) return;
    var btn = document.getElementById('btn-novo-saving');
    if (btn) btn.addEventListener('click', function () { openSavingModal(null); });
    document.querySelectorAll('#page-gestao tr[data-sav-id]').forEach(function (tr) {
      var id = Number(tr.getAttribute('data-sav-id'));
      tr.querySelector('[data-action="edit-sav"]').addEventListener('click', function (e) { e.stopPropagation(); openSavingModal(id); });
      tr.querySelector('[data-action="del-sav"]').addEventListener('click', function (e) { e.stopPropagation(); deleteSaving(id); });
      tr.addEventListener('click', function () { openSavingModal(id); });
    });
  }

  function openSavingModal(id) {
    var it = id == null ? null : SAVING.find(function (x) { return x.id === id; });
    var html =
      '<div class="modal">' +
      '<div class="modal-head"><h2>' + (it ? 'Editar saving' : 'Novo saving') + ' — ' + ANO_ATIVO + '</h2>' +
      '<button class="modal-close" id="modal-close-btn">✕</button></div>' +
      '<div class="modal-body"><div class="form-grid">' +
      fieldText('codigoFornecedor', 'Código Fornecedor (Protheus)', it ? it.codigoFornecedor : '') +
      fieldText('prestador', 'Prestador', it ? it.prestador : '', { required: true }) +
      fieldText('tipoServico', 'Tipo de Serviço', it ? it.tipoServico : '') +
      fieldSelect('bidConcluido', 'BID concluído', it ? it.bidConcluido : 'Sim',
        [{ value: 'Sim', label: 'Sim' }, { value: 'Não', label: 'Não' }]) +
      fieldText('pagamento', 'Pagamento (ex.: de 30 dias P/ 60 dias)', it ? it.pagamento : '', { full: true }) +
      '<div class="field full"><label>Saving por mês em ' + ANO_ATIVO + ' (R$)</label><div class="meses-grid">' +
      R.MESES.map(function (m) {
        var f = R.campoSaving(m.key);
        return '<div class="mes-cell">' + fieldMoney(f, m.label.toUpperCase() + ' R$', it ? it[f] : null) + '</div>';
      }).join('') + '</div></div>' +
      fieldMoney('savSemMes', 'Saving do ano sem detalhe mensal (R$)', it ? it.savSemMes : null) +
      '</div></div>' +
      '<div class="modal-foot"><button class="btn" id="modal-cancel-btn">Cancelar</button>' +
      '<button class="btn btn-primary" id="modal-save-btn">Salvar</button></div></div>';
    document.getElementById('modal-root').innerHTML = html;
    document.getElementById('modal-overlay').classList.remove('hidden');
    document.getElementById('modal-close-btn').addEventListener('click', closeModal);
    document.getElementById('modal-cancel-btn').addEventListener('click', closeModal);
    document.getElementById('modal-save-btn').addEventListener('click', function () { submitSavingModal(id); });
  }

  function submitSavingModal(id) {
    var prestador = val('prestador').trim();
    var ok = true;
    setFieldError('prestador', !prestador);
    if (!prestador) ok = false;
    var campos = R.MESES.map(function (m) { return R.campoSaving(m.key); }).concat(['savSemMes']);
    var valores = {};
    campos.forEach(function (f) {
      var raw = val(f);
      var v = R.parseMoneyInput(raw);
      setFieldError(f, raw.trim() !== '' && v == null);
      if (raw.trim() !== '' && v == null) ok = false;
      valores[f] = v;
    });
    if (!ok) return;
    var fields = Object.assign({
      codigoFornecedor: val('codigoFornecedor').trim(),
      prestador: prestador,
      tipoServico: val('tipoServico').trim(),
      bidConcluido: val('bidConcluido') || 'Sim',
      pagamento: val('pagamento').trim()
    }, valores);
    if (id == null) {
      SAVING.push(Object.assign({ id: NEXT_TEMP_SAV_ID--, isNew: true, _row0: null }, fields));
    } else {
      Object.assign(SAVING.find(function (x) { return x.id === id; }), fields);
    }
    markDirty();
    closeModal();
    toast('Saving ' + (id == null ? 'cadastrado' : 'atualizado') + '. Clique em "Salvar" para gravar na planilha.');
    renderCurrentPage();
  }

  function deleteSaving(id) {
    var it = SAVING.find(function (x) { return x.id === id; });
    if (!it) return;
    if (!confirm('Excluir o saving de "' + it.prestador + '"? Só é gravado na planilha ao clicar em "Salvar".')) return;
    if (!it.isNew && it._row0 != null) DELETED_SAVING0.push(it._row0);
    SAVING = SAVING.filter(function (x) { return x.id !== id; });
    markDirty();
    toast('Saving removido. Clique em "Salvar" para gravar na planilha.');
    renderCurrentPage();
  }

  // "Inserir ano": adiciona o bloco de 26 colunas de um ano novo na planilha
  // (em memória) e marca sujo — o "Salvar" grava em disco. O inverso
  // (planilha → app) é automático: ao abrir, o app lê os anos do cabeçalho.
  function handleInserirAno() {
    if (IS_SNAPSHOT) return;
    if (!planilhaAberta()) {
      toast('Abra a planilha antes de inserir um ano.', true);
      return;
    }
    var ano = window.prompt('Informe o ano a inserir (ex.: ' + R.anoAtual() + '). As 12 colunas de mês em R$, as 12 em dólar e o Total Faturado do ano serão criadas ao lado das demais; a cotação de cada mês fica na linha 2, como nos outros anos:', R.anoAtual());
    if (ano == null || String(ano).trim() === '') return;
    var r = ExcelIO.ensureYear(String(ano).trim());
    if (!r.ok) {
      toast(r.error || 'Não foi possível inserir o ano.', true);
      return;
    }
    ANOS_MENSAL = r.anos.slice();
    ANO_ATIVO = String(ano).trim();
    markDirty();
    renderSeletoresAno();
    toast('Ano ' + ano + ' adicionado. Clique em "Salvar" para gravar na planilha.');
    renderCurrentPage();
  }

  // 13 colunas fixas + os 12 pares de mês do exercício ativo + "Ações" no modo
  // live. Contar aqui evita o colspan literal sair de sincronia com o cabeçalho.
  function colunasDaTabelaGestao() {
    return 13 + (MODO_BASE ? 1 : 0) + R.MESES.length * 2 + (IS_SNAPSHOT ? 0 : 1);
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
        var hay = R.stripAccents([c.codigoFornecedor, c.razaoSocial, c.tipoServico, c.departamento, c.comentarios].join(' ')).toLowerCase();
        if (hay.indexOf(q) === -1) return false;
      }
      return true;
    });
  }

  function renderGestaoTable() {
    var rows = applyGestaoFilters();
    rows = rows.slice().sort(function (a, b) {
      var f = GESTAO_SORT.field, dir = GESTAO_SORT.dir === 'asc' ? 1 : -1;
      var va, vb;
      if (f === 'status') { va = R.statusFinal(a, refAtiva()); vb = R.statusFinal(b, refAtiva()); }
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
      var st = R.statusFinal(c, refAtiva());
      return '<tr class="' + (IS_SNAPSHOT ? '' : 'row-clickable') + '" data-id="' + c.id + '">' +
        (MODO_BASE ? '<td>' + escapeHtml(c.codigoFornecedor) + '</td>' : '') +
        '<td>' + escapeHtml(c.razaoSocial) + '</td>' +
        '<td>' + escapeHtml(c.tipoServico) + '</td>' +
        '<td>' + escapeHtml(c.departamento) + '</td>' +
        '<td>' + escapeHtml(c.prazoPagamentoRaw) + '</td>' +
        '<td>' + escapeHtml(c.prazoNegociadoRaw) + '</td>' +
        '<td class="cell-num">' + R.formatBRL(c.valorMensal2026) + '</td>' +
        '<td class="cell-num">' + R.formatBRL(R.totalAnualReais(c, ANO_ATIVO)) + '</td>' +
        '<td class="cell-num">' + R.formatUSD(R.totalAnualDolares(c, ANO_ATIVO, COTACOES)) + '</td>' +
        '<td>' + escapeHtml(c.vigencia) + '</td>' +
        '<td>' + R.formatDateBR(c.dataInicio) + '</td>' +
        '<td>' + R.formatDateBR(c.dataFim) + '</td>' +
        '<td>' + statusBadge(st) + '</td>' +
        '<td>' + escapeHtml(c.comentarios) + '</td>' +
        R.MESES.map(function (m) {
          return '<td class="cell-num">' + R.formatBRL(c[R.fieldBrlAno(ANO_ATIVO, m.key)]) + '</td>' +
            '<td class="cell-num">' + R.formatUSD(R.mesUSD(c, m, ANO_ATIVO, COTACOES)) + '</td>';
        }).join('') +
        (IS_SNAPSHOT ? '' :
          '<td class="td-actions live-only">' +
          '<button class="btn btn-sm btn-icon" data-action="edit" title="Editar">Editar</button>' +
          '<button class="btn btn-sm btn-icon btn-danger" data-action="del" title="Excluir">Excluir</button>' +
          '</td>') +
        '</tr>';
    }).join('') || '<tr><td colspan="' + colunasDaTabelaGestao() + '" class="cell-muted" style="text-align:center;padding:24px;">Nenhum contrato encontrado com os filtros atuais.</td></tr>';

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
  // Os 12 meses em R$ de CADA ano são o que a Hendrya alimenta (por PROCV no
  // Excel ou por aqui); o valor em dólar de cada mês é derivado da cotação
  // daquele mês/ano e aparece só como leitura, igual à fórmula na planilha.
  function fieldMesesDispendio(c) {
    return '<div class="field full">' +
      anosDoFormulario().map(function (ano) {
        return '<label>Gasto Mensal ' + ano + ' (R$)</label>' +
          '<details class="meses-details"' + (ano === ANO_ATIVO ? ' open' : '') + '>' +
          '<summary>Editar os 12 meses de ' + ano + '</summary>' +
          '<div class="meses-grid">' +
          R.MESES.map(function (m) {
            var f = R.fieldBrlAno(ano, m.key);
            return '<div class="mes-cell">' +
              fieldMoney(f, m.label.toUpperCase() + ' R$', c ? c[f] : null) +
              '<div class="mes-usd" id="usd-' + f + '">—</div></div>';
          }).join('') +
          '</div></details>';
      }).join('') +
      '<div class="field-hint" id="meses-totais"></div>' +
      '</div>';
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
      (MODO_BASE ? fieldText('codigoFornecedor', 'Código Fornecedor (Protheus)', c ? c.codigoFornecedor : '') : '') +
      fieldText('razaoSocial', 'Fornecedor (Razão Social)', c ? c.razaoSocial : '', { required: true, full: !MODO_BASE }) +
      (MODO_BASE ? fieldText('nomeCurto', 'Nome curto (rótulo do gráfico)', c ? c.nomeCurto : '') : '') +
      fieldText('departamento', 'Departamento', c ? c.departamento : '', { list: deptos }) +
      fieldText('tipoServico', 'Tipo de Serviço', c ? c.tipoServico : '', { list: tipos }) +
      fieldDate('dataInicio', 'Início', c ? c.dataInicio : null) +
      fieldDate('dataFim', 'Fim', c ? c.dataFim : null) +
      fieldMoney('valorMensal2026', MODO_BASE ? 'Valor Mensal do Contrato (R$)' : 'Valor do Contrato (R$)', c ? c.valorMensal2026 : null) +
      fieldMoney('valorAnual2026', MODO_BASE ? 'Valor Anual do Contrato (R$)' : 'Valor Anual (R$)', c ? c.valorAnual2026 : null) +
      fieldText('vigencia', 'Vigência', c ? c.vigencia : '') +
      fieldText('prazoPagamentoRaw', 'Prazo de Pagamento', c ? c.prazoPagamentoRaw : '', { errorMsg: '' }) +
      fieldText('prazoNegociadoRaw', 'Prazo Negociado', c ? c.prazoNegociadoRaw : '') +
      // Na base do Power BI o status é fórmula da planilha (pela data de Fim): sem override manual.
      (MODO_BASE ? '' : fieldSelect('statusManual', 'Status Manual', c ? (c.statusManual || '') : '', statusOptions)) +
      fieldTextarea('comentarios', 'Comentários', c ? c.comentarios : '') +
      fieldMesesDispendio(c) +
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

    anosDoFormulario().forEach(function (ano) {
      R.mesFieldsAno(ano).brl.forEach(function (f) {
        var input = document.getElementById('in-' + f);
        if (input) input.addEventListener('input', recomputeTotaisAnuais);
      });
    });
    recomputeTotaisAnuais();
  }

  // Espelha na tela o que a planilha vai calcular: cada mês em dólar pela
  // cotação daquele mês/ano, e o "Total Faturado" de cada ano como soma dos 12
  // meses (é isso que a fórmula da planilha grava).
  function recomputeTotaisAnuais() {
    var linhas = [];
    anosDoFormulario().forEach(function (ano) {
      var somaBrl = 0, somaUsd = 0, algumBrl = false, algumUsd = false;
      R.MESES.forEach(function (m) {
        var f = R.fieldBrlAno(ano, m.key);
        var v = R.parseMoneyInput(val(f));
        var usd = R.convertToUSD(v, cotaDe(ano, m.key));
        var hint = document.getElementById('usd-' + f);
        if (hint) hint.textContent = usd == null ? '—' : R.formatUSD(usd);
        if (v != null) { somaBrl += v; algumBrl = true; }
        if (usd != null) { somaUsd += usd; algumUsd = true; }
      });
      linhas.push(ano + ': ' +
        (algumBrl ? R.formatBRL(somaBrl) : '—') +
        (algumUsd ? ' · ' + R.formatUSD(somaUsd) : ''));
    });
    var totais = document.getElementById('meses-totais');
    if (totais) {
      totais.textContent = MODO_BASE
        ? 'Dispêndio ' + linhas.join(' — ') + ' — a planilha soma os meses e converte para US$ sozinha.'
        : 'Total Faturado por ano: ' + linhas.join(' — ') +
          ' — gravado na planilha como fórmula (soma dos 12 meses de cada ano).';
    }
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

    var mesesValues = {};
    anosDoFormulario().forEach(function (ano) {
      R.mesFieldsAno(ano).brl.forEach(function (f) {
        var raw = val(f);
        var v = R.parseMoneyInput(raw);
        mesesValues[f] = v;
        setFieldError(f, false);
        if (raw.trim() !== '' && v == null) { setFieldError(f, true); ok = false; }
      });
    });
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
    if (MODO_BASE) {
      fields.codigoFornecedor = val('codigoFornecedor').trim();
      fields.nomeCurto = val('nomeCurto').trim();
    }
    anosDoFormulario().forEach(function (ano) {
      R.mesFieldsAno(ano).brl.forEach(function (f) { fields[f] = mesesValues[f]; });
    });

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
