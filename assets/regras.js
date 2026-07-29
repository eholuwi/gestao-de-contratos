/* ============================================================
   regras.js — Lógica de negócio do Dashboard de Gestão de Contratos
   Porta para JS as mesmas regras validadas no modelo Power BI
   (ver 04_Dashboard Gestão de Contratos/PowerBI/Claude/00_Especificacao_e_Modelo/).

   Não depende de DOM — pode ser testado isoladamente e é reaproveitado
   sem alteração dentro do snapshot estático (modo visualização).
   ============================================================ */
(function (global) {
  'use strict';

  var STATUS_ORDER = ['No Prazo', 'A Vencer', 'Vencido', 'Em Análise'];

  var STATUS_COLORS = {
    'No Prazo': '#16A34A',
    'A Vencer': '#FF851B',
    'Vencido': '#DC2626',
    'Em Análise': '#7a7f87'
  };

  var STATUS_CSS_CLASS = {
    'No Prazo': 'badge-no-prazo',
    'A Vencer': 'badge-a-vencer',
    'Vencido': 'badge-vencido',
    'Em Análise': 'badge-em-analise'
  };

  var VENCIDO_CRITICO_LIMIAR = 100000;
  var VIGENCIA_LONGA_ANOS = 3;
  var JANELA_A_VENCER_DIAS = 30;

  // ---------------------------------------------------------------
  // Normalização de texto
  // ---------------------------------------------------------------
  function stripAccents(s) {
    return String(s == null ? '' : s)
      .normalize('NFD')
      .replace(/[̀-ͯ]/g, '');
  }

  function normSpace(s) {
    return String(s == null ? '' : s).replace(/\s+/g, ' ').trim();
  }

  function normalizeRazaoSocial(raw) {
    var s = normSpace(raw);
    s = s.replace(/\.+$/, ''); // remove ponto final duplicado
    return s.trim();
  }

  function normalizeDepartamento(raw) {
    var s = normSpace(raw);
    if (!s) return '';
    if (/^financial$/i.test(s)) return 'Finance';
    return s;
  }

  function normalizeTipoServico(raw) {
    var s = normSpace(raw);
    return s === '' ? 'Não Classificado' : s;
  }

  // "30 dias" / "30 DIAS" / "A VISTA" / "07 DIAS" -> número de dias (ou null)
  function parsePrazoPagamentoDias(raw) {
    var s = stripAccents(String(raw == null ? '' : raw)).toUpperCase().trim();
    if (s === '') return null;
    if (s.indexOf('VISTA') > -1) return 0;
    var digits = s.replace(/[^0-9]/g, '');
    if (digits === '') return null;
    return parseInt(digits, 10);
  }

  // Tenta mapear texto legado de status ("Á VENCER", "VENCIDO"...) para
  // um dos 4 valores canônicos. Retorna null se não reconhecer (ou vazio).
  function parseStatusLegado(raw) {
    var s = stripAccents(String(raw == null ? '' : raw)).toUpperCase().trim();
    if (s === '') return null;
    if (s.indexOf('NO PRAZO') > -1 || s === 'EM DIA') return 'No Prazo';
    if (s.indexOf('A VENCER') > -1) return 'A Vencer';
    if (s.indexOf('VENCID') > -1) return 'Vencido';
    if (s.indexOf('ANALISE') > -1 || s.indexOf('EM ABERTO') > -1) return 'Em Análise';
    return null;
  }

  // ---------------------------------------------------------------
  // Datas
  // ---------------------------------------------------------------
  function startOfDay(d) {
    if (!d) return null;
    var x = new Date(d.getFullYear(), d.getMonth(), d.getDate());
    return x;
  }

  function addDays(d, n) {
    var x = new Date(d);
    x.setDate(x.getDate() + n);
    return x;
  }

  function addYears(d, n) {
    var x = new Date(d);
    x.setFullYear(x.getFullYear() + n);
    return x;
  }

  // ---------------------------------------------------------------
  // status_calculado — mesma régua do DAX:
  //   sem FIM -> "Em Análise"
  //   FIM <= hoje -> "Vencido"
  //   FIM <= hoje+30 -> "A Vencer"
  //   senão -> "No Prazo"
  // ---------------------------------------------------------------
  function statusCalculado(dataFim, hoje) {
    var t = startOfDay(hoje || new Date());
    if (!dataFim) return 'Em Análise';
    var f = startOfDay(dataFim);
    if (f <= t) return 'Vencido';
    if (f <= addDays(t, JANELA_A_VENCER_DIAS)) return 'A Vencer';
    return 'No Prazo';
  }

  // status_final: override manual (dropdown "Status Manual") tem prioridade;
  // "Automático" (null/vazio) usa a régua de data.
  function statusFinal(contrato, hoje) {
    if (contrato.statusManual && STATUS_ORDER.indexOf(contrato.statusManual) > -1) {
      return contrato.statusManual;
    }
    return statusCalculado(contrato.dataFim, hoje);
  }

  // ---------------------------------------------------------------
  // Flags de qualidade de dados (ex-página "Auditoria")
  //
  // "Status inconsistente" compara o status calculado contra uma referência:
  // - se já existe um override ativo (contrato.statusManual, definido por
  //   quem preencheu o formulário), compara contra ELE — sinaliza que o
  //   override em vigor diverge da régua automática (útil pra revisão).
  // - senão, compara contra o texto legado da coluna STATUS. (statusLegadoRaw,
  //   só leitura) — reproduz o tipo de achado que motivou trocar o status
  //   manual pelo automático (planilhas legadas tendem a acumular status
  //   desatualizado), agora como um item acionável na tela de Gestão em vez
  //   de um relatório estático.
  // Nenhum dos dois casos altera o status EXIBIDO (isso é statusFinal, que só
  // olha contrato.statusManual) — este flag é só um alerta de qualidade.
  // ---------------------------------------------------------------
  function qualityFlags(contrato, hoje) {
    var calc = statusCalculado(contrato.dataFim, hoje);
    var referencia = contrato.statusManual || parseStatusLegado(contrato.statusLegadoRaw);
    var statusDivergente = !!(referencia && referencia !== calc);
    return {
      semData: !contrato.dataFim,
      semValor: contrato.valorAnual2026 == null,
      statusInconsistente: statusDivergente,
      algum: !contrato.dataFim || contrato.valorAnual2026 == null || statusDivergente
    };
  }

  // ---------------------------------------------------------------
  // KPIs — as mesmas 14 medidas validadas no Power BI (_Medidas.tmdl)
  // ---------------------------------------------------------------
  function computeKPIs(contratos, hoje) {
    hoje = hoje || new Date();
    var total = contratos.length;
    var noPrazo = 0, aVencer = 0, vencido = 0, emAnalise = 0;
    var despesaTotal = 0, semValor = 0, despesaVencidos = 0, vencidosCriticos = 0;
    var somaPrazoPagamento = 0, nPrazoPagamento = 0;
    var vigenciaLonga = 0;
    var limiteVigenciaLonga = addYears(startOfDay(hoje), VIGENCIA_LONGA_ANOS);

    for (var i = 0; i < contratos.length; i++) {
      var c = contratos[i];
      var st = statusFinal(c, hoje);
      if (st === 'No Prazo') noPrazo++;
      else if (st === 'A Vencer') aVencer++;
      else if (st === 'Vencido') vencido++;
      else emAnalise++;

      if (c.valorAnual2026 != null) {
        despesaTotal += c.valorAnual2026;
      } else {
        semValor++;
      }

      if (st === 'Vencido' && c.valorAnual2026 != null) despesaVencidos += c.valorAnual2026;
      if (st === 'Vencido' && c.valorAnual2026 != null && c.valorAnual2026 > VENCIDO_CRITICO_LIMIAR) vencidosCriticos++;

      if (c.prazoPagamentoDias != null) {
        somaPrazoPagamento += c.prazoPagamentoDias;
        nPrazoPagamento++;
      }

      if (c.dataFim && startOfDay(c.dataFim) > limiteVigenciaLonga) vigenciaLonga++;
    }

    var comValor = total - semValor;

    return {
      totalContratos: total,
      contratosNoPrazo: noPrazo,
      contratosAVencer: aVencer,
      contratosVencidos: vencido,
      contratosEmAnalise: emAnalise,
      pctContratosVencidos: total ? vencido / total : 0,
      pctContratosAVencer: total ? aVencer / total : 0,
      despesaTotalAnual: despesaTotal,
      contratosSemValorInformado: semValor,
      despesaMediaPorContrato: comValor ? despesaTotal / comValor : 0,
      despesaContratosVencidos: despesaVencidos,
      contratosVencidosCriticos: vencidosCriticos,
      prazoMedioPagamentoDias: nPrazoPagamento ? somaPrazoPagamento / nPrazoPagamento : null,
      contratosVigenciaLonga: vigenciaLonga
    };
  }

  // ---------------------------------------------------------------
  // Formatadores
  // ---------------------------------------------------------------
  function formatBRL(v, compact) {
    if (v == null || isNaN(v)) return '—';
    if (compact) {
      var abs = Math.abs(v);
      if (abs >= 1e6) return 'R$ ' + (v / 1e6).toLocaleString('pt-BR', { minimumFractionDigits: 1, maximumFractionDigits: 1 }) + 'M';
      if (abs >= 1e3) return 'R$ ' + (v / 1e3).toLocaleString('pt-BR', { maximumFractionDigits: 0 }) + 'K';
    }
    return v.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL', minimumFractionDigits: 0, maximumFractionDigits: 0 });
  }

  function formatNumberBR(v, decimals) {
    if (v == null || isNaN(v)) return '—';
    return v.toLocaleString('pt-BR', { minimumFractionDigits: decimals || 0, maximumFractionDigits: decimals || 0 });
  }

  function formatPercent(v) {
    if (v == null || isNaN(v)) return '—';
    return (v * 100).toLocaleString('pt-BR', { maximumFractionDigits: 0 }) + '%';
  }

  function formatDateBR(d) {
    if (!d) return '—';
    var dt = (d instanceof Date) ? d : new Date(d);
    if (isNaN(dt.getTime())) return '—';
    var dd = String(dt.getDate()).padStart(2, '0');
    var mm = String(dt.getMonth() + 1).padStart(2, '0');
    return dd + '/' + mm + '/' + dt.getFullYear();
  }

  function formatDiasSuffix(v) {
    if (v == null || isNaN(v)) return '—';
    return Math.round(v) + ' dias';
  }

  // Converte string "dd/mm/aaaa" ou "aaaa-mm-dd" (input type=date) em Date local
  function parseDateInput(s) {
    if (!s) return null;
    s = String(s).trim();
    var m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (m) return new Date(parseInt(m[1], 10), parseInt(m[2], 10) - 1, parseInt(m[3], 10));
    m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (m) return new Date(parseInt(m[3], 10), parseInt(m[2], 10) - 1, parseInt(m[1], 10));
    var d = new Date(s);
    return isNaN(d.getTime()) ? null : d;
  }

  // Date -> "aaaa-mm-dd" (valor de <input type=date>)
  function toDateInputValue(d) {
    if (!d) return '';
    var dt = (d instanceof Date) ? d : new Date(d);
    if (isNaN(dt.getTime())) return '';
    var mm = String(dt.getMonth() + 1).padStart(2, '0');
    var dd = String(dt.getDate()).padStart(2, '0');
    return dt.getFullYear() + '-' + mm + '-' + dd;
  }

  function parseMoneyInput(s) {
    if (s == null || s === '') return null;
    if (typeof s === 'number') return s;
    var t = String(s).trim().replace(/R\$|\s/g, '');
    if (t === '') return null;
    var hasComma = t.indexOf(',') > -1, hasDot = t.indexOf('.') > -1;
    if (hasComma && hasDot) t = t.replace(/\./g, '').replace(',', '.');
    else if (hasComma) t = t.replace(',', '.');
    var n = parseFloat(t);
    return isNaN(n) ? null : n;
  }

  // ---------------------------------------------------------------
  // Agregações auxiliares para gráficos
  // ---------------------------------------------------------------
  function topN(pairs, n) {
    // pairs: [{label, value}] já somados; retorna top N + "Outros"
    var sorted = pairs.slice().sort(function (a, b) { return b.value - a.value; });
    if (sorted.length <= n) return sorted;
    var top = sorted.slice(0, n);
    var outros = sorted.slice(n).reduce(function (acc, p) { return acc + p.value; }, 0);
    top.push({ label: 'Outros', value: outros });
    return top;
  }

  function sumBy(contratos, keyFn, valueFn) {
    var map = {};
    contratos.forEach(function (c) {
      var k = keyFn(c) || 'Não Classificado';
      var v = valueFn(c);
      if (v == null) return;
      map[k] = (map[k] || 0) + v;
    });
    return Object.keys(map).map(function (k) { return { label: k, value: map[k] }; });
  }

  function countBy(contratos, keyFn) {
    var map = {};
    contratos.forEach(function (c) {
      var k = keyFn(c) || 'Não Classificado';
      map[k] = (map[k] || 0) + 1;
    });
    return Object.keys(map).map(function (k) { return { label: k, value: map[k] }; });
  }

  global.Regras = {
    STATUS_ORDER: STATUS_ORDER,
    stripAccents: stripAccents,
    normSpace: normSpace,
    STATUS_COLORS: STATUS_COLORS,
    STATUS_CSS_CLASS: STATUS_CSS_CLASS,
    VENCIDO_CRITICO_LIMIAR: VENCIDO_CRITICO_LIMIAR,
    VIGENCIA_LONGA_ANOS: VIGENCIA_LONGA_ANOS,
    JANELA_A_VENCER_DIAS: JANELA_A_VENCER_DIAS,

    normalizeRazaoSocial: normalizeRazaoSocial,
    normalizeDepartamento: normalizeDepartamento,
    normalizeTipoServico: normalizeTipoServico,
    parsePrazoPagamentoDias: parsePrazoPagamentoDias,
    parseStatusLegado: parseStatusLegado,

    startOfDay: startOfDay,
    addDays: addDays,
    addYears: addYears,

    statusCalculado: statusCalculado,
    statusFinal: statusFinal,
    qualityFlags: qualityFlags,
    computeKPIs: computeKPIs,

    formatBRL: formatBRL,
    formatNumberBR: formatNumberBR,
    formatPercent: formatPercent,
    formatDateBR: formatDateBR,
    formatDiasSuffix: formatDiasSuffix,
    parseDateInput: parseDateInput,
    toDateInputValue: toDateInputValue,
    parseMoneyInput: parseMoneyInput,

    topN: topN,
    sumBy: sumBy,
    countBy: countBy
  };
})(window);
