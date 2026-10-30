/* ============================================================
   regras.js — Lógica de negócio do Dashboard de Gestão de Contratos
   Porta para JS as mesmas regras validadas no modelo Power BI
   (ver 04_Dashboard Gestão de Contratos/PowerBI/especificacao/00_Especificacao_e_Modelo/).

   Não depende de DOM — pode ser testado isoladamente e é reaproveitado
   sem alteração dentro do snapshot estático (modo visualização).
   ============================================================ */
(function (global) {
  'use strict';

  var STATUS_ORDER = ['No Prazo', 'A Vencer', 'Vencido', 'Em Análise', 'Distrato/Expirado'];

  var STATUS_COLORS = {
    'No Prazo': '#16A34A',
    'A Vencer': '#FF851B',
    'Vencido': '#DC2626',
    'Em Análise': '#7a7f87',
    'Distrato/Expirado': '#6D28D9'
  };

  var STATUS_CSS_CLASS = {
    'No Prazo': 'badge-no-prazo',
    'A Vencer': 'badge-a-vencer',
    'Vencido': 'badge-vencido',
    'Em Análise': 'badge-em-analise',
    'Distrato/Expirado': 'badge-distrato'
  };

  var VIGENCIA_LONGA_ANOS = 3;
  var JANELA_A_VENCER_DIAS = 30;

  // Cada mês ocupa DUAS colunas na planilha: "JAN R$" (dispêndio do mês) e
  // "JAN $" (o mesmo valor convertido pela cotação daquele mês, que fica na
  // linha 2 da própria coluna "$" — ver excel.js).
  // `nomeLongo` é o que vai no rótulo da cotação na linha 1 da planilha
  // ("cotaçao do dolar janeiro 2026") — a grafia que a Hendrya já usa lá.
  var MESES = [
    { key: 'jan', label: 'Jan', nomeLongo: 'janeiro', fieldBrl: 'mesJanBrl', fieldUsd: 'mesJanUsd', headerBrl: 'jan r$', headerUsd: 'jan $', headerLegado: 'jan' },
    { key: 'fev', label: 'Fev', nomeLongo: 'fevereiro', fieldBrl: 'mesFevBrl', fieldUsd: 'mesFevUsd', headerBrl: 'fev r$', headerUsd: 'fev $', headerLegado: 'fev' },
    { key: 'mar', label: 'Mar', nomeLongo: 'março', fieldBrl: 'mesMarBrl', fieldUsd: 'mesMarUsd', headerBrl: 'mar r$', headerUsd: 'mar $', headerLegado: 'mar' },
    { key: 'abr', label: 'Abr', nomeLongo: 'abril', fieldBrl: 'mesAbrBrl', fieldUsd: 'mesAbrUsd', headerBrl: 'abr r$', headerUsd: 'abr $', headerLegado: 'abr' },
    { key: 'mai', label: 'Mai', nomeLongo: 'maio', fieldBrl: 'mesMaiBrl', fieldUsd: 'mesMaiUsd', headerBrl: 'mai r$', headerUsd: 'mai $', headerLegado: 'mai' },
    { key: 'jun', label: 'Jun', nomeLongo: 'junho', fieldBrl: 'mesJunBrl', fieldUsd: 'mesJunUsd', headerBrl: 'jun r$', headerUsd: 'jun $', headerLegado: 'jun' },
    { key: 'jul', label: 'Jul', nomeLongo: 'julho', fieldBrl: 'mesJulBrl', fieldUsd: 'mesJulUsd', headerBrl: 'jul r$', headerUsd: 'jul $', headerLegado: 'jul' },
    { key: 'ago', label: 'Ago', nomeLongo: 'agosto', fieldBrl: 'mesAgoBrl', fieldUsd: 'mesAgoUsd', headerBrl: 'ago r$', headerUsd: 'ago $', headerLegado: 'ago' },
    { key: 'set', label: 'Set', nomeLongo: 'setembro', fieldBrl: 'mesSetBrl', fieldUsd: 'mesSetUsd', headerBrl: 'set r$', headerUsd: 'set $', headerLegado: 'set' },
    { key: 'out', label: 'Out', nomeLongo: 'outubro', fieldBrl: 'mesOutBrl', fieldUsd: 'mesOutUsd', headerBrl: 'out r$', headerUsd: 'out $', headerLegado: 'out' },
    { key: 'nov', label: 'Nov', nomeLongo: 'novembro', fieldBrl: 'mesNovBrl', fieldUsd: 'mesNovUsd', headerBrl: 'nov r$', headerUsd: 'nov $', headerLegado: 'nov' },
    { key: 'dez', label: 'Dez', nomeLongo: 'dezembro', fieldBrl: 'mesDezBrl', fieldUsd: 'mesDezUsd', headerBrl: 'dez r$', headerUsd: 'dez $', headerLegado: 'dez' }
  ];
  var MESES_FIELDS_BRL = MESES.map(function (m) { return m.fieldBrl; });
  var MESES_FIELDS_USD = MESES.map(function (m) { return m.fieldUsd; });

  var MESES_BY_KEY = {};
  MESES.forEach(function (m) { MESES_BY_KEY[m.key] = m; });

  // Anos com colunas mensais na planilha. A ordem é só o padrão usado quando a
  // planilha ainda não foi lida; quem manda são os anos detectados no cabeçalho
  // em excel.js (vice-versa: a Hendrya pode acrescentar um ano no Excel).
  var ANOS_MENSAIS = ['2026', '2025'];

  // Exercício corrente (o ano em que estamos). É o padrão do seletor de ano:
  // se a planilha ainda não tiver o bloco desse ano, ele é criado no seletor
  // como "2027 — em preparação" e o mês corrente aparece zerado.
  function anoAtual() { return String(new Date().getFullYear()); }

  // Cada ano tem seu bloco de 24 colunas (12 R$ + 12 $) na planilha. Os campos
  // de 2026 mantêm os nomes históricos (mesJanBrl/mesJanUsd) para compatibilidade
  // com snapshots antigos; anos novos recebem sufixo: mesJanBrl2025, mesJanBrl2024.
  function cap(s) { return s.charAt(0).toUpperCase() + s.slice(1); }

  function fieldBrlAno(ano, mesKey) {
    return ano === '2026' ? MESES_BY_KEY[mesKey].fieldBrl : 'mes' + cap(mesKey) + 'Brl' + ano;
  }

  function fieldUsdAno(ano, mesKey) {
    return ano === '2026' ? MESES_BY_KEY[mesKey].fieldUsd : 'mes' + cap(mesKey) + 'Usd' + ano;
  }

  function mesFieldsAno(ano) {
    var brl = [], usd = [];
    MESES.forEach(function (m) {
      brl.push(fieldBrlAno(ano, m.key));
      usd.push(fieldUsdAno(ano, m.key));
    });
    return { brl: brl, usd: usd };
  }

  // Cada ano fecha o seu bloco com o par de totais "Total Faturado <ano> R$" e
  // "Total Faturado <ano> $" — a soma dos 12 meses daquele ano. Os nomes de
  // campo saem daqui para excel.js e app.js não montarem a string por conta.
  function fieldTotalAno(ano) { return 'totalFaturado' + ano; }
  function fieldTotalUsdAno(ano) { return 'totalFaturadoUsd' + ano; }

  // Lê a cotação do mês ignorando a forma do objeto cotacoes: suporta o shape
  // novo {ano: {mes: taxa}} e o legado chato {mes: taxa} (pré-multi-ano).
  function cotaMes(cotacoes, ano, mesKey) {
    if (!cotacoes) return null;
    var porAno = cotacoes[ano];
    if (porAno && typeof porAno === 'object') return porAno[mesKey] != null ? porAno[mesKey] : null;
    return cotacoes[mesKey] != null ? cotacoes[mesKey] : null;
  }

  // Soma só os meses preenchidos (sem prorata); null se nenhum mês tiver valor.
  // `ano` opcional (padrão '2026'); para 2026 usa os campos legados (mesJanBrl...).
  function somaMesesBRL(contrato, ano) {
    ano = ano || '2026';
    var fields = mesFieldsAno(ano).brl;
    var soma = 0, algum = false;
    fields.forEach(function (f) {
      if (contrato[f] != null) { soma += contrato[f]; algum = true; }
    });
    return algum ? soma : null;
  }

  // Valor do mês em dólar = valor do mês em R$ ÷ cotação DAQUELE mês.
  // Assinatura atual: mesUSD(contrato, mes, ano, cotacoes). Aceita também a
  // forma legada (contrato, mes, cotacoes) tratando o 3º argumento como cotacoes.
  function mesUSD(contrato, mes, ano, cotacoes) {
    if (typeof ano === 'object') { cotacoes = ano; ano = '2026'; }
    return convertToUSD(contrato[fieldBrlAno(ano, mes.key)], cotaMes(cotacoes, ano, mes.key));
  }

  function somaMesesUSD(contrato, ano, cotacoes) {
    if (typeof ano === 'object') { cotacoes = ano; ano = '2026'; }
    var soma = 0, algum = false;
    MESES.forEach(function (m) {
      var v = mesUSD(contrato, m, ano, cotacoes);
      if (v != null) { soma += v; algum = true; }
    });
    return algum ? soma : null;
  }

  // "Total Faturado <ano>" — puramente o somatório dos valores mensais daquele
  // ano (o que a planilha calcula como fórmula). Null se nenhum mês lançado.
  function totalFaturado(contrato, ano) {
    return somaMesesBRL(contrato, ano);
  }

  // "Despesa" do ano para os KPIs: o total faturado quando há meses; senão cai
  // para o valor declarado (Valor Anual do Contrato, só existe para 2026) e,
  // por fim, para o valor que a planilha tiver em Total Faturado <ano>.
  // V1.7.0: sem cair para o Valor Anual — dispêndio do ano é o que foi lançado nos
  // meses, como na planilha do Power BI ("Dispêndio do Ano (soma dos meses)").
  function despesaAno(contrato, ano) {
    var t = somaMesesBRL(contrato, ano);
    if (t != null) return t;
    var col = contrato[fieldTotalAno(ano)];
    if (col != null) return col;
    return null;
  }

  // "Total Faturado <ano>" em R$ e em US$: o app recalcula a partir dos meses; o
  // valor lido da planilha só entra quando o contrato ainda não tem mês lançado
  // (fórmula recém-gravada não traz valor em cache).
  function totalAnualReais(contrato, ano) {
    ano = ano || '2026';
    var soma = somaMesesBRL(contrato, ano);
    if (soma != null) return soma;
    var col = contrato[fieldTotalAno(ano)];
    return col != null ? col : null;
  }

  function totalAnualDolares(contrato, ano, cotacoes) {
    if (typeof ano === 'object') { cotacoes = ano; ano = '2026'; }
    var soma = somaMesesUSD(contrato, ano, cotacoes);
    if (soma != null) return soma;
    var col = contrato[fieldTotalUsdAno(ano)];
    return col != null ? col : null;
  }

  // Valor negociado do contrato ("Despesa Total Anual"), a mesma regra da coluna
  // "Valor negociado do ano" da planilha do Power BI (V1.7.0, decisão do Luis em
  // 06/10/2026: os dois painéis mostram o mesmo número): Valor Anual; vazio ou
  // zero -> Valor Mensal x 12; nenhum dos dois -> sem valor (null). `ano` ficou
  // na assinatura só por compatibilidade.
  function valorTotalContrato(contrato, ano) {
    if (contrato.valorAnual2026 > 0) return contrato.valorAnual2026;
    if (contrato.valorMensal2026 > 0) return contrato.valorMensal2026 * 12;
    return null;
  }

  // ---------------------------------------------------------------
  // Saving (aba Saving da base do Power BI, V1.9.0) — a mesma régua da planilha:
  // saving em R$ por mês (campos savJan..savDez) + total sem detalhe mensal
  // (savSemMes). US$ do mês = R$ do mês ÷ cotação daquele mês (sem cotação fica
  // de fora, como no dispêndio); US$ do "sem mês" = ÷ MÉDIA das cotações
  // preenchidas do ano. "Desde" = 1º mês com saving.
  // ---------------------------------------------------------------
  function campoSaving(mesKey) { return 'sav' + cap(mesKey); }

  function savingMes(item, mesKey) {
    var v = item[campoSaving(mesKey)];
    return (v == null || v === '' || isNaN(Number(v))) ? null : Number(v);
  }

  function savingBRL(item) {
    var t = 0;
    MESES.forEach(function (m) { var v = savingMes(item, m.key); if (v != null) t += v; });
    return t + (Number(item.savSemMes) || 0);
  }

  function savingUSD(item, ano, cotacoes) {
    var porAno = (cotacoes || {})[ano] || {};
    var t = 0;
    MESES.forEach(function (m) {
      var v = savingMes(item, m.key);
      var c = porAno[m.key];
      if (v != null && c) t += v / c;
    });
    var sem = Number(item.savSemMes) || 0;
    var cots = MESES.map(function (m) { return porAno[m.key]; }).filter(function (c) { return c != null && !isNaN(c); });
    if (sem && cots.length) {
      var media = cots.reduce(function (a, b) { return a + b; }, 0) / cots.length;
      if (media) t += sem / media;
    }
    return t;
  }

  function savingDesde(item) {
    for (var i = 0; i < MESES.length; i++) {
      var v = savingMes(item, MESES[i].key);
      if (v != null && v > 0) return MESES[i].label;
    }
    return '';
  }

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
    return statusPorComentario(contrato.comentarios) || statusCalculado(contrato.dataFim, hoje);
  }

  // Marcador no Comentário manda no status (= fórmula Status da planilha, montar_base_pbi.py
  // MARCADORES_STATUS; o SEARCH do Excel também não liga para maiúscula/minúscula).
  function statusPorComentario(comentarios) {
    var s = String(comentarios || '').toUpperCase();
    if (s.indexOf('STATUS: DISTRATO') > -1 || s.indexOf('STATUS: EXPIRADO') > -1) return 'Distrato/Expirado';
    if (s.indexOf('STATUS: VENCIDO') > -1) return 'Vencido';
    return null;
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
  // `ano` decide de qual exercício sai a despesa faturada (dólar, vencidos).
  // A "Despesa Total Anual" NÃO depende do ano: é a soma do valor total de cada
  // contrato (ver `valorTotalContrato`). As medidas de contagem/status também não.
  function computeKPIs(contratos, hoje, cotacoes, ano) {
    hoje = hoje || new Date();
    ano = ano || ANOS_MENSAIS[0];
    var total = contratos.length;
    var noPrazo = 0, aVencer = 0, vencido = 0, emAnalise = 0, distrato = 0;
    var despesaTotal = 0, semValor = 0, despesaVencidos = 0;
    var despesaTotalUSD = 0, despesaDoAno = 0;
    var somaPrazoPagamento = 0, nPrazoPagamento = 0;
    var prazoNegociado = 0;
    var vigenciaLonga = 0;
    var limiteVigenciaLonga = addYears(startOfDay(hoje), VIGENCIA_LONGA_ANOS);

    for (var i = 0; i < contratos.length; i++) {
      var c = contratos[i];
      var st = statusFinal(c, hoje);
      if (st === 'No Prazo') noPrazo++;
      else if (st === 'A Vencer') aVencer++;
      else if (st === 'Vencido') vencido++;
      else if (st === 'Distrato/Expirado') distrato++;
      else emAnalise++;

      // "Despesa Total Anual" = valor negociado dos contratos (Valor Anual, ou
      // Mensal x 12), não a soma do que foi faturado no exercício — pedido da
      // Hendrya em 28/09/2026, regra da planilha do Power BI desde a V1.7.0.
      var valorContrato = valorTotalContrato(c, ano);
      var despesa = despesaAno(c, ano);
      if (valorContrato != null) {
        despesaTotal += valorContrato;
      } else {
        semValor++;
      }
      if (despesa != null) despesaDoAno += despesa;

      var anualUSD = totalAnualDolares(c, ano, cotacoes);
      if (anualUSD != null) despesaTotalUSD += anualUSD;

      if (st === 'Vencido' && despesa != null) {
        despesaVencidos += despesa;
      }

      if (c.prazoPagamentoDias != null) {
        somaPrazoPagamento += c.prazoPagamentoDias;
        nPrazoPagamento++;
      }

      // "Prazo de Pagamento Negociado": qualquer texto preenchido na coluna
      // Prazo Negociado (não é sobre prazo ter sido ou não alterado — é sobre o
      // fornecedor já ter passado por negociação de condição de pagamento).
      if (String(c.prazoNegociadoRaw || '').trim() !== '') prazoNegociado++;

      if (c.dataFim && startOfDay(c.dataFim) > limiteVigenciaLonga) vigenciaLonga++;
    }

    var comValor = total - semValor;

    return {
      totalContratos: total,
      contratosNoPrazo: noPrazo,
      contratosAVencer: aVencer,
      contratosVencidos: vencido,
      contratosEmAnalise: emAnalise,
      contratosDistrato: distrato,
      pctContratosVencidos: total ? vencido / total : 0,
      pctContratosAVencer: total ? aVencer / total : 0,
      despesaTotalAnual: despesaTotal,
      despesaTotalAnualUSD: despesaTotalUSD,
      despesaTotalAno: despesaDoAno,
      contratosSemValorInformado: semValor,
      despesaMediaPorContrato: comValor ? despesaTotal / comValor : 0,
      despesaContratosVencidos: despesaVencidos,
      prazoMedioPagamentoDias: nPrazoPagamento ? somaPrazoPagamento / nPrazoPagamento : null,
      contratosVigenciaLonga: vigenciaLonga,
      contratosPrazoPagamentoNegociado: prazoNegociado
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

  // Conversão manual via taxa de câmbio informada pelo usuário (não é cotação
  // ao vivo — ver decisão registrada: alternativa (a), campo preenchido periodicamente).
  function formatUSD(v, compact) {
    if (v == null || isNaN(v)) return '—';
    if (compact) {
      var abs = Math.abs(v);
      if (abs >= 1e6) return 'US$ ' + (v / 1e6).toLocaleString('en-US', { minimumFractionDigits: 1, maximumFractionDigits: 1 }) + 'M';
      if (abs >= 1e3) return 'US$ ' + (v / 1e3).toLocaleString('en-US', { maximumFractionDigits: 0 }) + 'K';
    }
    return v.toLocaleString('en-US', { style: 'currency', currency: 'USD', minimumFractionDigits: 0, maximumFractionDigits: 0 });
  }

  function convertToUSD(valorBRL, taxaCambio) {
    if (valorBRL == null || isNaN(valorBRL) || taxaCambio == null || !taxaCambio) return null;
    return valorBRL / taxaCambio;
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
    VIGENCIA_LONGA_ANOS: VIGENCIA_LONGA_ANOS,
    JANELA_A_VENCER_DIAS: JANELA_A_VENCER_DIAS,
    MESES: MESES,
    MESES_FIELDS_BRL: MESES_FIELDS_BRL,
    MESES_FIELDS_USD: MESES_FIELDS_USD,
    ANOS_MENSAIS: ANOS_MENSAIS,
    anoAtual: anoAtual,
    fieldBrlAno: fieldBrlAno,
    fieldUsdAno: fieldUsdAno,
    fieldTotalAno: fieldTotalAno,
    fieldTotalUsdAno: fieldTotalUsdAno,
    mesFieldsAno: mesFieldsAno,
    cotaMes: cotaMes,
    somaMesesBRL: somaMesesBRL,
    somaMesesUSD: somaMesesUSD,
    mesUSD: mesUSD,
    totalFaturado: totalFaturado,
    despesaAno: despesaAno,
    valorTotalContrato: valorTotalContrato,
    campoSaving: campoSaving,
    savingMes: savingMes,
    savingBRL: savingBRL,
    savingUSD: savingUSD,
    savingDesde: savingDesde,
    totalAnualReais: totalAnualReais,
    totalAnualDolares: totalAnualDolares,

    normalizeRazaoSocial: normalizeRazaoSocial,
    normalizeDepartamento: normalizeDepartamento,
    normalizeTipoServico: normalizeTipoServico,
    parsePrazoPagamentoDias: parsePrazoPagamentoDias,
    parseStatusLegado: parseStatusLegado,

    startOfDay: startOfDay,
    addDays: addDays,
    addYears: addYears,

    statusCalculado: statusCalculado,
    statusPorComentario: statusPorComentario,
    statusFinal: statusFinal,
    qualityFlags: qualityFlags,
    computeKPIs: computeKPIs,

    formatBRL: formatBRL,
    formatUSD: formatUSD,
    convertToUSD: convertToUSD,
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
