/* ============================================================
   excel.js — Leitura/escrita da MATRIZ DE CONTRATOS 2026.xlsx
   direto no navegador via File System Access API + ExcelJS.

   Dois modos de execução:
   - Modo navegador (padrão): File System Access API — abre uma
     PASTA para resolver o snapshot na mesma concessão.
   - Modo desktop (quando window.DesktopBridge existe, ver
     desktop/): opera sobre caminhos de ARQUIVO fornecidos pelo
      processo principal do Electron — auto-open da planilha ao
      lado do executável, seletor de arquivo (não pasta), escrita
      com o snapshot no mesmo diretório do arquivo.

   V1.7.0 — modo "base do Power BI": se a pasta (ou subpasta) tiver as
   planilhas MATRIZ DE POWERBI CONTRATOS <ano>.xlsx, o app lê e grava NELAS
   (a mesma base do Power BI). Lá a gravação é cirúrgica no XML, nunca pelo
   ExcelJS — ver o bloco "Modo base do Power BI" (parseBase/gravarBase).
   O resto deste arquivo é o layout antigo (MATRIZ DE CONTRATOS), mantido.

   Estratégia do layout antigo (ver docs/ARQUITETURA.md):
   - "Abrir planilha" pede acesso a uma PASTA (não só o arquivo), porque
     assim conseguimos, sem diálogo extra, resolver também o snapshot de
     distribuição na mesma pasta, com uma única concessão de permissão.
   - Escrita atualiza só as células das colunas gerenciadas (ver COLUNAS_BASE e
     buildColunasGeridas abaixo); todas as outras colunas/abas da planilha
     (TOTAL ANUAL 2024/2025, "Contratos P&D", "Tabela de prazo...") são
     preservadas intactas porque nunca são tocadas.
   - As colunas "MÊS $" e os dois "Total Faturado <ano>" são gravados como
     FÓRMULA (não como valor): assim a planilha continua viva — mudou a cotação
     do mês na linha 2, o Excel recalcula tudo sozinho.
   - Salvar também (re)aplica a FORMATAÇÃO da planilha (aplicarFormatacao):
     a Hendrya abre no Excel um arquivo apresentável, e ele não degrada a
     cada salvamento.
   - "Excluir" um contrato não remove a linha física (isso desalinharia
     colunas não gerenciadas de outras linhas em algumas versões de
     planilha) — em vez disso limpa os campos gerenciados da linha,
     inclusive Razão Social, que é o sinal usado no parse para pular a
     linha (mesma regra do Power Query: RemoveBlankRows).

   ⚠️ Por que ExcelJS e não SheetJS (trocado na v1.3): o SheetJS community
   NÃO grava estilo — cor de fundo, negrito, borda e painel congelado são
   descartados na escrita, e as larguras de coluna se perdiam. Como o
   objetivo passou a ser entregar uma planilha formatada, ele deixou de
   servir. Ver MAPA_DO_PROJETO.md, seção 9.

   ⚠️ Datas e fuso: o ExcelJS devolve a data da célula como um Date em UTC
   (29/10/2026 vira 2026-10-29T00:00:00Z, cujo DIA LOCAL em UTC-3 é 28).
   Ler/gravar sem normalizar desloca o vencimento em um dia e muda o status
   calculado do contrato. Por isso toda data passa por cellToDate/dateToCell.
   ============================================================ */
(function (global) {
  'use strict';

  var COTACAO_ROW_INDEX0 = 1; // linha 2 (0-based = 1) — cotação do dólar de cada mês
  var HEADER_ROW_INDEX0 = 2; // linha 3 (0-based = 2)
  var FIRST_DATA_ROW_INDEX0 = 3; // linha 4 (0-based = 3)
var MAIN_FILE_NAME_HINT = 'matriz de contratos';
// V1.6.0 — modo "planilha por ano": além da planilha única multi-ano, o app
// aceita uma pasta com UM arquivo por exercício ("MATRIZ DE CONTRATOS 2025.xlsx",
// "MATRIZ DE CONTRATOS 2026.xlsx"...). O exercício do arquivo vem do NOME.
var ANUAL_FILE_RE = /^matriz de contratos\s*(\d{4})\.xlsx$/i;
function anoDoArquivo(fileName) {
  var m = ANUAL_FILE_RE.exec(fileName || '');
  return m ? m[1] : null;
}
  var SNAPSHOT_FILE_NAME = 'Dashboard_Contratos_snapshot.html';

  // Paleta do design system (mesma de theme.css)
  var COR_PRIMARIA = 'FFF58220';
  var COR_PRIMARIA_ESCURA = 'FFB35C10';
  var COR_TEXTO_CLARO = 'FFFFFFFF';
  var COR_ZEBRA = 'FFFAFAFA';
  var COR_BORDA = 'FFE5E5E5';
  var COR_MUTED = 'FF7A7F87';
  var COR_FAIXA_COTACAO = 'FFFDF1E6';

  var FMT_BRL = 'R$ #,##0.00';
  var FMT_USD = '$ #,##0.00';
  var FMT_DATA = 'dd/mm/yyyy';

  // Modo de escrita de cada coluna:
  // - 'value'   : o app grava o valor digitado (comportamento histórico)
  // - 'formula' : o app grava uma FÓRMULA do Excel, para a planilha continuar
  //               viva (mudou a cotação da linha 2 no Excel, tudo recalcula)
  // - 'read'    : o app lê e nunca escreve
  //
  // `fmt`, `width` e `mes` alimentam aplicarFormatacao — uma fonte de verdade só.
  // header da planilha (normalizado) -> campo interno
  //
  // Os meses e os totais são colunas por ANO (V1.4): cada ano com coluna mensal
  // ocupa um bloco REGULAR de 26 colunas — 24 mensais (12 R$ + 12 $) seguidas de
  // "Total Faturado <ano> R$" e "Total Faturado <ano> $". Os anos são DETECTADOS
  // no cabeçalho ao abrir (a Hendrya pode acrescentar "JAN 2027 R$" no Excel e o
  // app pega sozinho). Os campos de 2026 mantêm os nomes históricos; anos novos
  // ganham sufixo.
  //
  // ⚠️ O header do total TEM que levar o ano. Até a V1.4 havia uma coluna única
  // "Total Faturado" (do ano em curso): ao inserir um ano novo, o header genérico
  // voltava a casar com a coluna do ano anterior e o Salvar seguinte a
  // sobrescrevia com o SUM do ano novo — perda silenciosa do total antigo.
  var COLUNAS_BASE = [
    { header: 'razao social', field: 'razaoSocial', type: 'text', mode: 'value', width: 42 },
    { header: 'tipo de servico', field: 'tipoServico', type: 'text', mode: 'value', width: 30 },
    { header: 'departamento', field: 'departamento', type: 'text', mode: 'value', width: 18 },
    { header: 'prazo pagamento', field: 'prazoPagamentoRaw', type: 'text', mode: 'value', width: 15 },
    { header: 'prazo negociado', field: 'prazoNegociadoRaw', type: 'text', mode: 'value', width: 15 },
    { header: 'valor mensal do contrato', alias: 'valor mensal 2', field: 'valorMensal2026', type: 'number', mode: 'value', width: 16, fmt: FMT_BRL },
    { header: 'valor anual do contrato', alias: 'valor anual2', field: 'valorAnual2026', type: 'number', mode: 'value', width: 16, fmt: FMT_BRL },
    { header: 'vigencia', field: 'vigencia', type: 'text', mode: 'value', width: 13 },
    { header: 'inicio', field: 'dataInicio', type: 'date', mode: 'value', width: 12, fmt: FMT_DATA },
    { header: 'fim', field: 'dataFim', type: 'date', mode: 'value', width: 12, fmt: FMT_DATA },
    { header: 'status.', field: 'statusManual', type: 'text', mode: 'value', width: 14 },
    { header: 'comentarios', field: 'comentarios', type: 'text', mode: 'value', width: 40 }
  ];

  // Headers dos dois totais que fecham o bloco de um ano. Usados tanto para
  // casar a coluna existente quanto para escrever o cabeçalho de um ano novo.
  function headerTotalBrl(ano) { return 'Total Faturado ' + ano + ' R$'; }
  function headerTotalUsd(ano) { return 'Total Faturado ' + ano + ' $'; }

  // Rótulo da linha 1, na mesma grafia que a planilha já usa desde a V1.2.
  function rotuloCotacao(mes, ano) {
    return 'cotaçao do dolar ' + mes.nomeLongo + ' ' + ano;
  }

  function buildColunasGeridas(anos) {
    var out = COLUNAS_BASE.slice();
    anos.forEach(function (ano) {
      Regras.MESES.forEach(function (m) {
        out.push({
          header: m.label + ' ' + ano + ' R$',
          alias: ano === '2026' ? [m.headerBrl, m.headerLegado] : null,
          field: Regras.fieldBrlAno(ano, m.key),
          type: 'number', mode: 'value', width: 15, fmt: FMT_BRL, mes: true, ano: ano
        });
        out.push({
          header: m.label + ' ' + ano + ' $',
          alias: ano === '2026' ? [m.headerUsd, m.headerLegado] : null,
          field: Regras.fieldUsdAno(ano, m.key),
          type: 'number', mode: 'formula', width: 15, fmt: FMT_USD, mes: true, ano: ano
        });
      });
      // Os totais ficam FORA do grupo recolhível (sem `mes: true`): é o que faz
      // o Excel fechar um [+] por ano em vez de um único grupo gigante, e o que
      // mantém o total visível com os meses recolhidos.
      out.push({
        header: headerTotalBrl(ano),
        alias: ano === '2026' ? ['total anual reais', 'total anual', 'total faturado'] : null,
        field: Regras.fieldTotalAno(ano),
        type: 'number', mode: 'formula', width: 18, fmt: FMT_BRL, ano: ano, total: true
      });
      out.push({
        header: headerTotalUsd(ano),
        alias: ano === '2026' ? ['total anual dolares'] : null,
        field: Regras.fieldTotalUsdAno(ano),
        type: 'number', mode: 'formula', width: 18, fmt: FMT_USD, ano: ano, total: true
      });
    });
    return out;
  }

  // Colunas gerenciadas VIGENTES — reconstruídas a cada abertura conforme os
  // anos detectados no cabeçalho (nunca editar como constante).
  var colunasGeridas = buildColunasGeridas(Regras.ANOS_MENSAIS);

  // normaliza removendo acentos/pontuação para casar cabeçalhos como "Valor Anual2"
  // mesmo com pequenas variações de grafia/espaçamento na planilha de origem.
  // "R$" e "$" viram marcadores próprios ANTES de a pontuação ser descartada:
  // sem isso "JAN $" e o header legado "JAN" cairiam os dois em "jan".
  function normHeaderLoose(s) {
    return Regras.stripAccents(String(s == null ? '' : s))
      .toLowerCase()
      .replace(/r\$/g, ' rs ')
      .replace(/\$/g, ' usd ')
      .replace(/[^a-z0-9]+/g, ' ')
      .trim();
  }

  var supported = !!(global.showDirectoryPicker);
  var desktop = global.DesktopBridge || null;
  var nativeMode = !!desktop;

  function isSupported() { return nativeMode || supported; }

  // ------------------------------------------------------------
  // Estado do módulo
  // ------------------------------------------------------------
  var state = {
    dirHandle: null,
    fileHandle: null,
    filePath: null,       // modo desktop: caminho absoluto do .xlsx aberto
    fileName: null,
    workbook: null,       // objeto ExcelJS (mantido em memória para preservar outras abas)
    sheetName: null,
    lastReadBuffer: null, // bytes originais da última leitura bem-sucedida
    headerMap: null,      // {field: colIndex0}
    anosMensais: null,    // anos com colunas mensais, detectados no cabeçalho (ex.: ['2026','2025','2024'])
    lastDataRow0: null,   // índice 0-based da última linha de dados conhecida
    contratos: null,      // último parse (usado por testes/smoke)
    cotacoes: null,       // {ano: {jan: 4, fev: 5.09, ...}} — ver readCotacoes
    migracoes: null,      // o que migrarLayout ajustou na última abertura
    // V1.6.0 — modo "planilha por ano": mapa ano -> {substate} de cada MATRIZ
    // anual lida na pasta; `estadoAnual=true` quando o modo está ativo e o
    // singleton acima espelha o substate de `anoAtivo` (mesmo workbook, headerMap,
    // lastReadBuffer... — quem troca de ano chama ativarAno()).
    anosPlanilhas: null,
    estadoAnual: false,
    anoAtivo: null,
    // V1.7.0 — modo "base do Power BI" (MATRIZ DE POWERBI CONTRATOS <ano>.xlsx):
    // usa o mesmo mapa de substates do modo anual, cada um com `base` (linhas,
    // colunas, valores lidos, data de referência) — ver parseBase/saveBase.
    modoBase: false
  };

  function anosMensais() { return state.anosMensais || Regras.ANOS_MENSAIS; }

  // ------------------------------------------------------------
  // Ponte com o ExcelJS (1-based) a partir dos índices 0-based do app
  // ------------------------------------------------------------
  function cellAt(ws, r0, c0) { return ws.getRow(r0 + 1).getCell(c0 + 1); }
  function colLetter(ws, c0) { return ws.getColumn(c0 + 1).letter; }

  // Desembrulha os formatos de valor do ExcelJS (texto rico, fórmula com
  // resultado, hyperlink) para o valor simples que o parse espera.
  function cellValue(cell) {
    var v = cell.value;
    if (v == null) return undefined;
    if (typeof v === 'object' && !(v instanceof Date)) {
      if (v.richText) return cell.text;
      if (v.formula !== undefined || v.sharedFormula !== undefined) return v.result == null ? undefined : v.result;
      if (v.text !== undefined) return v.text;
      if (v.error !== undefined) return undefined;
    }
    return v;
  }

  function findManagedSheet(workbook) {
    var hit = workbook.worksheets.find(function (w) {
      return normHeaderLoose(w.name).indexOf('gestao de contratos') > -1;
    });
    return hit || workbook.worksheets[0];
  }

  // Descobre quais anos têm bloco de colunas mensais no cabeçalho (padrão
  // "JAN 2027 R$" / "JAN 2027 $"). Cabeçalhos legados "JAN R$"/"JAN $" (sem ano)
  // contam como 2026. É isso que permite o "vice-versa": a Hendrya adiciona um
  // ano no Excel e o app passa a ler/gravar aquele bloco sem recompilar.
  // ⚠️ O casamento é contra as 12 chaves de mês, nunca contra "três letras
  // quaisquer": o header "FIM" também tem três letras e fazia TODA planilha
  // declarar que tem colunas de 2026, inclusive uma que só tivesse 2027.
  function labelsDeMes() {
    return Regras.MESES.map(function (m) { return normHeaderLoose(m.label); });
  }

  function detectAnosMensais(ws) {
    var anos = {};
    var meses = labelsDeMes();
    for (var c0 = 0; c0 < ws.columnCount; c0++) {
      var v = cellValue(cellAt(ws, HEADER_ROW_INDEX0, c0));
      if (v == null || v === '') continue;
      var norm = normHeaderLoose(v);
      var m = norm.match(/^([a-z]{3}) (\d{4})( rs| usd)?$/);
      if (m && meses.indexOf(m[1]) > -1) { anos[m[2]] = true; continue; }
      var leg = norm.match(/^([a-z]{3})( rs| usd)?$/);
      if (leg && meses.indexOf(leg[1]) > -1) anos['2026'] = true;
    }
    var lista = Object.keys(anos).sort().reverse();
    return lista.length ? lista : Regras.ANOS_MENSAIS.slice();
  }

  function headersNormalizados(ws) {
    var out = [];
    for (var c0 = 0; c0 < ws.columnCount; c0++) {
      var v = cellValue(cellAt(ws, HEADER_ROW_INDEX0, c0));
      out.push((v == null || v === '') ? null : normHeaderLoose(v));
    }
    return out;
  }

  function ehHeaderDeMes(norm) {
    if (norm == null) return false;
    var meses = labelsDeMes();
    var m = norm.match(/^([a-z]{3}) (\d{4})( rs| usd)?$/);
    if (m && meses.indexOf(m[1]) > -1) return true;
    var leg = norm.match(/^([a-z]{3})( rs| usd)?$/);
    return !!(leg && meses.indexOf(leg[1]) > -1);
  }

  // Coluna do "DEZ <ano> $" — é ela que marca o fim do bloco mensal do ano.
  function colDezUsd(headers, ano) {
    var c = headers.indexOf(normHeaderLoose('Dez ' + ano + ' $'));
    if (c > -1) return c;
    return ano === '2026' ? headers.indexOf(normHeaderLoose('Dez $')) : -1;
  }

  // MIGRAÇÃO DE LAYOUT (V1.4) — idempotente, roda a cada abertura ANTES do
  // parse. Leva a planilha ao formato regular "bloco de 26 colunas por ano":
  //   1. apaga a coluna "Total Faturado"/"TOTAL ANUAL" solta no meio dos campos
  //      fixos (o total passou a morar no fim do bloco de cada ano);
  //   2. renomeia "Total Anual Reais"/"Total Anual Dolares" — que já ficam logo
  //      depois de DEZ 2026 $ — para "Total Faturado 2026 R$/$" (nada se move);
  //   3. insere o par de totais que faltar nos demais anos.
   // Só mexe no workbook EM MEMÓRIA: nada vai a disco sem um Salvar. Rodar duas
   // vezes não faz nada na segunda.
  function migrarLayout(ws) {
    var acoes = [];

    var headers = headersNormalizados(ws);
    var primeiroMes = -1;
    for (var i = 0; i < headers.length; i++) {
      if (ehHeaderDeMes(headers[i])) { primeiroMes = i; break; }
    }
    if (primeiroMes > -1) {
      for (var j = primeiroMes - 1; j >= 0; j--) {
        if (headers[j] === 'total faturado' || headers[j] === 'total anual') {
          ws.spliceColumns(j + 1, 1);
          acoes.push('removida a coluna "Total Faturado" solta');
          break;
        }
      }
    }

    detectAnosMensais(ws).forEach(function (ano) {
      // recalcula a cada ano: os passos anteriores deslocam colunas à direita
      var hs = headersNormalizados(ws);
      var dez = colDezUsd(hs, ano);
      if (dez < 0) return;
      var alvoBrl = normHeaderLoose(headerTotalBrl(ano));
      var alvoUsd = normHeaderLoose(headerTotalUsd(ano));
      if (hs[dez + 1] === alvoBrl && hs[dez + 2] === alvoUsd) return; // já migrado

      var legado = hs[dez + 1] === 'total anual reais' && hs[dez + 2] === 'total anual dolares';
      // Só desloca colunas se as duas vagas seguintes estiverem ocupadas por
      // algo de verdade. Fora o fim da planilha, isso cobre a coluna vazia que
      // o ExcelJS deixa para trás ao remover uma coluna (o spliceColumns de
      // remoção desloca os valores mas não encurta o columnCount): sem esta
      // checagem o último ano inseriria um par a mais e sobraria uma coluna
      // órfã no extremo direito.
      var vagaLivre = function (i) { return hs[i] == null || hs[i] === ''; };
      if (!legado && !(vagaLivre(dez + 1) && vagaLivre(dez + 2))) {
        ws.spliceColumns(dez + 2, 0, [], []);
      }

      var head = ws.getRow(HEADER_ROW_INDEX0 + 1);
      head.getCell(dez + 2).value = headerTotalBrl(ano);
      head.getCell(dez + 3).value = headerTotalUsd(ano);
      acoes.push((legado ? 'renomeados' : 'criados') + ' os totais de ' + ano);
    });

    return acoes;
  }

  // Casa os cabeçalhos da linha 3 com os campos internos. O `alias` (string ou
  // array) só é usado numa segunda passada, para os campos que o header
  // principal não achou — é o que permite abrir uma cópia antiga da planilha
  // (colunas "JAN R$".."DEZ $" e "TOTAL ANUAL") sem quebrar.
  function buildHeaderMap(ws) {
    var map = {};
    var headers = [];
    for (var c0 = 0; c0 < ws.columnCount; c0++) {
      var v = cellValue(cellAt(ws, HEADER_ROW_INDEX0, c0));
      headers.push({ col: c0, norm: (v == null || v === '') ? null : normHeaderLoose(v) });
    }
    [['header', 'header'], ['alias', 'alias']].forEach(function (par) {
      headers.forEach(function (h) {
        if (h.norm == null) return;
        colunasGeridas.forEach(function (def) {
          var prop = par[0];
          if (def[prop] == null || map[def.field] != null) return;
          var chaves = Array.isArray(def[prop]) ? def[prop] : [def[prop]];
          chaves.forEach(function (ch) {
            if (map[def.field] != null) return;
            if (normHeaderLoose(ch) === h.norm) map[def.field] = h.col;
          });
        });
      });
    });
    return map;
  }

  // `instanceof Date` não vale aqui: a data vem construída pelo ExcelJS, que
  // pode estar em outro realm (é o caso dos testes headless, que rodam o app
  // dentro de um vm context). Comparar a marca interna funciona sempre.
  function isDate(v) {
    return Object.prototype.toString.call(v) === '[object Date]';
  }

  // O ExcelJS entrega a data da célula em UTC. O dia que interessa é o que
  // está escrito na planilha, então lemos pelas partes UTC e montamos uma
  // data local à meia-noite (o resto do app trabalha em data local).
  function cellToDate(v) {
    if (v == null || v === '') return null;
    if (isDate(v)) {
      if (isNaN(v.getTime())) return null;
      return new Date(v.getUTCFullYear(), v.getUTCMonth(), v.getUTCDate());
    }
    if (typeof v === 'number') return null;
    return Regras.parseDateInput(v);
  }

  // Caminho inverso: grava meia-noite UTC do dia escolhido, para o Excel
  // mostrar exatamente esse dia em qualquer fuso.
  function dateToCell(d) {
    if (!d) return null;
    var dt = isDate(d) ? d : Regras.parseDateInput(d);
    if (!dt || isNaN(dt.getTime())) return null;
    return new Date(Date.UTC(dt.getFullYear(), dt.getMonth(), dt.getDate()));
  }

  function parseContratos(ws) {
    var headerMap = buildHeaderMap(ws);
    var contratos = [];
    var lastDataRow0 = FIRST_DATA_ROW_INDEX0 - 1;

    for (var r0 = FIRST_DATA_ROW_INDEX0; r0 < ws.rowCount; r0++) {
      var get = function (field) {
        if (headerMap[field] == null) return undefined;
        return cellValue(cellAt(ws, r0, headerMap[field]));
      };

      var razaoRaw = get('razaoSocial');
      var linhaVazia = colunasGeridas.every(function (def) {
        var v = get(def.field);
        return v == null || v === '';
      });
      if (linhaVazia) continue;
      if (razaoRaw == null || String(razaoRaw).trim() === '') continue;

      var prazoRaw = get('prazoPagamentoRaw');
      var c = {
        _row0: r0,
        id: r0 + 1, // número de linha do Excel (1-based) — chave estável
        isNew: false,
        razaoSocial: Regras.normalizeRazaoSocial(razaoRaw),
        tipoServico: Regras.normalizeTipoServico(get('tipoServico')),
        departamento: Regras.normalizeDepartamento(get('departamento')),
        prazoPagamentoRaw: prazoRaw == null ? '' : String(prazoRaw).trim(),
        prazoPagamentoDias: Regras.parsePrazoPagamentoDias(prazoRaw),
        prazoNegociadoRaw: (function () { var v = get('prazoNegociadoRaw'); return v == null ? '' : String(v).trim(); })(),
        valorMensal2026: (function () { var v = get('valorMensal2026'); return (v === '' || v == null) ? null : Number(v); })(),
        valorAnual2026: (function () { var v = get('valorAnual2026'); return (v === '' || v == null) ? null : Number(v); })(),
        vigencia: (function () { var v = get('vigencia'); return v == null ? '' : String(v).trim(); })(),
        dataInicio: cellToDate(get('dataInicio')),
        dataFim: cellToDate(get('dataFim')),
        comentarios: (function () { var v = get('comentarios'); return v == null ? '' : String(v).trim(); })(),
        // "Status Manual" (override ativo) SEMPRE começa em null (Automático) ao abrir
        // a planilha — mesmo que a coluna STATUS. legada tenha um valor. Em planilhas
        // legadas mantidas manualmente essa coluna tende a ficar desatualizada frente
        // à data real de fim. O texto original fica em statusLegadoRaw só para
        // referência/auditoria (ver Regras.qualityFlags) — nunca para decidir o
        // status exibido.
        statusManual: null,
        statusLegadoRaw: (function () { var v = get('statusManual'); return v == null ? '' : String(v).trim(); })()
      };
      // Campos numéricos derivados (meses por ano, Total Faturado por ano e os
      // totais legados) — tudo resolvido pelo headerMap, nunca por coluna fixa.
      var camposNumericos = [];
      colunasGeridas.forEach(function (def) {
        if (def.type !== 'number') return;
        if (def.field === 'valorMensal2026' || def.field === 'valorAnual2026') return;
        if (camposNumericos.indexOf(def.field) === -1) camposNumericos.push(def.field);
      });
      camposNumericos.forEach(function (f) {
        var v = get(f);
        c[f] = (v === '' || v == null || isNaN(Number(v))) ? null : Number(v);
      });
      contratos.push(c);
      lastDataRow0 = r0;
    }

    state.headerMap = headerMap;
    state.lastDataRow0 = lastDataRow0;
    return contratos;
  }

  // ------------------------------------------------------------
  // Abrir pasta + planilha
  // ------------------------------------------------------------
  function abrirPastaHandle(dirHandle) {
    resetarModoAnual();
    state.dirHandle = dirHandle;
    // V1.7.0: a base do Power BI (pasta e subpastas) tem prioridade sobre o layout antigo.
    return listarBaseDir(dirHandle).then(function (base) {
      if (base.length) return abrirBase(base);
      return listarAnuaisDir(dirHandle).then(function (lista) {
        if (lista.length) return abrirAnuais(lista);
        // V1.4 (back-compat): cadeia única com blocos multi-ano numa só planilha.
        return findXlsxHandle(dirHandle).then(function (fileHandle) {
          state.fileHandle = fileHandle;
          state.fileName = fileHandle.name;
          return loadFromFileHandle(fileHandle);
        });
      });
    });
  }

  function openFolder() {
    if (nativeMode) return pickDesktop();
    if (!supported) {
      return Promise.reject(new Error('Este navegador não suporta a File System Access API (necessária para editar). Use Microsoft Edge ou Google Chrome.'));
    }
    return global.showDirectoryPicker({ mode: 'readwrite' }).then(function (dirHandle) {
      guardarPastaPrincipal(dirHandle);
      return abrirPastaHandle(dirHandle);
    });
  }

  // V1.9.1 — "abrir automático": a última pasta escolhida fica no IndexedDB (File System
  // Access API guarda o handle). No próximo carregamento, se a permissão já foi concedida
  // ('granted'), a planilha abre sozinha; senão o app só avisa para dar um clique.
  var IDB_NAME = 'dashboard-contratos';
  var IDB_STORE = 'navegacao';
  var IDB_KEY_PASTA = 'pastaPrincipal';

  function abrirIdb() {
    return new Promise(function (resolve, reject) {
      var req = global.indexedDB.open(IDB_NAME, 1);
      req.onupgradeneeded = function () {
        if (!req.result.objectStoreNames.contains(IDB_STORE)) req.result.createObjectStore(IDB_STORE);
      };
      req.onsuccess = function () { resolve(req.result); };
      req.onerror = function () { reject(req.error || new Error('IndexedDB indisponível.')); };
    });
  }

  function idbExecuta(tx, cb) {
    return new Promise(function (resolve, reject) {
      tx.oncomplete = resolve;
      tx.onerror = function () { reject(tx.error); };
      cb(tx);
    });
  }

  function guardarPastaPrincipal(handle) {
    if (nativeMode || !handle) return Promise.resolve();
    return abrirIdb().then(function (db) {
      return idbExecuta(db.transaction(IDB_STORE, 'readwrite'), function (tx) {
        tx.objectStore(IDB_STORE).put(handle, IDB_KEY_PASTA);
      }).then(function () { db.close(); });
    }).catch(function (e) { console.warn('[excel] não persistiu a última pasta:', e); });
  }

  function recuperarPastaPrincipal() {
    if (nativeMode || !global.indexedDB) return Promise.resolve(null);
    return abrirIdb().then(function (db) {
      return new Promise(function (resolve, reject) {
        var req;
        try {
          var tx = db.transaction(IDB_STORE, 'readonly');
          req = tx.objectStore(IDB_STORE).get(IDB_KEY_PASTA);
          req.onsuccess = function () { resolve(req.result || null); };
          req.onerror = function () { reject(req.error); };
        } catch (e) { reject(e); }
      }).then(function (v) { db.close(); return v; });
    }).catch(function () { return null; });
  }

  // Sem exigir gesto do usuário quando a permissão já está concedida. Resultado:
  // null (nada gravado) | { precisaGestor: true } | resultado do abrirPastaHandle.
  function abrirPastaPrincipal() {
    if (nativeMode) return Promise.resolve(null);
    if (!supported) return Promise.resolve(null);
    return recuperarPastaPrincipal().then(function (handle) {
      if (!handle) return null;
      return handle.queryPermission({ mode: 'readwrite' }).then(function (perm) {
        if (perm !== 'granted') return { precisaGestor: true };
        return abrirPastaHandle(handle);
      });
    }).catch(function (e) {
      console.warn('[excel] reabertura automática falhou:', e);
      return null;
    });
  }

  // V1.6.0 — todos os "MATRIZ DE CONTRATOS <ano>.xlsx" da pasta, em ordem de
  // exercício decrescente. Arquivos de outro nome seguem o fluxo antigo.
  function listarAnuaisDir(dirHandle) {
    var lista = [];
    return (async function () {
      for await (var entry of dirHandle.values()) {
        if (entry.kind !== 'file' || !/\.xlsx$/i.test(entry.name) || /~\$/.test(entry.name)) continue;
        var ano = anoDoArquivo(entry.name);
        if (ano) lista.push({ ano: ano, fileHandle: entry });
      }
      lista.sort(function (a, b) { return a.ano < b.ano ? 1 : -1; });
      return lista;
    })();
  }

  // Nova abertura sempre começa sem o modo anual — o fluxo seguinte decide se
  // a pasta tinha MATRIZ por exercício (e rearma) ou se segue no modo clássico.
  function resetarModoAnual() {
    state.anosPlanilhas = null;
    state.estadoAnual = false;
    state.anoAtivo = null;
    state.modoBase = false;
  }

  // Lê cada arquivo anual (navegador), guarda um substate por ano e arma o modo.
  function abrirAnuais(lista) {
    var mapa = {};
    return lista.reduce(function (prev, item) {
      return prev
        .then(function () { return item.fileHandle.getFile(); })
        .then(function (file) { return file.arrayBuffer(); })
        .then(function (buffer) { return parseWorkbook(buffer, item.fileHandle.name); })
        .then(function (modelo) {
          modelo.fileHandle = item.fileHandle; // substate guarda o handle do SEU ano
          mapa[item.ano] = modelo;
        });
    }, Promise.resolve()).then(function () {
      return armarModoAnual(mapa);
    });
  }

  // Lê cada arquivo anual (desktop) e arma o modo.
  function abrirAnuaisDesktop(dir, anuais) {
    var mapa = {};
    return anuais.reduce(function (prev, item) {
      var caminho = null;
      return prev
        .then(function () { return desktop.join(dir, item.nome); })
        .then(function (p) { caminho = p; return desktop.read(p); })
        .then(function (res) {
          if (!res || !res.ok) return Promise.reject(new Error('Não foi possível ler ' + item.nome + '.'));
          return parseWorkbook(res.data, item.nome);
        })
        .then(function (modelo) {
          modelo.filePath = caminho; // substate guarda o caminho do SEU ano
          mapa[item.ano] = modelo;
        });
    }, Promise.resolve()).then(function () {
      return armarModoAnual(mapa);
    });
  }

  // Ativa o modo anual e devolve o resultado de abertura para a app: o exercício
  // ativo é o ano corrente se ele tiver planilha, senão o mais recente.
  function armarModoAnual(mapa) {
    state.anosPlanilhas = mapa;
    state.estadoAnual = true;
    var anos = Object.keys(mapa).sort().reverse();
    var alvo = String(new Date().getFullYear());
    var ativo = anos.indexOf(alvo) > -1 ? alvo : anos[0];
    state.anoAtivo = ativo;
    // União dos anos internos de todos os arquivos (para snapshot cobrir tudo),
    // SEM reassign o ano ativo: ele vem só dos nomes de arquivo.
    var uniao = {};
    Object.keys(mapa).forEach(function (ano) {
      uniao[ano] = true;
      (mapa[ano].anosMensais || []).forEach(function (a) { uniao[a] = true; });
    });
    var anosUnion = Object.keys(uniao).sort().reverse();
    return ativarAno(ativo).then(function (res) {
      var porArquivo = {};
      Object.keys(mapa).forEach(function (ano) { porArquivo[ano] = { fileName: mapa[ano].fileName }; });
      return {
        contratos: res.contratos,
        fileName: res.fileName,
        cotacoes: res.cotacoes,
        anosMensais: anosUnion,
        anoAtivo: ativo,
        anosPlanilhas: porArquivo,
        estadoAnual: true,
        modoBase: res.modoBase,
        saving: res.saving,
        temSaving: res.temSaving
      };
    });
  }

  // Espelha no singleton `state` (e no módulo `colunasGeridas`) o substate do
  // ano escolhido — é o que deixa save()/ensureYear()/formatação funcionarem
  // sem saber que estão dentro do modo anual. O workbook fica em memória por
  // ano, então trocar de exercício não recarrega nada do disco.
  function espelharSubstate(sub) {
    state.workbook = sub.workbook;
    state.sheetName = sub.sheetName;
    state.fileName = sub.fileName;
    state.filePath = sub.filePath || null;
    state.fileHandle = sub.fileHandle || null;
    state.headerMap = sub.headerMap;
    state.anosMensais = sub.anosMensais.slice();
    colunasGeridas = sub.colunasGeridas.slice();
    state.contratos = sub.contratos;
    state.cotacoes = sub.cotacoes;
    state.lastReadBuffer = sub.lastReadBuffer;
    state.lastDataRow0 = sub.lastDataRow0;
    state.migracoes = sub.migracoes;
  }

  // Troca de exercício no modo anual. Devolve o modelo do ano, para a app
  // trocar CONTRATOS/COTACOES/ANOS_MENSAL junto.
  function ativarAno(ano) {
    var anoStr = String(ano);
    var sub = (state.anosPlanilhas || {})[anoStr];
    if (!state.estadoAnual || !sub || !sub.workbook) {
      return Promise.reject(new Error('Não há planilha para ' + anoStr + '.'));
    }
    espelharSubstate(sub);
    state.anoAtivo = anoStr;
    var porArquivo = {};
    Object.keys(state.anosPlanilhas).forEach(function (a) { porArquivo[a] = { fileName: state.anosPlanilhas[a].fileName }; });
    return Promise.resolve({
      contratos: sub.contratos,
      fileName: sub.fileName,
      cotacoes: sub.cotacoes,
      anosMensais: Object.keys(state.anosPlanilhas).sort().reverse(),
      anoAtivo: anoStr,
      anosPlanilhas: porArquivo,
      modoBase: !!sub.base,
      saving: sub.saving || [],
      temSaving: !!(sub.base && sub.base.saving)
    });
  }

  // ------------------------------------------------------------
  // V1.7.0 — Modo "base do Power BI"
  //
  // A pasta escolhida (ou qualquer subpasta dela) tem as planilhas que o Power BI
  // lê: "MATRIZ DE POWERBI CONTRATOS <ano>.xlsx", uma por exercício, no molde do
  // montar_base_pbi.py (tabela tContratos na aba Contratos, cabeçalho na linha 3,
  // linhas livres no fim da tabela). O app lê e grava nelas: os dois painéis usam a
  // mesma base, e as regras de regras.js são as da planilha (mesmos números).
  //
  // ⚠️ NUNCA gravar essa planilha com o ExcelJS. Medido em 06/10/2026: o
  // writeBuffer derruba as 218 fórmulas de matriz dinâmica (a aba KPI inteira, que
  // é o que o Power BI lê), as colunas calculadas da tabela e a formatação
  // condicional. Aqui o ExcelJS só LÊ; gravarBase() abre o .xlsx como zip (JSZip)
  // e troca o valor de cada célula alterada direto no XML da aba Contratos — o
  // resto do arquivo sai idêntico.
  //
  // Quem recalcula status, dólar e KPI é o Excel. Ao gravar, o app escreve no
  // VALOR CALCULADO da conferência "recalculada no Excel" (nomes pHtmlResultado e
  // pHtmlStatus) um ATENÇÃO: o Power BI avisa em vermelho até alguém abrir a
  // planilha no Excel e salvar — o recálculo devolve OK sozinho.
  // ------------------------------------------------------------
  var BASE_FILE_RE = /^matriz de powerbi contratos (\d{4})\.xlsx$/i;
  var BASE_SEM_TIPO = 'Não Classificado'; // o que Regras.normalizeTipoServico põe no lugar do vazio
  var BASE_CAMPOS = [
    { header: 'Código Fornecedor', field: 'codigoFornecedor', type: 'text' },
    { header: 'Razão Social', field: 'razaoSocial', type: 'text' },
    { header: 'Nome curto', field: 'nomeCurto', type: 'text' },
    { header: 'Tipo de Serviço', field: 'tipoServico', type: 'text' },
    { header: 'Departamento', field: 'departamento', type: 'text' },
    { header: 'Prazo Pagamento', field: 'prazoPagamentoRaw', type: 'text' },
    { header: 'Prazo Negociado', field: 'prazoNegociadoRaw', type: 'text' },
    { header: 'Valor Mensal do Contrato', field: 'valorMensal2026', type: 'number' },
    { header: 'Valor Anual do Contrato', field: 'valorAnual2026', type: 'number' },
    { header: 'Vigência', field: 'vigencia', type: 'text' },
    { header: 'Início', field: 'dataInicio', type: 'date' },
    { header: 'Fim', field: 'dataFim', type: 'date' },
    { header: 'Comentários', field: 'comentarios', type: 'text' }
  ];

  // Colunas de entrada (amarelas) da planilha do ano: as fixas + "JAN R$".."DEZ R$",
  // que viram os campos do exercício do arquivo (Regras.fieldBrlAno).
  function camposBase(ano) {
    return BASE_CAMPOS.concat(Regras.MESES.map(function (m) {
      return { header: m.label.toUpperCase() + ' R$', field: Regras.fieldBrlAno(ano, m.key), type: 'number' };
    }));
  }

  // V1.9.0 — aba Saving (tabela tSaving, molde v3): uma linha por BID/negociação, saving em
  // R$ por mês + total sem detalhe mensal. As colunas cinzas (US$, Desde, totais) são fórmula.
  var BASE_SAVING_CAMPOS = [
    { header: 'Código Fornecedor', field: 'codigoFornecedor', type: 'text' },
    { header: 'Prestador', field: 'prestador', type: 'text' },
    { header: 'Tipo de Serviço', field: 'tipoServico', type: 'text' },
    { header: 'BID concluído', field: 'bidConcluido', type: 'text' },
    { header: 'Pagamento', field: 'pagamento', type: 'text' }
  ];
  function camposSaving() {
    return BASE_SAVING_CAMPOS.concat(Regras.MESES.map(function (m) {
      return { header: m.label.toUpperCase() + ' R$', field: Regras.campoSaving(m.key), type: 'number' };
    }), [{ header: 'Saving do ano sem mês R$', field: 'savSemMes', type: 'number' }]);
  }

  function letraColuna(n) {
    var s = '';
    while (n > 0) { var r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); }
    return s;
  }
  function colunaDaLetra(letras) {
    var n = 0;
    for (var i = 0; i < letras.length; i++) n = n * 26 + (letras.charCodeAt(i) - 64);
    return n;
  }

  // Linha do cabeçalho e última linha (livres incluídas) da tabela, lidas do XML.
  function lerRefTabela(zip, nome, opcional) {
    var arquivos = Object.keys(zip.files).filter(function (n) { return /^xl\/tables\/[^/]+\.xml$/.test(n); });
    return Promise.all(arquivos.map(function (n) { return zip.file(n).async('string'); })).then(function (xmls) {
      for (var i = 0; i < xmls.length; i++) {
        var tag = /<table\b[^>]*>/.exec(xmls[i]);
        if (tag && new RegExp('\\s(?:name|displayName)="' + nome + '"').test(tag[0])) {
          var ref = /\sref="[A-Z]+(\d+):[A-Z]+(\d+)"/.exec(tag[0]);
          if (ref) return { linhaCab: Number(ref[1]), ultimaLinha: Number(ref[2]) };
        }
      }
      if (opcional) return null; // ex.: tSaving numa planilha anterior ao molde v3
      throw new Error('A planilha não tem a tabela ' + nome + ' (não está no molde do Power BI).');
    });
  }

  // tMeta (aba KPI): 'referencia' (dd/mm/aaaa, a data do status) e 'posicao' (texto do cabeçalho).
  function lerMetaBase(wb) {
    var meta = {};
    var ws = wb.getWorksheet('KPI');
    if (!ws) return meta;
    ws.eachRow(function (row) {
      var chave = cellValue(row.getCell(2));
      if (chave === 'referencia' || chave === 'posicao') meta[chave] = cellValue(row.getCell(3));
    });
    return meta;
  }

  function dataDmy(s) {
    var m = /^(\d{2})\/(\d{2})\/(\d{4})$/.exec(String(s == null ? '' : s).trim());
    return m ? new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1])) : null;
  }

  // Aba Cotações (tabela tCotacoes: MesNum, Mês, Cotação), só leitura.
  function lerCotacoesBase(wb, ano) {
    var porAno = {};
    var ws = wb.getWorksheet('Cotações');
    if (ws) {
      var cab = {};
      ws.getRow(HEADER_ROW_INDEX0 + 1).eachCell(function (cell, n) { cab[normHeaderLoose(cellValue(cell))] = n; });
      if (cab.mesnum && cab.cotacao) {
        for (var r = HEADER_ROW_INDEX0 + 2; r <= HEADER_ROW_INDEX0 + 13; r++) {
          var mes = Number(cellValue(ws.getRow(r).getCell(cab.mesnum)));
          var v = Number(cellValue(ws.getRow(r).getCell(cab.cotacao)));
          if (mes >= 1 && mes <= 12 && v > 0) porAno[Regras.MESES[mes - 1].key] = v;
        }
      }
    }
    var out = {};
    out[ano] = porAno;
    return out;
  }

  function contratoBase(b, campos, linha) {
    var texto = function (v) { return v == null ? '' : String(v).trim(); };
    var num = function (v) { return (v == null || isNaN(Number(v))) ? null : Number(v); };
    // Data digitada como texto ("28/03/205") fica vazia aqui, como na planilha (Fim sem data =
    // Em Análise); o texto continua na célula, porque o Salvar só grava campo alterado.
    var data = function (v) { var d = cellToDate(v); return d && !isNaN(d.getTime()) ? d : null; };
    var c = {
      _row0: linha - 1,
      id: linha, // linha do Excel — chave estável, como no modo clássico
      isNew: false,
      codigoFornecedor: texto(b.codigoFornecedor),
      razaoSocial: Regras.normalizeRazaoSocial(b.razaoSocial),
      nomeCurto: texto(b.nomeCurto),
      tipoServico: Regras.normalizeTipoServico(b.tipoServico),
      departamento: Regras.normalizeDepartamento(b.departamento),
      prazoPagamentoRaw: texto(b.prazoPagamentoRaw),
      prazoPagamentoDias: Regras.parsePrazoPagamentoDias(b.prazoPagamentoRaw),
      prazoNegociadoRaw: texto(b.prazoNegociadoRaw),
      valorMensal2026: num(b.valorMensal2026),
      valorAnual2026: num(b.valorAnual2026),
      vigencia: texto(b.vigencia),
      dataInicio: data(b.dataInicio),
      dataFim: data(b.dataFim),
      comentarios: texto(b.comentarios),
      statusManual: null, // na base o status é fórmula da planilha
      statusLegadoRaw: ''
    };
    campos.forEach(function (d) { if (d.field.indexOf('mes') === 0) c[d.field] = num(b[d.field]); });
    return c;
  }

  // Forma comparável de um campo: o Salvar só grava o que mudou desde a abertura.
  function chaveBase(d, v) {
    if (v == null || v === '') return '';
    if (d.type === 'date') return Regras.toDateInputValue(v);
    if (d.type === 'number') return isNaN(Number(v)) ? '' : String(Number(v));
    var s = String(v).trim();
    return (d.field === 'tipoServico' && s === BASE_SEM_TIPO) ? '' : s;
  }
  function valoresBase(c, campos) {
    var o = {};
    campos.forEach(function (d) { o[d.field] = chaveBase(d, c[d.field]); });
    return o;
  }
  function valorParaGravar(d, v) {
    var k = chaveBase(d, v);
    if (k === '') return null;
    if (d.type === 'number') return Number(v);
    if (d.type === 'date') return Regras.parseDateInput(k);
    return k;
  }

  function itemSaving(b, campos, linha) {
    var texto = function (v) { return v == null ? '' : String(v).trim(); };
    var num = function (v) { return (v == null || isNaN(Number(v))) ? null : Number(v); };
    var it = {
      _row0: linha - 1,
      id: linha, // linha do Excel, como nos contratos (lista separada)
      isNew: false,
      codigoFornecedor: texto(b.codigoFornecedor),
      prestador: texto(b.prestador),
      tipoServico: texto(b.tipoServico),
      bidConcluido: texto(b.bidConcluido),
      pagamento: texto(b.pagamento),
      savSemMes: num(b.savSemMes)
    };
    Regras.MESES.forEach(function (m) { var f = Regras.campoSaving(m.key); it[f] = num(b[f]); });
    return it;
  }

  // Uma tabela de entradas (Contratos ou Saving): colunas pelo cabeçalho, itens, linhas livres
  // e os valores lidos de cada linha (é com eles que o Salvar compara para gravar só o que mudou).
  // `est` é a "estrutura" que alteracoesTabela/aplicarGravado usam.
  function lerTabelaBase(ws, ref, campos, campoNome, montar, fileName, aba) {
    var colDe = {};
    ws.getRow(ref.linhaCab).eachCell(function (cell, n) {
      var h = normHeaderLoose(cellValue(cell));
      campos.forEach(function (d) { if (colDe[d.field] == null && normHeaderLoose(d.header) === h) colDe[d.field] = n; });
    });
    var faltam = campos.filter(function (d) { return colDe[d.field] == null; }).map(function (d) { return d.header; });
    if (faltam.length) {
      throw new Error(fileName + ' (aba ' + aba + '): faltam as colunas ' + faltam.join(', ') + ' (a planilha não está no molde do Power BI).');
    }
    var itens = [], livres = [], originais = {}, orfas = 0;
    for (var linha = ref.linhaCab + 1; linha <= ref.ultimaLinha; linha++) {
      var row = ws.getRow(linha);
      var bruto = {}, algum = false;
      campos.forEach(function (d) {
        var v = cellValue(row.getCell(colDe[d.field]));
        if (typeof v === 'string' && v.trim() === '') v = undefined;
        bruto[d.field] = v;
        if (v != null) algum = true;
      });
      if (bruto[campoNome] == null) {
        if (algum) orfas++; else livres.push(linha); // linha com dado e sem nome: nunca é reaproveitada
        continue;
      }
      var it = montar(bruto, campos, linha);
      originais[linha] = valoresBase(it, campos);
      itens.push(it);
    }
    return {
      itens: itens,
      est: { aba: aba, campos: campos, colDe: colDe, ultimaLinha: ref.ultimaLinha, livres: livres, originais: originais, orfas: orfas }
    };
  }

  function parseBase(arrayBuffer, fileName) {
    var ano = String(BASE_FILE_RE.exec(fileName)[1]);
    var wb = new ExcelJS.Workbook();
    var refs = null;
    return JSZip.loadAsync(arrayBuffer).then(function (zip) {
      return Promise.all([lerRefTabela(zip, 'tContratos'), lerRefTabela(zip, 'tSaving', true)]);
    }).then(function (r) {
      refs = r;
      return wb.xlsx.load(arrayBuffer);
    }).then(function () {
      var ws = wb.getWorksheet('Contratos');
      if (!ws) throw new Error(fileName + ': não achei a aba Contratos.');
      var con = lerTabelaBase(ws, refs[0], camposBase(ano), 'razaoSocial', contratoBase, fileName, 'Contratos');
      var wsS = refs[1] ? wb.getWorksheet('Saving') : null;
      var sav = wsS ? lerTabelaBase(wsS, refs[1], camposSaving(), 'prestador', itemSaving, fileName, 'Saving') : null;
      var meta = lerMetaBase(wb);
      var cotacoes = lerCotacoesBase(wb, ano);
      var base = con.est;
      base.referencia = dataDmy(meta.referencia);
      base.posicao = meta.posicao ? String(meta.posicao) : '';
      base.saving = sav ? sav.est : null; // null = planilha sem a aba Saving (molde anterior ao v3)
      state.headerMap = base.colDe;
      state.lastDataRow0 = null;
      state.contratos = con.itens;
      state.cotacoes = cotacoes;
      return {
        ano: ano,
        fileName: fileName,
        workbook: wb,
        sheetName: 'Contratos',
        headerMap: base.colDe,
        anosMensais: [ano],
        colunasGeridas: [],
        contratos: con.itens,
        saving: sav ? sav.itens : [],
        cotacoes: cotacoes,
        lastReadBuffer: arrayBuffer,
        lastDataRow0: null,
        migracoes: [],
        base: base
      };
    });
  }

  // Todas as planilhas da base na pasta e nas subpastas (a mesma regra do Power BI).
  function listarBaseDir(dirHandle) {
    var lista = [];
    return (async function varrer(dir, caminho, nivel) {
      for await (var entry of dir.values()) {
        if (entry.kind === 'directory') {
          if (nivel < 6) await varrer(entry, caminho + entry.name + '/', nivel + 1);
        } else {
          var m = BASE_FILE_RE.exec(entry.name);
          if (m) lista.push({ ano: m[1], fileHandle: entry, pasta: dir, caminho: caminho + entry.name });
        }
      }
      return lista;
    })(dirHandle, '', 0);
  }

  function abrirBase(lista) {
    var porAno = {};
    lista.forEach(function (it) { (porAno[it.ano] = porAno[it.ano] || []).push(it); });
    var repetidos = Object.keys(porAno).filter(function (a) { return porAno[a].length > 1; });
    if (repetidos.length) {
      return Promise.reject(new Error('Mais de uma planilha do mesmo ano: ' + repetidos.map(function (a) {
        return porAno[a].map(function (i) { return i.caminho; }).join(' e ');
      }).join('; ') + '. Deixe uma por ano (uma cópia leva outro nome) — o Power BI também acusa isso.'));
    }
    var mapa = {};
    return lista.reduce(function (prev, item) {
      var arquivo = null;
      return prev
        .then(function () { return item.fileHandle.getFile(); })
        .then(function (file) { arquivo = file; return file.arrayBuffer(); })
        .then(function (buffer) { return parseBase(buffer, item.fileHandle.name); })
        .then(function (modelo) {
          modelo.fileHandle = item.fileHandle;
          modelo.base.pasta = item.pasta;
          modelo.base.caminho = item.caminho;
          modelo.base.lastModified = arquivo.lastModified;
          modelo.base.size = arquivo.size;
          mapa[item.ano] = modelo;
        });
    }, Promise.resolve()).then(function () {
      state.modoBase = true;
      return armarModoAnual(mapa);
    });
  }

  // O que gravar numa tabela (est = Contratos ou Saving): só células de entrada que mudaram.
  // Item novo ocupa a 1ª linha livre; excluído tem as entradas limpas (a linha volta a ser livre).
  function alteracoesTabela(est, itens, removidos0) {
    var mudancas = [];
    var livres = est.livres.slice();
    var novasLinhas = {};
    (removidos0 || []).forEach(function (r0) {
      var linha = r0 + 1;
      var antes = est.originais[linha];
      if (!antes) return;
      var celulas = est.campos.filter(function (d) { return antes[d.field] !== ''; }).map(function (d) {
        return { col: est.colDe[d.field], type: d.type, value: null };
      });
      if (celulas.length) mudancas.push({ aba: est.aba, linha: linha, celulas: celulas });
    });
    (itens || []).forEach(function (c) {
      var linha, antes;
      if (c.isNew || c._row0 == null) {
        linha = livres.shift();
        if (linha == null) {
          throw new Error('A tabela da aba ' + est.aba + ' não tem mais linha livre. Abra a planilha no Excel e digite logo abaixo da tabela.');
        }
        antes = null;
        novasLinhas[c.id] = linha;
      } else {
        linha = c._row0 + 1;
        antes = est.originais[linha];
      }
      var celulas = [];
      est.campos.forEach(function (d) {
        if (chaveBase(d, c[d.field]) !== (antes ? antes[d.field] : '')) {
          celulas.push({ col: est.colDe[d.field], type: d.type, value: valorParaGravar(d, c[d.field]) });
        }
      });
      if (celulas.length) mudancas.push({ aba: est.aba, linha: linha, celulas: celulas });
    });
    return { mudancas: mudancas, novasLinhas: novasLinhas, livresRestantes: livres };
  }

  function alteracoesBase(sub, contratos, deletedRows0) {
    return alteracoesTabela(sub.base, contratos, deletedRows0);
  }

  // Depois de gravar: itens novos ganham a linha, excluídos voltam a ser livres e os valores
  // gravados viram a referência do próximo Salvar.
  function aplicarGravado(est, alt, itens, removidos0) {
    (removidos0 || []).forEach(function (r0) { delete est.originais[r0 + 1]; alt.livresRestantes.push(r0 + 1); });
    est.livres = alt.livresRestantes.filter(function (l, i, a) { return a.indexOf(l) === i; }).sort(function (x, y) { return x - y; });
    (itens || []).forEach(function (c) {
      var linha = alt.novasLinhas[c.id];
      if (linha != null) { c._row0 = linha - 1; c.isNew = false; }
      if (c._row0 != null) est.originais[c._row0 + 1] = valoresBase(c, est.campos);
    });
  }

  // ── Gravação direto no XML (sem ExcelJS) ─────────────────────────────
  function xmlEsc(s) {
    return String(s).replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, '')
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }

  // Data local -> número de série do Excel (sistema 1900): o dia escrito na tela.
  function serialData(d) {
    return (Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()) - Date.UTC(1899, 11, 30)) / 86400000;
  }

  function celulaXml(ref, s, type, value) {
    var st = s != null ? ' s="' + s + '"' : '';
    if (value == null || value === '') return '<c r="' + ref + '"' + st + '/>';
    if (type === 'number') return '<c r="' + ref + '"' + st + '><v>' + String(value) + '</v></c>';
    if (type === 'date') return '<c r="' + ref + '"' + st + '><v>' + serialData(value) + '</v></c>';
    return '<c r="' + ref + '"' + st + ' t="inlineStr"><is><t xml:space="preserve">' + xmlEsc(value) + '</t></is></c>';
  }

  function caminhoDaAba(wbXml, rels, nome) {
    var esc = xmlEsc(nome).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    var sheet = new RegExp('<sheet\\b[^>]*\\bname="' + esc + '"[^>]*/>').exec(wbXml);
    if (!sheet) throw new Error('A planilha não tem a aba ' + nome + '.');
    var rid = /\br:id="([^"]+)"/.exec(sheet[0])[1];
    var rel = new RegExp('<Relationship\\b[^>]*\\bId="' + rid + '"[^>]*/>').exec(rels);
    var alvo = /\bTarget="([^"]+)"/.exec(rel[0])[1];
    return alvo.charAt(0) === '/' ? alvo.slice(1) : 'xl/' + alvo.replace(/^xl\//, '');
  }

  function nomeDefinido(wbXml, nome) {
    var m = new RegExp('<definedName\\b[^>]*\\bname="' + nome + '"[^>]*>([^<]*)</definedName>').exec(wbXml);
    if (!m) return null;
    var p = /^'?(.*?)'?!\$?([A-Z]+)\$?(\d+)$/.exec(m[1].replace(/&apos;/g, "'"));
    return p ? { aba: p[1].replace(/''/g, "'"), ref: p[2] + p[3] } : null;
  }

  // Quais estilos (índice de cellXfs) têm formato de data — data gravada em célula
  // com formato de texto apareceria como número no Excel.
  function estilosDeData(stylesXml) {
    var custom = {};
    (stylesXml.match(/<numFmt\b[^>]*\/>/g) || []).forEach(function (t) {
      var id = /numFmtId="(\d+)"/.exec(t), code = /formatCode="([^"]*)"/.exec(t);
      if (id && code) custom[id[1]] = code[1].replace(/&quot;.*?&quot;|"[^"]*"/g, '');
    });
    var bloco = /<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/.exec(stylesXml);
    var xfs = bloco ? (bloco[1].match(/<xf\b[^>]*?(?:\/>|>[\s\S]*?<\/xf>)/g) || []) : [];
    return xfs.map(function (xf) {
      var id = Number((/numFmtId="(\d+)"/.exec(xf) || [0, 0])[1]);
      if ((id >= 14 && id <= 22) || (id >= 45 && id <= 47)) return true;
      var code = custom[String(id)];
      return !!code && /[dy]/i.test(code);
    });
  }

  // Estilo de cada coluna numa linha (a última da tabela é livre: tem o formato
  // padrão de cada entrada, do molde).
  function estilosDaLinha(xml, linha) {
    var out = {};
    var m = new RegExp('<row\\b[^>]*?\\br="' + linha + '"[^>]*?(?:/>|>([\\s\\S]*?)</row>)').exec(xml);
    if (!m || !m[1]) return out;
    var re = /<c\b([^>]*?)\/?>/g, c;
    while ((c = re.exec(m[1]))) {
      var r = /\br="([A-Z]+)\d+"/.exec(c[1]), s = /\bs="(\d+)"/.exec(c[1]);
      if (r && s) out[colunaDaLetra(r[1])] = s[1];
    }
    return out;
  }

  // Troca (ou cria, na ordem) a célula da linha/coluna no XML da aba, mantendo o estilo.
  function gravarCelula(xml, linha, col, type, value, ctx) {
    var ref = letraColuna(col) + linha;
    var mRow = new RegExp('<row\\b[^>]*?\\br="' + linha + '"[^>]*?(?:/>|>[\\s\\S]*?</row>)').exec(xml);
    if (!mRow) {
      var nova = '<row r="' + linha + '">' + celulaXml(ref, ctx.estilo(col, null, type), type, value) + '</row>';
      var reRow = /<row\b[^>]*?\br="(\d+)"/g, mr, pos = -1;
      while ((mr = reRow.exec(xml))) { if (Number(mr[1]) > linha) { pos = mr.index; break; } }
      if (pos === -1) pos = xml.indexOf('</sheetData>');
      return xml.slice(0, pos) + nova + xml.slice(pos);
    }
    var rowXml = mRow[0];
    if (rowXml.indexOf('</row>') === -1) rowXml = rowXml.replace(/\s*\/>$/, '>') + '</row>';
    var mCel = new RegExp('<c\\b[^>]*?\\br="' + ref + '"[^>]*?(?:/>|>[\\s\\S]*?</c>)').exec(rowXml);
    var novoRow;
    if (mCel) {
      var abre = /^<c\b[^>]*>/.exec(mCel[0])[0];
      var s = /\ss="(\d+)"/.exec(abre);
      if (/\st="s"/.test(abre)) ctx.sstLiberadas++;
      novoRow = rowXml.slice(0, mCel.index) + celulaXml(ref, ctx.estilo(col, s ? s[1] : null, type), type, value) +
        rowXml.slice(mCel.index + mCel[0].length);
    } else {
      var reC = /<c\b[^>]*?\br="([A-Z]+)\d+"/g, mc, posC = -1;
      while ((mc = reC.exec(rowXml))) { if (colunaDaLetra(mc[1]) > col) { posC = mc.index; break; } }
      if (posC === -1) posC = rowXml.lastIndexOf('</row>');
      novoRow = rowXml.slice(0, posC) + celulaXml(ref, ctx.estilo(col, null, type), type, value) + rowXml.slice(posC);
    }
    return xml.slice(0, mRow.index) + novoRow + xml.slice(mRow.index + mRow[0].length);
  }

  // Troca só o VALOR CALCULADO (<v>) de uma célula de fórmula; a fórmula fica.
  function gravarValorCalculado(xml, ref, texto) {
    var m = new RegExp('<c\\b[^>]*?\\br="' + ref + '"[^>]*?>([\\s\\S]*?)</c>').exec(xml);
    if (!m || m[1].indexOf('<f') === -1) return xml;
    var abre = /^<c\b[^>]*>/.exec(m[0])[0];
    var abreNovo = /\st="[^"]*"/.test(abre) ? abre.replace(/\st="[^"]*"/, ' t="str"') : abre.replace(/>$/, ' t="str">');
    var corpo = m[1].replace(/<v>[\s\S]*?<\/v>|<v\/>/, '') + '<v>' + xmlEsc(texto) + '</v>';
    return xml.slice(0, m.index) + abreNovo + corpo + '</c>' + xml.slice(m.index + m[0].length);
  }

  // Excel recalcula tudo ao abrir (status, dólar, KPI, conferência do HTML).
  function ligarRecalculo(wbXml) {
    if (/<calcPr\b/.test(wbXml)) {
      return wbXml.replace(/<calcPr\b([^>]*?)\s*\/>/, function (_m, a) {
        return '<calcPr' + a.replace(/\s*fullCalcOnLoad="[^"]*"/, '') + ' fullCalcOnLoad="1"/>';
      });
    }
    var depois = wbXml.indexOf('</definedNames>') > -1 ? '</definedNames>' : '</sheets>';
    return wbXml.replace(depois, depois + '<calcPr fullCalcOnLoad="1"/>');
  }

  // bytes do .xlsx + mudanças -> bytes novos. Puro (sem File System Access API):
  // é o que o conferir_html_base.js testa no Node com as planilhas de verdade.
  // `mudancas[].aba` diz a aba (Contratos ou Saving; sem aba = Contratos). `opcoes.ultimas` =
  // última linha de cada tabela (a linha livre de onde sai o estilo padrão de cada coluna).
  function gravarBase(bytes, mudancas, opcoes) {
    var zip, wbXml, rels, abaK, nomes, dataEh;
    var porAba = {};
    mudancas.forEach(function (m) { var a = m.aba || 'Contratos'; (porAba[a] = porAba[a] || []).push(m); });
    var abas = Object.keys(porAba);
    var ultimas = opcoes.ultimas || { Contratos: opcoes.ultimaLinha };
    var caminhos = [];
    return JSZip.loadAsync(bytes).then(function (z) {
      zip = z;
      return Promise.all(['xl/workbook.xml', 'xl/_rels/workbook.xml.rels', 'xl/styles.xml'].map(function (n) {
        return zip.file(n).async('string');
      }));
    }).then(function (t) {
      wbXml = t[0]; rels = t[1]; dataEh = estilosDeData(t[2]);
      caminhos = abas.map(function (a) { return caminhoDaAba(wbXml, rels, a); });
      nomes = { resultado: nomeDefinido(wbXml, 'pHtmlResultado'), status: nomeDefinido(wbXml, 'pHtmlStatus') };
      abaK = nomes.status ? caminhoDaAba(wbXml, rels, nomes.status.aba) : null;
      return Promise.all(caminhos.map(function (p) { return zip.file(p).async('string'); })
        .concat([abaK ? zip.file(abaK).async('string') : null]));
    }).then(function (t) {
      var ctx = {
        sstLiberadas: 0,
        canon: {},
        estilo: function (col, sAtual, type) {
          var canon = ctx.canon;
          if (type === 'date' && !(sAtual != null && dataEh[Number(sAtual)])) {
            return canon[col] != null && dataEh[Number(canon[col])] ? canon[col] : sAtual;
          }
          return sAtual != null ? sAtual : (canon[col] != null ? canon[col] : null);
        }
      };
      abas.forEach(function (a, i) {
        var xml = t[i];
        ctx.canon = estilosDaLinha(xml, ultimas[a]);
        porAba[a].forEach(function (m) {
          m.celulas.forEach(function (c) { xml = gravarCelula(xml, m.linha, c.col, c.type, c.value, ctx); });
        });
        zip.file(caminhos[i], xml);
      });
      var xmlK = t[abas.length];
      if (xmlK && opcoes.marca && nomes.resultado) {
        xmlK = gravarValorCalculado(xmlK, nomes.resultado.ref, opcoes.marca.resultado);
        xmlK = gravarValorCalculado(xmlK, nomes.status.ref, opcoes.marca.status);
        zip.file(abaK, xmlK);
      }
      zip.file('xl/workbook.xml', ligarRecalculo(wbXml));
      var sst = zip.file('xl/sharedStrings.xml');
      if (ctx.sstLiberadas && sst) {
        return sst.async('string').then(function (x) {
          zip.file('xl/sharedStrings.xml', x.replace(/(<sst\b[^>]*\bcount=")(\d+)"/, function (_m, a, n) {
            return a + Math.max(0, Number(n) - ctx.sstLiberadas) + '"';
          }));
        });
      }
    }).then(function () {
      return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
    });
  }

  // Descartar (trocar de ano sem salvar): o app edita os contratos no lugar (Object.assign),
  // então "descartar" relê o ano dos bytes em memória — senão a edição descartada voltaria
  // no próximo Salvar, porque a gravação compara com os valores lidos.
  function descartarAlteracoes(ano) {
    var sub = subDoAno(ano);
    if (!sub || !sub.base) return Promise.resolve();
    var antes = sub.base;
    return parseBase(sub.lastReadBuffer, sub.fileName).then(function (m) {
      ['pasta', 'caminho', 'lastModified', 'size'].forEach(function (k) { m.base[k] = antes[k]; });
      sub.contratos = m.contratos;
      sub.saving = m.saving;
      sub.cotacoes = m.cotacoes;
      sub.workbook = m.workbook;
      sub.base = m.base;
    });
  }

  function marcaGravacao(agora) {
    var hh = String(agora.getHours()).padStart(2, '0');
    var mi = String(agora.getMinutes()).padStart(2, '0');
    return {
      resultado: 'gravada pelo app HTML em ' + Regras.formatDateBR(agora) + ' ' + hh + ':' + mi + ' — falta abrir no Excel e salvar',
      status: 'ATENÇÃO'
    };
  }

  function existeArquivo(pasta, nome) {
    return pasta.getFileHandle(nome).then(function () { return true; }, function () { return false; });
  }

  // Salvar no modo base: grava só na planilha do ano em exibição — contratos e, se a planilha
  // tiver a aba Saving, o saving (`extra = {saving, savingRemovidos}`), num arquivo só.
  async function saveBase(contratos, deletedRows0, extra) {
    var sub = state.anosPlanilhas && state.anosPlanilhas[String(state.anoAtivo)];
    if (!sub || !sub.base) throw new Error('Nenhuma planilha aberta.');
    var b = sub.base;
    var nome = sub.fileName;
    // Trava 1: aberta no Excel (o Excel deixa o arquivo "~$<nome>" na pasta).
    if (await existeArquivo(b.pasta, '~$' + nome) || await existeArquivo(b.pasta, '~$' + nome.slice(2))) {
      throw new Error('A planilha ' + nome + ' está aberta no Excel. Feche no Excel e clique em Salvar de novo — nada foi gravado.');
    }
    // Trava 2: alguém salvou depois que o app abriu (gravar por cima apagaria o que ele fez).
    var atual = await sub.fileHandle.getFile();
    if (atual.lastModified !== b.lastModified || atual.size !== b.size) {
      throw new Error('A planilha ' + nome + ' mudou desde que você abriu (alguém salvou no Excel ou em outro computador). Nada foi gravado: clique em "Abrir planilha" de novo e refaça as alterações.');
    }
    var alt = alteracoesBase(sub, contratos, deletedRows0);
    var temSaving = !!(b.saving && extra && extra.saving);
    var altS = temSaving ? alteracoesTabela(b.saving, extra.saving, extra.savingRemovidos) : null;
    var mudancas = alt.mudancas.concat(altS ? altS.mudancas : []);
    if (!mudancas.length) return { gravado: false, arquivo: nome };
    var ultimas = { Contratos: b.ultimaLinha };
    if (b.saving) ultimas.Saving = b.saving.ultimaLinha;
    var novo = await gravarBase(new Uint8Array(sub.lastReadBuffer), mudancas,
      { ultimas: ultimas, marca: marcaGravacao(new Date()) });
    await writeFile(sub.fileHandle, novo);
    var gravado = await sub.fileHandle.getFile();
    b.lastModified = gravado.lastModified;
    b.size = gravado.size;
    sub.lastReadBuffer = novo.buffer.slice(novo.byteOffset, novo.byteOffset + novo.byteLength);
    state.lastReadBuffer = sub.lastReadBuffer;
    // Linhas e valores de referência para o próximo Salvar.
    aplicarGravado(b, alt, contratos, deletedRows0);
    if (altS) {
      aplicarGravado(b.saving, altS, extra.saving, extra.savingRemovidos);
      sub.saving = extra.saving;
    }
    sub.contratos = contratos;
    state.contratos = contratos;
    return { gravado: true, arquivo: nome };
  }

  // ------------------------------------------------------------
  // Modo desktop — auto-open, seletor de arquivo, leitura por caminho
  // ------------------------------------------------------------
  function openDesktopDefault() {
    if (!desktop) return Promise.reject(new Error('Modo desktop indisponível neste ambiente.'));
    return desktop.getDefaultPath().then(function (r) {
      resetarModoAnual();
      if (!r.exists) {
        return Promise.reject(new Error('Planilha não encontrada ao lado do aplicativo. Use "Abrir planilha" para escolher o arquivo.'));
      }
      // V1.6.0 — se a pasta tem MATRIZ por exercício, abre todas e ativa o ano.
      var dir = null;
      return desktop.dirname(r.path).then(function (d) {
        dir = d;
        return desktop.readdir(d);
      }).then(function (nomes) {
        var anuais = (nomes || []).map(function (n) { return { ano: anoDoArquivo(n), nome: n }; })
          .filter(function (x) { return x.ano; })
          .sort(function (a, b) { return a.ano < b.ano ? 1 : -1; });
        if (!anuais.length) return readDesktopFile(r.path);
        return abrirAnuaisDesktop(dir, anuais);
      });
    });
  }

  function openDesktopPath(filePath) {
    if (!desktop) return Promise.reject(new Error('Modo desktop indisponível neste ambiente.'));
    resetarModoAnual();
    return readDesktopFile(filePath);
  }

  function pickDesktop() {
    if (!desktop) return Promise.reject(new Error('Modo desktop indisponível neste ambiente.'));
    return desktop.pick().then(function (r) {
      resetarModoAnual();
      if (r.canceled) return Promise.reject(new Error('Nenhuma planilha selecionada.'));
      return readDesktopFile(r.path);
    });
  }

  function readDesktopFile(filePath) {
    return desktop.basename(filePath).then(function (name) {
      state.filePath = filePath;
      state.fileName = name;
      return desktop.read(filePath);
    }).then(function (res) {
      if (!res || !res.ok) return Promise.reject(new Error('Não foi possível ler o arquivo: ' + state.fileName));
      return loadFromBuffer(res.data, state.fileName);
    });
  }

  // Parse de um buffer: deixa o singleton apontando para essa planilha e
  // devolve um substate clonado (o modo anual guarda um por arquivo).
  function parseWorkbook(arrayBuffer, fileName) {
    state.lastReadBuffer = arrayBuffer;
    var wb = new ExcelJS.Workbook();
    return wb.xlsx.load(arrayBuffer).then(function () {
      state.workbook = wb;
      var ws = findManagedSheet(wb);
      state.sheetName = ws.name;
      state.migracoes = migrarLayout(ws);
      state.anosMensais = detectAnosMensais(ws);
      colunasGeridas = buildColunasGeridas(state.anosMensais);
      var contratos = parseContratos(ws);
      state.contratos = contratos;
      state.cotacoes = readCotacoes(ws, state.headerMap);
      return {
        ano: state.anoAtivo,
        fileName: fileName,
        workbook: wb,
        sheetName: ws.name,
        headerMap: state.headerMap,
        anosMensais: state.anosMensais.slice(),
        colunasGeridas: colunasGeridas.slice(),
        contratos: contratos,
        cotacoes: state.cotacoes,
        lastReadBuffer: arrayBuffer,
        lastDataRow0: state.lastDataRow0,
        migracoes: state.migracoes,
        filePath: state.filePath,
        fileHandle: state.fileHandle
      };
    });
  }

  function loadFromBuffer(arrayBuffer, fileName) {
    return parseWorkbook(arrayBuffer, fileName).then(function (modelo) {
      return { contratos: modelo.contratos, fileName: fileName, cotacoes: modelo.cotacoes, anosMensais: modelo.anosMensais.slice() };
    });
  }

  function findXlsxHandle(dirHandle) {
    var candidates = [];
    return (async function () {
      for await (var entry of dirHandle.values()) {
        if (entry.kind === 'file' && /\.xlsx$/i.test(entry.name) && !/~\$/.test(entry.name)) {
          candidates.push(entry);
        }
      }
      if (!candidates.length) throw new Error('Nenhum arquivo .xlsx encontrado na pasta selecionada.');
      var exact = candidates.find(function (h) { return normHeaderLoose(h.name).indexOf(MAIN_FILE_NAME_HINT) > -1; });
      return exact || candidates[0];
    })();
  }

  function loadFromFileHandle(fileHandle) {
    return fileHandle.getFile().then(function (file) {
      return file.arrayBuffer();
    }).then(function (buffer) {
      return loadFromBuffer(buffer, state.fileName);
    });
  }

  // ------------------------------------------------------------
  // Escrita de volta na planilha
  // ------------------------------------------------------------
  function setCell(ws, r0, c0, value, type) {
    var cell = cellAt(ws, r0, c0);
    if (value == null || value === '') {
      cell.value = null;
      return;
    }
    if (type === 'date') {
      cell.value = dateToCell(value);
      cell.numFmt = FMT_DATA;
    } else if (type === 'number') {
      cell.value = Number(value);
    } else {
      cell.value = String(value);
    }
  }

  // Grava uma FÓRMULA com o resultado já calculado, para o número aparecer
  // antes mesmo de o Excel recalcular. Resultado null representa o "" do
  // IFERROR (mês ainda sem cotação).
  // As funções vão em inglês porque é assim que o arquivo .xlsx guarda;
  // o Excel em pt-BR mostra SEERRO/SOMA na tela.
  function setFormulaCell(ws, r0, c0, formula) {
    var cell = cellAt(ws, r0, c0);
    if (!formula || !formula.expr) {
      cell.value = null;
      return;
    }
    cell.value = { formula: formula.expr, result: formula.value == null ? '' : Number(formula.value) };
  }

  // Cotação do dólar de cada mês: fica na LINHA 2 da própria coluna "MÊS $"
  // (rótulo na linha 1). É preenchida no Excel pela Hendrya — o app só lê;
  // nunca escreve nessas células. Agora há UMA faixa por ano (cada bloco de
  // meses tem sua cotação na linha 2 das próprias colunas "$").
  function readCotacoes(ws, headerMap) {
    var out = {};
    anosMensais().forEach(function (ano) {
      var porAno = {};
      Regras.MESES.forEach(function (m) {
        var c0 = headerMap[Regras.fieldUsdAno(ano, m.key)];
        if (c0 == null) return;
        var v = cellValue(cellAt(ws, COTACAO_ROW_INDEX0, c0));
        if (v == null || v === '') return;
        var n = Number(v);
        if (!isNaN(n) && n > 0) porAno[m.key] = n;
      });
      out[ano] = porAno;
    });
    return out;
  }

  // Fórmulas gravadas em cada linha de contrato, resolvidas pelo headerMap
  // (nunca por coluna fixa): "MÊS $ <ano>" = "MÊS R$ <ano>" ÷ cotação do mês
  // (linha 2 da própria coluna "$"), e o par "Total Faturado <ano> R$/$" como
  // soma dos 12 meses daquele ano — cada ano fecha o seu próprio bloco.
  function formulasDaLinha(ws, contrato, headerMap, r0, cotacoes) {
    var linha = r0 + 1; // referência A1 é 1-based
    var linhaCotacao = COTACAO_ROW_INDEX0 + 1;
    var out = {};

    anosMensais().forEach(function (ano) {
      var refsBrl = [], refsUsd = [];
      var somaBrl = 0, somaUsd = 0, temMes = false, temUsd = false;

      Regras.MESES.forEach(function (m) {
        var fieldBrl = Regras.fieldBrlAno(ano, m.key);
        var fieldUsd = Regras.fieldUsdAno(ano, m.key);
        var colBrl = headerMap[fieldBrl];
        var colUsd = headerMap[fieldUsd];
        if (colBrl == null || colUsd == null) return;
        var refBrl = colLetter(ws, colBrl) + linha;
        refsBrl.push(refBrl);
        refsUsd.push(colLetter(ws, colUsd) + linha);
        if (contrato[fieldBrl] == null) {
          out[fieldUsd] = null;
          return;
        }
        temMes = true;
        somaBrl += contrato[fieldBrl];
        var usd = Regras.mesUSD(contrato, m, ano, cotacoes);
        if (usd != null) { somaUsd += usd; temUsd = true; }
        // IFERROR mantém a célula em branco enquanto a cotação do mês não for
        // preenchida, em vez de espalhar #DIV/0! pela planilha.
        out[fieldUsd] = {
          expr: 'IFERROR(' + refBrl + '/' + colLetter(ws, colUsd) + '$' + linhaCotacao + ',"")',
          value: usd
        };
      });

      out[Regras.fieldTotalAno(ano)] = temMes
        ? { expr: 'SUM(' + refsBrl.join(',') + ')', value: somaBrl } : null;
      out[Regras.fieldTotalUsdAno(ano)] = temMes
        ? { expr: 'SUM(' + refsUsd.join(',') + ')', value: temUsd ? somaUsd : null } : null;
    });
    return out;
  }

  // contratos: array completo (existentes + novos + excluídos já removidos da lista
  // pelo chamador, mas contratos excluídos entram separadamente em `deletedRows0`)
  function applyContratosToWorkbook(contratos, deletedRows0) {
    var ws = state.workbook.getWorksheet(state.sheetName);
    var headerMap = state.headerMap;
    var nextRow0 = state.lastDataRow0 + 1;

    (deletedRows0 || []).forEach(function (r0) {
      colunasGeridas.forEach(function (def) {
        if (def.mode === 'read' || headerMap[def.field] == null) return;
        setCell(ws, r0, headerMap[def.field], null, def.type);
      });
    });

    contratos.forEach(function (c) {
      var r0;
      if (c.isNew || c._row0 == null) {
        r0 = nextRow0++;
        c._row0 = r0;
        c.isNew = false;
        c.id = r0 + 1;
      } else {
        r0 = c._row0;
      }
      var values = {
        razaoSocial: c.razaoSocial,
        tipoServico: c.tipoServico,
        departamento: c.departamento,
        prazoPagamentoRaw: c.prazoPagamentoRaw,
        prazoNegociadoRaw: c.prazoNegociadoRaw,
        valorMensal2026: c.valorMensal2026,
        valorAnual2026: c.valorAnual2026,
        vigencia: c.vigencia,
        dataInicio: c.dataInicio,
        dataFim: c.dataFim,
        statusManual: c.statusManual || '',
        comentarios: c.comentarios
      };
      anosMensais().forEach(function (ano) {
        Regras.mesFieldsAno(ano).brl.forEach(function (f) { values[f] = c[f]; });
      });
      var formulas = formulasDaLinha(ws, c, headerMap, r0, state.cotacoes);
      colunasGeridas.forEach(function (def) {
        if (def.mode === 'read' || headerMap[def.field] == null) return;
        if (def.mode === 'formula') {
          setFormulaCell(ws, r0, headerMap[def.field], formulas[def.field]);
        } else {
          setCell(ws, r0, headerMap[def.field], values[def.field], def.type);
        }
      });
    });

    state.lastDataRow0 = Math.max(state.lastDataRow0, nextRow0 - 1);
    aplicarFormatacao(ws, headerMap, state.lastDataRow0);
  }

  // ------------------------------------------------------------
  // FORMATAÇÃO DA PLANILHA
  //
  // Roda a cada Salvar e é idempotente: a planilha se conserta no primeiro
  // salvamento e não degrada nos seguintes. A fonte de verdade de largura e
  // formato numérico são as colunas geridas (campos `width` / `fmt`); as
  // colunas resolvidas pelo headerMap, nunca por letra fixa.
  // ------------------------------------------------------------
  function bordaInferior(cor, estilo) {
    return { bottom: { style: estilo || 'thin', color: { argb: cor } } };
  }

  function aplicarFormatacao(ws, headerMap, lastDataRow0) {
    var porColuna = {}; // colIndex0 -> def
    colunasGeridas.forEach(function (def) {
      if (headerMap[def.field] != null) porColuna[headerMap[def.field]] = def;
    });
    var totalCols = Math.max(ws.columnCount, 1);

    // --- colunas: largura, formato numérico e agrupamento dos meses ---
    for (var c0 = 0; c0 < totalCols; c0++) {
      var col = ws.getColumn(c0 + 1);
      var def = porColuna[c0];
      col.width = def && def.width ? def.width : 16;
      if (def && def.fmt) col.numFmt = def.fmt;
      // As colunas mensais entram num grupo recolhível. Como o par de totais de
      // cada ano fica em nível 0, o Excel fecha UM GRUPO POR ANO: a planilha
      // abre com os campos fixos + os totais, e a Hendrya expande no [+] só o
      // ano que for lançar.
      col.outlineLevel = (def && def.mes) ? 1 : 0;
    }
    ws.properties.outlineLevelCol = 1;

    // --- linhas 1-2: a faixa das cotações do dólar ---
    // A faixa pertence às colunas "MÊS $" e a mais nenhuma: é lá que a Hendrya
    // digita a cotação (linha 2) sob o rótulo (linha 1). Montamos o conjunto
    // dessas colunas pelo headerMap — assim a regra sobrevive a qualquer
    // deslocamento de coluna.
    var colunasCotacao = {};
    anosMensais().forEach(function (ano) {
      Regras.MESES.forEach(function (m) {
        var c0 = headerMap[Regras.fieldUsdAno(ano, m.key)];
        if (c0 != null) colunasCotacao[c0] = { mes: m, ano: ano };
      });
    });

    for (var c1 = 0; c1 < totalCols; c1++) {
      var cotacaoDaColuna = colunasCotacao[c1];
      var rotulo = cellAt(ws, 0, c1);
      var valor = cellAt(ws, COTACAO_ROW_INDEX0, c1);
      if (cotacaoDaColuna) {
        // rótulo faltante (ex.: SET–DEZ de 2026) nasce aqui, para a Hendrya só
        // precisar digitar o número embaixo
        if (rotulo.value == null || rotulo.value === '') {
          rotulo.value = rotuloCotacao(cotacaoDaColuna.mes, cotacaoDaColuna.ano);
        }
      } else if (valor.value != null && valor.value !== '') {
        // rótulo solto na linha das cotações (a planilha trazia "2026" e
        // "FATURAMENTO" perdidos ali) — confunde a leitura e ganhava formato
        // de cotação na formatação
        valor.value = null;
      }
      rotulo.font = { italic: true, size: 9, color: { argb: COR_MUTED } };
      rotulo.alignment = { horizontal: 'right', vertical: 'middle' };
      valor.font = { bold: true, size: 10 };
      valor.alignment = { horizontal: 'left', vertical: 'middle' };
      if (typeof valor.value === 'number') {
        valor.numFmt = FMT_BRL;
        valor.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COR_FAIXA_COTACAO } };
      }
    }
    ws.getRow(1).height = 15;
    ws.getRow(COTACAO_ROW_INDEX0 + 1).height = 18;

    // --- linha 3: cabeçalho ---
    var head = ws.getRow(HEADER_ROW_INDEX0 + 1);
    head.height = 34;
    for (var c2 = 0; c2 < totalCols; c2++) {
      var h = head.getCell(c2 + 1);
      h.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COR_PRIMARIA } };
      h.font = { bold: true, size: 10, color: { argb: COR_TEXTO_CLARO } };
      h.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
      h.border = bordaInferior(COR_PRIMARIA_ESCURA, 'medium');
    }

    // --- desocultar TODAS as linhas ---
    // A planilha chegou com 96 das 97 linhas ocultas — resíduo de um autofiltro
    // salvo. Quem abria o .xlsx via só o cabeçalho e uma linha. Vai até
    // ws.rowCount, e não até lastDataRow0, porque as linhas além do último
    // contrato (a planilha tem algumas sem Razão Social, que o parse ignora)
    // também precisam estar visíveis para quem for revisar no Excel.
    for (var rh = 1; rh <= ws.rowCount; rh++) ws.getRow(rh).hidden = false;

    // --- linhas de dados: zebra + borda + alinhamento ---
    for (var r0 = FIRST_DATA_ROW_INDEX0; r0 <= lastDataRow0; r0++) {
      var row = ws.getRow(r0 + 1);
      row.height = 18;
      var zebra = (r0 - FIRST_DATA_ROW_INDEX0) % 2 === 1;
      for (var c3 = 0; c3 < totalCols; c3++) {
        var cell = row.getCell(c3 + 1);
        cell.alignment = { vertical: 'middle', wrapText: false };
        cell.border = bordaInferior(COR_BORDA);
        cell.fill = zebra
          ? { type: 'pattern', pattern: 'solid', fgColor: { argb: COR_ZEBRA } }
          : { type: 'pattern', pattern: 'none' };
      }
    }

    // --- congelar cabeçalho + Razão Social, e autofiltro ---
    ws.views = [{ state: 'frozen', xSplit: 1, ySplit: HEADER_ROW_INDEX0 + 1 }];
    ws.autoFilter = {
      from: { row: HEADER_ROW_INDEX0 + 1, column: 1 },
      to: { row: lastDataRow0 + 1, column: totalCols }
    };
  }

  // "Inserir ano" (mecanismo app → planilha): adiciona o bloco REGULAR de 26
  // colunas de um ano novo (24 mensais + os dois totais) mais os rótulos de
  // cotação, no EXTREMO direito da aba, e reaprende o mapa de colunas na hora.
  // A gravação em disco só acontece no próximo Salvar (padrão DIRTY do app).
  // Não move nenhuma célula existente.
  function ensureYear(ano) {
    var anoStr = String(ano).trim();
    if (!/^\d{4}$/.test(anoStr)) return { ok: false, error: 'Ano inválido: informe um ano com 4 dígitos (ex.: 2027).' };
    if (!state.workbook) return { ok: false, error: 'Nenhuma planilha aberta.' };
    if (state.modoBase) {
      return { ok: false, error: 'Na base do Power BI o ano novo é uma planilha nova: no Excel, Salvar como com o ano no nome (aba "Como atualizar" da planilha).' };
    }
    if (anosMensais().indexOf(anoStr) > -1) {
      return { ok: false, error: 'O ano ' + anoStr + ' já existe na planilha.' };
    }

    var ws = state.workbook.getWorksheet(state.sheetName);
    var headerRow = ws.getRow(HEADER_ROW_INDEX0 + 1);
    var labelRow = ws.getRow(1);
    var startC0 = ws.columnCount; // primeira coluna livre (0-based)

    Regras.MESES.forEach(function (m, i) {
      var cBrl = startC0 + i * 2;
      var cUsd = cBrl + 1;
      headerRow.getCell(cBrl + 1).value = m.label + ' ' + anoStr + ' R$';
      headerRow.getCell(cUsd + 1).value = m.label + ' ' + anoStr + ' $';
      labelRow.getCell(cUsd + 1).value = rotuloCotacao(m, anoStr);
    });
    // fecha o bloco com o par de totais — sem isso o ano novo herdaria a coluna
    // de total do ano anterior no casamento de cabeçalho e a sobrescreveria
    headerRow.getCell(startC0 + 25).value = headerTotalBrl(anoStr);
    headerRow.getCell(startC0 + 26).value = headerTotalUsd(anoStr);

    state.anosMensais = anosMensais().concat([anoStr]).sort().reverse();
    colunasGeridas = buildColunasGeridas(state.anosMensais);
    state.headerMap = buildHeaderMap(ws);
    state.cotacoes = readCotacoes(ws, state.headerMap);
    return { ok: true, anos: state.anosMensais.slice() };
  }

  function workbookToArrayBuffer() {
    return state.workbook.xlsx.writeBuffer().then(function (buf) {
      return new Uint8Array(buf);
    });
  }

  async function writeFile(fileHandle, data) {
    var writable = await fileHandle.createWritable();
    await writable.write(data);
    await writable.close();
  }

  // Salva o novo conteúdo no arquivo principal.
  async function save(contratos, deletedRows0, extra) {
    if (state.modoBase) return saveBase(contratos, deletedRows0, extra);
    if (nativeMode) return saveDesktop(contratos, deletedRows0);
    if (!state.dirHandle || !state.fileHandle || !state.workbook) {
      throw new Error('Nenhuma planilha aberta.');
    }
    applyContratosToWorkbook(contratos, deletedRows0);
    var newBuffer = await workbookToArrayBuffer();

    await writeFile(state.fileHandle, newBuffer);
    state.lastReadBuffer = newBuffer;
    atualizarSubstateSalvo();
    return true;
  }

  // Modo desktop: principal e snapshot ficam no mesmo diretório do arquivo.
  async function saveDesktop(contratos, deletedRows0) {
    if (!state.filePath || !state.workbook) {
      throw new Error('Nenhuma planilha aberta.');
    }
    applyContratosToWorkbook(contratos, deletedRows0);
    var newBuffer = await workbookToArrayBuffer();

    await desktop.write(state.filePath, newBuffer);
    state.lastReadBuffer = newBuffer;
    atualizarSubstateSalvo();
    return true;
  }

  // Modo anual: o workbook é o mesmo objeto do substate do ano ativo, mas o
  // lastReadBuffer (base da próxima gravação) vive no substate — sincroniza.
  function atualizarSubstateSalvo() {
    if (!state.estadoAnual || !state.anosPlanilhas || state.anoAtivo == null) return;
    var sub = state.anosPlanilhas[String(state.anoAtivo)];
    if (sub) sub.lastReadBuffer = state.lastReadBuffer;
  }

  // fileName opcional — o "Exportar HTML do dia" passa um nome com data.
  async function saveSnapshotHtml(htmlString, fileName) {
    var name = fileName || SNAPSHOT_FILE_NAME;
    if (nativeMode) {
      if (!state.filePath) throw new Error('Nenhuma planilha aberta.');
      var dir = await desktop.dirname(state.filePath);
      var data = new TextEncoder().encode(htmlString);
      await desktop.write(await desktop.join(dir, name), data);
      return name;
    }
    if (!state.dirHandle) throw new Error('Nenhuma pasta aberta.');
    var handle = await state.dirHandle.getFileHandle(name, { create: true });
    var writable = await handle.createWritable();
    await writable.write(htmlString);
    await writable.close();
    return name;
  }

  global.ExcelIO = {
    isSupported: isSupported,
    openFolder: openFolder,
    abrirPastaPrincipal: abrirPastaPrincipal,
    openDesktopDefault: openDesktopDefault,
    openPath: openDesktopPath,
    save: save,
    saveSnapshotHtml: saveSnapshotHtml,
    ensureYear: ensureYear,
    ativarAno: ativarAno,
    getState: function () { return state; },
    // V1.7.0 — modo base: cada exercício vem do seu arquivo
    modoBase: function () { return !!state.modoBase; },
    descartarAlteracoes: descartarAlteracoes,
    contratosDoAno: function (ano) { var s = subDoAno(ano); return s ? s.contratos : null; },
    referenciaDoAno: function (ano) { var s = subDoAno(ano); return s && s.base ? s.base.referencia : null; },
    posicaoDoAno: function (ano) { var s = subDoAno(ano); return s && s.base ? s.base.posicao : ''; },
    // Só para o teste no Node (conferir_html_base.js, na raiz do 04): leitura e
    // gravação puras, sem File System Access API.
    _base: {
      parse: parseBase, alteracoes: alteracoesBase, alteracoesTabela: alteracoesTabela, gravar: gravarBase,
      marca: marcaGravacao, aplicarGravado: aplicarGravado
    }
  };

  function subDoAno(ano) {
    return (state.modoBase && state.anosPlanilhas) ? state.anosPlanilhas[String(ano)] || null : null;
  }
})(window);
