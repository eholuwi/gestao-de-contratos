/* ============================================================
   excel.js — Leitura/escrita da MATRIZ DE CONTRATOS 2026.xlsx
   direto no navegador via File System Access API + SheetJS.

   Estratégia (ver docs/ARQUITETURA.md):
   - "Abrir planilha" pede acesso a uma PASTA (não só o arquivo), porque
     assim conseguimos, sem diálogo extra, resolver também o backup
     (MATRIZ DE CONTRATOS 2026.backup.xlsx) e o snapshot de distribuição
     na mesma pasta, com uma única concessão de permissão.
   - Escrita atualiza só as células das colunas gerenciadas (ver
     COLUNAS_GERENCIADAS abaixo); todas as outras colunas/abas da planilha
     (TOTAL ANUAL, JAN..DEZ, "Contratos P&D", "Tabela de prazo...") são
     preservadas intactas porque nunca são tocadas.
   - "Excluir" um contrato não remove a linha física (isso desalinharia
     colunas não gerenciadas de outras linhas em algumas versões de
     planilha) — em vez disso limpa os campos gerenciados da linha,
     inclusive Razão Social, que é o sinal usado no parse para pular a
     linha (mesma regra do Power Query: RemoveBlankRows).
   ============================================================ */
(function (global) {
  'use strict';

  var SHEET_NAME_HINT = 'gestao de contratos bi';
  var HEADER_ROW_INDEX0 = 2; // linha 3 (0-based = 2)
  var FIRST_DATA_ROW_INDEX0 = 3; // linha 4 (0-based = 3)
  var MAIN_FILE_NAME_HINT = 'matriz de contratos';
  var BACKUP_SUFFIX = '.backup.xlsx';
  var SNAPSHOT_FILE_NAME = 'Dashboard_Contratos_snapshot.html';

  // header da planilha (normalizado) -> campo interno
  var COLUNAS_GERENCIADAS = [
    { header: 'razao social', field: 'razaoSocial', type: 'text' },
    { header: 'tipo de servico', field: 'tipoServico', type: 'text' },
    { header: 'departamento', field: 'departamento', type: 'text' },
    { header: 'prazo pagamento', field: 'prazoPagamentoRaw', type: 'text' },
    { header: 'prazo negociado', field: 'prazoNegociadoRaw', type: 'text' },
    { header: 'valor mensal 2', field: 'valorMensal2026', type: 'number' },
    { header: 'valor anual2', field: 'valorAnual2026', type: 'number' },
    { header: 'vigencia', field: 'vigencia', type: 'text' },
    { header: 'inicio', field: 'dataInicio', type: 'date' },
    { header: 'fim', field: 'dataFim', type: 'date' },
    { header: 'status.', field: 'statusManual', type: 'text' },
    { header: 'comentarios', field: 'comentarios', type: 'text' }
  ];

  // normaliza removendo acentos/pontuação para casar cabeçalhos como "Valor Anual2"
  // mesmo com pequenas variações de grafia/espaçamento na planilha de origem.
  function normHeaderLoose(s) {
    return Regras.stripAccents(String(s == null ? '' : s))
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, ' ')
      .trim();
  }

  var supported = !!(global.showDirectoryPicker);

  function isSupported() { return supported; }

  // ------------------------------------------------------------
  // Estado do módulo
  // ------------------------------------------------------------
  var state = {
    dirHandle: null,
    fileHandle: null,
    fileName: null,
    workbook: null,       // objeto SheetJS (mantido em memória para preservar outras abas)
    sheetName: null,
    lastReadBuffer: null, // bytes originais da última leitura bem-sucedida (usado pro backup)
    headerMap: null,      // {field: colIndex0}
    lastDataRow0: null    // índice 0-based da última linha de dados conhecida
  };

  function findManagedSheetName(workbook) {
    var hit = workbook.SheetNames.find(function (n) {
      return normHeaderLoose(n).indexOf('gestao de contratos') > -1;
    });
    return hit || workbook.SheetNames[0];
  }

  function buildHeaderMap(ws) {
    var range = XLSX.utils.decode_range(ws['!ref']);
    var map = {};
    for (var c = range.s.c; c <= range.e.c; c++) {
      var cellAddr = XLSX.utils.encode_cell({ r: HEADER_ROW_INDEX0, c: c });
      var cell = ws[cellAddr];
      if (!cell || cell.v == null) continue;
      var norm = normHeaderLoose(cell.v);
      COLUNAS_GERENCIADAS.forEach(function (def) {
        if (normHeaderLoose(def.header) === norm && map[def.field] == null) {
          map[def.field] = c;
        }
      });
    }
    return map;
  }

  function cellToDate(v) {
    if (v == null || v === '') return null;
    if (v instanceof Date) return isNaN(v.getTime()) ? null : v;
    if (typeof v === 'number') {
      var d = XLSX.SSF.parse_date_code(v);
      if (!d) return null;
      return new Date(d.y, d.m - 1, d.d);
    }
    return Regras.parseDateInput(v);
  }

  function parseContratos(workbook, sheetName) {
    var ws = workbook.Sheets[sheetName];
    var headerMap = buildHeaderMap(ws);
    var range = XLSX.utils.decode_range(ws['!ref']);
    var contratos = [];
    var lastDataRow0 = FIRST_DATA_ROW_INDEX0 - 1;

    for (var r = FIRST_DATA_ROW_INDEX0; r <= range.e.r; r++) {
      var get = function (field) {
        if (headerMap[field] == null) return undefined;
        var addr = XLSX.utils.encode_cell({ r: r, c: headerMap[field] });
        var cell = ws[addr];
        return cell ? cell.v : undefined;
      };

      var razaoRaw = get('razaoSocial');
      var linhaVazia = COLUNAS_GERENCIADAS.every(function (def) {
        var v = get(def.field);
        return v == null || v === '';
      });
      if (linhaVazia) continue;
      if (razaoRaw == null || String(razaoRaw).trim() === '') continue;

      var prazoRaw = get('prazoPagamentoRaw');
      var c = {
        _row0: r,
        id: r + 1, // número de linha do Excel (1-based) — chave estável
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
        // à data real de fim. Usá-la para PILOTAR o
        // para resolver. O texto original fica em statusLegadoRaw só para referência/
        // auditoria (ver Regras.qualityFlags) — nunca para decidir o status exibido.
        statusManual: null,
        statusLegadoRaw: (function () { var v = get('statusManual'); return v == null ? '' : String(v).trim(); })()
      };
      contratos.push(c);
      lastDataRow0 = r;
    }

    state.headerMap = headerMap;
    state.lastDataRow0 = lastDataRow0;
    return contratos;
  }

  // ------------------------------------------------------------
  // Abrir pasta + planilha
  // ------------------------------------------------------------
  function openFolder() {
    if (!supported) {
      return Promise.reject(new Error('Este navegador não suporta a File System Access API (necessária para editar). Use Microsoft Edge ou Google Chrome.'));
    }
    return global.showDirectoryPicker({ mode: 'readwrite' }).then(function (dirHandle) {
      state.dirHandle = dirHandle;
      return findXlsxHandle(dirHandle);
    }).then(function (fileHandle) {
      state.fileHandle = fileHandle;
      state.fileName = fileHandle.name;
      return loadFromFileHandle(fileHandle);
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
      state.lastReadBuffer = buffer;
      var wb = XLSX.read(buffer, { type: 'array', cellDates: true });
      state.workbook = wb;
      state.sheetName = findManagedSheetName(wb);
      var contratos = parseContratos(wb, state.sheetName);
      return { contratos: contratos, fileName: state.fileName };
    });
  }

  // ------------------------------------------------------------
  // Escrita de volta na planilha
  // ------------------------------------------------------------
  function setCell(ws, r, c, value, type) {
    var addr = XLSX.utils.encode_cell({ r: r, c: c });
    if (value == null || value === '') {
      // limpa a célula (usado em exclusão / campo apagado)
      delete ws[addr];
      return;
    }
    var cell = { v: value };
    if (type === 'date' && value instanceof Date) {
      cell = { v: value, t: 'd', z: 'dd/mm/yyyy' };
    } else if (type === 'number') {
      cell = { v: Number(value), t: 'n' };
    } else {
      cell = { v: String(value), t: 's' };
    }
    ws[addr] = cell;
  }

  function extendRangeIfNeeded(ws, maxRow0) {
    var range = XLSX.utils.decode_range(ws['!ref']);
    if (maxRow0 > range.e.r) {
      range.e.r = maxRow0;
      ws['!ref'] = XLSX.utils.encode_range(range);
    }
  }

  // contratos: array completo (existentes + novos + excluídos já removidos da lista
  // pelo chamador, mas contratos excluídos entram separadamente em `deletedRows0`)
  function applyContratosToWorkbook(contratos, deletedRows0) {
    var ws = state.workbook.Sheets[state.sheetName];
    var headerMap = state.headerMap;
    var nextRow0 = state.lastDataRow0 + 1;

    (deletedRows0 || []).forEach(function (r0) {
      COLUNAS_GERENCIADAS.forEach(function (def) {
        if (headerMap[def.field] == null) return;
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
      COLUNAS_GERENCIADAS.forEach(function (def) {
        if (headerMap[def.field] == null) return;
        setCell(ws, r0, headerMap[def.field], values[def.field], def.type);
      });
    });

    state.lastDataRow0 = Math.max(state.lastDataRow0, nextRow0 - 1);
    extendRangeIfNeeded(ws, state.lastDataRow0);
  }

  function workbookToArrayBuffer() {
    return XLSX.write(state.workbook, { bookType: 'xlsx', type: 'array', cellDates: true });
  }

  async function writeFile(fileHandle, arrayBuffer) {
    var writable = await fileHandle.createWritable();
    await writable.write(arrayBuffer);
    await writable.close();
  }

  // Salva: 1) backup com o conteúdo ANTERIOR à gravação  2) grava o novo conteúdo no arquivo principal
  async function save(contratos, deletedRows0) {
    if (!state.dirHandle || !state.fileHandle || !state.workbook) {
      throw new Error('Nenhuma planilha aberta.');
    }
    applyContratosToWorkbook(contratos, deletedRows0);
    var newBuffer = workbookToArrayBuffer();

    var backupName = backupFileName(state.fileName);
    var backupHandle = await state.dirHandle.getFileHandle(backupName, { create: true });
    if (state.lastReadBuffer) {
      await writeFile(backupHandle, state.lastReadBuffer);
    }

    await writeFile(state.fileHandle, newBuffer);
    state.lastReadBuffer = newBuffer;
    return true;
  }

  function backupFileName(fileName) {
    var base = fileName.replace(/\.xlsx$/i, '');
    return base + BACKUP_SUFFIX;
  }

  async function saveSnapshotHtml(htmlString) {
    if (!state.dirHandle) throw new Error('Nenhuma pasta aberta.');
    var handle = await state.dirHandle.getFileHandle(SNAPSHOT_FILE_NAME, { create: true });
    var writable = await handle.createWritable();
    await writable.write(htmlString);
    await writable.close();
    return SNAPSHOT_FILE_NAME;
  }

  global.ExcelIO = {
    isSupported: isSupported,
    openFolder: openFolder,
    save: save,
    saveSnapshotHtml: saveSnapshotHtml,
    getState: function () { return state; }
  };
})(window);
