define(['qlik', 'jquery', 'text!./qollect.css', './exceljs.min'], function (qlik, $, cssContent, ExcelJS) {
  'use strict';

  // ------- shared export flags (used by XLSX) -------
  const AUTO_FILTERS = true;
  const FREEZE_HEADER = true;
  const AUTO_WIDTHS = true;

  // ExcelJS AMD/module fallback
  ExcelJS = ExcelJS || (window && window.ExcelJS);

  // ------- load external CSS (no inline styles) -------
  (function ensureStyle() {
    if (!document.getElementById('qollect-style')) {
      const s = document.createElement('style');
      s.id = 'qollect-style';
      s.textContent = cssContent;
      document.head.appendChild(s);
    }
  })();

  // ------- plain text downloader (for .qvs) -------
  function downloadText(filename, text) {
    const name = filename && filename.trim() ? filename : 'script';
    try {
      const blob = new Blob([text], { type: 'text/plain;charset=utf-8' });
      const url = (window.URL || window.webkitURL).createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = name.endsWith('.qvs') ? name : `${name}.qvs`;
      document.body.appendChild(a);
      a.click();
      setTimeout(() => {
        document.body.removeChild(a);
        (window.URL || window.webkitURL).revokeObjectURL(url);
      }, 0);
    } catch (e) {
      const data = 'data:text/plain;charset=utf-8,' + encodeURIComponent(text);
      const a = document.createElement('a');
      a.href = data;
      a.download = name.endsWith('.qvs') ? name : `${name}.qvs`;
      document.body.appendChild(a);
      a.click();
      setTimeout(() => a.remove(), 0);
    }
  }

  // ===================== ExcelJS XLSX export helpers =====================
  function pxToExcelCharWidth(px) {
    const chars = Math.round((Number(px) || 80) / 7);
    return Math.max(8, Math.min(80, chars));
  }

  function normalizeSheetName(name) {
    const s = String(name || 'Sheet').replace(/[:\\/?*\[\]]/g, ' ').trim();
    return (s || 'Sheet').slice(0, 31);
  }

  function addHeaderStyle(ws, headerRowNumber, colsCount) {
    const headerRow = ws.getRow(headerRowNumber);
    headerRow.font = { bold: true };
    for (let c = 1; c <= colsCount; c++) {
      const cell = headerRow.getCell(c);
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9E1F2' } };
      cell.border = { bottom: { style: 'thin' } };
      cell.alignment = { vertical: 'middle', wrapText: true };
    }
  }

  function applyUnusedRowStyle(row) {
    row.eachCell({ includeEmpty: true }, (cell) => {
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFCE4E4' } };
    });
  }

  function applyWrapTopToRow(row) {
    row.eachCell({ includeEmpty: true }, (cell) => {
      cell.alignment = Object.assign({}, cell.alignment || {}, { wrapText: true, vertical: 'top' });
    });
  }

  function applyWrapTopToCell(cellObj) {
    cellObj.alignment = Object.assign({}, cellObj.alignment || {}, { wrapText: true, vertical: 'top' });
  }

  function setAutoFilter(ws, rowsCount, colsCount) {
    try {
      ws.autoFilter = {
        from: { row: 1, column: 1 },
        to: { row: Math.max(1, rowsCount), column: Math.max(1, colsCount) }
      };
    } catch (e) {}
  }

  function setFreezeHeader(ws) {
    ws.views = [{ state: 'frozen', ySplit: 1 }];
  }

  async function downloadXlsx(filename, workbook) {
    const name = filename && filename.trim() ? filename.trim() : 'Qollect_Metadata';
    const finalName = name.toLowerCase().endsWith('.xlsx') ? name : `${name}.xlsx`;

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = (window.URL || window.webkitURL).createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = finalName;
    document.body.appendChild(a);
    a.click();
    setTimeout(() => {
      document.body.removeChild(a);
      (window.URL || window.webkitURL).revokeObjectURL(url);
    }, 0);
  }

  function estimateColumnWidths(headers, matrix) {
    const PX_PER_CHAR = 7, PADDING = 16, MIN_W = 80, MAX_W = 600;
    const cols = headers.length;
    const maxLens = Array.from({ length: cols }, (_, i) => String(headers[i] ?? '').length);

    for (const r of (matrix || [])) {
      for (let c = 0; c < cols; c++) {
        const len = String((r || [])[c] ?? '').replace(/\r?\n/g, '').length;
        if (len > maxLens[c]) maxLens[c] = len;
      }
    }
    return maxLens.map(len => Math.max(MIN_W, Math.min(MAX_W, len * PX_PER_CHAR + PADDING)));
  }

  function addWorksheetXlsx(workbook, spec) {
    const ws = workbook.addWorksheet(normalizeSheetName(spec.name));

    const headers = spec.headers || [];
    const matrix = spec.matrix || [];
    const colsCount = headers.length;
    const rowsCount = matrix.length + 1;

    const widthsPx = Array.isArray(spec.widths) && spec.widths.length
      ? spec.widths.slice(0, colsCount)
      : (AUTO_WIDTHS ? estimateColumnWidths(headers, matrix) : []);

    ws.columns = headers.map((h, idx) => ({
      header: String(h ?? ''),
      key: `c${idx + 1}`,
      width: widthsPx.length ? pxToExcelCharWidth(widthsPx[idx] || 80) : 16
    }));

    addHeaderStyle(ws, 1, colsCount);

    for (let r = 0; r < matrix.length; r++) {
      const rawRow = matrix[r] || [];
      const excelRow = ws.addRow(rawRow.map(v => (v == null ? '' : v)));

      if (spec.wrapAll) applyWrapTopToRow(excelRow);

      if (Array.isArray(spec.wrapCols) && spec.wrapCols.length) {
        for (const colIndex1Based of spec.wrapCols) {
          const c = excelRow.getCell(colIndex1Based);
          applyWrapTopToCell(c);
        }
      }

      excelRow.eachCell({ includeEmpty: true }, (c) => {
        const v = c.value;
        const s = (v == null) ? '' : String(v);
        if (/[\r\n]/.test(s)) applyWrapTopToCell(c);
      });

      if (typeof spec.rowStyleId === 'function') {
        const styleId = spec.rowStyleId(rawRow);
        if (styleId === 'sUnused') applyUnusedRowStyle(excelRow);
        if (styleId === 'sWrapTop') applyWrapTopToRow(excelRow);
      }
    }

    if (AUTO_FILTERS) setAutoFilter(ws, rowsCount, colsCount);
    if (FREEZE_HEADER) setFreezeHeader(ws);

    return ws;
  }

  // ------- Engine helpers -------
  function getDoc(app) {
    const doc = app && app.model && app.model.enigmaModel;
    if (!doc) throw new Error('Engine document not available (app.model.enigmaModel missing).');
    return doc;
  }

  const YN = v => (v === true ? 'Y' : v === false ? 'N' : '');
  const sortAsc = (arr, key) => arr.slice().sort((a, b) => String(a?.[key] ?? '').localeCompare(String(b?.[key] ?? ''), undefined, { sensitivity: 'base' }));

  // ===================== field parsing helpers =====================
  const reInlineDollar = /\$\(\s*=\s*([^)]*?)\s*\)/g;
  const reVarDollar = /\$\(\s*([A-Za-z_]\w*)(?:\([^\)]*\))?\s*\)/g;
  const reBlockComments = new RegExp('/' + '\\*' + '[\\s\\S]*?' + '\\*' + '/', 'g');
  const reHasSet = /\{\s*<[\s\S]*?>\s*\}/;

  function buildVariableMap(vars) {
    const m = new Map();
    for (const v of vars || []) if (v?.name) m.set(v.name, v.definition ?? '');
    return m;
  }

  function expandDollar(expr, varMap, depth = 0, seen = new Set()) {
    if (!expr || typeof expr !== 'string') return '';
    if (depth > 5) return expr;
    expr = expr.replace(reInlineDollar, (_, inner) => inner || '');
    expr = expr.replace(reVarDollar, (_, vname) => {
      if (!varMap || !varMap.has(vname)) return '';
      if (seen.has(vname)) return '';
      seen.add(vname);
      return expandDollar(varMap.get(vname) || '', varMap, depth + 1, seen);
    });
    return expr;
  }

  function extractFieldsFromExpr(expr, allFieldsSet) {
    const used = new Set();
    if (!expr || typeof expr !== 'string') return used;

    let s = expr.replace(/\/\/.*$/gm, '').replace(reBlockComments, '');

    const br = s.match(/\[([^\]\\]|\\.)+\]/g) || [];
    for (const tok of br) {
      const name = tok.slice(1, -1);
      const base = name.replace(/\.autoCalendar\..*$/, '');
      if (allFieldsSet.has(name)) used.add(name);
      if (allFieldsSet.has(base)) used.add(base);
    }

    const sa = s.match(/\{\s*<([\s\S]*?)>\s*\}/g) || [];
    for (const block of sa) {
      const inside = block.slice(block.indexOf('<') + 1, block.lastIndexOf('>'));
      const parts = inside.split(/,(?=(?:[^'"]|'[^']*'|"[^"]*")*$)/g);
      for (const part of parts) {
        const m = part.match(/(^|[,<\s])(?:(?:\w+::)?)(\[?[^\]=,]+?\]?)(?==)/);
        if (m) {
          let lhs = m[2] || '';
          lhs = lhs.replace(/^\[|\]$/g, '');
          const base = lhs.replace(/\.autoCalendar\..*$/, '');
          if (allFieldsSet.has(lhs)) used.add(lhs);
          if (allFieldsSet.has(base)) used.add(base);
        }
      }
    }

    const ac = s.match(/([A-Za-z_][\w ]*)\.autoCalendar\.[A-Za-z]+/g) || [];
    for (const t of ac) {
      const base = t.replace(/\.autoCalendar\..*$/, '').trim();
      if (allFieldsSet.has(base)) used.add(base);
    }

    for (const fname of allFieldsSet) {
      if (!/^[A-Za-z_][\w.]*$/.test(fname)) continue;
      const escaped = fname.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
      const re = new RegExp(`(?<![\\w$])${escaped}(?![\\w$])`);
      if (re.test(s)) used.add(fname);
    }

    return used;
  }

  function collectObjectExpressions(props) {
    const out = [];
    const seen = new WeakSet();
    const MAX_DEPTH = 50;

    function pushExpr(val, path, force = false) {
      if (typeof val === 'string' && val.trim()) {
        if (force || /[=\[\]{}$]/.test(val) || /autoCalendar/.test(val) || /[A-Za-z_]\w*\s*\(/.test(val)) {
          out.push({ expr: val, path });
        }
      }
    }

    function pushField(val, path) {
      if (typeof val === 'string' && val.trim()) {
        out.push({ expr: /^\[.*\]$/.test(val) ? val : `[${val}]`, path });
      }
    }

    function walk(node, path, depth = 0) {
      if (!node || depth > MAX_DEPTH) return;
      if (typeof node === 'string') { pushExpr(node, path); return; }

      if (Array.isArray(node)) {
        for (let i = 0; i < node.length; i++) {
          const v = node[i];
          if (typeof v === 'string') pushExpr(v, `${path}[${i}]`);
          else if (v && typeof v === 'object' && !seen.has(v)) { seen.add(v); walk(v, `${path}[${i}]`, depth + 1); }
        }
        return;
      }

      if (typeof node === 'object') {
        if (seen.has(node)) return;
        seen.add(node);

        if (node.qListObjectDef) {
          const lo = node.qListObjectDef;
          const fdefs = Array.isArray(lo?.qDef?.qFieldDefs) && lo.qDef.qFieldDefs.length ? lo.qDef.qFieldDefs
            : (lo?.qDef?.qFieldDef ? [lo.qDef.qFieldDef] : []);
          fdefs.forEach((f, idx) => pushField(f, `${path}/qListObjectDef/qDef/qFieldDefs[${idx}]`));
          pushExpr(lo?.qDef?.qLabelExpression, `${path}/qListObjectDef/qDef/qLabelExpression`);
          pushExpr(lo?.qCalcCond, `${path}/qListObjectDef/qCalcCond`);
        }

        if (node.qHyperCubeDef) {
          const hc = node.qHyperCubeDef;
          (hc.qDimensions || []).forEach((d, i) => {
            const fdefs = Array.isArray(d?.qDef?.qFieldDefs) && d.qDef.qFieldDefs.length ? d.qDef.qFieldDefs
              : (d?.qDef?.qFieldDef ? [d.qDef.qFieldDef] : []);
            fdefs.forEach((f, idx) => pushField(f, `${path}/qHyperCubeDef/qDimensions[${i}]/qDef/qFieldDefs[${idx}]`));
            pushExpr(d?.qDef?.qLabelExpression, `${path}/qHyperCubeDef/qDimensions[${i}]/qDef/qLabelExpression`);
            pushExpr(d?.qCalcCond, `${path}/qHyperCubeDef/qDimensions[${i}]/qCalcCond`);
          });
          (hc.qMeasures || []).forEach((m, i) => {
            pushExpr(m?.qDef?.qDef, `${path}/qHyperCubeDef/qMeasures[${i}]/qDef/qDef`, true);
            pushExpr(m?.qDef?.qLabelExpression, `${path}/qHyperCubeDef/qMeasures[${i}]/qDef/qLabelExpression`);
            pushExpr(m?.qSortByExpression?.qExpression, `${path}/qHyperCubeDef/qMeasures[${i}]/qSortByExpression/qExpression`, true);
          });
          pushExpr(hc.qCalcCond, `${path}/qHyperCubeDef/qCalcCond`);
          pushExpr(hc.qSuppressZero, `${path}/qHyperCubeDef/qSuppressZero`);
          pushExpr(hc.qSuppressMissing, `${path}/qHyperCubeDef/qSuppressMissing`);
        }

        for (const k of Object.keys(node)) {
          const v = node[k];
          if (typeof v === 'string') pushExpr(v, `${path}/${k}`);
          else if (v && typeof v === 'object' && !seen.has(v)) { seen.add(v); walk(v, `${path}/${k}`, depth + 1); }
        }
      }
    }

    walk(props, 'object');
    return out;
  }

  // ---- robust object discovery (getChildInfos + recursion) ----
  async function fetchSheetsViaEngine(app) {
    const doc = getDoc(app);
    const listHandle = await doc.createSessionObject({
      qInfo: { qType: 'SheetList' },
      qAppObjectListDef: { qType: 'sheet', qData: { rank: '/rank' } }
    });
    const layout = await listHandle.getLayout();
    const items = layout?.qAppObjectList?.qItems || [];
    return items.map(it => ({
      id: it.qInfo?.qId || '',
      title: it.qMeta?.title || '',
      description: it.qMeta?.description || '',
      owner: (it.qMeta && it.qMeta.owner && (it.qMeta.owner.name || it.qMeta.owner.userId)) || ''
    }));
  }

  async function fetchObjectPropsForSheets(app) {
    const objs = [];
    const sheets = await fetchSheetsViaEngine(app);
    const visited = new Set();

    async function addObjectById(objId, sheet) {
      if (!objId || visited.has(objId)) return;
      visited.add(objId);
      try {
        const objModel = await app.getObject(objId);

        const p = await objModel.getProperties().catch(() => null);
        const l = await objModel.getLayout().catch(() => null);

        const propsForScan = (() => {
          const exL = l && l.qLayoutExclude ? l.qLayoutExclude : null;
          const exP = p && p.qLayoutExclude ? p.qLayoutExclude : null;
          if (p && (exL || exP)) return { ...p, qLayoutExclude: (exL || exP) };
          if (p) return p;
          if (exL || exP) return { qLayoutExclude: (exL || exP) };
          return {};
        })();

        objs.push({ sheetTitle: sheet.title, sheetId: sheet.id, objectId: objId, props: propsForScan });

        let kids = [];
        try { kids = await objModel.getChildInfos(); } catch (e) {}

        if (Array.isArray(kids) && kids.length) {
          for (const ch of kids) await addObjectById(ch.qId || ch.id || ch.Id, sheet);
        } else {
          const cells = Array.isArray(p?.cells) ? p.cells : [];
          for (const c of cells) await addObjectById(c.name || c.qId || c.id, sheet);
        }
      } catch (e) {}
    }

    for (const sh of sheets) {
      try {
        const sheetModel = await app.getObject(sh.id);
        let childInfos = [];
        try { childInfos = await sheetModel.getChildInfos(); } catch (e) {}

        if (Array.isArray(childInfos) && childInfos.length) {
          for (const ch of childInfos) await addObjectById(ch.qId || ch.id || ch.Id, sh);
        } else {
          const props = await sheetModel.getProperties().catch(() => null);
          const cells = Array.isArray(props?.cells) ? props.cells : [];
          for (const c of cells) await addObjectById(c.name || c.qId || c.id, sh);
        }
      } catch (e) {}
    }

    return objs;
  }

  // ===================== MASTER ITEM USAGE (primary + alternates) =====================
  async function computeMasterUsage(app) {
    const objs = await fetchObjectPropsForSheets(app);

    const dimSlots = new Map();
    const msrSlots = new Map();

    const addSlot = (map, id, slotKey) => {
      if (!id || !slotKey) return;
      let s = map.get(id);
      if (!s) { s = new Set(); map.set(id, s); }
      s.add(slotKey);
    };

    function scanPropsForLibraryIds(props, objId) {
      function walk(node, path, lastDimIdx = null, lastMsrIdx = null) {
        if (!node) return;

        if (Array.isArray(node)) {
          node.forEach((v, i) => {
            if (/qDimensions$/.test(path)) walk(v, `${path}[${i}]`, i, lastMsrIdx);
            else if (/qMeasures$/.test(path)) walk(v, `${path}[${i}]`, lastDimIdx, i);
            else walk(v, `${path}[${i}]`, lastDimIdx, lastMsrIdx);
          });
          return;
        }

        if (typeof node === 'object') {
          for (const [k, v] of Object.entries(node)) {
            const p = `${path}/${k}`;

            if (k === 'qLibraryId' && typeof v === 'string' && v.trim()) {
              const isDim = /\/qDimensions(\[|\/)/.test(path) || /qListObjectDef/.test(path);
              const isMsr = /\/qMeasures(\[|\/)/.test(path);

              const mD = path.match(/qDimensions\[\d+\]/g);
              const mM = path.match(/qMeasures\[\d+\]/g);
              const slot = mD?.[mD.length - 1] || mM?.[mM.length - 1] || 'misc';
              const slotKey = `${objId}:${slot}`;

              if (isMsr) addSlot(msrSlots, v, slotKey);
              else if (isDim) addSlot(dimSlots, v, slotKey);
              else {
                if (lastMsrIdx !== null) addSlot(msrSlots, v, `${objId}:qMeasures[${lastMsrIdx}]`);
                else addSlot(dimSlots, v, `${objId}:qDimensions[${lastDimIdx ?? '0'}]`);
              }
            }

            if (v && (typeof v === 'object' || Array.isArray(v))) walk(v, p, lastDimIdx, lastMsrIdx);
          }
        }
      }
      walk(props, 'object', null, null);
    }

    for (const o of objs) {
      scanPropsForLibraryIds(o.props, o.objectId);

      const masterId = o?.props?.qExtendsId;
      if (masterId) {
        try {
          const masterModel = await app.getObject(masterId);
          const mp = await masterModel.getProperties().catch(() => null);
          const ml = await masterModel.getLayout().catch(() => null);
          const propsForScan = mp
            ? (ml && ml.qLayoutExclude ? { ...mp, qLayoutExclude: ml.qLayoutExclude } : mp)
            : (ml && ml.qLayoutExclude ? { qLayoutExclude: ml.qLayoutExclude } : {});
          scanPropsForLibraryIds(propsForScan, o.objectId);
        } catch (e) {}
      }
    }

    const dimCount = new Map();
    const msrCount = new Map();
    for (const [id, set] of dimSlots) dimCount.set(id, set.size);
    for (const [id, set] of msrSlots) msrCount.set(id, set.size);
    return { dimCount, msrCount };
  }

  // ===================== FETCHERS =====================
  async function fetchDimensionsViaEngine(app) {
    const doc = getDoc(app);
    const listHandle = await doc.createSessionObject({
      qInfo: { qType: 'DimensionList' },
      qDimensionListDef: { qType: 'dimension', qData: { title: '/qMetaDef/title', description: '/qMetaDef/description', tags: '/qMetaDef/tags' } }
    });
    const layout = await listHandle.getLayout();
    const items = layout?.qDimensionList?.qItems || [];
    if (!items.length) return [];

    const results = [];
    for (const it of items) {
      const id = it.qInfo?.qId;
      if (!id) continue;

      let props = null;
      try {
        if (typeof app.getDimension === 'function') {
          const dimModel = await app.getDimension(id);
          props = await dimModel.getProperties();
        } else if (typeof getDoc(app).getDimension === 'function') {
          const dimHandle = await getDoc(app).getDimension({ qId: id });
          props = await dimHandle.getProperties();
        }
      } catch (e) {}

      if (props) {
        const meta = props.qMetaDef || {};
        const qDim = props.qDim || {};
        const fieldsArray = Array.isArray(qDim.qFieldDefs) ? qDim.qFieldDefs : [];
        const drillDownFieldsArray = Array.isArray(qDim.qDrillDownFieldDefs) ? qDim.qDrillDownFieldDefs : [];
        results.push({
          id,
          title: meta.title || it.qMeta?.title || '',
          description: meta.description || it.qMeta?.description || '',
          tags: (meta.tags || it.qMeta?.tags || []).join(', '),
          fields: fieldsArray.join(', '),
          fieldsArray,
          drillDownFieldsArray,
          labelExpr: qDim.qLabelExpression || '',
          usedCount: 0
        });
      } else {
        results.push({
          id,
          title: it.qMeta?.title || '',
          description: it.qMeta?.description || '',
          tags: (it.qMeta?.tags || []).join(', '),
          fields: '',
          fieldsArray: [],
          drillDownFieldsArray: [],
          labelExpr: '',
          usedCount: 0
        });
      }
    }

    return results;
  }

  async function fetchMeasuresViaEngine(app) {
    const doc = getDoc(app);
    const listHandle = await doc.createSessionObject({
      qInfo: { qType: 'MeasureList' },
      qMeasureListDef: { qType: 'measure', qData: { title: '/qMetaDef/title', description: '/qMetaDef/description', tags: '/qMetaDef/tags' } }
    });
    const layout = await listHandle.getLayout();
    const items = layout?.qMeasureList?.qItems || [];
    if (!items.length) return [];

    const results = [];
    for (const it of items) {
      const id = it.qInfo?.qId;
      if (!id) continue;

      let props = null;
      try {
        if (typeof app.getMeasure === 'function') {
          const msrModel = await app.getMeasure(id);
          props = await msrModel.getProperties();
        } else if (typeof getDoc(app).getMeasure === 'function') {
          const msrHandle = await getDoc(app).getMeasure({ qId: id });
          props = await msrHandle.getProperties();
        }
      } catch (e) {}

      if (props) {
        const meta = props.qMetaDef || {};
        const qMeasure = props.qMeasure || {};
        results.push({
          id,
          title: meta.title || it.qMeta?.title || '',
          description: meta.description || it.qMeta?.description || '',
          tags: (meta.tags || it.qMeta?.tags || []).join(', '),
          expression: qMeasure.qDef || '',
          label: qMeasure.qLabel || '',
          labelExpr: qMeasure.qLabelExpression || '',
          usedCount: 0
        });
      } else {
        results.push({
          id,
          title: it.qMeta?.title || '',
          description: it.qMeta?.description || '',
          tags: (it.qMeta?.tags || []).join(', '),
          expression: '',
          label: '',
          labelExpr: '',
          usedCount: 0
        });
      }
    }

    return results;
  }

  async function fetchFieldsViaEngine(app) {
    const doc = getDoc(app);
    const listHandle = await doc.createSessionObject({
      qInfo: { qType: 'FieldList' },
      qFieldListDef: { qShowSystem: false, qShowHidden: false, qShowSemantic: true, qShowDerivedFields: true, qShowImplicit: true, qShowSrcTables: true }
    });
    const layout = await listHandle.getLayout();
    const items = layout?.qFieldList?.qItems || [];
    return items.map(it => ({
      name: it.qName || '',
      tags: (it.qTags || []).join(', '),
      srcTables: Array.isArray(it.qSrcTables) ? it.qSrcTables.join(', ') : (it.qSrcTables || '')
    }));
  }

  // ------- Variables fetcher -------
  async function fetchVariablesViaEngine(app) {
    const doc = getDoc(app);
    let items = [];

    try {
      const h = await doc.createSessionObject({
        qInfo: { qType: 'VariableList' },
        qVariableListDef: {
          qType: 'variable',
          qShowReserved: true,
          qShowConfig: true,
          qData: { tags: '/qMetaDef/tags', definition: '/qDefinition', comment: '/qComment' }
        }
      });
      const layout = await h.getLayout();
      items = layout?.qVariableList?.qItems || [];
    } catch (e) {}

    if (!items.length) {
      items = await new Promise(resolve => {
        try {
          app.getList('VariableList', function (reply) {
            const arr = reply?.qVariableList?.qItems || [];
            if (reply?.qInfo?.qId) app.destroySessionObject(reply.qInfo.qId);
            resolve(arr);
          });
        } catch (e) {
          resolve([]);
        }
      });
    }

    if (!items.length) return [];

    const vars = [];
    for (const it of items) {
      const name = it.qName || it.qInfo?.qName || '';
      if (!name) continue;

      let props = null;
      try {
        if (app.variable && typeof app.variable.getByName === 'function') {
          const vm = await app.variable.getByName(name);
          props = await vm.getProperties();
        } else if (typeof getDoc(app).getVariableByName === 'function') {
          let vh = null;
          try {
            vh = await getDoc(app).getVariableByName(name);
          } catch (e1) {
            try {
              vh = await getDoc(app).getVariableByName({ qName: name });
            } catch (e2) {}
          }
          if (vh && typeof vh.getProperties === 'function') props = await vh.getProperties();
        }
      } catch (e) {}

      const definition = props?.qDefinition ?? it.qDefinition ?? it.qData?.definition ?? '';
      const comment = props?.qComment ?? it.qComment ?? it.qData?.comment ?? '';
      const tagsArr = props?.qMetaDef?.tags ?? it.qTags ?? it.qData?.tags ?? [];
      const tags = Array.isArray(tagsArr) ? tagsArr.join(', ') : (tagsArr || '');
      const isScript = props?.qIsScriptCreated ?? it.qIsScriptCreated ?? it.qIsScript;
      const isReserved = props?.qIsReserved ?? it.qIsReserved;

      vars.push({ name, definition, comment, tags, isScript: YN(isScript), isReserved: YN(isReserved) });
    }

    return vars;
  }

  // ------- helper: mark fields from primary dimensions in charts -------
  function markFieldsFromDimensions(props, addUse, allFieldsSet, markExpr) {
    const hc = props && props.qHyperCubeDef;
    if (!hc) return;
    const asArray = v => Array.isArray(v) ? v : (v != null ? [v] : []);

    for (const d of asArray(hc.qDimensions)) {
      const qd = d && d.qDef ? d.qDef : null;
      const defs = qd
        ? (Array.isArray(qd.qFieldDefs) && qd.qFieldDefs.length ? qd.qFieldDefs : (qd.qFieldDef ? [qd.qFieldDef] : []))
        : [];

      for (const f of defs) {
        const name = String(f || '').replace(/^\[|\]$/g, '').replace(/\.autoCalendar\..*$/, '');
        if (typeof f === 'string' && f.trim().startsWith('=')) {
          markExpr(f, 'Chart');
        } else {
          if (name && (!allFieldsSet || allFieldsSet.has(name))) addUse(name, 'Chart');
        }
      }

      if (qd?.qLabelExpression) markExpr(qd.qLabelExpression, 'Chart');
      if (d?.qCalcCond) markExpr(d.qCalcCond, 'Chart');
    }
  }

  // ------- helper: mark fields from alt dimensions/measures -------
  async function markFieldsFromLayoutExclude(app, props, addUse, allFieldsSet, markExpr) {
    const ex = (props && props.qLayoutExclude && props.qLayoutExclude.qHyperCubeDef)
      || (props && props.qHyperCubeDef && props.qHyperCubeDef.qLayoutExclude && props.qHyperCubeDef.qLayoutExclude.qHyperCubeDef)
      || null;
    if (!ex) return;

    const asArray = v => Array.isArray(v) ? v : (v != null ? [v] : []);

    for (const d of asArray(ex.qDimensions)) {
      const qd = d && d.qDef ? d.qDef : null;
      const defs = qd
        ? (Array.isArray(qd.qFieldDefs) && qd.qFieldDefs.length ? qd.qFieldDefs : (qd.qFieldDef ? [qd.qFieldDef] : []))
        : [];

      for (const f of defs) {
        const name = String(f || '').replace(/^\[|\]$/g, '').replace(/\.autoCalendar\..*$/, '');
        if (typeof f === 'string' && f.trim().startsWith('=')) {
          markExpr(f, 'Chart');
        } else {
          if (name && (!allFieldsSet || allFieldsSet.has(name))) addUse(name, 'Chart');
        }
      }

      if (d && typeof d.qLibraryId === 'string' && d.qLibraryId.trim()) {
        try {
          const dimModel = await app.getDimension(d.qLibraryId);
          const dp = await dimModel.getProperties();
          const arr = Array.isArray(dp?.qDim?.qFieldDefs) ? dp.qDim.qFieldDefs : [];
          const dd = Array.isArray(dp?.qDim?.qDrillDownFieldDefs) ? dp.qDim.qDrillDownFieldDefs : [];
          for (const f of [...arr, ...dd]) {
            const name = String(f || '').replace(/^\[|\]$/g, '').replace(/\.autoCalendar\..*$/, '');
            if (name && (!allFieldsSet || allFieldsSet.has(name))) addUse(name, 'Chart');
            if (typeof f === 'string' && f.trim().startsWith('=')) markExpr(f, 'Chart');
          }
          if (dp?.qDim?.qLabelExpression) markExpr(dp.qDim.qLabelExpression, 'Chart');
        } catch (e) {}
      }

      if (qd?.qLabelExpression) markExpr(qd.qLabelExpression, 'Chart');
      if (d?.qCalcCond) markExpr(d.qCalcCond, 'Chart');
    }

    for (const m of asArray(ex.qMeasures)) {
      if (m?.qDef?.qDef) markExpr(m.qDef.qDef, 'Chart');
      if (m?.qDef?.qLabelExpression) markExpr(m.qDef.qLabelExpression, 'Chart');
      if (m?.qSortByExpression?.qExpression) markExpr(m.qSortByExpression.qExpression, 'Chart');

      if (m && typeof m.qLibraryId === 'string' && m.qLibraryId.trim()) {
        try {
          const msrModel = await app.getMeasure(m.qLibraryId);
          const mp = await msrModel.getProperties();
          if (mp?.qMeasure?.qDef) markExpr(mp.qMeasure.qDef, 'Chart');
          if (mp?.qMeasure?.qLabelExpression) markExpr(mp.qMeasure.qLabelExpression, 'Chart');
        } catch (e) {}
      }
    }
  }

  // ===================== nested qHyperCubeDef support =====================
  function markFieldsFromHyperCubeDef(hc, addUse, allFieldsSet, markExpr, cat) {
    if (!hc) return;
    const asArray = v => Array.isArray(v) ? v : (v != null ? [v] : []);

    for (const d of asArray(hc.qDimensions)) {
      const qd = d && d.qDef ? d.qDef : null;
      const defs = qd
        ? (Array.isArray(qd.qFieldDefs) && qd.qFieldDefs.length ? qd.qFieldDefs : (qd.qFieldDef ? [qd.qFieldDef] : []))
        : [];

      for (const f of defs) {
        const raw = String(f || '').trim();
        if (!raw) continue;
        if (raw.startsWith('=')) {
          markExpr(raw, cat);
        } else {
          const name = raw.replace(/^\[|\]$/g, '').replace(/\.autoCalendar\..*$/, '');
          if (name && (!allFieldsSet || allFieldsSet.has(name))) addUse(name, cat);
        }
      }

      if (qd?.qLabelExpression) markExpr(qd.qLabelExpression, cat);
      if (d?.qCalcCond) markExpr(d.qCalcCond, cat);
    }

    for (const m of asArray(hc.qMeasures)) {
      if (m?.qDef?.qDef) markExpr(m.qDef.qDef, cat);
      if (m?.qDef?.qLabelExpression) markExpr(m.qDef.qLabelExpression, cat);
      if (m?.qSortByExpression?.qExpression) markExpr(m.qSortByExpression.qExpression, cat);
    }

    if (hc.qCalcCond) markExpr(hc.qCalcCond, cat);
  }

  async function markFieldsFromAllHypercubes(app, props, addUse, allFieldsSet, markExpr) {
    const seen = new WeakSet();
    const MAX_DEPTH = 60;

    async function walk(node, depth) {
      if (!node || depth > MAX_DEPTH) return;
      if (typeof node !== 'object') return;
      if (seen.has(node)) return;
      seen.add(node);

      if (node.qHyperCubeDef) {
        markFieldsFromHyperCubeDef(node.qHyperCubeDef, addUse, allFieldsSet, markExpr, 'Chart');
      }

      await markFieldsFromLayoutExclude(app, node, addUse, allFieldsSet, markExpr);

      if (Array.isArray(node)) {
        for (const v of node) await walk(v, depth + 1);
      } else {
        for (const v of Object.values(node)) {
          if (v && typeof v === 'object') await walk(v, depth + 1);
        }
      }
    }

    await walk(props, 0);
  }

  // ===================== UNUSED FIELDS + USED IN =====================
  async function findUnusedFields(app, allFields, dims, msrs, vars) {
    const allFieldsSet = new Set((allFields || []).map(f => f.name).filter(Boolean));
    const varMap = buildVariableMap(vars);
    const used = new Set();
    const usedInMap = new Map();

    const keyFields = new Set(
      (allFields || [])
        .filter(f => {
          const t = String(f?.tags || '').toLowerCase();
          return /\$key\b/.test(t);
        })
        .map(f => f.name)
        .filter(Boolean)
    );

    const addUse = (fname, cat) => {
      if (!fname) return;
      used.add(fname);
      let s = usedInMap.get(fname);
      if (!s) { s = new Set(); usedInMap.set(fname, s); }
      s.add(cat);
    };

    for (const k of keyFields) addUse(k, 'Key');

    const markExpr = (expr, cat) => {
      const ex = expandDollar(expr, varMap);
      const fields = extractFieldsFromExpr(ex, allFieldsSet);
      const hasSet = reHasSet.test(ex || '');
      fields.forEach(f => {
        addUse(f, cat);
        if (hasSet) addUse(f, 'Set analysis');
      });
    };

    for (const d of dims || []) {
      const rawList = Array.isArray(d.fieldsArray) ? d.fieldsArray : (typeof d.fields === 'string' ? d.fields.split(',') : []);
      const ddList = Array.isArray(d.drillDownFieldsArray) ? d.drillDownFieldsArray : [];
      const allDimDefs = [...rawList, ...ddList].map(s => String(s || '').trim()).filter(Boolean);

      for (const def of allDimDefs) {
        if (def.startsWith('=')) markExpr(def, 'Dimension');
        else {
          const base = def.replace(/^\[|\]$/g, '').replace(/\.autoCalendar\..*$/, '');
          if (allFieldsSet.has(base)) addUse(base, 'Dimension');
          const ex = /^\[.*\]$/.test(def) ? def : `[${def}]`;
          markExpr(ex, 'Dimension');
        }
      }

      if (d.labelExpr) markExpr(d.labelExpr, 'Dimension');
    }

    for (const m of msrs || []) {
      if (m.expression) markExpr(m.expression, 'Measure');
      if (m.labelExpr) markExpr(m.labelExpr, 'Measure');
    }

    for (const v of vars || []) {
      if (!v?.definition) continue;
      markExpr(v.definition, 'Variable');
    }

    const objs = await fetchObjectPropsForSheets(app);
    for (const o of objs) {
      const props = o.props || {};
      const items = collectObjectExpressions(props);
      for (const it of items) markExpr(it.expr, 'Chart');

      await markFieldsFromAllHypercubes(app, props, addUse, allFieldsSet, markExpr);
      markFieldsFromDimensions(props, addUse, allFieldsSet, markExpr);
      await markFieldsFromLayoutExclude(app, props, addUse, allFieldsSet, markExpr);
    }

    for (const f of allFieldsSet) {
      if (f.includes('.autoCalendar.')) {
        const base = f.replace(/\.autoCalendar\..*$/, '');
        if (used.has(f)) addUse(base, 'Chart');
      }
    }

    const unused = new Set();
    for (const f of allFieldsSet) {
      if (f.includes('.autoCalendar.')) continue;
      if (keyFields.has(f)) continue;
      if (!used.has(f)) unused.add(f);
    }

    return { used, unused, usedInMap, keyFields };
  }

  // ------- Master item caches & resolvers -------
  const masterDimCache = new Map();
  const masterMsrCache = new Map();

  async function warmMasterCaches(app) {
    try {
      const dims = await fetchDimensionsViaEngine(app);
      for (let i = 0; i < dims.length; i++) {
        const d = dims[i];
        masterDimCache.set(d.id, {
          id: d.id,
          title: d.title || (Array.isArray(d.fieldsArray) && d.fieldsArray[0]) || d.id,
          levels: (Array.isArray(d.drillDownFieldsArray) && d.drillDownFieldsArray.length ? d.drillDownFieldsArray : (d.fieldsArray || []))
        });
      }
    } catch (e) {}

    try {
      const msrs = await fetchMeasuresViaEngine(app);
      for (let j = 0; j < msrs.length; j++) {
        const m = msrs[j];
        masterMsrCache.set(m.id, {
          id: m.id,
          title: m.title || m.label || m.id,
          expr: m.expression || '',
          label: m.label || ''
        });
      }
    } catch (e) {}
  }

  async function resolveMasterDimension(app, id) {
    if (!id) return null;
    if (masterDimCache.has(id)) return masterDimCache.get(id);

    try {
      const dm = await app.getDimension(id);
      const [p, l] = await Promise.all([
        dm.getProperties().catch(() => null),
        dm.getLayout().catch(() => null)
      ]);

      const fieldsArr = (Array.isArray(p?.qDim?.qDrillDownFieldDefs) && p.qDim.qDrillDownFieldDefs.length
        ? p.qDim.qDrillDownFieldDefs
        : (Array.isArray(p?.qDim?.qFieldDefs) ? p.qDim.qFieldDefs : []))
        .map(x => String(x || '').replace(/^\[|\]$/g, ''));

      const title = (p?.qMetaDef?.title) || (l?.qMeta?.title) || (fieldsArr[0] || id);
      const out = { id, title, levels: fieldsArr };
      masterDimCache.set(id, out);
      return out;
    } catch (e) {
      if (!masterDimCache.has(id)) masterDimCache.set(id, { id, title: id, levels: [] });
      return masterDimCache.get(id);
    }
  }

  async function resolveMasterMeasure(app, id) {
    if (!id) return null;
    if (masterMsrCache.has(id)) return masterMsrCache.get(id);

    try {
      const mm = await app.getMeasure(id);
      const [p, l] = await Promise.all([
        mm.getProperties().catch(() => null),
        mm.getLayout().catch(() => null)
      ]);

      const expr = p?.qMeasure?.qDef || '';
      const label = p?.qMeasure?.qLabel || l?.qMeasure?.qLabel || '';
      const title = p?.qMetaDef?.title || l?.qMeta?.title || label || id;
      const out = { id, title, expr, label };
      masterMsrCache.set(id, out);
      return out;
    } catch (e) {
      if (!masterMsrCache.has(id)) masterMsrCache.set(id, { id, title: id, expr: '', label: '' });
      return masterMsrCache.get(id);
    }
  }

  function sanitizeInlineText(value) {
    return String(value == null ? '' : value)
      .replace(/[\r\n]+/g, ' ')
      .replace(/\s{2,}/g, ' ')
      .trim();
  }

  const trunc = (s, n = 120) => {
    const str = sanitizeInlineText(s || '');
    return str.length > n ? (str.slice(0, n - 1) + '…') : str;
  };

  // ---------- Items summary ----------
  function extractAltBlocks(src) {
    if (!src) return { dims: [], msrs: [] };
    const ex = (src.qLayoutExclude && src.qLayoutExclude.qHyperCubeDef)
      || (src.qHyperCubeDef && src.qHyperCubeDef.qLayoutExclude && src.qHyperCubeDef.qLayoutExclude.qHyperCubeDef)
      || null;
    return { dims: (ex && ex.qDimensions) || [], msrs: (ex && ex.qMeasures) || [] };
  }

  function dedupeByKey(arr) {
    const out = [];
    const seen = new Set();
    for (const it of arr || []) {
      const key = it && it.qLibraryId ? ('lib:' + it.qLibraryId) : ('def:' + JSON.stringify(it?.qDef || {}));
      if (seen.has(key)) continue;
      seen.add(key);
      out.push(it);
    }
    return out;
  }

  const fmtMaster = item => {
    const title = sanitizeInlineText(item?.title || '(no title)');
    const id = sanitizeInlineText(item?.id || '');
    return id
      ? `[Master Item: ${title}, ID: ${id}]`
      : `[Master Item: ${title}]`;
  };

  const fmtField = name => `[Field: ${sanitizeInlineText(name || '')}]`;

  const fmtFieldWithLabel = (fieldText, labelText) => {
    const f = sanitizeInlineText(fieldText || '');
    const l = sanitizeInlineText(labelText || '');
    if (!l || !f) return `[Field: ${f || l}]`;
    if (l === f) return `[Field: ${f}]`;
    return `[Field: ${f}, Label: ${l}]`;
  };

  const fmtExpr = expr => `[Expression: ${trunc(expr)}]`;

  const fmtExprWithLabel = (exprText, labelText) => {
    const e = sanitizeInlineText(exprText || '');
    const l = sanitizeInlineText(labelText || '');
    if (!l) return `[Expression: ${trunc(e)}]`;
    return `[Expression: ${trunc(e)}, Label: ${l}]`;
  };

  function findFirstNestedHyperCubeDef(root) {
    const seen = new WeakSet();
    const MAX_DEPTH = 60;

    const hasContent = hc =>
      !!hc && ((Array.isArray(hc.qDimensions) && hc.qDimensions.length) || (Array.isArray(hc.qMeasures) && hc.qMeasures.length));

    function walk(node, depth) {
      if (!node || depth > MAX_DEPTH) return null;
      if (typeof node !== 'object') return null;
      if (seen.has(node)) return null;
      seen.add(node);

      if (node.qHyperCubeDef) {
        if (hasContent(node.qHyperCubeDef)) return node.qHyperCubeDef;
        const candidate = node.qHyperCubeDef;

        if (Array.isArray(node)) {
          for (const v of node) {
            const hit = walk(v, depth + 1);
            if (hit) return hit;
          }
        } else {
          for (const v of Object.values(node)) {
            const hit = walk(v, depth + 1);
            if (hit) return hit;
          }
        }

        return candidate;
      }

      if (Array.isArray(node)) {
        for (const v of node) {
          const hit = walk(v, depth + 1);
          if (hit) return hit;
        }
      } else {
        for (const v of Object.values(node)) {
          const hit = walk(v, depth + 1);
          if (hit) return hit;
        }
      }

      return null;
    }

    return walk(root, 0);
  }

  function findFirstNestedLayoutHyperCube(layout) {
    const seen = new WeakSet();
    const MAX_DEPTH = 60;

    function walk(node, depth) {
      if (!node || depth > MAX_DEPTH) return null;
      if (typeof node !== 'object') return null;
      if (seen.has(node)) return null;
      seen.add(node);

      if (node.qHyperCube) return node.qHyperCube;

      if (Array.isArray(node)) {
        for (const v of node) {
          const hit = walk(v, depth + 1);
          if (hit) return hit;
        }
      } else {
        for (const v of Object.values(node)) {
          const hit = walk(v, depth + 1);
          if (hit) return hit;
        }
      }

      return null;
    }

    return walk(layout, 0);
  }

  async function buildItemsSummary(app, props, layout) {
    let baseProps = props || {};

    if ((!baseProps.qHyperCubeDef || !Array.isArray(baseProps.qHyperCubeDef.qMeasures)) && baseProps.qExtendsId) {
      try {
        const masterModel = await app.getObject(baseProps.qExtendsId);
        const mp = await masterModel.getProperties();
        if (mp) baseProps = { ...mp, qLayoutExclude: (baseProps.qLayoutExclude || mp.qLayoutExclude) };
      } catch (e) {}
    }

    let hc = baseProps.qHyperCubeDef || null;
    const hcHasContent = h =>
      !!h && ((Array.isArray(h.qDimensions) && h.qDimensions.length) || (Array.isArray(h.qMeasures) && h.qMeasures.length));

    if (!hcHasContent(hc)) {
      const nested = findFirstNestedHyperCubeDef(baseProps);
      if (nested) hc = nested;
    }
    hc = hc || {};

    const pAlt = extractAltBlocks(baseProps);
    const lAlt = extractAltBlocks(layout);
    const alt = dedupeByKey([...(pAlt.dims || []), ...(lAlt.dims || [])]);
    const altMs = dedupeByKey([...(pAlt.msrs || []), ...(lAlt.msrs || [])]);

    const layoutHC = layout?.qHyperCube || findFirstNestedLayoutHyperCube(layout) || null;

    const dimInfos = (layoutHC && Array.isArray(layoutHC.qDimensionInfo)) ? layoutHC.qDimensionInfo : [];
    const msrInfos = (layoutHC && Array.isArray(layoutHC.qMeasureInfo)) ? layoutHC.qMeasureInfo : [];

    const sections = [];
    const pushSection = (title, lines) => {
      if (!lines.length) return;
      sections.push(`${title}`);
      sections.push(...lines.map(s => `   ${s}`));
    };

    const dimLines = [];
    if (Array.isArray(hc.qDimensions)) {
      for (let i = 0; i < hc.qDimensions.length; i++) {
        const d = hc.qDimensions[i];
        const visName = sanitizeInlineText((dimInfos[i] && (dimInfos[i].qFallbackTitle || dimInfos[i].qGroupFieldDefs?.[0])) || '');
        if (d.qLibraryId) {
          const md = await resolveMasterDimension(app, d.qLibraryId);
          const title = visName || md?.title || (md?.levels?.[0]) || d.qLibraryId;
          dimLines.push(`• ${fmtMaster({ title, id: d.qLibraryId })}`);
        } else {
          const defs = Array.isArray(d?.qDef?.qFieldDefs) && d.qDef.qFieldDefs.length
            ? d.qDef.qFieldDefs
            : (d?.qDef?.qFieldDef ? [d.qDef.qFieldDef] : []);

          const fieldText = defs
            .map(x => sanitizeInlineText(String(x || '').replace(/^\[|\]$/g, '')))
            .filter(Boolean)
            .join('→') || 'Field';

          dimLines.push(`• ${fmtFieldWithLabel(fieldText, visName)}`);
        }
      }
    }

    const msrLines = [];
    if (Array.isArray(hc.qMeasures)) {
      for (let i = 0; i < hc.qMeasures.length; i++) {
        const m = hc.qMeasures[i];
        const visName = sanitizeInlineText((msrInfos[i] && msrInfos[i].qFallbackTitle) || '');
        if (m.qLibraryId) {
          const mm = await resolveMasterMeasure(app, m.qLibraryId);
          const title = visName || mm?.label || mm?.title || m.qLibraryId;
          msrLines.push(`• ${fmtMaster({ title, id: m.qLibraryId })}`);
        } else {
          const expr = m?.qDef?.qDef || '';
          const label = sanitizeInlineText((m?.qDef?.qLabel || visName || '').trim());
          msrLines.push(`• ${fmtExprWithLabel(expr, label)}`);
        }
      }
    }

    const altDimLines = [];
    for (const d of alt) {
      if (d.qLibraryId) {
        const md = await resolveMasterDimension(app, d.qLibraryId);
        const title = md?.title || (md?.levels?.[0]) || d.qLibraryId;
        altDimLines.push(`• Alt: ${fmtMaster({ title, id: d.qLibraryId })}`);
      } else {
        const defs = Array.isArray(d?.qDef?.qFieldDefs) && d.qDef.qFieldDefs.length
          ? d.qDef.qFieldDefs
          : (d?.qDef?.qFieldDef ? [d.qDef.qFieldDef] : []);
        const fieldText = defs.map(x => sanitizeInlineText(String(x || '').replace(/^\[|\]$/g, ''))).filter(Boolean).join('→') || 'Field';
        const label = sanitizeInlineText((d?.qDef?.qLabel || d?.qDef?.qFallbackTitle || '').trim());
        altDimLines.push(`• Alt: ${fmtFieldWithLabel(fieldText, label)}`);
      }
    }

    const altMsrLines = [];
    for (const m of altMs) {
      if (m.qLibraryId) {
        const mm = await resolveMasterMeasure(app, m.qLibraryId);
        const title = mm?.label || mm?.title || m.qLibraryId;
        altMsrLines.push(`• Alt: ${fmtMaster({ title, id: m.qLibraryId })}`);
      } else {
        const expr = m?.qDef?.qDef || '';
        const label = sanitizeInlineText((m?.qDef?.qLabel || '').trim());
        altMsrLines.push(`• Alt: ${fmtExprWithLabel(expr, label)}`);
      }
    }

    pushSection(`Dimensions (${dimLines.length})`, dimLines);
    pushSection(`Measures (${msrLines.length})`, msrLines);
    if (altDimLines.length) pushSection(`Alternate Dimensions (${altDimLines.length})`, altDimLines);
    if (altMsrLines.length) pushSection(`Alternate Measures (${altMsrLines.length})`, altMsrLines);

    return sections.join('\r\n');
  }

  // ===================== Container support =====================
  function safeTitleFromProps(p) {
    if (!p) return '';
    if (typeof p.title === 'string') return p.title;
    if (p.title && typeof p.title.qStringExpression === 'string') return p.title.qStringExpression;
    if (p.qMetaDef && typeof p.qMetaDef.title === 'string') return p.qMetaDef.title;
    return '';
  }

  function isGuidLike(s) {
    const t = String(s || '').trim();
    if (!t) return false;
    return /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(t);
  }

  function prettyVisName(vis) {
    const v = String(vis || '').trim().toLowerCase();
    const map = {
      barchart: 'Bar chart',
      linechart: 'Line chart',
      combochart: 'Combo chart',
      piechart: 'Pie chart',
      table: 'Table',
      pivot: 'Pivot table',
      pivotchart: 'Pivot table',
      kpi: 'KPI',
      filterpane: 'Filter pane',
      listbox: 'List box',
      treemap: 'Treemap',
      map: 'Map',
      scatterplot: 'Scatter plot',
      distributionplot: 'Distribution plot',
      boxplot: 'Box plot',
      gauge: 'Gauge',
      textimage: 'Text and image',
      container: 'Container',
      'sn-layout-container': 'Layout container',
      button: 'Action button'
    };
    if (map[v]) return map[v];

    const cleaned = v
      .replace(/[_-]+/g, ' ')
      .replace(/([a-z])([A-Z])/g, '$1 $2')
      .trim();
    if (!cleaned) return 'Chart';
    return cleaned.charAt(0).toUpperCase() + cleaned.slice(1);
  }

  function isContainerLikeType(visType) {
    const t = String(visType || '').trim().toLowerCase();
    return t === 'container' || t === 'sn-layout-container';
  }

  function getCandidateObjectId(node) {
    if (!node || typeof node !== 'object') return '';

    const direct =
      node.qId ||
      node.id ||
      node.objectId ||
      node.qObjectId ||
      '';

    if (direct && typeof direct === 'string') return direct;

    const nested =
      node.qInfo?.qId ||
      node.qChild?.qId ||
      '';

    return (typeof nested === 'string') ? nested : '';
  }

  function getCandidateObjectType(node) {
    if (!node || typeof node !== 'object') return '';
    return (
      node.visualization ||
      node.type ||
      node.qType ||
      node.chartType ||
      node.component ||
      node.qData?.visualization ||
      node.qData?.type ||
      ''
    );
  }

  function getCandidateObjectTitle(node) {
    if (!node || typeof node !== 'object') return '';
    return (
      node.title ||
      node.label ||
      node.qMeta?.title ||
      node.qMeta?.name ||
      node.qData?.title ||
      ''
    );
  }

  function collectContainerCandidates(root, selfId) {
    const out = [];
    const seen = new WeakSet();
    const addedIds = new Set();
    const MAX_DEPTH = 20;

    function maybeAdd(node) {
      if (!node || typeof node !== 'object') return;

      const id = String(getCandidateObjectId(node) || '').trim();
      if (!id) return;
      if (selfId && id === selfId) return;

      const type = String(getCandidateObjectType(node) || '').trim();
      const title = String(getCandidateObjectTitle(node) || '').trim();

      const looksLikeRealObject =
        !!type ||
        !!title ||
        !!node.qInfo?.qId ||
        !!node.qMeta?.title ||
        !!node.qData?.visualization ||
        !!node.qData?.type;

      if (!looksLikeRealObject) return;

      if (addedIds.has(id)) return;
      addedIds.add(id);

      out.push({ id, type, title });
    }

    function walk(node, depth) {
      if (!node || depth > MAX_DEPTH) return;
      if (typeof node !== 'object') return;
      if (seen.has(node)) return;
      seen.add(node);

      if (Array.isArray(node)) {
        for (const item of node) {
          maybeAdd(item);
          walk(item, depth + 1);
        }
        return;
      }

      const interestingKeys = ['items', 'children', 'childItems', 'tabs', 'panels', 'content', 'components'];
      for (const k of interestingKeys) {
        if (node[k] && (Array.isArray(node[k]) || typeof node[k] === 'object')) {
          walk(node[k], depth + 1);
        }
      }

      for (const [k, v] of Object.entries(node)) {
        if (!v || typeof v !== 'object') continue;

        if (
          k === 'qHyperCubeDef' ||
          k === 'qHyperCube' ||
          k === 'qLayoutExclude' ||
          k === 'qDataPages' ||
          k === 'qStackedDataPages' ||
          k === 'qGrandTotalRow'
        ) {
          continue;
        }

        maybeAdd(v);
        walk(v, depth + 1);
      }
    }

    walk(root, 0);
    return out;
  }

  // ------- Dynamic extension detection -------
  function getInstalledExtensionIdSet() {
    return new Promise((resolve) => {
      try {
        if (!qlik || typeof qlik.getExtensionList !== 'function') {
          resolve(new Set());
          return;
        }
        qlik.getExtensionList(function (list) {
          const set = new Set();
          (list || []).forEach(x => {
            const id = String(x?.id || x?.value?.id || '').trim().toLowerCase();
            if (id) set.add(id);
          });
          resolve(set);
        });
      } catch (e) {
        resolve(new Set());
      }
    });
  }

  function containerChildDisplayTitle(rawTitle, childType, childId) {
    const t = String(rawTitle || '').trim();
    const hasRealTitle = t && t !== String(childId || '').trim() && !isGuidLike(t);
    if (hasRealTitle) return { text: t, isFallback: false };
    return { text: `Container > ${prettyVisName(childType)}`, isFallback: true };
  }

  function extractContainerChildTitleMap(layout, props, selfId) {
    const map = new Map();

    const pushItem = (id, title, type) => {
      const key = String(id || '').trim();
      if (!key) return;
      if (selfId && key === selfId) return;

      const cleanTitle = String(title || '').trim();
      const cleanType = String(type || '').trim();

      if (!map.has(key)) {
        map.set(key, {
          title: cleanTitle,
          type: cleanType
        });
        return;
      }

      const existing = map.get(key);
      if (!existing.title && cleanTitle) existing.title = cleanTitle;
      if (!existing.type && cleanType) existing.type = cleanType;
    };

    const items =
      layout?.qChildList?.qItems ||
      layout?.qChildListObject?.qItems ||
      layout?.qChildren?.qItems ||
      [];

    if (Array.isArray(items)) {
      for (const it of items) {
        const id =
          it?.qInfo?.qId ||
          it?.qId ||
          it?.id ||
          it?.qIdentifier ||
          '';

        const title =
          it?.qMeta?.title ||
          it?.qMeta?.name ||
          it?.qData?.title ||
          it?.title ||
          '';

        const type =
          it?.qData?.visualization ||
          it?.qData?.type ||
          it?.qType ||
          it?.type ||
          '';

        pushItem(id, title, type);
      }
    }

    const fallbackCandidates = [
      ...collectContainerCandidates(layout, selfId),
      ...collectContainerCandidates(props, selfId)
    ];

    for (const c of fallbackCandidates) {
      pushItem(c.id, c.title, c.type);
    }

    return map;
  }

  function dedupeStrings(arr) {
    const out = [];
    const seen = new Set();
    for (const v of arr || []) {
      const s = String(v || '').trim();
      if (!s || seen.has(s)) continue;
      seen.add(s);
      out.push(s);
    }
    return out;
  }

  async function buildContainerItemsSummary(app, containerModel, containerProps, containerLayout, depth = 0) {
    if (depth > 3) return '';

    const selfId =
      containerLayout?.qInfo?.qId ||
      containerProps?.qInfo?.qId ||
      containerModel?.id ||
      '';

    let kids = [];
    try { kids = await containerModel.getChildInfos(); } catch (e) {}

    const idsFromChildInfos = Array.isArray(kids)
      ? kids
          .map(k => k?.qId || k?.id || k?.Id || '')
          .filter(Boolean)
          .filter(id => String(id) !== String(selfId))
      : [];

    const titleMap = extractContainerChildTitleMap(containerLayout, containerProps, selfId);
    const idsFromLayout = Array.from(titleMap.keys()).filter(id => String(id) !== String(selfId));

    const idsFromPropsCells = Array.isArray(containerProps?.cells)
      ? containerProps.cells
          .map(c => c?.name || c?.qId || c?.id || '')
          .filter(Boolean)
          .filter(id => String(id) !== String(selfId))
      : [];

    const childIds = dedupeStrings([...idsFromChildInfos, ...idsFromLayout, ...idsFromPropsCells]);
    if (!childIds.length) return '';

    const lines = [];
    lines.push(`Container items (${childIds.length})`);

    for (const cid of childIds) {
      try {
        const childModel = await app.getObject(cid);
        const [cp, cl] = await Promise.all([
          childModel.getProperties().catch(() => null),
          childModel.getLayout().catch(() => null)
        ]);

        const mapInfo = titleMap.get(cid) || {};
        const cType = cp?.visualization || cl?.visualization || mapInfo.type || '';

        const rawTitle =
          safeTitleFromProps(cp) ||
          cl?.qMeta?.title ||
          mapInfo.title ||
          '';

        const disp = containerChildDisplayTitle(rawTitle, cType, cid);

        let childSummary = cp ? await buildItemsSummary(app, cp, cl) : '';

        if ((!childSummary || !childSummary.trim()) && isContainerLikeType(cType) && String(cid) !== String(selfId)) {
          childSummary = await buildContainerItemsSummary(app, childModel, cp, cl, depth + 1);
        }

        const headerLine = disp.isFallback
          ? `• ${disp.text}`
          : `• ${disp.text}${cType ? ` (${cType})` : ''}`;

        lines.push(headerLine);

        if (childSummary && childSummary.trim()) {
          const indented = childSummary
            .split(/\r?\n/)
            .map(x => `   ${x}`)
            .join('\r\n');
          lines.push(indented);
        }
      } catch (e) {}
    }

    return lines.join('\r\n');
  }

  // ------- Script metadata helpers -------
  function extractQvdsFromLine(line) {
    const hits = [];

    const bracketSegs = (line.match(/\[([^\]]+)\]/g) || []);
    for (const seg of bracketSegs) {
      const inner = seg.slice(1, -1);
      const last = inner.split(/[\/\\]/).pop();
      if (last && /\.qvd\b/i.test(last)) hits.push(last);
    }

    const quotedSegs = (line.match(/["']([^"']*?\.qvd[^"']*)["']/ig) || []);
    for (const q of quotedSegs) {
      const inner = q.replace(/^["']|["']$/g, '');
      const last = inner.split(/[\/\\]/).pop();
      if (last && /\.qvd\b/i.test(last)) hits.push(last);
    }

    const bare = (line.match(/\b[^\s"'()[\];,]+\.qvd\b/ig) || []);
    for (let tok of bare) {
      tok = tok.split(/[\/\\]/).pop().replace(/[\]\)\.;,]+$/g, '');
      if (tok && /\.qvd\b/i.test(tok)) hits.push(tok);
    }

    const seen = new Set();
    const out = [];
    for (const h of hits) {
      const clean = String(h).trim();
      if (clean && !seen.has(clean)) {
        seen.add(clean);
        out.push(clean);
      }
    }
    return out;
  }

  function parseScriptMetadata(scriptText) {
    const rows = [];

    const withoutBlocks = String(scriptText || '').replace(/\/\*[\s\S]*?\*\//g, '');
    const lines = withoutBlocks.replace(/\r\n/g, '\n').split('\n');

    let current = null;

    const makeRow = (tabName) => ({
      tab: (tabName && tabName.trim()) || 'Main',
      loads: 0,
      stores: 0,
      joins: 0,
      residents: 0,
      qvds: new Set(),
      vars: new Set()
    });

    const flush = () => {
      if (!current) return;
      if (
        current.loads || current.stores || current.joins || current.residents ||
        (current.vars && current.vars.size) || (current.qvds && current.qvds.size)
      ) {
        rows.push({
          tab: current.tab,
          loads: current.loads,
          stores: current.stores,
          joins: current.joins,
          residents: current.residents,
          qvds: Array.from(current.qvds),
          vars: Array.from(current.vars)
        });
      }
      current = null;
    };

    for (const raw of lines) {
      const trimmed = (raw || '').trim();

      const tabMatch =
        trimmed.match(/^\/\/\/\s*\$tab\s*(.*)$/i) ||
        trimmed.match(/^\/\/\s*\$tab\s*(.*)$/i);

      if (tabMatch) {
        flush();
        current = makeRow(tabMatch[1] || 'Untitled');
        continue;
      }

      const line = (raw || '').replace(/\/\/.*$/, '').trim();
      if (!line) continue;
      if (!current) current = makeRow('Main');

      if (/\bLOAD\b/i.test(line)) current.loads++;
      if (/\bSTORE\b/i.test(line)) current.stores++;
      if (/\bJOIN\b/i.test(line)) current.joins++;
      if (/\bRESIDENT\b/i.test(line)) current.residents++;

      const qvds = extractQvdsFromLine(raw || '');
      qvds.forEach(q => current.qvds.add(q));

      const vm = line.match(/^\s*(SET|LET)\s+([A-Za-z_]\w*)\s*=/i);
      if (vm && vm[2]) current.vars.add(vm[2]);
    }

    flush();
    return rows;
  }

  // ------- Charts fetcher -------
  async function fetchChartsViaEngine(app) {
    await warmMasterCaches(app);

    const extIdSet = await getInstalledExtensionIdSet();
    const sheets = await fetchSheetsViaEngine(app);
    const charts = [];

    for (const sh of sheets) {
      try {
        const sheetModel = await app.getObject(sh.id);
        let childInfos = [];
        try { childInfos = await sheetModel.getChildInfos(); } catch (e) {}

        const props = childInfos?.length ? null : await sheetModel.getProperties().catch(() => null);
        const cells = childInfos?.length ? childInfos.map(ci => ({ name: ci.qId })) : (props?.cells || []);

        for (const c of cells) {
          const objId = c.name;
          try {
            const objModel = await app.getObject(objId);
            const p = await objModel.getProperties().catch(() => null);
            const l = await objModel.getLayout().catch(() => null);

            const rawType = (p?.visualization || c.type || l?.visualization || '');
            const visType = String(rawType || '').toLowerCase();

            let itemsSummary = p ? await buildItemsSummary(app, p, l) : '';

            if (isContainerLikeType(visType)) {
              const containerSummary = await buildContainerItemsSummary(app, objModel, p, l);
              if (containerSummary && containerSummary.trim()) itemsSummary = containerSummary;
            }

            const isExtension = extIdSet.size ? (extIdSet.has(visType) ? 'Y' : 'N') : '';

            charts.push({
              sheetTitle: sh.title,
              sheetId: sh.id,
              objectId: objId,
              type: rawType,
              isExtension,
              title: (typeof p?.title === 'string' ? p.title : (p?.title && p.title.qStringExpression) || ''),
              isMaster: p?.qExtendsId ? 'Y' : 'N',
              masterId: p?.qExtendsId || '',
              itemsSummary
            });
          } catch (e) {}
        }
      } catch (e) {}
    }

    return charts;
  }

  // ------- Export orchestrator -------
  function getAppIdSafe(app) {
    return app?.id
      || app?.model?.enigmaModel?.id
      || app?.model?.enigmaModel?.appId
      || app?.model?.id
      || '';
  }

  // ===================== Export Selected (XLSX) =====================
  async function exportSelectedXlsx(app, fileName, opts) {
    if (!ExcelJS) {
      alert('XLSX export unavailable: ExcelJS was not loaded. Please add exceljs.min.js into your extension folder and reference it as ./lib/exceljs.min in define().');
      return;
    }

    const { overview, dims, msrs, flds, shts, chrs, vars, scrmeta } = opts;
    if (!overview && !dims && !msrs && !flds && !shts && !chrs && !vars && !scrmeta) {
      alert('Nothing selected to export.');
      return;
    }

    const needDims = overview || dims || flds;
    const needMsrs = overview || msrs || flds;
    const needVars = overview || vars || flds;
    const needFlds = overview || flds;
    const needShts = overview || shts;
    const needChrs = overview || chrs;

    const [DIMS, MSRS, VARS, FLDS, SHTS, CHRS] = await Promise.all([
      needDims ? fetchDimensionsViaEngine(app) : Promise.resolve([]),
      needMsrs ? fetchMeasuresViaEngine(app) : Promise.resolve([]),
      needVars ? fetchVariablesViaEngine(app) : Promise.resolve([]),
      needFlds ? fetchFieldsViaEngine(app) : Promise.resolve([]),
      needShts ? fetchSheetsViaEngine(app) : Promise.resolve([]),
      needChrs ? fetchChartsViaEngine(app) : Promise.resolve([])
    ]);

    if (dims || msrs) {
      const { dimCount, msrCount } = await computeMasterUsage(app);
      for (const d of DIMS) d.usedCount = dimCount.get(d.id) || 0;
      for (const m of MSRS) m.usedCount = msrCount.get(m.id) || 0;
    }

    let unusedSet = null;
    let usedInMap = null;
    if (flds) {
      const res = await findUnusedFields(app, FLDS, DIMS, MSRS, VARS);
      unusedSet = res.unused;
      usedInMap = res.usedInMap;
    }

    let scriptMetaRows = null;
    let scriptDenied = false;
    if (scrmeta) {
      try {
        const res = await app.getScript();
        const text = (res && res.qScript) ? res.qScript : (typeof res === 'string' ? res : '');
        if (text) scriptMetaRows = parseScriptMetadata(text);
        else scriptDenied = true;
      } catch (e) {
        console.error('Script metadata extraction failed:', e);
        scriptDenied = true;
      }
    }

    const workbook = new ExcelJS.Workbook();
    workbook.creator = 'Qollect';
    workbook.created = new Date();

    if (overview) {
      let appName = '';
      let appId = getAppIdSafe(app);
      try {
        const layout = await getDoc(app).getAppLayout();
        appName = layout?.qTitle || '';
      } catch (e) {}

      const headers = [
        'Application name', 'Application ID',
        '# of Dimensions', '# of Measures', '# of Fields',
        '# of Sheets', '# of Charts', '# of Variables'
      ];
      const matrix = [[
        appName || '', appId || '',
        Number(DIMS.length || 0), Number(MSRS.length || 0), Number(FLDS.length || 0),
        Number(SHTS.length || 0), Number(CHRS.length || 0), Number(VARS.length || 0)
      ]];
      addWorksheetXlsx(workbook, { name: 'App Overview', headers, matrix });
    }

    if (dims) {
      const headers = ['ID', 'Title', 'Fields', 'Label Expression', 'Description', 'Tags', 'Used Count'];
      const dd = sortAsc(DIMS, 'title');
      const matrix = dd.map(d => [d.id, d.title, d.fields, d.labelExpr, d.description, d.tags, Number(d.usedCount || 0)]);
      addWorksheetXlsx(workbook, {
        name: 'Dimensions',
        headers,
        matrix,
        rowStyleId: r => (Number(r?.[6]) === 0 ? 'sUnused' : null)
      });
    }

    if (msrs) {
      const headers = ['ID', 'Title', 'Expression', 'Label', 'Label Expression', 'Description', 'Tags', 'Used Count'];
      const mm = sortAsc(MSRS, 'title');
      const matrix = mm.map(m => [m.id, m.title, m.expression, m.label, m.labelExpr, m.description, m.tags, Number(m.usedCount || 0)]);
      addWorksheetXlsx(workbook, {
        name: 'Measures',
        headers,
        matrix,
        rowStyleId: r => (Number(r?.[7]) === 0 ? 'sUnused' : null)
      });
    }

    if (flds) {
      const headers = ['Field', 'Source Tables', 'Tags', 'Usage', 'Used In'];
      const order = ['Key', 'Chart', 'Set analysis', 'Dimension', 'Measure', 'Variable'];
      const fmtUsedIn = name => {
        const s = usedInMap && usedInMap.get(name);
        if (!s || !s.size) return '';
        const arr = Array.from(s);
        arr.sort((a, b) => order.indexOf(a) - order.indexOf(b));
        return arr.join(', ');
      };

      const ff = sortAsc(FLDS, 'name');
      const matrix = ff.map(f => [
        f.name,
        f.srcTables,
        f.tags,
        unusedSet && unusedSet.has(f.name) ? 'UNUSED' : 'USED',
        fmtUsedIn(f.name)
      ]);

      addWorksheetXlsx(workbook, {
        name: 'Fields',
        headers,
        matrix,
        rowStyleId: r => (r?.[3] === 'UNUSED' ? 'sUnused' : null)
      });
    }

    if (shts) {
      const headers = ['ID', 'Sheet Title', 'Description', 'Owner'];
      const ss = sortAsc(SHTS, 'title');
      const matrix = ss.map(s => [s.id, s.title, s.description, s.owner]);
      addWorksheetXlsx(workbook, { name: 'Sheets', headers, matrix });
    }

    if (chrs) {
      const headers = ['Chart ID', 'Type', 'Extension?', 'Title', 'Sheet', 'Sheet ID', 'Master?', 'Master ID', 'Items'];
      const cc = sortAsc(CHRS, 'title');
      const matrix = cc.map(o => [o.objectId, o.type, o.isExtension, o.title, o.sheetTitle, o.sheetId, o.isMaster, o.masterId, o.itemsSummary || '']);
      const widths = [90, 120, 70, 160, 200, 220, 60, 200, 360];
      addWorksheetXlsx(workbook, {
        name: 'Charts',
        headers,
        matrix,
        widths,
        rowStyleId: () => 'sWrapTop',
        wrapCols: [9]
      });
    }

    if (vars) {
      const headers = ['Name', 'Definition', 'Comment', 'Tags', 'Script Variable?', 'Reserved?'];
      const vv = sortAsc(VARS, 'name');
      const matrix = vv.map(v => [v.name, v.definition, v.comment, v.tags, v.isScript, v.isReserved]);
      addWorksheetXlsx(workbook, {
        name: 'Variables',
        headers,
        matrix,
        wrapAll: true
      });
    }

    if (scrmeta) {
      if (scriptMetaRows && scriptMetaRows.length) {
        const headers = ['Tab', 'LOADs', 'STOREs', 'JOINs', 'RESIDENTs', 'QVDs', 'Variables'];
        const matrix = scriptMetaRows.map(r => [
          r.tab,
          Number(r.loads || 0),
          Number(r.stores || 0),
          Number(r.joins || 0),
          Number(r.residents || 0),
          (r.qvds || []).join(', '),
          (r.vars || []).join(', ')
        ]);
        const widths = [180, 70, 70, 70, 90, 320, 260];
        addWorksheetXlsx(workbook, {
          name: 'Script',
          headers,
          matrix,
          widths,
          rowStyleId: () => 'sWrapTop',
          wrapAll: true
        });
      } else if (scriptDenied) {
        const headers = ['Info'];
        const matrix = [[String('Script metadata not available for this session.')]];
        addWorksheetXlsx(workbook, {
          name: 'Script',
          headers,
          matrix,
          widths: [600],
          wrapAll: true
        });
      }
    }

    await downloadXlsx(fileName || 'Qollect_Metadata', workbook);
  }

  // ------- Export App Script (.qvs) -------
  async function exportAppScript(app, fileNameBase) {
    let text = null;
    try {
      const res = await app.getScript();
      text = (res && res.qScript) ? res.qScript : (typeof res === 'string' ? res : '');
    } catch (e) {
      console.error('getScript failed:', e);
      alert('Script export unavailable: your user session does not have permission to read the app load script. Use a development copy or ask an administrator to grant Edit + Data Load Editor access.');
      return;
    }

    if (!text) {
      alert('Script export unavailable for this app/user (insufficient permissions or published/embedded context). Use a development copy or request Edit + Data Load Editor access.');
      return;
    }

    const normalized = text.replace(/\r?\n/g, '\r\n');
    const base = (fileNameBase && fileNameBase.trim()) || 'Qollect_App_Script';
    downloadText(base, normalized);
  }

  // ------- Extension -------
  return {
    definition: {
      type: 'items',
      component: 'accordion',
      items: {
        settings: {
          uses: 'settings',
          items: {
            fileName: {
              type: 'string',
              label: 'Default file name',
              ref: 'props.fileName',
              defaultValue: 'Qollect_Metadata'
            },
            scriptFileName: {
              type: 'string',
              label: 'Script file name (base)',
              ref: 'props.scriptFileName',
              defaultValue: 'Qollect_App_Script'
            }
          }
        },
        about: {
          label: 'About',
          type: 'items',
          items: {
            aboutTitle: { component: 'text', label: 'Qollect' },
            aboutVer: { component: 'text', label: 'Version: 1.4.1' },
            aboutAuth: { component: 'text', label: 'Author: Eli Gohar' },
            supportHdr: { component: 'text', label: 'Support development (Ko-fi):' },
            supportLnk: { component: 'link', label: 'ko-fi.com/eligohar', url: 'https://ko-fi.com/eligohar' }
          }
        }
      }
    },

    paint: function ($element, layout) {
      const app = qlik.currApp(this);
      const id = layout.qInfo.qId;
      const fileName = layout?.props?.fileName || 'Qollect_Metadata';
      const scriptFileName = layout?.props?.scriptFileName || 'Qollect_App_Script';

      $element.html(`
        <div class="qollect__wrap">
          <div class="qollect__card" role="group" aria-labelledby="qollect-title-${id}">
            <h3 id="qollect-title-${id}" class="qollect__title">Qollect - export app metadata</h3>
            <ul class="qollect__list" aria-label="Metadata types">
              <li class="qollect__list-item"><label class="qollect__item"><input id="ovw-${id}" type="checkbox" class="qollect__chk" checked><span class="qollect__label">App Overview</span></label></li>
              <li class="qollect__list-item"><label class="qollect__item"><input id="dims-${id}" type="checkbox" class="qollect__chk" checked><span class="qollect__label">Dimensions</span></label></li>
              <li class="qollect__list-item"><label class="qollect__item"><input id="msrs-${id}" type="checkbox" class="qollect__chk" checked><span class="qollect__label">Measures</span></label></li>
              <li class="qollect__list-item"><label class="qollect__item"><input id="vars-${id}" type="checkbox" class="qollect__chk" checked><span class="qollect__label">Variables</span></label></li>
              <li class="qollect__list-item"><label class="qollect__item"><input id="flds-${id}" type="checkbox" class="qollect__chk" checked><span class="qollect__label">Fields</span></label></li>
              <li class="qollect__list-item"><label class="qollect__item"><input id="shts-${id}" type="checkbox" class="qollect__chk" checked><span class="qollect__label">Sheets</span></label></li>
              <li class="qollect__list-item"><label class="qollect__item"><input id="chrs-${id}" type="checkbox" class="qollect__chk" checked><span class="qollect__label">Charts</span></label></li>
              <li class="qollect__list-item"><label class="qollect__item"><input id="scrmeta-${id}" type="checkbox" class="qollect__chk" checked><span class="qollect__label">Script</span></label></li>
            </ul>

            <div class="qollect__actions">
              <button id="btn-xlsx-${id}" class="qollect__btn" type="button">Export Selected</button>
              <button id="btn-script-${id}" class="qollect__btn qollect__btn--secondary" type="button" title="Download the app's load script as a .qvs file">Export App Script (.qvs)</button>
            </div>
          </div>
        </div>
      `);

      const $btnXlsx = $element.find(`#btn-xlsx-${id}`);
      const $btnScript = $element.find(`#btn-script-${id}`);

      function readSelection() {
        return {
          overview: $element.find(`#ovw-${id}`).is(':checked'),
          dims: $element.find(`#dims-${id}`).is(':checked'),
          msrs: $element.find(`#msrs-${id}`).is(':checked'),
          vars: $element.find(`#vars-${id}`).is(':checked'),
          flds: $element.find(`#flds-${id}`).is(':checked'),
          shts: $element.find(`#shts-${id}`).is(':checked'),
          chrs: $element.find(`#chrs-${id}`).is(':checked'),
          scrmeta: $element.find(`#scrmeta-${id}`).is(':checked')
        };
      }

      $btnXlsx.off('click').on('click', async () => {
        const sel = readSelection();
        $btnXlsx.prop('disabled', true).text('Exporting…');
        try {
          await exportSelectedXlsx(app, fileName, sel);
        } catch (err) {
          console.error(err);
          alert('XLSX export failed: ' + (err?.message || err));
        } finally {
          $btnXlsx.prop('disabled', false).text('Export Selected');
        }
      });

      $btnScript.off('click').on('click', async () => {
        $btnScript.prop('disabled', true).text('Fetching script…');
        try {
          await exportAppScript(app, scriptFileName);
        } catch (err) {
          console.error(err);
          alert('Script export failed: ' + (err?.message || err));
        } finally {
          $btnScript.prop('disabled', false).text('Export App Script (.qvs)');
        }
      });

      return qlik.Promise.resolve();
    }
  };
});
