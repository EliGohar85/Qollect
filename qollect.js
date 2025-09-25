define(['qlik', 'jquery', 'text!./qollect.css'], function (qlik, $, cssContent) {
  'use strict';

  // ------- load external CSS (no inline styles) -------
  (function ensureStyle() {
    if (!document.getElementById('qollect-style')) {
      const s = document.createElement('style');
      s.id = 'qollect-style';
      s.textContent = cssContent;
      document.head.appendChild(s);
    }
  })();

  // ------- SpreadsheetML helpers -------
  const xmlEscape = s => String(s ?? '')
    .replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')
    .replace(/"/g,'&quot;').replace(/'/g,'&apos;');

  const cell = v => `<Cell><Data ss:Type="${typeof v==='number'?'Number':'String'}">${xmlEscape(v)}</Data></Cell>`;
  const row  = cells => `<Row>${cells.join('')}</Row>`;
  const rowStyled = (cells, styleId) => `<Row${styleId ? ` ss:StyleID="${styleId}"` : ''}>${cells.join('')}</Row>`;
  const xlHeaderRow = headers =>
    `<Row>${headers.map(h=>`<Cell ss:StyleID="sHeader"><Data ss:Type="String">${xmlEscape(h)}</Data></Cell>`).join('')}</Row>`;

  const AUTO_FILTERS = true, FREEZE_HEADER = true, AUTO_WIDTHS = true;

  function estimateColumnWidths(headers, matrix){
    const PX_PER_CHAR = 7, PADDING = 16, MIN_W = 80, MAX_W = 600;
    const cols = headers.length;
    const maxLens = Array.from({length: cols}, (_, i) => String(headers[i] ?? '').length);
    for (const r of matrix) for (let c=0;c<cols;c++){
      const len = String(r[c] ?? '').length;
      if (len > maxLens[c]) maxLens[c] = len;
    }
    return maxLens.map(len => Math.max(MIN_W, Math.min(MAX_W, len * PX_PER_CHAR + PADDING)));
  }

  function worksheetWithFeatures(name, headers, matrix, opts = {}){
    const colsCount = headers.length, rowsCount = matrix.length + 1;

    let columnsXml = '';
    if (AUTO_WIDTHS) {
      const widths = estimateColumnWidths(headers, matrix);
      columnsXml = widths.map((w, idx) => `<Column ss:Index="${idx+1}" ss:Width="${w}"/>`).join('');
    }

    const headerXml = xlHeaderRow(headers);
    const dataXml = matrix.map(r => {
      const styledId = typeof opts.rowStyleId === 'function' ? (opts.rowStyleId(r) || null) : null;
      const cells = r.map(cell);
      return styledId ? rowStyled(cells, styledId) : row(cells);
    }).join('');

    const filterXml = AUTO_FILTERS ? `<AutoFilter x:Range="R1C1:R${rowsCount}C${colsCount}" xmlns:x="urn:schemas-microsoft-com:office:excel"/>` : '';
    const freezeXml = FREEZE_HEADER ? `<WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel">
      <FreezePanes/><FrozenNoSplit/><SplitHorizontal>1</SplitHorizontal>
      <TopRowBottomPane>1</TopRowBottomPane><ActivePane>2</ActivePane>
    </WorksheetOptions>` : '';

    return `<Worksheet ss:Name="${xmlEscape(name)}"><Table>${columnsXml}${headerXml}${dataXml}</Table>${filterXml}${freezeXml}</Worksheet>`;
  }

  const xlWorkbook = body => `<?xml version="1.0"?>
<?mso-application progid="Excel.Sheet"?>
<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:o="urn:schemas-microsoft-com:office:office"
 xmlns:x="urn:schemas-microsoft-com:office:excel"
 xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:html="http://www.w3.org/TR/REC-html40">
 <Styles>
  <Style ss:ID="sHeader"><Font ss:Bold="1"/><Interior ss:Color="#D9E1F2" ss:Pattern="Solid"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/></Borders></Style>
  <Style ss:ID="sUnused"><Interior ss:Color="#FCE4E4" ss:Pattern="Solid"/></Style>
 </Styles>
 ${body}
</Workbook>`;

  function downloadXml(filename, xmlString){
    try{
      const blob = new Blob([xmlString], { type: 'application/vnd.ms-excel' });
      const url = (window.URL||window.webkitURL).createObjectURL(blob);
      const a = document.createElement('a'); a.href=url;
      a.download = filename.endsWith('.xls') ? filename : `${filename}.xls`;
      document.body.appendChild(a); a.click();
      setTimeout(()=>{ document.body.removeChild(a); (window.URL||window.webkitURL).revokeObjectURL(url); },0);
    }catch{
      const data = 'data:application/vnd.ms-excel;charset=utf-8,' + encodeURIComponent(xmlString);
      const a = document.createElement('a'); a.href=data;
      a.download = filename.endsWith('.xls') ? filename : `${filename}.xls`;
      document.body.appendChild(a); a.click(); setTimeout(()=>a.remove(),0);
    }
  }

  // ------- Engine helpers -------
  function getDoc(app){
    const doc = app && app.model && app.model.enigmaModel;
    if (!doc) throw new Error('Engine document not available (app.model.enigmaModel missing).');
    return doc;
  }
  const YN = v => (v === true ? 'Y' : v === false ? 'N' : '');
  const sortAsc = (arr, key) => arr.slice().sort((a,b)=> String(a?.[key] ?? '').localeCompare(String(b?.[key] ?? ''), undefined, {sensitivity:'base'}));

  // ===================== field parsing helpers =====================
  const reInlineDollar = /\$\(\s*=\s*([^)]*?)\s*\)/g;
  const reVarDollar = /\$\(\s*([A-Za-z_]\w*)(?:\([^\)]*\))?\s*\)/g;
  const reBlockComments = new RegExp('/' + '\\*' + '[\\s\\S]*?' + '\\*' + '/', 'g');
  const reHasSet = /\{\s*<[\s\S]*?>\s*\}/; // detect set analysis in an expression

  function buildVariableMap(vars){ const m=new Map(); for(const v of vars||[]) if(v?.name) m.set(v.name, v.definition ?? ''); return m; }

  function expandDollar(expr, varMap, depth = 0, seen = new Set()){
    if (!expr || typeof expr !== 'string') return '';
    if (depth > 5) return expr;
    expr = expr.replace(reInlineDollar, (_, inner) => inner || '');
    expr = expr.replace(reVarDollar, (_, vname) => {
      if (!varMap || !varMap.has(vname)) return '';
      if (seen.has(vname)) return '';
      seen.add(vname);
      return expandDollar(varMap.get(vname) || '', varMap, depth+1, seen);
    });
    return expr;
  }

  function extractFieldsFromExpr(expr, allFieldsSet){
    const used = new Set();
    if (!expr || typeof expr !== 'string') return used;

    let s = expr.replace(/\/\/.*$/gm, '').replace(reBlockComments, '');

    // [Field]
    const br = s.match(/\[([^\]\\]|\\.)+\]/g) || [];
    for (const tok of br) {
      const name = tok.slice(1,-1);
      const base = name.replace(/\.autoCalendar\..*$/,'');
      if (allFieldsSet.has(name)) used.add(name);
      if (allFieldsSet.has(base)) used.add(base);
    }

    // {<Field=...>} (set analysis)
    const sa = s.match(/\{\s*<([\s\S]*?)>\s*\}/g) || [];
    for (const block of sa) {
      const inside = block.slice(block.indexOf('<')+1, block.lastIndexOf('>'));
      const parts = inside.split(/,(?=(?:[^'"]|'[^']*'|"[^"]*")*$)/g);
      for (const part of parts) {
        const m = part.match(/(^|[,<\s])(?:(?:\w+::)?)(\[?[^\]=,]+?\]?)(?==)/);
        if (m) {
          let lhs = m[2] || '';
          lhs = lhs.replace(/^\[|\]$/g,'');
          const base = lhs.replace(/\.autoCalendar\..*$/,'');
          if (allFieldsSet.has(lhs)) used.add(lhs);
          if (allFieldsSet.has(base)) used.add(base);
        }
      }
    }

    // Field.autoCalendar.*
    const ac = s.match(/([A-Za-z_][\w ]*)\.autoCalendar\.[A-Za-z]+/g) || [];
    for (const t of ac) {
      const base = t.replace(/\.autoCalendar\..*$/,'').trim();
      if (allFieldsSet.has(base)) used.add(base);
    }

    // conservative unbracketed
    for (const fname of allFieldsSet) {
      if (!/^[A-Za-z_]\w*$/.test(fname)) continue;
      const re = new RegExp(`(?<![\\w$])${fname}(?![\\w$])`);
      if (re.test(s)) used.add(fname);
    }

    return used;
  }

  function collectObjectExpressions(props){
    const out = [];
    const seen = new WeakSet(); const MAX_DEPTH = 50;

    function pushExpr(val, path, force=false){
      if (typeof val === 'string' && val.trim()) {
        if (force || /[=\[\]{}$]/.test(val) || /autoCalendar/.test(val) || /[A-Za-z_]\w*\s*\(/.test(val)) {
          out.push({ expr: val, path });
        }
      }
    }
    function pushField(val, path){
      if (typeof val === 'string' && val.trim()) {
        out.push({ expr: /^\[.*\]$/.test(val) ? val : `[${val}]`, path });
      }
    }

    function walk(node, path, depth=0){
      if (!node || depth > MAX_DEPTH) return;
      if (typeof node === 'string') { pushExpr(node, path); return; }
      if (Array.isArray(node)) {
        for (let i=0;i<node.length;i++){
          const v=node[i];
          if (typeof v === 'string') pushExpr(v, `${path}[${i}]`);
          else if (v && typeof v === 'object' && !seen.has(v)) { seen.add(v); walk(v, `${path}[${i}]`, depth+1); }
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
          fdefs.forEach((f,idx)=> pushField(f, `${path}/qListObjectDef/qDef/qFieldDefs[${idx}]`));
          pushExpr(lo?.qDef?.qLabelExpression, `${path}/qListObjectDef/qDef/qLabelExpression`);
          pushExpr(lo?.qCalcCond, `${path}/qListObjectDef/qCalcCond`);
        }
        if (node.qHyperCubeDef) {
          const hc = node.qHyperCubeDef;
          (hc.qDimensions||[]).forEach((d,i)=>{
            const fdefs = Array.isArray(d?.qDef?.qFieldDefs) && d.qDef.qFieldDefs.length ? d.qDef.qFieldDefs
                         : (d?.qDef?.qFieldDef ? [d.qDef.qFieldDef] : []);
            fdefs.forEach((f,idx)=> pushField(f, `${path}/qHyperCubeDef/qDimensions[${i}]/qDef/qFieldDefs[${idx}]`));
            pushExpr(d?.qDef?.qLabelExpression, `${path}/qHyperCubeDef/qDimensions[${i}]/qDef/qLabelExpression`);
            pushExpr(d?.qCalcCond, `${path}/qHyperCubeDef/qDimensions[${i}]/qCalcCond`);
          });
          (hc.qMeasures||[]).forEach((m,i)=>{
            pushExpr(m?.qDef?.qDef, `${path}/qHyperCubeDef/qMeasures[${i}]/qDef/qDef`, true);
            pushExpr(m?.qDef?.qLabelExpression, `${path}/qHyperCubeDef/qMeasures[${i}]/qDef/qLabelExpression`);
            pushExpr(m?.qSortByExpression?.qExpression, `${path}/qHyperCubeDef/qMeasures[${i}]/qSortByExpression/qExpression`, true);
          });
          pushExpr(hc.qCalcCond, `${path}/qHyperCubeDef/qCalcCond`);
          pushExpr(hc.qSuppressZero, `${path}/qHyperCubeDef/qSuppressZero`);
          pushExpr(hc.qSuppressMissing, `${path}/qHyperCubeDef/qSuppressMissing`);
        }

        pushExpr(node.title, `${path}/title`);
        pushExpr(node.subtitle, `${path}/subtitle`);
        pushExpr(node.footnote, `${path}/footnote`);
        pushExpr(node.qTitleExpression, `${path}/qTitleExpression`);
        pushExpr(node.qSubtitleExpression, `${path}/qSubtitleExpression`);
        pushExpr(node.qFootnoteExpression, `${path}/qFootnoteExpression`);
        pushExpr(node.showCondition, `${path}/showCondition`);
        pushExpr(node.colorExpression, `${path}/colorExpression`);
        pushExpr(node.sortByExpression?.qExpression, `${path}/sortByExpression/qExpression`, true);

        for (const k of Object.keys(node)) {
          const v=node[k];
          if (typeof v === 'string') pushExpr(v, `${path}/${k}`);
          else if (v && typeof v === 'object' && !seen.has(v)) { seen.add(v); walk(v, `${path}/${k}`, depth+1); }
        }
      }
    }

    walk(props,'object');
    return out;
  }

  // ---- robust object discovery (getChildInfos + recursion) ----
  async function fetchSheetsViaEngine(app){
    const doc = getDoc(app);
    const listHandle = await doc.createSessionObject({
      qInfo: { qType: 'SheetList' },
      qAppObjectListDef: { qType: 'sheet', qData: { rank: '/rank' } }
    });
    const layout = await listHandle.getLayout();
    const items = layout?.qAppObjectList?.qItems || [];
    return items.map(it => ({
      id: it.qInfo?.qId || '', title: it.qMeta?.title || '', description: it.qMeta?.description || '',
      owner: (it.qMeta && it.qMeta.owner && (it.qMeta.owner.name || it.qMeta.owner.userId)) || ''
    }));
  }

  async function fetchObjectPropsForSheets(app){
    const objs = [];
    const sheets = await fetchSheetsViaEngine(app);
    const visited = new Set();

    async function addObjectById(objId, sheet){
      if (!objId || visited.has(objId)) return;
      visited.add(objId);
      try {
        const objModel = await app.getObject(objId);
        const p = await objModel.getProperties();
        objs.push({ sheetTitle: sheet.title, sheetId: sheet.id, objectId: objId, props: p });
        let kids = [];
        try { kids = await objModel.getChildInfos(); } catch {}
        if (Array.isArray(kids) && kids.length) {
          for (const ch of kids) await addObjectById(ch.qId || ch.id || ch.Id, sheet);
        } else {
          const cells = Array.isArray(p?.cells) ? p.cells : [];
          for (const c of cells) await addObjectById(c.name || c.qId || c.id, sheet);
        }
      } catch(e){}
    }

    for (const sh of sheets) {
      try {
        const sheetModel = await app.getObject(sh.id);
        let childInfos = [];
        try { childInfos = await sheetModel.getChildInfos(); } catch {}
        if (Array.isArray(childInfos) && childInfos.length) {
          for (const ch of childInfos) await addObjectById(ch.qId || ch.id || ch.Id, sh);
        } else {
          const props = await sheetModel.getProperties();
          const cells = Array.isArray(props?.cells) ? props.cells : [];
          for (const c of cells) await addObjectById(c.name || c.qId || c.id, sh);
        }
      } catch(e){}
    }
    return objs;
  }

  // ===================== MASTER ITEM USAGE (primary + alternates, slot-unique) =====================
  async function computeMasterUsage(app){
    const objs = await fetchObjectPropsForSheets(app);

    const dimSlots = new Map(); // id -> Set(slotKey)
    const msrSlots = new Map();

    const addSlot = (map, id, slotKey) => {
      if (!id || !slotKey) return;
      let s = map.get(id);
      if (!s) { s = new Set(); map.set(id, s); }
      s.add(slotKey);
    };

    function scanPropsForLibraryIds(props, objId){
      function walk(node, path, lastDimIdx = null, lastMsrIdx = null){
        if (!node) return;
        if (Array.isArray(node)) {
          node.forEach((v,i)=>{
            if (/qDimensions$/.test(path)) walk(v, `${path}[${i}]`, i, lastMsrIdx);
            else if (/qMeasures$/.test(path)) walk(v, `${path}[${i}]`, lastDimIdx, i);
            else walk(v, `${path}[${i}]`, lastDimIdx, lastMsrIdx);
          });
          return;
        }
        if (typeof node === 'object') {
          for (const [k,v] of Object.entries(node)) {
            const p = `${path}/${k}`;

            if (k === 'qLibraryId' && typeof v === 'string' && v.trim()) {
              const isDim = /\/qDimensions(\[|\/)/.test(path) || /qListObjectDef/.test(path);
              const isMsr = /\/qMeasures(\[|\/)/.test(path);
              let slot = null;

              const mD = path.match(/qDimensions\[\d+\]/g);
              const mM = path.match(/qMeasures\[\d+\]/g);
              if (mD && mD.length) slot = mD[mD.length-1];
              else if (mM && mM.length) slot = mM[mM.length-1];
              else if (path.includes('qListObjectDef')) slot = 'qListObjectDef';

              const slotKey = `${objId}:${slot || p}`;

              if (isMsr) addSlot(msrSlots, v, slotKey);
              else if (isDim) addSlot(dimSlots, v, slotKey);
              else {
                if (lastMsrIdx !== null) addSlot(msrSlots, v, `${objId}:qMeasures[${lastMsrIdx}]`);
                else addSlot(dimSlots, v, `${objId}:qDimensions[${lastDimIdx ?? '0'}]`);
              }
            }

            if (v && (typeof v === 'object' || Array.isArray(v))) {
              walk(v, p, lastDimIdx, lastMsrIdx);
            }
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
          const mp = await masterModel.getProperties();
          scanPropsForLibraryIds(mp, o.objectId);
        } catch(e){}
      }
    }

    const dimCount = new Map();
    const msrCount = new Map();
    for (const [id, set] of dimSlots) dimCount.set(id, set.size);
    for (const [id, set] of msrSlots) msrCount.set(id, set.size);
    return { dimCount, msrCount };
  }

  // ===================== FETCHERS =====================
  async function fetchDimensionsViaEngine(app){
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
      const id = it.qInfo?.qId; if (!id) continue;
      let props = null;
      try {
        if (typeof app.getDimension === 'function') { const dimModel = await app.getDimension(id); props = await dimModel.getProperties(); }
        else if (typeof getDoc(app).getDimension === 'function') { const dimHandle = await getDoc(app).getDimension({ qId: id }); props = await dimHandle.getProperties(); }
      } catch(e){}
      if (props) {
        const meta = props.qMetaDef || {}, qDim = props.qDim || {};
        const fieldsArray = Array.isArray(qDim.qFieldDefs) ? qDim.qFieldDefs : [];
        const drillDownFieldsArray = Array.isArray(qDim.qDrillDownFieldDefs) ? qDim.qDrillDownFieldDefs : [];
        results.push({
          id, title: meta.title || it.qMeta?.title || '', description: meta.description || it.qMeta?.description || '',
          tags: (meta.tags || it.qMeta?.tags || []).join(', '),
          fields: fieldsArray.join(', '), fieldsArray, drillDownFieldsArray,
          labelExpr: qDim.qLabelExpression || '', usedCount: 0
        });
      } else {
        results.push({ id, title: it.qMeta?.title || '', description: it.qMeta?.description || '', tags: (it.qMeta?.tags || []).join(', '), fields: '', fieldsArray: [], drillDownFieldsArray: [], labelExpr: '', usedCount: 0 });
      }
    }
    return results;
  }

  async function fetchMeasuresViaEngine(app){
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
      const id = it.qInfo?.qId; if (!id) continue;
      let props = null;
      try {
        if (typeof app.getMeasure === 'function') { const msrModel = await app.getMeasure(id); props = await msrModel.getProperties(); }
        else if (typeof getDoc(app).getMeasure === 'function') { const msrHandle = await getDoc(app).getMeasure({ qId: id }); props = await msrHandle.getProperties(); }
      } catch(e){}
      if (props) {
        const meta = props.qMetaDef || {}, qMeasure = props.qMeasure || {};
        results.push({
          id, title: meta.title || it.qMeta?.title || '', description: meta.description || it.qMeta?.description || '',
          tags: (meta.tags || it.qMeta?.tags || []).join(', '),
          expression: qMeasure.qDef || '', label: qMeasure.qLabel || '', labelExpr: qMeasure.qLabelExpression || '', usedCount: 0
        });
      } else {
        results.push({ id, title: it.qMeta?.title || '', description: it.qMeta?.description || '', tags: (it.qMeta?.tags || []).join(', '), expression: '', label: '', labelExpr: '', usedCount: 0 });
      }
    }
    return results;
  }

  async function fetchFieldsViaEngine(app){
    const doc = getDoc(app);
    const listHandle = await doc.createSessionObject({
      qInfo: { qType: 'FieldList' },
      qFieldListDef: { qShowSystem: false, qShowHidden: false, qShowSemantic: true, qShowDerivedFields: true, qShowImplicit: true, qShowSrcTables: true }
    });
    const layout = await listHandle.getLayout();
    const items = layout?.qFieldList?.qItems || [];
    return items.map(it => ({
      name: it.qName || '', tags: (it.qTags || []).join(', '),
      srcTables: Array.isArray(it.qSrcTables) ? it.qSrcTables.join(', ') : (it.qSrcTables || '')
    }));
  }

  // ===================== UNUSED FIELDS + "USED IN" FINDER =====================
  async function findUnusedFields(app, allFields, dims, msrs, vars){
    const allFieldsSet = new Set((allFields||[]).map(f=>f.name).filter(Boolean));
    const varMap = buildVariableMap(vars);
    const used = new Set();
    const usedInMap = new Map(); // field -> Set(categories)

    const addUse = (fname, cat) => {
      if (!fname) return;
      used.add(fname);
      let s = usedInMap.get(fname);
      if (!s) { s = new Set(); usedInMap.set(fname, s); }
      s.add(cat);
    };
    const markExpr = (expr, cat) => {
      const s = extractFieldsFromExpr(expr, allFieldsSet);
      const hasSet = reHasSet.test(expr || '');
      s.forEach(f => { addUse(f, cat); if (hasSet) addUse(f, 'Set analysis'); });
    };

    // Master dimensions (single + drill-down + labels)
    for (const d of dims || []) {
      const rawList = Array.isArray(d.fieldsArray) ? d.fieldsArray : (typeof d.fields === 'string' ? d.fields.split(',') : []);
      const ddList  = Array.isArray(d.drillDownFieldsArray) ? d.drillDownFieldsArray : [];
      const allDimDefs = [...rawList, ...ddList].map(s=>String(s||'').trim()).filter(Boolean);

      for (const def of allDimDefs) {
        if (def.startsWith('=')) {
          const ex = expandDollar(def, varMap);
          markExpr(ex, 'Dimension');
        } else {
          const base = def.replace(/^\[|\]$/g,'').replace(/\.autoCalendar\..*$/,'');
          if (allFieldsSet.has(base)) addUse(base, 'Dimension');
          const ex = /^\[.*\]$/.test(def) ? def : `[${def}]`;
          markExpr(ex, 'Dimension');
        }
      }
      if (d.labelExpr) {
        const ex = expandDollar(d.labelExpr, varMap);
        markExpr(ex, 'Dimension');
      }
    }

    // Master measures
    for (const m of msrs || []) {
      if (m.expression) {
        const ex = expandDollar(m.expression, varMap);
        markExpr(ex, 'Measure');
      }
      if (m.labelExpr) {
        const exl = expandDollar(m.labelExpr, varMap);
        markExpr(exl, 'Measure');
      }
    }

    // Variables
    for (const v of vars || []) {
      if (!v?.definition) continue;
      const ex = expandDollar(v.definition, varMap);
      markExpr(ex, 'Variable');
    }

    // Objects on sheets (incl. listboxes / alternates)
    const objs = await fetchObjectPropsForSheets(app);
    for (const o of objs) {
      const items = collectObjectExpressions(o.props || {});
      for (const it of items) {
        const ex = expandDollar(it.expr, varMap);
        markExpr(ex, 'Chart');
      }
    }

    // Map autoCalendar children -> base
    for (const f of allFieldsSet) {
      if (f.includes('.autoCalendar.')) {
        const base = f.replace(/\.autoCalendar\..*$/,'');
        if (used.has(f)) addUse(base, 'Chart'); // conservative attribution
      }
    }

    const unused = new Set();
    for (const f of allFieldsSet) {
      if (f.includes('.autoCalendar.')) continue;
      if (!used.has(f)) unused.add(f);
    }
    return { used, unused, usedInMap };
  }

  // ------- Sheet builders -------
  const buildOverviewSheet = (info) => {
    const headers = [
      'Application name','Application ID',
      '# of Dimensions','# of Measures','# of Fields',
      '# of Sheets','# of Charts','# of Variables'
    ];
    const matrix = [[
      info.name || '', info.id || '',
      Number(info.dims||0), Number(info.msrs||0), Number(info.flds||0),
      Number(info.shts||0), Number(info.chrs||0), Number(info.vars||0)
    ]];
    return worksheetWithFeatures('App Overview', headers, matrix);
  };

  const buildDimSheet = dims => {
    const headers = ['ID','Title','Fields','Label Expression','Description','Tags','Used Count'];
    const matrix = dims.map(d => [d.id, d.title, d.fields, d.labelExpr, d.description, d.tags, Number(d.usedCount||0)]);
    return worksheetWithFeatures('Dimensions', headers, matrix, { rowStyleId: r => (Number(r?.[6])===0 ? 'sUnused' : null) });
  };
  const buildMsrSheet = msrs => {
    const headers = ['ID','Title','Expression','Label','Label Expression','Description','Tags','Used Count'];
    const matrix = msrs.map(m => [m.id, m.title, m.expression, m.label, m.labelExpr, m.description, m.tags, Number(m.usedCount||0)]);
    return worksheetWithFeatures('Measures', headers, matrix, { rowStyleId: r => (Number(r?.[7])===0 ? 'sUnused' : null) });
  };
  const buildFldSheet = (flds, unusedSet, usedInMap) => {
    const headers = ['Field','Source Tables','Tags','Usage','Used In'];
    const order = ['Chart','Set analysis','Dimension','Measure','Variable'];
    const fmtUsedIn = name => {
      const s = usedInMap && usedInMap.get(name);
      if (!s || !s.size) return '';
      const arr = Array.from(s);
      arr.sort((a,b)=> order.indexOf(a) - order.indexOf(b));
      return arr.join(', ');
    };
    const matrix = flds.map(f => [
      f.name, f.srcTables, f.tags,
      unusedSet && unusedSet.has(f.name) ? 'UNUSED' : 'USED',
      fmtUsedIn(f.name)
    ]);
    return worksheetWithFeatures('Fields', headers, matrix, { rowStyleId: r => (r?.[3]==='UNUSED' ? 'sUnused' : null) });
  };
  const buildShtSheet = shts => {
    const headers = ['ID','Sheet Title','Description','Owner'];
    const matrix = shts.map(s => [s.id, s.title, s.description, s.owner]);
    return worksheetWithFeatures('Sheets', headers, matrix);
  };
  const buildChrSheet = chrs => {
    const headers = ['Chart ID','Type','Title','Sheet','Sheet ID','Master?','Master ID'];
    const matrix = chrs.map(o => [o.objectId, o.type, o.title, o.sheetTitle, o.sheetId, o.isMaster, o.masterId]);
    return worksheetWithFeatures('Charts', headers, matrix);
  };
  const buildVarSheet = vars => {
    const headers = ['Name','Definition','Comment','Tags','Script Variable?','Reserved?'];
    const matrix = vars.map(v => [v.name, v.definition, v.comment, v.tags, v.isScript, v.isReserved]);
    return worksheetWithFeatures('Variables', headers, matrix);
  };

  // ------- Charts (for Charts sheet) -------
  async function fetchChartsViaEngine(app){
    const sheets = await fetchSheetsViaEngine(app);
    const charts = [];
    for (const sh of sheets) {
      try {
        const sheetModel = await app.getObject(sh.id);
        let childInfos = [];
        try { childInfos = await sheetModel.getChildInfos(); } catch {}
        const props = childInfos?.length ? null : await sheetModel.getProperties();
        const cells = childInfos?.length ? childInfos.map(ci => ({ name: ci.qId })) : (props?.cells || []);
        for (const c of cells) {
          const objId = c.name;
          try {
            const objModel = await app.getObject(objId);
            const p = await objModel.getProperties();
            charts.push({
              sheetTitle: sh.title, sheetId: sh.id, objectId: objId,
              type: p.visualization || c.type || '',
              title: (typeof p.title === 'string' ? p.title : (p.title && p.title.qStringExpression) || ''),
              isMaster: p.qExtendsId ? 'Y' : 'N', masterId: p.qExtendsId || ''
            });
          } catch(e){}
        }
      } catch(e){}
    }
    return charts;
  }

  // ------- Variables -------
  async function fetchVariablesViaEngine(app){
    const doc = getDoc(app);
    let items = [];
    try {
      const h = await doc.createSessionObject({
        qInfo: { qType: 'VariableList' },
        qVariableListDef: {
          qType: 'variable', qShowReserved: true, qShowConfig: true,
          qData: { tags: '/qMetaDef/tags', definition: '/qDefinition', comment: '/qComment' }
        }
      });
      const layout = await h.getLayout();
      items = layout?.qVariableList?.qItems || [];
    } catch(e){}
    if (!items.length) {
      items = await new Promise(resolve => {
        try {
          app.getList('VariableList', function(reply){
            const arr = reply?.qVariableList?.qItems || [];
            if (reply?.qInfo?.qId) { app.destroySessionObject(reply.qInfo.qId); }
            resolve(arr);
          });
        } catch(e){ resolve([]); }
      });
    }
    if (!items.length) return [];
    const vars = [];
    for (const it of items) {
      const name = it.qName || it.qInfo?.qName || ''; if (!name) continue;
      let props = null;
      try {
        if (app.variable && typeof app.variable.getByName === 'function') {
          const vm = await app.variable.getByName(name);
          props = await vm.getProperties();
        } else if (typeof getDoc(app).getVariableByName === 'function') {
          let vh = null;
          try { vh = await getDoc(app).getVariableByName(name); } catch(e1) {
            try { vh = await getDoc(app).getVariableByName({ qName: name }); } catch(e2){}
          }
          if (vh && typeof vh.getProperties === 'function') props = await vh.getProperties();
        }
      } catch(e){}
      const definition = props?.qDefinition ?? it.qDefinition ?? it.qData?.definition ?? '';
      const comment    = props?.qComment    ?? it.qComment    ?? it.qData?.comment    ?? '';
      const tagsArr    = props?.qMetaDef?.tags ?? it.qTags ?? it.qData?.tags ?? [];
      const tags       = Array.isArray(tagsArr) ? tagsArr.join(', ') : (tagsArr || '');
      const isScript   = props?.qIsScriptCreated ?? it.qIsScriptCreated ?? it.qIsScript;
      const isReserved = props?.qIsReserved ?? it.qIsReserved;
      vars.push({ name, definition, comment, tags, isScript: YN(isScript), isReserved: YN(isReserved) });
    }
    return vars;
  }

  // ------- Export orchestrator -------
  function getAppIdSafe(app){
    return app?.id
      || app?.model?.enigmaModel?.id
      || app?.model?.enigmaModel?.appId
      || app?.model?.id
      || '';
  }

  async function exportSelected(app, fileName, opts){
    const { overview, dims, msrs, flds, shts, chrs, vars } = opts;
    if (!overview && !dims && !msrs && !flds && !shts && !chrs && !vars) { alert('Nothing selected to export.'); return; }

    // decide what we need to fetch (avoid duplicates)
    const needDims = overview || dims || flds;
    const needMsrs = overview || msrs || flds;
    const needVars = overview || vars || flds;
    const needFlds = overview || flds;
    const needShts = overview || shts;
    const needChrs = overview || chrs;

    const [DIMS, MSRS, VARS, FLDS, SHTS, CHRS] = await Promise.all([
      needDims ? fetchDimensionsViaEngine(app) : Promise.resolve([]),
      needMsrs ? fetchMeasuresViaEngine(app)   : Promise.resolve([]),
      needVars ? fetchVariablesViaEngine(app)  : Promise.resolve([]),
      needFlds ? fetchFieldsViaEngine(app)     : Promise.resolve([]),
      needShts ? fetchSheetsViaEngine(app)     : Promise.resolve([]),
      needChrs ? fetchChartsViaEngine(app)     : Promise.resolve([]),
    ]);

    // Only compute master usage when needed (dims/measures sheets shown)
    if (dims || msrs) {
      const { dimCount, msrCount } = await computeMasterUsage(app);
      for (const d of DIMS) d.usedCount = dimCount.get(d.id) || 0;
      for (const m of MSRS) m.usedCount = msrCount.get(m.id) || 0;
    }

    let unusedSet = null, usedInMap = null;
    if (flds) {
      const res = await findUnusedFields(app, FLDS, DIMS, MSRS, VARS);
      unusedSet = res.unused;
      usedInMap = res.usedInMap;
    }

    const sheets = [];

    // App Overview first (if requested)
    if (overview) {
      let appName = '', appId = getAppIdSafe(app);
      try {
        const layout = await getDoc(app).getAppLayout();
        appName = layout?.qTitle || '';
      } catch(e){}
      sheets.push(buildOverviewSheet({
        name: appName,
        id: appId,
        dims: DIMS.length,
        msrs: MSRS.length,
        flds: FLDS.length,
        shts: SHTS.length,
        chrs: CHRS.length,
        vars: VARS.length
      }));
    }

    if (dims) sheets.push(buildDimSheet(sortAsc(DIMS, 'title')));
    if (msrs) sheets.push(buildMsrSheet(sortAsc(MSRS, 'title')));
    if (flds) sheets.push(buildFldSheet(sortAsc(FLDS, 'name'), unusedSet, usedInMap));
    if (shts) sheets.push(buildShtSheet(sortAsc(SHTS, 'title')));
    if (chrs) sheets.push(buildChrSheet(sortAsc(CHRS, 'title')));
    if (vars) sheets.push(buildVarSheet(sortAsc(VARS, 'name')));

    downloadXml(fileName || 'Qollect_Metadata', xlWorkbook(sheets.join('')));
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
            fileName: { type: 'string', label: 'Default file name', ref: 'props.fileName', defaultValue: 'Qollect_Metadata' }
          }
        },
        about: {
          label: 'About',
          type: 'items',
          items: {
            aboutTitle: { component: 'text', label: 'Qollect' },
            aboutVer:   { component: 'text', label: 'Version: 1.1.0' },
            aboutAuth:  { component: 'text', label: 'Author: Eli Gohar' },
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
            </ul>

            <div class="qollect__actions">
              <button id="btn-${id}" class="qollect__btn" type="button">Export Selected (XLS)</button>
            </div>
          </div>
        </div>
      `);

      const $btn = $element.find(`#btn-${id}`);
      $btn.off('click').on('click', async () => {
        const overview = $element.find(`#ovw-${id}`).is(':checked');
        const dims = $element.find(`#dims-${id}`).is(':checked');
        const msrs = $element.find(`#msrs-${id}`).is(':checked');
        const vars = $element.find(`#vars-${id}`).is(':checked');
        const flds = $element.find(`#flds-${id}`).is(':checked');
        const shts = $element.find(`#shts-${id}`).is(':checked');
        const chrs = $element.find(`#chrs-${id}`).is(':checked');
        $btn.prop('disabled', true).text('Exporting…');
        try { await exportSelected(app, fileName, { overview, dims, msrs, flds, shts, chrs, vars }); }
        catch (err) { console.error(err); alert('Export failed: ' + (err?.message || err)); }
        finally { $btn.prop('disabled', false).text('Export Selected (XLS)'); }
      });

      return qlik.Promise.resolve();
    }
  };
});
