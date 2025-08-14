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
  const xlHeaderRow = headers =>
    `<Row>${headers.map(h=>`<Cell ss:StyleID="sHeader"><Data ss:Type="String">${xmlEscape(h)}</Data></Cell>`).join('')}</Row>`;

  // New: worksheet builder with auto-widths, filters, and frozen header
  const AUTO_FILTERS = true;
  const FREEZE_HEADER = true;
  const AUTO_WIDTHS = true;

  function estimateColumnWidths(headers, matrix){
    // rough px-per-character; Excel uses ~7 px per char in default font
    const PX_PER_CHAR = 7;
    const PADDING = 16;
    const MIN_W = 80;   // px
    const MAX_W = 600;  // px

    const cols = headers.length;
    const maxLens = Array.from({length: cols}, (_, i) => String(headers[i] ?? '').length);

    for (const row of matrix) {
      for (let c = 0; c < cols; c++) {
        const val = row[c];
        const len = String(val ?? '').length;
        if (len > maxLens[c]) maxLens[c] = len;
      }
    }

    return maxLens.map(len => {
      const px = len * PX_PER_CHAR + PADDING;
      return Math.max(MIN_W, Math.min(MAX_W, px));
    });
  }

  function worksheetWithFeatures(name, headers, matrix){
    const colsCount = headers.length;
    const rowsCount = matrix.length + 1; // + header row

    // Columns (auto widths)
    let columnsXml = '';
    if (AUTO_WIDTHS) {
      const widths = estimateColumnWidths(headers, matrix);
      columnsXml = widths.map((w, idx) =>
        `<Column ss:Index="${idx+1}" ss:Width="${w}"/>`
      ).join('');
    }

    // Build rows_xml
    const headerXml = xlHeaderRow(headers);
    const dataXml = matrix.map(r => row(r.map(cell))).join('');

    // Filters (AutoFilter on header row range)
    const filterXml = AUTO_FILTERS
      ? `<AutoFilter x:Range="R1C1:R${rowsCount}C${colsCount}" xmlns:x="urn:schemas-microsoft-com:office:excel"/>`
      : '';

    // Freeze header row
    const freezeXml = FREEZE_HEADER
      ? `<WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel">
           <FreezePanes/>
           <FrozenNoSplit/>
           <SplitHorizontal>1</SplitHorizontal>
           <TopRowBottomPane>1</TopRowBottomPane>
           <ActivePane>2</ActivePane>
         </WorksheetOptions>`
      : '';

    return `<Worksheet ss:Name="${xmlEscape(name)}">
      <Table>${columnsXml}${headerXml}${dataXml}</Table>
      ${filterXml}
      ${freezeXml}
    </Worksheet>`;
  }

  const xlWorkbook = body => `<?xml version="1.0"?>
<?mso-application progid="Excel.Sheet"?>
<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:o="urn:schemas-microsoft-com:office:office"
 xmlns:x="urn:schemas-microsoft-com:office:excel"
 xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:html="http://www.w3.org/TR/REC-html40">
 <Styles>
  <Style ss:ID="sHeader">
    <Font ss:Bold="1"/>
    <Interior ss:Color="#D9E1F2" ss:Pattern="Solid"/>
    <Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/></Borders>
  </Style>
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

  // ------- sort helper (case-insensitive, stable via copy) -------
  const sortAsc = (arr, key) =>
    arr.slice().sort((a,b) => String(a?.[key] ?? '').localeCompare(String(b?.[key] ?? ''), undefined, { sensitivity:'base' }));

  // ------- fetchers -------
  async function fetchDimensionsViaEngine(app){
    const doc = getDoc(app);
    const listHandle = await doc.createSessionObject({
      qInfo: { qType: 'DimensionList' },
      qDimensionListDef: {
        qType: 'dimension',
        qData: { title: '/qMetaDef/title', description: '/qMetaDef/description', tags: '/qMetaDef/tags' }
      }
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
      } catch(e){}

      if (props) {
        const meta = props.qMetaDef || {};
        const qDim = props.qDim || {};
        results.push({
          id,
          title: meta.title || it.qMeta?.title || '',
          description: meta.description || it.qMeta?.description || '',
          tags: (meta.tags || it.qMeta?.tags || []).join(', '),
          fields: (qDim.qFieldDefs || []).join(', '),
          labelExpr: qDim.qLabelExpression || ''
        });
      } else {
        results.push({
          id,
          title: it.qMeta?.title || '',
          description: it.qMeta?.description || '',
          tags: (it.qMeta?.tags || []).join(', '),
          fields: '',
          labelExpr: ''
        });
      }
    }
    return results;
  }

  async function fetchMeasuresViaEngine(app){
    const doc = getDoc(app);
    const listHandle = await doc.createSessionObject({
      qInfo: { qType: 'MeasureList' },
      qMeasureListDef: {
        qType: 'measure',
        qData: { title: '/qMetaDef/title', description: '/qMetaDef/description', tags: '/qMetaDef/tags' }
      }
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
      } catch(e){}

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
          labelExpr: qMeasure.qLabelExpression || ''
        });
      } else {
        results.push({
          id,
          title: it.qMeta?.title || '',
          description: it.qMeta?.description || '',
          tags: (it.qMeta?.tags || []).join(', '),
          expression: '',
          label: '',
          labelExpr: ''
        });
      }
    }
    return results;
  }

  async function fetchFieldsViaEngine(app){
    const doc = getDoc(app);
    const listHandle = await doc.createSessionObject({
      qInfo: { qType: 'FieldList' },
      qFieldListDef: {
        qShowSystem: false,
        qShowHidden: false,
        qShowSemantic: true,
        qShowDerivedFields: true,
        qShowImplicit: true,
        qShowSrcTables: true
      }
    });
    const layout = await listHandle.getLayout();
    const items = layout?.qFieldList?.qItems || [];
    return items.map(it => ({
      name: it.qName || '',
      tags: (it.qTags || []).join(', '),
      srcTables: Array.isArray(it.qSrcTables) ? it.qSrcTables.join(', ') : (it.qSrcTables || '')
    }));
  }

  // SHEETS
  async function fetchSheetsViaEngine(app){
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

  // CHARTS (objects on sheets)
  async function fetchChartsViaEngine(app){
    const sheets = await fetchSheetsViaEngine(app);
    const charts = [];
    for (const sh of sheets) {
      try {
        const sheetModel = await app.getObject(sh.id);
        const props = await sheetModel.getProperties();
        const cells = props?.cells || [];
        for (const c of cells) {
          const objId = c.name;
          let visType = c.type || '';
          let title = '';
          let isMaster = c.qExtendsId ? 'Y' : 'N';
          let masterId = c.qExtendsId || '';
          try {
            const objModel = await app.getObject(objId);
            const p = await objModel.getProperties();
            visType = p.visualization || visType || '';
            const metaTitle = p.qMetaDef && p.qMetaDef.title ? p.qMetaDef.title : '';
            const rawTitle = typeof p.title === 'string' ? p.title : (p.title && p.title.qStringExpression) || '';
            title = rawTitle || metaTitle || '';
            if (p.qExtendsId) { isMaster = 'Y'; masterId = p.qExtendsId; }
          } catch(e){}
          charts.push({
            sheetTitle: sh.title,
            sheetId: sh.id,
            objectId: objId,
            type: visType,
            title,
            isMaster,
            masterId
          });
        }
      } catch(e){}
    }
    return charts;
  }

  // VARIABLES (robust: engine list + capability API fallback)
  async function fetchVariablesViaEngine(app){
    const doc = getDoc(app);

    let items = [];
    try {
      const h = await doc.createSessionObject({
        qInfo: { qType: 'VariableList' },
        qVariableListDef: {
          qType: 'variable',
          qShowReserved: true,
          qShowConfig: true,
          qData: {
            tags: '/qMetaDef/tags',
            definition: '/qDefinition',
            comment: '/qComment'
          }
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
      const name = it.qName || it.qInfo?.qName || '';
      if (!name) continue;

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
          if (vh && typeof vh.getProperties === 'function') {
            props = await vh.getProperties();
          }
        }
      } catch(e){}

      const definition = props?.qDefinition ?? it.qDefinition ?? it.qData?.definition ?? '';
      const comment    = props?.qComment    ?? it.qComment    ?? it.qData?.comment    ?? '';
      const tagsArr    = props?.qMetaDef?.tags ?? it.qTags ?? it.qData?.tags ?? [];
      const tags       = Array.isArray(tagsArr) ? tagsArr.join(', ') : (tagsArr || '');
      const isScript   = props?.qIsScriptCreated ?? it.qIsScriptCreated ?? it.qIsScript;
      const isReserved = props?.qIsReserved ?? it.qIsReserved;

      vars.push({
        name,
        definition,
        comment,
        tags,
        isScript: YN(isScript),
        isReserved: YN(isReserved)
      });
    }
    return vars;
  }

  // ------- Sheet builders (now using matrix -> worksheetWithFeatures) -------
  const buildDimSheet = dims => {
    const headers = ['ID','Title','Fields','Label Expression','Description','Tags'];
    const matrix = dims.map(d => [d.id, d.title, d.fields, d.labelExpr, d.description, d.tags]);
    return worksheetWithFeatures('Dimensions', headers, matrix);
  };

  const buildMsrSheet = msrs => {
    const headers = ['ID','Title','Expression','Label','Label Expression','Description','Tags'];
    const matrix = msrs.map(m => [m.id, m.title, m.expression, m.label, m.labelExpr, m.description, m.tags]);
    return worksheetWithFeatures('Measures', headers, matrix);
  };

  const buildFldSheet = flds => {
    const headers = ['Field','Source Tables','Tags'];
    const matrix = flds.map(f => [f.name, f.srcTables, f.tags]);
    return worksheetWithFeatures('Fields', headers, matrix);
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

  // ------- Export orchestrator (now with sorting) -------
  async function exportSelected(app, fileName, opts){
    const { dims, msrs, flds, shts, chrs, vars } = opts;
    if (!dims && !msrs && !flds && !shts && !chrs && !vars) {
      alert('Nothing selected to export.');
      return;
    }
    const sheets = [];
    if (dims) sheets.push(buildDimSheet(sortAsc(await fetchDimensionsViaEngine(app), 'title')));
    if (msrs) sheets.push(buildMsrSheet(sortAsc(await fetchMeasuresViaEngine(app), 'title')));
    if (flds) sheets.push(buildFldSheet(sortAsc(await fetchFieldsViaEngine(app), 'name')));
    if (shts) sheets.push(buildShtSheet(sortAsc(await fetchSheetsViaEngine(app), 'title')));
    if (chrs) sheets.push(buildChrSheet(sortAsc(await fetchChartsViaEngine(app), 'title')));
    if (vars) sheets.push(buildVarSheet(sortAsc(await fetchVariablesViaEngine(app), 'name')));
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
            aboutVer:   { component: 'text', label: 'Version: 1.0.0' },
            aboutAuth:  { component: 'text', label: 'Author: Eli Gohar' }
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
            <h3 id="qollect-title-${id}" class="qollect__title">Qollect — export app metadata</h3>

            <ul class="qollect__list" aria-label="Metadata types">
              <li class="qollect__list-item">
                <label class="qollect__item">
                  <input id="dims-${id}" type="checkbox" class="qollect__chk" checked>
                  <span class="qollect__label">Dimensions</span>
                </label>
              </li>
              <li class="qollect__list-item">
                <label class="qollect__item">
                  <input id="msrs-${id}" type="checkbox" class="qollect__chk" checked>
                  <span class="qollect__label">Measures</span>
                </label>
              </li>
              <li class="qollect__list-item">
                <label class="qollect__item">
                  <input id="vars-${id}" type="checkbox" class="qollect__chk" checked>
                  <span class="qollect__label">Variables</span>
                </label>
              </li>
              <li class="qollect__list-item">
                <label class="qollect__item">
                  <input id="flds-${id}" type="checkbox" class="qollect__chk" checked>
                  <span class="qollect__label">Fields</span>
                </label>
              </li>
              <li class="qollect__list-item">
                <label class="qollect__item">
                  <input id="shts-${id}" type="checkbox" class="qollect__chk" checked>
                  <span class="qollect__label">Sheets</span>
                </label>
              </li>
              <li class="qollect__list-item">
                <label class="qollect__item">
                  <input id="chrs-${id}" type="checkbox" class="qollect__chk" checked>
                  <span class="qollect__label">Charts</span>
                </label>
              </li>
            </ul>

            <button id="btn-${id}" class="qollect__btn" type="button">Export Selected (XLS)</button>
            <p class="qollect__note">Each selection becomes its own sheet.</p>
          </div>
        </div>
      `);

      const $btn = $element.find(`#btn-${id}`);
      $btn.off('click').on('click', async () => {
        const dims = $element.find(`#dims-${id}`).is(':checked');
        const msrs = $element.find(`#msrs-${id}`).is(':checked');
        const vars = $element.find(`#vars-${id}`).is(':checked');
        const flds = $element.find(`#flds-${id}`).is(':checked');
        const shts = $element.find(`#shts-${id}`).is(':checked');
        const chrs = $element.find(`#chrs-${id}`).is(':checked');
        $btn.prop('disabled', true).text('Exporting…');
        try {
          await exportSelected(app, fileName, { dims, msrs, flds, shts, chrs, vars });
        } catch (err) {
          console.error(err);
          alert('Export failed: ' + (err?.message || err));
        } finally {
          $btn.prop('disabled', false).text('Export Selected (XLS)');
        }
      });

      return qlik.Promise.resolve();
    }
  };
});
