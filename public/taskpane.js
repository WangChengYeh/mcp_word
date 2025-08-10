/* Office Add-in Task Pane client: listens to Socket.IO events and applies Word operations.
   Event names and args follow tool.md. Sends results back via 'word:result'. */

(function () {
  const socket = io();

  const state = {
    ranges: new Map(), // id -> Word.Range (tracked)
    tables: new Map(), // id -> Word.Table (tracked)
  };

  const genId = (p) => `${p}:${Date.now().toString(36)}${Math.random().toString(36).slice(2, 8)}`;

  const toInsertLocation = (loc) => {
    const WordLoc = Word.InsertLocation;
    switch (loc) {
      case "start": return WordLoc.start;
      case "end": return WordLoc.end;
      case "before": return WordLoc.before;
      case "after": return WordLoc.after;
      case "replace":
      default:
        return WordLoc.replace;
    }
  };

  async function fetchAsBase64(url) {
    const res = await fetch(url);
    if (!res.ok) throw new Error(`fetch failed: ${res.status}`);
    const blob = await res.blob();
    const reader = new FileReader();
    const base64 = await new Promise((resolve, reject) => {
      reader.onloadend = () => resolve(reader.result.split(",")[1]);
      reader.onerror = reject;
      reader.readAsDataURL(blob);
    });
    return base64;
  }

  function respond(op, data, diagnostics) {
    socket.emit('word:result', { ok: true, op, data, diagnostics });
  }
  function respondErr(op, err, code = 'E_RUNTIME') {
    const msg = (err && err.message) || String(err);
    socket.emit('word:result', { ok: false, op, code, diagnostics: [{ level: 'error', msg }] });
  }

  async function getTarget(context, scope) {
    if (!scope || scope === 'selection') return context.document.getSelection();
    if (scope === 'document') return context.document.body;
    if (typeof scope === 'string' && scope.startsWith('rangeId:')) {
      const r = state.ranges.get(scope);
      if (!r) throw new Error(`range not found: ${scope}`);
      return r;
    }
    return context.document.getSelection();
  }

  Office.onReady(() => {
    // insertText
    socket.on('word:insertText', async (args = {}) => {
      try {
        await Word.run(async (context) => {
          const scope = args.scope || 'selection';
          const location = toInsertLocation(args.location || 'replace');
          const target = await getTarget(context, scope);
          const text = args.text ?? '';
          if (args.newParagraph && target.insertParagraph) {
            target.insertParagraph(text, location);
          } else {
            target.insertText(text, location);
          }
          if (target.track) { try { target.track(); context.trackedObjects.add(target); } catch {}
          }
          await context.sync();
          const id = typeof scope === 'string' && scope.startsWith('rangeId:') ? scope : genId('rangeId');
          if (!state.ranges.has(id) && target) state.ranges.set(id, target);
          respond('insertText', { rangeId: id, length: text.length });
        });
      } catch (e) { respondErr('insertText', e); }
    });

    // getSelection
    socket.on('word:getSelection', async () => {
      try {
        await Word.run(async (context) => {
          const sel = context.document.getSelection();
          sel.load(["text", "start", "end"]);
          sel.track();
          context.trackedObjects.add(sel);
          await context.sync();
          const id = genId('rangeId');
          state.ranges.set(id, sel);
          respond('getSelection', { text: sel.text || '', rangeId: id, start: sel.start, end: sel.end });
        });
      } catch (e) { respondErr('getSelection', e); }
    });

    // search
    socket.on('word:search', async (args = {}) => {
      try {
        await Word.run(async (context) => {
          const base = await getTarget(context, args.scope || 'document');
          const options = {
            matchCase: !!args.matchCase,
            matchWholeWord: !!args.matchWholeWord,
            matchPrefix: !!args.matchPrefix,
            matchSuffix: !!args.matchSuffix,
            ignoreSpace: !!args.ignoreSpace,
            ignorePunct: !!args.ignorePunct,
            matchWildcards: !!args.useRegex, // best effort; Word uses wildcards
          };
          const results = base.search(String(args.query || ''), options);
          results.load(["text"]);
          await context.sync();
          const out = [];
          for (let i = 0; i < results.items.length; i++) {
            const r = results.items[i];
            try { r.track(); context.trackedObjects.add(r); } catch {}
            const id = genId('rangeId');
            state.ranges.set(id, r);
            out.push({ rangeId: id, text: r.text || '' });
          }
          respond('search', { results: out });
        });
      } catch (e) { respondErr('search', e); }
    });

    // replace
    socket.on('word:replace', async (args = {}) => {
      try {
        await Word.run(async (context) => {
          const replaceOne = async (r, text) => { r.insertText(text, Word.InsertLocation.replace); };
          let replaced = 0;
          if (args.target === 'searchQuery') {
            const searchArgs = { ...args, scope: args.scope || 'document' };
            const base = await getTarget(context, searchArgs.scope);
            const results = base.search(String(searchArgs.query || ''), {
              matchCase: !!searchArgs.matchCase,
              matchWholeWord: !!searchArgs.matchWholeWord,
              matchPrefix: !!searchArgs.matchPrefix,
            });
            await context.sync();
            const items = results.items;
            const count = args.mode === 'replaceFirst' ? Math.min(1, items.length) : items.length;
            for (let i = 0; i < count; i++) { replaceOne(items[i], String(args.replaceWith || '')); replaced++; }
          } else {
            const target = await getTarget(context, args.target);
            replaceOne(target, String(args.replaceWith || ''));
            replaced++;
          }
          await context.sync();
          respond('replace', { replaced });
        });
      } catch (e) { respondErr('replace', e); }
    });

    // insertPicture
    socket.on('word:insertPicture', async (args = {}) => {
      try {
        await Word.run(async (context) => {
          const scope = args.scope || 'selection';
          const location = toInsertLocation(args.location || 'replace');
          const base64 = args.source === 'url' ? await fetchAsBase64(String(args.data)) : String(args.data);
          const target = await getTarget(context, scope);
          const insertAt = target.insertInlinePictureFromBase64 ? target : context.document.body;
          const pic = insertAt.insertInlinePictureFromBase64(base64, location);
          pic.load(["width", "height"]);
          await context.sync();
          if (args.width) pic.width = args.width;
          if (args.height) pic.height = args.height;
          if (args.altText) { try { pic.altTextDescription = String(args.altText); } catch {} }
          await context.sync();
          const rid = genId('rangeId');
          respond('insertPicture', { rangeId: rid });
        });
      } catch (e) { respondErr('insertPicture', e, e?.message?.includes('fetch') ? 'E_RUNTIME' : 'E_RUNTIME'); }
    });

    // ---- Tables ----
    function resolveTableRef(ref) {
      if (!ref) return null;
      if (ref.startsWith('tableId:')) return state.tables.get(ref) || null;
      if (ref.startsWith('rangeId:')) return state.ranges.get(ref) || null;
      return null;
    }

    socket.on('word:table.create', async (args = {}) => {
      try {
        await Word.run(async (context) => {
          const scope = args.scope || 'selection';
          const location = toInsertLocation(args.location || 'end');
          const base = await getTarget(context, scope);
          const rows = Number(args.rows || 1), cols = Number(args.cols || 1);
          const data = Array.isArray(args.data) ? args.data : undefined;
          const tbl = (base.insertTable ? base : context.document.body).insertTable(rows, cols, location, data);
          await context.sync();
          if (args.header === true) { try { tbl.headerRow = true; } catch {} }
          try { tbl.track(); context.trackedObjects.add(tbl); } catch {}
          const id = genId('tableId');
          state.tables.set(id, tbl);
          respond('table.create', { tableId: id });
        });
      } catch (e) { respondErr('table.create', e); }
    });

    socket.on('word:table.setCellText', async (args = {}) => {
      try { await Word.run(async (context) => {
        const t = resolveTableRef(String(args.tableRef));
        if (!t) throw new Error('tableRef not found');
        t.getCell(Number(args.row), Number(args.col)).insertText(String(args.text || ''), Word.InsertLocation.replace);
        await context.sync();
        respond('table.setCellText', {});
      }); } catch (e) { respondErr('table.setCellText', e); }
    });

    socket.on('word:table.insertRows', async (args = {}) => {
      try { await Word.run(async (context) => {
        const t = resolveTableRef(String(args.tableRef));
        if (!t) throw new Error('tableRef not found');
        const at = Number(args.at || 0);
        const count = Number(args.count || 1);
        const row = t.rows.getItemAt(at);
        row.insertRows(Word.InsertLocation.after, count);
        await context.sync();
        respond('table.insertRows', {});
      }); } catch (e) { respondErr('table.insertRows', e); }
    });

    socket.on('word:table.insertColumns', async (args = {}) => {
      try { await Word.run(async (context) => {
        const t = resolveTableRef(String(args.tableRef));
        if (!t) throw new Error('tableRef not found');
        const at = Number(args.at || 0);
        const count = Number(args.count || 1);
        const col = t.columns.getItemAt(at);
        col.insertColumns(Word.InsertLocation.after, count);
        await context.sync();
        respond('table.insertColumns', {});
      }); } catch (e) { respondErr('table.insertColumns', e); }
    });

    socket.on('word:table.deleteRows', async (args = {}) => {
      try { await Word.run(async (context) => {
        const t = resolveTableRef(String(args.tableRef));
        if (!t) throw new Error('tableRef not found');
        const idx = (args.indexes || []).map(Number).sort((a,b)=>b-a);
        idx.forEach(i => { try { t.rows.getItemAt(i).delete(); } catch {} });
        await context.sync();
        respond('table.deleteRows', { deleted: idx.length });
      }); } catch (e) { respondErr('table.deleteRows', e); }
    });

    socket.on('word:table.deleteColumns', async (args = {}) => {
      try { await Word.run(async (context) => {
        const t = resolveTableRef(String(args.tableRef));
        if (!t) throw new Error('tableRef not found');
        const idx = (args.indexes || []).map(Number).sort((a,b)=>b-a);
        idx.forEach(i => { try { t.columns.getItemAt(i).delete(); } catch {} });
        await context.sync();
        respond('table.deleteColumns', { deleted: idx.length });
      }); } catch (e) { respondErr('table.deleteColumns', e); }
    });

    socket.on('word:table.mergeCells', async (args = {}) => {
      try { await Word.run(async (context) => {
        const t = resolveTableRef(String(args.tableRef));
        if (!t) throw new Error('tableRef not found');
        const r1 = t.getCell(Number(args.startRow), Number(args.startCol));
        const r2 = t.getCell(Number(args.startRow) + Number(args.rowSpan) - 1, Number(args.startCol) + Number(args.colSpan) - 1);
        r1.merge(r2);
        await context.sync();
        respond('table.mergeCells', {});
      }); } catch (e) { respondErr('table.mergeCells', e); }
    });

    socket.on('word:table.applyStyle', async (args = {}) => {
      try { await Word.run(async (context) => {
        const t = resolveTableRef(String(args.tableRef));
        if (!t) throw new Error('tableRef not found');
        const s = args.style;
        if (typeof s === 'string') { try { t.style = s; } catch {} }
        else if (s && typeof s === 'object') {
          ['bandedRows','bandedColumns','firstRow','lastRow','firstColumn','lastColumn','totalRow'].forEach(k=>{ if (k in s) { try { t[k] = !!s[k]; } catch {} } });
        }
        await context.sync();
        respond('table.applyStyle', {});
      }); } catch (e) { respondErr('table.applyStyle', e); }
    });

    // ---- applyStyle ----
    socket.on('word:applyStyle', async (args = {}) => {
      try {
        await Word.run(async (context) => {
          const scope = args.scope || 'selection';
          const r = await getTarget(context, scope);
          // precedence
          const order = args.precedence || 'styleThenOverrides';
          const applyNamed = () => { if (args.namedStyle) { try { r.style = String(args.namedStyle); } catch {} } };
          const applyPara = () => {
            if (!args.para) return;
            const pf = r.paragraphFormat;
            const p = args.para;
            if (p.alignment) pf.alignment = p.alignment;
            if (p.lineSpacing != null) pf.lineSpacing = p.lineSpacing;
            if (p.spaceBefore != null) pf.spaceBefore = p.spaceBefore;
            if (p.spaceAfter != null) pf.spaceAfter = p.spaceAfter;
            if (p.leftIndent != null) pf.leftIndent = p.leftIndent;
            if (p.rightIndent != null) pf.rightIndent = p.rightIndent;
            if (p.firstLineIndent != null) pf.firstLineIndent = p.firstLineIndent;
            if (p.list && p.list !== 'none') {
              r.paragraphs.load();
              return context.sync().then(() => {
                r.paragraphs.items.forEach((para) => { try { para.startNewList(); } catch {} });
              });
            }
          };
          const applyChar = () => {
            if (!args.char) return;
            const f = r.font;
            const c = args.char;
            if (c.bold != null) f.bold = c.bold;
            if (c.italic != null) f.italic = c.italic;
            if (c.underline) f.underline = c.underline;
            if (c.strikeThrough != null) f.strikeThrough = c.strikeThrough;
            if (c.doubleStrikeThrough != null) f.doubleStrikeThrough = c.doubleStrikeThrough;
            if (c.allCaps != null) f.allCaps = c.allCaps;
            if (c.smallCaps != null) f.smallCaps = c.smallCaps;
            if (c.superscript != null) f.superscript = c.superscript;
            if (c.subscript != null) f.subscript = c.subscript;
            if (c.fontName) f.name = c.fontName;
            if (c.fontSize != null) f.size = c.fontSize;
            if (c.color) f.color = c.color;
            if (c.highlight) f.highlightColor = c.highlight;
          };

          if (args.resetDirectFormatting) { try { r.style = 'Normal'; } catch {} }
          if (order === 'styleThenOverrides') { applyNamed(); await context.sync(); await applyPara(); applyChar(); }
          else { await applyPara(); applyChar(); await context.sync(); applyNamed(); }
          await context.sync();
          const id = genId('rangeId');
          respond('applyStyle', { rangeId: id });
        });
      } catch (e) { respondErr('applyStyle', e); }
    });

    // listStyles
    socket.on('word:listStyles', async (args = {}) => {
      try {
        const paragraphStyles = [
          'Normal','No Spacing','Title','Subtitle','Heading 1','Heading 2','Heading 3','Heading 4','Heading 5','Heading 6','Quote','Intense Quote','List Paragraph','Caption','TOC 1','TOC 2','TOC 3'
        ];
        const characterStyles = [
          'Strong','Emphasis','Intense Emphasis','Subtle Emphasis','Subscript','Superscript','Hyperlink','FollowedHyperlink'
        ];
        const tableStyles = [
          'Table Grid','Table Grid Light','Grid Table 1 Light','Grid Table 2','Grid Table 4 Accent 1','List Table 1 Light','List Table 2','List Table 3'
        ];
        let data = { paragraphStyles: paragraphStyles.map(n=>({name:n,builtIn:true})), characterStyles: characterStyles.map(n=>({name:n,builtIn:true})), tableStyles: tableStyles.map(n=>({name:n,builtIn:true})) };
        const q = (args.query || '').toLowerCase();
        if (q) {
          const filter = (arr) => arr.filter(s => s.name.toLowerCase().includes(q));
          data = { paragraphStyles: filter(data.paragraphStyles), characterStyles: filter(data.characterStyles), tableStyles: filter(data.tableStyles) };
        }
        respond('listStyles', data);
      } catch (e) { respondErr('listStyles', e); }
    });

    // Connected
    try { socket.emit('client:ready', { host: 'word' }); } catch {}
  });
})();

