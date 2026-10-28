/* Fills the original DepEd monitoring checklist workbooks (.xlsx) with a school's responses.
   Works on the original file's XML, so the form keeps its exact look, check boxes and formulas.
   Usage: await CidXlsx.build(JSZip, templateBytes, form, items, {mode:'single'|'all'})
   items: [{ schoolId, sheet, response|null, school:{full_name, head}, psds, ... }] */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.CidXlsx = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  const esc = s => String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/\r/g, '');
  const colNum = c => { let n = 0; for (const ch of c) n = n * 26 + ch.charCodeAt(0) - 64; return n; };
  const splitRef = ref => { const m = /^([A-Z]+)(\d+)$/.exec(ref); return { col: m[1], row: +m[2] }; };

  function rowRange(xml, row) {
    const re = new RegExp('<row r="' + row + '"[^>]*?(?:/>|>[\\s\\S]*?</row>)');
    const m = re.exec(xml);
    return m ? { start: m.index, end: m.index + m[0].length, text: m[0] } : null;
  }

  // Replace (or create) a cell. kind: 'str' | 'bool' | 'num' | 'empty'
  function setCell(xml, ref, kind, val) {
    const { col, row } = splitRef(ref);
    const re = new RegExp('<c r="' + ref + '"((?:\\s[^>]*?)?)(?:/>|>[\\s\\S]*?</c>)');
    const m = re.exec(xml);
    let style = '';
    if (m) { const sm = /\ss="(\d+)"/.exec(m[1]); if (sm) style = ' s="' + sm[1] + '"'; }
    let cell;
    if (kind === 'str') cell = val === '' || val == null ? '<c r="' + ref + '"' + style + '/>' :
      '<c r="' + ref + '"' + style + ' t="inlineStr"><is><t xml:space="preserve">' + esc(val) + '</t></is></c>';
    else if (kind === 'bool') cell = '<c r="' + ref + '"' + style + ' t="b"><v>' + (val ? 1 : 0) + '</v></c>';
    else if (kind === 'num') cell = '<c r="' + ref + '"' + style + '><v>' + val + '</v></c>';
    else cell = '<c r="' + ref + '"' + style + '/>';
    if (m) return xml.slice(0, m.index) + cell + xml.slice(m.index + m[0].length);
    const rr = rowRange(xml, row);
    if (!rr) return xml; // row not present in the template: nothing to fill
    let t = rr.text;
    if (/\/>$/.test(t) && !/<\/row>$/.test(t)) t = t.replace(/\/>$/, '>' + cell + '</row>');
    else {
      const cells = [...t.matchAll(/<c r="([A-Z]+)\d+"/g)];
      let insertAt = t.lastIndexOf('</row>');
      for (const c of cells) if (colNum(c[1]) > colNum(col)) { insertAt = c.index; break; }
      t = t.slice(0, insertAt) + cell + t.slice(insertAt);
    }
    return xml.slice(0, rr.start) + t + xml.slice(rr.end);
  }

  // Change only the cached <v> of a formula cell
  function setCached(xml, ref, n) {
    const re = new RegExp('(<c r="' + ref + '"(?:\\s[^>]*?)?>)([\\s\\S]*?)(</c>)');
    const m = re.exec(xml);
    if (!m) return xml;
    let inner = m[2];
    if (/^<c[^>]*\st="(str|s|inlineStr)"/.test(m[1])) return xml;
    if (/<v>[\s\S]*?<\/v>|<v\/>/.test(inner)) inner = inner.replace(/<v>[\s\S]*?<\/v>|<v\/>/, '<v>' + n + '</v>');
    else inner += '<v>' + n + '</v>';
    return xml.slice(0, m.index) + m[1] + inner + m[3] + xml.slice(m.index + m[0].length);
  }

  function itemKey(si, ii) { return si + '.' + ii; }

  /* Fill one school sheet's XML. resp may be null (blank form). */
  function fillSheetXml(xml, form, ctx) {
    const resp = ctx.response, ans = (resp && resp.answers) || {};
    const yes = {}, no = {};
    let out = xml;
    out = setCell(out, 'B11', 'str', String(ctx.school.full_name || '').toUpperCase());
    let term = resp && resp.term ? String(resp.term).toUpperCase() : '';
    if (term && resp.visit_date) term += '  –  Visit date: ' + fmtDate(resp.visit_date);
    out = setCell(out, 'B12', 'str', term);
    out = setCell(out, 'B13', 'str', (resp && resp.monitored_by) || '');
    form.sections.forEach((sec, si) => sec.items.forEach((it, ii) => {
      const a = ans[itemKey(si, ii)] || {};
      const y = a.a === 'Y', n = a.a === 'N';
      yes[it.row] = y; no[it.row] = n;
      out = setCell(out, 'AA' + it.row, 'bool', y);
      out = setCell(out, 'AB' + it.row, 'bool', n);
      out = setCell(out, 'D' + it.row, 'str', a.r || '');
      const ac = y ? 1 : 0, ad = n ? 1 : 0, ae = ac + ad;
      out = setCached(out, 'AC' + it.row, ac);
      out = setCached(out, 'AD' + it.row, ad);
      out = setCached(out, 'AE' + it.row, ae);
      out = setCached(out, 'AF' + it.row, ae === 1 ? ac : 0);
      out = setCached(out, 'AG' + it.row, ae === 1 ? ad : 0);
    }));
    out = setCell(out, 'A' + (form.txt.S_row + 1), 'str', (resp && resp.strengths) || '');
    out = setCell(out, 'A' + (form.txt.A_row + 1), 'str', (resp && resp.areas) || '');
    out = setCell(out, 'A' + (form.txt.N_row + 1), 'str', (resp && resp.next_action) || '');
    out = setCell(out, form.headCell, 'str', String(ctx.school.head || 'NAME OF SCHOOL HEAD').toUpperCase());
    out = setCell(out, form.psdsCell, 'str', ctx.psds);

    // Recompute cached totals (section subtotals and the grand total), so previews show right numbers
    const vals = {};
    const get = ref => vals[ref] != null ? vals[ref] : 0;
    form.sections.forEach(sec => sec.items.forEach(it => {
      const ae = (yes[it.row] ? 1 : 0) + (no[it.row] ? 1 : 0);
      vals['AF' + it.row] = ae === 1 && yes[it.row] ? 1 : 0;
      vals['AG' + it.row] = ae === 1 && no[it.row] ? 1 : 0;
    }));
    const cellRe = /<c r="(B|C|AF|AG)(\d+)"(?:\s[^>]*?)?>(?:(?!<\/c>)[\s\S])*?<f>([^<]*)<\/f>(?:(?!<\/c>)[\s\S])*?<\/c>/g;
    const formulas = [];
    let m;
    while ((m = cellRe.exec(out))) formulas.push({ col: m[1], row: +m[2], f: m[3] });
    formulas.sort((a, b) => a.row - b.row || (a.col === 'AF' || a.col === 'AG' ? -1 : 1));
    for (const c of formulas) {
      const f = c.f.replace(/\s/g, '');
      let total = null;
      const sm = /^SUM\((.+)\)$/.exec(f);
      if (sm) {
        total = 0;
        for (const part of sm[1].split(',')) {
          const r = /^(AF|AG)(\d+):(AF|AG)(\d+)$/.exec(part);
          if (r) { for (let i = +r[2]; i <= +r[4]; i++) total += get(r[1] + i); continue; }
          if (/^(AF|AG)\d+$/.test(part)) total += get(part);
        }
      } else if (/^(AF|AG)\d+$/.test(f)) total = get(f);
      if (total == null) continue;
      vals[c.col + c.row] = total;
      out = setCached(out, c.col + c.row, total);
    }
    return { xml: out, yes, no };
  }

  function fmtDate(d) {
    const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(String(d));
    if (!m) return String(d);
    const mon = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'][+m[2] - 1];
    return +m[3] + ' ' + mon + ' ' + m[1];
  }

  const dirOf = p => p.replace(/[^/]*$/, '');
  function resolve(base, target) {
    if (target.startsWith('/')) return target.slice(1);
    const parts = (dirOf(base) + target).split('/'); const out = [];
    for (const p of parts) { if (p === '..') out.pop(); else if (p !== '.') out.push(p); }
    return out.join('/');
  }

  function sheetPathMap(wbXml, relsXml) {
    const rels = {};
    for (const m of relsXml.matchAll(/<Relationship\s[^>]*?>/g)) {
      const id = /Id="([^"]+)"/.exec(m[0]), tg = /Target="([^"]+)"/.exec(m[0]);
      if (id && tg) rels[id[1]] = resolve('xl/workbook.xml', tg[1]);
    }
    const out = [];
    for (const m of wbXml.matchAll(/<sheet\s[^>]*?\/>/g)) {
      const name = /name="([^"]*)"/.exec(m[0])[1].replace(/&amp;/g, '&');
      const rid = /r:id="([^"]+)"/.exec(m[0])[1];
      out.push({ name, path: rels[rid], rid, tag: m[0] });
    }
    return out;
  }

  async function setCheckboxes(zip, sheetPath, yes, no) {
    const relPath = dirOf(sheetPath) + '_rels/' + sheetPath.split('/').pop() + '.rels';
    const rf = zip.file(relPath);
    if (!rf) return;
    const rels = await rf.async('string');
    const wanted = ref => {
      const r = /^A([AB])(\d+)$/.exec(ref.replace(/\$/g, ''));
      if (!r) return false;
      return r[1] === 'A' ? !!yes[+r[2]] : !!no[+r[2]];
    };
    for (const m of rels.matchAll(/<Relationship\s[^>]*?>/g)) {
      const tg = /Target="([^"]+)"/.exec(m[0])[1], type = /Type="([^"]+)"/.exec(m[0])[1];
      const p = resolve(sheetPath, tg);
      if (/ctrlProp$/.test(type)) {
        let x = await zip.file(p).async('string');
        const link = /fmlaLink="([^"]+)"/.exec(x);
        if (!link) continue;
        x = x.replace(/\schecked="[^"]*"/, '');
        if (wanted(link[1])) x = x.replace(/objectType="CheckBox"/, 'objectType="CheckBox" checked="Checked"');
        zip.file(p, x);
      } else if (/vmlDrawing$/.test(type)) {
        let x = await zip.file(p).async('string');
        x = x.replace(/<v:shape [\s\S]*?<\/v:shape>/g, shape => {
          const link = /<x:FmlaLink>([^<]*)<\/x:FmlaLink>/.exec(shape);
          if (!link) return shape;
          shape = shape.replace(/\s*<x:Checked>\d*<\/x:Checked>/, '');
          if (wanted(link[1])) shape = shape.replace(/(\s*)<x:FmlaLink>/, '$1<x:Checked>1</x:Checked>$1<x:FmlaLink>');
          return shape;
        });
        zip.file(p, x);
      }
    }
  }

  async function removeSheetParts(zip, sheetPath, ct) {
    const relPath = dirOf(sheetPath) + '_rels/' + sheetPath.split('/').pop() + '.rels';
    const rf = zip.file(relPath);
    const kill = [sheetPath, relPath];
    if (rf) {
      const rels = await rf.async('string');
      for (const m of rels.matchAll(/<Relationship\s[^>]*?>/g)) {
        const tg = /Target="([^"]+)"/.exec(m[0])[1];
        const p = resolve(sheetPath, tg);
        kill.push(p);
        if (/drawings\/drawing\d+\.xml$/.test(p)) kill.push(dirOf(p) + '_rels/' + p.split('/').pop() + '.rels');
      }
    }
    for (const p of kill) {
      if (zip.file(p)) zip.remove(p);
      ct.xml = ct.xml.replace(new RegExp('<Override\\s[^>]*PartName="/' + p.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '"[^>]*/>'), '');
    }
  }

  async function build(JSZip, templateBytes, form, items, opts) {
    opts = opts || {};
    const zip = await JSZip.loadAsync(templateBytes);
    let wbXml = await zip.file('xl/workbook.xml').async('string');
    let wbRels = await zip.file('xl/_rels/workbook.xml.rels').async('string');
    const sheets = sheetPathMap(wbXml, wbRels);
    const ct = { xml: await zip.file('[Content_Types].xml').async('string') };
    const byName = {}; sheets.forEach(s => byName[s.name] = s);
    const doneNames = new Set();

    for (const it of items) {
      const sh = byName[it.sheet];
      if (!sh) continue;
      const xml = await zip.file(sh.path).async('string');
      const r = fillSheetXml(xml, form, { response: it.response || null, school: it.school, psds: it.psds });
      zip.file(sh.path, r.xml);
      await setCheckboxes(zip, sh.path, r.yes, r.no);
      doneNames.add(it.sheet);
    }

    if (opts.mode === 'single') {
      const keep = items[0].sheet;
      const drop = sheets.filter(s => s.name !== keep);
      for (const s of drop) {
        await removeSheetParts(zip, s.path, ct);
        wbXml = wbXml.replace(s.tag, '');
        wbRels = wbRels.replace(new RegExp('<Relationship\\s[^>]*Id="' + s.rid + '"[^>]*/>'), '');
      }
      // defined names (print areas) refer to sheets by position: keep only the one for this sheet, re-indexed to 0
      wbXml = wbXml.replace(/<definedNames>[\s\S]*?<\/definedNames>/, dn => {
        const keepName = new RegExp("'?" + keep.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + "'?!");
        const kept = [...dn.matchAll(/<definedName\s[^>]*>[^<]*<\/definedName>/g)].map(x => x[0]).filter(x => keepName.test(x)).map(x => x.replace(/localSheetId="\d+"/, 'localSheetId="0"'));
        return kept.length ? '<definedNames>' + kept.join('') + '</definedNames>' : '';
      });
    } else {
      // all schools: blank out any school sheet that has no response, so stray template answers do not show
      for (const s of sheets) {
        if (doneNames.has(s.name) || !opts.blankSheets || !opts.blankSheets[s.name]) continue;
        const xml = await zip.file(s.path).async('string');
        const r = fillSheetXml(xml, form, { response: null, school: opts.blankSheets[s.name], psds: opts.psds });
        zip.file(s.path, r.xml);
        await setCheckboxes(zip, s.path, r.yes, r.no);
      }
    }
    wbXml = wbXml.replace(/activeTab="\d+"/, 'activeTab="0"');
    wbXml = /<calcPr[^>]*\/>/.test(wbXml) ? wbXml.replace(/<calcPr([^>]*?)\/>/, (m, a) => '<calcPr' + a.replace(/\sfullCalcOnLoad="[^"]*"/, '') + ' fullCalcOnLoad="1"/>') : wbXml.replace('</workbook>', '<calcPr fullCalcOnLoad="1"/></workbook>');
    // drop the cached calculation chain, Excel rebuilds it
    if (zip.file('xl/calcChain.xml')) {
      zip.remove('xl/calcChain.xml');
      wbRels = wbRels.replace(/<Relationship\s[^>]*calcChain[^>]*\/>/, '');
      ct.xml = ct.xml.replace(/<Override\s[^>]*calcChain[^>]*\/>/, '');
    }
    // sheets drop out of the first view: make only the first visible sheet selected
    zip.file('xl/workbook.xml', wbXml);
    zip.file('xl/_rels/workbook.xml.rels', wbRels);
    zip.file('[Content_Types].xml', ct.xml);
    return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
  }

  return { build, fillSheetXml, fmtDate };
});
