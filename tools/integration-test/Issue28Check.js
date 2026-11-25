/**
 * INTEGRATION TEST ONLY -- issue #28 (phases A, B, C) on DISPOSABLE COPIES.
 *
 * Pushed only into the bound script of a disposable copy, next to the candidate, by
 * tools/integration-test/push-issue28-copy.mjs. Never part of a release bundle; it refuses to run in the master
 * template, the five SA4#138 reports and the Legacy 6G AH report.
 *
 * Run from the Apps Script editor of the copy (Run ▸ function; the Run button passes no arguments), read the Execution log:
 *   it28Preflight()         isolation: document, triggers, script/document properties, release. Reads only.
 *   it28Snap()              the next numbered snapshot (s1, s2, ...): what changed since the one before and since s1
 *   it28Snapshots()         the snapshots kept, each verified (file, SHA-256, size, complete)
 *   it28Probe()             reads and walks the document without storing anything: sizes and times
 *   it28TimedUpdate()       one update (continuousUpdate(), the menu's path), its real duration
 *   it28Candidates()        TDocs of this report for the phase B test (a registered successor / none yet)
 *   it28Fingerprint('x') / it28Compare('x','y')   the same with labels (from code only)
 *   it28InstallRunner()     once per copy: the runner, so that every later step is pushed as a job and runs by itself
 * Snapshots: one Docs REST read with a field mask (fingerprint v2, see below); the detail is kept as gzip files in a private
 * Drive folder "DISPOSABLE Issue28 snapshots - <document>", the index (IT28_SNAP_*) in the document properties.
 * it28TimedUpdate() runs the report's own update; nothing else here changes the document.
 */

// This repository is public: the protected documents are known here by the SHA-256 digest of their ID only.
var IT28_REFUSED_DOCUMENT_DIGESTS_ = {
  '52d82baa6f520368257374a29e6cfaed042b885500d863326faab02db8cca5cc': 'SA4 Report Template (master)',
  '04c6f4603a9188d1a1c2c6c59071a47a981424511c597eac6ebe9692e9c1bac4': 'MBS SA4#138',
  '12a5d26658048eb1c104f580187ac76613aed8338d5db89c96b2b1e87fcc4996': 'Video SA4#138',
  'f0718379ccee0d1750ba2cbf97e5b2670e03895e87754284a8687f6afff726e3': '6G Media SA4#138',
  'efa7506097faaa0b4fb66f31c8c549cc502fb2dbefa621a8acbfa68fbea0855d': 'RTC SA4#138',
  '37d155fa5e383c2fa487000f43e232de7db26a8e0bddaaa66484c80ef41505da': 'Audio SA4#138',
  'e985bf59891602c65b9b3ab38329b5640c4db1da93768f412709f8c8fd4a1550': 'the Legacy 6G AH report'
};
var IT28_TEMPLATE_DOCUMENT_DIGEST_ = '52d82baa6f520368257374a29e6cfaed042b885500d863326faab02db8cca5cc';
/** The keys of the version 1 snapshots (removed by the first version 2 snapshot). */
var IT28_FP_PREFIX_ = 'IT28_FP_';

/** SHA-256 of a text, hex (Apps Script). */
function it28IdDigest_(text) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(text), Utilities.Charset.UTF_8)
    .map(function (b) { return ('0' + ((b + 256) % 256).toString(16)).slice(-2); }).join('');
}

function it28Document_() {
  const doc = DocumentApp.getActiveDocument();
  const id = doc.getId();
  const refused = IT28_REFUSED_DOCUMENT_DIGESTS_[it28IdDigest_(id)];
  if (refused) throw new Error('Issue28Check.js runs in disposable copies only; this is ' + refused + '. Nothing done.');
  if (String(doc.getName()).indexOf('DISPOSABLE') !== 0) throw new Error('The name of this document does not start with "DISPOSABLE": nothing done.');
  return doc;
}

/**
 * Pure: the Drive-like IDs in the values of a property map other than `ownId`, each with its key and what it is.
 * digestOf(id) -> hex. Values themselves are not returned (a token stays in the copy); the IDs are.
 */
function it28ForeignIds_(props, ownId, digestOf) {
  const out = [];
  Object.keys(props).sort().forEach(function (k) {
    if (k.indexOf(IT28_FP_PREFIX_) === 0 || k === 'REVIEWER_API_TOKEN') return;
    const seen = {};
    (String(props[k] || '').match(/[-\w]{25,}/g) || []).forEach(function (id) {
      if (id === ownId || seen[id] || !/[A-Za-z]/.test(id) || !/\d/.test(id)) return;
      seen[id] = true;
      const d = digestOf(id);
      out.push({ key: k, id: id, refused: IT28_REFUSED_DOCUMENT_DIGESTS_[d] || null, template: d === IT28_TEMPLATE_DOCUMENT_DIGEST_ });
    });
  });
  return out;
}

/** Editor: is this copy isolated? Reads only. */
function it28Preflight() {
  const doc = it28Document_();
  const out = { document: doc.getId(), name: doc.getName(), problems: [], review: [], notes: [] };
  try { out.release = typeof SA4_RELEASE_ !== 'undefined' ? SA4_RELEASE_.releaseId + ' / ' + SA4_RELEASE_.codeVersion : 'no Release.js'; } catch (e) { out.release = '?'; }
  const triggers = ScriptApp.getProjectTriggers();
  out.triggers = triggers.map(function (t) { return t.getHandlerFunction() + ' (' + String(t.getEventType()) + ')'; });
  if (triggers.length) out.problems.push(triggers.length + ' trigger(s) of this project for you: ' + out.triggers.join(', ') + ' -- delete them (Triggers page) before testing.');
  const scriptProps = PropertiesService.getScriptProperties().getProperties();
  out.scriptPropertyNames = Object.keys(scriptProps).sort();
  out.scriptPropertyNames.forEach(function (k) {
    if (/^SA4_REGISTRY|^SA4_STATE\|/.test(k)) out.problems.push('script property ' + k + ': a scheduler registry / central state copied from the original -- remove it before testing.');
    if (k === 'REVIEWER_API_TOKEN') out.notes.push('REVIEWER_API_TOKEN present: the copy may read the Reviewer as the original does (read-only).');
  });
  const docProps = PropertiesService.getDocumentProperties().getProperties();
  out.documentPropertyNames = Object.keys(docProps).sort();
  const ids = it28ForeignIds_(scriptProps, doc.getId(), it28IdDigest_).map(function (x) { x.where = 'script'; return x; })
    .concat(it28ForeignIds_(docProps, doc.getId(), it28IdDigest_).map(function (x) { x.where = 'document'; return x; }));
  ids.forEach(function (x) {
    const at = x.where + ' property ' + x.key + ': ' + x.id;
    if (x.template) out.notes.push(at + ' is the master template (read only: the Status dropdown definition)');
    else if (x.refused) out.problems.push(at + ' is ' + x.refused + ' -- a production report named in the copy\'s settings');
    else out.review.push(at + ' (not this copy; check what it is)');
  });
  out.meetingType = docProps.MEETING_TYPE || '';
  out.sessionRule = docProps.ADHOC_SESSION_RULE || '(upload)';
  out.carryForward = docProps.ADHOC_CARRY_FORWARD_MODE || '(off)';
  out.headings = docProps.DISPOSITION_CLASS_HEADINGS || '(normal)';
  out.localRevisionsUrl = docProps.LOCAL_REVISIONS_URL || '(none)';
  out.verdict = out.problems.length ? 'NOT ISOLATED: ' + out.problems.length + ' problem(s)' :
    out.review.length ? 'no problem; ' + out.review.length + ' foreign ID(s) to review' : 'isolated (no trigger, no registry, no production report in the settings, document named DISPOSABLE)';
  Logger.log('[IT28] preflight ' + JSON.stringify(out, null, 1));
  return out;
}

// ============================================================== fingerprint v2 (2026-11-25)
// v1 hashed objects of the Advanced Docs service as they came (JSON.stringify of colours and other objects): two
// reads of the same document gave other digests (key order is not guaranteed), and the object graph of a large
// document ran out of memory. v2 reads the document once through the Docs REST API with a field mask (only what is
// fingerprinted), canonicalises every object (sorted keys; identical adjacent runs merged; indices and heading IDs
// dropped), keeps per location separate digests for text, links and formatting with short excerpts, and compares by
// TDoc number and row label -- never by the order objects are iterated. The detailed snapshot is kept as a gzip
// file in a private Drive folder "DISPOSABLE Issue28 snapshots - <document>"; document properties keep only an
// index (file, SHA-256, size, counts). A snapshot whose file is missing, altered or incomplete is refused.

var IT28_SNAP_KEY_ = 'IT28_SNAP_';
var IT28_FOLDER_KEY_ = 'IT28_SNAP_FOLDER';
var IT28_EXCERPT_ = 600;
var IT28_LOG_LINES_ = 120;
var IT28_DROP_KEYS_ = { startIndex: true, endIndex: true, headingId: true };

var IT28_TEXTSTYLE_MASK_ = 'textStyle(bold,italic,underline,strikethrough,smallCaps,baselineOffset,fontSize,weightedFontFamily,foregroundColor,backgroundColor,link)';
var IT28_PARAGRAPH_MASK_ = 'paragraph(elements(textRun(content,suggestedInsertionIds,suggestedDeletionIds,' + IT28_TEXTSTYLE_MASK_ + '),' +
  'dropdown,person,richLink,dateElement,inlineObjectElement(inlineObjectId),footnoteReference(footnoteNumber),pageBreak,horizontalRule,columnBreak,autoText(type),equation),' +
  'paragraphStyle(namedStyleType,alignment,direction,indentStart,indentFirstLine,indentEnd,lineSpacing,spaceAbove,spaceBelow,keepWithNext,keepLinesTogether,pageBreakBefore,shading),bullet(nestingLevel))';
// The content of one tab. A field mask cannot recurse: tabs are read one nesting level per request (it28ReadSnapshot_).
var IT28_BODY_MASK_ = 'documentTab(body(content(' + IT28_PARAGRAPH_MASK_ + ',' +
  'sectionBreak(sectionStyle(sectionType)),tableOfContents(content(paragraph(elements(textRun(content))))),' +
  'table(rows,columns,tableRows(tableCells(tableCellStyle(backgroundColor,columnSpan,rowSpan,contentAlignment),content(' + IT28_PARAGRAPH_MASK_ + ',table)))))))';
var IT28_TABPROPS_MASK_ = 'tabProperties(tabId,title,index,nestingLevel,parentTabId)';
/** Tab levels read (Docs nests tabs up to three levels); a tab below them is refused, never left out. */
var IT28_TAB_LEVELS_ = 3;

/** The fields of the structure read: the tab tree, IDs and titles, no content; one level more to see a tab below the levels read. */
function it28StructureFields_() {
  let inner = 'childTabs(tabProperties(tabId))';
  for (let d = IT28_TAB_LEVELS_ - 1; d >= 1; d--) inner = 'childTabs(' + IT28_TABPROPS_MASK_ + ',' + inner + ')';
  return 'revisionId,tabs(' + IT28_TABPROPS_MASK_ + ',' + inner + ')';
}
/** The fields of the content read of the tabs at nesting level `depth` (0: the top tabs) -- and of no other tab. */
function it28LevelFields_(depth) {
  let s = 'tabProperties(tabId),' + IT28_BODY_MASK_;
  for (let d = 0; d < depth; d++) s = 'childTabs(' + s + ')';
  return 'revisionId,tabs(' + s + ')';
}

/** Pure: a 64-bit digest (two FNV-1a 32 with other seeds), hex. */
function it28Hash_(text) {
  const s = String(text);
  let h1 = 0x811c9dc5, h2 = 0x01000193 ^ 0x5bd1e995;
  for (let i = 0; i < s.length; i++) {
    const c = s.charCodeAt(i);
    h1 ^= c; h1 = Math.imul(h1, 0x01000193) >>> 0;
    h2 ^= c; h2 = Math.imul(h2, 0x01000193 ^ 0x2f) >>> 0;
  }
  return ('00000000' + h1.toString(16)).slice(-8) + ('00000000' + h2.toString(16)).slice(-8);
}

/** Pure: canonical text of a value -- keys sorted, indices and heading IDs dropped, numbers rounded. */
function it28Canon_(v) {
  if (v === null || v === undefined) return 'null';
  if (Array.isArray(v)) return '[' + v.map(it28Canon_).join(',') + ']';
  if (typeof v === 'object') {
    return '{' + Object.keys(v).filter(function (k) { return !IT28_DROP_KEYS_[k]; }).sort()
      .map(function (k) { return JSON.stringify(k) + ':' + it28Canon_(v[k]); }).join(',') + '}';
  }
  if (typeof v === 'number') return String(Math.round(v * 1000) / 1000);
  return JSON.stringify(v);
}

/** Pure: the string leaves of an object, in canonical order (to show a chip by its text). */
function it28Strings_(v, out) {
  out = out || [];
  if (typeof v === 'string') out.push(v);
  else if (Array.isArray(v)) v.forEach(function (x) { it28Strings_(x, out); });
  else if (v && typeof v === 'object') Object.keys(v).filter(function (k) { return !IT28_DROP_KEYS_[k]; }).sort().forEach(function (k) { it28Strings_(v[k], out); });
  return out;
}

/** Pure: the words of a chip: display values first; values that look like IDs left out. */
function it28ChipWords_(v) {
  const preferred = [];
  const walk = function (o) {
    if (!o || typeof o !== 'object') return;
    Object.keys(o).sort().forEach(function (k) {
      if (typeof o[k] === 'string' && /^(displayValue|displayText|title|name|text|email|uri)$/.test(k)) preferred.push(o[k]);
      else if (typeof o[k] === 'object') walk(o[k]);
    });
  };
  walk(v);
  const words = preferred.length ? preferred : it28Strings_(v).filter(function (s) { return !/^(kix|opt|suggest)\.|^[A-Za-z0-9_-]{20,}$/.test(s); });
  return words.join(' ').slice(0, 80);
}

function it28ColorHex_(c) {
  const rgb = c && c.color && c.color.rgbColor;
  if (!rgb) return c && c.color ? 'theme' : 'none';
  const h = function (x) { return ('0' + Math.round((x || 0) * 255).toString(16)).slice(-2); };
  return '#' + h(rgb.red) + h(rgb.green) + h(rgb.blue);
}

/** Pure: what a link points to: its URL, or the kind of internal target (IDs of headings change when rebuilt). */
function it28LinkTok_(link) {
  if (!link) return '';
  if (link.url) return link.url;
  if (link.bookmarkId || link.bookmark) return '#bookmark:' + (link.bookmarkId || (link.bookmark && link.bookmark.id) || '');
  if (link.headingId || link.heading) return '#heading';
  if (link.tabId) return '#tab';
  return '#internal';
}

/** Pure: readable formatting tokens of a text style (for the diff; the digest has everything). */
function it28StyleTokens_(s) {
  const t = [];
  if (s.bold) t.push('bold');
  if (s.italic) t.push('italic');
  if (s.underline) t.push('underline');
  if (s.strikethrough) t.push('strikethrough');
  if (s.smallCaps) t.push('smallcaps');
  if (s.baselineOffset && s.baselineOffset !== 'NONE' && s.baselineOffset !== 'BASELINE_OFFSET_UNSPECIFIED') t.push(String(s.baselineOffset).toLowerCase());
  if (s.fontSize && s.fontSize.magnitude) t.push('size ' + s.fontSize.magnitude);
  if (s.weightedFontFamily && s.weightedFontFamily.fontFamily) t.push('font ' + s.weightedFontFamily.fontFamily + (s.weightedFontFamily.weight && s.weightedFontFamily.weight !== 400 ? ' ' + s.weightedFontFamily.weight : ''));
  if (s.foregroundColor) t.push('color ' + it28ColorHex_(s.foregroundColor));
  if (s.backgroundColor) t.push('background ' + it28ColorHex_(s.backgroundColor));
  return t;
}

/**
 * Pure: one paragraph as a unit: text, links ("url @ text"), formatting digest material and readable tokens.
 * Adjacent runs with the same style and link are merged: how the API splits runs is not a change.
 */
function it28ParagraphUnit_(p) {
  const runs = [];
  const tokens = {};
  (p.elements || []).forEach(function (e) {
    if (e.textRun) {
      const s = {};
      const ts = e.textRun.textStyle || {};
      Object.keys(ts).forEach(function (k) { if (k !== 'link') s[k] = ts[k]; });
      const link = it28LinkTok_(ts.link);
      const key = it28Canon_(s) + (e.textRun.suggestedInsertionIds ? '+suggested-insertion' : '') + (e.textRun.suggestedDeletionIds ? '+suggested-deletion' : '');
      it28StyleTokens_(s).forEach(function (x) { tokens[x] = true; });
      if (e.textRun.suggestedInsertionIds || e.textRun.suggestedDeletionIds) tokens.suggestion = true;
      const last = runs[runs.length - 1];
      if (last && !last.chip && last.k === key && last.l === link) last.t += e.textRun.content || '';
      else runs.push({ t: e.textRun.content || '', k: key, l: link });
    } else {
      const kind = Object.keys(e).filter(function (k) { return !IT28_DROP_KEYS_[k]; }).sort()[0] || 'element';
      // A chip (dropdown, person, link chip, date ...): its words and a digest of all of it are text, so any change of it
      // is a text change; its formatting key is its kind.
      const words = it28ChipWords_(e[kind]);
      runs.push({ t: '[' + kind + (words ? ': ' + words : '') + ' #' + it28Hash_(it28Canon_(e)).slice(0, 6) + ']', k: 'chip:' + kind, l: '', chip: true });
    }
  });
  const ps = p.paragraphStyle || {};
  if (ps.namedStyleType && ps.namedStyleType !== 'NORMAL_TEXT') tokens['style ' + ps.namedStyleType] = true;
  if (ps.alignment && ps.alignment !== 'START') tokens['align ' + ps.alignment] = true;
  if (p.bullet) tokens['bullet ' + (p.bullet.nestingLevel || 0)] = true;
  const text = runs.map(function (r) { return r.t; }).join('');
  // The text with style markers ("a «b»{bold} c"), only where some run is styled: shown for a text change, so a style
  // that moved with the text is visible too.
  const styled = runs.some(function (r) { return !r.chip && r.k !== '{}'; });
  const marked = styled ? runs.map(function (r) {
    if (r.chip || r.k === '{}') return r.t;
    let s = {};
    try { s = JSON.parse(r.k.replace(/\+suggested-(insertion|deletion)/g, '')); } catch (e) { s = {}; }
    const tk = it28StyleTokens_(s).concat(/suggested/.test(r.k) ? ['suggestion'] : []);
    return tk.length ? '«' + r.t.replace(/\n$/, '') + '»{' + tk.join(',') + '}' + (/\n$/.test(r.t) ? '\n' : '') : r.t;
  }).join('').replace(/\n$/, '') : '';
  return {
    text: text.replace(/\n$/, ''),
    marked: marked,
    links: runs.filter(function (r) { return r.l; }).map(function (r) { return r.l + ' @ ' + r.t.replace(/\n$/, '').slice(0, 60); }),
    // The formatting twice: the sequence of styles (fmt), and the same with where each run ends (fmtL). A text change
    // moves the ends; fmtL counts only where the text is the same (it28EntryChanges_).
    fmt: it28Canon_(ps) + '|' + (p.bullet ? 'b' + (p.bullet.nestingLevel || 0) : '') + '|' + runs.map(function (r) { return r.k + (r.l ? ':L' : ''); }).join('~'),
    fmtL: runs.map(function (r) { return r.t.length; }).join(','),
    tokens: Object.keys(tokens).sort(),
    style: ps.namedStyleType || 'NORMAL_TEXT'
  };
}

/** Pure: the content of a table cell (paragraphs; a nested table by digest) as one unit. */
function it28CellUnit_(cell) {
  const parts = (cell.content || []).map(function (c) {
    if (c.paragraph) return it28ParagraphUnit_(c.paragraph);
    const d = it28Hash_(it28Canon_(c));
    return { text: '[' + (c.table ? 'table' : 'element') + ' ' + d + ']', marked: '', links: [], fmt: 'x' + d, fmtL: '', tokens: [], style: '' };
  });
  const tokens = {};
  parts.forEach(function (u) { u.tokens.forEach(function (x) { tokens[x] = true; }); });
  const cs = cell.tableCellStyle || {};
  if (cs.backgroundColor) tokens['cell ' + it28ColorHex_(cs.backgroundColor)] = true;
  return {
    text: parts.map(function (u) { return u.text; }).join('\n'),
    marked: parts.some(function (u) { return u.marked; }) ? parts.map(function (u) { return u.marked || u.text; }).join('\n') : '',
    links: [].concat.apply([], parts.map(function (u) { return u.links; })),
    fmt: it28Canon_(cs) + '/' + parts.map(function (u) { return u.fmt; }).join('/'),
    fmtL: parts.map(function (u) { return u.fmtL; }).join('/'),
    tokens: Object.keys(tokens).sort()
  };
}

/** Pure: the units of a table row (cells from `from`): joined text/links/format, readable tokens. */
function it28RowUnit_(row, from) {
  const cells = (row.tableCells || []).slice(from).map(it28CellUnit_);
  const tokens = {};
  cells.forEach(function (u) { u.tokens.forEach(function (x) { tokens[x] = true; }); });
  return { text: cells.map(function (u) { return u.text; }).join(' || '), textKey: cells.map(function (u) { return u.text; }).join('\u241E'),
    marked: cells.some(function (u) { return u.marked; }) ? cells.map(function (u) { return u.marked || u.text; }).join(' || ') : '', links: [].concat.apply([], cells.map(function (u) { return u.links; })), fmt: cells.map(function (u) { return u.fmt; }).join('||'), fmtL: cells.map(function (u) { return u.fmtL; }).join('||'), tokens: Object.keys(tokens).sort() };
}

/** Pure: a stored entry of a unit: excerpt, digests of text / links / format, readable tokens. */
function it28Entry_(u, extra) {
  const cut = function (s) { return s.length > IT28_EXCERPT_ ? s.slice(0, IT28_EXCERPT_) + '…' : s; };
  const e = { x: cut(u.text), t: it28Hash_(u.textKey !== undefined ? u.textKey : u.text), l: u.links.slice(0, 50), ld: it28Hash_(u.links.join('\n')), f: it28Hash_(u.fmt), fl: it28Hash_(u.fmtL || ''), ft: u.tokens.join(', ').slice(0, 300) };
  if (u.marked) e.mx = cut(u.marked);
  Object.keys(extra || {}).forEach(function (k) { e[k] = extra[k]; });
  return e;
}

function it28CellText_(cell) {
  return (cell && cell.content || []).map(function (c) { return c.paragraph ? (c.paragraph.elements || []).map(function (e) { return e.textRun ? e.textRun.content : ''; }).join('') : ''; }).join('').replace(/\n$/, '').trim();
}

/**
 * Pure: the tabs of a Docs API tab tree in document order (depth first): id, title, depth, parent, path ("2.1": the
 * first child of the second tab), and the node. A tab below the levels read is refused.
 */
function it28TabList_(tabs) {
  const out = [];
  const walk = function (list, depth, parent, prefix) {
    (list || []).forEach(function (t, i) {
      const pr = t.tabProperties || {};
      const path = prefix + (i + 1);
      if (depth >= IT28_TAB_LEVELS_) throw new Error('tab ' + path + ' is nested deeper than the ' + IT28_TAB_LEVELS_ + ' levels read -- snapshot refused, nothing stored');
      out.push({ id: pr.tabId || (depth === 0 && list.length === 1 ? '' : '?' + path), title: pr.title || '', depth: depth, parent: parent, path: path, node: t });
      walk(t.childTabs, depth + 1, pr.tabId || '', path + '.');
    });
  };
  walk(tabs, 0, '', '');
  return out;
}

/** Pure: the fingerprint of the content of one tab (its body). Deterministic, whatever the key order or run splitting. */
function it28TabContentOf_(content) {
  const snap = { tdocs: [], summary: [], other: [], headings: [], counts: { tables: 0, tdocTables: 0, paragraphs: 0, sectionBreaks: 0, links: 0 } };
  let inSummary = false;
  let under = '(start)';
  content.forEach(function (el) {
    if (el.sectionBreak) { snap.counts.sectionBreaks++; (inSummary ? snap.summary : snap.other).push({ k: 'section', x: '(section break)', t: 'section', l: [], ld: '', f: it28Hash_(it28Canon_(el.sectionBreak)), fl: '', ft: '', at: under }); return; }
    if (el.tableOfContents) { snap.other.push(it28Entry_({ text: '(table of contents) ' + it28Hash_(it28Canon_(el.tableOfContents)), links: [], fmt: '', fmtL: '', tokens: [] }, { k: 'toc', at: under })); return; }
    if (el.table) {
      snap.counts.tables++;
      const rows = el.table.tableRows || [];
      const head = (rows[0] && rows[0].tableCells) || [];
      const number = head.length >= 2 && it28CellText_(head[0]) === 'TDoc' ? it28CellText_(head[1]) : '';
      const units = rows.map(function (row) { return it28RowUnit_(row, 1); });
      units.forEach(function (u) { snap.counts.links += u.links.length; });
      rows.forEach(function (row) { const c0 = (row.tableCells || [])[0]; if (c0) snap.counts.links += it28CellUnit_(c0).links.length; });
      if (!inSummary && /\d/.test(number)) {
        const seen = {};
        const rec = { n: number, at: under, rows: [] };
        rows.forEach(function (row, r) {
          const label = it28CellText_((row.tableCells || [])[0]) || ('row ' + (r + 1));
          seen[label] = (seen[label] || 0) + 1;
          const key = label + (seen[label] > 1 ? ' #' + seen[label] : '');
          const labelUnit = it28CellUnit_((row.tableCells || [])[0] || {});
          rec.rows.push(it28Entry_(units[r], { key: key, lf: it28Hash_(labelUnit.fmt), lt: it28Hash_(labelUnit.text), ll: labelUnit.links.join(' | ') }));
        });
        snap.tdocs.push(rec);
      } else {
        const target = inSummary ? snap.summary : snap.other;
        rows.forEach(function (row) { target.push(it28Entry_(it28RowUnit_(row, 0), { k: 'row', at: under })); });
      }
      return;
    }
    if (el.paragraph) {
      snap.counts.paragraphs++;
      const u = it28ParagraphUnit_(el.paragraph);
      snap.counts.links += u.links.length;
      if (/^HEADING_|^TITLE/.test(u.style)) {
        if (u.style === 'HEADING_1') inSummary = u.text === 'Disposition Summary';
        snap.headings.push(u.style.replace('HEADING_', 'H') + ' ' + u.text.slice(0, 200));
        under = u.text.slice(0, 80);
      }
      (inSummary ? snap.summary : snap.other).push(it28Entry_(u, { k: 'p', s: u.style, at: under }));
    }
  });
  snap.counts.tdocTables = snap.tdocs.length;
  return snap;
}

/** Pure: the snapshot from its tabs (in document order) and the content fingerprint of each. */
function it28AssembleSnapshot_(revisionId, list, contents) {
  const snap = { v: 3, revisionId: revisionId || '', tabs: [], counts: { tabs: list.length, tables: 0, tdocTables: 0, paragraphs: 0, sectionBreaks: 0, links: 0 }, complete: false };
  list.forEach(function (t) {
    const c = contents[t.id];
    if (!c) throw new Error('tab ' + t.path + ' "' + t.title + '" (' + t.id + ') could not be read completely -- snapshot refused, nothing stored');
    snap.tabs.push({ id: t.id, title: t.title, depth: t.depth, parent: t.parent, path: t.path, content: c });
    Object.keys(c.counts).forEach(function (k) { snap.counts[k] += c.counts[k]; });
  });
  snap.complete = true;
  return snap;
}

/**
 * Pure: the snapshot of a whole Docs API document (all tabs with their content, as tests or a one-request read give
 * it). A document without tabs (body only) is one tab.
 */
function it28SnapshotOf_(document) {
  const list = document.tabs && document.tabs.length ? it28TabList_(document.tabs) : [{ id: '', title: '', depth: 0, parent: '', path: '1', node: { documentTab: document } }];
  const contents = {};
  list.forEach(function (t) {
    const content = t.node.documentTab && t.node.documentTab.body && t.node.documentTab.body.content;
    if (!Array.isArray(content)) throw new Error('tab ' + t.path + ' "' + t.title + '" has no body content -- snapshot refused, nothing stored');
    if (t.id in contents) throw new Error('tab ID ' + t.id + ' twice -- snapshot refused');
    contents[t.id] = it28TabContentOf_(content);
  });
  return it28AssembleSnapshot_(document.revisionId, list, contents);
}

/** Pure: entries of two sequences that are not in the other (LCS when small, else by count), and whether the order differs. */
function it28SeqDiff_(a, b, key) {
  const ka = a.map(key), kb = b.map(key);
  if (ka.join('\u0001') === kb.join('\u0001')) return null;
  const added = [], removed = [];
  if (ka.length * kb.length <= 4000000) {
    const n = ka.length, m = kb.length;
    const dp = [];
    for (let i = 0; i <= n; i++) dp.push(new Int32Array(m + 1));
    for (let i = n - 1; i >= 0; i--) for (let j = m - 1; j >= 0; j--) dp[i][j] = ka[i] === kb[j] ? dp[i + 1][j + 1] + 1 : Math.max(dp[i + 1][j], dp[i][j + 1]);
    let i = 0, j = 0;
    while (i < n && j < m) {
      if (ka[i] === kb[j]) { i++; j++; } else if (dp[i + 1][j] >= dp[i][j + 1]) { removed.push({ i: i, e: a[i] }); i++; } else { added.push({ i: j, e: b[j] }); j++; }
    }
    for (; i < n; i++) removed.push({ i: i, e: a[i] });
    for (; j < m; j++) added.push({ i: j, e: b[j] });
  } else {
    const count = {};
    ka.forEach(function (k) { count[k] = (count[k] || 0) + 1; });
    kb.forEach(function (k, j) { if (count[k]) count[k]--; else added.push({ i: j, e: b[j] }); });
    const back = {};
    kb.forEach(function (k) { back[k] = (back[k] || 0) + 1; });
    ka.forEach(function (k, i) { if (back[k]) back[k]--; else removed.push({ i: i, e: a[i] }); });
  }
  return { added: added, removed: removed, orderOnly: !added.length && !removed.length };
}

/** Pure: what changed between two entries of the same location: text, links, formatting -- old and new. */
function it28EntryChanges_(where, a, b) {
  const out = [];
  if ((a.lt || '') !== (b.lt || '') || (a.ll || '') !== (b.ll || '')) out.push({ where: where, field: 'label cell (text or links)', old: a.ll || '(no link)', new: b.ll || '(no link)' });
  // A text change in styled text is shown with style markers: a style that moved with the text is seen as well.
  const ox = a.mx || b.mx ? (a.mx || a.x) : a.x, nx = a.mx || b.mx ? (b.mx || b.x) : b.x;
  if (a.t !== b.t) out.push({ where: where, field: a.mx || b.mx ? 'text (with its styles)' : 'text', old: ox, new: nx === ox ? nx + ' (the first ' + IT28_EXCERPT_ + ' characters are the same; it differs after)' : nx });
  if (a.ld !== b.ld) out.push({ where: where, field: 'links', old: a.l.join(' | ') || '(none)', new: b.l.join(' | ') || '(none)' });
  if (a.f !== b.f || (a.t === b.t && a.fl !== b.fl) || (a.lf || '') !== (b.lf || '')) out.push({ where: where, field: 'formatting', old: a.ft || '(plain)', new: b.ft === a.ft ? (b.ft || '(plain)') + ' (the same styles at other places, or other paragraph/cell settings)' : (b.ft || '(plain)') });
  return out;
}

/** Pure: what changed in one tab's content from a to b (P: the tab named before each place, '' in a document of one tab). */
function it28TabDiff_(a, b, P) {
  const out = { identical: false, tdocsAdded: [], tdocsRemoved: [], duplicates: [], tdocOrder: null, titles: [], changes: [], headings: null, summary: null, other: null };
  const byNumber = function (s) { const m = {}; s.tdocs.forEach(function (t) { (m[t.n] = m[t.n] || []).push(t); }); return m; };
  const ma = byNumber(a), mb = byNumber(b);
  const numbers = Object.keys(ma).concat(Object.keys(mb).filter(function (n) { return !ma[n]; })).sort();
  numbers.forEach(function (n) {
    const xa = ma[n] || [], xb = mb[n] || [];
    if (!xa.length) { out.tdocsAdded.push(P + n + ' (under "' + xb[0].at + '")'); return; }
    if (!xb.length) { out.tdocsRemoved.push(P + n + ' (was under "' + xa[0].at + '")'); return; }
    if (xa.length !== xb.length || xb.length > 1) out.duplicates.push(P + n + ': ' + xa.length + ' table(s) -> ' + xb.length);
    for (let o = 0; o < Math.min(xa.length, xb.length); o++) {
      const ra = {}, rb = {};
      xa[o].rows.forEach(function (r) { ra[r.key] = r; });
      xb[o].rows.forEach(function (r) { rb[r.key] = r; });
      const keys = xa[o].rows.map(function (r) { return r.key; }).concat(xb[o].rows.map(function (r) { return r.key; }).filter(function (k) { return !ra[k]; }));
      keys.forEach(function (k) {
        const where = P + 'TDoc ' + n + (xb.length > 1 ? ' (table ' + (o + 1) + ')' : '') + ', row "' + k + '"';
        if (!ra[k]) { out.changes.push({ where: where, field: 'row added', old: '', new: rb[k].x }); return; }
        if (!rb[k]) { out.changes.push({ where: where, field: 'row removed', old: ra[k].x, new: '' }); return; }
        const ch = it28EntryChanges_(where, ra[k], rb[k]);
        if (/^title$/i.test(k) && ra[k].t !== rb[k].t) out.titles.push({ tdoc: P + n, old: ra[k].x, new: rb[k].x });
        Array.prototype.push.apply(out.changes, ch);
      });
    }
  });
  const order = function (s) { return s.tdocs.map(function (t) { return t.n; }).filter(function (n) { return ma[n] && mb[n]; }); };
  const oa = order(a), ob = order(b);
  for (let i = 0; i < Math.max(oa.length, ob.length); i++) if (oa[i] !== ob[i]) { out.tdocOrder = { firstDifferenceAt: i + 1, was: oa[i] || '', now: ob[i] || '' }; if (P) out.tdocOrder.tab = P.replace(/: $/, ''); break; }
  const hd = it28SeqDiff_(a.headings, b.headings, function (h) { return h; });
  if (hd) out.headings = { added: hd.added.map(function (x) { return P + '#' + (x.i + 1) + ' ' + x.e; }), removed: hd.removed.map(function (x) { return P + '#' + (x.i + 1) + ' ' + x.e; }), orderOnly: hd.orderOnly };
  const entryKey = function (e) { return e.k + '|' + (e.s || '') + '|' + e.t + '|' + e.ld + '|' + e.f + '|' + e.fl; };
  const seq = function (xa, xb) {
    const d = it28SeqDiff_(xa, xb, entryKey);
    if (!d) return null;
    // A removed and an added entry at the same place: one changed entry, shown field by field.
    const changed = [];
    const addedAt = {};
    d.added.forEach(function (x) { addedAt[x.e.at + '|' + x.e.k] = (addedAt[x.e.at + '|' + x.e.k] || []).concat([x]); });
    const addedKind = {};
    d.added.forEach(function (x) { addedKind[x.e.k] = (addedKind[x.e.k] || []).concat([x]); });
    const take = function (pool) { while (pool && pool.length) { const y = pool.shift(); if (!y.used) return y; } return null; };
    const removed = [];
    const moved = [];
    d.removed.forEach(function (x) {
      // The same place first; else the next unpaired entry of the same kind (a line that became a heading moves its place).
      const y = take(addedAt[x.e.at + '|' + x.e.k]) || take(addedKind[x.e.k]);
      if (y) {
        y.used = true;
        const ch = it28EntryChanges_(P + '"' + x.e.at + '" (' + x.e.k + ' ' + (x.i + 1) + ')', x.e, y.e);
        // Nothing differs but the place: a move.
        if (ch.length) changed.push.apply(changed, ch); else moved.push({ x: x.e.x.slice(0, 200), from: x.i + 1 + ' (under "' + P + x.e.at + '")', to: y.i + 1 + ' (under "' + P + y.e.at + '")' });
      } else removed.push(x);
    });
    return { changed: changed, added: d.added.filter(function (x) { return !x.used; }).map(function (x) { return { at: P + x.e.at, x: x.e.x.slice(0, 200) }; }),
      removed: removed.map(function (x) { return { at: P + x.e.at, x: x.e.x.slice(0, 200) }; }), moved: moved, orderOnly: d.orderOnly };
  };
  out.summary = seq(a.summary, b.summary);
  out.other = seq(a.other, b.other);
  out.identical = !out.tdocsAdded.length && !out.tdocsRemoved.length && !out.duplicates.filter(function (d) { return !/: (\d+) table\(s\) -> \1$/.test(d); }).length &&
    !out.tdocOrder && !out.changes.length && !out.headings && !out.summary && !out.other;
  return out;
}

/**
 * Pure: what changed from snapshot a to b: the tabs (by tab ID: added, removed, renamed, moved), then each tab present
 * in both, located and with old and new values. Identical snapshots: identical true, nothing listed.
 */
function it28DiffOf_(a, b) {
  const out = { identical: false, release: (a.release || '?') + ' -> ' + (b.release || '?'), counts: {}, tabs: null, tdocsAdded: [], tdocsRemoved: [], duplicates: [], tdocOrder: null, tdocOrders: [], titles: [], changes: [], headings: null, summary: null, other: null };
  Object.keys(a.counts).concat(Object.keys(b.counts).filter(function (k) { return !(k in a.counts); })).forEach(function (k) { if (a.counts[k] !== b.counts[k]) out.counts[k] = a.counts[k] + ' -> ' + b.counts[k]; });
  const ta = {}, tb = {};
  a.tabs.forEach(function (x) { ta[x.id] = x; });
  b.tabs.forEach(function (x) { tb[x.id] = x; });
  const name = function (x) { return 'tab ' + x.path + ' "' + x.title + '"'; };
  const tabs = { added: [], removed: [], renamed: [], moved: [] };
  b.tabs.forEach(function (x) { if (!(x.id in ta)) tabs.added.push(name(x) + ' (' + x.id + ', ' + x.content.counts.tdocTables + ' TDoc tables)'); });
  a.tabs.forEach(function (x) { if (!(x.id in tb)) tabs.removed.push(name(x) + ' (' + x.id + ', ' + x.content.counts.tdocTables + ' TDoc tables)'); });
  b.tabs.forEach(function (x) {
    const y = ta[x.id];
    if (!y) return;
    if (y.title !== x.title) tabs.renamed.push('tab ' + x.id + ': "' + y.title + '" -> "' + x.title + '"');
    // Moved: another parent, or another place among its siblings (a child of a moved tab is not moved itself).
    if (y.parent !== x.parent || y.path.split('.').pop() !== x.path.split('.').pop()) tabs.moved.push('tab "' + x.title + '" (' + x.id + '): ' + y.path + (y.parent ? ' under ' + y.parent : '') + ' -> ' + x.path + (x.parent ? ' under ' + x.parent : ''));
  });
  if (tabs.added.length || tabs.removed.length || tabs.renamed.length || tabs.moved.length) out.tabs = tabs;
  const multi = a.tabs.length > 1 || b.tabs.length > 1;
  const mergeSeq = function (acc, s) {
    if (!s) return acc;
    if (!acc) return { changed: s.changed.slice(), added: s.added.slice(), removed: s.removed.slice(), moved: s.moved.slice(), orderOnly: s.orderOnly };
    ['changed', 'added', 'removed', 'moved'].forEach(function (k) { Array.prototype.push.apply(acc[k], s[k]); });
    acc.orderOnly = acc.orderOnly && s.orderOnly;
    return acc;
  };
  let allIdentical = true;
  b.tabs.forEach(function (x) {
    const y = ta[x.id];
    if (!y) return;
    const d = it28TabDiff_(y.content, x.content, multi ? 'tab "' + x.title + '": ' : '');
    if (!d.identical) allIdentical = false;
    ['tdocsAdded', 'tdocsRemoved', 'duplicates', 'titles', 'changes'].forEach(function (k) { Array.prototype.push.apply(out[k], d[k]); });
    if (d.tdocOrder) { out.tdocOrders.push(d.tdocOrder); if (!out.tdocOrder) out.tdocOrder = d.tdocOrder; }
    if (d.headings) {
      if (!out.headings) out.headings = { added: [], removed: [], orderOnly: true };
      Array.prototype.push.apply(out.headings.added, d.headings.added);
      Array.prototype.push.apply(out.headings.removed, d.headings.removed);
      out.headings.orderOnly = out.headings.orderOnly && d.headings.orderOnly;
    }
    out.summary = mergeSeq(out.summary, d.summary);
    out.other = mergeSeq(out.other, d.other);
  });
  out.identical = allIdentical && !out.tabs && !Object.keys(out.counts).length;
  return out;
}

/** Pure: the diff as text lines (all of them; the log shows the first IT28_LOG_LINES_). */
function it28DiffLines_(d) {
  const L = [];
  L.push('release ' + d.release + (d.identical ? ': IDENTICAL (no difference)' : ''));
  Object.keys(d.counts).forEach(function (k) { L.push('count ' + k + ': ' + d.counts[k]); });
  if (d.tabs) {
    d.tabs.added.forEach(function (x) { L.push('tab added: ' + x); });
    d.tabs.removed.forEach(function (x) { L.push('tab removed: ' + x); });
    d.tabs.renamed.forEach(function (x) { L.push('tab renamed: ' + x); });
    d.tabs.moved.forEach(function (x) { L.push('tab moved: ' + x); });
  }
  d.tdocsAdded.forEach(function (n) { L.push('TDoc table added: ' + n); });
  d.tdocsRemoved.forEach(function (n) { L.push('TDoc table removed: ' + n); });
  d.duplicates.forEach(function (n) { L.push('TDoc table more than once: ' + n); });
  (d.tdocOrders && d.tdocOrders.length ? d.tdocOrders : d.tdocOrder ? [d.tdocOrder] : []).forEach(function (o) { L.push((o.tab ? o.tab + ': ' : '') + 'order of the TDoc tables changed at #' + o.firstDifferenceAt + ': ' + o.was + ' -> ' + o.now); });
  d.titles.forEach(function (t) { L.push('title of ' + t.tdoc + ': "' + t.old + '" -> "' + t.new + '"'); });
  d.changes.forEach(function (c) { L.push(c.where + ' -- ' + c.field + ': "' + String(c.old).slice(0, 300) + '" -> "' + String(c.new).slice(0, 300) + '"'); });
  if (d.headings) { d.headings.removed.forEach(function (h) { L.push('heading removed: ' + h); }); d.headings.added.forEach(function (h) { L.push('heading added: ' + h); }); if (d.headings.orderOnly) L.push('headings: same, other order'); }
  [['summary', d.summary], ['outside the TDoc tables', d.other]].forEach(function (p) {
    if (!p[1]) return;
    p[1].changed.forEach(function (c) { L.push(p[0] + ': ' + c.where + ' -- ' + c.field + ': "' + String(c.old).slice(0, 300) + '" -> "' + String(c.new).slice(0, 300) + '"'); });
    p[1].removed.forEach(function (x) { L.push(p[0] + ': removed under "' + x.at + '": "' + x.x + '"'); });
    p[1].added.forEach(function (x) { L.push(p[0] + ': added under "' + x.at + '": "' + x.x + '"'); });
    p[1].moved.forEach(function (m) { L.push(p[0] + ': moved from #' + m.from + ' to #' + m.to + ': "' + m.x + '"'); });
    if (p[1].orderOnly) L.push(p[0] + ': same entries, other order');
  });
  if (!d.identical && L.length === 1) L.push('NOT IDENTICAL, but no difference could be itemised -- report this (issue #28)');
  return L;
}

// ---- impure: Docs API, Drive, properties

/** One Docs REST read with the given fields, parsed; the raw text is not kept. Logs size and times. */
function it28DocsGet_(documentId, fields, timing, what) {
  const url = 'https://docs.googleapis.com/v1/documents/' + encodeURIComponent(documentId) + '?includeTabsContent=true&fields=' + encodeURIComponent(fields).replace(/%28/g, '(').replace(/%29/g, ')').replace(/%2C/g, ',');
  if (url.length > 2000) throw new Error('the Docs API request for ' + what + ' is too long (' + url.length + ')');
  let started = Date.now();
  const response = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true });
  const code = response.getResponseCode();
  let text = response.getContentText();
  const read = { what: what, chars: text.length, fetchMs: Date.now() - started };
  if (code !== 200) throw new Error('the Docs API answered ' + code + ' (' + what + '): ' + text.slice(0, 300));
  console.log('[IT28] ' + what + ': read ' + text.length + ' characters in ' + read.fetchMs + ' ms');
  started = Date.now();
  const parsed = JSON.parse(text);
  text = null;
  read.parseMs = Date.now() - started;
  timing.reads.push(read);
  timing.responseChars += read.chars;
  return parsed;
}

/**
 * The snapshot of the document: the tab tree first (no content), then one request per nesting level with the content
 * of that level's tabs only; each tab is walked and released at once. Refused (nothing stored) when the document
 * changes between the requests, when a tab of the tree is not returned with its whole body, or when an unknown tab is.
 */
function it28ReadSnapshot_(documentId, timing) {
  timing.reads = []; timing.responseChars = 0; timing.walkMs = 0;
  const structure = it28DocsGet_(documentId, it28StructureFields_(), timing, 'tab structure');
  const list = it28TabList_(structure.tabs || []);
  if (!list.length) throw new Error('the Docs API returned no tab -- snapshot refused');
  const known = {};
  list.forEach(function (t) { if (t.id in known) throw new Error('tab ID ' + t.id + ' twice -- snapshot refused'); known[t.id] = t; });
  console.log('[IT28] ' + list.length + ' tab(s): ' + list.map(function (t) { return t.path + ' "' + t.title + '" (' + t.id + ')'; }).join('; '));
  const contents = {};
  const depths = list.map(function (t) { return t.depth; }).filter(function (d, i, all) { return all.indexOf(d) === i; }).sort();
  depths.forEach(function (depth) {
    let part = it28DocsGet_(documentId, it28LevelFields_(depth), timing, 'tabs of level ' + (depth + 1));
    if (part.revisionId !== structure.revisionId) throw new Error('the document changed while it was read (revision ' + String(structure.revisionId).slice(0, 12) + ' -> ' + String(part.revisionId).slice(0, 12) + ') -- snapshot refused, nothing stored');
    const got = [];
    const walk = function (tabs, d) { (tabs || []).forEach(function (x) { if (d === depth) got.push(x); else walk(x.childTabs, d + 1); }); };
    walk(part.tabs, 0);
    got.forEach(function (x) {
      const id = (x.tabProperties && x.tabProperties.tabId) || (list.length === 1 ? '' : null);
      const content = x.documentTab && x.documentTab.body && x.documentTab.body.content;
      if (id === null || !(id in known) || known[id].depth !== depth) throw new Error('the Docs API returned a tab that is not in the tab structure (' + id + ') -- snapshot refused, nothing stored');
      if (!Array.isArray(content)) throw new Error('tab ' + known[id].path + ' "' + known[id].title + '" was returned without its body -- snapshot refused, nothing stored');
      const t0 = Date.now();
      contents[id] = it28TabContentOf_(content);
      timing.walkMs += Date.now() - t0;
      x.documentTab = null;
    });
    part = null;
  });
  return it28AssembleSnapshot_(structure.revisionId, list, contents);
}

function it28Folder_() {
  const props = PropertiesService.getDocumentProperties();
  const id = props.getProperty(IT28_FOLDER_KEY_);
  if (id) { try { const f = DriveApp.getFolderById(id); if (!f.isTrashed()) return f; } catch (e) { /* made again */ } }
  const folder = DriveApp.createFolder('DISPOSABLE Issue28 snapshots - ' + DocumentApp.getActiveDocument().getName());
  props.setProperty(IT28_FOLDER_KEY_, folder.getId());
  return folder;
}
/** Stores text (gzip) in the snapshot folder; returns the file ID. Replaced in the tests. */
function it28PutFile_(name, text, gzip) {
  const blob = Utilities.newBlob(text, gzip ? 'application/json' : 'text/plain', name);
  return it28Folder_().createFile(gzip ? Utilities.gzip(blob, name + '.gz') : blob).getId();
}
/** The text of a stored file. Replaced in the tests. */
function it28GetFile_(id) {
  return Utilities.ungzip(DriveApp.getFileById(id).getBlob()).getDataAsString('UTF-8');
}

/** The SHA-256 of a snapshot's content: without time, release, revision (valid 24 h, per user) and itself. */
function it28ContentSha_(snap) {
  const c = {};
  Object.keys(snap).forEach(function (k) { if (['at', 'release', 'revisionId', 'contentSha'].indexOf(k) === -1) c[k] = snap[k]; });
  return it28IdDigest_(JSON.stringify(c));
}

/** Stores a snapshot; the index (meta) goes into the document properties. */
function it28SaveSnapshot_(label, snap) {
  const text = JSON.stringify(snap);
  const meta = { v: 3, label: label, tabs: snap.counts.tabs, sha256: it28IdDigest_(text), contentSha: snap.contentSha, chars: text.length, at: snap.at, release: snap.release, revisionId: snap.revisionId,
    tdocTables: snap.counts.tdocTables, tables: snap.counts.tables, paragraphs: snap.counts.paragraphs, links: snap.counts.links, fileId: '' };
  meta.fileId = it28PutFile_('IT28 ' + label + '.json', text, true);
  PropertiesService.getDocumentProperties().setProperty(IT28_SNAP_KEY_ + label, JSON.stringify(meta));
  return meta;
}

/** A stored snapshot, verified: the file exists, its SHA-256 and size are those of the index, it is complete. */
function it28LoadSnapshot_(label) {
  const raw = PropertiesService.getDocumentProperties().getProperty(IT28_SNAP_KEY_ + label);
  if (!raw) throw new Error('no snapshot "' + label + '"');
  const meta = JSON.parse(raw);
  let text;
  try { text = it28GetFile_(meta.fileId); } catch (e) { throw new Error('snapshot "' + label + '": its file cannot be read (' + e.message + ') -- refused'); }
  if (text.length !== meta.chars || it28IdDigest_(text) !== meta.sha256) throw new Error('snapshot "' + label + '": its file is not the one stored (size or SHA-256 differ) -- refused');
  const snap = JSON.parse(text);
  const tdocTables = (snap.tabs || []).reduce(function (n, x) { return n + ((x.content && x.content.tdocs) || []).length; }, 0);
  if (snap.v !== 3 || snap.complete !== true || tdocTables !== meta.tdocTables || snap.counts.tdocTables !== meta.tdocTables || snap.counts.tables !== meta.tables || snap.tabs.length !== meta.tabs || snap.counts.tabs !== meta.tabs) throw new Error('snapshot "' + label + '" is incomplete or of another version -- refused');
  if (it28ContentSha_(snap) !== snap.contentSha || snap.contentSha !== meta.contentSha) throw new Error('snapshot "' + label + '": its content does not match its content SHA-256 -- refused');
  return snap;
}

/** Pure: the numbered snapshots (s1, s2, ...) of a property map, in order. */
function it28SnapLabels_(all) {
  return Object.keys(all).map(function (k) { const m = /^IT28_SNAP_(s\d+)$/.exec(k); return m ? m[1] : null; }).filter(Boolean)
    .sort(function (x, y) { return Number(x.slice(1)) - Number(y.slice(1)); });
}

/** Editor: keeps a snapshot of the report now as `label`; logs counts, size and the time of each step. */
function it28Fingerprint(label) {
  const doc = it28Document_();
  const name = String(label || 'x' + new Date().toISOString().slice(11, 19).replace(/:/g, ''));
  const props = PropertiesService.getDocumentProperties();
  const old = Object.keys(props.getProperties()).filter(function (k) { return k.indexOf(IT28_FP_PREFIX_) === 0; });
  if (old.length) { old.forEach(function (k) { props.deleteProperty(k); }); console.log('[IT28] ' + old.length + ' property key(s) of the version 1 snapshots removed (not comparable).'); }
  const timing = {};
  const started = Date.now();
  const snap = it28ReadSnapshot_(doc.getId(), timing);
  let t0;
  // The content alone (before time and release are added): the same document gives the same contentSha.
  snap.contentSha = it28ContentSha_(snap);
  snap.at = new Date().toISOString();
  try { snap.release = SA4_RELEASE_.releaseId + ' / ' + SA4_RELEASE_.codeVersion; } catch (e) { snap.release = '?'; }
  t0 = Date.now();
  const meta = it28SaveSnapshot_(name, snap);
  timing.storeMs = Date.now() - t0;
  timing.totalMs = Date.now() - started;
  Logger.log('[IT28] snapshot "' + name + '" (' + snap.release + '): ' + snap.counts.tdocTables + ' TDoc tables, ' + snap.counts.tables + ' tables, ' + snap.counts.paragraphs + ' paragraphs, ' +
    snap.counts.links + ' links, ' + snap.counts.sectionBreaks + ' section breaks, ' + snap.counts.tabs + ' tab(s); revision ' + String(snap.revisionId).slice(0, 12) + '; stored ' + meta.chars +
    ' characters; content SHA-256 ' + snap.contentSha.slice(0, 16) + '.\n[IT28] time: ' + JSON.stringify(timing));
  return name;
}

/** Editor: what changed from snapshot a to b; the full list is also stored as a text file next to the snapshots. */
function it28Compare(a, b) {
  it28Document_();
  const t0 = Date.now();
  const d = it28DiffOf_(it28LoadSnapshot_(a), it28LoadSnapshot_(b));
  const lines = it28DiffLines_(d);
  let file = '';
  if (!d.identical) file = it28PutFile_('IT28 compare ' + a + ' to ' + b + '.txt', lines.join('\n'), false);
  Logger.log('[IT28] compare "' + a + '" -> "' + b + '" (' + (Date.now() - t0) + ' ms): ' + (d.identical ? 'IDENTICAL' : lines.length + ' line(s)' + (file ? '; all of them in the snapshot folder, file ' + file : '')) +
    '\n' + lines.slice(0, IT28_LOG_LINES_).join('\n') + (lines.length > IT28_LOG_LINES_ ? '\n… ' + (lines.length - IT28_LOG_LINES_) + ' more' : ''));
  return d;
}

/** Editor (no arguments): the next numbered snapshot, compared with the one before and with s1 (the baseline). */
function it28Snap() {
  it28Document_();
  const labels = it28SnapLabels_(PropertiesService.getDocumentProperties().getProperties());
  const name = 's' + (labels.length ? Number(labels[labels.length - 1].slice(1)) + 1 : 1);
  it28Fingerprint(name);
  if (labels.length) {
    it28Compare(labels[labels.length - 1], name);
    if (labels[0] !== labels[labels.length - 1]) it28Compare(labels[0], name);
  } else {
    Logger.log('[IT28] "' + name + '" is the baseline: later snapshots are compared with it.');
  }
  return name;
}

/** Editor (no arguments): the snapshots kept (from the index; each file verified). */
function it28Snapshots() {
  it28Document_();
  const all = PropertiesService.getDocumentProperties().getProperties();
  const names = Object.keys(all).filter(function (k) { return k.indexOf(IT28_SNAP_KEY_) === 0 && k !== IT28_FOLDER_KEY_; }).map(function (k) { return k.slice(IT28_SNAP_KEY_.length); });
  const rows = it28SnapLabels_(all).concat(names.filter(function (n) { return !/^s\d+$/.test(n); }).sort()).map(function (n) {
    const m = JSON.parse(all[IT28_SNAP_KEY_ + n]);
    let ok;
    try { it28LoadSnapshot_(n); ok = 'verified'; } catch (e) { ok = 'NOT USABLE: ' + e.message; }
    return n + ': ' + m.at + ', ' + m.release + ', revision ' + String(m.revisionId).slice(0, 12) + ', ' + m.tdocTables + ' TDoc tables, ' + m.chars + ' characters, content SHA-256 ' + String(m.contentSha).slice(0, 16) + ' -- ' + ok;
  });
  Logger.log('[IT28] snapshots:\n' + rows.join('\n'));
  return rows;
}

/** Editor (no arguments): reads and walks the document without storing anything; per tab what it holds; sizes and times. */
function it28Probe() {
  const doc = it28Document_();
  const timing = {};
  const snap = it28ReadSnapshot_(doc.getId(), timing);
  const text = JSON.stringify(snap);
  snap.contentSha = it28ContentSha_(snap);
  const tabs = snap.tabs.map(function (x) {
    const c = x.content;
    const heads = c.headings.filter(function (h) { return /^H[12] /.test(h); }).slice(0, 4).map(function (h) { return h.slice(0, 70); });
    return '  tab ' + x.path + ' "' + x.title + '" (' + x.id + (x.parent ? ', in ' + x.parent : '') + '): ' + c.counts.tdocTables + ' TDoc tables, ' + c.counts.tables + ' tables, ' +
      c.counts.paragraphs + ' paragraphs, summary ' + (c.summary.length ? 'yes' : 'no') + '; first headings: ' + (heads.join(' | ') || '(none)');
  });
  Logger.log('[IT28] probe: ' + snap.counts.tabs + ' tab(s), ' + snap.counts.tdocTables + ' TDoc tables, ' + snap.counts.tables + ' tables, ' + snap.counts.paragraphs + ' paragraphs; the snapshot would be ' + text.length +
    ' characters (content SHA-256 ' + snap.contentSha.slice(0, 16) + ')\n' + tabs.join('\n') + '\n[IT28] time: ' + JSON.stringify(timing));
  return timing;
}

/** Editor: one update, as the menu runs it, and its duration. */
function it28TimedUpdate() {
  it28Document_();
  const started = Date.now();
  continuousUpdate();
  const ms = Date.now() - started;
  Logger.log('[IT28] update took ' + ms + ' ms (' + (ms / 1000).toFixed(1) + ' s).');
  return ms;
}

/** Editor: TDocs for the phase B test -- originals whose successor is registered and undecided in the report, and originals revised to a TDoc not yet listed. */
function it28Candidates() {
  const doc = it28Document_();
  const all = flattenTdocGroups_(downloadAndGroupTdocs_(getReportConfig_(), null));
  const facts = plenaryListFactsOf_(all);
  const index = buildTdocTableIndex_(doc.getBody().getTables());
  const statusOf = function (n) { const t = lookupTdocTableInIndex_(index, n); return t ? provisionalStatusOfTable_(t) : null; };
  const out = { successorUndecided: [], successorDecided: [], namedNotListed: [] };
  Object.keys(facts).forEach(function (n) {
    if (statusOf(n) === null) return;
    const s = provisionalSuccessorsOf_(facts, n);
    s.listed.forEach(function (m) { const st = statusOf(m); if (st === null) return; (PROVREV_UNDECIDED_.indexOf(st) !== -1 ? out.successorUndecided : out.successorDecided).push(n + ' -> ' + m + ' (' + st + ')'); });
    if (s.named.length) out.namedNotListed.push(n + ' -> ' + s.named.join(', '));
  });
  Logger.log('[IT28] candidates ' + JSON.stringify(out, null, 1));
  return out;
}

// ============================================================== the runner (2026-11-25)
// So that nobody has to run editor functions by hand at every stage: it28InstallRunner() -- once per DISPOSABLE copy,
// from the editor -- installs a time trigger IN THIS COPY ONLY that runs it28RunnerTick() every 5 minutes. A tick
// does nothing unless the pushed job file (Issue28Job.js: var IT28_JOB_ = { id, steps }) names a job not done yet;
// then it runs the steps (a fixed list: snapshot, compare, timed update, configure the meeting, sessions, phase B and C
// checks), saving after each step (a long job goes
// on in the next tick), and writes the result as the description of Drive files "IT28 result <job> <i>/<n>" in the
// snapshot folder, where tools/integration-test/run-issue28-validation.mjs reads it (Drive metadata only). It never
// throws (an error is a result); a failed step that changes something stops the job. it28Cleanup_() -- at a job step
// 'uninstall' and when IT28_RUNNER_HOURS_ hours are over -- removes every trigger of the project and every document
// and script property but the snapshot index (the copies had none), and reports what is left. It refuses to run in
// anything but a DISPOSABLE copy. No web app, no API executable, no new permission.

var IT28_RUNNER_HANDLER_ = 'it28RunnerTick';
var IT28_RUNNER_KEY_ = 'IT28_RUNNER';
var IT28_RUNNER_HOURS_ = 72;
var IT28_RUNNER_TICK_MS_ = 4 * 60 * 1000;
var IT28_RESULT_CHUNK_ = 3000;

function it28RunnerTriggers_() {
  return ScriptApp.getProjectTriggers().filter(function (t) { return t.getHandlerFunction() === IT28_RUNNER_HANDLER_; });
}

/** Editor, once per copy: installs the runner (a 5-minute trigger in this copy only, for IT28_RUNNER_HOURS_ hours). */
function it28InstallRunner() {
  const doc = it28Document_();
  it28RunnerTriggers_().forEach(function (t) { ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger(IT28_RUNNER_HANDLER_).timeBased().everyMinutes(5).create();
  const state = { installedAt: new Date().toISOString(), deadline: Date.now() + IT28_RUNNER_HOURS_ * 3600000, done: {}, current: null };
  PropertiesService.getDocumentProperties().setProperty(IT28_RUNNER_KEY_, JSON.stringify(state));
  it28WriteResult_('runner', { event: 'installed', at: state.installedAt, document: doc.getName(), release: it28Release_(), deadline: new Date(state.deadline).toISOString(), triggers: it28RunnerTriggers_().length });
  Logger.log('[IT28] runner installed in this copy: every 5 minutes until ' + new Date(state.deadline).toISOString() + '. Nothing else to do here.');
}

function it28Release_() {
  try { return SA4_RELEASE_.releaseId + ' / ' + SA4_RELEASE_.codeVersion + (SA4_RELEASE_.gitCommit ? ' @' + String(SA4_RELEASE_.gitCommit).slice(0, 7) : ''); } catch (e) { return '?'; }
}

/** The result of a job as Drive files: the summary JSON split over file descriptions (readable with Drive metadata). Replaced in the tests. */
function it28WriteResult_(jobId, result) {
  if (!result.document) result.document = DocumentApp.getActiveDocument().getName();
  const text = JSON.stringify(result);
  const n = Math.max(1, Math.ceil(text.length / IT28_RESULT_CHUNK_));
  const folder = it28Folder_();
  for (let i = 0; i < n; i++) {
    const f = folder.createFile(Utilities.newBlob(i === 0 ? JSON.stringify(result, null, 1) : '', 'text/plain', 'IT28 result ' + jobId + ' ' + (i + 1) + '/' + n + '.txt'));
    f.setDescription(text.slice(i * IT28_RESULT_CHUNK_, (i + 1) * IT28_RESULT_CHUNK_));
  }
}

/** Pure: one compare as a compact result (identical, how many lines, the first ones). */
function it28CompareResult_(a, b, d) {
  const lines = it28DiffLines_(d);
  // Counters by kind, for rules checked without reading the lines (e.g. no manual minutes changed).
  const rows = {};
  d.changes.forEach(function (c) {
    const m = /row "([^"#]+?)(?: #\d+)?"$/.exec(c.where);
    const k = (m ? m[1] : '?') + ' / ' + c.field;
    rows[k] = (rows[k] || 0) + 1;
  });
  const n = function (s) { return s ? { changed: s.changed.length, added: s.added.length, removed: s.removed.length, moved: s.moved.length } : null; };
  return { compare: a + ' -> ' + b, identical: d.identical, release: d.release, lines: lines.length, first: lines.slice(0, 40).map(function (l) { return l.slice(0, 240); }),
    counters: { tabs: d.tabs ? { added: d.tabs.added.length, removed: d.tabs.removed.length, renamed: d.tabs.renamed.length, moved: d.tabs.moved.length } : null,
      tdocsAdded: d.tdocsAdded.length, tdocsRemoved: d.tdocsRemoved.length, titles: d.titles.length, rows: rows, headings: d.headings ? { added: d.headings.added.length, removed: d.headings.removed.length } : null,
      summary: n(d.summary), other: n(d.other), counts: d.counts } };
}

/** Steps that change something: when one fails, the rest of the job is not run. */
var IT28_CHANGING_STEPS_ = ['update', 'configureMeeting', 'saveSessions', 'setProps', 'expectRevision', 'resolveRevision', 'sessionRuleSet', 'carrySet', 'uninstall'];

/**
 * Removes what the tests left in this copy: every trigger of the project, every document property but the snapshot
 * index (IT28_SNAP_*: the evidence), every script property. The copies had none of them before (preflight
 * 2026-11-25). Returns what is left, as the proof of the cleanup.
 */
function it28Cleanup_() {
  ScriptApp.getProjectTriggers().forEach(function (t) { ScriptApp.deleteTrigger(t); });
  const docProps = PropertiesService.getDocumentProperties();
  Object.keys(docProps.getProperties()).forEach(function (k) { if (k.indexOf(IT28_SNAP_KEY_) !== 0) docProps.deleteProperty(k); });
  const scriptProps = PropertiesService.getScriptProperties();
  Object.keys(scriptProps.getProperties()).forEach(function (k) { scriptProps.deleteProperty(k); });
  const left = Object.keys(docProps.getProperties()).filter(function (k) { return k.indexOf(IT28_SNAP_KEY_) !== 0; });
  return { triggersLeft: ScriptApp.getProjectTriggers().length, documentPropertiesLeft: left, scriptPropertiesLeft: Object.keys(scriptProps.getProperties()),
    clean: !ScriptApp.getProjectTriggers().length && !left.length && !Object.keys(scriptProps.getProperties()).length };
}

/** Pure: the counts of a Close Report result, and whether a TDoc is named among the blocking findings or the information. */
function it28CloseReportSummary_(result, tdoc) {
  const blocking = JSON.stringify([result.problems || [], result.blockingRevisions || [], result.unresolved || []]);
  const information = JSON.stringify(result.informationNotes || []);
  return { ready: !!result.ready, problems: (result.problems || []).length, blockingRevisions: (result.blockingRevisions || []).length, unresolved: (result.unresolved || []).length,
    information: (result.informationNotes || []).length, tdoc: tdoc || '', tdocBlocking: !!tdoc && blocking.indexOf(tdoc) !== -1, tdocInformation: !!tdoc && information.indexOf(tdoc) !== -1,
    informationFirst: (result.informationNotes || []).slice(0, 3).map(function (x) { return String(typeof x === 'string' ? x : JSON.stringify(x)).slice(0, 200); }) };
}

/** One step of a job; returns its result (an error is a result). */
function it28RunStep_(step, state) {
  const started = Date.now();
  const out = { do: step.do };
  try {
    const props = PropertiesService.getDocumentProperties();
    if (step.do === 'snap') {
      const labels = it28SnapLabels_(props.getProperties());
      const name = 's' + (labels.length ? Number(labels[labels.length - 1].slice(1)) + 1 : 1);
      it28Fingerprint(name);
      const meta = JSON.parse(props.getProperty(IT28_SNAP_KEY_ + name));
      out.label = name; out.contentSha = String(meta.contentSha).slice(0, 16); out.tabs = meta.tabs; out.tdocTables = meta.tdocTables; out.tables = meta.tables;
      out.paragraphs = meta.paragraphs; out.links = meta.links; out.release = meta.release;
      state.lastSnaps = (state.lastSnaps || []).concat([name]).slice(-20);
    } else if (step.do === 'compare') {
      const snaps = state.lastSnaps || [];
      const b = step.to || snaps[snaps.length - 1];
      const a = step.from === 'baseline' ? it28SnapLabels_(props.getProperties())[0] : (step.from === 'previous' || !step.from) ? snaps[snaps.length - 2] : step.from;
      if (!a || !b) throw new Error('nothing to compare (' + a + ' -> ' + b + ')');
      Object.assign(out, it28CompareResult_(a, b, it28DiffOf_(it28LoadSnapshot_(a), it28LoadSnapshot_(b))));
    } else if (step.do === 'update') {
      // As continuousUpdate() does (its document lock, its core), but the core's result is read: continuousUpdate()
      // drops it, and a refused or failed update returns { success: false, notUpdated } without throwing (found
      // 2026-11-25: every update of the AHG copy took under 2 s and was taken for a pass).
      it28Document_();
      const lock = LockService.getDocumentLock();
      if (!lock.tryLock(30000)) throw new Error('another update holds the document lock: not run');
      const t0 = Date.now();
      let r;
      try { r = continuousUpdateCore_(); } finally { lock.releaseLock(); }
      out.updateMs = Date.now() - t0;
      const log = String(Logger.getLog ? Logger.getLog() : '');
      const downloaded = /Downloaded (\d+) TDOCs/.exec(log);
      const total = /TOTAL continuousUpdate[^\d]*(\d+) ?ms/.exec(log);
      out.tdocsDownloaded = downloaded ? Number(downloaded[1]) : null;
      out.perfTotalMs = total ? Number(total[1]) : null;
      out.result = r ? { success: r.success !== false, notUpdated: !!r.notUpdated, error: r.error ? String(r.error).slice(0, 600) : '', incomplete: r.incomplete || null } : null;
      if (!r || r.success === false || r.notUpdated) throw new Error('the update did not run' + (r && r.notUpdated ? ' (nothing changed)' : '') + ': ' + ((r && r.error) || 'no result'));
    } else if (step.do === 'probe') {
      const timing = {};
      const snap = it28ReadSnapshot_(it28Document_().getId(), timing);
      out.timing = timing;
      out.tabs = snap.tabs.map(function (x) {
        const c = x.content;
        return { path: x.path, id: x.id, title: x.title, parent: x.parent, tdocTables: c.counts.tdocTables, tables: c.counts.tables, paragraphs: c.counts.paragraphs, summary: c.summary.length > 0,
          headings: c.headings.slice(0, 6).map(function (h) { return h.slice(0, 90); }), firstText: c.other.filter(function (e) { return e.k === 'p' && e.x; }).slice(0, 4).map(function (e) { return e.x.slice(0, 120); }) };
      });
    } else if (step.do === 'props') {
      const all = props.getProperties();
      out.names = Object.keys(all).filter(function (k) { return k.indexOf('IT28_') !== 0; }).sort();
      out.values = {};
      (step.show || ['MEETING_TYPE', 'MEETING_NUMBER', 'MEETING_FOLDER', 'MEETING_ID', 'REPORT_SUFFIX', 'TDOC_LIST_URL', 'ADHOC_SESSION_RULE', 'ADHOC_CARRY_FORWARD_MODE', 'DISPOSITION_CLASS_HEADINGS'])
        .forEach(function (k) { if (all[k] !== undefined) out.values[k] = String(all[k]).slice(0, 200); });
    } else if (step.do === 'configure') {
      if (!step.config || typeof step.config !== 'object') throw new Error('configure needs a config');
      saveConfigurationSettings(step.config);
      out.saved = Object.keys(step.config).sort();
    } else if (step.do === 'setProps') {
      // Test settings of this copy only (e.g. DISPOSITION_CLASS_HEADINGS); never the harness's own IT28_* keys.
      Object.keys(step.props || {}).forEach(function (k) { if (k.indexOf('IT28_') === 0) throw new Error('setProps may not touch ' + k); });
      Object.keys(step.props || {}).forEach(function (k) { if (step.props[k] === null) props.deleteProperty(k); else props.setProperty(k, String(step.props[k])); });
      out.set = Object.keys(step.props || {}).sort();
    } else if (step.do === 'configureMeeting') {
      // As Configure Meeting does: Resolve (and Discover Agenda / TDocs for an ad-hoc meeting) from the 3GPP Portal, then Save.
      const resolved = resolveMeetingForConfigDialog_(String(step.meetingId));
      if (!resolved.ok) throw new Error('Resolve ' + step.meetingId + ': ' + resolved.error);
      let preview = resolved.preview;
      if (step.expectType === 'adhoc') {
        const found = discoverAgendaForConfigDialog_(String(step.meetingId), resolved.resolved);
        if (!found.ok) throw new Error('Discover ' + step.meetingId + ': ' + found.error);
        preview = found.preview;
      }
      const v = function (f) { return preview[f] && preview[f].value ? String(preview[f].value) : ''; };
      if (v('meetingType') !== step.expectType) throw new Error('meeting ' + step.meetingId + ' is "' + v('meetingType') + '", not ' + step.expectType + ' -- nothing saved');
      const config = Object.assign({ meetingId: v('meetingId'), meetingType: v('meetingType'), meetingName: v('meetingName'), meetingDate: v('meetingDate'), ftpBase: v('ftpBase'),
        agendaTdoc: v('agendaTdoc'), agendaCsvUrl: v('agendaCsvUrl'), tdocUrl: v('tdocListUrl'), revisionsUrl: v('revisionsUrl'), reportType: step.reportType,
        meetingFolder: '', meetingNumber: '', mailingList: '', mailingListMode: 'derived', agendaSourceDocId: '', showPreview: false }, step.overrides || {});
      saveConfigurationSettings(config);
      out.config = { meetingId: config.meetingId, meetingType: config.meetingType, meetingName: config.meetingName, meetingDate: config.meetingDate, ftpBase: config.ftpBase, tdocUrl: config.tdocUrl, agendaTdoc: config.agendaTdoc, reportType: config.reportType };
    } else if (step.do === 'saveSessions') {
      const r = saveAdhocSessionsConfiguration(step.rows);
      if (r && r.ok === false) throw new Error('sessions not saved: ' + (r.errors || []).join('; '));
      out.sessions = (step.rows || []).map(function (s) { return s.label + ' ' + s.date + (s.endDate ? '..' + s.endDate : '') + ' ' + s.start + '-' + s.end; });
    } else if (step.do === 'candidates') {
      out.candidates = it28Candidates();
      state.candidates = out.candidates;
    } else if (step.do === 'expectRevision') {
      // Phase B: "revision expected" for a TDoc of this report whose successor is never registered.
      const named = ((state.candidates || {}).namedNotListed || []).map(function (x) { return x.split(' -> ')[0]; });
      let original = named[0] || '';
      if (!original) {
        const index = buildTdocTableIndex_(it28Document_().getBody().getTables());
        const all = flattenTdocGroups_(downloadAndGroupTdocs_(getReportConfig_(), null));
        const facts = plenaryListFactsOf_(all);
        original = Object.keys(facts).sort().filter(function (n) { const tb = lookupTdocTableInIndex_(index, n); return tb && PROVREV_UNDECIDED_.indexOf(provisionalStatusOfTable_(tb)) !== -1; })[0] || '';
      }
      if (!original) throw new Error('no TDoc of this report is undecided: nothing to record');
      out.message = saveExpectedRevision({ original: original, decision: step.decision || 'noted', plenary: '', agenda: '', successor: '' });
      out.original = original;
      state.revisionOriginal = original;
    } else if (step.do === 'resolveRevision') {
      if (!state.revisionOriginal) throw new Error('no expected revision recorded in this job');
      out.message = changeExpectedRevision('resolve', state.revisionOriginal, { outcome: step.outcome || 'not-submitted', reason: step.reason || 'Integration test: the successor was never registered.' });
      out.original = state.revisionOriginal;
    } else if (step.do === 'closeReportCheck') {
      // Close Report's checks only (its core, under its lock); not the dialog, not the writes after it.
      Object.assign(out, it28CloseReportSummary_(withDispositionLock_('Close Report', closeReportCore_), state.revisionOriginal));
    } else if (step.do === 'meetingControl') {
      const data = meetingControlData();
      out.meetingControl = JSON.stringify(data).slice(0, 1500);
      out.errors = (data && data.errors) || [];
    } else if (step.do === 'sessionRulePreview') {
      const r = previewAdhocSessionRule();
      out.preview = JSON.stringify(r).slice(0, 1200);
      out.previewOk = !!(r && r.ok);
      state.sessionRulePreviewOk = out.previewOk;
    } else if (step.do === 'sessionRuleSet') {
      if (step.rule !== 'upload' && !state.sessionRulePreviewOk) { out.skipped = 'the preview refused the rule (see the step before): not switched'; }
      else out.message = setAdhocSessionRule(step.rule);
    } else if (step.do === 'carryPreview') {
      const r = previewAdhocCarryForward();
      out.preview = JSON.stringify(r).slice(0, 1200);
      out.previewOk = !!(r && r.ok);
      state.carryPreviewOk = out.previewOk;
    } else if (step.do === 'carrySet') {
      if (step.on && !state.carryPreviewOk) { out.skipped = 'the preview refused carry-forward: not switched'; }
      else out.message = setAdhocCarryForward(!!step.on);
    } else if (step.do === 'uninstall') {
      Object.assign(out, it28Cleanup_());
      state.uninstalled = true;
      if (!out.clean) throw new Error('cleanup incomplete: ' + JSON.stringify({ triggers: out.triggersLeft, documentProperties: out.documentPropertiesLeft, scriptProperties: out.scriptPropertiesLeft }));
    } else {
      throw new Error('unknown step ' + step.do);
    }
    out.ok = true;
  } catch (e) {
    out.ok = false;
    out.error = String((e && e.message) || e).slice(0, 600);
  }
  out.ms = Date.now() - started;
  return out;
}

/** The trigger: runs the pushed job if it is new; never throws. */
function it28RunnerTick() {
  let lock = null;
  let jobId = '?';
  try {
    it28Document_();
    // The script lock, not the document lock: the report's update takes the document lock itself.
    lock = LockService.getScriptLock();
    if (!lock.tryLock(1000)) { lock = null; return; }
    const props = PropertiesService.getDocumentProperties();
    const state = JSON.parse(props.getProperty(IT28_RUNNER_KEY_) || 'null');
    if (!state) return;
    if (Date.now() > state.deadline) {
      it28WriteResult_('runner', { event: 'expired', at: new Date().toISOString(), cleanup: it28Cleanup_() });
      return;
    }
    const job = typeof IT28_JOB_ !== 'undefined' ? IT28_JOB_ : null;
    if (!job || !job.id || state.done[job.id]) return;
    jobId = job.id;
    if (!state.current || state.current.id !== job.id) state.current = { id: job.id, next: 0, results: [], startedAt: new Date().toISOString(), release: it28Release_() };
    const started = Date.now();
    // At least one step per tick (progress), then more while the tick has time.
    do {
      const step = job.steps[state.current.next];
      const r = it28RunStep_(step, state);
      state.current.results.push(r);
      state.current.next++;
      // A failed step that changes something stops the job: the rest is not run (the orchestrator cleans up).
      if (!r.ok && IT28_CHANGING_STEPS_.indexOf(step.do) !== -1 && step.do !== 'uninstall') { state.current.stoppedAt = state.current.next; state.current.next = job.steps.length; }
      if (state.uninstalled) break;
      props.setProperty(IT28_RUNNER_KEY_, JSON.stringify(state));
    } while (state.current.next < job.steps.length && Date.now() - started < IT28_RUNNER_TICK_MS_);
    if (state.current.next < job.steps.length) return;
    const result = { job: job.id, release: state.current.release, startedAt: state.current.startedAt, finishedAt: new Date().toISOString(),
      ok: state.current.results.every(function (r) { return r.ok; }) && !state.current.stoppedAt, stoppedAt: state.current.stoppedAt || null, steps: state.current.results };
    if (state.uninstalled) {
      // it28Cleanup_ removed the runner's state with everything else; nothing is written back.
      it28WriteResult_(job.id, result);
      const u = result.steps[result.steps.length - 1];
      it28WriteResult_('runner', { event: u && u.clean ? 'uninstalled' : 'uninstall incomplete', at: result.finishedAt, job: job.id, cleanup: u });
      return;
    }
    state.done[job.id] = true;
    state.current = null;
    props.setProperty(IT28_RUNNER_KEY_, JSON.stringify(state));
    it28WriteResult_(job.id, result);
  } catch (e) {
    try { it28WriteResult_(jobId + ' ERROR', { job: jobId, error: String((e && e.message) || e).slice(0, 600), at: new Date().toISOString() }); } catch (e2) { /* nothing more can be done */ }
  } finally {
    if (lock) lock.releaseLock();
  }
}
