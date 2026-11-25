/**
 * Issue #28: tools/integration-test/Issue28Check.js (disposable copies only).
 * Fingerprint v2: deterministic over repeated reads (key order, run splitting), compared by TDoc number and row
 * label; every real change of text, links, formatting, chips, headings, summary or other content is reported with
 * its place and old/new values; stored snapshots are verified (file, SHA-256, size, complete). Refusals, preflight.
 */
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const { loadCode } = require('./helpers/load-code');

let failures = 0;
function check(name, actual, expected) {
  const ok = JSON.stringify(actual) === JSON.stringify(expected);
  if (!ok) failures++;
  console.log((ok ? '  ok   ' : '  FAIL ') + name + (ok ? '' : '\n         expected ' + JSON.stringify(expected) + '\n         actual   ' + JSON.stringify(actual)));
}

const { sandbox: S } = loadCode();
// Apps Script's computeDigest: signed bytes.
S.Utilities.DigestAlgorithm = { SHA_256: 'sha256' }; S.Utilities.Charset = { UTF_8: 'utf8' };
S.Utilities.computeDigest = (alg, text) => Array.from(require('crypto').createHash(alg).update(String(text), 'utf8').digest()).map((b) => (b > 127 ? b - 256 : b));
S.console = { log: () => {} };
vm.runInContext(fs.readFileSync(path.join(__dirname, '..', 'tools', 'integration-test', 'Issue28Check.js'), 'utf8'), S, { filename: 'Issue28Check.js' });
// This repository is public: no real ID here. Synthetic IDs, their digests added to the tables of the sandbox.
const R = require('../tools/integration-test/issue28-copy-rules.js');
const realTable = Object.keys(S.IT28_REFUSED_DOCUMENT_DIGESTS_).sort();
const FAKE = { prod: '1FakeProductionReport6G000000000000000000x1', legacy: '1FakeLegacyAhReport0000000000000000000000x2', own: '1FakeDisposableCopy00000000000000000000000x3',
  template: '1FakeTemplateDocument0000000000000000000000x4', agenda: '1FakeUnknownAgendaDocument000000000000000x5' };
S.IT28_REFUSED_DOCUMENT_DIGESTS_[R.idDigest(FAKE.prod)] = '6G Media SA4#138';
S.IT28_REFUSED_DOCUMENT_DIGESTS_[R.idDigest(FAKE.legacy)] = 'the Legacy 6G AH report';
S.IT28_TEMPLATE_DOCUMENT_DIGEST_ = R.idDigest(FAKE.template);

// ---------------------------------------------------------------- a Docs API document (as the REST API returns it)
const run = (text, style) => ({ textRun: { content: text, textStyle: style || {} } });
const para = (elements, named) => ({ paragraph: { elements: elements, paragraphStyle: { namedStyleType: named || 'NORMAL_TEXT' } } });
const cell = (...paras) => ({ content: paras });
const row = (label, ...values) => ({ tableCells: [cell(para([run(label + '\n')]))].concat(values) });
const RED = { color: { rgbColor: { red: 1 } } };
const GREEN = { color: { rgbColor: { green: 1 } } };
const tdoc = (n, o) => ({ table: { rows: 5, columns: 2, tableRows: [
  row('TDoc', cell(para([run(n + '\n', o.link === null ? {} : { link: { url: o.link || 'https://x/' + n + '.zip' }, foregroundColor: { color: { rgbColor: { blue: 1 } } } })]))),
  row('Title', cell(para([run((o.title || 'A title') + '\n')]))),
  row('Revisions', cell(para([run('none\n', { foregroundColor: GREEN })]))),
  row('Discussion', cell(para(o.minutesRuns || [run((o.minutes || 'Typed minutes.') + '\n', o.minutesStyle || {})]))),
  row('Status', cell(para(o.chip ? [{ dropdown: { dropdownProperties: { selectedOption: o.chip } } }, run('\n')] : [run((o.status || 'agreed') + '\n', { foregroundColor: o.statusColor || RED })])))
].concat(o.extraRows || []) } });
const doc = (o) => {
  o = o || {};
  const tdocs = o.order || ['S4-261001', 'S4-261002'];
  return { revisionId: 'r1', documentId: 'd', tabs: [{ tabProperties: { tabId: 't.0' }, documentTab: { body: { content: [
    { sectionBreak: { sectionStyle: { sectionType: 'CONTINUOUS' } } },
    para([run((o.heading || '7 Audio') + '\n')], 'HEADING_1'),
  ].concat(tdocs.map((n) => tdoc(n, (o.t || {})[n] || {})), [
    para([run(o.other || 'Typed text.\n')]),
  ], o.insert ? [para([run(o.insert)])] : [], [
    para([run('Disposition Summary\n')], 'HEADING_1'),
    para([run('Agreed, Endorsed and Approved Documents (2)\n', o.h2 ? {} : { bold: true })], o.h2 ? 'HEADING_2' : 'NORMAL_TEXT'),
    { table: { rows: 2, columns: 2, tableRows: [
      { tableCells: [cell(para([run('TDoc\n')])), cell(para([run('Title\n')]))] },
      { tableCells: [cell(para([run('S4-261001\n')])), cell(para([run((o.summaryTitle || 'A title') + '\n')]))] }] } }
  ]) } } }] };
};
/** The same document with every object's keys in reverse order (another serialization of the same content). */
const reversedKeys = (v) => (Array.isArray(v) ? v.map(reversedKeys) : v && typeof v === 'object' ? Object.keys(v).reverse().reduce((o, k) => { o[k] = reversedKeys(v[k]); return o; }, {}) : v);
/** The same document with each text run split in two runs of the same style. */
const splitRuns = (v) => {
  if (Array.isArray(v)) return v.map(splitRuns);
  if (!v || typeof v !== 'object') return v;
  const o = {};
  Object.keys(v).forEach((k) => {
    if (k === 'elements') o[k] = [].concat(...v[k].map((e) => (e.textRun && e.textRun.content.length > 3 ? [{ textRun: Object.assign({}, e.textRun, { content: e.textRun.content.slice(0, 2) }) }, { textRun: Object.assign({}, e.textRun, { content: e.textRun.content.slice(2) }) }] : [splitRuns(e)])));
    else o[k] = splitRuns(v[k]);
  });
  return o;
};
const snapOf = (d) => S.it28SnapshotOf_(JSON.parse(JSON.stringify(d)));
const diff = (a, b) => S.it28DiffOf_(snapOf(doc(a)), snapOf(doc(b)));
const lines = (a, b) => S.it28DiffLines_(diff(a, b));

console.log('fingerprint v2: repeated reads are identical');
const base = snapOf(doc());
check('two reads of the same document: the same snapshot text, byte for byte', JSON.stringify(snapOf(doc())), JSON.stringify(base));
check('... with every key order reversed (the API does not promise an order): the same', JSON.stringify(snapOf(reversedKeys(doc()))), JSON.stringify(base));
check('... with the runs split differently: the same', JSON.stringify(snapOf(splitRuns(doc()))), JSON.stringify(base));
const same = diff({}, {});
check('nothing changed: identical, no summary change, nothing outside reported, one line saying so', [same.identical, same.summary, same.other, same.changes, S.it28DiffLines_(same)], [true, null, null, [], ['release ? -> ?: IDENTICAL (no difference)']]);
check('counts and classification: 1 tab, 2 TDoc tables (the summary table "TDoc | Title" is summary), 3 tables, 1 section break', [base.counts.tabs, base.counts.tdocTables, base.counts.tables, base.counts.sectionBreaks, base.tabs[0].content.tdocs.map((t) => t.n), base.tabs[0].content.summary.length], [1, 2, 3, 1, ['S4-261001', 'S4-261002'], 4]);

console.log('fingerprint v2: real changes, located, with old and new values');
const c1 = diff({}, { t: { 'S4-261001': { minutes: 'Typed minutes, changed.' } } });
check('minutes changed: TDoc and row named, old and new text', [c1.identical, c1.changes], [false, [{ where: 'TDoc S4-261001, row "Discussion"', field: 'text', old: 'Typed minutes.', new: 'Typed minutes, changed.' }]]);
const c2 = diff({}, { t: { 'S4-261002': { link: 'https://y/S4-261002.zip' } } });
check('a link changed: old and new URL', c2.changes, [{ where: 'TDoc S4-261002, row "TDoc"', field: 'links', old: 'https://x/S4-261002.zip @ S4-261002', new: 'https://y/S4-261002.zip @ S4-261002' }]);
const c3 = diff({}, { t: { 'S4-261001': { link: null } } });
check('a link lost: links and formatting (its colour) of the row, one link less', [c3.changes.map((c) => c.field + ': ' + c.old + ' -> ' + c.new), c3.counts], [['links: https://x/S4-261001.zip @ S4-261001 -> (none)', 'formatting: color #0000ff -> (plain)'], { links: '2 -> 1' }]);
const c4 = diff({}, { t: { 'S4-261001': { minutesStyle: { bold: true } } } });
check('minutes made bold: a formatting change, old and new styles', c4.changes, [{ where: 'TDoc S4-261001, row "Discussion"', field: 'formatting', old: '(plain)', new: 'bold' }]);
const runsA = [run('First', { bold: true }), run(' second\n')];
const runsB = [run('First '), run('second', { bold: true }), run('\n')];
const c5 = S.it28DiffOf_(snapOf(doc({ t: { 'S4-261001': { minutesRuns: runsA } } })), snapOf(doc({ t: { 'S4-261001': { minutesRuns: runsB } } })));
check('bold moved to another word (same text, same styles): still a formatting change, said so', [c5.changes.length, c5.changes[0].field, /same styles at other places/.test(c5.changes[0].new)], [1, 'formatting', true]);
const c6 = diff({}, { t: { 'S4-261002': { statusColor: GREEN } } });
check('the colour of a Status changed: old and new colour', c6.changes.map((c) => c.where + ' ' + c.field + ': ' + c.old + ' -> ' + c.new), ['TDoc S4-261002, row "Status" formatting: color #ff0000 -> color #00ff00']);
const c7 = diff({ t: { 'S4-261001': { chip: 'agreed' } } }, { t: { 'S4-261001': { chip: 'noted' } } });
check('a Status dropdown chip changed: shown by its text (v1 saw only "dropdown")', [c7.changes.length, c7.changes[0].field, /^\[dropdown: agreed #[0-9a-f]{6}\]$/.test(c7.changes[0].old), /^\[dropdown: noted #[0-9a-f]{6}\]$/.test(c7.changes[0].new)], [1, 'text', true, true]);
const c7b = diff({ t: { 'S4-261001': { chip: 'agreed' } } }, { t: { 'S4-261001': { chip: 'agreed' } } });
check('... the same chip twice: identical', c7b.identical, true);
const c8 = diff({}, { t: { 'S4-261001': { title: 'A new title' } } });
check('a title changed: in titles and as a row change', [c8.titles, c8.changes.length], [[{ tdoc: 'S4-261001', old: 'A title', new: 'A new title' }], 1]);
const c9 = diff({}, { heading: '7 Audio and speech' });
check('a heading changed: heading removed/added; the text under it is located under the new heading', [c9.headings.removed, c9.headings.added, c9.changes], [['#1 H1 7 Audio'], ['#1 H1 7 Audio and speech'], []]);
const c10 = diff({}, { h2: true });
check('the class line restyled (bold Normal -> Heading 2): a summary change with old and new style, a heading added', [c10.summary && c10.summary.changed.map((c) => c.field + ': ' + c.old + ' -> ' + c.new), c10.headings.added], [['formatting: bold -> style HEADING_2'], ['#3 H2 Agreed, Endorsed and Approved Documents (2)']]);
const c11 = diff({}, { summaryTitle: 'Another title' });
check('a summary table cell changed: located in the summary, old and new row', c11.summary.changed.map((c) => c.field + ': ' + c.old + ' -> ' + c.new), ['text: S4-261001 || A title -> S4-261001 || Another title']);
const c12 = diff({}, { insert: 'A paragraph typed by hand.\n' });
check('a paragraph added outside the TDoc tables: added, with its heading and text', [c12.other.added, c12.other.removed, c12.other.changed], [[{ at: '7 Audio', x: 'A paragraph typed by hand.' }], [], []]);
const c13 = diff({}, { other: 'Typed text, changed.\n' });
check('a paragraph changed outside the TDoc tables: one change with old and new', c13.other.changed.map((c) => c.field + ': ' + c.old + ' -> ' + c.new), ['text: Typed text. -> Typed text, changed.']);
const c14 = diff({}, { order: ['S4-261001'] });
const c15 = diff({ order: ['S4-261001'] }, { order: ['S4-261001', 'S4-261002', 'S4-261003'] });
check('a TDoc table removed, two added (located)', [c14.tdocsRemoved, c15.tdocsAdded], [['S4-261002 (was under "7 Audio")'], ['S4-261002 (under "7 Audio")', 'S4-261003 (under "7 Audio")']]);
const c16 = diff({}, { order: ['S4-261002', 'S4-261001'] });
check('two TDoc tables swapped: no row change (compared by number), the order reported', [c16.changes, c16.tdocOrder, c16.identical], [[], { firstDifferenceAt: 1, was: 'S4-261001', now: 'S4-261002' }, false]);
const c17 = diff({}, { order: ['S4-261001', 'S4-261002', 'S4-261001'] });
const c18 = diff({ order: ['S4-261001', 'S4-261002', 'S4-261001'] }, { order: ['S4-261001', 'S4-261002', 'S4-261001'], t: { 'S4-261001': { minutes: 'x' } } });
check('a TDoc table there twice: reported; a change in it located by table', [c17.duplicates, c18.changes.map((c) => c.where)], [['S4-261001: 1 table(s) -> 2'], ['TDoc S4-261001 (table 1), row "Discussion"', 'TDoc S4-261001 (table 2), row "Discussion"']]);
check('... the same duplicate in both: not a difference', diff({ order: ['S4-261001', 'S4-261001'] }, { order: ['S4-261001', 'S4-261001'] }).identical, true);
const long = 'x'.repeat(900);
const c19 = diff({ t: { 'S4-261001': { minutes: long + 'a' } } }, { t: { 'S4-261001': { minutes: long + 'b' } } });
check('a change after the first 600 characters (the excerpt): still found by its digest, said so', [c19.changes.length, /differs after/.test(c19.changes[0].new)], [1, true]);
const twoRows = [row('Discussion', cell(para([run('Second block.\n')])))];
const c20 = diff({ t: { 'S4-261001': { extraRows: twoRows } } }, { t: { 'S4-261001': { extraRows: [row('Discussion', cell(para([run('Second block, changed.\n')])))] } } });
check('two rows with the same label: told apart ("#2"), the change in the second', c20.changes.map((c) => c.where), ['TDoc S4-261001, row "Discussion #2"']);
const c21 = diff({}, { t: { 'S4-261001': { minutesRuns: [{ textRun: { content: 'Typed minutes.\n', textStyle: {}, suggestedInsertionIds: ['s1'] } }] } } });
check('a suggested insertion: a formatting change ("suggestion")', c21.changes.map((c) => c.field + ': ' + c.new), ['formatting: suggestion']);
check('the diff as lines: a located line per change', lines({}, { t: { 'S4-261001': { minutes: 'New.' } } }), ['release ? -> ?', 'TDoc S4-261001, row "Discussion" -- text: "Typed minutes." -> "New."']);

console.log('fingerprint v2: the findings of the independent review (2026-11-25)');
{
  // 1. a paragraph moved outside the TDoc tables: reported as moved, never an unexplained "not identical"
  const free = (texts) => ({ revisionId: 'r', tabs: [{ documentTab: { body: { content: [para([run('7 Audio\n')], 'HEADING_1')].concat(texts.map((x) => para([run(x + '\n')]))) } } }] });
  const mv = S.it28DiffOf_(snapOf(free(['a', 'b', 'c'])), snapOf(free(['b', 'c', 'a'])));
  const mvLines = S.it28DiffLines_(mv);
  check('review 1: a paragraph moved: reported as moved with its old and new place', [mv.identical, mv.other.moved.length, mvLines.some((l) => /moved from #\d+ .* to #\d+ .*: "a"/.test(l))], [false, 1, true]);
  const mv2 = S.it28DiffLines_(S.it28DiffOf_(snapOf(free(['a', 'b', 'c', 'd'])), snapOf(free(['b', 'c', 'a', 'd2']))));
  check('review 1: a move and a change: both listed', [mv2.some((l) => /moved .*: "a"/.test(l)), mv2.some((l) => /"d" -> "d2"/.test(l))], [true, true]);
  // 2. the label cell of a TDoc row: its link and its text
  const withLabel = (url, label) => { const d = doc(); const rowStatus = d.tabs[0].documentTab.body.content[2].table.tableRows[4]; rowStatus.tableCells[0] = cell(para([run(label || 'Status', url ? { link: { url: url } } : {}), run('\n')])); return d; };
  const lb = S.it28DiffOf_(snapOf(withLabel('https://a')), snapOf(withLabel('https://b')));
  check('review 2: the link in a label cell changed: reported, old and new', lb.changes.map((x) => x.where + ' ' + x.field + ': ' + x.old + ' -> ' + x.new), ['TDoc S4-261001, row "Status" label cell (text or links): https://a @ Status -> https://b @ Status']);
  const lb2 = S.it28DiffOf_(snapOf(withLabel(null, 'Status')), snapOf(withLabel(null, 'Status  ')));
  check('review 2: the label text changed only by spaces (same row key): reported', [lb2.identical, lb2.changes.map((x) => x.field)], [false, ['label cell (text or links)']]);
  // 4. the style moved with a text change: the text change is shown with style markers
  const st = S.it28DiffOf_(snapOf(doc({ t: { 'S4-261001': { minutesRuns: [run('a '), run('b', { bold: true }), run(' c\n')] } } })), snapOf(doc({ t: { 'S4-261001': { minutesRuns: [run('a b '), run('c', { bold: true }), run(' d\n')] } } })));
  check('review 4: bold moved together with a text change: visible in the old and new text', st.changes.map((x) => x.field + ': ' + x.old + ' -> ' + x.new), ['text (with its styles): a «b»{bold} c -> a b «c»{bold} d']);
  // 5. contentSha: the revision (valid 24 h, per user) is not part of it
  const r1 = snapOf(doc()), r2 = snapOf(Object.assign(doc(), { revisionId: 'another' }));
  check('review 5: the same content under another revision ID: the same content SHA-256, identical', [S.it28ContentSha_(r1) === S.it28ContentSha_(r2), S.it28DiffOf_(r1, r2).identical], [true, true]);
  // 6. a bookmark link keeps its target; cells joined cannot be mistaken for text
  const bm = (id) => { const d = doc(); d.tabs[0].documentTab.body.content[4].paragraph.elements = [run('See here\n', { link: { bookmarkId: id } })]; return d; };
  check('review 6: a link to another bookmark: reported', S.it28DiffOf_(snapOf(bm('id.1')), snapOf(bm('id.2'))).other.changed.map((x) => x.field), ['links']);
  const cells2 = (a, b) => ({ revisionId: 'r', tabs: [{ documentTab: { body: { content: [{ table: { tableRows: [{ tableCells: [cell(para([run(a + '\n')])), cell(para([run(b + '\n')]))] }] } }] } } }] });
  check('review 6: "a || b" | "c" vs "a" | "b || c": a text change (not taken for formatting)', S.it28DiffOf_(snapOf(cells2('a || b', 'c')), snapOf(cells2('a', 'b || c'))).other.changed.map((x) => x.field), ['text']);
  // 7. a chip reads by its display value
  const chip = S.it28ParagraphUnit_({ elements: [{ dropdown: { dropdownId: 'kix.abc123', dropdownProperties: { displayValue: 'agreed', optionId: 'opt.1234567890' } } }] });
  check('review 7: a dropdown chip shown by its display value, not its IDs', /^\[dropdown: agreed #[0-9a-f]{6}\]$/.test(chip.text), true);
}

console.log('fingerprint v2: documents with tabs (the AHG copy has 4)');
// A document of tabs: each tab its own body; child tabs nested; IDs stable, titles shown.
const body = (o) => doc(o).tabs[0].documentTab;
const tabNode = (id, title, o, children) => Object.assign({ tabProperties: { tabId: id, title: title }, documentTab: body(o) }, children ? { childTabs: children } : {});
const tabbed = (o) => {
  o = o || {};
  const s = o.s || {};
  const tabs = (o.order || ['t.a', 't.b']).map((id) => tabNode(id, (o.titles || {})[id] || { 't.a': 'Session 1', 't.b': 'Session 2', 't.c': 'Session 3' }[id], s[id] || {}, id === (o.parentOfChild || 't.b') && !o.noChild ? [tabNode('t.k', 'Notes', s['t.k'] || {})] : null));
  return { revisionId: o.rev || 'r1', tabs: tabs };
};
const tb = snapOf(tabbed());
check('tabs walked in document order with their IDs, titles, nesting paths and parents; counts summed', [tb.tabs.map((x) => x.path + ' ' + x.id + ' ' + x.title + ' ' + x.parent), tb.counts.tabs, tb.counts.tdocTables],
  [['1 t.a Session 1 ', '2 t.b Session 2 ', '2.1 t.k Notes t.b'], 3, 6]);
check('tabs: deterministic (keys reversed, runs split: the same snapshot)', [JSON.stringify(snapOf(reversedKeys(tabbed()))) === JSON.stringify(tb), JSON.stringify(snapOf(splitRuns(tabbed()))) === JSON.stringify(tb), S.it28DiffOf_(tb, snapOf(tabbed())).identical], [true, true, true]);
const tdiff = (a, b) => S.it28DiffOf_(snapOf(tabbed(a)), snapOf(tabbed(b)));
const tc1 = tdiff({}, { s: { 't.b': { t: { 'S4-261001': { minutes: 'Changed in session 2.' } } } } });
check('the same TDoc in two tabs: a change in one tab is located in that tab only', tc1.changes.map((x) => x.where + ' ' + x.field), ['tab "Session 2": TDoc S4-261001, row "Discussion" text']);
const tc2 = tdiff({}, { s: { 't.k': { t: { 'S4-261002': { link: 'https://other/x.zip', statusColor: GREEN } } } } });
check('a link and a colour changed in a child tab: located there', tc2.changes.map((x) => x.where + ' ' + x.field), ['tab "Notes": TDoc S4-261002, row "TDoc" links', 'tab "Notes": TDoc S4-261002, row "Status" formatting']);
const tc3 = tdiff({}, { titles: { 't.a': 'Session 1 (Monday)' } });
check('a tab renamed: reported by its ID, old and new title; its content compared as before', [tc3.tabs && tc3.tabs.renamed, tc3.changes, tc3.identical], [['tab t.a: "Session 1" -> "Session 1 (Monday)"'], [], false]);
const tc4 = tdiff({}, { order: ['t.a', 't.b', 't.c'] });
const tc5 = tdiff({ order: ['t.a', 't.b', 't.c'] }, { order: ['t.a', 't.c'] });
check('a tab added; a tab removed (with its child): reported with their TDoc tables', [tc4.tabs.added, tc5.tabs.removed, tc5.counts.tabs], [['tab 3 "Session 3" (t.c, 2 TDoc tables)'], ['tab 2 "Session 2" (t.b, 2 TDoc tables)', 'tab 2.1 "Notes" (t.k, 2 TDoc tables)'], '4 -> 2']);
const tc6 = tdiff({}, { order: ['t.b', 't.a'] });
check('two tabs swapped: both reported as moved; their content unchanged', [tc6.tabs.moved, tc6.changes], [['tab "Session 2" (t.b): 2 -> 1', 'tab "Session 1" (t.a): 1 -> 2'], []]);
const tc7 = tdiff({}, { parentOfChild: 't.a' });
check('a child tab moved under another tab: reported (path and parent)', tc7.tabs.moved, ['tab "Notes" (t.k): 2.1 under t.b -> 1.1 under t.a']);
const tc8 = tdiff({}, { s: { 't.a': { insert: 'Typed by hand in session 1.\n', heading: '7 Audio and speech' } } });
check('free text and a heading in one tab: located in that tab', [tc8.other.added.map((x) => x.at), tc8.headings.added], [['tab "Session 1": 7 Audio and speech'], ['tab "Session 1": #1 H1 7 Audio and speech']]);
check('the diff lines name the tab', S.it28DiffLines_(tc3).concat(S.it28DiffLines_(tc6)).filter((l) => /^tab /.test(l)), ['tab renamed: tab t.a: "Session 1" -> "Session 1 (Monday)"', 'tab moved: tab "Session 2" (t.b): 2 -> 1', 'tab moved: tab "Session 1" (t.a): 1 -> 2']);
check('a tab nested deeper than three levels: refused, not left out', (() => { const d = tabbed(); d.tabs[1].childTabs[0].childTabs = [tabNode('t.x', 'x', {}, [tabNode('t.y', 'y', {})])]; try { snapOf(d); return 'taken'; } catch (e) { return /nested deeper than the 3 levels/.test(e.message); } })(), true);
check('a tab without a body: refused', (() => { const d = tabbed(); delete d.tabs[0].documentTab; try { snapOf(d); return 'taken'; } catch (e) { return /tab 1 "Session 1" has no body content/.test(e.message); } })(), true);

console.log('fingerprint v2: the read of a document with tabs (structure, then one request per level)');
{
  // The Docs API as it answers a field mask: the structure request has no documentTab; a level request has the
  // documentTab of that level's tabs only.
  const answer = (full, url, revs) => {
    const fields = decodeURIComponent(url.split('fields=')[1]);
    const depth = fields.includes('documentTab') ? (fields.split('documentTab')[0].match(/childTabs\(/g) || []).length : -1;
    const cut = (tabs, d) => (tabs || []).map((x) => { const y = { tabProperties: x.tabProperties }; if (depth === d) y.documentTab = x.documentTab; if (x.childTabs && (depth === -1 || d < depth)) y.childTabs = cut(x.childTabs, d + 1); return y; });
    return { revisionId: revs.shift() || full.revisionId, tabs: cut(full.tabs, 0) };
  };
  const reader = (full, revs, tamper) => {
    const urls = [];
    S.ScriptApp.getOAuthToken = () => 'token';
    S.UrlFetchApp = { fetch: (url) => { urls.push(url); let r = JSON.parse(JSON.stringify(answer(full, url, revs))); if (tamper) r = tamper(r, urls.length); return { getResponseCode: () => 200, getContentText: () => JSON.stringify(r) }; } };
    return urls;
  };
  const full = tabbed();
  let urls = reader(full, []);
  const timing = {};
  const read = S.it28ReadSnapshot_(FAKE.own, timing);
  check('three tabs on two levels: three requests (structure, level 1, level 2), each under 2000 characters; the same snapshot as the whole document gives',
    [urls.length, urls.every((u) => u.length < 2000), !urls[0].includes('documentTab'), JSON.stringify(read) === JSON.stringify(snapOf(full)), timing.reads.map((r) => r.what)], [3, true, true, true, ['tab structure', 'tabs of level 1', 'tabs of level 2']]);
  urls = reader(full, ['r1', 'r1', 'r2']);
  check('the document changed between the requests (another revision): refused, nothing stored', (() => { try { S.it28ReadSnapshot_(FAKE.own, {}); return 'read'; } catch (e) { return /changed while it was read/.test(e.message); } })(), true);
  reader(full, [], (r, n) => { if (n === 2) delete r.tabs[1].documentTab.body; return r; });
  check('a tab returned without its body: refused', (() => { try { S.it28ReadSnapshot_(FAKE.own, {}); return 'read'; } catch (e) { return /tab 2 "Session 2" was returned without its body/.test(e.message); } })(), true);
  reader(full, [], (r, n) => { if (n === 2) r.tabs.pop(); return r; });
  check('a tab of the structure not returned at all: refused', (() => { try { S.it28ReadSnapshot_(FAKE.own, {}); return 'read'; } catch (e) { return /tab 2 "Session 2" \(t\.b\) could not be read completely/.test(e.message); } })(), true);
  reader(full, [], (r, n) => { if (n === 2) r.tabs.push({ tabProperties: { tabId: 't.new' }, documentTab: body({}) }); return r; });
  check('a tab returned that is not in the structure: refused', (() => { try { S.it28ReadSnapshot_(FAKE.own, {}); return 'read'; } catch (e) { return /not in the tab structure \(t\.new\)/.test(e.message); } })(), true);
  check('the request fields: structure 3 levels deep, no content; level 2 asks only child tabs for content', [S.it28StructureFields_().includes('documentTab'), (S.it28StructureFields_().match(/childTabs\(/g) || []).length, S.it28LevelFields_(1).startsWith('revisionId,tabs(childTabs(tabProperties(tabId),documentTab(')], [false, 3, true]);
}

console.log('refusals');
S.DocumentApp.getActiveDocument = () => ({ getId: () => FAKE.prod, getName: () => 'DISPOSABLE fake' });
check('refused in a protected report (6G Media SA4#138), whatever its name', (() => { try { S.it28Preflight(); return 'ran'; } catch (e) { return /disposable copies only; this is 6G Media SA4#138/.test(e.message); } })(), true);
S.DocumentApp.getActiveDocument = () => ({ getId: () => FAKE.own, getName: () => 'Audio report' });
check('refused in a document not named DISPOSABLE', (() => { try { S.it28Preflight(); return 'ran'; } catch (e) { return /does not start with "DISPOSABLE"/.test(e.message); } })(), true);
S.DocumentApp.getActiveDocument = () => ({ getId: () => FAKE.legacy, getName: () => 'DISPOSABLE fake' });
check('refused in the Legacy 6G AH report (known by digest only)', (() => { try { S.it28Preflight(); return 'ran'; } catch (e) { return /this is the Legacy 6G AH report/.test(e.message); } })(), true);
check('the refused documents of Issue28Check.js are those of the push tool (7 digests: template, five reports, Legacy 6G AH)', realTable, Object.keys(R.REFUSED_DOCUMENT_DIGESTS).filter((d) => d !== R.idDigest('1FakeProdDoc')).sort());
check('no real protected or copy ID in Issue28Check.js or this test (by digest; the repository is public)', R.plainIdsIn(fs.readFileSync(path.join(__dirname, '..', 'tools', 'integration-test', 'Issue28Check.js'), 'utf8') + fs.readFileSync(__filename, 'utf8')), []);

console.log('it28Preflight: triggers, properties, foreign IDs');
const OWN = FAKE.own;
const pre = (o) => {
  S.DocumentApp.getActiveDocument = () => ({ getId: () => OWN, getName: () => 'DISPOSABLE Issue28 - test' });
  S.ScriptApp.getProjectTriggers = () => (o.triggers || []).map((h) => ({ getHandlerFunction: () => h, getEventType: () => 'CLOCK' }));
  S.PropertiesService.getScriptProperties = () => ({ getProperties: () => o.script || {} });
  S.PropertiesService.getDocumentProperties = () => ({ getProperties: () => o.doc || {} });
  const log = S.Logger.log; S.Logger.log = () => {}; try { return S.it28Preflight(); } finally { S.Logger.log = log; }
};
const clean = pre({ doc: { MEETING_TYPE: 'main', REPORT_DOC_ID: OWN, STATUS_TEMPLATE_DOC_ID: FAKE.template, IT28_FP_x: FAKE.prod } });
check('isolated copy: no problem, nothing to review; its own ID and its snapshots ignored; the template only noted', [clean.problems, clean.review, clean.notes.length, /^isolated/.test(clean.verdict)], [[], [], 1, true]);
const bad = pre({ triggers: ['continuousUpdate'], script: { SA4_REGISTRY_INDEX: '[]', REVIEWER_API_TOKEN: 'secret-token-0123456789abcdefghij' }, doc: { PLENARY_DOC: 'https://docs.google.com/document/d/' + FAKE.prod + '/edit', AGENDA: FAKE.agenda } });
check('a trigger, a registry and a production report in the settings: three problems', [bad.problems.length, /trigger/.test(bad.problems[0]), /SA4_REGISTRY_INDEX/.test(bad.problems[1]), /PLENARY_DOC.*6G Media SA4#138/.test(bad.problems[2]), /^NOT ISOLATED: 3/.test(bad.verdict)], [3, true, true, true, true]);
check('an unknown ID: to review, by key; the token value is never scanned or shown', [bad.review, JSON.stringify(bad).includes('secret-token')], [['document property AGENDA: ' + FAKE.agenda + ' (not this copy; check what it is)'], false]);

console.log('it28Snap / it28Compare / it28Snapshots / it28Probe: the Docs REST read, stored and verified snapshots');
{
  const store = {};
  S.PropertiesService.getDocumentProperties = () => ({ getProperties: () => Object.assign({}, store), getProperty: (k) => (k in store ? store[k] : null), setProperty: (k, v) => { store[k] = String(v); }, deleteProperty: (k) => { delete store[k]; } });
  S.DocumentApp.getActiveDocument = () => ({ getId: () => FAKE.own, getName: () => 'DISPOSABLE Issue28 - test' });
  const files = {};
  let fileNo = 0;
  S.it28PutFile_ = (name, text) => { const id = 'file' + (++fileNo); files[id] = { name: name, text: text }; return id; };
  S.it28GetFile_ = (id) => { if (!files[id]) throw new Error('not found'); return files[id].text; };
  let current = doc();
  const requests = [];
  let answer = 200;
  S.ScriptApp.getOAuthToken = () => 'token';
  S.UrlFetchApp = { fetch: (url, o) => { requests.push({ url: url, auth: o.headers.Authorization }); return { getResponseCode: () => answer, getContentText: () => (answer === 200 ? JSON.stringify(reversedKeys(current)) : '{"error":"denied"}') }; } };
  S.SA4_RELEASE_ = { releaseId: 'T-2026.10.33', codeVersion: '2.27.1' };
  const logs = [];
  S.Logger.log = (m) => logs.push(String(m));
  store.IT28_FP_s1_0 = '{}'; store.IT28_FP_s1_N = '1';
  check('the first snap: the baseline s1; the version 1 keys removed; two Docs REST reads (tab structure, then the content of the one tab) with the token', [S.it28Snap(), logs.some((l) => /"s1" is the baseline/.test(l)), Object.keys(store).filter((k) => /^IT28_FP_/.test(k)), requests.length,
    /^https:\/\/docs\.googleapis\.com\/v1\/documents\/1FakeDisposableCopy.*\?includeTabsContent=true&fields=revisionId,tabs\(tabProperties\(tabId,title/.test(requests[0].url), /dropdown/.test(requests[1].url), requests.every((r) => r.url.length < 2000), requests[0].auth], ['s1', true, [], 2, true, true, true, 'Bearer token']);
  logs.length = 0;
  current = splitRuns(doc());
  check('the second (nothing changed; another serialization): IDENTICAL, the same content SHA-256, no compare file', [S.it28Snap(), logs.filter((l) => /\[IT28\] compare "s1" -> "s2" \(\d+ ms\): IDENTICAL/.test(l)).length, JSON.parse(store.IT28_SNAP_s1).contentSha === JSON.parse(store.IT28_SNAP_s2).contentSha, Object.values(files).filter((f) => /compare/.test(f.name)).length], ['s2', 1, true, 0]);
  logs.length = 0;
  current = doc({ t: { 'S4-261001': { minutes: 'Typed minutes, changed.' } } });
  S.SA4_RELEASE_ = { releaseId: 'T-2026.10.33', codeVersion: '2.28.0 (issue #28 phase A candidate, NOT RELEASED)' };
  const s3 = S.it28Snap();
  const cmp = logs.filter((l) => /\[IT28\] compare/.test(l));
  check('the third: compared with s2 and s1, releases named, the change located with old and new, the full list stored', [s3, cmp.length, /compare "s2" -> "s3"/.test(cmp[0]), /compare "s1" -> "s3"/.test(cmp[1]), /2\.27\.1 -> T-2026\.10\.33 \/ 2\.28\.0/.test(cmp[1]),
    /TDoc S4-261001, row "Discussion" -- text: "Typed minutes\." -> "Typed minutes, changed\."/.test(cmp[1]), Object.values(files).filter((f) => /^IT28 compare s1 to s3\.txt$/.test(f.name)).length], ['s3', 2, true, true, true, true, 1]);
  logs.length = 0;
  const rows = S.it28Snapshots();
  check('it28Snapshots: s1..s3 in order, each verified', [rows.length, rows.every((r) => /-- verified$/.test(r)), /^s1: .*2\.27\.1, revision r1, 2 TDoc tables/.test(rows[0])], [3, true, true]);
  check('the numbering is numeric (s10 after s9)', S.it28SnapLabels_({ IT28_SNAP_s9: '{}', IT28_SNAP_s10: '{}', IT28_SNAP_s2: '{}', IT28_SNAP_FOLDER: 'f', IT28_SNAP_x: '{}' }), ['s2', 's9', 's10']);
  const meta2 = JSON.parse(store.IT28_SNAP_s2);
  const keep = files[meta2.fileId].text;
  files[meta2.fileId].text = keep.replace('Typed minutes.', 'Typed minutes!');
  check('a stored snapshot altered: refused (SHA-256)', (() => { try { S.it28LoadSnapshot_('s2'); return 'loaded'; } catch (e) { return /not the one stored/.test(e.message); } })(), true);
  const incomplete = keep.replace('"complete":true', '"complete":false');
  files[meta2.fileId].text = incomplete;
  store.IT28_SNAP_s2 = JSON.stringify(Object.assign({}, meta2, { chars: incomplete.length, sha256: R.idDigest(incomplete) }));
  check('a stored snapshot not complete (even with a matching index): refused', (() => { try { S.it28LoadSnapshot_('s2'); return 'loaded'; } catch (e) { return /incomplete/.test(e.message); } })(), true);
  delete files[meta2.fileId];
  check('a stored snapshot whose file is gone: refused', (() => { try { S.it28LoadSnapshot_('s2'); return 'loaded'; } catch (e) { return /cannot be read/.test(e.message); } })(), true);
  check('... it28Snapshots shows it as not usable', S.it28Snapshots().filter((r) => /NOT USABLE/.test(r)).map((r) => r.slice(0, 3)), ['s2:']);
  answer = 403;
  check('the Docs API refusing: an error, nothing stored', [(() => { try { S.it28Snap(); return 'stored'; } catch (e) { return /answered 403/.test(e.message); } })(), 'IT28_SNAP_s4' in store], [true, false]);
  answer = 200;
  const before = Object.keys(store).length + '/' + Object.keys(files).length;
  logs.length = 0;
  const t = S.it28Probe();
  check('it28Probe: reads and walks, times every request and the walk, lists each tab, stores nothing', [t.reads.length, typeof t.reads[0].fetchMs, typeof t.reads[0].parseMs, typeof t.walkMs, t.responseChars > 0, Object.keys(store).length + '/' + Object.keys(files).length, /probe: 1 tab\(s\), 2 TDoc tables/.test(logs[0]), /tab 1 "" \(t\.0\): 2 TDoc tables, 3 tables, \d+ paragraphs, summary yes; first headings: H1 7 Audio \| H1 Disposition Summary/.test(logs[0])], [2, 'number', 'number', 'number', true, before, true, true]);
}

console.log('the runner: jobs pushed with a stage run by themselves in the copy (one install, then no editor)');
{
  const store = {};
  const sprops = {};
  S.PropertiesService.getDocumentProperties = () => ({ getProperties: () => Object.assign({}, store), getProperty: (k) => (k in store ? store[k] : null), setProperty: (k, v) => { store[k] = String(v); }, deleteProperty: (k) => { delete store[k]; } });
  S.PropertiesService.getScriptProperties = () => ({ getProperties: () => Object.assign({}, sprops), deleteProperty: (k) => { delete sprops[k]; } });
  let docName = 'DISPOSABLE Issue28 - test';
  S.DocumentApp.getActiveDocument = () => ({ getId: () => FAKE.own, getName: () => docName, getBody: () => ({ getTables: () => [] }) });
  const files = {};
  let fileNo = 0;
  S.it28PutFile_ = (name, text) => { const id = 'f' + (++fileNo); files[id] = { name, text }; return id; };
  S.it28GetFile_ = (id) => files[id].text;
  const results = [];
  S.it28WriteResult_ = (id, r) => { if (!r.document) r.document = docName; results.push({ id: id, r: JSON.parse(JSON.stringify(r)) }); };
  const status = [];
  S.it28Status_ = (s) => status.push(s);
  let triggers = [];
  S.ScriptApp.getProjectTriggers = () => triggers.slice();
  S.ScriptApp.deleteTrigger = (t) => { triggers = triggers.filter((x) => x !== t); };
  S.ScriptApp.newTrigger = (h) => ({ timeBased: () => ({ everyMinutes: (m) => ({ create: () => { const t = { getHandlerFunction: () => h, minutes: m }; triggers.push(t); return t; } }) }) });
  let lockFree = true;
  S.LockService = { getScriptLock: () => ({ tryLock: () => lockFree, releaseLock: () => {} }), getDocumentLock: () => docLock };
  let current = doc();
  S.ScriptApp.getOAuthToken = () => 'token';
  S.UrlFetchApp = { fetch: () => ({ getResponseCode: () => 200, getContentText: () => JSON.stringify(current) }) };
  S.SA4_RELEASE_ = { releaseId: 'T-2026.10.33', codeVersion: '2.27.1', gitCommit: 'b7299c12d12b' };
  let updates = 0;
  let refuse = '';
  const docLock = { tryLock: () => true, releaseLock: () => {} };
  S.continuousUpdateCore_ = () => { if (refuse) return { success: false, error: refuse, notUpdated: true }; updates++; if (updates === 1) current = doc({ t: { 'S4-261001': { status: 'noted' } } }); return { success: true, error: null }; };
  S.continuousUpdate = () => { throw new Error('the runner must not call continuousUpdate (it drops the result)'); };
  S.Logger.getLog = () => 'Downloaded 128 TDOCs\n[PERF] TOTAL continuousUpdate: 41234 ms';
  S.Logger.log = () => {};
  const J = require('../tools/integration-test/issue28-jobs.js');
  // A job with updates takes several ticks (an update only first in an execution): tick until its result is written.
  const tickUntilDone = () => { const k = results.length; for (let i = 0; i < 12 && results.length === k; i++) S.it28RunnerTick(); };
  S.it28InstallRunner();
  check('install: one 5-minute trigger for the runner, its state, an "installed" result with the document name', [triggers.length, triggers[0].getHandlerFunction(), triggers[0].minutes, !!store.IT28_RUNNER, results[0].id, results[0].r.event, results[0].r.document], [1, 'it28RunnerTick', 5, true, 'runner', 'installed', 'DISPOSABLE Issue28 - test']);
  S.it28InstallRunner();
  check('install again: still one trigger', triggers.length, 1);
  S.it28RunnerTick();
  check('a tick without a job: nothing happens', results.length, 2);
  const job = J.makeJob('main', 'stage', 't1');
  S.IT28_JOB_ = { id: job.id, steps: job.steps };
  S.IT28_RUNNER_TICK_MS_ = -1;
  S.it28RunnerTick();
  check('a long job goes on in the next tick: one step per tick here, the state kept', [JSON.parse(store.IT28_RUNNER).current.next, results.length], [1, 2]);
  S.IT28_RUNNER_TICK_MS_ = 4 * 60 * 1000;
  S.it28RunnerTick();
  check('an update only as the first step of an execution: the tick ran the first update fresh and stopped before the second', [JSON.parse(store.IT28_RUNNER).current.next, /step 2\/8 \(update\) started/.test(status.join('|'))], [4, true]);
  S.it28RunnerTick();
  const res = results[results.length - 1];
  check('... and finishes: one result for the job, every step ok, the update timed, the compares as the job asks', [res.id, res.r.ok, res.r.steps.map((s) => s.do + (s.ok ? '' : '!')).join(','), typeof res.r.steps[1].updateMs, res.r.steps[3].identical, res.r.steps[6].identical, res.r.release],
    [job.id, true, 'snap,update,snap,compare,update,snap,compare,compare', 'number', false, true, 'T-2026.10.33 / 2.27.1 @b7299c1']);
  check('... the compare counters name the row and field, for the rules; the job checks pass', [res.r.steps[3].counters.rows, J.evaluate(job, res.r)], [{ 'Status / text (with its styles)': 1 }, []]);
  const n = results.length;
  S.it28RunnerTick();
  check('the same job again: not run twice', results.length, n);
  check('the status file says what runs, step by step', [status.some((s) => /step 2\/8 \(update\) started/.test(s)), /finished$/.test(status[status.length - 1])], [true, true]);
  // an execution ended inside a step (the 6-minute limit): recorded, the job stops, the step is not run again
  const killed = { id: 'main-killed-t1c', steps: [{ do: 'update' }, { do: 'snap' }] };
  S.IT28_JOB_ = killed;
  const st0 = JSON.parse(store.IT28_RUNNER); st0.current = { id: killed.id, next: 0, results: [], startedAt: 'x', release: 'y', inProgress: 0, inProgressAt: Date.now() - 360000 }; store.IT28_RUNNER = JSON.stringify(st0);
  const updatesBefore = updates;
  S.it28RunnerTick();
  const kr = results[results.length - 1].r;
  check('an execution ended during an update: a failed step (not run again), the job stops', [kr.job, kr.ok, kr.stoppedAt, /execution ended during this step/.test(kr.steps[0].error), kr.steps.length, updates === updatesBefore], [killed.id, false, 1, true, 1, true]);
  check('an update records the TDocs downloaded and its own total time (from its log)', [res.r.steps[1].tdocsDownloaded, res.r.steps[1].perfTotalMs, res.r.steps[1].result.success], [128, 41234, true]);
  refuse = 'Cannot build report yet. Open Configure Meeting to finish the setup.';
  S.IT28_JOB_ = { id: 'ahg-diagnose-t1b', steps: J.makeJob('ahg', 'diagnose', 'x').steps.concat([{ do: 'snap' }]) };
  tickUntilDone();
  const dg = results[results.length - 1].r;
  check('an update that is refused (success false, not updated -- no exception): a failed step with its error, and the job stops (found 2026-11-25)', [dg.ok, dg.stoppedAt, dg.steps[1].ok, /the update did not run \(nothing changed\): Cannot build report yet/.test(dg.steps[1].error), dg.steps.length], [false, 2, false, true, 2]);
  refuse = '';

  // configureMeeting: Resolve (+ Discover for ad-hoc) from the Portal, then Save -- as the dialog does
  const saved = [];
  S.resolveMeetingForConfigDialog_ = (id) => ({ ok: true, resolved: { id: id }, preview: { meetingId: { value: id }, meetingType: { value: id === '86178' ? 'adhoc' : 'main' }, meetingName: { value: id === '86178' ? '3GPPSA4-e (AH) on FS_6G_MED' : '3GPPSA4#138' }, meetingDate: { value: 'x' }, ftpBase: { value: 'https://ftp/' + id }, agendaTdoc: { value: '' }, revisionsUrl: { value: '' }, tdocListUrl: { value: '' } } });
  S.discoverAgendaForConfigDialog_ = (id, r) => ({ ok: true, preview: Object.assign({}, S.resolveMeetingForConfigDialog_(id).preview, { agendaTdoc: { value: 'S4aP260149' }, tdocListUrl: { value: 'https://portal/list/' + id } }) });
  S.saveConfigurationSettings = (cfg) => { saved.push(cfg); store.MEETING_TYPE = cfg.meetingType; store.MEETING_ID = cfg.meetingId; };
  const sessionsSaved = [];
  S.saveAdhocSessionsConfiguration = (rows) => { sessionsSaved.push(rows); store.ADHOC_SESSIONS = JSON.stringify(rows); return { ok: true }; };
  const cfgJob = J.makeJob('ahg', 'configure-ahg', 't2');
  S.IT28_JOB_ = { id: cfgJob.id, steps: cfgJob.steps };
  S.it28RunnerTick();
  const cfg = results[results.length - 1].r;
  check('configure the AHG copy: meeting 86178 resolved and discovered from the Portal, saved as ad-hoc 6G; its two historical sessions saved', [cfg.ok, saved[0].meetingId, saved[0].meetingType, saved[0].reportType, saved[0].agendaTdoc, saved[0].tdocUrl, cfg.steps[1].sessions, J.evaluate(cfgJob, cfg)],
    [true, '86178', 'adhoc', '6G', 'S4aP260149', 'https://portal/list/86178', ['Call 1 2026-09-22 15:00-18:00', 'Call 2 2026-10-26..2026-10-28 15:00-18:00'], []]);
  const wrong = { id: 'main-wrong-t3', steps: [{ do: 'configureMeeting', meetingId: 86178, expectType: 'main', reportType: '6G' }, { do: 'snap' }] };
  S.IT28_JOB_ = wrong;
  S.it28RunnerTick();
  const w = results[results.length - 1].r;
  check('a meeting of another type than expected: nothing saved, and the job stops there (a step that changes something failed)', [w.ok, w.stoppedAt, w.steps.length, /is "adhoc", not main -- nothing saved/.test(w.steps[0].error), saved.length], [false, 1, 1, true, 1]);
  const bad = { id: 'main-bad-t4', steps: [{ do: 'props' }, { do: 'setProps', props: { IT28_RUNNER: 'x' } }, { do: 'snap' }] };
  S.IT28_JOB_ = bad;
  S.it28RunnerTick();
  const b = results[results.length - 1].r;
  check('setProps may not touch the harness keys: refused, the job stops; an unknown step is refused too (no free calls)', [b.steps.map((s) => s.ok), b.stoppedAt, /may not touch IT28_RUNNER/.test(b.steps[1].error)], [[true, false], 2, true]);
  S.IT28_JOB_ = { id: 'main-call-t5', steps: [{ do: 'call', fn: 'it28Cleanup_' }] };
  S.it28RunnerTick();
  check('... a "call" step does not exist', /unknown step call/.test(results[results.length - 1].r.steps[0].error), true);

  // phase B and C steps
  S.it28Candidates = () => ({ successorUndecided: [], successorDecided: [], namedNotListed: ['S4-261002 -> S4-269999'] });
  const records = {};
  S.saveExpectedRevision = (input) => { records[input.original] = 'pending'; return 'Recorded ' + input.original + ' ' + input.decision; };
  S.changeExpectedRevision = (action, original, details) => { records[original] = action + ':' + details.outcome; return 'Closed ' + original; };
  S.withDispositionLock_ = (name, fn) => fn();
  S.closeReportCore_ = () => (Object.values(records)[0] === 'pending' ? { ready: false, problems: ['S4-261002: revision expected, not registered'], blockingRevisions: [], unresolved: [], informationNotes: [] }
    : { ready: true, problems: [], blockingRevisions: [], unresolved: [], informationNotes: ['S4-261002: revision record closed (not-submitted)'] });
  const pb = J.makeJob('main', 'phase-b', 't6');
  S.IT28_JOB_ = { id: pb.id, steps: pb.steps };
  tickUntilDone();
  const pbr = results[results.length - 1].r;
  check('phase B: a TDoc whose named successor is not listed gets "revision expected"; it blocks Close Report; closed with an outcome and a reason it is information', [pbr.ok, pbr.steps[1].original, pbr.steps[6].tdocBlocking, pbr.steps[8].tdocBlocking, pbr.steps[8].tdocInformation, records['S4-261002'], J.evaluate(pb, pbr)],
    [true, 'S4-261002', true, false, true, 'resolve:not-submitted', []]);
  S.meetingControlData = () => ({ errors: [], rows: 3 });
  S.previewAdhocSessionRule = () => ({ ok: false, error: 'only 3 TDoc(s) of the list have both times' });
  let rule = 'upload';
  S.setAdhocSessionRule = (r) => { rule = r; return 'rule ' + r; };
  S.previewAdhocCarryForward = () => ({ ok: true, changes: [], kept: [] });
  let carry = false;
  S.setAdhocCarryForward = (on) => { carry = on; return 'carry ' + on; };
  const pc = J.makeJob('ahg', 'phase-c', 't7');
  S.IT28_JOB_ = { id: pc.id, steps: pc.steps };
  tickUntilDone();
  const pcr = results[results.length - 1].r;
  check('phase C: the reservation rule is not switched when its preview refuses (said so); carry-forward on after its preview, then both back off', [pcr.ok, pcr.steps[2].previewOk, /preview refused the rule/.test(pcr.steps[3].skipped), pcr.steps[8].message, rule, carry, J.evaluate(pc, pcr)],
    [true, false, true, 'carry true', 'upload', false, []]);

  lockFree = false;
  S.IT28_JOB_ = { id: 'main-x-t8', steps: [{ do: 'props' }] };
  const before = results.length;
  S.it28RunnerTick();
  check('another run holds the lock: the tick does nothing', results.length, before);
  lockFree = true;

  // uninstall: every trigger, every document property but the snapshot index, every script property -- proven
  triggers.push({ getHandlerFunction: () => 'continuousUpdate', minutes: 60 });
  sprops.REVIEWER_API_TOKEN = 'x';
  store.DISPOSITION_CLASS_HEADINGS = 'heading2';
  const snapKeys = Object.keys(store).filter((k) => k.indexOf('IT28_SNAP_') === 0);
  S.IT28_JOB_ = { id: 'main-uninstall-t9', steps: [{ do: 'uninstall' }] };
  S.it28RunnerTick();
  const u = results.slice(-2);
  check('uninstall: all triggers (not only the runner\'s), all test settings and script properties removed; the snapshot index kept; proven in the result and an "uninstalled" record',
    [triggers.length, Object.keys(store).filter((k) => k.indexOf('IT28_SNAP_') !== 0), Object.keys(sprops), Object.keys(store).filter((k) => k.indexOf('IT28_SNAP_') === 0).length === snapKeys.length && snapKeys.length > 0, u[0].r.steps[0].clean, u[1].id, u[1].r.event, J.evaluate(J.makeJob('main', 'uninstall', 'x'), u[0].r)],
    [0, [], [], true, true, 'runner', 'uninstalled', []]);
  S.it28InstallRunner();
  store.MEETING_TYPE = 'main';
  const st = JSON.parse(store.IT28_RUNNER); st.deadline = Date.now() - 1; store.IT28_RUNNER = JSON.stringify(st);
  S.it28RunnerTick();
  const ex = results[results.length - 1].r;
  check('past its 72-hour limit: the runner cleans up by itself (triggers and test settings), proven', [triggers.length, 'IT28_RUNNER' in store, 'MEETING_TYPE' in store, ex.event, ex.cleanup.clean], [0, false, false, 'expired', true]);
  docName = 'Not disposable';
  const before2 = results.length;
  S.it28RunnerTick();
  check('in a document not named DISPOSABLE: refuses (an error result, nothing run)', [results.length, / ERROR$/.test(results[results.length - 1].id)], [before2 + 1, true]);
}

if (failures) { console.log(`\n${failures} check(s) FAILED`); process.exit(1); }
console.log('\nAll checks passed.');
