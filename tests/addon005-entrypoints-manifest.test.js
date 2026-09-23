/**
 * ADDON-005 — add-on entry points (onInstall/onOpen), legacy menu
 * preservation, and locally-testable manifest expectations.
 *
 * Run: node tests/addon005-entrypoints-manifest.test.js
 */

const fs = require('fs');
const path = require('path');
const { loadCode, CODE_JS_PATH } = require('./helpers/load-code.js');

let failures = 0;
function check(name, actual, expected) {
  const a = JSON.stringify(actual);
  const e = JSON.stringify(expected);
  if (a === e) {
    console.log(`  ok   ${name}`);
  } else {
    failures++;
    console.log(`  FAIL ${name}\n         expected ${e}\n         actual   ${a}`);
  }
}

// -------------------------------------------------------------- onInstall

console.log('onInstall(e) -- standard delegation to onOpen(e), does nothing else');

{
  const { sandbox } = loadCode();
  let onOpenCalledWith = 'not-called';
  sandbox.onOpen = (e) => { onOpenCalledWith = e; };

  sandbox.onInstall({ marker: 'install-event' });

  check('onInstall(e) delegates to onOpen(e) with the SAME event object',
    onOpenCalledWith, { marker: 'install-event' });
}

{
  const { sandbox } = loadCode();
  const fakeMenu = () => ({ addItem: () => fakeMenu(), addSeparator: () => fakeMenu(), addSubMenu: () => fakeMenu(), addToUi: () => {} });
  sandbox.DocumentApp.getUi = () => ({ createMenu: () => fakeMenu() });
  // onInstall must not adopt, register, or create a trigger as a side effect.
  let adoptCalls = 0, registerCalls = 0;
  sandbox.adoptReportDocumentForAddon_ = () => { adoptCalls++; };
  sandbox.registerReportDocument_ = () => { registerCalls++; };
  sandbox.ScriptApp = { getProjectTriggers: () => [] };

  sandbox.onInstall();

  check('onInstall() has no adoption/registration side effects', [adoptCalls, registerCalls], [0, 0]);
}

// ------------------------------------------------------------------ onOpen

console.log('onOpen(e) -- accepts (and ignores) the add-on event object; builds the SAME menu either way');

{
  function buildMenuCallLog(sandbox) {
    const calls = [];
    const fakeMenu = () => ({
      addItem: (...a) => { calls.push(['addItem', ...a]); return fakeMenu(); },
      addSeparator: () => { calls.push(['addSeparator']); return fakeMenu(); },
      addSubMenu: (...a) => { calls.push(['addSubMenu']); return fakeMenu(); },
      addToUi: () => { calls.push(['addToUi']); }
    });
    sandbox.DocumentApp.getUi = () => ({ createMenu: (name) => { calls.push(['createMenu', name]); return fakeMenu(); } });
    return calls;
  }

  const { sandbox: sandboxBound } = loadCode();
  const callsBound = buildMenuCallLog(sandboxBound);
  sandboxBound.onOpen(); // legacy bound-script simple trigger: no event object at all

  const { sandbox: sandboxAddon } = loadCode();
  const callsAddon = buildMenuCallLog(sandboxAddon);
  sandboxAddon.onOpen({ authMode: 'FULL' }); // add-on: real event object present

  check('menu-building call sequence is IDENTICAL whether or not an event object is passed',
    JSON.stringify(callsBound) === JSON.stringify(callsAddon), true);
}

console.log('onOpen() -- the new add-on submenu is ADDITIVE: every pre-existing menu item is still present');

{
  const { sandbox } = loadCode();
  const items = [];
  const fakeMenu = () => ({
    addItem: (label, fn) => { items.push([label, fn]); return fakeMenu(); },
    addSeparator: () => fakeMenu(),
    addSubMenu: () => fakeMenu(),
    addToUi: () => {}
  });
  sandbox.DocumentApp.getUi = () => ({ createMenu: () => fakeMenu() });

  sandbox.onOpen();

  const targets = items.map((i) => i[1]);
  // A representative sample of pre-existing, unrelated-to-ADDON-005 menu targets.
  ['configureMeetingSettings', 'runFullReportBuild', 'continuousUpdate', 'manageTriggers',
   'buildSkeletonWithTdocTables', 'updateReportIncremental', 'clearAllCaches'].forEach((fn) => {
    check(`pre-existing menu item "${fn}" is still present`, targets.indexOf(fn) !== -1, true);
  });

  // The new ADDON-005 items.
  ['enableAutomaticUpdatesForAddon', 'disableAutomaticUpdatesForAddon',
   'setAutomaticUpdateIntervalForAddon', 'showAddonSchedulerStatusForAddon'].forEach((fn) => {
    check(`new add-on menu item "${fn}" is present`, targets.indexOf(fn) !== -1, true);
  });
}

console.log('onOpen() -- every menu target referenced actually exists as a callable function');

{
  const { sandbox } = loadCode();
  const items = [];
  const fakeMenu = () => ({
    addItem: (label, fn) => { items.push(fn); return fakeMenu(); },
    addSeparator: () => fakeMenu(),
    addSubMenu: () => fakeMenu(),
    addToUi: () => {}
  });
  sandbox.DocumentApp.getUi = () => ({ createMenu: () => fakeMenu() });
  sandbox.onOpen();

  const missing = items.filter((name) => typeof sandbox[name] !== 'function');
  check('no menu item references a missing function', missing, []);
}

// ---------------------------------------------------- manifest expectations

console.log('appsscript.json -- locally-testable manifest expectations');

{
  const manifestPath = path.join(path.dirname(CODE_JS_PATH), 'appsscript.json');
  const manifest = JSON.parse(fs.readFileSync(manifestPath, 'utf8'));

  check('is valid JSON with the expected top-level shape', typeof manifest, 'object');
  check('existing Advanced Drive service is preserved',
    manifest.dependencies && manifest.dependencies.enabledAdvancedServices &&
    manifest.dependencies.enabledAdvancedServices.some(s => s.userSymbol === 'Drive' && s.serviceId === 'drive'),
    true);
  check('runtimeVersion is still V8 (unchanged)', manifest.runtimeVersion, 'V8');
  check('timeZone is unchanged', manifest.timeZone, 'Europe/Berlin');

  const scopes = manifest.oauthScopes || [];
  const requiredScopes = [
    'https://www.googleapis.com/auth/documents',          // DocumentApp incl. openById on other documents
    'https://www.googleapis.com/auth/drive',               // DriveApp temp-file create/trash + Advanced Drive service
    'https://www.googleapis.com/auth/spreadsheets',        // SpreadsheetApp.open() for TDoc-list xlsx conversion
    'https://www.googleapis.com/auth/script.external_request', // UrlFetchApp
    'https://www.googleapis.com/auth/script.scriptapp',    // ScriptApp trigger lifecycle
    'https://www.googleapis.com/auth/script.container.ui'  // DocumentApp.getUi()/HtmlService dialogs
  ];
  requiredScopes.forEach((scope) => {
    check(`oauthScopes includes ${scope}`, scopes.indexOf(scope) !== -1, true);
  });
  check('no unexpected extra scopes were added', scopes.length, requiredScopes.length);

  // ADDON-005 deliberately does NOT add an `addOns` manifest block yet --
  // see the ADDON-005 report: primary Google documentation states the
  // (classic) Editor add-on manifest has no add-on-specific required
  // properties, and the exact required shape of a Google Workspace
  // add-on's `addOns.common` block (e.g. whether logoUrl is mandatory)
  // could not be confirmed from documentation alone. Adding a guessed
  // block risks a deploy-time validation failure for no local-testing
  // benefit; this is finalized at the moment of the actual first
  // deployment attempt instead.
  check('no addOns block was speculatively added', manifest.addOns, undefined);
}

// ========================================================== summary =======

if (failures > 0) {
  console.log(`\n${failures} check(s) FAILED`);
  process.exit(1);
} else {
  console.log('\nAll checks passed.');
}
