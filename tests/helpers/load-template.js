/**
 * Loads Code.js and then template/ReportCreator.js into ONE sandbox -- the
 * same shared global scope Apps Script gives the files of one project.
 * Optionally defines SA4_RELEASE_ first, as the generated Release.js would.
 */

const fs = require('fs');
const path = require('path');
const vm = require('vm');
const { loadCode } = require('./load-code.js');

const REPORT_CREATOR_PATH = path.join(__dirname, '..', '..', 'template', 'ReportCreator.js');

function loadTemplateRuntime(options) {
  const opts = options || {};
  const loaded = loadCode(opts);
  if (opts.release) {
    vm.runInContext('var SA4_RELEASE_ = ' + JSON.stringify(opts.release) + ';', loaded.sandbox, { filename: 'Release.js' });
  }
  vm.runInContext(fs.readFileSync(REPORT_CREATOR_PATH, 'utf8'), loaded.sandbox, { filename: 'ReportCreator.js' });
  return loaded;
}

module.exports = { loadTemplateRuntime, REPORT_CREATOR_PATH };
