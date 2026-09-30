/**
 * ADDON-009: the Apps Script stubs the discussion e-mail exporter needs,
 * ported from the Legacy loader (sa4-report-legacy tests/helpers/
 * load-code.js, LEGACY-UPGRADE-006/006D). Installed onto a sandbox returned
 * by tests/helpers/load-code.js, so the shared loader -- and every other
 * suite -- is unchanged.
 *
 * newBlob()/base64Encode() are real (Node Buffer, UTF-8, SIGNED bytes like
 * Apps Script); zip() is a recording fake (no ZIP bytes -- it records which
 * blobs were archived in `_zipEntryNames`); formatDate() is a deterministic
 * token formatter over the Date's UTC fields.
 */

function installExportStubs(sandbox) {
  sandbox.Session.getActiveUser = () => ({ getEmail: () => '' });

  const u = sandbox.Utilities;
  u.formatDate = (date, tz, fmt) => {
    const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
    const pad = (n) => String(n).padStart(2, '0');
    const y = date.getUTCFullYear(), mo = date.getUTCMonth(), d = date.getUTCDate();
    const h = date.getUTCHours(), mi = date.getUTCMinutes(), s = date.getUTCSeconds();
    return String(fmt)
      .replace(/yyyy/g, String(y))
      .replace(/MMM/g, MONTHS[mo])
      .replace(/MM/g, pad(mo + 1))
      .replace(/dd/g, pad(d))
      .replace(/HH/g, pad(h))
      .replace(/mm/g, pad(mi))
      .replace(/ss/g, pad(s));
  };
  u.zip = (blobs, name) => {
    const entryNames = (blobs || []).map((b) => (typeof b.getName === 'function' ? b.getName() : ''));
    let blobName = name || 'archive.zip';
    return {
      getBytes: () => [],
      getName: () => blobName,
      setName: (n) => { blobName = n; },
      getContentType: () => 'application/zip',
      _zipEntryNames: entryNames
    };
  };
  u.newBlob = (data, contentType, name) => {
    const bytes = Array.from(Buffer.from(String(data === undefined ? '' : data), 'utf8')).map((b) => (b > 127 ? b - 256 : b));
    let blobName = name || '';
    return {
      getBytes: () => bytes.slice(),
      getName: () => blobName,
      setName: (n) => { blobName = n; },
      getContentType: () => contentType || 'application/octet-stream'
    };
  };
  u.base64Encode = (input) => {
    const buf = Array.isArray(input) ? Buffer.from(input.map((b) => (b < 0 ? b + 256 : b))) : Buffer.from(String(input), 'utf8');
    return buf.toString('base64');
  };
  return sandbox;
}

/** The UTF-8 text of a blob made by the newBlob() stub above. */
function blobText(blob) {
  return Buffer.from(blob.getBytes().map((b) => (b < 0 ? b + 256 : b))).toString('utf8');
}

module.exports = { installExportStubs, blobText };
