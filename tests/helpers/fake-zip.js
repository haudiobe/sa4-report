/**
 * Utilities.newBlob() / Utilities.unzip() for the tests, on real ZIP bytes
 * (a .docx is a ZIP). Reads the central directory; stored and deflated
 * entries only (what Word and 3GPP ZIPs use). No dependency outside Node.
 *
 * A blob here: getName(), setName(), getBytes() (an array of numbers, as in
 * Apps Script), getContentType(), setContentType(), getDataAsString(charset).
 */

const zlib = require('zlib');

function makeBlob(bytes, contentType, name) {
  const buf = Buffer.from(bytes);
  const blob = {
    _name: name || '', _type: contentType === undefined ? null : contentType,
    getName: () => blob._name,
    setName: (n) => { blob._name = n; return blob; },
    getContentType: () => blob._type,
    setContentType: (t) => { blob._type = t; return blob; },
    getBytes: () => Array.from(buf),
    getDataAsString: (charset) => buf.toString(/^utf-?8$/i.test(charset || 'UTF-8') ? 'utf8' : 'latin1'),
    _buffer: buf
  };
  return blob;
}

/** The entries of a ZIP as blobs, in central-directory order. Throws on anything that is not a ZIP. */
function unzip(blob) {
  const buf = blob._buffer || Buffer.from(blob.getBytes());
  let eocd = -1;
  for (let i = buf.length - 22; i >= Math.max(0, buf.length - 65557); i--) {
    if (buf.readUInt32LE(i) === 0x06054b50) { eocd = i; break; }
  }
  if (eocd < 0) throw new Error('Could not unzip: not a ZIP file');
  const count = buf.readUInt16LE(eocd + 10);
  let p = buf.readUInt32LE(eocd + 16);
  const out = [];
  for (let n = 0; n < count; n++) {
    if (buf.readUInt32LE(p) !== 0x02014b50) throw new Error('Could not unzip: bad central directory');
    const method = buf.readUInt16LE(p + 10);
    const compressed = buf.readUInt32LE(p + 20);
    const nameLen = buf.readUInt16LE(p + 28), extraLen = buf.readUInt16LE(p + 30), commentLen = buf.readUInt16LE(p + 32);
    const local = buf.readUInt32LE(p + 42);
    const name = buf.slice(p + 46, p + 46 + nameLen).toString('utf8');
    const dataStart = local + 30 + buf.readUInt16LE(local + 26) + buf.readUInt16LE(local + 28);
    const raw = buf.slice(dataStart, dataStart + compressed);
    const data = method === 0 ? raw : method === 8 ? zlib.inflateRawSync(raw) : null;
    if (data === null) throw new Error('Could not unzip: method ' + method);
    if (!name.endsWith('/')) out.push(makeBlob(data, null, name));
    p += 46 + nameLen + extraLen + commentLen;
  }
  return out;
}

module.exports = { makeBlob, unzip };
