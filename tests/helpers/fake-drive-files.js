/**
 * A fake of the part of Google Drive the personal settings use: the Advanced
 * Drive service (Drive.Files.list / get / create) and DriveApp.getFileById()
 * (content, setContent, setTrashed). One user's view of Drive: the files
 * the user owns and files other people shared with them.
 *
 * A file: { id, name, mimeType, content, ownedByMe, shared, trashed,
 * properties }. `add()` fills in what is not given as for a private settings
 * file of the user.
 *
 * Drive.Files.list() evaluates the ONE query the production code sends
 * (owner, not trashed, type, marker) and refuses any other, so a changed
 * query fails loudly. With `leaky: true` it returns every file whatever the
 * query says -- for the tests that show the code checks each file itself.
 *
 * Switches for failures: failList, failRead, failWrite, failCreate,
 * failMeta, failTrash (each: true, or a function(id) -> boolean);
 * `dropWrites` (a write or a create reports success but stores nothing);
 * `createShared` (a new file comes out shared).
 * Every call is recorded in `calls`, as [name, id or query].
 */

const EXPECTED_QUERY = "'me' in owners and trashed = false and mimeType = 'application/json' and properties has { key='sa4ReportUserSettings' and value='1' }";
const MARKER = { sa4ReportUserSettings: '1' };
const NAME = 'SA4 Report – private settings (do not share).json';
const byteLength = (text) => Buffer.byteLength(String(text), 'utf8');

function makeFakeDrive(options) {
  const o = options || {};
  const drive = { files: [], calls: [], leaky: !!o.leaky, failList: false, failRead: false, failWrite: false, failCreate: false, failMeta: false, failTrash: false, dropWrites: false, createShared: false, seq: 0 };
  const fails = (flag, id) => (typeof flag === 'function' ? !!flag(id) : !!flag);
  const metaOf = (f) => ({ id: f.id, name: f.name, mimeType: f.mimeType, size: String(byteLength(f.content)), ownedByMe: f.ownedByMe, shared: f.shared, trashed: f.trashed, properties: f.properties ? Object.assign({}, f.properties) : undefined });
  const find = (id) => { const f = drive.files.filter((x) => x.id === id)[0]; if (!f) throw new Error('File not found: ' + id); return f; };

  /** Adds a file; what is not given is that of a private, marked settings file of the user. */
  drive.add = (file) => {
    const f = Object.assign({ id: 'file' + (++drive.seq), name: NAME, mimeType: 'application/json', content: '', ownedByMe: true, shared: false, trashed: false, properties: Object.assign({}, MARKER) }, file || {});
    drive.files.push(f);
    return f;
  };
  /** A settings file holding `token`. */
  drive.addSettings = (token, file) => drive.add(Object.assign({ content: JSON.stringify({ schema: 'sa4-report-user-settings/1', reviewerApiToken: token }) }, file || {}));
  drive.count = (name) => drive.calls.filter((c) => c[0] === name).length;
  drive.live = () => drive.files.filter((f) => !f.trashed);

  drive.Drive = { Files: {
    list: (params) => {
      const q = params && params.q;
      drive.calls.push(['list', q]);
      if (fails(drive.failList)) throw new Error('Drive is not available (synthetic)');
      if (q !== EXPECTED_QUERY) throw new Error('Unexpected Drive query: ' + q);
      const hits = drive.leaky ? drive.files.slice()
        : drive.files.filter((f) => f.ownedByMe === true && f.trashed === false && f.mimeType === 'application/json' && f.properties && f.properties.sa4ReportUserSettings === '1');
      return { files: hits.map(metaOf) };
    },
    get: (id) => {
      drive.calls.push(['meta', id]);
      if (fails(drive.failMeta, id)) throw new Error('Drive is not available (synthetic)');
      return metaOf(find(id));
    },
    create: (resource, blob) => {
      drive.calls.push(['create', resource && resource.name]);
      if (fails(drive.failCreate)) throw new Error('Drive is not available (synthetic)');
      const f = { id: 'created' + (++drive.seq), name: resource.name, mimeType: resource.mimeType, content: drive.dropWrites ? '' : blob._text, ownedByMe: true, shared: !!drive.createShared, trashed: false,
        properties: resource.properties ? Object.assign({}, resource.properties) : undefined };
      drive.files.push(f);
      return { id: f.id };
    }
  } };
  drive.DriveApp = {
    getFileById: (id) => ({
      getBlob: () => ({ getDataAsString: () => {
        drive.calls.push(['read', id]);
        if (fails(drive.failRead, id)) throw new Error('Drive is not available (synthetic)');
        return find(id).content;
      } }),
      setContent: (text) => {
        drive.calls.push(['write', id]);
        if (fails(drive.failWrite, id)) throw new Error('Drive is not available (synthetic)');
        if (!drive.dropWrites) find(id).content = String(text);
      },
      setTrashed: (value) => {
        drive.calls.push(['trash', id]);
        if (fails(drive.failTrash, id)) throw new Error('Drive is not available (synthetic)');
        find(id).trashed = !!value;
      }
    })
  };
  drive.newBlob = (text, mimeType, name) => ({ _text: String(text), _mimeType: mimeType, _name: name });

  /** Puts the fake into a sandbox, beside whatever DriveApp and Utilities it has. */
  drive.install = (sandbox) => {
    sandbox.Drive = drive.Drive;
    sandbox.DriveApp = Object.assign({}, sandbox.DriveApp || {}, drive.DriveApp);
    sandbox.Utilities = Object.assign({}, sandbox.Utilities || {}, { newBlob: drive.newBlob });
    return drive;
  };
  return drive;
}

module.exports = { makeFakeDrive, EXPECTED_QUERY, MARKER, NAME };
