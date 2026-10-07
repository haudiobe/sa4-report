/**
 * A fake of the part of Google Drive the Shared Minutes use: the Advanced
 * Drive service -- Drive.Files.list / get / create and Drive.Permissions.list
 * / create. One user's Drive, as one script project sees it.
 *
 * A file: { id, name, mimeType, trashed, appProperties, parents, createdTime,
 * html, permissions: [{ id, type, role }], visible }. `visible: false` is a
 * file Drive has but does not find yet (its search lags behind).
 *
 * appProperties are private to the script project that wrote them: the fake
 * is ONE project (`project`), a file remembers the project that marked it,
 * and another project (drive.asProject()) sees the same files without their
 * markers -- as a copy of a report does.
 *
 * The report itself is a file too (`reportId`), in `folder` ({ id, name,
 * writable }, or null for a report without a folder the user can see).
 *
 * Drive.Files.list() evaluates the ONE query the production code sends for
 * the report it is asked about, answers the query of the personal settings
 * with no file, and refuses any other, so a changed query fails loudly.
 *
 * Switches (each: true, or a function(id) -> boolean, unless said otherwise):
 *   failSearch, failMeta, failFolder, failCreate, failPermissions, failShare
 *   createThenThrow     the document is created, then the request fails (a timeout)
 *   shareThenThrow      the permission is set, then the request fails (a timeout)
 *   shareIgnored        sharing reports success and sets nothing (a policy that drops it)
 *   shareRole           the role sharing really sets ('writer'; 'reader' for a policy that lowers it)
 *   searchLag           a created document is not found by a search until drive.settle()
 * Every call is recorded in `calls`, as [name, id or query, body].
 */

const DOC_MIME = 'application/vnd.google-apps.document';
const SETTINGS_QUERY = "'me' in owners and trashed = false and mimeType = 'application/json' and properties has { key='sa4ReportUserSettings' and value='1' }";
const queryFor = (reportId) => "appProperties has { key='sa4ReportId' and value='" + reportId + "' } and appProperties has { key='sa4Purpose' and value='shared-minutes' }" +
  " and mimeType = '" + DOC_MIME + "' and trashed = false";

function makeSharedMinutesDrive(options) {
  const o = options || {};
  const state = o.state || { files: [], seq: 0, permissionSeq: 0 };
  const drive = { files: state.files, calls: [], project: o.project || 'report-project', reportId: o.reportId, folder: o.folder === undefined ? { id: 'FOLDERid00000000000000000000000000', name: 'Meeting folder', writable: true } : o.folder,
    failSearch: false, failMeta: false, failFolder: false, failCreate: false, failPermissions: false, failShare: false,
    createThenThrow: false, shareThenThrow: false, shareIgnored: false, shareRole: 'writer', searchLag: false };
  const fails = (flag, id) => (typeof flag === 'function' ? !!flag(id) : !!flag);
  const find = (id) => { const f = drive.files.filter((x) => x.id === id)[0]; if (!f) throw new Error('File not found: ' + id); return f; };
  // The markers of a file, as this project sees them.
  const markerOf = (f) => (f.appProperties && f.markedBy === drive.project ? Object.assign({}, f.appProperties) : undefined);
  // The text and the type of a blob: of the blob this fake makes, and of the one the e-mail export stubs make.
  const textOf = (blob) => (!blob ? '' : (blob._text !== undefined ? blob._text : Buffer.from(blob.getBytes().map((b) => (b < 0 ? b + 256 : b))).toString('utf8')));
  const typeOf = (blob) => (!blob ? '' : (blob._mimeType !== undefined ? blob._mimeType : blob.getContentType()));
  const metaOf = (f) => ({ id: f.id, name: f.name, mimeType: f.mimeType, trashed: f.trashed, appProperties: markerOf(f), createdTime: f.createdTime });

  /** Adds a file; what is not given is that of a Shared Minutes document marked by this project. */
  drive.add = (file) => {
    const n = ++state.seq;
    const f = Object.assign({ id: 'SHAREDMINUTESdoc' + String(n).padStart(28, '0'), name: 'Shared Minutes ' + n, mimeType: DOC_MIME, trashed: false, appProperties: null, parents: [], createdTime: '2026-10-01T08:00:00.000Z',
      html: '', permissions: [{ id: 'owner', type: 'user', role: 'owner' }], visible: true, markedBy: drive.project }, file || {});
    drive.files.push(f);
    return f;
  };
  drive.count = (name) => drive.calls.filter((c) => c[0] === name).length;
  drive.names = () => drive.calls.map((c) => c[0]);
  drive.live = () => drive.files.filter((f) => !f.trashed);
  /** Drive finds everything it has. */
  drive.settle = () => { drive.files.forEach((f) => { f.visible = true; }); };
  /** The same Drive, seen by another script project (a copy of the report): the same files, none of this project's markers. */
  drive.asProject = (project, reportId) => makeSharedMinutesDrive({ state: state, project: project, reportId: reportId, folder: drive.folder });
  /** Whether anyone with the link can edit a file. */
  drive.anyoneCanEdit = (id) => find(id).permissions.some((p) => p.type === 'anyone' && p.role === 'writer');

  drive.Drive = {
    Files: {
      list: (params) => {
        const q = params && params.q;
        if (q === SETTINGS_QUERY) { drive.calls.push(['settings-list', q]); return { files: [] }; }
        drive.calls.push(['search', q, params]);
        if (fails(drive.failSearch)) throw new Error('Drive is not available (synthetic)');
        const m = String(q).match(/^appProperties has \{ key='sa4ReportId' and value='([A-Za-z0-9_-]+)' \}/);
        if (!m || q !== queryFor(m[1])) throw new Error('Unexpected Drive query: ' + q);
        const hits = drive.files.filter((f) => { const marker = markerOf(f); return f.visible && !f.trashed && f.mimeType === DOC_MIME && marker && marker.sa4ReportId === m[1] && marker.sa4Purpose === 'shared-minutes'; });
        return { files: hits.map(metaOf) };
      },
      get: (id, args) => {
        const fields = String((args && args.fields) || '');
        if (fields === 'parents') {
          drive.calls.push(['report-parents', id]);
          if (fails(drive.failFolder, id)) throw new Error('Drive is not available (synthetic)');
          if (id !== drive.reportId) throw new Error('File not found: ' + id);
          return { parents: drive.folder ? [drive.folder.id] : undefined };
        }
        if (/capabilities/.test(fields)) {
          drive.calls.push(['folder', id]);
          if (fails(drive.failFolder, id)) throw new Error('Drive is not available (synthetic)');
          if (!drive.folder || id !== drive.folder.id) throw new Error('File not found: ' + id);
          return { id: drive.folder.id, name: drive.folder.name, capabilities: { canAddChildren: !!drive.folder.writable } };
        }
        drive.calls.push(['meta', id]);
        if (fails(drive.failMeta, id)) throw new Error('Drive is not available (synthetic)');
        return metaOf(find(id));
      },
      create: (resource, blob, args) => {
        drive.calls.push(['create', resource && resource.name, { resource: JSON.parse(JSON.stringify(resource)), html: textOf(blob), mimeType: typeOf(blob), args: args }]);
        if (fails(drive.failCreate)) throw new Error('The user has exceeded their Drive storage quota (synthetic)');
        const f = drive.add({ name: resource.name, mimeType: resource.mimeType, appProperties: resource.appProperties ? Object.assign({}, resource.appProperties) : null, parents: (resource.parents || []).slice(),
          createdTime: '2026-10-07T09:00:00.000Z', html: textOf(blob), visible: !drive.searchLag });
        if (fails(drive.createThenThrow)) throw new Error('Exceeded maximum execution time (synthetic)');
        return metaOf(f);
      }
    },
    Permissions: {
      list: (id, args) => {
        drive.calls.push(['permissions', id, args]);
        if (fails(drive.failPermissions, id)) throw new Error('Drive is not available (synthetic)');
        return { permissions: find(id).permissions.map((p) => Object.assign({}, p)) };
      },
      create: (resource, id, args) => {
        drive.calls.push(['share', id, { resource: JSON.parse(JSON.stringify(resource)), args: args }]);
        if (fails(drive.failShare, id)) throw new Error('Sharing outside the organization is not allowed (synthetic)');
        if (!drive.shareIgnored) find(id).permissions.push({ id: 'anyoneWithLink', type: resource.type, role: drive.shareRole });
        if (fails(drive.shareThenThrow, id)) throw new Error('Exceeded maximum execution time (synthetic)');
        return { id: 'anyoneWithLink' };
      }
    }
  };
  drive.newBlob = (text, mimeType, name) => ({ _text: String(text), _mimeType: mimeType, _name: name });

  /** Puts the fake into a sandbox, beside whatever Utilities it has. */
  drive.install = (sandbox) => {
    sandbox.Drive = drive.Drive;
    sandbox.Utilities = Object.assign({}, sandbox.Utilities || {}, { newBlob: drive.newBlob });
    return drive;
  };
  return drive;
}

module.exports = { makeSharedMinutesDrive, queryFor, DOC_MIME, SETTINGS_QUERY };
