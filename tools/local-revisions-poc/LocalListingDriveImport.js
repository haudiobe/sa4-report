/**
 * DISPOSABLE PROOF OF CONCEPT -- not part of any release bundle.
 *
 * The report side of the local listing agent (sync-local-listings.ps1): reads
 * the listing the agent wrote for this report's working group into a Google
 * Drive folder, and imports it with the existing Import Local Inbox Listing
 * code (saveLocalRevisionsListing(): same parser, same per-folder refresh,
 * same merge with the 3GPP revisions, same rendering). Nothing new is parsed
 * or stored here.
 *
 * PoC configuration (Document Property of the disposable copy):
 *   LOCAL_LISTING_DRIVE_FOLDER_ID   the Drive folder the agent writes to
 * Run pocImportLocalListingFromDrive() from the editor of a DISPOSABLE copy.
 */

var POC_LOCAL_LISTING_MAX_AGE_MINUTES_ = 30;

/** Pure: the agent's file -- { folder, fetchedAt, html } from its first line, or null when it is not one. */
function pocParseAgentListing_(text) {
  const m = /^<!-- sa4-local-listing folder="([^"]+)" fetched="([^"]+)" -->\r?\n/.exec(String(text || ''));
  if (!m) return null;
  return { folder: m[1], fetchedAt: m[2], html: String(text).slice(m[0].length) };
}

/** Pure: whether a listing fetched at `fetchedAt` is recent enough at `now` (Date). */
function pocAgentListingFresh_(fetchedAt, now, maxAgeMinutes) {
  const t = Date.parse(fetchedAt);
  return !isNaN(t) && now.getTime() - t <= maxAgeMinutes * 60000 && t <= now.getTime() + 60000;
}

/** The working-group file name of this report: the drafts folder of its family (DRAFTS_FOLDERS). */
function pocAgentGroupOfReport_() {
  return getReportConfig_().DRAFTS_FOLDER;
}

/**
 * Editor (disposable copy only): imports the agent's listing for this report.
 * Returns { ok, message } as saveLocalRevisionsListing() does, or why nothing was imported.
 */
function pocImportLocalListingFromDrive() {
  const name = currentReportDocument_().getName();
  if (name.indexOf('DISPOSABLE TEST COPY') !== 0) throw new Error('PoC: disposable copies only.');
  const folderId = String(getReportStateStore_().getProperty('LOCAL_LISTING_DRIVE_FOLDER_ID') || '').trim();
  if (!folderId) return { ok: false, message: 'No LOCAL_LISTING_DRIVE_FOLDER_ID set.' };
  const group = pocAgentGroupOfReport_();
  const files = DriveApp.getFolderById(folderId).getFilesByName(group + '.html');
  if (!files.hasNext()) return { ok: false, message: 'Not checked: the agent has written no listing for ' + group + ' yet.' };
  const listing = pocParseAgentListing_(files.next().getBlob().getDataAsString('UTF-8'));
  if (!listing) return { ok: false, message: 'Not imported: ' + group + '.html is not a listing written by the agent.' };
  if (!pocAgentListingFresh_(listing.fetchedAt, new Date(), POC_LOCAL_LISTING_MAX_AGE_MINUTES_)) {
    return { ok: false, message: 'Not imported: the newest listing of ' + group + ' is from ' + listing.fetchedAt + ' (meeting-room server or agent unavailable).' };
  }
  const result = saveLocalRevisionsListing({ folder: listing.folder, html: listing.html, text: '' });
  Logger.log('PoC local listing import (' + group + ', fetched ' + listing.fetchedAt + '): ' + result.message);
  return result;
}
