const ARCHIVE_HEADERS = [
  'CD#', 'Case#', 'Image #', 'Calibration', 'On Server?', 'Study',
  'Comments', 'Segmentation Status',
];
const ARCHIVE_TAPS = ['0.36', '0.58', '1.1', '2.9'];

function onOpen() {
  SpreadsheetApp.getUi().createMenu('RASS archive')
    .addItem('Log a Drive folder', 'openArchiveIntake').addToUi();
}

function openArchiveIntake() {
  SpreadsheetApp.getUi().showSidebar(
    HtmlService.createHtmlOutputFromFile('Intake').setTitle('Log a Drive folder'),
  );
}

function archiveFolderId_(input) {
  const match = String(input).trim().match(
    /^(?:https:\/\/drive\.google\.com\/drive\/(?:u\/\d+\/)?folders\/)?([\w-]{15,})(?:[?#].*)?$/,
  );
  if (!match) throw new Error('Paste a Google Drive folder link or its folder ID.');
  return match[1];
}

function inspectArchiveFolder(input) {
  const id = archiveFolderId_(input);
  const fields = 'name,mimeType,parents';
  const folder = Drive.Files.get(id, {fields, supportsAllDrives: true});
  if (folder.mimeType !== 'application/vnd.google-apps.folder') {
    throw new Error('Choose the folder for one case/block.');
  }
  const ancestors = [folder.name];
  let parent = folder.parents && folder.parents[0];
  for (let depth = 0; parent && depth < 8; depth++) {
    // Parent visibility can be more limited than the selected folder.
    try {
      const item = Drive.Files.get(parent, {fields: 'name,parents', supportsAllDrives: true});
      ancestors.unshift(item.name);
      parent = item.parents && item.parents[0];
    } catch (error) {
      break;
    }
  }
  const cd = ancestors.map(name => name.match(/^CD\s*#?\s*(\d+)\b/i)).find(Boolean);
  const samples = [];
  let imageCount = 0;
  let folderCount = 0;
  let pageToken;
  do {
    const page = Drive.Files.list({
      q: `'${id}' in parents and trashed = false`,
      fields: 'nextPageToken,files(name,mimeType)',
      supportsAllDrives: true,
      includeItemsFromAllDrives: true,
      pageSize: 1000,
      pageToken,
    });
    for (const item of page.files || []) {
      if (item.mimeType === 'application/vnd.google-apps.folder') folderCount++;
      if (!item.name.startsWith('._') && /\.(tiff?|png|jpe?g)$/i.test(item.name)) {
        imageCount++;
        if (samples.length < 8) samples.push(item.name);
      }
    }
    pageToken = page.nextPageToken;
  } while (pageToken);
  return {
    id, path: ancestors.join(' / '), cd: cd ? Number(cd[1]) : '',
    block: folder.name, imageCount, folderCount, samples,
  };
}

function archiveKey_(value) {
  return String(value).toLowerCase().trim()
    .replace(/\([^)]*\)/g, '')
    .replace(/\bmch00-0(\d)\b/g, 'mch00-$1')
    .replace(/\bblk\s*(?=\d)/g, 'b')
    .replace(/\bb\s+(?=\d)/g, 'b')
    .replace(/\s+(?:measured|borderline|rejected)$/g, '')
    .replace(/\s+/g, ' ').trim();
}

function archiveText_(value) {
  const text = String(value || '').trim();
  return text.startsWith('=') ? "'" + text : text;
}

function saveArchiveFolder(input) {
  const id = archiveFolderId_(input.folderId);
  const folder = Drive.Files.get(id, {fields: 'mimeType', supportsAllDrives: true});
  if (folder.mimeType !== 'application/vnd.google-apps.folder') {
    throw new Error('Choose a Drive folder.');
  }
  const cd = Number(input.cd);
  const block = String(input.block || '').trim();
  if (!Number.isInteger(cd) || cd <= 0 || !archiveKey_(block)) {
    throw new Error('Enter the CD number and case/block label.');
  }
  const requested = Array.isArray(input.taps) ? input.taps : [];
  if (requested.some(tap => !ARCHIVE_TAPS.includes(tap))) throw new Error('Invalid tap.');
  const taps = ARCHIVE_TAPS.filter(tap => requested.includes(tap)).join(', ');
  const reference = String(input.reference || '').trim();
  const link = `https://drive.google.com/drive/folders/${id}`;
  const lock = LockService.getDocumentLock();
  lock.waitLock(10000);
  try {
    const sheet = SpreadsheetApp.getActive().getSheetByName('Sheet1');
    if (!sheet || sheet.getRange(1, 1, 1, 8).getDisplayValues()[0].some(
      (value, i) => value.trim() !== ARCHIVE_HEADERS[i],
    )) throw new Error('The archive log headers have changed.');

    const count = sheet.getLastRow() - 1;
    const existing = count > 0 ? sheet.getRange(2, 1, count, 8).getDisplayValues() : [];
    const matches = [];
    existing.forEach((row, i) => {
      if (Number(row[0]) === cd && archiveKey_(row[1]) === archiveKey_(block) && !row[2].trim()) {
        matches.push({row, number: i + 2});
      }
    });
    if (matches.length > 1) throw new Error('Multiple block rows match. Check the log manually.');
    if (existing.some((row, i) => row[6].includes(link) && i + 2 !== matches[0]?.number)) {
      throw new Error('This folder is already linked under another case/block label.');
    }
    if (matches.length === 1) {
      const {row, number} = matches[0];
      const current = row[3].split(',').map(value => value.trim()).filter(Boolean).sort();
      const proposed = taps.split(', ').filter(Boolean).sort();
      if (current.length && proposed.length && current.join(',') !== proposed.join(',')) {
        throw new Error(`Row ${number} already has different taps. Review that row in the log.`);
      }
      const note = [reference, `Drive folder: ${link}`].filter(
        value => value && !row[6].includes(value),
      );
      const comments = note.length ? sheet.getRange(number, 7) : null;
      const calibration = !row[3].trim() && taps ? sheet.getRange(number, 4) : null;
      if (comments && comments.getFormula()) throw new Error(`Row ${number} has a formula in Comments.`);
      if (calibration && calibration.getFormula()) throw new Error(`Row ${number} has a formula in Calibration.`);
      if (comments) {
        const prior = comments.getRichTextValue();
        const combined = [row[6], ...note].filter(Boolean).join('\n');
        const rich = SpreadsheetApp.newRichTextValue().setText(combined);
        if (prior && prior.getText() === row[6]) {
          for (const run of prior.getRuns()) {
            const start = run.getStartIndex();
            const end = run.getEndIndex();
            const url = run.getLinkUrl();
            if (url) rich.setLinkUrl(start, end, url);
            rich.setTextStyle(start, end, run.getTextStyle());
          }
        }
        const start = combined.lastIndexOf(link);
        if (start >= 0) rich.setLinkUrl(start, start + link.length, link);
        comments.setRichTextValue(rich.build());
      }
      if (calibration) calibration.setValue(taps);
      return {row: number, action: comments || calibration ? 'Updated' : 'Already logged at'};
    }
    sheet.appendRow([cd, archiveText_(block), '', taps, '', 'RASS',
                     archiveText_([reference, `Drive folder: ${link}`].filter(Boolean).join('\n')), '']);
    return {row: sheet.getLastRow(), action: 'Added'};
  } finally {
    SpreadsheetApp.flush();
    lock.releaseLock();
  }
}
