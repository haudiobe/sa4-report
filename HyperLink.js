function addHyperlinksToDocTables() {
  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();
  const baseUrl = 'https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/TSGS4_136_Montreal/Docs/';
  const pattern = /^S4-26\d{4}$/;

  const tables = body.getTables();

  tables.forEach(table => {
    const numRows = table.getNumRows();
    for (let i = 0; i < numRows; i++) {
      const row = table.getRow(i);
      if (row.getNumCells() < 2) continue; // Ensure at least 2 columns
      const cell = row.getCell(1); // Second column (index 1)
      const text = cell.getText().trim();

      if (pattern.test(text)) {
        const textElement = cell.editAsText();
        const existingUrl = textElement.getLinkUrl(0);
        if (!existingUrl) {
          const fullUrl = baseUrl + text + '.zip';
          textElement.setLinkUrl(fullUrl);
        }
      }
    }
  });
}

