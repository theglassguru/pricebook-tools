/**
 * Opens the Setup Guide sidebar with links to Guru HQ resources.
 * Wired to ⚙ Pricebook Tools → 📘 Setup Guide.
 */
function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('GeminiSidebar')
    .setTitle('Pricebook Setup Guide');
  SpreadsheetApp.getUi().showSidebar(html);
}
