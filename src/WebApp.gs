/**
 * Handles incoming HTTP GET requests to the Web App.
 * Expected URL format: https://script.google.com/macros/s/WEB_APP_ID/exec?studentId=123&year=2026
 *
 * @param {GoogleAppsScript.Events.DoGet} params
 * @returns {GoogleAppsScript.HTML.HtmlOutput}
 */
function doGet({ parameter }) {
  const { studentId, year } = parameter;

  if (!studentId || !year) {
    return renderError(
      "Matrícula e Ano são obrigatórios para realizar a consulta.",
    );
  }

  try {
    const fileId = findReportPdfId(studentId);

    if (!fileId) {
      return renderError(
        "O boletim ainda não foi gerado ou não foi encontrado para esta matrícula.",
      );
    }

    const template = HtmlService.createTemplateFromFile("RedirectDialog");
    template.downloadUrl = `https://drive.google.com/uc?export=download&id=${fileId}`;

    return template
      .evaluate()
      .setTitle("Boletim Escolar")
      .addMetaTag("viewport", "width=device-width, initial-scale=1.0");
  } catch (error) {
    return renderError("Erro inesperado: " + error.message);
  }
}

/**
 * Função auxiliar para padronizar a renderização de erros
 * @param {string} message
 */
function renderError(message) {
  const template = HtmlService.createTemplateFromFile("ErrorDialog");
  template.errorMessage = message;

  return template
    .evaluate()
    .setTitle("Erro no Sistema")
    .addMetaTag("viewport", "width=device-width, initial-scale=1.0");
}

/**
 * Searches Google Drive for the student's PDF report card.
 *
 * @param {string} studentId
 * @returns {string|null} The Google Drive File ID or null if not found.
 */
function findReportPdfId(studentId) {
  const searchQuery = `title contains '${studentId}_' and mimeType = 'application/pdf' and trashed = false`;
  const files = DriveApp.searchFiles(searchQuery);

  return files.hasNext() ? files.next().getId() : null;
}
