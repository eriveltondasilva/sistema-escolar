/**
 * @param {GoogleAppsScript.Events.DoGet} params
 * @returns {GoogleAppsScript.HTML.HtmlOutput}
 */
function doGet({ parameter }) {
  const { studentId, year } = parameter;

  if (!studentId || !year) {
    return renderError("Matrícula e Ano são obrigatórios para realizar a consulta.");
  }

  try {
    const fileId = findReportPdfId(studentId, year);

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
 * @param {string} studentId
 * @param {string} year
 * @returns {string|null}
 */
function findReportPdfId(studentId, year) {
  let config;
  try {
    config = loadConfig();
  } catch (e) {
    throw new Error(`Erro ao carregar configuração: ${e.message}`);
  }

  const pdfFolder = DriveApp.getFolderById(config.pdfsFolderId);
  const yearFolder = pdfFolder.getFoldersByName(year);

  if (!yearFolder.hasNext()) return null;

  const searchQuery = `title contains '${studentId}_' and mimeType = 'application/pdf' and trashed = false`;
  const files = yearFolder.next().searchFiles(searchQuery);

  return files.hasNext() ? files.next().getId() : null;
}
