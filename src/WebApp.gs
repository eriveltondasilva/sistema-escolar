/**
 * IMPORTANTE: ATUALIZAÇÃO DE PRODUÇÃO
 * Sempre que alterar o código deste Web App, gerar uma nova versão é obrigatório.
 * Salvar o arquivo apenas atualiza o rascunho de desenvolvimento.
 *
 * Passo a passo para aplicar as alterações na URL pública:
 * 1. Clique em "Implantar" (canto superior direito) > "Gerenciar implantações".
 * 2. Selecione a implantação ativa e clique no ícone de Lápis (Editar).
 * 3. No campo "Versão", altere para "Nova versão".
 * 4. Adicione uma descrição (opcional) e clique em "Implantar".
 */

/**
 * @param {GoogleAppsScript.Events.DoGet} params
 * @returns {GoogleAppsScript.HTML.HtmlOutput}
 */
function doGet({ parameter }) {
  const { studentId, year, className } = parameter;

  if (!studentId || !year || !className) {
    return renderError("Matrícula, Ano e Turma são obrigatórios para realizar a consulta.");
  }

  try {
    const fileId = findReportPdfId(studentId, year, className);

    if (!fileId) {
      return renderError(
        "O boletim ainda não foi gerado ou não foi encontrado para esta matrícula na turma informada.",
      );
    }

    const template = HtmlService.createTemplateFromFile("RedirectDialog");
    template.downloadUrl = `https://drive.google.com/uc?export=download&id=${fileId}`;

    return template
      .evaluate()
      .setTitle("Boletim Escolar")
      .addMetaTag("viewport", "width=device-width, initial-scale=1.0");
  } catch (error) {
    return renderError(`Erro inesperado: ${error.message}`);
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
 * @param {string} className
 * @returns {string|null}
 */
function findReportPdfId(studentId, year, className) {
  let config;

  try {
    config = loadConfig();
  } catch (e) {
    throw new Error(`Erro ao carregar configuração: ${e.message}`);
  }

  const pdfFolder = DriveApp.getFolderById(config.pdfsFolderId);
  const yearFolderIterator = pdfFolder.getFoldersByName(year);

  if (!yearFolderIterator.hasNext()) return null;

  const yearFolder = yearFolderIterator.next();
  const classFolderIterator = yearFolder.getFoldersByName(className);

  if (!classFolderIterator.hasNext()) return null;

  const classFolder = classFolderIterator.next();

  const prefix = `${studentId}_`;
  const searchQuery = `title contains '${prefix}' and mimeType = 'application/pdf' and trashed = false`;

  const files = classFolder.searchFiles(searchQuery);

  while (files.hasNext()) {
    const file = files.next();
    if (file.getName().startsWith(prefix)) {
      return file.getId();
    }
  }

  return null;
}
