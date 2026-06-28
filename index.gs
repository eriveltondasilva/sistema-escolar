/**
 * Sistema Escolar — Geração de Boletins
 * Apps Script único (Code.gs)
 */

// ============================================================
// DEFINIÇÕES DE TIPOS (JSDOC)
// ============================================================

/**
 * @typedef {Object} AppConfig
 * @property {string} schoolYearsFolderId
 * @property {string} pdfsFolderId
 * @property {string} tempFolderId
 * @property {string} reportTemplateFileId
 * @property {string} classTemplateFileId
 * @property {string} studentsSpreadsheetId
 */

/**
 * @typedef {Object} Subject
 * @property {string} name
 * @property {string} code
 */

/**
 * @typedef {Object} StudentData
 * @property {string} name
 * @property {string} address
 * @property {string} nationality
 * @property {string} birthDate
 * @property {string} sex
 */

/**
 * @typedef {Object} PersonalData
 * @property {string} name
 * @property {string} address
 * @property {string} nationality
 * @property {string} birthDate
 * @property {string} sex
 * @property {string} guardianNames
 */

/**
 * @typedef {Object} ReportContext
 * @property {number} yearNumber
 * @property {GoogleAppsScript.Drive.File} templateFile
 * @property {GoogleAppsScript.Drive.Folder} tempFolder
 * @property {GoogleAppsScript.Drive.Folder} pdfFolder
 * @property {Map<string, StudentData>} studentsMap
 * @property {Map<string, string[]>} guardiansMap
 * @property {Map<string, any[][]>} gradesBySubject
 */

/**
 * @typedef {Object} PlaceholderField
 * @property {string} suffix
 * @property {string} field
 * @property {(value: any) => string} format
 */

// ============================================================
// CONFIGURAÇÃO
// ============================================================

const DEFAULT_LOCALE = "pt-BR";
const DEFAULT_TIMEZONE = "America/Sao_Paulo";

// primeira linha de dados nas abas de turma/disciplina
const FIRST_DATA_ROW = 5;
// Início da primeira linha de configuração na aba "Configuração"
const CONFIG_START_ROW = 4;

const GRADE_COLUMNS_COUNT = 17;

/** @type {string[]} */
const VALID_CLASSES = ["6º Ano", "7º Ano", "8º Ano", "9º Ano"];

/** @type {Subject[]} */
const VALID_SUBJECTS = [
  { name: "Arte", code: "ART" },
  { name: "Ciências", code: "CIE" },
  { name: "Educação Física", code: "EDF" },
  { name: "Geografia", code: "GEO" },
  { name: "História", code: "HIS" },
  { name: "Língua Inglesa", code: "ING" },
  { name: "Língua Portuguesa", code: "LPO" },
  { name: "Matemática", code: "MAT" },
];

const CONFIG_KEY_MAP = {
  PASTA_ANOS_LETIVOS_ID: "schoolYearsFolderId",
  PASTA_PDFS_ID: "pdfsFolderId",
  PASTA_TEMP_ID: "tempFolderId",
  MODELO_BOLETIM_ID: "reportTemplateFileId",
  MODELO_TURMA_ID: "classTemplateFileId",
  CADASTRO_ALUNOS_ID: "studentsSpreadsheetId",
};

const SHEET_NAMES = {
  CONFIG: "Configuração",
  STUDENTS: "Alunos",
  GUARDIANS: "Responsáveis",
  SUMMARY: "Resumo",
};

// Colunas da aba "Alunos" (0-indexed)
const STUDENT_COLUMNS = {
  id: 0,
  name: 1,
  address: 2,
  nationality: 3,
  birthDate: 4,
  sex: 5,
};

// Colunas da aba "Responsáveis" (0-indexed)
const GUARDIAN_COLUMNS = {
  studentId: 0,
  name: 1,
};

// Colunas da planilha de notas por disciplina (0-indexed dentro da linha)
const GRADE_COLUMNS = {
  grade1Q: 2,
  absences1Q: 3,
  grade2Q: 4,
  absences2Q: 5,
  makeup1S: 6,
  average1S: 7,
  grade3Q: 8,
  absences3Q: 9,
  grade4Q: 10,
  absences4Q: 11,
  makeup2S: 12,
  average2S: 13,
  totalAbsences: 14,
  finalGrade: 15,
  status: 16,
};

/** @type {PlaceholderField[]} */
const SUBJECT_PLACEHOLDER_FIELDS = [
  { suffix: "nota1", field: "grade1Q", format: formatGrade },
  { suffix: "falt1", field: "absences1Q", format: formatValue },
  { suffix: "nota2", field: "grade2Q", format: formatGrade },
  { suffix: "falt2", field: "absences2Q", format: formatValue },
  { suffix: "rec1", field: "makeup1S", format: formatGrade },
  { suffix: "media1", field: "average1S", format: formatGrade },
  { suffix: "nota3", field: "grade3Q", format: formatGrade },
  { suffix: "falt3", field: "absences3Q", format: formatValue },
  { suffix: "nota4", field: "grade4Q", format: formatGrade },
  { suffix: "falt4", field: "absences4Q", format: formatValue },
  { suffix: "rec2", field: "makeup2S", format: formatGrade },
  { suffix: "media2", field: "average2S", format: formatGrade },
  { suffix: "faltastotal", field: "totalAbsences", format: formatValue },
  { suffix: "final", field: "finalGrade", format: formatGrade },
  { suffix: "situacao", field: "status", format: (status) => status ?? "" },
];

// ---

/**
 * Lê e valida as configurações da aba "Configuração".
 * * @returns {AppConfig}
 * @throws {Error} Se a aba não for encontrada ou faltarem configurações.
 */
function loadConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.CONFIG);

  if (!sheet) {
    throw new Error(`Aba "${SHEET_NAMES.CONFIG}" não encontrada.`);
  }

  const configKeysCount = Object.keys(CONFIG_KEY_MAP).length;
  const rows = sheet
    .getRange(CONFIG_START_ROW, 1, configKeysCount, 2)
    .getValues();

  const rawConfig = Object.fromEntries(
    rows
      .map(([key, value]) => [String(key ?? "").trim(), value])
      .filter(([key]) => key.length > 0),
  );

  const missingKeys = Object.keys(CONFIG_KEY_MAP).filter(
    (key) => !rawConfig[key],
  );

  if (missingKeys.length > 0) {
    throw new Error(
      `Configuração incompleta: faltam os valores de ${missingKeys.join(", ")}.`,
    );
  }

  return Object.fromEntries(
    Object.entries(CONFIG_KEY_MAP).map(([sheetKey, codeKey]) => [
      codeKey,
      rawConfig[sheetKey],
    ]),
  );
}

// ============================================================
// MENU
// ============================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("<< Sistema Escolar >>")
    .addItem("Gerar boletim do aluno", "generateStudentReport")
    .addItem("Gerar boletins da turma", "generateClassReports")
    .addSeparator()
    .addItem("Criar ano letivo", "createSchoolYear")
    .addSeparator()
    .addItem("Verificar configuração", "checkConfiguration")
    .addToUi();
}

// ============================================================
// VALIDAÇÃO
// ============================================================

/**
 * Encontra a aba de uma disciplina na planilha de turma.
 * * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} classSpreadsheet
 * @param {Subject} subject
 * @returns {GoogleAppsScript.Spreadsheet.Sheet | null}
 */
function findSubjectSheet(classSpreadsheet, subject) {
  return (
    classSpreadsheet.getSheetByName(subject.code) ??
    classSpreadsheet.getSheetByName(subject.name)
  );
}

/**
 * Confere quais disciplinas esperadas existem como aba na planilha de turma.
 * * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} classSpreadsheet
 * @returns {{found: Subject[], missing: string[]}}
 */
function checkClassSubjects(classSpreadsheet) {
  return VALID_SUBJECTS.reduce(
    (acc, subject) => {
      if (findSubjectSheet(classSpreadsheet, subject)) {
        acc.found.push(subject);
      } else {
        acc.missing.push(subject.name);
      }
      return acc;
    },
    { found: [], missing: [] },
  );
}

/**
 * Lê todos os alunos (matrícula + nome) da aba "Resumo".
 * * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} classSpreadsheet
 * @returns {Array<{studentId: string, name: string, row: number}>}
 */
function getClassStudentsFromResumo(classSpreadsheet) {
  const resumoSheet = classSpreadsheet.getSheetByName(SHEET_NAMES.SUMMARY);
  if (!resumoSheet) return [];

  const lastRow = resumoSheet.getLastRow();
  if (lastRow < FIRST_DATA_ROW) return [];

  const values = resumoSheet
    .getRange(FIRST_DATA_ROW, 1, lastRow - FIRST_DATA_ROW + 1, 2)
    .getValues();

  return values
    .map(([studentId, name], index) => ({
      studentId: String(studentId ?? "").trim(),
      name: String(name ?? "").trim(),
      row: FIRST_DATA_ROW + index,
    }))
    .filter(({ studentId }) => studentId.length > 0);
}

/**
 * Lê a aba "Alunos" do Cadastro e devolve um mapa.
 * * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} registrationSheet
 * @returns {Map<string, StudentData>}
 */
function loadStudentsMap(registrationSheet) {
  const studentsSheet = registrationSheet.getSheetByName(SHEET_NAMES.STUDENTS);
  if (!studentsSheet) {
    throw new Error(
      `Cadastro de Alunos: a aba "${SHEET_NAMES.STUDENTS}" não existe.`,
    );
  }

  const rows = studentsSheet.getDataRange().getValues().slice(1);
  /** @type {Map<string, StudentData>} */
  const map = new Map();

  for (const row of rows) {
    const studentId = String(row[STUDENT_COLUMNS.id] ?? "").trim();
    if (!studentId) continue;

    map.set(studentId, {
      name: String(row[STUDENT_COLUMNS.name] ?? "").trim(),
      address: row[STUDENT_COLUMNS.address],
      nationality: row[STUDENT_COLUMNS.nationality],
      birthDate: formatLongDate(row[STUDENT_COLUMNS.birthDate]),
      sex: row[STUDENT_COLUMNS.sex],
    });
  }

  return map;
}

/**
 * Verifica se há matrículas duplicadas na aba "Alunos".
 * * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} registrationSheet
 * @returns {string[]} Mensagens de erro
 */
function findDuplicateStudentIds(registrationSheet) {
  const studentsSheet = registrationSheet.getSheetByName(SHEET_NAMES.STUDENTS);
  if (!studentsSheet) return [];

  const rows = studentsSheet.getDataRange().getValues().slice(1);
  const rowsByStudentId = new Map();

  rows.forEach((row, index) => {
    const studentId = String(row[STUDENT_COLUMNS.id] ?? "").trim();
    if (!studentId) return;

    const dataRow = index + 2;
    const existingRows = rowsByStudentId.get(studentId) ?? [];
    existingRows.push(dataRow);
    rowsByStudentId.set(studentId, existingRows);
  });

  return [...rowsByStudentId.entries()]
    .filter(([, dataRows]) => dataRows.length > 1)
    .map(
      ([studentId, dataRows]) =>
        `Cadastro de Alunos: matrícula ${studentId} duplicada nas linhas ${dataRows.join(", ")}.`,
    );
}

/**
 * Lê a aba "Responsáveis" do Cadastro e agrupa por matrícula.
 * * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} registrationSheet
 * @returns {Map<string, string[]>}
 */
function loadGuardiansMap(registrationSheet) {
  const guardiansSheet = registrationSheet.getSheetByName(
    SHEET_NAMES.GUARDIANS,
  );
  if (!guardiansSheet) return new Map();

  const rows = guardiansSheet.getDataRange().getValues().slice(1);
  const validRows = rows
    .map((row) => ({
      studentId: String(row[GUARDIAN_COLUMNS.studentId] ?? "").trim(),
      name: row[GUARDIAN_COLUMNS.name],
    }))
    .filter(({ studentId }) => studentId.length > 0);

  const grouped = Map.groupBy(validRows, ({ studentId }) => studentId);

  for (const [studentId, entries] of grouped) {
    grouped.set(
      studentId,
      entries.map(({ name }) => name),
    );
  }

  return grouped;
}

/**
 * Verifica se há matrículas duplicadas na aba "Resumo".
 * * @param {Array<{studentId: string, row: number}>} students
 * @param {number|string} year
 * @param {string} className
 * @returns {string[]}
 */
function findDuplicateResumoIds(students, year, className) {
  const rowsByStudentId = new Map();

  for (const { studentId, row } of students) {
    const existingRows = rowsByStudentId.get(studentId) ?? [];
    existingRows.push(row);
    rowsByStudentId.set(studentId, existingRows);
  }

  return [...rowsByStudentId.entries()]
    .filter(([, rows]) => rows.length > 1)
    .map(
      ([studentId, rows]) =>
        `[${year} / ${className} / Resumo] Matrícula ${studentId} duplicada nas linhas ${rows.join(", ")}.`,
    );
}

/**
 * Compara os alunos da aba "Resumo" de uma turma com o Cadastro de Alunos.
 * * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} classSpreadsheet
 * @param {Map<string, StudentData>} registeredStudentsMap
 * @param {number|string} year
 * @param {string} className
 * @returns {string[]} Array de mensagens de aviso/erro
 */
function validateClassStudents(
  classSpreadsheet,
  registeredStudentsMap,
  year,
  className,
) {
  const issues = [];
  const students = getClassStudentsFromResumo(classSpreadsheet);

  if (students.length === 0) {
    issues.push(
      `[${year} / ${className}] Turma sem alunos cadastrados na aba "${SHEET_NAMES.SUMMARY}".`,
    );
    return issues;
  }

  issues.push(...findDuplicateResumoIds(students, year, className));

  for (const { studentId, name, row } of students) {
    const registeredStudent = registeredStudentsMap.get(studentId);

    if (registeredStudent === undefined) {
      issues.push(
        `[${year} / ${className} / Resumo, linha ${row}] Matrícula ${studentId} não consta no Cadastro de Alunos.`,
      );
      continue;
    }

    const namesDiffer =
      registeredStudent.name.localeCompare(name, DEFAULT_LOCALE, {
        sensitivity: "base",
      }) !== 0;

    if (namesDiffer) {
      issues.push(
        `[${year} / ${className} / Resumo, linha ${row}] Nome "${name}" diverge do Cadastro ("${registeredStudent.name}") para a matrícula ${studentId}.`,
      );
    }
  }

  return issues;
}

/**
 * Verifica todas as configurações e estrutura de pastas.
 * Disparada pelo menu do usuário.
 */
function checkConfiguration() {
  const ui = SpreadsheetApp.getUi();
  let config;

  try {
    config = loadConfig();
  } catch (e) {
    ui.alert(`Erro: ${e.message}`);
    return;
  }

  const issues = [];

  try {
    getReportTemplateFile(config);
  } catch (e) {
    issues.push(e.message);
  }
  try {
    getClassTemplateFile(config);
  } catch (e) {
    issues.push(e.message);
  }

  try {
    DriveApp.getFolderById(config.pdfsFolderId);
  } catch {
    issues.push("PDFs: pasta não encontrada ou sem acesso.");
  }

  let registeredStudentsMap;
  try {
    const registrationSheet = SpreadsheetApp.openById(
      config.studentsSpreadsheetId,
    );

    if (!registrationSheet.getSheetByName(SHEET_NAMES.STUDENTS)) {
      issues.push(
        `Cadastro de Alunos: a aba "${SHEET_NAMES.STUDENTS}" não existe.`,
      );
    }
    if (!registrationSheet.getSheetByName(SHEET_NAMES.GUARDIANS)) {
      issues.push(
        `Cadastro de Alunos: a aba "${SHEET_NAMES.GUARDIANS}" não existe.`,
      );
    }

    registeredStudentsMap = loadStudentsMap(registrationSheet);
    issues.push(...findDuplicateStudentIds(registrationSheet));
  } catch (e) {
    issues.push(`Cadastro de Alunos: ${e.message}`);
  }

  let years = [];
  try {
    years = listSchoolYears(config);
  } catch (e) {
    issues.push(e.message);
  }

  if (years.length === 0) {
    issues.push(
      'Nenhuma pasta de ano letivo encontrada dentro de "Anos Letivos".',
    );
  }

  for (const year of years) {
    let yearFolder;
    try {
      yearFolder = getSchoolYearFolder(config, year);
    } catch (e) {
      issues.push(e.message);
      continue;
    }

    for (const className of VALID_CLASSES) {
      let classFile;
      try {
        classFile = getClassSpreadsheetFile(yearFolder, year, className);
      } catch (e) {
        issues.push(`[${year}] ${e.message}`);
        continue;
      }

      let classSpreadsheet;
      try {
        classSpreadsheet = SpreadsheetApp.openById(classFile.getId());
      } catch (e) {
        issues.push(
          `[${year} / ${className}] Erro ao abrir a planilha: ${e.message}`,
        );
        continue;
      }

      const { missing } = checkClassSubjects(classSpreadsheet);
      if (missing.length > 0) {
        issues.push(
          `[${year} / ${className}] Disciplinas faltando: ${missing.join(", ")}`,
        );
      }

      if (registeredStudentsMap) {
        issues.push(
          ...validateClassStudents(
            classSpreadsheet,
            registeredStudentsMap,
            year,
            className,
          ),
        );
      }
    }
  }

  if (issues.length > 0) {
    ui.alert(
      `Foram encontrados ${issues.length} problema(s):\n\n${issues.join("\n")}`,
    );
    return;
  }

  ui.alert(
    "Tudo certo! Estrutura, disciplinas e alunos validados em todos os anos letivos.",
  );
}

// ============================================================
// BUSCA NO DRIVE
// ============================================================

/**
 * @param {AppConfig} config
 * @return {string[]}
 */
function listSchoolYears(config) {
  const rootFolder = DriveApp.getFolderById(config.schoolYearsFolderId);
  const folderIterator = rootFolder.getFolders();

  const folderNames = [];
  while (folderIterator.hasNext()) {
    const folder = folderIterator.next().getName();
    if (!/\d{4}/.test(folder)) continue;
    folderNames.push(folder);
  }

  return folderNames.sort();
}

/**
 * @param {AppConfig} config
 * @param {string} schoolYearLabel
 * @returns {GoogleAppsScript.Drive.Folder}
 */
function getSchoolYearFolder(config, schoolYearLabel) {
  const rootFolder = DriveApp.getFolderById(config.schoolYearsFolderId);
  const subfolders = rootFolder.getFoldersByName(schoolYearLabel);

  if (!subfolders.hasNext()) {
    throw new Error(
      `Pasta do ano letivo "${schoolYearLabel}" não encontrada dentro de "Anos Letivos".`,
    );
  }

  return subfolders.next();
}

/**
 * @param {string} schoolYearLabel
 * @return {number}
 */
function extractYearNumber(schoolYearLabel) {
  const match = schoolYearLabel.match(/\d{4}/);
  if (!match) {
    throw new Error(
      `Não foi possível identificar um ano de 4 dígitos no nome da pasta "${schoolYearLabel}".`,
    );
  }

  return Number(match[0]);
}

/**
 * @param {GoogleAppsScript.Drive.Folder} yearFolder
 * @param {string} schoolYearLabel
 * @param {string} className
 * @returns {GoogleAppsScript.Drive.File}
 */
function getClassSpreadsheetFile(yearFolder, schoolYearLabel, className) {
  const files = yearFolder.getFilesByName(className);
  if (!files.hasNext()) {
    throw new Error(
      `Planilha da turma "${className}" não encontrada dentro de "Anos Letivos/${schoolYearLabel}".`,
    );
  }

  const file = files.next();
  if (file.getMimeType() !== MimeType.GOOGLE_SHEETS) {
    throw new Error(
      `O arquivo "${className}" em "Anos Letivos/${schoolYearLabel}" não é uma planilha do Google Sheets.`,
    );
  }

  return file;
}

/**
 * @param {AppConfig} config
 * @returns {GoogleAppsScript.Drive.File}
 */
function getReportTemplateFile(config) {
  try {
    return DriveApp.getFileById(config.reportTemplateFileId);
  } catch {
    throw new Error(
      `Modelo de boletim não encontrado (ID: ${config.reportTemplateFileId}).`,
    );
  }
}

/**
 * @param {AppConfig} config
 * @returns {GoogleAppsScript.Drive.File}
 */
function getClassTemplateFile(config) {
  try {
    return DriveApp.getFileById(config.classTemplateFileId);
  } catch {
    throw new Error(
      `Modelo de planilha de turma não encontrado (ID: ${config.classTemplateFileId}).`,
    );
  }
}

/**
 * @param {AppConfig} config
 * @param {string} schoolYearLabel
 * @returns {boolean}
 */
function schoolYearFolderExists(config, schoolYearLabel) {
  const rootFolder = DriveApp.getFolderById(config.schoolYearsFolderId);
  return rootFolder.getFoldersByName(schoolYearLabel).hasNext();
}

/**
 * @param {AppConfig} config
 * @param {number} yearNumber
 * @param {string} className
 * @returns {GoogleAppsScript.Drive.Folder}
 */
function getOrCreateClassPdfFolder(config, yearNumber, className) {
  const rootFolder = DriveApp.getFolderById(config.pdfsFolderId);
  const yearLabel = String(yearNumber);

  const yearFolders = rootFolder.getFoldersByName(yearLabel);
  const yearFolder = yearFolders.hasNext()
    ? yearFolders.next()
    : rootFolder.createFolder(yearLabel);

  const classFolders = yearFolder.getFoldersByName(className);
  return classFolders.hasNext()
    ? classFolders.next()
    : yearFolder.createFolder(className);
}

// ============================================================
// AÇÕES DO MENU
// ============================================================

/**
 * @param {GoogleAppsScript.Base.Ui} ui
 * @param {string} title
 * @param {string} message
 * @returns {string | null}
 */
function promptForValue(ui, title, message) {
  const response = ui.prompt(title, message, ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) return null;
  return response.getResponseText().trim();
}

/**
 * @param {GoogleAppsScript.Base.Ui} ui
 * @param {string} title
 * @param {string[]} years
 * @returns {{schoolYearLabel: string, className: string} | null}
 */
function promptSchoolYearAndClass(ui, title, years) {
  const schoolYearLabel = promptForValue(
    ui,
    title,
    `Digite o ano letivo (opções: ${years.join(", ")}):`,
  );
  if (schoolYearLabel === null) return null;

  if (!years.includes(schoolYearLabel)) {
    ui.alert(
      `"${schoolYearLabel}" não é um ano letivo válido. Opções: ${years.join(", ")}`,
    );
    return null;
  }

  const className = promptForValue(
    ui,
    title,
    `Digite o nome da turma (opções: ${VALID_CLASSES.join(", ")}):`,
  );
  if (className === null) return null;

  if (!VALID_CLASSES.includes(className)) {
    ui.alert(
      `"${className}" não é uma turma válida. Opções: ${VALID_CLASSES.join(", ")}`,
    );
    return null;
  }

  return { schoolYearLabel, className };
}

/**
 * Wrapper com lock para criar o ano letivo
 */
function createSchoolYear() {
  withScriptLock(
    createSchoolYear_,
    "Já existe uma operação em andamento. Tente novamente em alguns instantes.",
  );
}

/**
 * @param {GoogleAppsScript.Base.Ui} ui
 */
function createSchoolYear_(ui) {
  let config;

  try {
    config = loadConfig();
  } catch (e) {
    ui.alert(`Erro na configuração: ${e.message}`);
    return;
  }

  const yearInput = promptForValue(
    ui,
    "Criar ano letivo",
    "Digite o ano letivo (ex: 2026):",
  );
  if (yearInput === null) return;

  if (!/^\d{4}$/.test(yearInput)) {
    ui.alert(`"${yearInput}" não é um ano válido. Digite 4 dígitos, ex: 2026.`);
    return;
  }

  const schoolYearLabel = `Ano Letivo — ${yearInput}`;
  if (schoolYearFolderExists(config, schoolYearLabel)) {
    ui.alert(
      `O ano letivo "${schoolYearLabel}" já existe. Nenhuma alteração foi feita.`,
    );
    return;
  }

  let classTemplateFile;
  try {
    classTemplateFile = getClassTemplateFile(config);
  } catch (e) {
    ui.alert(`Erro: ${e.message}`);
    return;
  }

  const rootFolder = DriveApp.getFolderById(config.schoolYearsFolderId);
  const yearFolder = rootFolder.createFolder(schoolYearLabel);
  const createdClasses = [];
  const errors = [];

  for (const className of VALID_CLASSES) {
    try {
      const classFile = classTemplateFile.makeCopy(className, yearFolder);
      const classSpreadsheet = SpreadsheetApp.openById(classFile.getId());
      fillClassHeaderPlaceholders(classSpreadsheet, className, yearInput);
      createdClasses.push(className);
    } catch (e) {
      errors.push(`${className}: ${e.message}`);
    }
  }

  if (errors.length > 0) {
    // Desfaz a criação para não deixar a pasta do ano num estado parcial:
    // sem isso, uma nova tentativa para o mesmo ano seria bloqueada pela
    // checagem "já existe" acima, mesmo com turmas faltando e sem nenhuma
    // forma de retomar ou completar pelo menu.
    yearFolder.setTrashed(true);
    ui.alert(
      `Não foi possível criar o ano letivo "${schoolYearLabel}". Nenhuma alteração foi feita.\n\n` +
        `Erros:\n${errors.join("\n")}`,
    );
    return;
  }

  ui.alert(
    `Ano letivo "${schoolYearLabel}" criado com ${createdClasses.length} turma(s): ${createdClasses.join(", ")}.`,
  );
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} classSpreadsheet
 * @param {string} className
 * @param {string} yearLabel
 */
function fillClassHeaderPlaceholders(classSpreadsheet, className, yearLabel) {
  for (const sheet of classSpreadsheet.getSheets()) {
    replaceSheetHeaderText(sheet, "{{school_class}}", className);
    replaceSheetHeaderText(sheet, "{{school_year}}", yearLabel);
  }
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {string} placeholder
 * @param {string} value
 */
function replaceSheetHeaderText(sheet, placeholder, value) {
  sheet
    .createTextFinder(placeholder)
    .matchEntireCell(false)
    .replaceAllWith(value);
}

/**
 * Executa uma ação utilizando o LockService do Apps Script.
 * * @param {(ui: GoogleAppsScript.Base.Ui) => void} action
 * @param {string} busyMessage
 */
function withScriptLock(action, busyMessage) {
  const ui = SpreadsheetApp.getUi();
  const lock = LockService.getScriptLock();

  if (!lock.tryLock(5000)) {
    ui.alert(busyMessage);
    return;
  }

  try {
    action(ui);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Wrapper para gerar o boletim de um único aluno
 */
function generateStudentReport() {
  withScriptLock(
    generateStudentReport_,
    "Já existe uma geração de boletim em andamento. Tente novamente em alguns instantes.",
  );
}

/**
 * @param {GoogleAppsScript.Base.Ui} ui
 */
function generateStudentReport_(ui) {
  let config;

  try {
    config = loadConfig();
  } catch (e) {
    ui.alert(`Erro na configuração: ${e.message}`);
    return;
  }

  let years;
  try {
    years = listSchoolYears(config);
  } catch (e) {
    ui.alert(`Erro: ${e.message}`);
    return;
  }

  if (years.length === 0) {
    ui.alert('Nenhum ano letivo encontrado dentro da pasta "Anos Letivos".');
    return;
  }

  const selection = promptSchoolYearAndClass(
    ui,
    "Gerar boletim de um aluno",
    years,
  );
  if (!selection) return;
  const { schoolYearLabel, className } = selection;

  const studentId = promptForValue(
    ui,
    "Gerar boletim de um aluno",
    "Digite a matrícula do aluno:",
  );
  if (studentId === null) return;
  if (!studentId) {
    ui.alert("Matrícula não pode ser vazia.");
    return;
  }

  try {
    const yearFolder = getSchoolYearFolder(config, schoolYearLabel);
    const classFile = getClassSpreadsheetFile(
      yearFolder,
      schoolYearLabel,
      className,
    );
    const classSpreadsheet = SpreadsheetApp.openById(classFile.getId());

    const { found: foundSubjects, missing } =
      checkClassSubjects(classSpreadsheet);
    if (missing.length > 0) {
      ui.alert(
        `Atenção: as seguintes disciplinas não foram encontradas nessa turma e serão ignoradas no boletim:\n` +
          `${missing.join(", ")}\n\nContinuando com as disciplinas disponíveis...`,
      );
    }

    if (foundSubjects.length === 0) {
      ui.alert(
        "Nenhuma disciplina reconhecida nessa turma. Não é possível gerar o boletim.",
      );
      return;
    }

    const rowNumber = findRowByStudentId(
      classSpreadsheet,
      studentId,
      foundSubjects,
    );
    if (rowNumber === null) {
      ui.alert(
        `Matrícula ${studentId} não encontrada na turma "${className}" (${schoolYearLabel}).`,
      );
      return;
    }

    const context = buildSingleStudentReportContext({
      config,
      classSpreadsheet,
      schoolYearLabel,
      className,
      foundSubjects,
      studentId,
    });
    const pdfUrl = generateReportForStudent({
      studentId,
      className,
      foundSubjects,
      context,
    });

    ui.alert(`Boletim gerado com sucesso!\n\n${pdfUrl}`);
  } catch (e) {
    ui.alert(`Erro ao gerar boletim: ${e.message}`);
  }
}

/**
 * Wrapper para gerar todos os boletins da turma
 */
function generateClassReports() {
  withScriptLock(
    generateClassReports_,
    "Já existe uma geração de boletins em andamento. Tente novamente em alguns instantes.",
  );
}

/**
 * @param {GoogleAppsScript.Base.Ui} ui
 */
function generateClassReports_(ui) {
  let config;

  try {
    config = loadConfig();
  } catch (e) {
    ui.alert(`Erro na configuração: ${e.message}`);
    return;
  }

  let years;
  try {
    years = listSchoolYears(config);
  } catch (e) {
    ui.alert(`Erro: ${e.message}`);
    return;
  }

  if (years.length === 0) {
    ui.alert('Nenhum ano letivo encontrado dentro da pasta "Anos Letivos".');
    return;
  }

  const selection = promptSchoolYearAndClass(
    ui,
    "Gerar boletins de uma turma",
    years,
  );
  if (!selection) return;
  const { schoolYearLabel, className } = selection;

  let classSpreadsheet;
  try {
    const yearFolder = getSchoolYearFolder(config, schoolYearLabel);
    const classFile = getClassSpreadsheetFile(
      yearFolder,
      schoolYearLabel,
      className,
    );
    classSpreadsheet = SpreadsheetApp.openById(classFile.getId());
  } catch (e) {
    ui.alert(`Erro: ${e.message}`);
    return;
  }

  const { found: foundSubjects, missing } =
    checkClassSubjects(classSpreadsheet);
  if (missing.length > 0) {
    ui.alert(
      `Atenção: as seguintes disciplinas não foram encontradas e serão ignoradas:\n` +
        `${missing.join(", ")}\n\nContinuando com as disciplinas disponíveis...`,
    );
  }

  if (foundSubjects.length === 0) {
    ui.alert(
      "Nenhuma disciplina reconhecida nessa turma. Não é possível gerar boletins.",
    );
    return;
  }

  const firstSheet = findSubjectSheet(classSpreadsheet, foundSubjects[0]);
  const lastRow = firstSheet.getLastRow();
  const studentIdRows =
    lastRow >= FIRST_DATA_ROW
      ? firstSheet
          .getRange(FIRST_DATA_ROW, 1, lastRow - FIRST_DATA_ROW + 1, 1)
          .getValues()
      : [];

  const context = buildReportContext({
    config,
    classSpreadsheet,
    schoolYearLabel,
    className,
    foundSubjects,
  });

  let successCount = 0;
  const errors = [];

  const startTime = Date.now();
  const MAX_RUNTIME_MS = 5 * 60 * 1_000; // safety margin before the 6-minute limit

  for (const [index, [studentId]] of studentIdRows.entries()) {
    if (Date.now() - startTime > MAX_RUNTIME_MS) {
      errors.push(
        `Execução interrompida por tempo: restam alunos a partir da linha ${FIRST_DATA_ROW + index}.`,
      );
      break;
    }

    if (!studentId) continue;

    const rowNumber = FIRST_DATA_ROW + index;
    try {
      generateReportForStudent({
        studentId: String(studentId),
        className,
        foundSubjects,
        context,
      });
      successCount++;
    } catch (e) {
      errors.push(`Linha ${rowNumber} (matrícula ${studentId}): ${e.message}`);
    }

    Utilities.sleep(200); // pequena folga para evitar erros transitórios de cota no Drive
  }

  if (errors.length > 0) {
    ui.alert(
      `Foram processados com alguns erros.\n\nErros:\n${errors.join("\n")}`,
    );
    return undefined;
  }

  ui.alert(
    `${successCount} boletim(ns) gerado(s) com sucesso para "${className}" (${schoolYearLabel}).`,
  );
}

/**
 * Monta o contexto que `generateReportForStudent` precisa.
 *
 * @param {Object} params
 * @param {AppConfig} params.config
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} params.classSpreadsheet
 * @param {string} params.schoolYearLabel
 * @param {string} params.className
 * @param {Subject[]} params.foundSubjects
 * @returns {ReportContext}
 */
function buildReportContext({
  config,
  classSpreadsheet,
  schoolYearLabel,
  className,
  foundSubjects,
}) {
  const yearNumber = extractYearNumber(schoolYearLabel);
  const registrationSheet = SpreadsheetApp.openById(
    config.studentsSpreadsheetId,
  );

  return {
    yearNumber,
    templateFile: getReportTemplateFile(config),
    tempFolder: DriveApp.getFolderById(config.tempFolderId),
    pdfFolder: getOrCreateClassPdfFolder(config, yearNumber, className),
    studentsMap: loadStudentsMap(registrationSheet),
    guardiansMap: loadGuardiansMap(registrationSheet),
    gradesBySubject: loadGradesBySubject(classSpreadsheet, foundSubjects),
  };
}

/**
 * Versão mais leve de `buildReportContext` para gerar o boletim de UM único
 * aluno: em vez de carregar o cadastro inteiro da escola (Alunos,
 * Responsáveis) e todas as linhas de cada disciplina da turma — útil quando
 * se está gerando para todos os alunos de uma vez —, busca diretamente pela
 * matrícula em cada planilha, lendo só a linha necessária.
 *
 * @param {Object} params
 * @param {AppConfig} params.config
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} params.classSpreadsheet
 * @param {string} params.schoolYearLabel
 * @param {string} params.className
 * @param {Subject[]} params.foundSubjects
 * @param {string} params.studentId
 * @returns {ReportContext}
 */
function buildSingleStudentReportContext({
  config,
  classSpreadsheet,
  schoolYearLabel,
  className,
  foundSubjects,
  studentId,
}) {
  const yearNumber = extractYearNumber(schoolYearLabel);
  const registrationSheet = SpreadsheetApp.openById(
    config.studentsSpreadsheetId,
  );

  return {
    yearNumber,
    templateFile: getReportTemplateFile(config),
    tempFolder: DriveApp.getFolderById(config.tempFolderId),
    pdfFolder: getOrCreateClassPdfFolder(config, yearNumber, className),
    studentsMap: loadSingleStudentMap(registrationSheet, studentId),
    guardiansMap: loadSingleStudentGuardiansMap(registrationSheet, studentId),
    gradesBySubject: loadGradesForSingleStudent(
      classSpreadsheet,
      foundSubjects,
      studentId,
    ),
  };
}

/**
 * Busca um único aluno na aba "Alunos" pela matrícula, sem ler a planilha
 * inteira do Cadastro. Devolve a mesma estrutura de `loadStudentsMap`
 * (com no máximo uma entrada), para ser usado por `getPersonalData`.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} registrationSheet
 * @param {string} studentId
 * @returns {Map<string, StudentData>}
 */
function loadSingleStudentMap(registrationSheet, studentId) {
  const studentsSheet = registrationSheet.getSheetByName(SHEET_NAMES.STUDENTS);
  if (!studentsSheet) {
    throw new Error(
      `Cadastro de Alunos: a aba "${SHEET_NAMES.STUDENTS}" não existe.`,
    );
  }

  const map = new Map();
  const lastRow = studentsSheet.getLastRow();
  if (lastRow < 2) return map;

  const match = studentsSheet
    .getRange(2, STUDENT_COLUMNS.id + 1, lastRow - 1, 1)
    .createTextFinder(studentId)
    .matchEntireCell(true)
    .findNext();
  if (!match) return map;

  const row = studentsSheet
    .getRange(match.getRow(), 1, 1, studentsSheet.getLastColumn())
    .getValues()[0];

  map.set(studentId, {
    name: String(row[STUDENT_COLUMNS.name] ?? "").trim(),
    address: row[STUDENT_COLUMNS.address],
    nationality: row[STUDENT_COLUMNS.nationality],
    birthDate: formatLongDate(row[STUDENT_COLUMNS.birthDate]),
    sex: row[STUDENT_COLUMNS.sex],
  });

  return map;
}

/**
 * Busca os responsáveis de um único aluno na aba "Responsáveis" pela
 * matrícula, sem ler a planilha inteira do Cadastro.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} registrationSheet
 * @param {string} studentId
 * @returns {Map<string, string[]>}
 */
function loadSingleStudentGuardiansMap(registrationSheet, studentId) {
  const guardiansSheet = registrationSheet.getSheetByName(
    SHEET_NAMES.GUARDIANS,
  );
  const map = new Map();
  if (!guardiansSheet) return map;

  const lastRow = guardiansSheet.getLastRow();
  if (lastRow < 2) return map;

  const matches = guardiansSheet
    .getRange(2, GUARDIAN_COLUMNS.studentId + 1, lastRow - 1, 1)
    .createTextFinder(studentId)
    .matchEntireCell(true)
    .findAll();

  const names = matches.map((cell) =>
    guardiansSheet
      .getRange(cell.getRow(), GUARDIAN_COLUMNS.name + 1, 1, 1)
      .getValue(),
  );

  if (names.length > 0) map.set(studentId, names);
  return map;
}

/**
 * Carrega as notas de um único aluno em cada disciplina da turma, lendo
 * apenas a linha correspondente em cada aba (em vez da aba inteira).
 * Devolve o mesmo formato de `loadGradesBySubject` (mapa por disciplina ->
 * mapa por matrícula -> linha), para ser consumido por `getGradesForStudent`
 * sem nenhuma alteração.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} classSpreadsheet
 * @param {Subject[]} foundSubjects
 * @param {string} studentId
 * @returns {Map<string, Map<string, any[]>>}
 */
function loadGradesForSingleStudent(
  classSpreadsheet,
  foundSubjects,
  studentId,
) {
  const map = new Map();

  for (const subject of foundSubjects) {
    const sheet = findSubjectSheet(classSpreadsheet, subject);
    const byStudentId = new Map();

    const lastRow = sheet?.getLastRow() ?? 0;
    if (sheet && lastRow >= FIRST_DATA_ROW) {
      const match = sheet
        .getRange(FIRST_DATA_ROW, 1, lastRow - FIRST_DATA_ROW + 1, 1)
        .createTextFinder(studentId)
        .matchEntireCell(true)
        .findNext();

      if (match) {
        const rowValues = sheet
          .getRange(match.getRow(), 1, 1, GRADE_COLUMNS_COUNT)
          .getValues()[0];
        byStudentId.set(studentId, rowValues);
      }
    }

    map.set(subject.name, byStudentId);
  }

  return map;
}

/**
 * @param {Object} params
 * @param {string} params.studentId
 * @param {string} params.className
 * @param {Subject[]} params.foundSubjects
 * @param {ReportContext} params.context
 * @returns {string} PDF URL
 */
function generateReportForStudent({
  studentId,
  className,
  foundSubjects,
  context,
}) {
  const personalData = getPersonalData(studentId, context);
  const gradesData = getGradesForStudent(studentId, foundSubjects, context);

  const fileName = `${studentId}_${personalData.name.replace(/\s+/g, "_").toLowerCase()}`;

  // A cópia de trabalho nasce em _temp, não na pasta final do PDF
  const docCopy = context.templateFile.makeCopy(fileName, context.tempFolder);

  try {
    const doc = DocumentApp.openById(docCopy.getId());
    const body = doc.getBody();

    replacePlaceholder(body, "nome", personalData.name);
    replacePlaceholder(body, "filiacao", personalData.guardianNames);
    replacePlaceholder(body, "endereco", personalData.address);
    replacePlaceholder(body, "data_nascimento", personalData.birthDate);
    replacePlaceholder(body, "nacionalidade", personalData.nationality);
    replacePlaceholder(body, "serie", className);
    replacePlaceholder(body, "turno", "");
    replacePlaceholder(body, "ano_letivo", String(context.yearNumber));
    replacePlaceholder(body, "data_entrega", formatLongDate(new Date()));

    for (const subject of foundSubjects) {
      const grades = gradesData[subject.name] ?? {};
      fillSubjectPlaceholders(body, subject.code, grades);
    }

    doc.saveAndClose();
    const pdfBlob = docCopy.getAs("application/pdf");

    const pdfFile = context.pdfFolder
      .createFile(pdfBlob)
      .setName(`${fileName}.pdf`);

    return pdfFile.getUrl();
  } finally {
    docCopy.setTrashed(true);
  }
}

/**
 * @param {GoogleAppsScript.Document.Body} body
 * @param {string} subjectCode
 * @param {Object} grades
 */
function fillSubjectPlaceholders(body, subjectCode, grades) {
  for (const { suffix, field, format } of SUBJECT_PLACEHOLDER_FIELDS) {
    replacePlaceholder(body, `${subjectCode}_${suffix}`, format(grades[field]));
  }
}

// ============================================================
// LEITURA DE DADOS
// ============================================================

/**
 * @param {string} studentId
 * @param {ReportContext} context
 * @returns {PersonalData}
 */
function getPersonalData(studentId, context) {
  const student = context.studentsMap.get(studentId);

  if (!student) {
    throw new Error(
      `Aluno com matrícula ${studentId} não encontrado no Cadastro de Alunos.`,
    );
  }

  const guardianNames = context.guardiansMap.get(studentId) ?? [];

  return { ...student, guardianNames: formatGuardianNames(guardianNames) };
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} classSpreadsheet
 * @param {Subject[]} foundSubjects
 * @returns {Map<string, Map<string, any[]>>}
 */
function loadGradesBySubject(classSpreadsheet, foundSubjects) {
  const map = new Map();

  for (const subject of foundSubjects) {
    const sheet = findSubjectSheet(classSpreadsheet, subject);
    if (!sheet) continue;

    const lastRow = sheet.getLastRow();
    const rows =
      lastRow >= FIRST_DATA_ROW
        ? sheet
            .getRange(
              FIRST_DATA_ROW,
              1,
              lastRow - FIRST_DATA_ROW + 1,
              GRADE_COLUMNS_COUNT,
            )
            .getValues()
        : [];

    // column 0 holds the student ID, just like "Resumo" — indexing by it
    // (instead of by row position) avoids assigning a student's grades to
    // a different student if a subject sheet ever gets sorted, has a row
    // added/removed, or otherwise loses alignment with the other sheets.
    const byStudentId = new Map(
      rows
        .map((row) => [String(row[0] ?? "").trim(), row])
        .filter(([studentId]) => studentId.length > 0),
    );

    map.set(subject.name, byStudentId);
  }

  return map;
}

/**
 * @param {string} studentId
 * @param {Subject[]} foundSubjects
 * @param {ReportContext} context
 * @returns {Record<string, Object | null>}
 */
function getGradesForStudent(studentId, foundSubjects, context) {
  const result = {};

  for (const subject of foundSubjects) {
    const rowValues = context.gradesBySubject.get(subject.name)?.get(studentId);

    result[subject.name] = rowValues
      ? Object.fromEntries(
          Object.entries(GRADE_COLUMNS).map(([field, index]) => [
            field,
            rowValues[index],
          ]),
        )
      : null;
  }

  return result;
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} classSpreadsheet
 * @param {string} studentId
 * @param {Subject[]} foundSubjects
 * @returns {number | null}
 */
function findRowByStudentId(classSpreadsheet, studentId, foundSubjects) {
  const sheet = findSubjectSheet(classSpreadsheet, foundSubjects[0]);
  if (!sheet) return null;

  const lastRow = sheet.getLastRow();
  if (lastRow < FIRST_DATA_ROW) return null;

  const ids = sheet
    .getRange(FIRST_DATA_ROW, 1, lastRow - FIRST_DATA_ROW + 1, 1)
    .getValues();

  const index = ids.findIndex(([id]) => String(id) === studentId);
  return index === -1 ? null : FIRST_DATA_ROW + index;
}

// ============================================================
// UTILITÁRIOS
// ============================================================

/**
 * Substitui um termo (placeholder) pelo valor dentro do corpo do documento.
 *
 * O `Body.replaceText` usa a biblioteca RE2 do Google, que também interpreta
 * "$" no texto de substituição como referência a grupo de captura (ex: "$1").
 * Como o valor pode vir de dados digitados na planilha (nome, endereço etc.),
 * escapamos "$" para garantir que ele sempre apareça como texto literal.
 *
 * @param {GoogleAppsScript.Document.Body} body
 * @param {string} key
 * @param {string | null | undefined} value
 */
function replacePlaceholder(body, key, value) {
  const safeValue = String(value ?? "").replace(/\$/g, "$$$$");
  body.replaceText("{{" + key + "}}", safeValue);
}

/**
 * Formata os nomes dos responsáveis concatenando-os corretamente.
 * * @param {string[]} names
 * @return {string}
 */
function formatGuardianNames(names) {
  if (names.length === 0) return "";
  if (names.length === 1) return names[0];

  return new Intl.ListFormat(DEFAULT_LOCALE, {
    style: "long",
    type: "conjunction",
  }).format(names);
}

/**
 * Formata os números que representam as notas com até 2 casas decimais.
 * * @param {any} value
 * @returns {string}
 */
function formatGrade(value) {
  if (value === "" || value === null || value === undefined) return "";
  const number = Number(value);

  if (Number.isNaN(number)) return String(value);

  return new Intl.NumberFormat(DEFAULT_LOCALE, {
    minimumFractionDigits: 1,
    maximumFractionDigits: 2,
  }).format(number);
}

/**
 * Formata um valor simples, inserindo "---" caso esteja vazio.
 * * @param {any} value
 * @return {string}
 */
function formatValue(value) {
  return value === "" || value === null || value === undefined
    ? "---"
    : String(value);
}

/**
 * Retorna a data no formato extenso e de acordo com as opções nativas do objeto Date.
 * * @param {Date} date
 * @param {Intl.DateTimeFormatOptions} [options]
 * @return {string}
 */
function formatLongDate(date, options) {
  if (!date || !(date instanceof Date)) return "";

  return new Intl.DateTimeFormat(DEFAULT_LOCALE, {
    dateStyle: "long",
    timeZone: DEFAULT_TIMEZONE,
    ...options,
  }).format(date);
}
