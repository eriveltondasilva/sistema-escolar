/**
 * Sistema Escolar — Geração de Boletins
 * Apps Script único (Code.gs)
 */

// ============================================================
// TIPOS
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
 * @typedef {Object} Student
 * @property {string} name
 * @property {string} address
 * @property {string} nationality
 * @property {string} birthDate
 * @property {string} sex
 */

/**
 * @typedef {Student} PersonalData
 * @property {string} guardianNames
 */

/**
 * @typedef {Object} ReportContext
 * @property {number} yearNumber
 * @property {GoogleAppsScript.Drive.File} templateFile
 * @property {GoogleAppsScript.Drive.Folder} tempFolder
 * @property {GoogleAppsScript.Drive.Folder} pdfFolder
 * @property {Map<string, Student>} studentsMap
 * @property {Map<string, string[]>} guardiansMap
 * @property {Map<string, any[][]>} gradesBySubject
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

const VALID_CLASSES = ["6º Ano", "7º Ano", "8º Ano", "9º Ano"];

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

// Colunas da aba "Alunos" (0-indexed) — confirme contra a planilha "Cadastro de Alunos"
const STUDENT_COLUMNS = {
  id: 0, // Matriculas
  name: 1, // Nome completo
  address: 2, // Endereço
  nationality: 3, // Nacionalidade
  birthDate: 4, // Data de Nasc.
  sex: 5, // Sexo
};

// Colunas da aba "Responsáveis" (0-indexed)
const GUARDIAN_COLUMNS = {
  studentId: 0, // Matriculas
  name: 1, // Nome Completo
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
 * @return {AppConfig}
 * @throws {Error} Se faltarem chaves na configuração
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
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} classSpreadsheet
 * @param {Subject} subject
 *
 * @returns {GoogleAppsScript.Spreadsheet.Sheet | null}
 */
function findSubjectSheet(classSpreadsheet, subject) {
  return (
    classSpreadsheet.getSheetByName(subject.code) ??
    classSpreadsheet.getSheetByName(subject.name)
  );
}

/**
 * Confere quais disciplinas esperadas existem como aba na planilha de turma
 * (por nome completo ou por code).
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
 * Lê todos os alunos (matrícula + nome) da aba "Resumo" de uma planilha de turma,
 * a partir de FIRST_DATA_ROW, ignorando linhas sem matrícula.
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
 * Lê a aba "Alunos" do Cadastro.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} registrationSheet
 * @returns {Map<string, Student>}
 */
function loadStudentsMap(registrationSheet) {
  const studentsSheet = registrationSheet.getSheetByName(SHEET_NAMES.STUDENTS);
  if (!studentsSheet) {
    throw new Error(
      `Cadastro de Alunos: a aba "${SHEET_NAMES.STUDENTS}" não existe.`,
    );
  }

  const rows = studentsSheet.getDataRange().getValues().slice(1);

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
 * Verifica se há matrículas duplicadas na aba "Alunos" do Cadastro.
 * Retorna uma mensagem por matrícula duplicada, com todas as linhas onde ela aparece.
 */
function findDuplicateStudentIds(registrationSheet) {
  const studentsSheet = registrationSheet.getSheetByName(SHEET_NAMES.STUDENTS);
  if (!studentsSheet) return [];

  const rows = studentsSheet.getDataRange().getValues().slice(1);
  const rowsByStudentId = new Map();

  rows.forEach((row, index) => {
    const studentId = String(row[STUDENT_COLUMNS.id] ?? "").trim();
    if (!studentId) return;

    const dataRow = index + 2; // +1 pelo cabeçalho, +1 por ser 1-indexed
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
 * Lê a aba "Responsáveis" do Cadastro UMA ÚNICA VEZ e agrupa os nomes por
 * matrícula, em vez de reler e refiltrar a aba inteira a cada aluno.
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
 * Verifica se há matrículas duplicadas na aba "Resumo" de uma turma.
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
 * Reporta matrículas não encontradas e divergências de nome (case/acento-insensível).
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
 *
 * @param config
 * @return {string[]}
 */
function listSchoolYears(config) {
  const rootFolder = DriveApp.getFolderById(config.schoolYearsFolderId);
  const folderIterator = rootFolder.getFolders();

  const folderNames = [];
  while (folderIterator.hasNext()) {
    const folder = folderIterator.next().getName();
    if (!folder.test(/\d{4}/)) continue;
    folderNames.push(folder);
  }

  return folderNames.sort();
}

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

function getReportTemplateFile(config) {
  try {
    return DriveApp.getFileById(config.reportTemplateFileId);
  } catch {
    throw new Error(
      `Modelo de boletim não encontrado (ID: ${config.reportTemplateFileId}).`,
    );
  }
}

function getClassTemplateFile(config) {
  try {
    return DriveApp.getFileById(config.classTemplateFileId);
  } catch {
    throw new Error(
      `Modelo de planilha de turma não encontrado (ID: ${config.classTemplateFileId}).`,
    );
  }
}

function schoolYearFolderExists(config, schoolYearLabel) {
  const rootFolder = DriveApp.getFolderById(config.schoolYearsFolderId);
  return rootFolder.getFoldersByName(schoolYearLabel).hasNext();
}

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

function promptForValue(ui, title, message) {
  const response = ui.prompt(title, message, ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) return null;
  return response.getResponseText().trim();
}

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
 * Cria a estrutura completa de um novo ano letivo:
 * pasta "Ano Letivo — AAAA" + uma planilha por turma (cópia do modelo).
 * Recusa se a pasta do ano já existir — não sobrescreve nem completa.
 */
function createSchoolYear() {
  withScriptLock(
    createSchoolYear_,
    "Já existe uma operação em andamento. Tente novamente em alguns instantes.",
  );
}

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

  let message = `Ano letivo "${schoolYearLabel}" criado com ${createdClasses.length} turma(s): ${createdClasses.join(", ")}.`;
  if (errors.length > 0) {
    message += `\n\nErros:\n${errors.join("\n")}`;
  }

  ui.alert(message);
}

function fillClassHeaderPlaceholders(classSpreadsheet, className, yearLabel) {
  for (const sheet of classSpreadsheet.getSheets()) {
    replaceSheetHeaderText(sheet, "{{school_class}}", className);
    replaceSheetHeaderText(sheet, "{{school_year}}", yearLabel);
  }
}

function replaceSheetHeaderText(sheet, placeholder, value) {
  sheet
    .createTextFinder(placeholder)
    .matchEntireCell(false)
    .replaceAllWith(value);
}

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

function generateStudentReport() {
  withScriptLock(
    generateStudentReport_,
    "Já existe uma geração de boletim em andamento. Tente novamente em alguns instantes.",
  );
}

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
    const { found, missing } = checkClassSubjects(classSpreadsheet);

    if (missing.length > 0) {
      ui.alert(
        `Atenção: as seguintes disciplinas não foram encontradas nessa turma e serão ignoradas no boletim:\n` +
          `${missing.join(", ")}\n\nContinuando com as disciplinas disponíveis...`,
      );
    }

    if (found.length === 0) {
      ui.alert(
        "Nenhuma disciplina reconhecida nessa turma. Não é possível gerar o boletim.",
      );
      return;
    }

    const rowNumber = findRowByStudentId(classSpreadsheet, studentId, found);
    if (rowNumber === null) {
      ui.alert(
        `Matrícula ${studentId} não encontrada na turma "${className}" (${schoolYearLabel}).`,
      );
      return;
    }

    const context = buildReportContext(
      config,
      classSpreadsheet,
      schoolYearLabel,
      className,
      found,
    );
    const pdfUrl = generateReportForStudent(
      studentId,
      rowNumber,
      className,
      found,
      context,
    );
    ui.alert(`Boletim gerado com sucesso!\n\n${pdfUrl}`);
  } catch (e) {
    ui.alert(`Erro ao gerar boletim: ${e.message}`);
  }
}

function generateClassReports() {
  withScriptLock(
    generateClassReports_,
    "Já existe uma geração de boletins em andamento. Tente novamente em alguns instantes.",
  );
}

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

  const { found, missing } = checkClassSubjects(classSpreadsheet);
  if (missing.length > 0) {
    ui.alert(
      `Atenção: as seguintes disciplinas não foram encontradas e serão ignoradas:\n` +
        `${missing.join(", ")}\n\nContinuando com as disciplinas disponíveis...`,
    );
  }

  if (found.length === 0) {
    ui.alert(
      "Nenhuma disciplina reconhecida nessa turma. Não é possível gerar boletins.",
    );
    return;
  }

  const firstSheet = findSubjectSheet(classSpreadsheet, found[0]);
  const lastRow = firstSheet.getLastRow();

  const studentIdRows =
    lastRow >= FIRST_DATA_ROW
      ? firstSheet
          .getRange(FIRST_DATA_ROW, 1, lastRow - FIRST_DATA_ROW + 1, 1)
          .getValues()
      : [];

  const context = buildReportContext(
    config,
    classSpreadsheet,
    schoolYearLabel,
    className,
    found,
  );

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

    const row = FIRST_DATA_ROW + index;

    try {
      generateReportForStudent(
        String(studentId),
        row,
        className,
        found,
        context,
      );
      successCount++;
    } catch (e) {
      errors.push(`Linha ${row} (matrícula ${studentId}): ${e.message}`);
    }

    Utilities.sleep(200); // pequena folga para evitar erros transitórios de cota no Drive
  }

  if (errors.length > 0) {
    ui.alert(`${summary}\n\nErros:\n${errors.join("\n")}`);
    return undefined;
  }

  ui.alert(
    `${successCount} boletim(ns) gerado(s) com sucesso para "${className}" (${schoolYearLabel}).`,
  );
}

/**
 * Monta, uma única vez por turma, tudo que generateReportForStudent precisa:
 * cadastro de alunos, responsáveis, notas de cada disciplina, modelo do boletim
 * e pasta de destino do PDF. Isso evita reabrir planilhas e refazer buscas no
 * Drive a cada aluno gerado.
 */
function buildReportContext(
  config,
  classSpreadsheet,
  schoolYearLabel,
  className,
  foundSubjects,
) {
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

function generateReportForStudent(
  studentId,
  rowNumber,
  className,
  foundSubjects,
  context,
) {
  const personalData = getPersonalData(studentId, context);
  const gradesData = getGradesForRow(rowNumber, foundSubjects, context);

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

function fillSubjectPlaceholders(body, subjectCode, grades) {
  for (const { suffix, field, format } of SUBJECT_PLACEHOLDER_FIELDS) {
    replacePlaceholder(body, `${subjectCode}_${suffix}`, format(grades[field]));
  }
}

// ============================================================
// LEITURA DE DADOS
// ============================================================

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

    map.set(subject.name, rows);
  }

  return map;
}

function getGradesForRow(rowNumber, foundSubjects, context) {
  const result = {};

  for (const subject of foundSubjects) {
    const rows = context.gradesBySubject.get(subject.name);
    const rowValues = rows?.[rowNumber - FIRST_DATA_ROW];

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
 * @param body
 * @param key
 * @param value
 */
function replacePlaceholder(body, key, value) {
  body.replaceText("{{" + key + "}}", value ?? "");
}

/**
 * @param {string[]} names
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
 * @param value
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
 * @param value
 * @return {string}
 */
function formatValue(value) {
  return value === "" || value === null || value === undefined
    ? "---"
    : String(value);
}

/**
 * @param {Date} date
 * @param {Intl.DateTimeFormatOptions} options
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
