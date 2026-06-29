/**
 * Sistema Escolar — Geração de Boletins
 * Apps Script único (Code.gs)
 *
 * GUIA DE SEPARAÇÃO PARA PRODUÇÃO
 * --------------------------------
 * O Apps Script trata todos os arquivos .gs de um projeto como um único
 * escopo global — não importa em quantos arquivos o código está dividido.
 * Por isso, esta divisão abaixo é puramente organizacional: quando for
 * migrar para múltiplos arquivos, basta copiar cada bloco demarcado por
 * "ARQUIVO: nome.gs" para um arquivo novo com esse nome, na mesma ordem.
 *
 *   1. Config.gs           — tipos (JSDoc), constantes e leitura da
 *                             aba "Configuração"
 *   2. Menu.gs              — menu da planilha (onOpen)
 *   3. DriveLookup.gs       — resolução de pastas/arquivos no Drive
 *                             (anos letivos, planilhas de turma, modelos,
 *                             pasta de PDFs)
 *   4. DataAccess.gs        — leitura de dados (Alunos, Responsáveis,
 *                             notas por disciplina); usada tanto pela
 *                             validação quanto pela geração de boletins
 *   5. Validation.gs        — regras de negócio de validação cruzada
 *                             (checkConfiguration)
 *   6. Actions.gs           — handlers do menu e orquestração de UI
 *                             (prompts, lock, criar ano letivo, gerar
 *                             boletim individual/da turma)
 *   7. ReportGeneration.gs  — montagem do contexto e geração do PDF
 *                             do boletim em si
 *   8. Utils.gs             — formatação de valores e substituição de
 *                             placeholders no documento
 */

// ============================================================
// ARQUIVO: Config.gs
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
 * @property {Map<string, Map<string, any[]>>} gradesBySubject
 */

/**
 * @typedef {Object} PlaceholderField
 * @property {string} suffix
 * @property {string} field
 * @property {(value: any) => string} format
 */

/**
 * @typedef {Object} ValidClass
 * @property {string} className
 * @property {string} stage
 * @property {string} shift
 */

/**
 * @typedef {Object} Issues
 * @property {"warning" | "error"} type
 * @property {string} message
 * @property {string} url
 *
 */

const DEFAULT_LOCALE = "pt-BR";
const DEFAULT_TIMEZONE = "America/Sao_Paulo";
const SCHOOL_YEAR_LABEL_PREFIX = "Ano Letivo - ";

const CONFIG_START_ROW = 4;
const SUMMARY_FIRST_DATA_ROW = 4;
const FIRST_DATA_ROW = 5;

const GRADE_COLUMNS_COUNT = 17;

/**
 * Turmas únicas, não insira duas vezes o mesmo className.
 * @type {ValidClass[]}
 */
const VALID_CLASSES = [
  { className: "6º Ano", stage: "Ensino Fundamental II", shift: "Vespertino" },
  { className: "7º Ano", stage: "Ensino Fundamental II", shift: "Vespertino" },
  { className: "8º Ano", stage: "Ensino Fundamental II", shift: "Vespertino" },
  { className: "9º Ano", stage: "Ensino Fundamental II", shift: "Vespertino" },
];

/** @type {Subject[]} */
const VALID_SUBJECTS = [
  { name: "Arte", code: "ART" },
  { name: "Ciências", code: "CIE" },
  { name: "Educação Física", code: "EDF" },
  { name: "Ensino Religioso", code: "REL" },
  { name: "Geografia", code: "GEO" },
  { name: "História", code: "HIS" },
  { name: "Língua Inglesa", code: "ING" },
  { name: "Língua Portuguesa", code: "LPO" },
  { name: "Matemática", code: "MAT" },
];

/** @type {PlaceholderField[]} */
const SUBJECT_PLACEHOLDER_FIELDS = [
  { suffix: "n1", field: "grade1Q", format: formatGrade },
  { suffix: "f1", field: "absences1Q", format: formatValue },
  { suffix: "n2", field: "grade2Q", format: formatGrade },
  { suffix: "f2", field: "absences2Q", format: formatValue },
  { suffix: "rs1", field: "makeup1S", format: formatGrade },
  { suffix: "ms1", field: "average1S", format: formatGrade },
  { suffix: "n3", field: "grade3Q", format: formatGrade },
  { suffix: "f3", field: "absences3Q", format: formatValue },
  { suffix: "n4", field: "grade4Q", format: formatGrade },
  { suffix: "f4", field: "absences4Q", format: formatValue },
  { suffix: "rs2", field: "makeup2S", format: formatGrade },
  { suffix: "ms2", field: "average2S", format: formatGrade },
  { suffix: "mf", field: "finalGrade", format: formatGrade },
  { suffix: "tf", field: "totalAbsences", format: formatValue },
  {
    suffix: "sf",
    field: "status",
    format: (status) => status.slice(0, 3).toUpperCase() ?? "",
  },
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

/**
 * Lê e valida as configurações da aba "Configuração".
 *
 * @returns {AppConfig}
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
// ARQUIVO: Menu.gs
// ============================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Sistema Escolar")
    .addItem("Gerar boletim do aluno", "generateStudentReport")
    .addItem("Gerar boletins da turma", "generateClassReports")
    .addSeparator()
    .addItem("Criar ano letivo", "createSchoolYear")
    .addSeparator()
    .addItem("Verificar configuração", "checkConfiguration")
    .addToUi();
}

// ============================================================
// ARQUIVO: DriveLookup.gs
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
// ARQUIVO: DataAccess.gs
// ------------------------------------------------------------
// Leitura de dados das planilhas (Alunos, Responsáveis, notas por
// disciplina). Compartilhado entre Validation.gs e ReportGeneration.gs —
// nenhuma das duas camadas deve duplicar essa leitura.
// ============================================================

/**
 * Encontra a aba de uma disciplina na planilha de turma.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} classSpreadsheet
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
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} classSpreadsheet
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
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} classSpreadsheet
 * @returns {Array<{studentId: string, name: string, row: number}>}
 */
function getClassStudentsFromResumo(classSpreadsheet) {
  const resumoSheet = classSpreadsheet.getSheetByName(SHEET_NAMES.SUMMARY);
  if (!resumoSheet) return [];

  const lastRow = resumoSheet.getLastRow();
  if (lastRow < SUMMARY_FIRST_DATA_ROW) return [];
  const values = resumoSheet
    .getRange(
      SUMMARY_FIRST_DATA_ROW,
      1,
      lastRow - SUMMARY_FIRST_DATA_ROW + 1,
      2,
    )
    .getValues();

  return values
    .map(([studentId, name], index) => ({
      studentId: String(studentId ?? "").trim(),
      name: String(name ?? "").trim(),
      row: SUMMARY_FIRST_DATA_ROW + index,
    }))
    .filter(({ studentId }) => studentId.length > 0);
}

/**
 * Verifica se um aluno pertence à turma, usando "Resumo" como lista
 * oficial de matrículas — a mesma fonte usada por `validateClassStudents`.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} classSpreadsheet
 * @param {string} studentId
 * @returns {boolean}
 */
function isStudentInClass(classSpreadsheet, studentId) {
  return getClassStudentsFromResumo(classSpreadsheet).some(
    (student) => student.studentId === studentId,
  );
}

/**
 * Lê a aba "Alunos" do Cadastro e devolve um mapa.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} registrationSheet
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
      birthDate: formatDate(row[STUDENT_COLUMNS.birthDate]),
      sex: row[STUDENT_COLUMNS.sex],
    });
  }

  return map;
}

/**
 * Verifica se há matrículas duplicadas na aba "Alunos".
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} registrationSheet
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
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} registrationSheet
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
 *
 * @param {Array<{studentId: string, row: number}>} students
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
    birthDate: formatDate(row[STUDENT_COLUMNS.birthDate]),
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

// ============================================================
// ARQUIVO: Validation.gs
// ============================================================

// ============================================================
// ARQUIVO: Validation.gs (REFATORADO PARA OPÇÃO 3)
// ============================================================

/**
 * Compara os alunos da aba "Resumo" de uma turma com o Cadastro de Alunos.
 * Retorna objetos detalhados de diagnóstico.
 */
function validateClassStudents(
  classSpreadsheet,
  registeredStudentsMap,
  year,
  className,
) {
  const issues = [];
  const students = getClassStudentsFromResumo(classSpreadsheet);
  const ssUrl = classSpreadsheet.getUrl();

  if (students.length === 0) {
    issues.push({
      type: "error",
      text: `[${year} / ${className}] Turma sem alunos cadastrados na aba "${SHEET_NAMES.SUMMARY}".`,
      url: ssUrl,
    });
    return issues;
  }

  const dupes = findDuplicateResumoIds(students, year, className);
  issues.push(
    ...dupes.map((msg) => ({ type: "error", text: msg, url: ssUrl })),
  );

  for (const { studentId, name, row } of students) {
    const registeredStudent = registeredStudentsMap.get(studentId);

    if (registeredStudent === undefined) {
      issues.push({
        type: "warning",
        text: `[${year} / ${className} / Resumo, linha ${row}] Matrícula ${studentId} não consta no Cadastro de Alunos.`,
        url: ssUrl,
      });
      continue;
    }

    const namesDiffer =
      registeredStudent.name.localeCompare(name, DEFAULT_LOCALE, {
        sensitivity: "base",
      }) !== 0;

    if (namesDiffer) {
      issues.push({
        type: "warning",
        text: `[${year} / ${className} / Resumo, linha ${row}] Nome "${name}" diverge do Cadastro ("${registeredStudent.name}") para a matrícula ${studentId}.`,
        url: ssUrl,
      });
    }
  }

  return issues;
}

/**
 * Verifica todas as configurações, estrutura de pastas e dados.
 * Disparada pelo menu do usuário.
 */
function checkConfiguration() {
  const ui = SpreadsheetApp.getUi();
  const issues = [];
  let config;

  try {
    config = loadConfig();
  } catch (e) {
    issues.push({
      type: "error",
      text: `Configuração: ${e.message}`,
      url: SpreadsheetApp.getActiveSpreadsheet().getUrl(),
    });
    showValidationDialog(issues);
    return;
  }

  try {
    getReportTemplateFile(config);
  } catch (e) {
    issues.push({ type: "error", text: e.message });
  }

  try {
    getClassTemplateFile(config);
  } catch (e) {
    issues.push({ type: "error", text: e.message });
  }

  try {
    DriveApp.getFolderById(config.pdfsFolderId);
  } catch {
    issues.push({
      type: "error",
      text: "PDFs: pasta não encontrada ou sem acesso.",
    });
  }

  let registeredStudentsMap;
  try {
    const registrationSheet = SpreadsheetApp.openById(
      config.studentsSpreadsheetId,
    );
    const regUrl = registrationSheet.getUrl();

    if (!registrationSheet.getSheetByName(SHEET_NAMES.STUDENTS)) {
      issues.push({
        type: "error",
        text: `Cadastro de Alunos: a aba "${SHEET_NAMES.STUDENTS}" não existe.`,
        url: regUrl,
      });
    }
    if (!registrationSheet.getSheetByName(SHEET_NAMES.GUARDIANS)) {
      issues.push({
        type: "error",
        text: `Cadastro de Alunos: a aba "${SHEET_NAMES.GUARDIANS}" não existe.`,
        url: regUrl,
      });
    }

    registeredStudentsMap = loadStudentsMap(registrationSheet);

    const dupes = findDuplicateStudentIds(registrationSheet);
    issues.push(
      ...dupes.map((msg) => ({ type: "error", text: msg, url: regUrl })),
    );
  } catch (e) {
    issues.push({ type: "error", text: `Cadastro de Alunos: ${e.message}` });
  }

  let years = [];
  try {
    years = listSchoolYears(config);
  } catch (e) {
    issues.push({ type: "error", text: e.message });
  }

  if (years.length === 0) {
    try {
      const rootFolder = DriveApp.getFolderById(config.schoolYearsFolderId);
      issues.push({
        type: "error",
        text: 'Nenhuma pasta de ano letivo encontrada dentro de "Anos Letivos".',
        url: rootFolder.getUrl(),
      });
    } catch {
      issues.push({
        type: "error",
        text: 'Nenhuma pasta de ano letivo encontrada dentro de "Anos Letivos".',
      });
    }
  }

  for (const year of years) {
    let yearFolder;
    try {
      yearFolder = getSchoolYearFolder(config, year);
    } catch (e) {
      issues.push({ type: "error", text: e.message });
      continue;
    }

    for (const { className } of VALID_CLASSES) {
      let classFile;
      try {
        classFile = getClassSpreadsheetFile(yearFolder, year, className);
      } catch (e) {
        issues.push({
          type: "error",
          text: `[${year}] ${e.message}`,
          url: yearFolder.getUrl(),
        });
        continue;
      }

      let classSpreadsheet;
      try {
        classSpreadsheet = SpreadsheetApp.openById(classFile.getId());
      } catch (e) {
        issues.push({
          type: "error",
          text: `[${year} / ${className}] Erro ao abrir a planilha: ${e.message}`,
          url: classFile.getUrl(),
        });
        continue;
      }

      const ssUrl = classSpreadsheet.getUrl();
      const { missing } = checkClassSubjects(classSpreadsheet);
      if (missing.length > 0) {
        issues.push({
          type: "warning",
          text: `[${year} / ${className}] Disciplinas faltando (serão ignoradas): ${missing.join(", ")}`,
          url: ssUrl,
        });
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

  showValidationDialog(issues);
}

/**
 * Renderiza o dialog HTML com os resultados.
 */
function showValidationDialog(issues) {
  const ui = SpreadsheetApp.getUi();
  const template = HtmlService.createTemplateFromFile("ValidationResultDialog");

  // Passamos os dados estruturados para o frontend
  template.issues = issues;

  const htmlOutput = template.evaluate().setWidth(600).setHeight(520);
  ui.showModalDialog(htmlOutput, "Diagnóstico do Sistema");
}

// ============================================================
// ARQUIVO: Actions.gs
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
 * Abre o diálogo unificado de seleção de Ano Letivo e Turma.
 * @param {"single" | "class"} actionType 'single' para aluno individual, 'class' para a turma toda.
 */
function openSelectYearClassDialog(actionType) {
  const ui = SpreadsheetApp.getUi();
  try {
    const config = loadConfig();
    const years = listSchoolYears(config);

    if (years.length === 0) {
      ui.alert('Nenhum ano letivo encontrado dentro da pasta "Anos Letivos".');
      return;
    }

    const template = HtmlService.createTemplateFromFile(
      "SelectYearClassDialog",
    );
    template.years = years;
    template.classes = VALID_CLASSES.map((c) => c.className);
    template.actionType = actionType;

    const height = actionType === "single" ? 320 : 240;

    const htmlOutput = template.evaluate().setWidth(400).setHeight(height);

    ui.showModalDialog(
      htmlOutput,
      actionType === "single"
        ? "Gerar Boletim do Aluno"
        : "Gerar Boletins da Turma",
    );
  } catch (e) {
    ui.alert(`Erro ao abrir seleção: ${e.message}`);
  }
}

function generateStudentReport() {
  openSelectYearClassDialog("single");
}

function generateClassReports() {
  openSelectYearClassDialog("class");
}

/**
 * Retorna a lista de objetos contendo ID e Nome dos alunos mapeados da aba Resumo.
 * Endpoint consumido pelo Alpine.js para alimentar o datalist de autocompletar.
 */
function getStudentsDataForClass(schoolYearLabel, className) {
  try {
    const config = loadConfig();
    const yearFolder = getSchoolYearFolder(config, schoolYearLabel);
    const classFile = getClassSpreadsheetFile(
      yearFolder,
      schoolYearLabel,
      className,
    );
    const classSpreadsheet = SpreadsheetApp.openById(classFile.getId());

    const students = getClassStudentsFromResumo(classSpreadsheet);

    // Retorna uma array de objetos [{studentId: "...", name: "..."}, ...]
    return students.map((s) => ({
      studentId: s.studentId,
      name: s.name,
    }));
  } catch (e) {
    throw new Error(e.message);
  }
}

function executeClassReportsGeneration(schoolYearLabel, className) {
  withScriptLock((ui) => {
    executeClassReportsGeneration_(ui, schoolYearLabel, className);
  }, "Já existe uma geração de boletins em andamento. Tente novamente em alguns instantes.");
}

function executeStudentReportGeneration(schoolYearLabel, className, studentId) {
  if (!studentId) throw new Error("Matrícula não pode ser vazia.");
  withScriptLock((ui) => {
    executeStudentReportGeneration_(ui, schoolYearLabel, className, studentId);
  }, "Já existe uma geração de boletim em andamento. Tente novamente em alguns instantes.");
}

function executeClassReportsGeneration_(ui, schoolYearLabel, className) {
  const config = loadConfig();
  const yearFolder = getSchoolYearFolder(config, schoolYearLabel);
  const classFile = getClassSpreadsheetFile(
    yearFolder,
    schoolYearLabel,
    className,
  );
  const classSpreadsheet = SpreadsheetApp.openById(classFile.getId());

  const { found, missing } = checkClassSubjects(classSpreadsheet);
  if (found.length === 0) {
    throw new Error("Nenhuma disciplina reconhecida nessa turma.");
  }
  const foundSubjects = found;

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
  const MAX_RUNTIME_MS = 5 * 60 * 1000;

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
    Utilities.sleep(200);
  }

  const MAX_ERRORS_SHOWN = 15;
  const errorsToShow = errors.slice(0, MAX_ERRORS_SHOWN);
  const truncatedCount = errors.length - errorsToShow.length;

  const template = HtmlService.createTemplateFromFile(
    "ClassReportResultDialog",
  );
  template.successCount = successCount;
  template.className = className;
  template.schoolYearLabel = schoolYearLabel;
  template.errors = errorsToShow;
  template.truncatedCount = truncatedCount;
  template.pdfFolderUrl = context.pdfFolder.getUrl();

  const htmlOutput = template.evaluate().setWidth(450).setHeight(460);
  ui.showModalDialog(htmlOutput, "Boletins gerados");
}

function executeStudentReportGeneration_(
  ui,
  schoolYearLabel,
  className,
  studentId,
) {
  const config = loadConfig();
  const yearFolder = getSchoolYearFolder(config, schoolYearLabel);
  const classFile = getClassSpreadsheetFile(
    yearFolder,
    schoolYearLabel,
    className,
  );
  const classSpreadsheet = SpreadsheetApp.openById(classFile.getId());

  const { found: foundSubjects } = checkClassSubjects(classSpreadsheet);
  if (foundSubjects.length === 0) {
    throw new Error("Nenhuma disciplina reconhecida nessa turma.");
  }

  if (!isStudentInClass(classSpreadsheet, studentId)) {
    throw new Error(
      `Matrícula ${studentId} não encontrada na turma "${className}" (${schoolYearLabel}).`,
    );
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

  const template = HtmlService.createTemplateFromFile("ReportSuccessDialog");
  template.studentId = studentId;
  template.className = className;
  template.schoolYearLabel = schoolYearLabel;
  template.pdfUrl = pdfUrl;
  template.pdfFolderUrl = context.pdfFolder.getUrl();

  const htmlOutput = template.evaluate().setWidth(420).setHeight(360);
  ui.showModalDialog(htmlOutput, "Boletim gerado");
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

  const schoolYearLabel = `${SCHOOL_YEAR_LABEL_PREFIX}${yearInput}`;
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

  for (const { className } of VALID_CLASSES) {
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

// ============================================================
// ARQUIVO: ReportGeneration.gs
// ============================================================

/**
 * Monta o contexto que `generateReportForStudent` precisa, carregando uma
 * única vez por turma o cadastro de alunos, responsáveis e as notas de
 * todas as disciplinas — usado quando se vai gerar para todos os alunos
 * da turma de uma vez.
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
 * @param {Object} params
 * @param {string} params.studentId
 * @param {string} params.className
 * @param {Subject[]} params.foundSubjects
 * @param {ReportContext} params.context
 * @returns {string} O URL do arquivo PDF gerado
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
  const docCopy = context.templateFile.makeCopy(fileName, context.tempFolder);
  const classInfo = VALID_CLASSES.find((c) => c.className === className);
  const date = new Date();

  try {
    const doc = DocumentApp.openById(docCopy.getId());
    const body = doc.getBody();

    replacePlaceholder(body, "nome", personalData.name);
    replacePlaceholder(body, "matricula", studentId);
    replacePlaceholder(body, "filiacao", personalData.guardianNames);
    replacePlaceholder(body, "endereco", personalData.address);

    replacePlaceholder(body, "data_nascimento", personalData.birthDate);
    replacePlaceholder(body, "nacionalidade", personalData.nationality);
    replacePlaceholder(body, "sexo", formatSex(personalData.sex));

    replacePlaceholder(body, "etapa", classInfo?.stage ?? "");
    replacePlaceholder(body, "serie", classInfo?.className ?? "");
    replacePlaceholder(body, "turma", "Única");
    replacePlaceholder(body, "turno", classInfo?.shift ?? "");

    replacePlaceholder(body, "ano_letivo", String(context.yearNumber));

    replacePlaceholder(body, "data_emissao", formatDate(date));
    replacePlaceholder(body, "hora_emissao", date.toLocaleTimeString());

    for (const subject of foundSubjects) {
      const grades = gradesData[subject.name] ?? {};
      fillSubjectPlaceholders(body, subject.code, grades);
    }

    doc.saveAndClose();

    const pdfBlob = docCopy.getAs("application/pdf");
    const pdfFile = context.pdfFolder
      .createFile(pdfBlob)
      .setName(`${fileName}.pdf`);

    trashPreviousPdfVersions(context.pdfFolder, fileName, pdfFile.getId());
    return pdfFile.getUrl();
  } finally {
    docCopy.setTrashed(true);
  }
}

/**
 * @param {GoogleAppsScript.Drive.Folder} pdfFolder
 * @param {string} fileName
 * @param {string} keepFileId
 */
function trashPreviousPdfVersions(pdfFolder, fileName, keepFileId) {
  const existingFiles = pdfFolder.getFilesByName(`${fileName}.pdf`);
  while (existingFiles.hasNext()) {
    const file = existingFiles.next();
    if (file.getId() !== keepFileId) {
      file.setTrashed(true);
    }
  }
}

/**
 * @param {GoogleAppsScript.Document.Body} body
 * @param {string} subjectCode
 * @param {Object} grades
 */
function fillSubjectPlaceholders(body, subjectCode, grades) {
  for (const { suffix, field, format } of SUBJECT_PLACEHOLDER_FIELDS) {
    replacePlaceholder(
      body,
      `${subjectCode}_${suffix}`.toLowerCase(),
      format(grades[field]),
    );
  }
}

// ============================================================
// ARQUIVO: Utils.gs
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
 *
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
 * Formata os números que representam as notas com até 2 casas decimais.
 *
 * @param {any} value
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
 * Formata um valor simples, inserindo "--" caso esteja vazio.
 *
 * @param {any} value
 * @return {string}
 */
function formatValue(value) {
  if (value === "" || value === null || value === undefined) {
    return "--";
  }

  return String(value);
}

/**
 *
 * @param {string} sex
 * @return {string}
 */
function formatSex(sex) {
  const gender = {
    F: "Feminino",
    M: "Masculino",
  };

  return gender[sex] ?? "";
}

/**
 * Retorna a data no formato extenso e de acordo com as opções nativas do objeto Date.
 *
 * @param {Date} date
 * @param {Intl.DateTimeFormatOptions} [options]
 * @return {string}
 */
function formatDate(date, options) {
  if (!date || !(date instanceof Date)) return "";
  return new Intl.DateTimeFormat(DEFAULT_LOCALE, {
    dateStyle: "short",
    timeZone: DEFAULT_TIMEZONE,
    ...options,
  }).format(date);
}
