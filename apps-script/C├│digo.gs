const SPREADSHEET_ID = "1t42qOqcZ-d-0OthMLt5kNFtPw_XALTQ4pCo_SySe4nM";
const SHEET_NAME = "Fichas";

const HEADERS = [
  "Data de envio", "Personagem", "Idade", "Aparência", "Personalidade",
  "Intelecto", "Presença", "Agilidade", "Vigor", "Força", "Habilidades",
  "Inventário", "Motivação para ir à vila", "O que não quer perder",
  "Fardo", "Maior sacrifício", "Maior segredo"
];

function doGet() {
  return ContentService.createTextOutput("API de A Última Notificação ativa.");
}

function doPost(e) {
  const lock = LockService.getScriptLock();

  try {
    if (!e || !e.parameter || !e.parameters) {
      throw new Error("Requisição inválida.");
    }

    const p = e.parameter;

    // Honeypot: bots simples costumam preencher este campo.
    if (String(p.website || "").trim() !== "") {
      throw new Error("Requisição rejeitada.");
    }

    const required = [
      "personagem", "idade", "aparencia", "personalidade",
      "intelecto", "presenca", "agilidade", "vigor", "forca",
      "inventario", "motivacao", "objetivo", "fardo", "sacrificio", "segredo"
    ];

    for (const field of required) {
      if (String(p[field] || "").trim() === "") {
        throw new Error("Campo obrigatório ausente: " + field);
      }
    }

    // Limites de comprimento no servidor.
    const textFields = [
      "personagem", "aparencia", "personalidade", "inventario",
      "motivacao", "objetivo", "fardo", "sacrificio", "segredo"
    ];

    for (const field of textFields) {
      if (String(p[field]).length > 3000) {
        throw new Error("Campo muito longo: " + field);
      }
    }

    const idade = Number(p.idade);
    if (!Number.isInteger(idade) || idade < 1 || idade > 500) {
      throw new Error("Idade inválida.");
    }

    const attributeNames = ["intelecto", "presenca", "agilidade", "vigor", "forca"];
    const values = attributeNames.map(name => Number(p[name]));

    if (values.some(value => !Number.isInteger(value) || value < 0 || value > 3)) {
      throw new Error("Cada atributo deve ser um número inteiro entre 0 e 3.");
    }

    if (values.reduce((sum, value) => sum + value, 0) > 9) {
      throw new Error("A soma dos atributos não pode ultrapassar 9.");
    }

    // e.parameters preserva todos os valores enviados para campos repetidos.
    const receivedSkills = (e.parameters.habilidades || [])
      .map(value => String(value).trim())
      .filter(Boolean);

    const allowedSkills = [
      "Acrobacia", "Adestramento", "Artes", "Atletismo", "Atualidades",
      "Ciências", "Crime", "Diplomacia", "Enganação", "Fortitude",
      "Iniciativa", "Intimidação", "Intuição", "Investigação", "Luta",
      "Medicina", "Ocultismo", "Percepção", "Pilotagem", "Pontaria",
      "Profissão", "Reflexos", "Religião", "Sobrevivência", "Tática", "Tecnologia"
    ];

    const skills = [...new Set(receivedSkills)];

    if (skills.some(skill => !allowedSkills.includes(skill))) {
      throw new Error("Uma habilidade não é válida.");
    }

    if (skills.length > Number(p.intelecto) + 1) {
      throw new Error("Quantidade de habilidades excedida.");
    }

    lock.waitLock(10000);

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName(SHEET_NAME);

    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
    }

    if (sheet.getLastRow() === 0) {
      sheet.appendRow(HEADERS);
      sheet.setFrozenRows(1);
    }

    const row = [
      new Date(),
      p.personagem,
      idade,
      p.aparencia,
      p.personalidade,
      ...values,
      skills.join(", "),
      p.inventario,
      p.motivacao,
      p.objetivo,
      p.fardo,
      p.sacrificio,
      p.segredo
    ];

    // Impede que textos iniciados por caracteres de fórmula sejam executados.
    const safeRow = row.map(value => {
      if (typeof value === "string" && /^[=+\-@]/.test(value)) {
        return "'" + value;
      }
      return value;
    });

    sheet.appendRow(safeRow);

    return ContentService
      .createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    console.error(error);
    return ContentService
      .createTextOutput(JSON.stringify({
        success: false,
        message: "Não foi possível registrar a ficha."
      }))
      .setMimeType(ContentService.MimeType.JSON);

  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
}
