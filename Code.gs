/**
 * AFRO RUN — Fumaça 2026
 * Google Apps Script conectado a uma planilha.
 * Configure SHEET_ID abaixo com o ID da planilha.
 */
const SHEET_ID = "COLE_AQUI_O_ID_DA_PLANILHA";
const SHEET_NAME = "Inscricoes";

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents || "{}");
    const required = ["nome","nascimento","cpf","whatsapp","comunidade","camisa"];
    required.forEach(key => {
      if (!data[key]) throw new Error("Campo obrigatório ausente: " + key);
    });
    if (data.aptidao !== true || data.imagem !== true || data.regulamento !== true) {
      throw new Error("Declarações obrigatórias não confirmadas.");
    }
    if (data.menor16 === true && (!data.responsavel || data.autorizacao_responsavel !== true)) {
      throw new Error("Menor de 16 anos precisa de responsável e autorização (Art. 6º).");
    }

    // Lock evita números duplicados em envios simultâneos.
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      const ss = SpreadsheetApp.openById(SHEET_ID);
      let sheet = ss.getSheetByName(SHEET_NAME);
      if (!sheet) {
        sheet = ss.insertSheet(SHEET_NAME);
        sheet.appendRow(["Número","Data de envio","Nome","Nascimento","CPF","WhatsApp","Cidade/Comunidade","Camisa","Sexo","Aptidão","Autorização de imagem","Aceite do regulamento","Menor de 16","Responsável","Autorização do responsável"]);
        sheet.setFrozenRows(1);
      }
      // A linha 1 é o cabeçalho; número seguinte segue o total de linhas.
      const number = Math.max(1, sheet.getLastRow());
      sheet.appendRow([
        number, new Date(), data.nome, data.nascimento, data.cpf, data.whatsapp,
        data.comunidade, data.camisa, data.sexo || "", Boolean(data.aptidao), Boolean(data.imagem), Boolean(data.regulamento),
        Boolean(data.menor16), data.responsavel || "", Boolean(data.autorizacao_responsavel)
      ]);
      return jsonResponse({success:true, number:number});
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    return jsonResponse({success:false, message:String(err.message || err)});
  }
}

function doGet() {
  return jsonResponse({success:true, message:"AFRO RUN inscrições online"});
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/*
PUBLICAÇÃO:
1. Crie uma Planilha Google; copie o ID que aparece entre /d/ e /edit na URL.
2. Abra Extensões > Apps Script e cole este código no arquivo Code.gs.
3. Substitua SHEET_ID pelo ID da planilha.
4. Clique em Implantar > Nova implantação > Tipo: Aplicativo da web.
5. Executar como: você; acesso: qualquer pessoa.
6. Autorize o script e copie a URL /exec do Web App.
7. No script.js do site, substitua SCRIPT_URL pela URL /exec.
8. Publique index.html, style.css, script.js, logo.png e favicon.png na raiz do repo.
9. Teste uma inscrição e confira a nova linha na planilha.
*/
