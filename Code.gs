/**
 * AFRO RUN — Fumaça 2026
 * Recebe as inscrições do site e grava na planilha.
 *
 * COMO CONFIGURAR (uma vez só):
 * 1. Crie uma Planilha Google (ex.: "Inscrições Afro Run 2026").
 * 2. Na planilha: Extensões > Apps Script. Apague o que estiver lá e cole TODO este código. Salve.
 * 3. Implantar > Nova implantação > engrenagem > "App da Web".
 *    Executar como: EU   |   Quem pode acessar: QUALQUER PESSOA   > Implantar.
 * 4. Autorize com sua conta Google (se aparecer "app não verificado": Avançado > Acessar).
 * 5. Copie a URL que termina em /exec e mande para colocar no site (SCRIPT_URL do script.js).
 *
 * Se mudar este código depois: Implantar > Gerenciar implantações > lápis > Versão: Nova > Implantar
 * (assim a URL continua a mesma).
 */

const ABA = "Inscrições";
const LIMITE_CAMISAS = 200;          // Regulamento: camisa para os 200 primeiros inscritos
const EMAIL_AVISO = "";              // Opcional: coloque um e-mail para receber aviso a cada inscrição

const CABECALHO = ["Nº", "Data/hora", "Nome", "Nascimento", "CPF", "WhatsApp", "Cidade/Comunidade",
  "Camisa (tamanho)", "Ganha camisa?", "Sexo", "Menor de 16?", "Responsável", "Autorização do responsável",
  "Apto(a)", "Uso de imagem", "Aceitou regulamento"];

function aba_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(ABA);
  if (!sh) {
    sh = ss.insertSheet(ABA);
    sh.appendRow(CABECALHO);
    sh.getRange(1, 1, 1, CABECALHO.length).setFontWeight("bold").setBackground("#2A140B").setFontColor("#FFC400");
    sh.setFrozenRows(1);
  }
  return sh;
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    const d = JSON.parse((e && e.postData && e.postData.contents) || "{}");
    ["nome", "nascimento", "cpf", "whatsapp", "comunidade", "camisa"].forEach(function (k) {
      if (!d[k]) throw new Error("Campo obrigatório ausente: " + k);
    });
    if (d.aptidao !== true || d.imagem !== true || d.regulamento !== true) {
      throw new Error("Declarações obrigatórias não confirmadas.");
    }
    if (d.menor16 === true && (!d.responsavel || d.autorizacao_responsavel !== true)) {
      throw new Error("Menor de 16 anos precisa de responsável e autorização (Art. 6º).");
    }

    lock.waitLock(30000); // evita número repetido em envios ao mesmo tempo
    const sh = aba_();
    const cpf = String(d.cpf).replace(/\D/g, "");
    const cpfs = sh.getLastRow() > 1 ? sh.getRange(2, 5, sh.getLastRow() - 1, 1).getValues() : [];
    for (let i = 0; i < cpfs.length; i++) {
      if (String(cpfs[i][0]).replace(/\D/g, "") === cpf) {
        return json_({ success: true, number: i + 1, repetida: true });
      }
    }
    const numero = sh.getLastRow(); // linha 1 é o cabeçalho
    const ganha = numero <= LIMITE_CAMISAS ? "SIM" : "NÃO";
    sh.appendRow([numero, new Date(), d.nome, d.nascimento, "'" + d.cpf, "'" + d.whatsapp, d.comunidade,
      d.camisa, ganha, d.sexo || "", d.menor16 ? "SIM" : "NÃO", d.responsavel || "",
      d.autorizacao_responsavel ? "SIM" : "", "SIM", "SIM", "SIM"]);

    if (EMAIL_AVISO) {
      MailApp.sendEmail(EMAIL_AVISO, "AFRO RUN: nova inscrição nº " + numero,
        d.nome + " — WhatsApp " + d.whatsapp + " — " + d.comunidade + " — camisa " + d.camisa + " (ganha: " + ganha + ")");
    }
    return json_({ success: true, number: numero, camisa: ganha });
  } catch (err) {
    return json_({ success: false, message: String(err.message || err) });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

function doGet() {
  const sh = aba_();
  return json_({ success: true, total: Math.max(0, sh.getLastRow() - 1) });
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
