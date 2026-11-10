// ===== AFRO RUN FUMAÇA 2026 – recebe as inscrições e grava na planilha =====
// Como usar: na planilha do Google > Extensões > Apps Script > cole tudo isto > Salvar
// > Implantar > Nova implantação > Tipo: App da Web > Executar como: Eu
// > Quem pode acessar: Qualquer pessoa > Implantar > copie a URL que termina em /exec.

const ABA = "Inscrições";
const LIMITE_CAMISAS = 200;
const CABECALHO = ["Nº", "Data/hora", "Nome", "Nascimento", "Idade", "Sexo", "WhatsApp",
  "Camisa (tamanho)", "Cidade/Comunidade", "Responsável", "WhatsApp responsável", "Ganha camisa?", "Aceitou regulamento e imagem"];

function aba_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(ABA);
  if (!sh) {
    sh = ss.insertSheet(ABA);
    sh.appendRow(CABECALHO);
    sh.getRange(1, 1, 1, CABECALHO.length).setFontWeight("bold").setBackground("#B5451B").setFontColor("#fff");
    sh.setFrozenRows(1);
  }
  return sh;
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function normal_(s) {
  return String(s || "").normalize("NFD").replace(/[̀-ͯ]/g, "").toLowerCase().replace(/\s+/g, " ").trim();
}

// GET: devolve só o total (para o contador de camisas da página)
function doGet() {
  return json_({ total: Math.max(0, aba_().getLastRow() - 1) });
}

// POST: grava uma inscrição
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const d = JSON.parse(e.postData.contents);
    if (!d.nome || !d.nascimento || !d.whatsapp) return json_({ ok: false, erro: "Dados incompletos." });

    const sh = aba_();
    const linhas = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues() : [];

    // Evita inscrição repetida (mesmo nome + mesma data de nascimento)
    const chave = normal_(d.nome) + "|" + d.nascimento;
    for (const l of linhas) {
      const nasc = l[3] instanceof Date ? Utilities.formatDate(l[3], "America/Bahia", "yyyy-MM-dd") : String(l[3]);
      if (normal_(l[2]) + "|" + nasc === chave) {
        return json_({ ok: true, duplicada: true, numero: l[0], camisa: l[0] <= LIMITE_CAMISAS, total: linhas.length });
      }
    }

    const numero = linhas.length + 1;
    const camisa = numero <= LIMITE_CAMISAS;
    sh.appendRow([numero, new Date(), d.nome, d.nascimento, d.idade, d.sexo, d.whatsapp,
      d.camisa, d.cidade, d.responsavel || "", d.telResp || "", camisa ? "SIM" : "NÃO", (d.aceiteRegulamento && d.aceiteImagem) ? "SIM" : "NÃO"]);
    sh.getRange(sh.getLastRow(), 4).setNumberFormat("@"); // guarda a data como texto

    return json_({ ok: true, numero, camisa, total: numero });
  } catch (err) {
    return json_({ ok: false, erro: String(err) });
  } finally {
    lock.releaseLock();
  }
}
