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

// Aba "Resumo": números atualizados sozinhos a cada inscrição
function resumo_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss.getSheetByName("Resumo")) return;
  const r = ss.insertSheet("Resumo", 0);
  const I = "'" + ABA + "'!";
  const linhas = [
    ["AFRO RUN FUMAÇA 2026 — RESUMO DAS INSCRIÇÕES", ""],
    ["", ""],
    ["Total de inscritos", "=COUNTA(" + I + "C2:C)"],
    ["Camisetas garantidas", "=COUNTIF(" + I + "L2:L,\"SIM\")"],
    ["Camisetas restantes", "=MAX(0," + LIMITE_CAMISAS + "-B4)"],
    ["Menores de 16 anos", "=COUNTIF(" + I + "E2:E,\"<16\")"],
    ["", ""],
    ["TAMANHO DAS CAMISETAS (só quem ganha)", ""],
    ["Infantil", "=COUNTIFS(" + I + "H2:H,\"Infantil\"," + I + "L2:L,\"SIM\")"],
    ["P",  "=COUNTIFS(" + I + "H2:H,\"P\","  + I + "L2:L,\"SIM\")"],
    ["M",  "=COUNTIFS(" + I + "H2:H,\"M\","  + I + "L2:L,\"SIM\")"],
    ["G",  "=COUNTIFS(" + I + "H2:H,\"G\","  + I + "L2:L,\"SIM\")"],
    ["GG", "=COUNTIFS(" + I + "H2:H,\"GG\"," + I + "L2:L,\"SIM\")"],
    ["XG", "=COUNTIFS(" + I + "H2:H,\"XG\"," + I + "L2:L,\"SIM\")"],
    ["", ""],
    ["SEXO", ""],
    ["Feminino", "=COUNTIF(" + I + "F2:F,\"Feminino\")"],
    ["Masculino", "=COUNTIF(" + I + "F2:F,\"Masculino\")"],
    ["Prefiro não informar", "=COUNTIF(" + I + "F2:F,\"Prefiro não informar\")"],
    ["", ""],
    ["INSCRITOS POR CIDADE / COMUNIDADE", ""],
    ["=IFERROR(QUERY(" + I + "I2:I,\"select I, count(I) where I is not null group by I order by count(I) desc label count(I) ''\",0),\"\")", ""]
  ];
  r.getRange(1, 1, linhas.length, 2).setValues(linhas);
  r.getRange("A1").setFontSize(14).setFontWeight("bold").setFontColor("#B5451B");
  ["A3:A6","A8","A16","A21"].forEach(x => r.getRange(x).setFontWeight("bold"));
  r.getRange("B3:B6").setFontWeight("bold").setFontSize(13);
  r.getRange("A8:B8").setBackground("#FFC928"); r.getRange("A16:B16").setBackground("#FFC928"); r.getRange("A21:B21").setBackground("#FFC928");
  r.setColumnWidth(1, 300); r.setColumnWidth(2, 120);
}

// Rode esta função uma vez pelo menu (Executar) para criar as abas antes da primeira inscrição
function configurar() { aba_(); resumo_(); }

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
    resumo_();
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
