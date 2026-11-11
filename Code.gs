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

// ID da planilha "Inscrições AFRO RUN Fumaça 2026" (conta secretariapindobacu@gmail.com).
// Se colar este código direto em Extensões > Apps Script da própria planilha, pode deixar vazio: "".
const SHEET_ID = "1aDGUQfjY8B-NyBbCb7z1VAQfZ0Qt4BcwPrcSxi1jOcY";
const ABA = "Inscrições";
const LIMITE_CAMISAS = 200;          // Regulamento: camisa para os 200 primeiros inscritos

const CABECALHO = ["Nº", "Data/hora", "Nome", "Nascimento", "CPF", "WhatsApp", "Cidade/Comunidade",
  "Camisa (tamanho)", "Ganha camisa?", "Sexo", "Menor de 16?", "Responsável", "Autorização do responsável",
  "Apto(a)", "Uso de imagem", "Aceitou regulamento"];

function aba_() {
  const ss = SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
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

/* =====================================================================
 * PAINEL: aba com totais e gráficos (camisas por tamanho, sexo,
 * comunidades e inscrições por dia). Atualiza sozinho a cada inscrição.
 * Para recriar do zero: no editor, escolha "montarPainel" e clique em Executar.
 * ===================================================================== */
function montarPainel() {
  const ssLoc = SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
  const SEP = /^en/.test(ssLoc.getSpreadsheetLocale()) ? "," : ";";
  // troca as vírgulas das fórmulas pelo separador do idioma da planilha (fora das aspas)
  const F_ = function (f) {
    let r = "", q = false;
    for (const c of f) { if (c === '"') q = !q; r += (!q && c === ",") ? SEP : c; }
    return r;
  };
  const ss = SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
  const ins = aba_();
  const COFFEE = "#2A140B", SUN = "#FFC400", ORANGE = "#FF5A1F", TURQ = "#00B3A7", PINK = "#F0368A", GREEN = "#2BB24C", CREAM = "#FFF4E2";

  // ---- aba Inscrições mais organizada
  ins.setFrozenRows(1);
  ins.setRowHeight(1, 34);
  ins.getRange(1, 1, 1, CABECALHO.length).setFontWeight("bold").setBackground(COFFEE).setFontColor(SUN).setVerticalAlignment("middle").setWrap(true);
  const larg = [45, 135, 230, 95, 115, 120, 150, 80, 85, 95, 80, 200, 110, 70, 80, 95];
  larg.forEach(function (w, i) { ins.setColumnWidth(i + 1, w); });
  ins.getRange("A:A").setHorizontalAlignment("center");
  ins.getRange("H:K").setHorizontalAlignment("center");
  ins.getRange("B2:B").setNumberFormat("dd/mm/yyyy HH:mm");
  ins.getRange("D2:D").setNumberFormat("dd/mm/yyyy");
  if (ins.getBandings().length === 0) {
    ins.getRange(1, 1, ins.getMaxRows(), CABECALHO.length).applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false)
      .setHeaderRowColor(COFFEE).setFirstRowColor("#FFFFFF").setSecondRowColor("#FFF4E2");
  }
  const regras = [
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("SIM").setBackground("#D9F2E1").setFontColor("#156B33").setRanges([ins.getRange("I2:I")]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("NÃO").setBackground("#FDE2E2").setFontColor("#9B1C1C").setRanges([ins.getRange("I2:I")]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("SIM").setBackground("#FFF1C2").setFontColor("#7A5200").setRanges([ins.getRange("K2:K")]).build()
  ];
  ins.setConditionalFormatRules(regras);

  // apaga a "Página1" vazia, se existir
  const p1 = ss.getSheetByName("Página1");
  if (p1 && p1.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(p1);

  // ---- aba Painel (recriada)
  let pn = ss.getSheetByName("Painel");
  if (pn) ss.deleteSheet(pn);
  pn = ss.insertSheet("Painel", 0);
  pn.setHiddenGridlines(true);
  pn.getRange("A:H").setFontFamily("Arial").setVerticalAlignment("middle");
  [150, 110, 165, 30, 150, 110, 30, 30].forEach(function (w, i) { pn.setColumnWidth(i + 1, w); });
  const I = "'" + ABA + "'!";

  pn.getRange("A1:F1").merge().setValue("AFRO RUN FUMAÇA 2026 · PAINEL DE INSCRIÇÕES")
    .setFontSize(16).setFontWeight("bold").setFontColor(SUN).setBackground(COFFEE).setHorizontalAlignment("center");
  pn.setRowHeight(1, 46);
  pn.getRange("A2:F2").merge().setFormula(F_('="Atualiza sozinho a cada inscrição · " & TEXT(NOW(),"dd/mm/yyyy HH:mm")'))
    .setFontColor("#7a6a5a").setFontSize(9).setHorizontalAlignment("center");

  // cartões de números
  const cards = [
    ["A4:B4", "A5:B5", "TOTAL DE INSCRITOS", "=COUNTA(" + I + "C2:C)", ORANGE],
    ["C4:C4", "C5:C5", "CAMISAS GARANTIDAS", "=COUNTIF(" + I + "I2:I,\"SIM\")", TURQ],
    ["E4:F4", "E5:F5", "CAMISAS RESTANTES (de " + LIMITE_CAMISAS + ")", "=MAX(0," + LIMITE_CAMISAS + "-COUNTIF(" + I + "I2:I,\"SIM\"))", PINK],
    ["A6:B6", "A7:B7", "MENORES DE 16 ANOS", "=COUNTIF(" + I + "K2:K,\"SIM\")", GREEN],
    ["C6:C6", "C7:C7", "INSCRITOS HOJE", "=COUNTIFS(" + I + "B2:B,\">=\"&TODAY())", SUN],
    ["E6:F6", "E7:F7", "SEM CAMISA (após 200)", "=COUNTIF(" + I + "I2:I,\"NÃO\")", "#8A6A55"]
  ];
  cards.forEach(function (c) {
    pn.getRange(c[0]).merge().setValue(c[2]).setBackground(c[4]).setFontColor(c[4] === SUN ? COFFEE : "#FFFFFF").setFontWeight("bold").setFontSize(9).setHorizontalAlignment("center");
    pn.getRange(c[1]).merge().setFormula(F_(c[3])).setBackground(CREAM).setFontColor(COFFEE).setFontWeight("bold").setFontSize(26).setHorizontalAlignment("center");
  });
  [5, 7].forEach(function (r) { pn.setRowHeight(r, 48); });

  // tabela: camisas por tamanho
  pn.getRange("A9:C9").merge().setValue("CAMISAS POR TAMANHO").setFontWeight("bold").setFontColor(COFFEE).setBackground(SUN);
  pn.getRange("A10:C10").setValues([["Tamanho", "Pedidos (todos)", "Com camisa garantida"]]).setFontWeight("bold").setBackground("#F3E3C7");
  const tam = ["PP", "P", "M", "G", "GG"];
  tam.forEach(function (t, i) {
    const r = 11 + i;
    pn.getRange(r, 1).setValue(t).setHorizontalAlignment("center").setFontWeight("bold");
    pn.getRange(r, 2).setFormula(F_("=COUNTIF(" + I + "H2:H,A" + r + ")")).setHorizontalAlignment("center");
    pn.getRange(r, 3).setFormula(F_("=COUNTIFS(" + I + "H2:H,A" + r + "," + I + "I2:I,\"SIM\")")).setHorizontalAlignment("center");
  });
  pn.getRange("A16").setValue("TOTAL").setFontWeight("bold").setHorizontalAlignment("center");
  pn.getRange("B16").setFormula(F_("=SUM(B11:B15)")).setFontWeight("bold").setHorizontalAlignment("center");
  pn.getRange("C16").setFormula(F_("=SUM(C11:C15)")).setFontWeight("bold").setHorizontalAlignment("center");
  pn.getRange("A10:C16").setBorder(true, true, true, true, true, true, "#D8C3A0", SpreadsheetApp.BorderStyle.SOLID);

  // tabela: sexo
  pn.getRange("E9:F9").merge().setValue("SEXO").setFontWeight("bold").setFontColor("#FFFFFF").setBackground(PINK);
  pn.getRange("E10:F10").setValues([["Sexo", "Inscritos"]]).setFontWeight("bold").setBackground("#F3E3C7");
  pn.getRange("E11:E14").setValues([["Feminino"], ["Masculino"], ["Outro"], ["Não informado"]]);
  pn.getRange("F11").setFormula(F_("=COUNTIF(" + I + "J2:J,E11)"));
  pn.getRange("F12").setFormula(F_("=COUNTIF(" + I + "J2:J,E12)"));
  pn.getRange("F13").setFormula(F_("=COUNTIF(" + I + "J2:J,E13)"));
  pn.getRange("F14").setFormula(F_("=COUNTA(" + I + "C2:C)-SUM(F11:F13)"));
  pn.getRange("F11:F14").setHorizontalAlignment("center");
  pn.getRange("E10:F14").setBorder(true, true, true, true, true, true, "#D8C3A0", SpreadsheetApp.BorderStyle.SOLID);

  // tabela: comunidades (automática)
  pn.getRange("A19:C19").merge().setValue("INSCRITOS POR CIDADE / COMUNIDADE").setFontWeight("bold").setFontColor("#FFFFFF").setBackground(TURQ);
  pn.getRange("A20").setFormula(F_("=IFERROR(QUERY(" + I + "G2:G,\"select G, count(G) where G <> '' group by G order by count(G) desc label G 'Cidade/Comunidade', count(G) 'Inscritos'\",0),\"Sem inscrições ainda\")"));

  // tabela: inscrições por dia (automática)
  pn.getRange("E19:F19").merge().setValue("INSCRIÇÕES POR DIA").setFontWeight("bold").setFontColor("#FFFFFF").setBackground(ORANGE);
  pn.getRange("E20").setFormula(F_("=IFERROR(QUERY({ARRAYFORMULA(IF(" + I + "B2:B=\"\",,INT(" + I + "B2:B)))},\"select Col1, count(Col1) where Col1 is not null group by Col1 order by Col1 label Col1 'Dia', count(Col1) 'Inscrições'\",0),\"Sem inscrições ainda\")"));
  pn.getRange("E21:E80").setNumberFormat("dd/mm/yyyy");

  // ---- gráficos (à direita)
  const graf = function (tipo, faixa, linha, titulo, cores, extra) {
    let b = pn.newChart().setChartType(tipo).addRange(pn.getRange(faixa)).setNumHeaders(1)
      .setPosition(linha, 8, 0, 0)
      .setOption("title", titulo).setOption("titleTextStyle", { color: COFFEE, fontSize: 15, bold: true })
      .setOption("width", 560).setOption("height", 300).setOption("colors", cores)
      .setOption("backgroundColor", "#FFFDF8").setOption("legend", { position: "top" });
    if (extra) Object.keys(extra).forEach(function (k) { b = b.setOption(k, extra[k]); });
    pn.insertChart(b.build());
  };
  graf(Charts.ChartType.COLUMN, "A10:C15", 1, "Camisas por tamanho", [ORANGE, TURQ], { vAxis: { minValue: 0, format: "0" } });
  graf(Charts.ChartType.PIE, "E10:F14", 17, "Inscritos por sexo", [PINK, TURQ, SUN, "#B9A68F"], { pieHole: 0.45 });
  graf(Charts.ChartType.BAR, "A20:B45", 33, "Inscritos por cidade / comunidade", [TURQ], { hAxis: { minValue: 0, format: "0" }, legend: { position: "none" } });
  graf(Charts.ChartType.COLUMN, "E20:F80", 49, "Inscrições por dia", [ORANGE], { vAxis: { minValue: 0, format: "0" }, legend: { position: "none" } });

  ss.setActiveSheet(pn);
}
