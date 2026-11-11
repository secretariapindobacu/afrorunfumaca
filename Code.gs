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
  "Ganha camisa?", "Sexo", "Menor de 16?", "Responsável", "Autorização do responsável",
  "Apto(a)", "Uso de imagem", "Aceitou regulamento"];

function aba_() {
  const ss = SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(ABA);
  if (!sh) {
    sh = ss.insertSheet(ABA);
    sh.appendRow(CABECALHO);
    sh.getRange(1, 1, 1, CABECALHO.length).setFontWeight("bold").setBackground("#2A140B").setFontColor("#FFC400");
    sh.setFrozenRows(1);
  } else if (sh.getMaxColumns() >= 8 && String(sh.getRange(1, 8).getValue()).indexOf("tamanho") >= 0) {
    sh.deleteColumn(8); // não há mais tamanho de camisa: remove a coluna antiga
  }
  return sh;
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    const d = JSON.parse((e && e.postData && e.postData.contents) || "{}");
    ["nome", "nascimento", "cpf", "whatsapp", "comunidade"].forEach(function (k) {
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
      ganha, d.sexo || "", d.menor16 ? "SIM" : "NÃO", d.responsavel || "",
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
 * PAINEL: totais e gráficos (pizza de camisas, sexo, faixa etária e menores,
 * barras por comunidade e inscrições por dia). Atualiza sozinho.
 * Para recriar do zero: no editor, escolha "montarPainel" e clique em Executar.
 * ===================================================================== */
function montarPainel() {
  const ss = SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
  const SEP = /^en/.test(ss.getSpreadsheetLocale()) ? "," : ";";
  const F_ = function (f) { let r = "", q = false; for (const c of f) { if (c === '"') q = !q; r += (!q && c === ",") ? SEP : c; } return r; };
  const ins = aba_();
  const COFFEE = "#2A140B", SUN = "#FFC400", ORANGE = "#FF5A1F", TURQ = "#00B3A7", PINK = "#F0368A", GREEN = "#2BB24C", CREAM = "#FFF4E2", PURPLE = "#7B3FC4";
  const I = "'" + ABA + "'!";

  // ---- aba Inscrições organizada
  ins.setFrozenRows(1);
  ins.setRowHeight(1, 34);
  ins.getRange(1, 1, 1, CABECALHO.length).setFontWeight("bold").setBackground(COFFEE).setFontColor(SUN).setVerticalAlignment("middle").setWrap(true);
  [45, 135, 230, 95, 115, 120, 150, 85, 95, 80, 200, 110, 70, 80, 95].forEach(function (w, i) { ins.setColumnWidth(i + 1, w); });
  ins.getRange("A:A").setHorizontalAlignment("center");
  ins.getRange("H:J").setHorizontalAlignment("center");
  ins.getRange("B2:B").setNumberFormat("dd/mm/yyyy HH:mm");
  ins.getRange("D2:D").setNumberFormat("dd/mm/yyyy");
  if (ins.getBandings().length === 0) {
    ins.getRange(1, 1, ins.getMaxRows(), CABECALHO.length).applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false)
      .setHeaderRowColor(COFFEE).setFirstRowColor("#FFFFFF").setSecondRowColor(CREAM);
  }
  ins.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("SIM").setBackground("#D9F2E1").setFontColor("#156B33").setRanges([ins.getRange("H2:H")]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("NÃO").setBackground("#FDE2E2").setFontColor("#9B1C1C").setRanges([ins.getRange("H2:H")]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("SIM").setBackground("#FFF1C2").setFontColor("#7A5200").setRanges([ins.getRange("J2:J")]).build()
  ]);
  const p1 = ss.getSheetByName("Página1");
  if (p1 && p1.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(p1);

  // ---- Painel
  let pn = ss.getSheetByName("Painel");
  if (pn) ss.deleteSheet(pn);
  pn = ss.insertSheet("Painel", 0);
  pn.setHiddenGridlines(true);
  pn.setTabColor(ORANGE);
  ins.setTabColor(TURQ);
  const W = [16, 140, 125, 125, 20, 140, 125, 125, 16];   // A..I (A, E e I são margens)
  W.forEach(function (w, i) { pn.setColumnWidth(i + 1, w); });
  pn.getRange("A1:I200").setFontFamily("Arial").setVerticalAlignment("middle").setBackground("#FFFBF3");

  const box = function (a1, txt, bg, fg, size) {
    return pn.getRange(a1).merge().setValue(txt).setBackground(bg).setFontColor(fg || "#FFFFFF").setFontWeight("bold").setFontSize(size || 10).setHorizontalAlignment("center");
  };
  const num = function (a1, f, color) {
    return pn.getRange(a1).merge().setFormula(F_(f)).setBackground("#FFFFFF").setFontColor(color || COFFEE).setFontWeight("bold").setFontSize(28).setHorizontalAlignment("center")
      .setBorder(true, true, true, true, false, false, "#EADBC2", SpreadsheetApp.BorderStyle.SOLID);
  };

  // título
  pn.setRowHeight(1, 12);
  box("B2:H2", "AFRO RUN FUMAÇA 2026  ·  PAINEL DE INSCRIÇÕES", COFFEE, SUN, 17);
  pn.setRowHeight(2, 50);
  pn.getRange("B3:H3").merge().setFormula(F_('="Cultura, Esporte e Igualdade em Movimento  ·  atualizado em " & TEXT(NOW(),"dd/mm/yyyy HH:mm")'))
    .setFontColor("#8A6A55").setFontSize(9).setHorizontalAlignment("center").setFontStyle("italic");
  pn.getRange("B4:H4").setBackground(SUN); pn.setRowHeight(4, 5);

  // cartões
  pn.setRowHeight(5, 10);
  box("B6:C6", "TOTAL DE INSCRITOS", ORANGE);
  box("D6:D6", "INSCRITOS HOJE", SUN, COFFEE);
  box("F6:G6", "CAMISAS GARANTIDAS", TURQ);
  box("H6:H6", "RESTAM (de " + LIMITE_CAMISAS + ")", PINK);
  num("B7:C7", "=COUNTA(" + I + "C2:C)", ORANGE);
  num("D7:D7", "=COUNTIFS(" + I + "B2:B,\">=\"&TODAY())");
  num("F7:G7", "=COUNTIF(" + I + "H2:H,\"SIM\")", TURQ);
  num("H7:H7", "=MAX(0," + LIMITE_CAMISAS + "-COUNTIF(" + I + "H2:H,\"SIM\"))", PINK);
  pn.setRowHeight(7, 58);
  pn.setRowHeight(8, 8);
  box("B9:C9", "MENORES DE 16 ANOS", GREEN);
  box("D9:D9", "SEM CAMISA", "#8A6A55");
  box("F9:H9", "CAMISAS ENTREGUES  ·  BARRA DE PROGRESSO", PURPLE);
  num("B10:C10", "=COUNTIF(" + I + "J2:J,\"SIM\")", GREEN);
  num("D10:D10", "=COUNTIF(" + I + "H2:H,\"NÃO\")", "#8A6A55");
  const COL = SEP === "," ? "," : "\\";
  pn.getRange("F10:H10").merge().setFormula("=SPARKLINE(COUNTIF(" + I + "H2:H" + SEP + "\"SIM\")" + SEP + "{\"charttype\"" + COL + "\"bar\";\"max\"" + COL + LIMITE_CAMISAS + ";\"color1\"" + COL + "\"" + TURQ + "\"})")
    .setBackground("#FFFFFF").setBorder(true, true, true, true, false, false, "#EADBC2", SpreadsheetApp.BorderStyle.SOLID);
  pn.setRowHeight(10, 58);

  // aviso
  pn.setRowHeight(11, 10);
  pn.getRange("B12:H12").merge().setFormula(F_('=IF(COUNTA(' + I + 'C2:C)=0,"📊  Os gráficos aparecem sozinhos a partir da primeira inscrição.","")'))
    .setFontColor(PINK).setFontWeight("bold").setHorizontalAlignment("center");

  // ---- tabelas que alimentam os gráficos (abaixo dos gráficos)
  const T = 80;  // linha inicial das tabelas
  box("B" + T + ":D" + T, "CAMISA OFICIAL", SUN, COFFEE);
  pn.getRange("B" + (T + 1) + ":C" + (T + 1)).setValues([["Situação", "Inscritos"]]).setFontWeight("bold").setBackground("#F3E3C7").setHorizontalAlignment("center");
  pn.getRange("B" + (T + 2) + ":B" + (T + 3)).setValues([["Com camisa"], ["Sem camisa"]]).setFontWeight("bold");
  pn.getRange("C" + (T + 2)).setFormula(F_("=COUNTIF(" + I + "H2:H,\"SIM\")")).setHorizontalAlignment("center");
  pn.getRange("C" + (T + 3)).setFormula(F_("=COUNTIF(" + I + "H2:H,\"NÃO\")")).setHorizontalAlignment("center");

  box("B" + (T + 5) + ":D" + (T + 5), "IDADE (Art. 6º)", GREEN);
  pn.getRange("B" + (T + 6) + ":C" + (T + 6)).setValues([["Grupo", "Inscritos"]]).setFontWeight("bold").setBackground("#F3E3C7").setHorizontalAlignment("center");
  pn.getRange("B" + (T + 7) + ":B" + (T + 8)).setValues([["Menores de 16"], ["16 anos ou mais"]]).setFontWeight("bold");
  pn.getRange("C" + (T + 7)).setFormula(F_("=COUNTIF(" + I + "J2:J,\"SIM\")")).setHorizontalAlignment("center");
  pn.getRange("C" + (T + 8)).setFormula(F_("=COUNTIF(" + I + "J2:J,\"NÃO\")")).setHorizontalAlignment("center");

  box("F" + T + ":H" + T, "SEXO", PINK);
  pn.getRange("F" + (T + 1) + ":G" + (T + 1)).setValues([["Sexo", "Inscritos"]]).setFontWeight("bold").setBackground("#F3E3C7");
  pn.getRange("F" + (T + 2) + ":F" + (T + 5)).setValues([["Feminino"], ["Masculino"], ["Outro"], ["Não informado"]]);
  for (let k = 0; k < 3; k++) pn.getRange("G" + (T + 2 + k)).setFormula(F_("=COUNTIF(" + I + "I2:I,F" + (T + 2 + k) + ")"));
  pn.getRange("G" + (T + 5)).setFormula(F_("=COUNTA(" + I + "C2:C)-SUM(G" + (T + 2) + ":G" + (T + 4) + ")"));

  const A = T + 10;
  box("B" + A + ":D" + A, "FAIXA ETÁRIA (no dia 28/11)", GREEN);
  pn.getRange("B" + (A + 1) + ":C" + (A + 1)).setValues([["Faixa", "Inscritos"]]).setFontWeight("bold").setBackground("#F3E3C7");
  const faixas = [
    ["Até 15 anos", "=COUNTIFS(" + I + "D2:D,\">\"&DATE(2010,11,28))"],
    ["16 a 29 anos", "=COUNTIFS(" + I + "D2:D,\"<=\"&DATE(2010,11,28)," + I + "D2:D,\">\"&DATE(1996,11,28))"],
    ["30 a 44 anos", "=COUNTIFS(" + I + "D2:D,\"<=\"&DATE(1996,11,28)," + I + "D2:D,\">\"&DATE(1981,11,28))"],
    ["45 a 59 anos", "=COUNTIFS(" + I + "D2:D,\"<=\"&DATE(1981,11,28)," + I + "D2:D,\">\"&DATE(1966,11,28))"],
    ["60 anos ou mais", "=COUNTIFS(" + I + "D2:D,\"<=\"&DATE(1966,11,28))"]
  ];
  faixas.forEach(function (f, i) { pn.getRange(A + 2 + i, 2).setValue(f[0]); pn.getRange(A + 2 + i, 3).setFormula(F_(f[1])).setHorizontalAlignment("center"); });

  box("F" + A + ":H" + A, "INSCRITOS POR CIDADE / COMUNIDADE", TURQ);
  pn.getRange("F" + (A + 1)).setFormula(F_("=IFERROR(QUERY(" + I + "G2:G,\"select G, count(G) where G <> '' group by G order by count(G) desc label G 'Comunidade', count(G) 'Inscritos'\",0),\"\")"));

  const D = A + 40;
  box("B" + D + ":D" + D, "INSCRIÇÕES POR DIA", ORANGE);
  pn.getRange("B" + (D + 1)).setFormula(F_("=IFERROR(QUERY({ARRAYFORMULA(IF(" + I + "B2:B=\"\",,INT(" + I + "B2:B)))},\"select Col1, count(Col1) where Col1 is not null group by Col1 order by Col1 label Col1 'Dia', count(Col1) 'Inscrições'\",0),\"\")"));
  pn.getRange("B" + (D + 2) + ":B" + (D + 60)).setNumberFormat("dd/mm");

  // ---- gráficos (logo abaixo dos números, em 2 colunas)
  const graf = function (tipo, faixa, linha, col, titulo, cores, extra, w, h) {
    let b = pn.newChart().setChartType(tipo).addRange(pn.getRange(faixa)).setNumHeaders(1)
      .setPosition(linha, col, 6, 0)
      .setOption("title", titulo).setOption("titleTextStyle", { color: COFFEE, fontSize: 14, bold: true })
      .setOption("width", w || 384).setOption("height", h || 320).setOption("colors", cores)
      .setOption("backgroundColor", { fill: "#FFFFFF", stroke: "#EADBC2", strokeWidth: 1 })
      .setOption("chartArea", { left: "6%", top: "14%", width: "88%", height: "66%" })
      .setOption("fontName", "Arial");
    if (extra) Object.keys(extra).forEach(function (k) { b = b.setOption(k, extra[k]); });
    pn.insertChart(b.build());
  };
  const pizza = { pieHole: 0.5, pieSliceText: "value", pieSliceTextStyle: { color: "#FFFFFF", fontSize: 13, bold: true }, legend: { position: "bottom", alignment: "center", textStyle: { color: COFFEE, fontSize: 11 } }, sliceVisibilityThreshold: 0 };
  graf(Charts.ChartType.PIE, "B" + (T + 1) + ":C" + (T + 3), 14, 2, "👕  Camisa oficial (200 primeiros)", [TURQ, "#C9B8A3"], pizza);
  graf(Charts.ChartType.PIE, "F" + (T + 1) + ":G" + (T + 5), 14, 6, "Inscritos por sexo", [PINK, TURQ, SUN, "#C9B8A3"], pizza);
  graf(Charts.ChartType.PIE, "B" + (A + 1) + ":C" + (A + 6), 31, 2, "Faixa etária", [SUN, ORANGE, PINK, TURQ, GREEN], pizza);
  graf(Charts.ChartType.PIE, "B" + (T + 6) + ":C" + (T + 8), 31, 6, "Menores de 16 anos", [GREEN, SUN], pizza);
  graf(Charts.ChartType.BAR, "F" + (A + 1) + ":G" + (A + 30), 48, 2, "📍  Inscritos por cidade / comunidade", [TURQ],
    { legend: { position: "none" }, hAxis: { minValue: 0, format: "0", gridlines: { color: "#F1E6D3" } }, chartArea: { left: "30%", top: "14%", width: "62%", height: "76%" } }, 790, 300);
  graf(Charts.ChartType.AREA, "B" + (D + 1) + ":C" + (D + 60), 65, 2, "📈  Inscrições por dia", [ORANGE],
    { legend: { position: "none" }, vAxis: { minValue: 0, format: "0", gridlines: { color: "#F1E6D3" } }, areaOpacity: 0.25, lineWidth: 3, pointSize: 6 }, 790, 280);

  ss.setActiveSheet(pn);
}
