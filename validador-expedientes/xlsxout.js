/* Genera el Excel del cruce (misma estructura que la skill) a partir del reporte del motor.
   Uso: CruceExcel.build(ExcelJS, report, { revisado: "19/10/2026" }) -> Workbook de ExcelJS */
(function (root) {
  "use strict";
  const X = {};
  const FONT = "Arial";
  const C = { hdr: "FF1F3864", white: "FFFFFFFF", border: "FFBFBFBF", blue: "FF0000FF", title: "FF1F3864" };
  const STATUS_FILL = { "S\u00ed": "FFC6EFCE", "Diferencia menor": "FFFFEB9C", "No": "FFFFC7CE", "Solo en un documento": "FFE7E6E6", "No legible": "FFF8CBAD", "No comprendido": "FFF8CBAD" };
  const PRIO_FILL = { Alta: "FFFFC7CE", Media: "FFFFEB9C", Baja: "FFE7E6E6" };
  const thin = { style: "thin", color: { argb: C.border } };
  const BORDER = { top: thin, left: thin, bottom: thin, right: thin };
  const body = (o = {}) => ({ name: FONT, size: 9, ...o });
  const colL = n => { let s = ""; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
  const toDate = iso => { if (!iso) return null; if (iso instanceof Date) return iso; const m = String(iso).match(/(\d{4})-(\d{2})-(\d{2})/); return m ? new Date(Date.UTC(+m[1], +m[2] - 1, +m[3], 12)) : null; };
  const MONEY = '#,##0.00;(#,##0.00);"-"';

  function title(ws, text, sub) {
    const a = ws.getCell("A1"); a.value = text; a.font = { name: FONT, size: 14, bold: true, color: { argb: C.title } };
    if (sub) { const b = ws.getCell("A2"); b.value = sub; b.font = body({ italic: true }); }
  }
  function header(ws, row, headers) {
    headers.forEach((h, j) => { const c = ws.getCell(row, j + 1); c.value = h; c.font = { name: FONT, size: 10, bold: true, color: { argb: C.white } }; c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: C.hdr } }; c.alignment = { wrapText: true, vertical: "middle" }; c.border = BORDER; });
  }
  function table(ws, start, headers, rows, widths, opt = {}) {
    header(ws, start, headers);
    rows.forEach((r, i) => {
      r.forEach((v, j) => { const c = ws.getCell(start + 1 + i, j + 1); c.value = v == null ? "" : v; c.font = body(); c.alignment = { wrapText: true, vertical: "top" }; c.border = BORDER; });
      if (opt.statusCol) { const c = ws.getCell(start + 1 + i, opt.statusCol); const f = STATUS_FILL[c.value]; if (f) { c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: f } }; c.font = body({ bold: true }); } }
      if (opt.prioCol) { const c = ws.getCell(start + 1 + i, opt.prioCol); const f = PRIO_FILL[c.value]; if (f) { c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: f } }; c.font = body({ bold: true }); } }
    });
    widths.forEach((w, j) => ws.getColumn(j + 1).width = w);
    if (opt.freeze !== false) ws.views = [{ state: "frozen", ySplit: start }];
    if (opt.filter !== false && rows.length) ws.autoFilter = { from: { row: start, column: 1 }, to: { row: start + rows.length, column: headers.length } };
    return start + rows.length;
  }
  const setF = (cell, formula, fmt) => { cell.value = { formula }; cell.font = body(); cell.border = BORDER; if (fmt) cell.numFmt = fmt; };
  const setV = (cell, v, o = {}) => { cell.value = v; cell.font = body(o.font || {}); cell.border = BORDER; if (o.fmt) cell.numFmt = o.fmt; if (o.wrap) cell.alignment = { wrapText: true, vertical: "top" }; };

  X.build = function (ExcelJS, rep, meta = {}) {
    const E = root.CruceEngine || (typeof require !== "undefined" ? require("./engine.js") : null);
    const wb = new ExcelJS.Workbook();
    wb.creator = "Validador de expedientes"; wb.calcProperties.fullCalcOnLoad = true;
    const xd = rep.excel || {};
    const nombre = rep.expediente || "Expediente";

    /* ---------- Portada ---------- */
    let ws = wb.addWorksheet("Portada");
    title(ws, "Cruce de expediente de cr\u00e9dito \u2014 " + nombre, "Revisi\u00f3n del " + (meta.revisado || "") + (rep.datos?.firma ? " sobre documentos fechados el " + rep.datos.firma.slice(8, 10) + "/" + rep.datos.firma.slice(5, 7) + "/" + rep.datos.firma.slice(0, 4) : "") + ". Generado autom\u00e1ticamente con reglas fijas; confirma contra los originales.");
    if (rep.validacion) {
      const v = rep.validacion, col = v.nivel === "ok" ? "C6EFCE" : v.nivel === "critico" ? "FFC7CE" : "FFEB9C";
      ws.getCell(3, 1).value = "\u2714 VALIDADO"; ws.getCell(3, 1).font = body({ bold: true, size: 11 });
      const c = ws.getCell(3, 2); c.value = v.titulo + ". " + v.detalle; c.font = body({ bold: true, size: 10 }); c.alignment = { wrapText: true, vertical: "middle" };
      ws.mergeCells(3, 2, 3, 6); ws.getRow(3).height = 30;
      [1, 2].forEach(k => { ws.getCell(3, k).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF" + col } }; ws.getCell(3, k).border = BORDER; });
    }
    const INFO = [
      ["Qu\u00e9 es este archivo", "Cruce de fechas, montos, tasas, personas, domicilios, sucursales, cuentas e identificadores entre los documentos del expediente. Cada dato indica documento y ubicaci\u00f3n exacta."],
      ["Conclusi\u00f3n", rep.conclusion || ""],
      ["C\u00f3mo usarlo", "1) 'Discrepancias y Verificaciones' = lo que hay que corregir o aclarar (las verificaciones son f\u00f3rmulas). 2) 'Dudas y No Confirmados' = preguntas concretas. 3) 'Resumen Cruzado' = todos los datos. 'Verificaci\u00f3n TA' recalcula intereses y totales con f\u00f3rmulas."],
      ["Leyenda de estatus", "S\u00ed = coincide en \u22652 documentos \u00b7 Diferencia menor = solo formato / explicable \u00b7 No = discrepancia real \u00b7 Solo en un documento = sin contraparte \u00b7 No legible / No comprendido = ver Dudas."],
      ["L\u00edmites", "El KYC y las hojas escaneadas se leen con OCR (puede confundir 0/O, 1/l, 5/S). La herramienta no ve firmas ni casillas marcadas: rev\u00edsalas en el original."],
    ];
    let r = 4;
    INFO.forEach(([k, v]) => { ws.getCell(r, 1).value = k; ws.getCell(r, 1).font = body({ bold: true }); const c = ws.getCell(r, 2); c.value = v; c.font = body(); c.alignment = { wrapText: true, vertical: "top" }; ws.mergeCells(r, 2, r, 6); ws.getRow(r).height = 42; r++; });
    r++; ws.getCell(r, 1).value = "Colores"; ws.getCell(r, 1).font = body({ bold: true });
    Object.entries(STATUS_FILL).forEach(([k, f], i) => { const c = ws.getCell(r, 2 + i); c.value = k; c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: f } }; c.font = body({ bold: true }); c.border = BORDER; });
    r += 2; ws.getCell(r, 1).value = "Inventario de documentos"; ws.getCell(r, 1).font = { name: FONT, size: 11, bold: true, color: { argb: C.title } };
    const inv = (rep.documentos || []).map(d => [d.nombre, d.tipo, d.paginas, d.calidad]);
    inv.push(["Documentos esperados NO recibidos", "\u2014", "\u2014", (rep.faltantes || []).join("; ")]);
    table(ws, r + 1, ["Archivo", "Documento identificado", "P\u00e1ginas / hojas", "Calidad de lectura / observaciones"], inv, [40, 34, 14, 70], { freeze: false, filter: false });
    ws.getColumn(5).width = 16; ws.getColumn(6).width = 16;

    /* ---------- Resumen y categor\u00edas ---------- */
    const hdrR = ["Categor\u00eda", "Campo", "Valor (como aparece)", "Documento fuente", "Apartado / ubicaci\u00f3n exacta", "Coincide entre documentos", "Observaciones"];
    const wR = [16, 30, 48, 26, 40, 16, 60];
    const rowsR = (rep.resumen || []).map(x => [x.cat, x.campo, x.valor, x.doc, x.ubic, x.st, x.obs]);
    ws = wb.addWorksheet("Resumen Cruzado"); title(ws, "Resumen cruzado \u2014 un rengl\u00f3n por dato y por documento"); table(ws, 3, hdrR, rowsR, wR, { statusCol: 6 });
    [["Fechas", "Fechas"], ["Montos y Tasas", "Montos y Tasas"], ["Personas", "Personas y Firmantes"], ["Domicilios", "Domicilios"], ["Sucursales", "Sucursales"]].forEach(([cat, name]) => {
      const w2 = wb.addWorksheet(name); title(w2, name + " \u2014 vista por categor\u00eda"); table(w2, 3, hdrR, rowsR.filter(x => x[0] === cat), wR, { statusCol: 6 });
    });

    /* ---------- Cuentas bancarias (f\u00f3rmulas) ---------- */
    ws = wb.addWorksheet("Cuentas Bancarias");
    title(ws, "Cuentas bancarias \u2014 validaci\u00f3n de CLABE con f\u00f3rmulas", "Pesos 3-7-1; verificador = (10 \u2212 suma mod 10) mod 10. Clave banco = 3 primeros d\u00edgitos; plaza = 4-6; cuenta embebida = 7-17.");
    const hC = ["Titular / uso", "Banco seg\u00fan documento", "Documento y ubicaci\u00f3n", "CLABE (texto)", "Longitud", "Verificador calculado", "D\u00edgito 18", "\u00bfVerificador OK?", "Clave banco", "Banco seg\u00fan documento (clave)", "\u00bfBanco coincide?", "Plaza", "Cuenta embebida (7-17)", "N\u00famero de cuenta en documento", "\u00bfCuenta contenida en CLABE?", "Observaciones"];
    header(ws, 4, hC);
    const accts = rep.cuentas || [];
    const nA = accts.length; const refStart = 5 + nA + 3; // tabla de claves
    const used = {};
    accts.forEach((a, k) => {
      const i = 5 + k; const D = "D" + i;
      const cl = E ? E.bankFromName(a.banco) : null; if (cl) used[cl] = E.BANKS[cl];
      [a.uso, a.banco || "", a.ubic].forEach((v, j) => setV(ws.getCell(i, j + 1), v, { wrap: true }));
      const cc = ws.getCell(i, 4); cc.value = String(a.clabe); cc.numFmt = "@"; cc.font = body(); cc.border = BORDER;
      setF(ws.getCell(i, 5), `LEN(${D})`);
      setF(ws.getCell(i, 6), `IF(LEN(${D})<17,"",MOD(10-MOD(SUMPRODUCT(MOD(VALUE(MID(${D},{1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17},1))*{3,7,1,3,7,1,3,7,1,3,7,1,3,7,1,3,7},10)),10),10))`);
      setF(ws.getCell(i, 7), `IF(LEN(${D})<18,"",VALUE(RIGHT(${D},1)))`);
      setF(ws.getCell(i, 8), `IF(AND(E${i}=18,F${i}=G${i}),"OK","ERROR")`);
      setF(ws.getCell(i, 9), `LEFT(${D},3)`);
      setV(ws.getCell(i, 10), cl ? E.BANKS[cl] : "No identificado");
      setF(ws.getCell(i, 11), `IFERROR(IF(I${i}=INDEX($B$${refStart + 1}:$B$${refStart + 40},MATCH(J${i},$A$${refStart + 1}:$A$${refStart + 40},0)),"S\u00ed","No"),"No verificado")`);
      setF(ws.getCell(i, 12), `MID(${D},4,3)`);
      setF(ws.getCell(i, 13), `MID(${D},7,11)`);
      const cn = ws.getCell(i, 14); cn.value = String(a.cuenta || ""); cn.numFmt = "@"; cn.font = body(); cn.border = BORDER;
      setF(ws.getCell(i, 15), `IF(N${i}="","\u2014",IF(ISNUMBER(SEARCH(SUBSTITUTE(N${i},"-",""),M${i})),"S\u00ed",IF(ISNUMBER(SEARCH(M${i},SUBSTITUTE(N${i},"-",""))),"S\u00ed","No")))`);
      setV(ws.getCell(i, 16), a.obs, { wrap: true });
    });
    accts.forEach(a => { const k = String(a.clabe).slice(0, 3); if (E && E.BANKS[k]) used[k] = E.BANKS[k]; });
    ws.getCell(refStart - 1, 1).value = "Claves de banco de referencia"; ws.getCell(refStart - 1, 1).font = body({ bold: true });
    ws.getCell(refStart, 1).value = "Banco"; ws.getCell(refStart, 2).value = "Clave"; ws.getCell(refStart, 1).font = ws.getCell(refStart, 2).font = body({ bold: true });
    Object.entries(used).forEach(([k, n], j) => { ws.getCell(refStart + 1 + j, 1).value = n; ws.getCell(refStart + 1 + j, 1).font = body(); const c = ws.getCell(refStart + 1 + j, 2); c.value = k; c.numFmt = "@"; c.font = body(); });
    [30, 26, 36, 22, 9, 12, 9, 11, 9, 18, 11, 8, 16, 18, 14, 50].forEach((w, j) => ws.getColumn(j + 1).width = w);

    /* ---------- Verificaci\u00f3n TA (f\u00f3rmulas) ---------- */
    const rows = xd.calendario || [];
    let VT = null, TA_TOT = null;
    if (rows.length && xd.monto) {
      ws = wb.addWorksheet("Verificaci\u00f3n TA"); VT = "'Verificaci\u00f3n TA'";
      title(ws, `Verificaci\u00f3n aritm\u00e9tica del calendario (fuente: ${xd.calSrc || "pagar\u00e9"})`, "Valores capturados del documento en azul. Columnas en negro recalculan con f\u00f3rmulas.");
      const blue = body({ color: { argb: C.blue } });
      [["A4", "Monto del cr\u00e9dito"], ["D4", "Tasa ordinaria"], ["G4", "Fecha de firma"], ["J4", "Base d\u00edas"]].forEach(([a, t]) => { ws.getCell(a).value = t; ws.getCell(a).font = body({ bold: true }); });
      ws.getCell("B4").value = xd.monto; ws.getCell("B4").numFmt = "$#,##0.00"; ws.getCell("B4").font = blue;
      ws.getCell("E4").value = (xd.tasa || 0) / 100; ws.getCell("E4").numFmt = "0.00%"; ws.getCell("E4").font = blue;
      ws.getCell("H4").value = toDate(xd.firma || rows[0].fecha); ws.getCell("H4").numFmt = "dd/mm/yyyy"; ws.getCell("H4").font = blue;
      ws.getCell("K4").value = 360; ws.getCell("K4").font = blue;
      const HR = 6;
      header(ws, HR, ["#", "Fecha de pago", "Capital", "Intereses", "IVA int.", "Comisi\u00f3n SDGC", "IVA comisi\u00f3n", "Pago (documento)", "Saldo (documento)", "D\u00edas (calc.)", "Inter\u00e9s recalculado", "Dif. inter\u00e9s", "Saldo recalculado", "Dif. saldo", "Pago recalculado", "Dif. pago", "Inter\u00e9s cotizador", "Dif. vs cotizador"]);
      const cot = xd.cotInt || [];
      rows.forEach((rw, k) => {
        const i = HR + 1 + k;
        const vals = { 1: rw.n, 2: toDate(rw.fecha), 3: rw.cap, 4: rw.int, 5: rw.iva || 0, 6: rw.com, 7: rw.ivacom, 8: rw.pago, 9: rw.saldo };
        if (cot.length === rows.length) vals[17] = cot[k];
        Object.entries(vals).forEach(([c, v]) => { const cell = ws.getCell(i, +c); cell.value = v; cell.font = blue; cell.border = BORDER; cell.numFmt = +c === 2 ? "dd/mm/yyyy" : +c === 1 ? "0" : MONEY; });
        const pd = k === 0 ? "$H$4" : `B${i - 1}`, pb = k === 0 ? "$B$4" : `I${i - 1}`, pbc = k === 0 ? "$B$4" : `M${i - 1}`;
        setF(ws.getCell(i, 10), `B${i}-${pd}`, "0");
        setF(ws.getCell(i, 11), `ROUND(${pb}*$E$4/$K$4*J${i},2)`, MONEY);
        setF(ws.getCell(i, 12), `D${i}-K${i}`, MONEY);
        setF(ws.getCell(i, 13), `${pbc}-C${i}`, MONEY);
        setF(ws.getCell(i, 14), `I${i}-M${i}`, MONEY);
        setF(ws.getCell(i, 15), `SUM(C${i}:G${i})`, MONEY);
        setF(ws.getCell(i, 16), `H${i}-O${i}`, MONEY);
        if (vals[17] != null) setF(ws.getCell(i, 18), `D${i}-Q${i}`, MONEY);
      });
      const last = HR + rows.length; TA_TOT = last + 1;
      ws.getCell(TA_TOT, 1).value = "Totales"; ws.getCell(TA_TOT, 1).font = body({ bold: true });
      [3, 4, 5, 6, 7, 8, 11, 12, 16, 17, 18].forEach(c => { if (c >= 17 && cot.length !== rows.length) return; const L = colL(c); setF(ws.getCell(TA_TOT, c), `SUM(${L}${HR + 1}:${L}${last})`, MONEY); ws.getCell(TA_TOT, c).font = body({ bold: true }); });
      [5, 12, 13, 12, 8, 12, 11, 13, 14, 7, 13, 10, 14, 10, 13, 9, 13, 10].forEach((w, j) => ws.getColumn(j + 1).width = w);
      ws.views = [{ state: "frozen", xSplit: 2, ySplit: HR }];
    }

    /* ---------- Ventas del cotizador (f\u00f3rmulas) ---------- */
    const ventas = rep.ventas || [];
    let VC = null, V = {};
    if (ventas.length) {
      ws = wb.addWorksheet("Ventas Cotizador"); VC = "'Ventas Cotizador'";
      const pago = rep.ventasInfo?.pago || rep.datos?.pago || 0, ret = (xd.ret ?? rep.ventasInfo?.ret ?? 0) / 100;
      title(ws, "Ventas con tarjeta usadas en el cotizador", `Fuente: Cotizador, hoja Inputs. Retenci\u00f3n ${Math.round(ret * 100)}% vs mensualidad.`);
      header(ws, 4, ["Mes", "Ventas con tarjeta", "Retenci\u00f3n", "Retenci\u00f3n \u2212 mensualidad"]);
      const blue = body({ color: { argb: C.blue } });
      [["G4", "Mensualidad", pago, "$#,##0.00"], ["G5", "% retenci\u00f3n", ret, "0%"], ["G6", "Meses de ventas (cotizador)", xd.mult ?? 2, "0.000000"]].forEach(([a, t, v, f]) => { ws.getCell(a).value = t; ws.getCell(a).font = body({ bold: true }); const c = ws.getCell("H" + a.slice(1)); c.value = v; c.numFmt = f; c.font = blue; });
      ventas.forEach((v, k) => { const i = 5 + k; const a = ws.getCell(i, 1); a.value = toDate(v.mes); a.numFmt = "mmm-yyyy"; a.font = body(); a.border = BORDER; const b = ws.getCell(i, 2); b.value = v.v; b.numFmt = "$#,##0.00"; b.font = blue; b.border = BORDER; setF(ws.getCell(i, 3), `B${i}*$H$5`, "$#,##0.00"); setF(ws.getCell(i, 4), `C${i}-$H$4`, '#,##0.00;[Red](#,##0.00);"-"'); });
      const e = 4 + ventas.length, b0 = e + 2;
      V = { suma: `${VC}!B${b0}`, prom: `${VC}!B${b0 + 1}`, monto: `${VC}!B${b0 + 2}`, low: `${VC}!B${b0 + 4}` };
      [[b0, "Suma de los meses", `SUM(B5:B${e})`, "$#,##0.00"], [b0 + 1, "Promedio (base del monto)", `AVERAGE(B5:B${e})`, "$#,##0.00"], [b0 + 2, "Monto = meses de ventas \u00d7 promedio", `ROUND($H$6*B${b0 + 1},0)`, "$#,##0.00"], [b0 + 3, "Promedio \u00faltimos 6 meses", `AVERAGE(B${Math.max(5, e - 5)}:B${e})`, "$#,##0.00"], [b0 + 4, "Meses en que la retenci\u00f3n no cubre la mensualidad", `COUNTIF(D5:D${e},"<0")`, "0"]].forEach(([i, l, f, fmt]) => { ws.getCell(i, 1).value = l; ws.getCell(i, 1).font = body({ bold: true }); setF(ws.getCell(i, 2), f, fmt); ws.getCell(i, 2).font = body({ bold: true }); });
      [44, 18, 16, 22, 3, 3, 30, 14].forEach((w, j) => ws.getColumn(j + 1).width = w);
    }

    /* ---------- Discrepancias y verificaciones ---------- */
    ws = wb.addWorksheet("Discrepancias y Verificaciones");
    title(ws, "Discrepancias encontradas (ordenadas por importancia)");
    const disc = (rep.discrepancias || []).map((d, i) => [i + 1, d.prio, d.campo, d.a, d.b, d.why]);
    if (!disc.length) disc.push(["\u2014", "", "Sin discrepancias detectadas por las reglas", "", "", "Revisa de todos modos las dudas y los originales."]);
    const e1 = table(ws, 3, ["#", "Prioridad", "Campo", "Qu\u00e9 dice un documento (ubicaci\u00f3n)", "Qu\u00e9 dice el otro (ubicaci\u00f3n)", "Por qu\u00e9 importa"], disc, [6, 9, 26, 55, 55, 50], { prioCol: 2, filter: false });
    const v0 = e1 + 3;
    ws.getCell(v0 - 1, 1).value = "Verificaciones aritm\u00e9ticas (con f\u00f3rmulas)"; ws.getCell(v0 - 1, 1).font = { name: FONT, size: 11, bold: true, color: { argb: C.title } };
    header(ws, v0, ["#", "Verificaci\u00f3n", "Valor A", "Valor B", "Diferencia", "Resultado", "Notas"]);
    const T = c => VT && TA_TOT ? `${VT}!${c}${TA_TOT}` : null;
    const checks = [];
    const add = (desc, a, b, note) => { if (a == null || b == null) return; checks.push([desc, a, b, note || ""]); };
    const fmt$ = n => "$" + Number(n).toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    const ad = xd.anexoD, m = xd.monto;
    if (T("C")) add(`Suma de capital = monto del cr\u00e9dito (${fmt$(m)})`, { f: T("C") }, m);
    if (T("H") && xd.totalContrato) add(`Suma de pagos = monto total a pagar (${fmt$(xd.totalContrato)})`, { f: T("H") }, xd.totalContrato);
    if (xd.sumaPrincipal && T("F")) add("Suma principal del pagar\u00e9 = cr\u00e9dito + SDGC con IVA (cl. 2.5)", xd.sumaPrincipal, { f: `${m}+${T("F")}+${T("G")}` });
    if (xd.montoPagareConv != null && xd.totalContrato != null) add("'Monto del Pagar\u00e9' del convenio = monto total a pagar (suma de los pagos)", xd.montoPagareConv, xd.totalContrato, "Regla acordada");
    if (ad && m) { add(`Comisi\u00f3n por investigaci\u00f3n = ${ad.pInv}% \u00d7 monto`, ad.inv, { f: `ROUND(${m}*${ad.pInv}%,2)` }); add(`Comisi\u00f3n por disposici\u00f3n = ${ad.pDisp}% \u00d7 monto`, ad.disp, { f: `ROUND(${m}*${ad.pDisp}%,2)` }); }
    if (ad?.sdgc && T("G")) { add(`IVA SDGC del calendario vs 16% \u00d7 ${fmt$(ad.sdgc)}`, { f: T("G") }, { f: `ROUND(${ad.sdgc}*16%,2)` }, "Centavos = redondeo mensual"); add("SDGC s/IVA del calendario = Anexo D", { f: T("F") }, ad.sdgc); }
    if (T("D")) add("Intereses del calendario vs rec\u00e1lculo (tasa/360 \u00d7 d\u00edas)", { f: T("D") }, { f: T("K") }, "Fila por fila en 'Verificaci\u00f3n TA'");
    if (VC && m) add(`Monto = ${xd.mult ?? 2} meses de ventas \u00d7 promedio`, m, { f: V.monto }, xd.mult != null && Math.abs(xd.mult - 2) > 1e-9 ? "Factor distinto de 2" : "");
    if (VC && ad?.sdgc && xd.sdgcRate && xd.base) add(`SDGC s/IVA = ${(xd.sdgcRate * 100).toFixed(3).replace(/0+$/, "")}%/1.16 \u00d7 promedio \u00d7 ${xd.base} (cotizador)`, ad.sdgc, { f: `ROUND(${xd.sdgcRate}/1.16*${V.prom}*${xd.base},2)` }, xd.base !== rep.datos?.plazo ? `Usa ${xd.base} meses; el contrato es a ${rep.datos?.plazo}` : "");
    if (xd.sinIVA?.rate?.v != null && ad?.pSDGC) { add(`% SDGC del Anexo D = tasa del bloque 'Cotizador s/IVA' (${xd.sinIVA.rate.ubic})`, ad.pSDGC, Math.round(xd.sinIVA.rate.v * 10000) / 100, "Regla acordada: el Anexo D se llena con el bloque sin IVA del cotizador"); if (ad.sdgc && xd.sinIVA.monto?.v != null) add(`Monto SDGC del Anexo D = 'Comisi\u00f3n SGDC Monto' del bloque s/IVA (${xd.sinIVA.monto.ubic})`, ad.sdgc, Math.round(xd.sinIVA.monto.v * 100) / 100, "Regla acordada"); }
    else if (VC && ad?.sdgc && ad.pSDGC && xd.base) add(`SDGC s/IVA = ${ad.pSDGC}% (escrito en Anexo D) \u00d7 promedio \u00d7 ${xd.base}`, ad.sdgc, { f: `ROUND(${ad.pSDGC}%*${V.prom}*${xd.base},2)` }, "Porcentaje tal como est\u00e1 escrito");
    if (xd.neto && ad) add("Neto a dispersar del cotizador = monto \u2212 (investigaci\u00f3n + disposici\u00f3n) \u00d7 1.16", xd.neto, { f: `ROUND(${m}-(${ad.inv}+${ad.disp})*1.16,2)` }, "El cotizador redondea a pesos");
    if (T("D") && (xd.cotInt || []).length === rows.length && rows.length) add("Intereses del contrato vs intereses del cotizador", { f: T("D") }, { f: T("Q") }, "El cotizador redondea hacia arriba");
    if (xd.totalContrato && xd.totalCot) add("Total del contrato vs total del cotizador", xd.totalContrato, xd.totalCot);
    if (VT && TA_TOT) add("Saldo final del calendario = 0", { f: `${VT}!I${TA_TOT - 1}` }, 0);
    if (xd.tasa && xd.tasaMor) add("Tasa moratoria = 1.5 \u00d7 tasa ordinaria (cl. 2.6)", xd.tasaMor, { f: `${xd.tasa}*1.5` });
    if (VC) add("Meses en que la retenci\u00f3n no cubre la mensualidad", { f: V.low }, 0, "0 = siempre cubre");
    if (VC && xd.ventasMin) add(`Ventas de los meses del cotizador vs casilla del KYC ('${xd.ventasMin.label}')`, { f: V.suma }, xd.ventasMin.min, "Negativo = no alcanza el m\u00ednimo de la casilla");
    checks.forEach(([desc, a, b, note], k) => {
      const i = v0 + 1 + k; const val = x => (x && typeof x === "object") ? { formula: x.f } : x;
      [[1, "V" + (k + 1)], [2, desc], [3, val(a)], [4, val(b)], [5, { formula: `ROUND(C${i}-D${i},2)` }], [6, { formula: `IF(ABS(E${i})<0.005,"Cuadra",IF(ABS(E${i})<1,"Diferencia por redondeo","NO CUADRA"))` }], [7, note]].forEach(([c, v]) => { const cell = ws.getCell(i, c); cell.value = v; cell.font = body(); cell.border = BORDER; cell.alignment = { wrapText: true, vertical: "top" }; if (c >= 3 && c <= 5) cell.numFmt = '#,##0.00;(#,##0.00);0.00'; });
    });
    ws.getColumn(7).width = 40;

    /* ---------- Dudas ---------- */
    ws = wb.addWorksheet("Dudas y No Confirmados");
    title(ws, "Dudas y datos no confirmados \u2014 preguntas para resolver");
    const du = (rep.dudas || []).map((d, i) => [i + 1, d.doc, d.ubic, d.ve, d.duda, d.preg]);
    if (!du.length) du.push(["\u2014", "", "", "Sin dudas registradas por las reglas.", "", ""]);
    table(ws, 3, ["#", "Documento", "Ubicaci\u00f3n exacta", "Qu\u00e9 se ve / qu\u00e9 dice", "Qu\u00e9 no est\u00e1 claro o por qu\u00e9 no se pudo confirmar", "Pregunta concreta"], du, [5, 26, 30, 50, 55, 50]);

    /* ---------- Inventario de fechas ---------- */
    const IF = rep.inventarioFechas || [];
    if (IF.length) {
      ws = wb.addWorksheet("Inventario de fechas");
      const nB = IF.filter(f => f.estado === "No coincide con la firma").length, nA = (rep.fechasProcesador || []).length;
      title(ws, "Todas las fechas encontradas en el expediente", `${IF.length} fechas. Grupo A (contrato con el procesador): ${nA} ubicaciones, deben coincidir entre s\u00ed. Grupo B (fecha de firma): ${nB} que no coinciden. Las dem\u00e1s son informativas (calendario, vencimiento, fideicomiso, constituci\u00f3n, poderes).`);
      table(ws, 4, ["#", "Documento", "P\u00e1gina", "Apartado", "Fecha (texto)", "Fecha", "Tipo", "Estado", "Contexto"], IF.map((f, i) => [i + 1, f.doc, f.pagina || "", f.apartado, f.texto, f.fecha ? f.fecha.split("-").reverse().join("/") : "", f.tipo, f.estado, f.contexto || ""]), [5, 20, 8, 30, 28, 12, 32, 26, 60], { freeze: true, filter: true });
      for (let i = 5; i <= 4 + IF.length; i++) { const c = ws.getCell(i, 8); const v = String(c.value); const col = /No coincide|Revisar/.test(v) ? "FFC7CE" : /Coincide/.test(v) ? "C6EFCE" : "EDEDED"; c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF" + col } }; }
    }

    /* ---------- Campos variables (los que en la plantilla van resaltados) ---------- */
    const CVR = rep.camposVariables || [];
    if (CVR.length) {
      ws = wb.addWorksheet("Campos variables");
      const nOk = CVR.filter(c => c.estado === "Poblado").length;
      title(ws, "Campos variables de contrato, pagar\u00e9 y convenio", `Son los campos que en la plantilla vienen resaltados (turquesa) m\u00e1s los datos clave del Anexo de Datos Generales. ${nOk} de ${CVR.length} poblados. Se excluyen los formatos en blanco por dise\u00f1o (Anexo C del contrato y Anexos C y D del convenio).`);
      const r0 = table(ws, 4, ["#", "Documento", "Campo", "Ubicaci\u00f3n", "Valor encontrado", "Estado"], CVR.map((c, i) => [i + 1, c.doc, c.campo, c.ubic, String(c.valor), c.estado]), [5, 22, 40, 36, 60, 22], { freeze: true, filter: true });
      for (let i = 5; i <= 4 + CVR.length; i++) { const c = ws.getCell(i, 6); const col = c.value === "Poblado" ? "C6EFCE" : c.value === "Marcador sin llenar" ? "FFC7CE" : "F8CBAD"; c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF" + col } }; c.font = body({ bold: true }); }
    }

    /* ---------- Cambios vs versi\u00f3n anterior ---------- */
    const cb = rep.cambios;
    if (cb) {
      ws = wb.addWorksheet("Cambios vs versi\u00f3n anterior");
      title(ws, "Cambios frente a la versi\u00f3n anterior del expediente", `Versi\u00f3n anterior: ${cb.previa.fecha} (${cb.previa.origen}). Las versiones se comparan archivo por archivo: texto en p\u00e1ginas digitales, imagen en p\u00e1ginas escaneadas y celda por celda en Excel.`);
      let rr = 4; ws.getCell(rr, 1).value = "Archivos"; ws.getCell(rr, 1).font = { name: FONT, size: 11, bold: true, color: { argb: C.title } };
      rr = table(ws, rr + 1, ["#", "Documento", "Estado", "Archivo anterior", "Archivo nuevo"], cb.archivos.map((x, i) => [i + 1, x.doc, x.estado, x.antes, x.despues]), [5, 22, 26, 45, 45], { freeze: false, filter: false }) + 2;
      ws.getCell(rr, 1).value = "Efecto en las discrepancias"; ws.getCell(rr, 1).font = { name: FONT, size: 11, bold: true, color: { argb: C.title } };
      const ef = [...cb.disc.resueltas.map(d => ["Resuelta", d.prio, d.campo]), ...cb.disc.nuevas.map(d => ["Nueva", d.prio, d.campo]), ...cb.disc.siguen.map(d => ["Sigue", d.prio, d.campo])].map((x, i) => [i + 1, ...x]);
      if (!ef.length) ef.push(["\u2014", "\u2014", "", cb.disc.sinPrevia ? "No hay cruce de la versi\u00f3n anterior para comparar." : "Ninguna discrepancia en ninguna de las dos versiones."]);
      rr = table(ws, rr + 1, ["#", "Efecto", "Prioridad", "Discrepancia"], ef, [5, 22, 26, 45], { freeze: false, filter: false, prioCol: 3 });
      ws.getCell(rr + 1, 1).value = `Dudas: ${cb.dudas.antes ?? "\u2014"} en la versi\u00f3n anterior \u2192 ${cb.dudas.despues} ahora.`; ws.getCell(rr + 1, 1).font = body({ italic: true });
      rr += 3; ws.getCell(rr, 1).value = "Detalle de cada cambio"; ws.getCell(rr, 1).font = { name: FONT, size: 11, bold: true, color: { argb: C.title } };
      const hdrRow = rr + 1; header(ws, hdrRow, ["#", "Documento", "Ubicaci\u00f3n", "Antes", "Despu\u00e9s", "Imagen antes", "Imagen despu\u00e9s"]);
      const det = cb.detalles.length ? cb.detalles : [{ doc: "\u2014", ubic: "", antes: "Sin cambios de contenido.", despues: "" }];
      let row = hdrRow + 1;
      det.forEach((d, i) => {
        [i + 1, d.doc, d.ubic, d.antes, d.despues].forEach((v, j) => { const c = ws.getCell(row, j + 1); c.value = v == null ? "" : v; c.font = body(); c.alignment = { wrapText: true, vertical: "top" }; c.border = BORDER; });
        [6, 7].forEach(j => { ws.getCell(row, j).border = BORDER; });
        let hpx = 0;
        [[d.imgA, 5], [d.imgB, 6]].forEach(([img, col]) => { if (!img) return; try { const id = wb.addImage({ base64: img, extension: "png" }); const w = Math.min(330, d.w || 330), h = Math.round(w * (d.h || 100) / (d.w || 330)); ws.addImage(id, { tl: { col: col + 0.05, row: row - 1 + 0.1 }, ext: { width: w, height: h }, editAs: "oneCell" }); hpx = Math.max(hpx, h); } catch (e) { } });
        const tl = Math.max(String(d.antes || "").length, String(d.despues || "").length);
        ws.getRow(row).height = Math.max(hpx ? hpx * 0.75 + 8 : 0, Math.min(400, 14 * Math.ceil(tl / 55)), 18);
        row++;
      });
      [5, 22, 26, 45, 45, 47, 47].forEach((w, j) => ws.getColumn(j + 1).width = w);
    }

    const order = ["Portada", "Discrepancias y Verificaciones", "Dudas y No Confirmados", "Inventario de fechas", "Campos variables", "Cambios vs versi\u00f3n anterior", "Resumen Cruzado", "Fechas", "Montos y Tasas", "Personas y Firmantes", "Domicilios", "Sucursales", "Cuentas Bancarias", "Ventas Cotizador", "Verificaci\u00f3n TA"];
    const byName = {}; wb.worksheets.forEach(w => byName[w.name] = w);
    order.filter(n => byName[n]).forEach((n, idx) => { byName[n].orderNo = idx; });
    wb.worksheets.forEach(w => { w.views = (w.views && w.views.length ? w.views : [{}]).map(v => ({ ...v, zoomScale: 90 })); });
    return wb;
  };
  X.fileName = rep => "Cruce_" + String(rep.expediente || "expediente").replace(/\u00b7/g, "_").normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/[^A-Za-z0-9_-]+/g, "_").replace(/_+/g, "_").replace(/^_|_$/g, "").slice(0, 90) + ".xlsx";

  if (typeof module !== "undefined" && module.exports) module.exports = X; else root.CruceExcel = X;
})(typeof self !== "undefined" ? self : this);
