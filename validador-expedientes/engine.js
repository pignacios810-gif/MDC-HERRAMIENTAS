/* Motor de cruce de expedientes Anticipa \u2014 reglas fijas, sin servicios externos.
   Entrada: documentos ya le\u00eddos (p\u00e1ginas con elementos posicionados o libros de Excel).
   Salida: reporte {expediente, conclusion, documentos, faltantes, resumen, discrepancias, dudas, verificaciones, cuentas, ventas}. */
(function (root) {
  "use strict";
  const E = {};

  /* ============ utilidades ============ */
  const deacc = s => String(s ?? "").normalize("NFD").replace(/[\u0300-\u036f]/g, "");
  const norm = s => deacc(s).toLowerCase().replace(/[^a-z0-9\u00f1@]+/g, " ").replace(/\s+/g, " ").trim();
  const squash = s => norm(s).replace(/ /g, "");
  const regimen = s => norm(s).replace(/\bs\s?a\s?p?\s?i?\s?(de)?\s?c\s?v\b/g, "").replace(/\bs\s?de\s?r\s?l\s?(de)?\s?c\s?v\b/g, "").trim();
  E.norm = norm;
  const money = s => { if (s == null || s === "") return null; if (typeof s === "number") return s; const m = String(s).replace(/[^\d.,-]/g, "").replace(/,/g, ""); const n = parseFloat(m); return isFinite(n) ? n : null; };
  const pct = s => { if (s == null || s === "") return null; if (typeof s === "number") return s <= 1 ? Math.round(s * 1e6) / 1e4 : s; const n = parseFloat(String(s).replace(/[^\d.]/g, "")); return isFinite(n) ? n : null; };
  const r2 = n => Math.round((n + (n >= 0 ? 1 : -1) * Number.EPSILON * Math.max(1, Math.abs(n))) * 100) / 100;
  const fmtMoney = n => n == null ? "" : "$" + Number(n).toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  E.fmtMoney = fmtMoney;
  const MESES = { enero: 1, febrero: 2, marzo: 3, abril: 4, mayo: 5, junio: 6, julio: 7, agosto: 8, septiembre: 9, setiembre: 9, octubre: 10, noviembre: 11, diciembre: 12, ene: 1, feb: 2, mar: 3, abr: 4, may: 5, jun: 6, jul: 7, ago: 8, sep: 9, sept: 9, oct: 10, nov: 11, dic: 12 };
  function parseDate(s) {
    if (s == null || s === "") return null;
    if (s instanceof Date && !isNaN(s)) { const d = new Date(s.getTime() + 12 * 3600e3); return d.toISOString().slice(0, 10); }
    const t = deacc(String(s)).toLowerCase();
    let m = t.match(/(\d{4})-(\d{2})-(\d{2})/); if (m) return `${m[1]}-${m[2]}-${m[3]}`;
    m = t.match(/(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})/); if (m) return `${m[3]}-${m[2].padStart(2, "0")}-${m[1].padStart(2, "0")}`;
    m = t.match(/(\d{1,2})\s*(?:de\s+)?([a-z]+)\s*(?:de[l]?\s+)?(\d{4})/); if (m && MESES[m[2]]) return `${m[3]}-${String(MESES[m[2]]).padStart(2, "0")}-${m[1].padStart(2, "0")}`;
    return null;
  }
  E.parseDate = parseDate;
  const fmtDate = iso => iso ? iso.slice(8, 10) + "/" + iso.slice(5, 7) + "/" + iso.slice(0, 4) : "";
  const dayDiff = (a, b) => (Date.parse(b) - Date.parse(a)) / 864e5;
  const yearsBetween = (a, b) => dayDiff(a, b) / 365.25;

  /* n\u00famero a letra (pesos) para validar n\u00famero vs letra */
  const U = ["", "uno", "dos", "tres", "cuatro", "cinco", "seis", "siete", "ocho", "nueve", "diez", "once", "doce", "trece", "catorce", "quince", "dieciseis", "diecisiete", "dieciocho", "diecinueve", "veinte", "veintiuno", "veintidos", "veintitres", "veinticuatro", "veinticinco", "veintiseis", "veintisiete", "veintiocho", "veintinueve"];
  const D = ["", "", "", "treinta", "cuarenta", "cincuenta", "sesenta", "setenta", "ochenta", "noventa"];
  const C = ["", "ciento", "doscientos", "trescientos", "cuatrocientos", "quinientos", "seiscientos", "setecientos", "ochocientos", "novecientos"];
  function w999(n) { if (n === 0) return ""; if (n === 100) return "cien"; const c = Math.floor(n / 100), r = n % 100; let s = C[c]; if (r) { s += (s ? " " : "") + (r < 30 ? U[r] : D[Math.floor(r / 10)] + (r % 10 ? " y " + U[r % 10] : "")); } return s; }
  function numToWords(n) {
    n = Math.floor(n); if (n === 0) return "cero";
    const mill = Math.floor(n / 1e6), miles = Math.floor((n % 1e6) / 1000), rest = n % 1000; const parts = [];
    if (mill) parts.push(mill === 1 ? "un millon" : w999(mill).replace(/uno$/, "un") + " millones");
    if (miles) parts.push(miles === 1 ? "mil" : w999(miles).replace(/uno$/, "un") + " mil");
    if (rest) parts.push(w999(rest));
    return parts.join(" ");
  }
  E.numToWords = numToWords;
  function letraOk(numero, letra) {
    // compara las palabras sin importar el orden (el texto puede venir partido en varias l\u00edneas)
    if (numero == null || !letra) return null;
    const cents = (String(letra).match(/(\d{2})\s*\/\s*100/) || [])[1];
    const skip = new Set(["pesos", "peso", "m", "n", "mn", "de", "con", "moneda", "nacional", "100"]);
    const apoc = w => w === "uno" ? "un" : w === "veintiuno" ? "veintiun" : w === "millones" ? "millones" : w === "mill\u00f3n" ? "millon" : w;
    const got = norm(letra).replace(/\d{2}\s*100/g, " ").split(" ").filter(w => w && !skip.has(w) && !/^\d+$/.test(w)).map(apoc).sort().join(" ");
    const expected = numToWords(numero).split(" ").filter(Boolean).map(apoc).sort().join(" ");
    const centsOk = cents == null || +cents === Math.round((numero % 1) * 100);
    return got === expected && centsOk;
  }

  /* CLABE */
  const BANKS = { "002": "Banamex", "006": "Bancomext", "009": "Banobras", "012": "BBVA M\u00e9xico", "014": "Santander", "019": "Banjercito", "021": "HSBC", "030": "BanBaj\u00edo", "036": "Inbursa", "042": "Mifel", "044": "Scotiabank", "058": "Banregio", "059": "Invex", "060": "Bans\u00ed", "062": "Afirme", "072": "Banorte", "106": "Bank of America", "113": "Ve por M\u00e1s", "127": "Banco Azteca", "130": "Compartamos", "132": "Multiva", "133": "Actinver", "136": "Intercam", "137": "BanCoppel", "138": "ABC Capital", "140": "Consubanco", "145": "Bbase", "166": "Banco del Bienestar", "646": "STP", "722": "Mercado Pago W", "728": "Spin by OXXO" };
  const BANK_KEYS = { "002": ["banamex", "citi", "nacional de mexico"], "012": ["bbva", "bancomer"], "014": ["santander"], "021": ["hsbc"], "030": ["bajio"], "036": ["inbursa"], "042": ["mifel"], "044": ["scotia"], "058": ["banregio", "regional"], "062": ["afirme"], "072": ["banorte", "mercantil del norte"], "127": ["azteca"], "137": ["coppel"], "646": ["stp", "transferencias y pagos"], "722": ["mercado pago"] };
  function checkClabe(raw) {
    const c = String(raw || "").replace(/\D/g, "");
    const r = { clabe: c, len: c.length, dvOk: false, clave: c.slice(0, 3), banco: BANKS[c.slice(0, 3)] || "Clave no verificada", plaza: c.slice(3, 6), emb: c.slice(6, 17) };
    if (c.length === 18) { const w = [3, 7, 1]; let s = 0; for (let i = 0; i < 17; i++) s += (+c[i] * w[i % 3]) % 10; r.dvCalc = (10 - s % 10) % 10; r.dvOk = r.dvCalc === +c[17]; }
    return r;
  }
  E.checkClabe = checkClabe;
  E.BANKS = BANKS;
  E.bankFromName = name => name ? (Object.keys(BANK_KEYS).find(k => BANK_KEYS[k].some(w => norm(name).includes(w))) || null) : null;
  const bankMatches = (clave, nombre) => { const kw = BANK_KEYS[clave]; if (!kw || !nombre) return null; return kw.some(w => norm(nombre).includes(w)); };
  const contained = (cuenta, emb) => { const a = String(cuenta || "").replace(/\D/g, ""); if (!a || !emb) return null; const e = emb.replace(/^0+/, ""), b = a.replace(/^0+/, ""); return e.includes(b) || b.includes(e); };

  /* ============ modelo de p\u00e1gina ============ */
  // page = {no, items:[{s,x0,x1,y,h}], source:'pdf'|'ocr', conf?}
  function buildLines(page) {
    if (page._lines) return page._lines;
    const items = page.items.filter(i => i.s && i.s.trim()).slice().sort((a, b) => a.y - b.y || a.x0 - b.x0);
    const hs = items.map(i => i.h).filter(h => h > 0).sort((a, b) => a - b);
    const medH = hs.length ? hs[Math.floor(hs.length / 2)] : 10;
    page.medH = medH;
    const tol = page.source === "ocr" ? medH * 0.55 : Math.max(2, medH * 0.3);
    const lines = [];
    for (const it of items) {
      let ln = lines.length ? lines[lines.length - 1] : null;
      if (ln && Math.abs(it.y - ln.y) <= tol) { ln.items.push(it); ln.y = (ln.y * (ln.items.length - 1) + it.y) / ln.items.length; }
      else lines.push({ y: it.y, items: [it] });
    }
    const gapJoin = page.source === "ocr" ? medH * 0.25 : 1.2;
    for (const ln of lines) {
      ln.items.sort((a, b) => a.x0 - b.x0);
      let t = ""; let prev = null;
      for (const it of ln.items) { if (prev) t += (it.x0 - prev.x1 > gapJoin ? " " : ""); t += it.s; prev = it; }
      ln.text = t.replace(/\s+/g, " ").trim();
      // segmentos (columnas) separados por huecos grandes
      const segGap = page.source === "ocr" ? medH * 1.6 : 9;
      ln.segs = []; let seg = null;
      for (const it of ln.items) { if (!seg || it.x0 - seg.x1 > segGap) { seg = { x0: it.x0, x1: it.x1, items: [it] }; ln.segs.push(seg); } else { seg.items.push(it); seg.x1 = Math.max(seg.x1, it.x1); } }
      ln.segs.forEach(sg => sg.text = sg.items.map(i => i.s).join(" ").replace(/\s+/g, " ").trim());
    }
    page._lines = lines;
    page.text = lines.map(l => l.text).join("\n");
    page.flat = lines.map(l => l.text).join(" ");
    return lines;
  }
  E.buildLines = buildLines;
  const pageText = p => (buildLines(p), p.text);
  const pageFlat = p => (buildLines(p), p.flat);

  // valor debajo de una etiqueta: busca la l\u00ednea con la etiqueta (opcionalmente despu\u00e9s de un ancla) y toma lo que est\u00e1 debajo en la misma columna
  function below(page, labelRe, opt = {}) {
    const lines = buildLines(page);
    let start = 0;
    if (opt.after) { const i = lines.findIndex(l => opt.after.test(l.text)); if (i < 0) return null; start = i; }
    let skip = opt.nth || 0;
    for (let i = start; i < lines.length; i++) {
      const ln = lines[i];
      const si = ln.segs.findIndex(sg => labelRe.test(sg.text));
      if (si < 0) continue;
      if (skip-- > 0) continue;
      const seg = ln.segs[si];
      const tolL = page.source === "ocr" ? page.medH * 1.2 : 6;
      const right = si + 1 < ln.segs.length ? ln.segs[si + 1].x0 - (page.source === "ocr" ? tolL : 0) : Infinity;
      const maxDown = (opt.lines || 1);
      const out = [];
      for (let j = i + 1, taken = 0; j < lines.length && taken < maxDown; j++) {
        const nl = lines[j];
        if (nl.y - ln.y > page.medH * (opt.maxGap || 4.5) * (taken + 1)) break;
        const its = nl.items.filter(it => it.x0 >= seg.x0 - tolL && it.x0 < right - 1);
        if (!its.length) { if (opt.lines > 1 && taken) break; continue; }
        const txt = its.map(i => i.s).join(" ").replace(/\s+/g, " ").trim();
        if (opt.stop && opt.stop.test(txt)) break;
        out.push(txt); taken++;
      }
      return out.length ? { v: out.join(" "), conf: page.source === "ocr" ? (page.conf ?? null) : null } : null;
    }
    return null;
  }
  E.below = below;
  // valor debajo de un segmento concreto
  function valUnder(page, i, seg, opt = {}) {
    const lines = buildLines(page), ln = lines[i]; const si = ln.segs.indexOf(seg);
    const tolL = page.source === "ocr" ? page.medH * 1.2 : 6;
    const right = si + 1 < ln.segs.length ? ln.segs[si + 1].x0 - (page.source === "ocr" ? tolL : 0) : Infinity;
    const out = [];
    for (let j = i + 1, taken = 0; j < lines.length && taken < (opt.lines || 1); j++) {
      const nl = lines[j]; if (nl.y - ln.y > page.medH * (opt.maxGap || 4.5) * (taken + 1)) break;
      const its = nl.items.filter(it => it.x0 >= seg.x0 - tolL && it.x0 < right - 1);
      if (!its.length) { if (taken) break; continue; }
      const txt = its.map(t => t.s).join(" ").replace(/\s+/g, " ").trim();
      if (opt.stop && opt.stop.test(txt)) break;
      out.push(txt); taken++;
    }
    return out.length ? out.join(" ") : null;
  }
  // bloque de domicilio del Anexo de Datos Generales: etiqueta de calle + colonia/municipio/estado a la derecha + C.P. (misma l\u00ednea o siguientes)
  function addrBlock(page, anchorRe) {
    const lines = buildLines(page);
    for (let i = 0; i < lines.length; i++) {
      const ln = lines[i]; const seg = ln.segs.find(sg => anchorRe.test(sg.text)); if (!seg) continue;
      const right = re => { const sg = ln.segs.find(x => x.x0 > seg.x0 && re.test(x.text)); return sg ? valUnder(page, i, sg) : null; };
      let cp = right(/C[\u00f3o]digo Postal/i);
      if (!cp) for (let j = i + 1; j < Math.min(lines.length, i + 4); j++) { const sg = lines[j].segs.find(x => /C[\u00f3o]digo Postal/i.test(x.text)); if (sg) { cp = valUnder(page, j, sg); break; } }
      const calle = valUnder(page, i, seg);
      if (!calle) return null;
      return { calle, colonia: right(/^Colonia/i), municipio: right(/Alcald/i), estado: right(/^Estado/i), cp: (String(cp || "").match(/\b(\d{5})\b/) || [])[1] || null };
    }
    return null;
  }
  const addrText = a => a ? [a.calle, a.colonia, a.municipio, a.estado, a.cp ? "C.P. " + a.cp : null].filter(Boolean).join(", ") : "";
  // \u00bfel domicilio aparece en un texto? (nombre de la calle + n\u00famero o C.P.)
  const STOPW = new Set(["calle", "av", "ave", "avenida", "c", "col", "colonia", "sin", "numero", "num", "no", "s", "n", "sn", "local", "int", "ext", "de", "del", "la", "las", "los", "el", "y", "cp", "c p", "mz", "lt", "alt", "ofn", "piso"]);
  function addrIn(a, text) {
    // todas las palabras significativas y n\u00fameros de la calle deben aparecer, y el C.P. si existe
    if (!a || !text) return null;
    const T = " " + norm(text) + " ";
    const toks = norm(a.calle).split(" ").filter(w => w && !STOPW.has(w) && (w.length > 1 || /^\d$/.test(w)) && w !== a.cp);
    if (!toks.length) return null;
    const has = w => T.includes(" " + w + " ");
    return toks.every(has) && (!a.cp || T.includes(a.cp));
  }
  E.addrIn = addrIn;
  const findPage = (pages, re) => pages.find(p => re.test(pageFlat(p)));
  const grab = (s, re, g = 1) => { const m = String(s || "").match(re); return m ? (m[g] || "").trim() : null; };
  const MONEY = "\\$\\s?([\\d,]+\\.\\d{2})";
  const DATE_WORDS = "(\\d{1,2}\\s+(?:de\\s+)?[A-Za-z\u00c1\u00c9\u00cd\u00d3\u00da\u00e1\u00e9\u00ed\u00f3\u00fa]+\\s+(?:del?\\s+)?\\d{4})";
  const garbledRatio = t => { const L = (t.match(/[A-Za-z\u00c1\u00c9\u00cd\u00d3\u00da\u00d1\u00e1\u00e9\u00ed\u00f3\u00fa\u00f1]/g) || []).length; const G = (t.match(/[\u00a1\u00bf@{}|\\\u00a7]/g) || []).length; return L ? G / L : 1; };

  /* ============ clasificaci\u00f3n ============ */
  E.classify = function (doc) {
    if (doc.type === "excel") {
      const names = doc.wb.SheetNames.map(norm);
      if (names.includes("inputs") && names.some(n => n.startsWith("cotizador"))) return "cotizador";
      const first = doc.wb.Sheets[doc.wb.SheetNames[0]];
      const a2 = String(first?.A2?.v || ""), a1 = String(first?.A1?.v || "");
      if (/plan de pago/i.test(a2 + a1)) return "ta";
      return "excel";
    }
    const all = doc.pages.map(pageFlat).join(" ");
    const t = deacc(all).toUpperCase();
    if (/CARATULA DE CREDITO/.test(t) && /ANEXO DE DATOS GENERALES/.test(t)) return "contrato";
    if (/CONVENIO DE ADHESION/.test(t) && /FIDEICOMISO/.test(t)) return "convenio";
    if (/P\s?A\s?G\s?A\s?R\s?E/.test(t) && /SUMA PRINCIPAL/.test(t) && doc.pages.length <= 6 && !/\[NOMBRE\]/.test(t)) return "pagare";
    const head = deacc(doc.pages.slice(0, 1).map(pageFlat).join(" ")).toUpperCase();
    if (/SOFIMEX/.test(t) && /CLIENTE PERSONA F[I1]SICA/.test(head)) { doc.pf = true; return "kyc"; }
    if (/SOFIMEX/.test(t) && /PERSONA MORAL/.test(t) && /PROPIETARIO REAL|ESTRUCTURA CORPORATIVA|DENOMINACION/.test(t)) return "kyc";
    if (/CONSTANCIA DE SITUACION FISCAL/.test(t)) return "csf";
    if (/CONTRATO\s*FINSUS/.test(t) && /PROCESADOR/.test(t) && /(RETENCI|LIQUIDACI|PLAZO DEL CREDITO)/.test(t)) return "captura";
    return "otro";
  };

  /* ============ extractores ============ */
  // cada extractor devuelve {fields:{key:{v,ubic,conf?,raw?}}, x:{...extra}}
  const ocrClean = v => typeof v === "string" ? v.replace(/^[|\s:;,.]+|[|\s:;,]+$/g, "").replace(/\s*\|\s*/g, " ").trim() : v;
  function F(fields, key, v, ubic, extra = {}) { if (extra.ocr) v = ocrClean(v); if (v == null || v === "" || (typeof v === "string" && !v.trim())) return; const e = { v: typeof v === "string" ? v.trim() : v, ubic, ...extra }; if (!fields[key]) { fields[key] = e; return; } const pre = u => String(u).split(" (")[0]; const all = [fields[key], ...(fields[key].more || [])]; if (all.some(o => pre(o.ubic) === pre(ubic))) return; (fields[key].more = fields[key].more || []).push(e); }

  function amortRows(text) {
    const rows = [];
    const re = /(?:^|\n)\s*(?:0\/)?(\d{1,2})\s+(\d{2}\/\d{2}\/\d{4})\s+\$?([\d,]+\.\d{2})\s+\$?([\d,]+\.\d{2})\s+\$?([\d,]+\.\d{2})\s+\$?([\d,]+\.\d{2})\s+\$?([\d,]+\.\d{2})\s+\$?([\d,]+\.\d{2})\s+\$?([\d,]+\.\d{2})/g;
    let m; while ((m = re.exec(text))) rows.push({ n: +m[1], fecha: parseDate(m[2]), cap: money(m[3]), int: money(m[4]), iva: money(m[5]), com: money(m[6]), ivacom: money(m[7]), pago: money(m[8]), saldo: money(m[9]) });
    return rows;
  }
  const modeOf = arr => { const c = {}; arr.forEach(v => c[v] = (c[v] || 0) + 1); return +Object.entries(c).sort((a, b) => b[1] - a[1])[0]?.[0]; };

  E.extractContrato = function (doc) {
    const f = {}, x = { clabes: [], domis: [], sucursales: [] }; const P = doc.pages;
    const pg = p => "p." + p.no;
    // car\u00e1tula
    const car = findPage(P, /CAR[\u00c1A]TULA DE CR[\u00c9E]DITO/);
    if (car) {
      const t = pageFlat(car);
      const pcts = [...t.matchAll(/(\d{1,3}\.\d{1,2})\s?%/g)].map(m => +m[1]);
      F(f, "cat", pcts[0], "Car\u00e1tula (" + pg(car) + ")");
      F(f, "tasaOrd", pct(grab(t, /Tasa Ordinaria:[\s\S]*?(\d{1,3}\.\d{1,2})\s?%\s*anual/i)), "Car\u00e1tula (" + pg(car) + ")");
      F(f, "tasaMor", pct(grab(t, /Tasa Moratoria:[\s\S]*?(\d{1,3}\.\d{1,2})\s?%\s*anual/i)), "Car\u00e1tula (" + pg(car) + ")");
      const ms = [...t.matchAll(/\$\s?([\d,]+\.\d{2})/g)].map(m => money(m[1]));
      F(f, "monto", ms[0], "Car\u00e1tula (" + pg(car) + ")"); F(f, "total", ms[1], "Car\u00e1tula (" + pg(car) + ")");
      F(f, "plazo", +grab(t, /PLAZO DEL CR[\u00c9E]DITO:\s*(\d+)/i) || null, "Car\u00e1tula (" + pg(car) + ")");
      x.diaPago = grab(t, /Fecha l[\u00edi]mite de pago:\s*(.+?)(?:PLAZO|Fecha de corte|$)/i);
    }
    // proemio y cierre
    const pro = findPage(P, /CONTRATO DE APERTURA DE CR[\u00c9E]DITO SIMPLE DE FECHA/);
    if (pro) F(f, "fechaFirma", grab(pageFlat(pro), new RegExp("SIMPLE DE FECHA " + DATE_WORDS, "i")), "Proemio (" + pg(pro) + ")");
    const cie = P.find(p => /FIRMAN DE CONFORMIDAD/i.test(pageFlat(p)));
    if (cie) x.fechaCierre = { v: grab(pageFlat(cie), new RegExp("EL D[\u00cdI]A " + DATE_WORDS, "i")), ubic: "Cierre del contrato (" + pg(cie) + ")" };
    const fid = grab(P.map(pageFlat).join(" "), /identificado con el n[\u00fau]mero (F\/\d+)/i); if (fid) F(f, "fideicomiso", fid, "Antecedente I");
    const ffid = grab(P.map(pageFlat).join(" "), new RegExp("Con fecha " + DATE_WORDS + " SOFIMEX", "i")); if (ffid) F(f, "fechaFideicomiso", ffid, "Antecedente I");
    // Anexo de Datos Generales
    const dg = findPage(P, /Anexo de Datos Generales/i) && P.find(p => /Referencia 2\. Datos Generales de la Acreditada/i.test(pageFlat(p)));
    if (dg) {
      const u = r => "Anexo de Datos Generales, " + r + " (" + pg(dg) + ")";
      const b = (re, after, o) => below(dg, re, { after, ...o })?.v;
      const R2 = /Referencia 2\./, R3 = /Referencia 3\./, R5 = /Referencia 5\./, R8 = /Referencia 8\./, R9 = /Referencia 9\./;
      F(f, "razon", b(/2\.1 Denominaci/), u("ref. 2.1"));
      F(f, "fechaConst", b(/Fecha de constituci/), u("ref. 2.1"));
      F(f, "escrituraConst", b(/2\.2 N[\u00fau]mero de escritura/), u("ref. 2.2"));
      F(f, "fedatarioConst", b(/^Fedatario:/, R2), u("ref. 2.2"));
      F(f, "folio", b(/2\.3 Folio/), u("ref. 2.3"));
      x.rppConst = b(/Lugar y Fecha de inscrip/, R2);
      F(f, "domCalle", b(/Domicilio Fiscal y Operativo/), u("ref. 2"));
      F(f, "domColonia", b(/Colonia o Fraccionamiento/, /Domicilio Fiscal/), u("ref. 2"));
      F(f, "domMunicipio", b(/Alcald[\u00edi]a o Municipio/, /Domicilio Fiscal/), u("ref. 2"));
      F(f, "domCP", b(/C[\u00f3o]digo Postal/, /Domicilio Fiscal/), u("ref. 2"));
      F(f, "rfc", b(/2\.4 RFC/), u("ref. 2.4"));
      const cta = b(/2\.5 Cuenta de la Acreditada/), clabe = b(/2\.5\.1 Clabe/), banco = b(/2\.5\.2 Instituci/);
      if (clabe) x.clabes.push({ uso: "Acreditado \u2014 cuenta de dispersi\u00f3n (ref. 2.5)", banco, clabe: clabe.replace(/\D/g, ""), cuenta: cta, ubic: u("ref. 2.5") });
      F(f, "idcliente", cta, u("ref. 2.5 'Cuenta de la Acreditada'"));
      F(f, "representante", b(/3\.1 Nombre/), u("ref. 3.1"));
      F(f, "fechaPoder", b(/Fecha de la escritura/, R3), u("ref. 3.1"));
      F(f, "escrituraPoder", b(/3\.2 N[\u00fau]mero de escritura/), u("ref. 3.2"));
      F(f, "obligado", b(/5\.1 Nombre completo/), u("ref. 5.1"));
      F(f, "nacObligado", b(/Fecha de nacimiento/, R5), u("ref. 5.1"));
      F(f, "curpObligado", b(/^CURP:/, R5), u("ref. 5.1"));
      x.estadoCivil = b(/Estado Civil/, R5);
      F(f, "rfcObligado", b(/5\.1\.2 RFC/), u("ref. 5.1.2"));
      F(f, "idObligado", b(/N[\u00fau]mero de Identificaci/, R5), u("ref. 5.1 'N\u00famero de Identificaci\u00f3n'"));
      F(f, "telefono", b(/N[\u00fau]mero de tel[\u00e9e]fono/, R5), u("ref. 5"));
      F(f, "emailObligado", b(/Correo electr[\u00f3o]nico:/, R5), u("ref. 5"));
      F(f, "avalCalle", b(/Calle y n[\u00fau]mero/, R5), u("ref. 5"));
      F(f, "avalCP", b(/C[\u00f3o]digo Postal/, R5), u("ref. 5"));
      F(f, "monto", money(b(/8\.1 Monto del Cr/)), u("ref. 8.1"));
      x.montoLetra = { v: b(/Monto del cr[\u00e9e]dito en letra/, null, { lines: 3, stop: /^\d+\.\d|Referencia|Tasa/i }), ubic: u("ref. 8.1 (letra)") };
      F(f, "fechaInicio", b(/8\.2 Fecha de inicio/), u("ref. 8.2"));
      F(f, "vencimiento", b(/8\.3 Fecha de vencimiento/), u("ref. 8.3"));
      F(f, "tasaOrd", pct(b(/8\.4 Tasa de Inter[\u00e9e]s Anual Ordinaria:?$/i) || b(/^8\.4 Tasa/)), u("ref. 8.4"));
      F(f, "tasaMor", pct(b(/8\.4\.1\.? Tasa/)), u("ref. 8.4.1"));
      const clabeF = b(/Cuenta Clabe/, R9), refP = b(/^Referencia:/, R9), bancoF = b(/Instituci[\u00f3o]n Bancaria/, R9), titF = b(/^Titular/, R9);
      F(f, "referencia", refP, u("ref. 9"));
      if (clabeF) x.clabes.push({ uso: "Fideicomiso \u2014 cuenta concentradora (ref. 9)", banco: bancoF, clabe: clabeF.replace(/\D/g, ""), cuenta: "", ubic: u("ref. 9"), titular: titF });
      F(f, "emailNotif", b(/11\.2 Correo/), u("ref. 11.2"));
      F(f, "emailCliente", b(/10\.1\.1 Correo/), u("ref. 10.1.1 (correo autorizado)"));
      const firmaDate = grab(pageFlat(dg), new RegExp("Ciudad de M[\u00e9e]xico,?\\s*(?:a\\s+)?" + DATE_WORDS, "i"));
      if (firmaDate) x.fechaBloqueFirma = { v: firmaDate, ubic: u("bloque de firmas") };
    }
    // Anexo de Datos Generales \u2014 persona f\u00edsica (referencia 4)
    const dgf = P.find(p => /Referencia 4\. Datos Generales de la? ?Acreditad[ao]\s*\(Persona F[\u00edi]sica\)/i.test(pageFlat(p)));
    if (dgf) {
      x.personaFisica = true;
      const u = r => "Anexo de Datos Generales, " + r + " (" + pg(dgf) + ")";
      const R4 = /Referencia 4\./, R9 = /Referencia 9\./;
      const b = (re, after, o) => below(dgf, re, { after, ...o })?.v;
      F(f, "razon", b(/4\.1 Nombre completo/), u("ref. 4.1"));
      F(f, "lugarNac", b(/4\.2 Lugar de nacimiento/), u("ref. 4.2"));
      F(f, "fechaNac", b(/^Fecha de Nacimiento/i, R4), u("ref. 4.2"));
      x.estadoCivilPF = b(/^Estado Civil/i, R4); F(f, "estadoCivil", x.estadoCivilPF, u("ref. 4.2"));
      F(f, "curp", b(/^CURP:?$/i, R4), u("ref. 4"));
      F(f, "rfc", b(/4\.3 RFC/), u("ref. 4.3"));
      F(f, "idPF", b(/N[\u00fau]mero de identificaci/i, R4), u("ref. 4 'N\u00famero de identificaci\u00f3n'"));
      F(f, "ocupacion", b(/^Ocupaci[\u00f3o]n/i, R4), u("ref. 4"));
      F(f, "telefono", b(/N[\u00fau]mero de tel[\u00e9e]fono/i, R4), u("ref. 4"));
      F(f, "emailCliente", b(/^Correo electr[\u00f3o]nico:?$/i, R4), u("ref. 4"));
      const part = addrBlock(dgf, /^Domicilio Particular/i), fis = addrBlock(dgf, /^Domicilio Fiscal/i), ope = addrBlock(dgf, /^Domicilio Operativo/i);
      const fisLabel = (pageFlat(dgf).match(/Domicilio Fiscal( y operativo)?/i) || [])[0] || "Domicilio Fiscal";
      x.domPF = [];
      if (part) x.domPF.push({ tipo: "Particular", ...part, ubic: u("ref. 4 'Domicilio Particular'") });
      if (fis) x.domPF.push({ tipo: /operativo/i.test(fisLabel) ? "Fiscal y operativo" : "Fiscal", ...fis, ubic: u("ref. 4 '" + fisLabel + "'") });
      if (ope) x.domPF.push({ tipo: "Operativo", ...ope, ubic: u("ref. 4 'Domicilio Operativo'") });
      if (fis) { F(f, "domCalle", fis.calle, u("ref. 4 domicilio fiscal")); F(f, "domColonia", fis.colonia, u("ref. 4 domicilio fiscal")); F(f, "domMunicipio", fis.municipio, u("ref. 4 domicilio fiscal")); F(f, "domCP", fis.cp, u("ref. 4 domicilio fiscal")); }
      const cta = b(/4\.4 Cuenta de Dispersi/i), clabe = b(/4\.4\.1 Clabe/i), banco = b(/4\.4\.2 Instituci/i);
      if (clabe) x.clabes.push({ uso: "Acreditada \u2014 cuenta de dispersi\u00f3n (ref. 4.4)", banco, clabe: clabe.replace(/\D/g, ""), cuenta: cta, ubic: u("ref. 4.4") });
      F(f, "monto", money(b(/8\.1 Monto del Cr/)), u("ref. 8.1"));
      x.montoLetra = { v: b(/Monto del cr[\u00e9e]dito en letra/, null, { lines: 3, stop: /^\d+\.\d|Referencia|Tasa/i }), ubic: u("ref. 8.1 (letra)") };
      F(f, "fechaInicio", b(/8\.2 Fecha de inicio/), u("ref. 8.2"));
      F(f, "vencimiento", b(/8\.3 Fecha de vencimiento/), u("ref. 8.3"));
      F(f, "tasaOrd", pct(b(/^8\.4 Tasa/)), u("ref. 8.4"));
      F(f, "tasaMor", pct(b(/8\.4\.1\.? Tasa/)), u("ref. 8.4.1"));
      const clabeF = b(/Cuenta Clabe/, R9), refP = b(/^Referencia:/, R9), bancoF = b(/Instituci[\u00f3o]n Bancaria/, R9), titF = b(/^Titular/, R9);
      F(f, "referencia", refP, u("ref. 9"));
      if (clabeF) x.clabes.push({ uso: "Fideicomiso \u2014 cuenta concentradora (ref. 9)", banco: bancoF, clabe: clabeF.replace(/\D/g, ""), cuenta: "", ubic: u("ref. 9"), titular: titF });
      F(f, "emailNotif", b(/11\.2 Correo/), u("ref. 11.2"));
      const firmaDate = grab(pageFlat(dgf), new RegExp("Ciudad de M[\u00e9e]xico,?\\s*(?:a\\s+)?" + DATE_WORDS, "i"));
      if (firmaDate) x.fechaBloqueFirma = { v: firmaDate, ubic: u("bloque de firmas") };
    }
    x.sinObligado = /referencia N\/A del Anexo de Datos Generales,?\s*\(el .Obligado/i.test(P.slice(0, 3).map(pageFlat).join(" "));
    // Anexo A
    const aa = P.find(p => /ANEXO A/.test(pageFlat(p)) && /LISTADO DE SUCURSALES/i.test(pageFlat(p)));
    if (aa) {
      const lines = buildLines(aa);
      const i0 = lines.findIndex(l => /^Sucursal\b/.test(l.text)), i1 = lines.findIndex((l, j) => j > i0 && /^CUENTAS BANCARIAS/.test(l.text));
      if (i0 >= 0 && i1 > i0) {
        const two = lines.slice(i0 + 1, i1).filter(l => l.segs.length >= 2).map(l => l.segs[1].x0); const xsplit = two.length ? Math.min(...two) - 2 : 300;
        let addr = "", name = "";
        const flush = () => { if (addr.trim()) x.sucursales.push({ nombre: name.trim(), dom: addr.trim().replace(/,\s*$/, "") }); addr = ""; name = ""; };
        for (const l of lines.slice(i0 + 1, i1)) {
          const left = l.items.filter(i => i.x0 < xsplit).map(i => i.s).join(" "), rightT = l.items.filter(i => i.x0 >= xsplit).map(i => i.s).join(" ");
          if (left.trim() && !addr.trim()) name += " " + left; else if (left.trim()) name += " " + left;
          addr += " " + rightT;
          if (/C\.\s?P\.?\s*\d{5}/.test(addr)) flush();
        }
        flush();
        F(f, "sucursales", x.sucursales.length, "Anexo A (" + pg(aa) + ")");
      }
      for (const l of lines) { const m = l.text.match(/(\d{18})/); if (m) x.clabes.push({ uso: "Acreditado \u2014 cuenta listada en Anexo A", banco: l.text.replace(m[1], "").trim(), clabe: m[1], cuenta: "", ubic: "Anexo A (" + pg(aa) + ")" }); }
    }
    // Anexo B
    const ab = P.find(p => /ANEXO B/.test(pageFlat(p)) && /procesamiento de pagos/i.test(pageFlat(p)));
    if (ab) {
      const t = pageFlat(ab);
      const m = t.match(/(\d{2}\/\d{2}\/\d{4})\s+(.+?)\s+Se entrega/);
      if (m) { F(f, "fechaProcesador", m[1], "Anexo B (" + pg(ab) + ")"); F(f, "procesador", m[2], "Anexo B (" + pg(ab) + ")"); }
      else { const d = grab(t, /Nombre del procesador\s+(\d{2}\/\d{2}\/\d{4})\s+Se entrega/i); if (d) { F(f, "fechaProcesador", d, "Anexo B (" + pg(ab) + ")"); x.procesadorVacio = { fecha: d, ubic: "Anexo B (" + pg(ab) + ")" }; } }
    }
    // Contrato del procesador: hoja de firmas
    const gsig = P.find(p => /REPRESENTADO POR/i.test(pageFlat(p)) && /AFILIADO/i.test(pageFlat(p)));
    if (gsig) {
      const t = pageText(gsig);
      let nm = grab(t, /REPRESENTADO POR[ \t]*:[ \t]*(.+)/i); if (nm && (/se.alar nombre|^\W*$|^REPRESENTADO/i.test(nm) || nm.replace(/[^A-Za-z\u00c1\u00c9\u00cd\u00d3\u00da\u00d1\u00e1\u00e9\u00ed\u00f3\u00fa\u00f1]/g, "").length < 4)) nm = null;
      const bad = garbledRatio(t) > 0.004 || (nm && /[{}|\\$]/.test(nm));
      F(f, "firmanteProcesador", nm, "Contrato del procesador, hoja de firmas (" + pg(gsig) + ")", { ilegible: bad && gsig.source !== "ocr", conf: gsig.source === "ocr" ? gsig.conf : null });
      const fd = grab(t, /fecha de afiliaci[\u00f3o]n\s*(.+)/i) || grab(t, /Fecha:?\s*(\d{2}\/\d{2}\/\d{4})/i) || grab(t, new RegExp("Fecha:?\\s*" + DATE_WORDS, "i"));
      F(f, "fechaProcesador", fd, "Contrato del procesador, hoja de firmas (" + pg(gsig) + ")");
      F(f, "razon", grab(t, /DENOMINACI[\u00d3O]N O NOMBRE\s*:\s*(.+)/i), "Contrato del procesador, hoja de firmas (" + pg(gsig) + ")", { conf: gsig.source === "ocr" ? gsig.conf : null, ocr: gsig.source === "ocr" });
      x.procSigPage = gsig.no;
    }
    const bp = P.find(p => /Billpocket/i.test(pageFlat(p)) && /El Cliente/.test(pageFlat(p)) && /ID No/i.test(pageFlat(p)));
    if (bp) {
      const lines = buildLines(bp); const i = lines.findIndex(l => /^El Cliente/.test(l.text));
      if (i >= 0) F(f, "firmanteProcesador", lines[i + 1]?.text, "Contrato Billpocket, firma (" + pg(bp) + ")");
      F(f, "idObligado", grab(pageFlat(bp), /ID No\.?:?\s*(\d+)/i), "Contrato Billpocket, firma (" + pg(bp) + ")");
      // Solo el bloque de firma del Cliente: el RFC que sigue a "Billpocket;" es el de Pocket de Latinoam\u00e9rica (PLA120807BK0)
      const bpCli = pageFlat(bp).split(/Billpocket;|Pocket de Latinoam/i)[0];
      F(f, "rfc", grab(bpCli, /RFC:\s*([A-Z&\u00d1]{3,4}\d{6}[A-Z0-9]{3})/), "Contrato Billpocket, firma del Cliente (" + pg(bp) + ")");
      x.procSigPage = bp.no; x.procFirmaTipografica = true;
    }
    const bpc = P.find(p => /Billpocket/i.test(pageFlat(p)) && /Cuenta Bancaria CLABE/i.test(pageFlat(p)));
    if (bpc) { const c = grab(pageFlat(bpc), /CLABE\s+(\d{15,19})\b/); if (c) x.clabes.push({ uso: "Dep\u00f3sito de ventas del procesador (Billpocket)", banco: grab(pageFlat(bpc), /Instituci[\u00f3o]n Financiera de la CLABE\s+(.+?)\s+3\./), clabe: c, cuenta: "", ubic: "Contrato Billpocket 2.3 (" + pg(bpc) + ")" }); x.descriptor = grab(pageFlat(bpc), /(BPK\*\S+\s+.+?)\s+Cuenta Bancaria/); }
    // Anexo D
    const ad = P.find(p => /LISTADO DE COMISIONES/i.test(pageFlat(p)) && /Comisi[\u00f3o]n por gesti[\u00f3o]n/i.test(pageFlat(p)));
    if (ad) {
      const t = pageFlat(ad);
      const pc = [...t.matchAll(/(\d{1,2}(?:\.\d{1,2})?)\s?%/g)].map(m => +m[1]);
      const am = [...t.matchAll(/\$\s?([\d,]+\.\d{2})\s*m[\u00e1a]s IVA/g)].map(m => money(m[1]));
      x.anexoD = { pInv: pc[0], pDisp: pc[1], pSDGC: pc[2], inv: am[0], disp: am[1], sdgc: am[2], pena: am[3], ubic: "Anexo D (" + pg(ad) + ")" };
      F(f, "comInv", am[0], "Anexo D (" + pg(ad) + ")"); F(f, "comDisp", am[1], "Anexo D (" + pg(ad) + ")"); F(f, "comSDGC", am[2], "Anexo D (" + pg(ad) + ")");
    }
    // Anexo F
    const af = P.find(p => /ANEXO F/.test(pageFlat(p)) && /TABLA DE AMORTIZACI/i.test(pageFlat(p)));
    if (af) {
      const t = pageText(af), u = "Anexo F (" + pg(af) + ")";
      F(f, "credito", grab(t, /Cr[\u00e9e]dito:\s*(\S+)/), u); F(f, "referencia", grab(t, /Referencia:\s*(\S+)/), u);
      F(f, "total", money(grab(t, /Monto total a pagar:\s*\$?([\d,.]+)/)), u);
      const rows = amortRows(t); x.rowsF = rows;
      if (rows.length) { F(f, "primerPago", rows[0].fecha, u); F(f, "vencimiento", rows[rows.length - 1].fecha, u + " \u00faltimo pago"); F(f, "pagoMensual", modeOf(rows.map(r => r.pago)), u); F(f, "numPagos", rows.length, u); }
    }
    // domiciliaciones
    P.filter(p => /domiciliaci[\u00f3o]n de pagos/i.test(pageFlat(p))).forEach(p => {
      const t = pageText(p), u = "Domiciliaci\u00f3n (" + pg(p) + ")";
      const d = { credito: grab(t, /Cr[\u00e9e]dito a pagar:\s*(\S+)/), cuenta: grab(t, /cargo:\s*(\S+)/), banco: grab(t, /Banco donde se maneja la cuenta:\s*(.+)/), clabe: grab(t, /CLABE de la cuenta:\s*(\d{18})/), monto: money(grab(t, /Monto del cargo mensual que se realizar[\u00e1a]:\s*\$?([\d,.]+)/)), autoriza: grab(t, /Autoriza:\s*(.+)/), multiplos: grab(t, /m[\u00fau]ltiplos de \$?([\d,]+)/), ubic: u };
      x.domis.push(d);
      if (d.clabe) x.clabes.push({ uso: "Acreditado \u2014 domiciliaci\u00f3n", banco: d.banco, clabe: d.clabe, cuenta: d.cuenta, ubic: u });
      F(f, "credito", d.credito, u); F(f, "pagoDomiciliado", d.monto, u);
    });
    x.ivaIntereses = /m[\u00e1a]s el Impuesto al Valor Agregado/i.test(P.map(pageFlat).join(" "));
    return { fields: f, x };
  };

  E.extractPagare = function (doc) {
    const f = {}, x = {}; const all = doc.pages.map(pageFlat).join(" "), txt = doc.pages.map(pageText).join("\n");
    const m = all.match(/\$\s?([\d,]+\.\d{2})\s*\(([^)]+)\)\s*\(la .Suma Principal.\)/i);
    if (m) { F(f, "sumaPrincipal", money(m[1]), "P\u00e1rrafo 1 (p.1)"); x.letra = { v: m[2], ubic: "P\u00e1rrafo 1 (p.1)" }; }
    F(f, "razon", grab(all, /POR VALOR RECIBIDO,\s*(.+?)\s*\(el .Suscriptor.\)/i), "P\u00e1rrafo 1 (p.1)");
    F(f, "tasaOrd", pct(grab(all, /tasa de inter[\u00e9e]s fija del\s*([\d.]+)%/i)), "p.1 tasa ordinaria");
    F(f, "tasaMor", pct(grab(all, /moratoria anual del\s*([\d.]+)%/i)), "p.1 tasa moratoria");
    const rows = amortRows(txt); x.rows = rows;
    if (rows.length) { F(f, "primerPago", rows[0].fecha, "Calendario, pago 1"); F(f, "vencimiento", rows[rows.length - 1].fecha, "Calendario, \u00faltimo pago"); F(f, "pagoMensual", modeOf(rows.map(r => r.pago)), "Calendario"); F(f, "numPagos", rows.length, "Calendario"); F(f, "total", r2(rows.reduce((a, r) => a + r.pago, 0)), "Calendario (suma de pagos)"); }
    const lines = doc.pages.flatMap(p => buildLines(p).map(l => ({ ...l, pno: p.no })));
    const iS = lines.findIndex(l => /^Suscriptor:$/.test(l.text)), iA = lines.findIndex(l => /^Aval:$/.test(l.text));
    if (iS >= 0) {
      const buf = []; for (let j = iS + 1; j < lines.length && j < iS + 6; j++) { const tx = lines[j].text; if (/^(Aval:|Este Pagar|\d+$|No\. RECA)/i.test(tx)) break; buf.push(tx); }
      const full = buf.join(" ").replace(/\s+/g, " ").trim(); const parts = full.split(/\s+y\/o\s+/i).map(t => t.trim()).filter(Boolean);
      x.domicilios = parts.map(t => ({ texto: t, cp: grab(t, /C\.?\s?P\.?\s*(\d{5})/i), ubic: "Domicilio del suscriptor (p." + lines[iS].pno + ")" }));
      if (parts.length === 1) { F(f, "domCalle", parts[0], "Domicilio del suscriptor (p." + lines[iS].pno + ")"); F(f, "domCP", x.domicilios[0].cp, "Domicilio del suscriptor (p." + lines[iS].pno + ")"); }
    }
    if (iA >= 0) { F(f, "avalCalle", lines[iA + 1]?.text, "Domicilio del aval (p." + lines[iA].pno + ")"); F(f, "avalCP", grab(lines[iA + 1]?.text, /C\.P\.\s*(\d{5})/), "Domicilio del aval (p." + lines[iA].pno + ")"); }
    F(f, "fechaFirma", grab(all, new RegExp("medios electr[\u00f3o]nicos, el " + DATE_WORDS, "i")), "Lugar y fecha de suscripci\u00f3n");
    x.hojasTexto = grab(all, /se suscribe en (\d+ Hojas?)/i);
    x.conAval = lines.some(l => /^.?Aval 1.?$/.test(l.text.trim()));
    const iR = lines.findIndex(l => /^Representante Legal$/.test(l.text));
    if (iR > 0) F(f, "representante", lines[iR - 1].text, "Firma del suscriptor (p." + lines[iR].pno + ")");
    const iAv = lines.findIndex(l => /Aval 1/.test(l.text));
    if (iAv >= 0) { const nm = lines.slice(iAv + 1).find(l => /[A-Z\u00c1\u00c9\u00cd\u00d3\u00da\u00d1]{3,}/.test(l.text) && !/^_+$/.test(l.text) && !/^\d+$/.test(l.text)); if (nm) F(f, "obligado", nm.text, "Firma del aval (p." + nm.pno + ")"); }
    return { fields: f, x };
  };

  E.extractConvenio = function (doc) {
    const f = {}, x = { clabes: [] }; const P = doc.pages, all = P.map(pageFlat).join(" ");
    const pr = all.match(/celebran, por una parte,\s*(.+?),\s*representada en este acto por\s*(.+?),\s*\(el .Fideicomitente/i);
    if (pr) { F(f, "razon", pr[1], "Proemio (p.1)"); F(f, "representante", pr[2], "Proemio (p.1)"); }
    else { const pf = all.match(/celebran, por una parte,\s*(.+?)\s*\(el .Fideicomitente Adherente/i); if (pf) { F(f, "razon", pf[1], "Proemio (p.1)"); x.personaFisica = true; } }
    x.estadoCivil = grab(all, /Es una persona f[\u00edi]sica,\s*([^,]+?)\s*(?:,|mismo)/i);
    F(f, "fideicomiso", grab(all, /Fuente de Pago No\.?\s*(F\/\d+)/i), "Proemio (p.1)");
    F(f, "fechaFideicomiso", grab(all, new RegExp("I\\.\\s*Con fecha " + DATE_WORDS, "i")), "Antecedente I");
    F(f, "fechaFirma", grab(all, new RegExp("II\\.\\s*Con fecha " + DATE_WORDS, "i")), "Antecedente II");
    const mm = all.match(/por la cantidad de \$([\d,]+\.\d{2})\s*M\.N\.\s*\(([^)]+)\)/i);
    if (mm) { F(f, "monto", money(mm[1]), "Antecedente II"); x.letra = { v: mm[2], ubic: "Antecedente II" }; }
    x.fechaCierre = { v: grab(all, new RegExp("celebran el presente Contrato, el d[\u00edi]a " + DATE_WORDS, "i")), ubic: "Cierre del convenio" };
    const oct = all.match(/Fideicomitente Adherente:\s*(.+?)\s+Correo Electr[\u00f3o]nico:\s*(\S+)/i);
    if (oct) { let dom = oct[1].replace(/^.*?C\.V\.\s*/i, ""); const pfConv = !!x.personaFisica; const nmC = f.razon?.v; if (nmC && norm(dom).startsWith(norm(nmC))) dom = dom.slice(dom.toUpperCase().indexOf(String(nmC).toUpperCase().split(" ").pop()) + String(nmC).split(" ").pop().length).trim(); x.domOct = dom; if (!pfConv) F(f, "domCalle", dom, "Cl\u00e1usula Octava, domicilios"); F(f, "domCP", grab(dom, /C\.P\.\s*(\d{5})/), "Cl\u00e1usula Octava, domicilios"); F(f, "emailNotif", oct[2], "Cl\u00e1usula Octava, domicilios"); }
    const aa = all.match(/Contratos de Servicios[\s\S]{0,80}?(\d{2}\/\d{2}\/\d{4})\s+(.+?)\s+El Fideicomitente/i);
    if (aa && !/^El Fideicomitente/i.test(aa[2]) && aa[2].trim()) { F(f, "fechaProcesador", aa[1], "Anexo A del convenio"); F(f, "procesador", aa[2].replace(/Nombre del procesador\s*/i, ""), "Anexo A del convenio"); }
    else { const d = grab(all, /Contratos de Servicios[\s\S]{0,80}?Nombre del procesador\s+(\d{2}\/\d{2}\/\d{4})\s+El Fideicomitente/i); if (d) { F(f, "fechaProcesador", d, "Anexo A del convenio"); x.procesadorVacio = { fecha: d, ubic: "Anexo A del convenio" }; } }
    const pb = P.find(p => /Anexo .B. del Convenio/i.test(pageFlat(p)) && /Distribuci[\u00f3o]n de Cobranza/i.test(pageFlat(p)));
    if (pb) {
      const t = pageFlat(pb);
      const b = t.match(/(\d-\d{4}-\d{3})[\s\S]{0,300}?(\d{2}\/\d{2}\/\d{1,4})[\s\S]{0,120}?\$?\s?([\d,]+\.\d{2})[\s\S]{0,120}?\$?\s?([\d,]+\.\d{2})[\s\S]{0,160}?(\d{1,2})\s?%[\s\S]{0,160}?\$\s?([\d,]+\.\d{2})/);
      if (b) { const u = "Anexo B 'Distribuci\u00f3n de Cobranza' (p." + pb.no + ")"; F(f, "credito", b[1], u); x.anexoBFecha = b[2]; F(f, "monto", money(b[3]), u); x.montoPagare = { v: money(b[4]), ubic: u }; F(f, "retencion", +b[5], u); F(f, "pagoMinimo", money(b[6]), u); }
    }
    const np = P.find(p => /RE:\s*Notificaci[\u00f3o]n/i.test(pageFlat(p)) && !/\{Procesador\}|\{Lugar/.test(pageFlat(p)));
    if (np) {
      const t = pageFlat(np), u = "Notificaci\u00f3n al procesador (p." + np.no + ")";
      const afi = grab(t, /Contrato\s+(.+?)\s+Estimados/i) || "";
      x.afiliaciones = [...afi.matchAll(/\b(\d{7})\b/g)].map(m => m[1]);
      if (x.afiliaciones.length) F(f, "afiliaciones", x.afiliaciones.length, u);
      x.descriptor = grab(afi, /(BP\*\S+.*)/);
      F(f, "fechaProcesador", grab(t, new RegExp("celebrado el " + DATE_WORDS, "i")), u);
      x.notifSinProcesador = /\by por virtud del cual se oblig|instruir a para que|derechos del respecto|presente, declara y manifiesta/i.test(t) ? u : null;
      x.notifEncabezado = afi.replace(/\s+/g, " ").trim();
    }
    const rest = P.map(pageFlat).join(" ");
    const cl = rest.match(/n[\u00fau]mero (\d{10}),?\s*CLABE No\.?\s*(\d{18})\s*que el Fideicomiso mantiene abierta con\s*(\w+)/i);
    if (cl) x.clabes.push({ uso: "Fideicomiso \u2014 cuenta concentradora", banco: cl[3], clabe: cl[2], cuenta: cl[1], ubic: "Notificaci\u00f3n al procesador" });
    return { fields: f, x };
  };

  function extractKycPF(doc) {
    const f = {}, x = { pf: true }; const P = doc.pages;
    const all = P.map(pageText).join("\n"); const up = deacc(all).toUpperCase();
    const p1 = P[0], last = P[P.length - 1];
    const u = (p, r) => "KYC p." + p.no + " '" + r + "'";
    const ocrF = (p, key, v, r) => F(f, key, v, u(p, r), { conf: p.conf ?? null, ocr: true }); // el KYC es imagen: su texto siempre viene de OCR
    x.texto = all;
    // fecha del formato
    const head = buildLines(p1).slice(0, 8).map(l => l.text).join(" ");
    const fd = grab(head, /(\d{2}\/\d{2}\/\d{4})/) || grab(deacc(head), /(\d{1,2}\s*\|?\s*(?:enero|febrero|marzo|abril|mayo|junio|julio|agosto|septiembre|octubre|noviembre|diciembre)\s*\|?\s*\d{4})/i);
    if (fd) ocrF(p1, "fechaFirma", fd.replace(/\|/g, " ").replace(/\s+/g, " "), "Fecha");
    ocrF(p1, "fideicomiso", grab(up, /\b(F\/\d{4})\b/), "N\u00famero de FID");
    x.ejecutivo = (below(p1, /NOMBRE DEL EJECUTIVO/i)?.v || "").replace(/^\s*F\/\d+\s*/, "").replace(/\s*\d{1,2}\s+[A-Za-z]+\s+\d{4}\s*$/, "").trim() || null;
    const rfc = grab(up.replace(/\s/g, " "), /\b([A-Z\u00d1&]{4}\d{6}[A-Z0-9]{3})\b/); if (rfc) ocrF(p1, "rfc", rfc, "RFC");
    let curp = grab(up, /\b([A-Z]{4}\d{6}[HM][A-Z]{5}[A-Z0-9][0-9O])\b/); if (curp) { curp = curp.slice(0, 17) + curp[17].replace("O", "0"); ocrF(p1, "curp", curp, "CURP"); }
    const em = all.replace(/\s+/g, " ").match(/([A-Za-z0-9._-]{3,}?)\s?[@O0\u00a9\u00ae(QG]{0,2}\s?((?:gmail|outlook|hotmail|yahoo|live|icloud)\.com)/i);
    if (em) ocrF(p1, "emailCliente", (em[1] + "@" + em[2]).toLowerCase(), "Correo electr\u00f3nico");
    x.telefonos = [...new Set((all.match(/\b\d{10}\b/g) || []))];
    x.estadoCivil = below(p1, /^ESTADO CIVIL/i)?.v || null;
    x.ocupacion = below(p1, /OCUPACI[\u00d3O]N O PROFESI/i)?.v || null;
    x.giro = below(p1, /ACTIVIDAD O GIRO|GIRO DEL NEGOCIO/i)?.v || null;
    x.empresa = below(p1, /EMPRESA DONDE LABORA/i)?.v || null;
    // domicilios: cada etiqueta CALLE con su n\u00famero, colonia y C.P.
    x.domicilios = [];
    P.forEach(pg => { const lines = buildLines(pg); lines.forEach((ln, i) => { ln.segs.forEach(sg => { if (/^CALLE$/i.test(sg.text.trim())) {
      const calle = ocrClean(valUnder(pg, i, sg) || ""); if (!calle || /^N\/?A$/i.test(calle) || !/[A-Za-z0-9]{2}/.test(calle) || /RESIDENCIA|COLONIA|POSTAL|EXTERIOR/i.test(calle)) return;
      const ext = (() => { const s2 = ln.segs.find(x2 => /EXTERIOR/i.test(x2.text)); return s2 ? valUnder(pg, i, s2) : null; })();
      const int = (() => { const s2 = ln.segs.find(x2 => /INTERIOR/i.test(x2.text)); return s2 ? valUnder(pg, i, s2) : null; })();
      let col = null, cp = null, anios = null;
      for (let j = i + 1; j < Math.min(lines.length, i + 5); j++) { const l2 = lines[j];
        const sc = l2.segs.find(x2 => /^COLONIA/i.test(x2.text)); if (sc && !col) col = valUnder(pg, j, sc);
        const sp = l2.segs.find(x2 => /POSTAL/i.test(x2.text)); if (sp && !cp) cp = grab(valUnder(pg, j, sp), /(\d{5})/);
        const sa = l2.segs.find(x2 => /A[\u00d1N]OS RESIDENCIA/i.test(x2.text)); if (sa && !anios) anios = grab(valUnder(pg, j, sa), /(\d+)/); }
      const head2 = lines.slice(Math.max(0, i - 3), i).map(l => l.text).join(" ");
      const tipo = /FISCAL/i.test(head2) ? "Fiscal" : /APODERADO/i.test(head2) ? "Apoderado" : /PROVEEDOR/i.test(head2) ? "Proveedor" : "Particular / del cliente";
      if (/Apoderado|Proveedor/.test(tipo)) return;
      x.domicilios.push({ tipo, calle: [calle, ext, int].filter(Boolean).join(" "), colonia: col, cp, anios, ubic: u(pg, "Domicilio " + tipo.toLowerCase()) });
    } }); }); });
    if (x.domicilios[0]) { ocrF(P[0], "domCP", x.domicilios.find(d => d.tipo === "Fiscal")?.cp || null, "C. Postal (domicilio fiscal)"); }
    // fines del fideicomiso: n\u00famero de contrato citado
    const t2 = up.replace(/\s+/g, " ");
    let cn = grab(t2, /#\s?(\d{6})\b/) || grab(t2, /(?:FINSUS|ADVANCE)\s*[#+\W]{0,3}[A-Za-z]?\s*([4H#T]?\d{6})\b/i);
    if (cn && cn.length === 7) cn = cn.slice(1);
    if (cn) { const pgc = P.find(pg => pageFlat(pg).replace(/\s/g, "").includes(cn)) || p1; x.contrato = { v: cn, ubic: u(pgc, "Fines que pretende dar al fideicomiso"), conf: pgc.conf }; }
    const opsPg = P.find(pg => /MONTO APROXIMADO DE (OPERACIONES|MOVIMIENTOS)/i.test(pageFlat(pg)));
    if (opsPg) x.opsMensuales = { v: money(below(opsPg, /MONTO APROXIMADO DE (OPERACIONES|MOVIMIENTOS)/i)?.v), ubic: u(opsPg, "Monto aproximado mensual") };
    const montoPg = P.find(pg => /MONTO DE LA OPERACI/i.test(pageFlat(pg)));
    if (montoPg) F(f, "monto", money(below(montoPg, /MONTO DE LA OPERACI/i)?.v), u(montoPg, "Monto de la operaci\u00f3n / contrato"), { conf: montoPg.conf, ocr: montoPg.source === "ocr" });
    // firmas: nombres en "NOMBRE Y FIRMA"
    const sigPg = [...P].reverse().find(pg => /NOMBRE Y FIRMA/i.test(pageFlat(pg))) || last;
    x.nombreCliente = below(sigPg, /NOMBRE Y FIRMA/i)?.v || null;
    x.entrevistador = below(sigPg, /NOMBRE Y FIRMA/i, { nth: 1 })?.v || null;
    x.seccionesPM = /PERSONA MORAL/i.test(up) && /VEINTICINCO POR CIENTO|25%/i.test(up);
    return { fields: f, x };
  }

  E.extractKyc = function (doc) {
    if (doc.pf) return extractKycPF(doc);
    const f = {}, x = {}; const P = doc.pages;
    const p1 = P[0], p2 = P[1], p4 = P[3] || P[P.length - 1];
    const u = (p, r) => "KYC p." + p.no + " '" + r + "'";
    const ocrF = (p, key, v, r) => F(f, key, v, u(p, r), { conf: p.conf ?? null, ocr: true });
    if (p1) {
      const b = (re, o) => below(p1, re, o)?.v;
      ocrF(p1, "fechaFirma", b(/^FECHA:?$/i) || b(/FECHA:/i), "Fecha");
      ocrF(p1, "fideicomiso", grab(b(/N[\u00daU]MERO DE FID/i), /(F\/\d+)/), "N\u00famero de FID");
      x.ejecutivo = (b(/NOMBRE DEL EJECUTIVO/i) || "").replace(/^\s*F\/\d+\s*/, "").trim() || null;
      ocrF(p1, "razon", b(/DENOMINACI[\u00d3O]N O RAZ[\u00d3O]N SOCIAL/i), "Denominaci\u00f3n o raz\u00f3n social");
      x.giro = b(/GIRO, ACTIVIDAD/i);
      ocrF(p1, "rfc", (b(/RFC \(HOMOCLAVE\)/i) || "").toUpperCase().replace(/\s/g, ""), "RFC");
      ocrF(p1, "escrituraConst", b(/ESCRITURA P[\u00daU]BLICA/i), "Escritura p\u00fablica");
      ocrF(p1, "fechaConst", b(/FECHA DE CONSTITUCI/i), "Fecha de constituci\u00f3n");
      ocrF(p1, "folio", grab(b(/INSCRIPCI[\u00d3O]N EN EL R\.?P\.?C/i), /(N-\d+)/), "Inscripci\u00f3n en el R.P.C.");
      x.administrador = b(/ADMINISTRADOR/i);
      // El encabezado "DIRECTOR GENERAL, GERENTE GENERAL O APODERADO LEGAL QUE CON SU FIRMA PUEDA OBLIGAR / A LA PERSONA MORAL:"
      // a veces se parte en dos renglones: el segundo ("A LA PERSONA MORAL:") no es la respuesta.
      let apod = b(/DIRECTOR GENERAL, GERENTE GENERAL/i);
      if (apod && /A LA PERSONA MORAL|PUEDA OBLIGAR/i.test(apod)) apod = b(/A LA PERSONA MORAL:?/i);
      if (!apod || /^(A LA|QUE CON SU)/i.test(apod)) apod = b(/^NOMBRE\(S\) DEL \(LOS\) ADMINISTRADOR/i) || apod;
      ocrF(p1, "representante", apod, "Apoderado legal");
      ocrF(p1, "escrituraPoder", b(/ESCRITURA P[\u00daU]BLICA/i, { after: /DIRECTOR GENERAL/i }), "Escritura del poder");
      ocrF(p1, "fechaPoder", b(/DE FECHA/i, { after: /DIRECTOR GENERAL/i }), "Fecha del poder");
      const calle = b(/^CALLE$/i, { after: /DOMICILIO PERSONA MORAL/i }), ext = b(/N[\u00daU]M\.? EXTERIOR/i), int = b(/N[\u00daU]M\.? INTERIOR/i);
      ocrF(p1, "domCalle", [calle, ext, int].filter(Boolean).join(" "), "Calle / N\u00fam. exterior / interior");
      ocrF(p1, "domColonia", b(/^COLONIA:?$/i, { after: /DOMICILIO PERSONA MORAL/i }), "Colonia");
      ocrF(p1, "domCP", grab(b(/C\.? ?POSTAL/i), /(\d{5})/), "C. Postal");
      ocrF(p1, "domMunicipio", b(/ALCALD[\u00cdI]A O MUNICIPIO/i), "Alcald\u00eda o municipio");
      x.anios = { v: grab(b(/A[\u00d1N]OS RESIDENCIA/i), /(\d+)/), ubic: u(p1, "A\u00f1os residencia"), conf: p1.conf };
    }
    if (p2) {
      const b = (re, o) => below(p2, re, o)?.v; const t = pageFlat(p2);
      x.telefonos = [b(/^OFICINA 1/i), b(/^CELULAR/i), b(/^OFICINA 2/i)].filter(Boolean);
      ocrF(p2, "emailCliente", (b(/CORREO ELECTR/i) || "").toLowerCase().replace(/\s/g, "").replace(/[(\u00a9\u00ae]/g, "@").replace(/\|.*$/, ""), "Correo electr\u00f3nico");
      x.contrato = { v: (c => c && c.length === 7 && /^[4H#]/.test(c) ? c.slice(1) : c)(grab(t, /FINSUS\s*#?\s*(\d{6,7})/i) || grab(t, /#\s?(\d{5,7})/)), ubic: u(p2, "Fines que pretende dar al fideicomiso"), conf: p2.conf };
      x.opsMensuales = { v: money(b(/MONTO APROXIMADO DE OPERACIONES/i)), ubic: u(p2, "Monto aproximado de operaciones mensuales") };
      x.patrimonio = money(b(/VALOR PROMEDIO ESPERADO/i));
    }
    if (p4) {
      const lines = buildLines(p4); const t = pageFlat(p4);
      const li = lines.findIndex(l => /IMPORTE DE VENTAS/i.test(l.text) || /0 A 1,?000,?000/.test(l.text));
      if (li >= 0) {
        const ln = lines[li]; const nl = lines.slice(li + 1, li + 3).find(l => /\bS[I\u00cd]\b/i.test(l.text));
        const BR = [["0 A 1,000,000", 0, 1e6], ["1,000,001 A 5,000,000", 1e6, 5e6], ["5,000,001 A 10,000,000", 5e6, 1e7], ["M\u00e1s de 10,000,000", 1e7, Infinity]];
        if (nl) { const si = nl.items.find(i => /^S[I\u00cd]$/i.test(i.s.trim())); if (si) { const segs = ln.segs.filter(s => /\d/.test(s.text)); let k = -1; segs.forEach((s, idx) => { if (si.x0 >= s.x0 - 20) k = idx; }); if (k >= 0 && segs.length >= 3) { const idx = Math.min(BR.length - 1, k + (BR.length - segs.length)); x.ventasAnuales = { label: BR[idx][0], min: BR[idx][1], max: BR[idx][2], ubic: u(p4, "Importe de ventas anuales"), conf: p4.conf }; } } }
      }
      const sl = lines.findIndex(l => /SUCURSALES/i.test(l.text) || /CUANTAS/i.test(l.text));
      if (sl >= 0) { const nums = lines.slice(sl, sl + 3).map(l => l.text).join(" ").match(/\b(\d{1,3})\b/g); if (nums) F(f, "sucursales", +nums[0], u(p4, "Cu\u00e1ntas sucursales tiene"), { conf: p4.conf, ocr: true }); }
      const ei = lines.findIndex(l => /ESTRUCTURA CORPORATIVA/i.test(l.text));
      x.accionistas = [];
      if (ei >= 0) for (const l of lines.slice(ei + 1, ei + 12)) { if (/NOTA:/i.test(l.text)) break; const m = l.text.match(/[-\u2013]?\s*([A-Z\u00c1\u00c9\u00cd\u00d3\u00da\u00d1][A-Z\u00c1\u00c9\u00cd\u00d3\u00da\u00d1 .]+?)\s*[-\u2013]?\s*(\d{1,3}(?:\.\d{1,2})?)\s?%/i); if (m) x.accionistas.push({ nombre: m[1].trim(), pct: +m[2] }); }
      x.entrevistador = below(p4, /^NOMBRE$/i, { nth: 1 })?.v || below(p4, /^NOMBRE$/i)?.v || null;
    }
    return { fields: f, x };
  };

  E.extractCsf = function (doc) {
    const f = {}, x = {}; const all = doc.pages.map(pageFlat).join(" ");
    F(f, "rfc", grab(all, /RFC:\s*([A-Z&\u00d1]{3,4}\d{6}[A-Z0-9]{3})/), "Datos de identificaci\u00f3n");
    F(f, "razon", grab(all, /Denominaci[\u00f3o]n\/Raz[\u00f3o]n Social:\s*(.+?)\s+R[\u00e9e]gimen Capital/i), "Datos de identificaci\u00f3n");
    x.regimen = grab(all, /R[\u00e9e]gimen Capital:\s*(.+?)\s+Nombre Comercial/i);
    F(f, "fechaConst", grab(all, /Fecha inicio de operaciones:\s*(.+?)\s+Estatus/i), "Fecha inicio de operaciones");
    F(f, "domCP", grab(all, /C[\u00f3o]digo Postal:\s*(\d{5})/i), "Domicilio registrado");
    const vial = grab(all, /Nombre de Vialidad:\s*(.+?)\s+N[\u00fau]mero Exterior/i), ext = grab(all, /N[\u00fau]mero Exterior:\s*(\S+)/i), int = grab(all, /N[\u00fau]mero Interior:\s*(.*?)\s+Nombre de la Colonia/i);
    F(f, "domCalle", [grab(all, /Tipo de Vialidad:\s*(.+?)\s+Nombre de Vialidad/i), vial, ext, int].filter(Boolean).join(" "), "Domicilio registrado");
    F(f, "domColonia", grab(all, /Nombre de la Colonia:\s*(.+?)\s+(?:Nombre de la Localidad|Nombre del Municipio)/i), "Domicilio registrado");
    F(f, "domMunicipio", grab(all, /Nombre del Municipio o Demarcaci[\u00f3o]n Territorial:\s*(.+?)\s+(?:Nombre de la Entidad|Nombre de la Localidad)/i), "Domicilio registrado");
    x.emision = parseDate(grab(all, /A\s+(\d{1,2}\s+DE\s+[A-Z]+\s+DE\s+\d{4})/i));
    return { fields: f, x };
  };

  const cell = (ws, a) => ws && ws[a] ? ws[a] : null;
  const cv = (ws, a) => { const c = cell(ws, a); return c ? c.v : null; };

  /* ---- pantalla "Datos Captura" del sistema (imagen) ---- */
  const CAP_LABELS = [
    ["fFinsusLetra", "F\\.?\\s*Contrato\\s*Finsus\\s*\\(\\s*L\\w*\\s*\\)\\s*:?"],
    ["fFinsus", "F\\.?\\s*Contrato\\s*Finsus\\s*:?"],
    ["nombreProc", "Nombre\\s*Contrato\\s*Procesador\\s*:?"],
    ["fProc", "F\\.?\\s*Contrato\\s*Procesador\\s*:?"],
    ["razonProc", "Raz[o\u00f3]n\\s*Social\\s*Procesador\\s*:?"],
    ["numProc", "[#*]?\\s*Contrato\\s*Procesador\\s*:?"],
    ["ocupacion", "Ocupaci[o\u00f3]n\\s*:?"],
    ["destino", "Destino\\s*de\\s*los\\s*Recursos\\s*:?"],
    ["tasaOrdLetra", "Tasa\\s*[Ii1l]?nter[e\u00e9]s\\s*Anual\\s*Ordinaria\\s*\\(\\s*L\\w*\\s*\\)\\s*:?"],
    ["tasaOrd", "Tasa\\s*[Ii1l]?nter[e\u00e9]s\\s*Anual\\s*Ordinaria\\s*:?"],
    ["tasaMorLetra", "Tasa\\s*[Ii1l]?nter[e\u00e9]s\\s*Moratoria\\s*\\(\\s*L\\w*\\s*\\)\\s*:?"],
    ["tasaMor", "Tasa\\s*[Ii1l]?nter[e\u00e9]s\\s*Moratoria\\s*:?"],
    ["clabeAcred", "Cuenta\\s*Clabe\\s*Acreditante\\s*:?"],
    ["cuentaAcred", "Cuenta\\s*del\\s*Acre\\w*\\s*:?"],
    ["lugar", "Lugar\\s*de\\s*f\\w*\\s*Contrato\\s*Finsus\\s*:?"],
    ["tipoSoc", "Tipo\\s*(?:de\\s*)?Sociedad\\s*:?"],
    ["pInv", "%?\\s*Comisi[o\u00f3]n\\s*(?:por\\s*)?Gastos\\s*de\\s*Investigaci[o\u00f3]n\\s*(?:[sS5]\\W?[il1I|]?\\W?va)?\\s*:?"],
    ["pDisp", "%?\\s*Comisi[o\u00f3]n\\s*(?:x|por)\\s*Disposici[o\u00f3]n\\s*(?:[sS5]\\W?[il1I|h]?\\W?[va]+)?\\s*:?"],
    ["pSDGC", "%?\\s*Comisi[o\u00f3]n\\s*por\\s*Gesti[o\u00f3]n\\s*de\\s*Cobranza\\s*(?:[sS5]\\W?[il1I|]?\\W?va)?\\s*:?"],
    ["cat", "%?\\s*CAT\\s*:?"],
    ["plazo", "Plazo\\s*del\\s*Cr[e\u00e9]dito\\s*:?"],
    ["ret", "%?\\s*Retenci[o\u00f3]n\\s*:?"],
    ["banco", "Instituci[o\u00f3]n\\s*Bancaria\\s*Acreditant\\w*\\s*:?"],
    ["fLiq", "F\\.?\\s*Liquidaci[o\u00f3]n\\s*:?"],
  ];
  const CAP_RE = new RegExp(CAP_LABELS.map(l => "(" + l[1] + ")").join("|"), "gi");
  function capDate(raw) {
    const s = String(raw || "");
    let m = s.match(/(\d{1,2})\s*\/\s*(\d{1,2})\s*\/\s*(\d{4})/);
    if (m) return { v: `${m[1].padStart(2, "0")}/${m[2].padStart(2, "0")}/${m[3]}`, fixed: false };
    const d = (s.match(/\d[\d\/]{6,10}\d/) || [""])[0].replace(/\D/g, "");
    if (d.length === 8) return { v: `${d.slice(0, 2)}/${d.slice(2, 4)}/${d.slice(4)}`, fixed: true };
    return null;
  }
  function capNum(raw, dec) {
    const m = String(raw || "").match(/\d[\d.,]*/); if (!m) return null;
    let t = m[0].replace(/,/g, "."), fixed = false;
    if (dec && !t.includes(".") && t.length >= 3) { t = t.slice(0, -2) + "." + t.slice(-2); fixed = true; }
    const v = parseFloat(t); return isFinite(v) ? { v, fixed, raw: m[0] } : null;
  }
  E.extractCaptura = function (doc) {
    const f = {}, x = { vals: {}, raw: {} }, p = doc.pages[0]; const LINES = buildLines(p);
    const conf = p.conf ?? 80, u = lab => `campo '${lab}'`;
    const LABTXT = { fFinsus: "F. Contrato Finsus", fFinsusLetra: "F. Contrato Finsus (Letra)", numProc: "# Contrato Procesador", razonProc: "Raz\u00f3n Social Procesador", nombreProc: "Nombre Contrato Procesador", fProc: "F. Contrato Procesador", ocupacion: "Ocupaci\u00f3n", destino: "Destino de los Recursos", tasaOrd: "Tasa Inter\u00e9s Anual Ordinaria", tasaOrdLetra: "Tasa Ordinaria (Letra)", tasaMor: "Tasa Inter\u00e9s Moratoria", tasaMorLetra: "Tasa Moratoria (Letra)", cuentaAcred: "Cuenta del Acreditante", clabeAcred: "Cuenta Clabe Acreditante", lugar: "Lugar de firma Contrato Finsus", tipoSoc: "Tipo Sociedad", pInv: "% Comisi\u00f3n Gastos de Investigaci\u00f3n s/IVA", pDisp: "% Comisi\u00f3n x Disposici\u00f3n s/IVA", pSDGC: "% Comisi\u00f3n por Gesti\u00f3n de Cobranza s/IVA", cat: "% CAT", plazo: "Plazo del Cr\u00e9dito", ret: "% Retenci\u00f3n", banco: "Instituci\u00f3n Bancaria Acreditante", fLiq: "F. Liquidaci\u00f3n" };
    x.labels = LABTXT;
    for (const ln of LINES) {
      const t = ln.text; const ms = [...t.matchAll(CAP_RE)];
      ms.forEach((m, i) => {
        const gi = m.slice(1).findIndex(g => g != null); if (gi < 0) return;
        const key = CAP_LABELS[gi][0]; const end = m.index + m[0].length; const nx = ms[i + 1] ? ms[i + 1].index : t.length;
        const val = t.slice(end, nx).replace(/^[\s:;.|\u2014-]+|[\s|\u2014-]+$/g, "").trim();
        if (!(key in x.raw)) x.raw[key] = val;
      });
    }
    const R = x.raw, low = 45;
    const D = (k, fk) => { const d = capDate(R[k]); if (d) { F(f, fk, d.v, u(LABTXT[k]), { ocr: true, conf: d.fixed ? low : conf }); x.vals[k] = d.v; } };
    D("fFinsus", "fechaFirma"); D("fProc", "fechaProcesador"); D("fLiq", "vencimiento");
    if (R.fFinsusLetra && /\d{4}/.test(R.fFinsusLetra)) F(f, "fechaFirma", R.fFinsusLetra.replace(/(\d{4})\s+\S{1,2}$/, "$1"), u(LABTXT.fFinsusLetra), { ocr: true, conf });
    if (R.razonProc) F(f, "procesador", R.razonProc.replace(/\b5\.\s?A\./g, "S.A.").replace(/C\.\s?v\./g, "C.V."), u(LABTXT.razonProc), { ocr: true, conf });
    const Pc = (k, fk, dec, maxv) => { const n = capNum(R[k], dec); if (!n) return; const sus = n.fixed || (maxv && n.v > maxv) || (k === "ret" && n.v < 5); x.vals[k] = n.v; if (fk) F(f, fk, String(n.v), u(LABTXT[k]), { ocr: true, conf: sus ? low : conf }); if (sus) (x.sospechosos = x.sospechosos || []).push(LABTXT[k] + ": '" + n.raw + "'"); };
    Pc("tasaOrd", "tasaOrd", true, 100); Pc("tasaMor", "tasaMor", true, 150); Pc("cat", "cat", true, 300); Pc("ret", "retencion", false, 100);
    Pc("pInv", null, true, 10); Pc("pDisp", null, true, 10); Pc("pSDGC", null, true, 10);
    const pl = capNum(R.plazo, false); if (pl) { F(f, "plazo", String(pl.v), u(LABTXT.plazo), { ocr: true, conf }); x.vals.plazo = pl.v; }
    const np = capNum(R.numProc, false); if (np) x.vals.numProc = np.raw;
    const cl = (R.clabeAcred || "").replace(/\D/g, ""); if (cl) x.vals.clabeAcred = cl;
    x.ocupacion = R.ocupacion != null ? R.ocupacion : undefined; x.tipoSoc = R.tipoSoc != null ? R.tipoSoc : undefined;
    if (R.ocupacion) F(f, "ocupacion", R.ocupacion, u(LABTXT.ocupacion), { ocr: true, conf });
    x.banco = R.banco || null; x.destino = R.destino || null; x.lugar = R.lugar || null; x.nombreProc = R.nombreProc || null;
    x.encontrados = Object.keys(R).length;
    x.texto = p.flat;
    return { fields: f, x };
  };

  /* ---- Word (.docx): convierte document.xml en "p\u00e1ginas" de renglones para usar los mismos extractores ---- */
  E.docxToPages = function (xml, perPage) {
    perPage = perPage || 60;
    const dec = t => t.replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, "&");
    const ptext = p => dec((p.match(/<w:t(?:\s[^>]*)?>[^<]*<\/w:t>|<w:tab\/>|<w:br\/>/g) || []).map(t => t.startsWith("<w:t") ? t.replace(/<[^>]+>/g, "") : " ").join(""));
    const body = (xml.match(/<w:body>([\s\S]*)<\/w:body>/) || [, xml])[1];
    const lines = []; // cada rengl\u00f3n = lista de celdas
    const re = /<w:tbl>[\s\S]*?<\/w:tbl>|<w:p[ >][\s\S]*?<\/w:p>|<w:p\/>/g; let m;
    while ((m = re.exec(body))) {
      const blk = m[0];
      if (blk.startsWith("<w:tbl")) {
        for (const tr of blk.match(/<w:tr[ >][\s\S]*?<\/w:tr>/g) || []) {
          const cells = (tr.match(/<w:tc>[\s\S]*?<\/w:tc>/g) || []).map(tc => (tc.match(/<w:p[ >][\s\S]*?<\/w:p>/g) || []).map(ptext).map(t => t.trim()).filter(Boolean).join(" "));
          if (cells.some(Boolean)) lines.push(cells);
        }
      } else { const t = ptext(blk).replace(/\s+/g, " ").trim(); if (t) lines.push([t]); }
    }
    const pages = [];
    for (let i = 0; i < lines.length; i += perPage) {
      const items = [];
      lines.slice(i, i + perPage).forEach((cells, k) => {
        const y = 40 + k * 14;
        cells.forEach((c, j) => { if (!c) return; const x0 = 40 + j * (cells.length > 1 ? 520 / cells.length : 0); items.push({ s: c, x0, x1: x0 + Math.min(c.length * 5, cells.length > 1 ? 520 / cells.length - 6 : 520), y, h: 10 }); });
      });
      pages.push({ no: pages.length + 1, source: "pdf", items });
    }
    return pages.length ? pages : [{ no: 1, source: "pdf", items: [] }];
  };
  E.extractTA = function (doc) {
    const f = {}, x = {}; const ws = doc.wb.Sheets[doc.wb.SheetNames[0]]; const sh = doc.wb.SheetNames[0];
    const A = i => String(cv(ws, "A" + i) || "");
    for (let i = 1; i < 15; i++) {
      const t = A(i);
      if (/^Cliente/i.test(t)) { F(f, "idcliente", grab(t, /:\s*([\d-]+)/), `${sh}!A${i}`); F(f, "razon", grab(t, /,\s*(.+)$/), `${sh}!A${i}`); }
      if (/^Referencia/i.test(t)) F(f, "referencia", grab(t, /:\s*(\S+)/), `${sh}!A${i}`);
      if (/^Credito/i.test(t)) F(f, "credito", grab(t, /:\s*(\S+)/), `${sh}!A${i}`);
      if (/^Monto/i.test(t)) { F(f, "monto", money(grab(t, /:\s*([\d,.]+)/)), `${sh}!A${i}`); F(f, "plazo", +grab(t, /Plazo:\s*(\d+)/), `${sh}!A${i}`); }
      if (/^Tasas/i.test(t)) { F(f, "tasaOrd", pct(grab(t, /Ord:\s*([\d.]+)/)), `${sh}!A${i}`); F(f, "tasaMor", pct(grab(t, /Mor:\s*([\d.]+)/)), `${sh}!A${i}`); }
      if (/^CAT/i.test(t)) F(f, "cat", pct(grab(t, /([\d.]+)/)), `${sh}!A${i}`);
      if (/\d{4} - \(/.test(t) && !x.fecha) { x.fecha = t; F(f, "fechaFirma", grab(t, /^(.+?\d{4})/), `${sh}!A${i}`); }
    }
    const rows = []; let last = 0;
    for (let i = 10; i < 200; i++) { const a = A(i); if (/^0\/\d+$/.test(a)) { rows.push({ n: +a.split("/")[1], fecha: parseDate(cv(ws, "B" + i)), dias: +cv(ws, "C" + i), cap: +cv(ws, "D" + i), int: +cv(ws, "E" + i), iva: +cv(ws, "F" + i), com: +cv(ws, "G" + i), ivacom: +cv(ws, "H" + i), pago: r2(+cv(ws, "I" + i)), saldo: +cv(ws, "J" + i), row: i }); last = i; } }
    x.rows = rows;
    if (rows.length) {
      F(f, "primerPago", rows[0].fecha, `${sh}!B${rows[0].row}`); F(f, "vencimiento", rows[rows.length - 1].fecha, `${sh}!B${last}`);
      F(f, "pagoMensual", modeOf(rows.map(r => r.pago)), `${sh}!I${rows[0].row}:I${last}`); F(f, "numPagos", rows.length, `${sh}!A${rows[0].row}:A${last}`);
      for (let i = last + 1; i < last + 5; i++) { const c = cell(ws, "I" + i); if (c && typeof c.v === "number" && c.v > 1000) { F(f, "total", r2(c.v), `${sh}!I${i}`); x.totalesFijos = ["D", "E"].filter(col => cell(ws, col + i) && !cell(ws, col + i).f).map(col => col + i); break; } }
    }
    return { fields: f, x };
  };

  E.extractCotizador = function (doc) {
    const f = {}, x = {}; const S = n => doc.wb.Sheets[doc.wb.SheetNames.find(s => norm(s) === norm(n))];
    const inp = S("Inputs"), c18 = S("Cotizador 18m"), pdf = S("PDF Oferta"), o12 = S("Oferta 12m");
    if (inp) {
      x.ventas = []; for (let i = 3; i <= 30; i++) { const d = cv(inp, "B" + i), v = cv(inp, "C" + i); if (d == null || v == null || typeof v !== "number") continue; const nota = cv(inp, "D" + i); x.ventas.push({ mes: parseDate(d instanceof Date ? d : (typeof d === "number" ? new Date(Math.round((d - 25569) * 864e5)) : d)), v, ubic: "Inputs!C" + i, nota: typeof nota === "string" ? nota : null, formula: cell(inp, "C" + i)?.f || null }); }
      F(f, "plazo", cv(inp, "E3"), "Inputs!E3"); F(f, "fechaFirma", cv(inp, "G3") instanceof Date ? cv(inp, "G3") : (typeof cv(inp, "G3") === "number" ? new Date(Math.round((cv(inp, "G3") - 25569) * 864e5)) : cv(inp, "G3")), "Inputs!G3");
      F(f, "razon", cv(inp, "F7"), "Inputs!F7"); F(f, "idExpediente", String(cv(inp, "F8") ?? ""), "Inputs!F8");
    }
    if (c18) {
      x.multFormula = !!cell(c18, "D4")?.f; x.sdgcFormula = cell(c18, "D6")?.f || null;
      // Bloque "Cotizador s/IVA" (normalmente columnas M:N): misma estructura que el bloque con IVA, con la tasa SDGC sin IVA ya redondeada
      {
        const colName = c => { let s = ""; c++; while (c > 0) { const m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = Math.floor((c - 1) / 26); } return s; };
        let found = null;
        for (const a in c18) { if (a[0] === "!") continue; const v = c18[a]?.v; if (typeof v === "string" && /Cotizador\s*s\/\s*IVA/i.test(v)) { const m = a.match(/^([A-Z]+)(\d+)$/); if (m) { let c = 0; for (const ch of m[1]) c = c * 26 + (ch.charCodeAt(0) - 64); found = { r: +m[2], c: c - 1 }; break; } } }
        if (found) {
          const col = colName(found.c + 1), row = n => col + (found.r + n);
          const lbl = n => String(cv(c18, colName(found.c) + (found.r + n)) || "");
          const pick = re => { for (let n = 1; n <= 14; n++) if (re.test(lbl(n))) return { v: cv(c18, row(n)), ubic: "'Cotizador 18m'!" + row(n), f: cell(c18, row(n))?.f || null }; return null; };
          x.sinIVA = { rate: pick(/Comisi[\u00f3o]n por SGDC/i), monto: pick(/Comisi[\u00f3o]n SGDC Monto/i), apertura: pick(/Comisi[\u00f3o]n Apertura/i), ret: pick(/Retenci/i), cat: pick(/CAT/i), vp: pick(/^VP/i) };
        }
      }
      x.intRows = []; for (let r = 21; r <= 40; r++) { const n = cv(c18, "B" + r), d = cv(c18, "D" + r); if (typeof n === "number" && n >= 1 && typeof d === "number") x.intRows.push(r2(d)); }
      x.promedio = cv(c18, "D3"); x.mult = cv(c18, "D4"); x.sdgcRate = cv(c18, "D6"); x.soloInteres = cv(c18, "D10"); x.apertura = cv(c18, "D11"); x.sdgcCIVA = cv(c18, "D12"); x.neto = cv(c18, "H18") != null ? Math.abs(cv(c18, "H18")) : null; x.plazoBase = cv(c18, "D9");
      F(f, "monto", cv(c18, "D8"), "'Cotizador 18m'!D8"); F(f, "tasaOrd", pct(cv(c18, "D5")), "'Cotizador 18m'!D5");
      F(f, "total", r2(cv(c18, "D14")), "'Cotizador 18m'!D14"); F(f, "pagoMensual", r2(cv(c18, "D13")), "'Cotizador 18m'!D13 (pago de la oferta)");
      F(f, "retencion", pct(cv(c18, "D15")), "'Cotizador 18m'!D15"); F(f, "cat", pct(cv(c18, "R9")), "'Cotizador 18m'!R9");
      const a21 = cv(c18, "A21"); if (a21) F(f, "primerPago", typeof a21 === "number" ? new Date(Math.round((a21 - 25569) * 864e5)) : a21, "'Cotizador 18m'!A21");
      for (const r of [36, 35, 37, 34]) { const a = cv(c18, "A" + r); if (a) { F(f, "vencimiento", typeof a === "number" ? new Date(Math.round((a - 25569) * 864e5)) : a, "'Cotizador 18m'!A" + r); break; } }
      if (x.sdgcCIVA) F(f, "comSDGC", r2(x.sdgcCIVA / 1.16), "'Cotizador 18m'!D12 \u00f7 1.16");
    }
    if (pdf) { x.plazoOferta = { v: cv(pdf, "C19"), ubic: "'PDF Oferta'!C19" }; x.fechaToday = /TODAY/i.test(cell(pdf, "F2")?.f || ""); }
    if (o12) x.alt12 = { monto: cv(o12, "D2"), tasa: cv(o12, "D4"), pago: cv(o12, "D9") };
    return { fields: f, x };
  };

  /* ============ definici\u00f3n de campos a cruzar ============ */
  const FIELDS = [
    ["razon", "Identificadores", "Raz\u00f3n social del acreditado", "razon"], ["rfc", "Identificadores", "RFC del acreditado", "id"], ["credito", "Identificadores", "N\u00famero de cr\u00e9dito", "id"],
    ["referencia", "Identificadores", "Referencia de pago", "id"], ["idcliente", "Identificadores", "N\u00famero de cliente Finsus", "id"], ["fideicomiso", "Identificadores", "N\u00famero de fideicomiso", "id"],
    ["procesador", "Identificadores", "Procesador de pagos", "razon"], ["folio", "Identificadores", "Folio mercantil", "id"], ["curp", "Identificadores", "CURP de la acreditada", "id"], ["idPF", "Identificadores", "Identificaci\u00f3n oficial de la acreditada", "id"],
    ["fechaFirma", "Fechas", "Fecha de firma", "date"], ["fechaInicio", "Fechas", "Fecha de inicio del contrato", "date"], ["fechaConst", "Fechas", "Fecha de constituci\u00f3n", "date"], ["fechaPoder", "Fechas", "Fecha de la escritura del poder", "date"],
    ["primerPago", "Fechas", "Primer pago", "date"], ["vencimiento", "Fechas", "Vencimiento / \u00faltimo pago", "date"], ["fechaProcesador", "Fechas", "Fecha del contrato con el procesador", "date"], ["fechaFideicomiso", "Fechas", "Constituci\u00f3n del fideicomiso", "date"],
    ["monto", "Montos y Tasas", "Monto del cr\u00e9dito", "money"], ["total", "Montos y Tasas", "Monto total a pagar", "money"], ["sumaPrincipal", "Montos y Tasas", "Suma principal del pagar\u00e9", "money"],
    ["pagoMensual", "Montos y Tasas", "Pago mensual", "money"], ["pagoDomiciliado", "Montos y Tasas", "Cargo mensual domiciliado", "money"], ["pagoMinimo", "Montos y Tasas", "Pago mensual m\u00ednimo garantizado (convenio)", "money"],
    ["plazo", "Montos y Tasas", "Plazo (meses)", "num"], ["numPagos", "Montos y Tasas", "N\u00famero de pagos", "num"], ["tasaOrd", "Montos y Tasas", "Tasa ordinaria (%)", "pct"], ["tasaMor", "Montos y Tasas", "Tasa moratoria (%)", "pct"],
    ["cat", "Montos y Tasas", "CAT (%)", "pct1"], ["retencion", "Montos y Tasas", "% de retenci\u00f3n diaria", "pct"],
    ["comInv", "Montos y Tasas", "Comisi\u00f3n por investigaci\u00f3n", "money"], ["comDisp", "Montos y Tasas", "Comisi\u00f3n por disposici\u00f3n", "money"], ["comSDGC", "Montos y Tasas", "Comisi\u00f3n por gesti\u00f3n de cobranza (s/IVA)", "money"],
    ["representante", "Personas", "Representante legal del acreditado", "name"], ["obligado", "Personas", "Obligado solidario / aval", "name"], ["firmanteProcesador", "Personas", "Firmante del contrato con el procesador", "name"],
    ["rfcObligado", "Personas", "RFC del obligado solidario", "id"], ["curpObligado", "Personas", "CURP del obligado solidario", "id"], ["idObligado", "Personas", "Identificaci\u00f3n del obligado solidario", "id"], ["nacObligado", "Personas", "Fecha de nacimiento del obligado", "date"],
    ["telefono", "Personas", "Tel\u00e9fono", "id"], ["emailObligado", "Personas", "Correo del representante / obligado", "email"], ["emailCliente", "Personas", "Correo de la acreditada", "email"], ["emailNotif", "Personas", "Correo de notificaciones", "email"], ["fechaNac", "Personas", "Fecha de nacimiento de la acreditada", "date"], ["lugarNac", "Personas", "Lugar de nacimiento", "name"], ["ocupacion", "Personas", "Ocupaci\u00f3n", "name"],
    ["escrituraConst", "Personas", "Escritura constitutiva", "num"], ["escrituraPoder", "Personas", "Escritura del poder", "num"], ["fedatarioConst", "Personas", "Fedatario", "name"],
    ["domCalle", "Domicilios", "Domicilio fiscal \u2014 calle y n\u00famero", "addr"], ["domColonia", "Domicilios", "Domicilio fiscal \u2014 colonia", "name"], ["domMunicipio", "Domicilios", "Domicilio fiscal \u2014 municipio", "name"], ["domCP", "Domicilios", "Domicilio fiscal \u2014 C.P.", "id"],
    ["avalCalle", "Domicilios", "Domicilio del aval \u2014 calle y n\u00famero", "addr"], ["avalCP", "Domicilios", "Domicilio del aval \u2014 C.P.", "id"],
    ["sucursales", "Sucursales", "N\u00famero de sucursales", "num"], ["afiliaciones", "Sucursales", "Afiliaciones del procesador", "num"],
  ];
  E.FIELDS = FIELDS;
  function similar(a, b) { a = String(a); b = String(b); if (!a.length || !b.length) return 0; const m = a.length, n = b.length; let prev = Array.from({ length: n + 1 }, (_, j) => j); for (let i = 1; i <= m; i++) { const cur = [i]; for (let j = 1; j <= n; j++) cur[j] = Math.min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1)); prev = cur; } return 1 - prev[n] / Math.max(m, n); }
  function canon(type, v) {
    if (v == null) return null;
    switch (type) {
      case "money": return money(v);
      case "pct": case "pct1": return pct(v);
      case "num": return Number(String(v).replace(/[^\d.]/g, ""));
      case "date": return parseDate(v);
      case "id": return squash(v).toUpperCase();
      case "email": return String(v).toLowerCase().replace(/\s/g, "");
      case "razon": return regimen(v).replace(/ /g, "");
      default: return squash(v);
    }
  }
  const display = (type, v) => type === "money" ? fmtMoney(money(v)) : type === "date" ? (parseDate(v) ? fmtDate(parseDate(v)) + (typeof v === "string" && !/^\d{4}-/.test(v) ? "  (" + v + ")" : "") : String(v)) : (v instanceof Date ? fmtDate(parseDate(v)) : String(v));

  function compare(type, entries) {
    const vals = entries.map(e => canon(type, e.v));
    if (type === "money") {
      const mn = Math.min(...vals), mx = Math.max(...vals);
      if (mx - mn < 0.005) return "S\u00ed";
      if (mx - mn < 1) return "Diferencia menor";
      return "No";
    }
    if (type === "pct" || type === "pct1") { const mn = Math.min(...vals), mx = Math.max(...vals); if (mx - mn < (type === "pct1" ? 0.011 : 0.0001)) return "S\u00ed"; return "No"; }
    if (type === "addr") {
      const nums = entries.map(e => [...new Set((String(e.v).match(/\d+/g) || []).filter(n => n.length !== 5))].sort().join(","));
      if (vals.every(v => v === vals[0])) return entries.every(e => e.v === entries[0].v) ? "S\u00ed" : "Diferencia menor";
      return nums.every(n => n === nums[0]) ? "Diferencia menor" : "No";
    }
    if (type === "razon") {
      if (vals.every(v => v === vals[0])) return entries.every(e => String(e.v).trim() === String(entries[0].v).trim()) ? "S\u00ed" : "Diferencia menor";
      const tk = e => regimen(e.v).split(" ").filter(Boolean).sort().join(" ");
      if (entries.every(e => tk(e) === tk(entries[0]))) return "Diferencia menor";
      return "No";
    }
    if (type === "name") { const L = vals.slice().sort((a, b) => b.length - a.length)[0]; if (vals.every(v => v && L.startsWith(v)) && !vals.every(v => v === vals[0])) return "Diferencia menor"; }
    if (vals.every(v => v === vals[0])) {
      if (type === "date" || type === "num") return "S\u00ed";
      return entries.every(e => String(e.v).trim() === String(entries[0].v).trim()) ? "S\u00ed" : "Diferencia menor";
    }
    return "No";
  }

  /* ============ cruce ============ */
  E.cross = function (docs, opts = {}) {
    const report = { resumen: [], discrepancias: [], dudas: [], verificaciones: [], cuentas: [], documentos: [], faltantes: [], ventas: [] };
    const byKind = k => docs.filter(d => d.kind === k);
    const one = k => byKind(k)[0];
    const LAB = { contrato: "Contrato", pagare: "Pagar\u00e9", convenio: "Convenio de adhesi\u00f3n", kyc: "KYC", ta: "Tabla de amortizaci\u00f3n (Excel)", cotizador: "Cotizador (Excel)", csf: "Constancia de Situaci\u00f3n Fiscal", captura: "Captura del sistema (Datos Captura)", otro: "No reconocido", excel: "Excel no reconocido" };
    docs.forEach(d => { d.label = LAB[d.kind] || d.kind; report.documentos.push({ nombre: d.name, tipo: d.label, paginas: d.pages ? d.pages.length + " pp." : (d.wb ? d.wb.SheetNames.length + " hojas" : ""), calidad: d.quality || "" }); });
    const EXPECT = [["contrato", "Contrato de apertura de cr\u00e9dito"], ["pagare", "Pagar\u00e9"], ["convenio", "Convenio de adhesi\u00f3n al fideicomiso"], ["kyc", "KYC"], ["ta", "Tabla de amortizaci\u00f3n"], ["cotizador", "Cotizador"], ["csf", "Constancia de Situaci\u00f3n Fiscal"]];
    EXPECT.forEach(([k, n]) => { if (!one(k)) report.faltantes.push(n); });
    report.faltantes.push("INE del representante y obligado solidario", "Acta constitutiva y poder", "Comprobante de domicilio", "Car\u00e1tulas o estados de cuenta de las cuentas listadas");

    const get = (k, key) => one(k)?.ex?.fields?.[key];
    const X = k => one(k)?.ex?.x || {};
    const C = X("contrato"), PG = X("pagare"), CV = X("convenio"), K = X("kyc"), T = X("ta"), Q = X("cotizador");
    const covered = new Set();
    const disc = (prio, campo, a, b, why, fieldKey) => { report.discrepancias.push({ prio, campo, a, b, why }); if (fieldKey) covered.add(fieldKey); };
    let nTop = 0; const duda = (doc, ubic, ve, d, preg, top) => { const o = { doc, ubic, ve, duda: d, preg }; if (top) report.dudas.splice(nTop++, 0, o); else report.dudas.push(o); };
    const ver = (desc, a, b, note = "") => { if (a == null || b == null || !isFinite(a) || !isFinite(b)) return; report.verificaciones.push({ desc, a: r2(a), b: r2(b), note }); };

    // campos cruzados
    const fieldStatus = {};
    for (const [key, cat, label, type] of FIELDS) {
      const entries = [];
      docs.forEach(d => { const e = d.ex?.fields?.[key]; if (e && e.v != null && e.v !== "") { entries.push({ ...e, more: undefined, doc: d.label }); (e.more || []).forEach(m => entries.push({ ...m, doc: d.label })); } });
      if (key === "fechaFirma" && C.fechaBloqueFirma) entries.push({ v: C.fechaBloqueFirma.v, ubic: C.fechaBloqueFirma.ubic, doc: "Contrato" });
      if (key === "fechaFirma" && C.fechaCierre?.v) entries.push({ v: C.fechaCierre.v, ubic: C.fechaCierre.ubic, doc: "Contrato" });
      if (key === "fechaFirma" && CV.fechaCierre?.v) entries.push({ v: CV.fechaCierre.v, ubic: CV.fechaCierre.ubic, doc: "Convenio de adhesi\u00f3n" });
      if (!entries.length) continue;
      let st;
      const lowConf = entries.filter(e => e.ilegible || (e.ocr && e.conf != null && e.conf < 60));
      if (entries.length === 1) st = "Solo en un documento";
      else st = compare(type, entries);
      if (lowConf.length && st !== "S\u00ed") st = st === "Solo en un documento" ? st : "No legible";
      if (st === "No" && entries.some(e => e.ocr)) {
        const nonOcr = entries.filter(e => !e.ocr).map(e => String(canon(type, e.v)));
        const ocrVals = entries.filter(e => e.ocr).map(e => String(canon(type, e.v)));
        const near = (v, w) => { if (similar(v, w) >= 0.8) return true; if (v.length !== w.length) return false; let d = 0; for (let i = 0; i < v.length; i++) if (v[i] !== w[i]) d++; return d <= (v.length >= 10 ? 2 : 1); };
        if (nonOcr.length && nonOcr.every(v => v === nonOcr[0]) && ocrVals.every(v => v !== nonOcr[0] && near(v, nonOcr[0]))) st = "No legible";
      }
      let obs = "";
      if (st === "No" || st === "Diferencia menor" || st === "No legible") {
        const groups = {}; entries.forEach(e => { const k = String(canon(type, e.v)); (groups[k] = groups[k] || []).push(e); });
        obs = Object.values(groups).map(g => display(type, g[0].v) + " \u2192 " + [...new Set(g.map(e => e.doc))].join(", ")).join(" \u00b7 ");
        if (st === "Diferencia menor" && type === "money") obs = "Diferencia de centavos (redondeo). " + obs;
        if (st === "Diferencia menor" && type !== "money") obs = "Solo difiere el formato. " + obs;
        if (st === "No legible") obs = "Al menos un valor sali\u00f3 de OCR o texto da\u00f1ado: confirma contra el original. " + obs;
      }
      if (entries.some(e => e.ocr)) obs = (obs ? obs + " " : "") + "(" + [...new Set(entries.filter(e => e.ocr).map(e => e.doc))].join(", ") + " le\u00eddo con OCR)";
      fieldStatus[key] = { st, entries, label, cat, type };
      entries.forEach(e => report.resumen.push({ cat, campo: label, valor: display(type, e.v), doc: e.doc, ubic: e.ubic, st, obs }));
    }
    const fs = k => fieldStatus[k];

    /* ---- reglas espec\u00edficas ---- */
    const rep = get("contrato", "representante")?.v || get("kyc", "representante")?.v;
    const fp = fs("firmanteProcesador");
    if (fp && rep) {
      fp.entries.forEach(e => {
        if (regimen(e.v) !== regimen(rep)) {
          const acc = (K.accionistas || []).map(a => a.nombre);
          const first = norm(e.v).split(" ")[0]; const parecido = acc.find(a => norm(a).split(" ")[0] === first && regimen(a) !== regimen(e.v));
          disc(e.ilegible ? "Media" : "Alta", "Firmante del contrato con el procesador", `${e.doc}, ${e.ubic}: firma "${e.v}"${e.ilegible ? " (texto de mala calidad; confirmar en el original)" : ""}.`, `Contrato, Anexo de Datos Generales ref. 3.1: el representante es "${rep}".${acc.length ? " Accionistas en KYC: " + K.accionistas.map(a => a.nombre + " " + a.pct + "%").join(", ") + "." : ""}${parecido ? " El nombre se parece a " + parecido + "." : ""}`, "La cesi\u00f3n de la cobranza al fideicomiso depende de ese contrato. Si quien firm\u00f3 no tiene poder, su validez queda en duda.", "firmanteProcesador");
          duda(e.doc, e.ubic, `Firma "${e.v}".`, "No hay poder ni identificaci\u00f3n de esa persona en el expediente.", "\u00bfQui\u00e9n firm\u00f3 el contrato con el procesador y con qu\u00e9 poder? \u00bfDebe volver a firmar el representante?");
        }
      });
    }
    const sp = get("pagare", "sumaPrincipal")?.v;
    // Regla acordada (06/11/2026): el 'Monto del Pagar\u00e9' del convenio (Anexo B) es la SUMA DE LOS PAGOS (columna H del cotizador,
    // = monto total a pagar de la car\u00e1tula/TA); la suma principal del pagar\u00e9 es D8 + D12 del cotizador. No se comparan entre s\u00ed.
    if (CV.montoPagare?.v != null) {
      const totalC = get("contrato", "total")?.v ?? get("ta", "total")?.v;
      const totCot = get("cotizador", "total")?.v;
      if (totalC != null) {
        ver("'Monto del Pagar\u00e9' del convenio = monto total a pagar (car\u00e1tula / TA)", CV.montoPagare.v, totalC, "Regla acordada");
        if (Math.abs(CV.montoPagare.v - totalC) >= 1) disc("Alta", "'Monto del Pagar\u00e9' del convenio vs. suma de pagos", `Convenio, ${CV.montoPagare.ubic}: ${fmtMoney(CV.montoPagare.v)}.`, `Monto total a pagar (suma de los pagos): ${fmtMoney(totalC)}${totCot != null ? "; cotizador (columna H) " + fmtMoney(totCot) : ""}.`, "El Anexo B del convenio debe ser la suma de los pagos.");
      }
      if (totCot != null) ver("'Monto del Pagar\u00e9' del convenio = suma de pagos del cotizador (columna H)", CV.montoPagare.v, totCot, "Regla acordada; centavos = redondeo del cotizador");
    }
    const plazoC = get("contrato", "plazo")?.v;
    // Regla acordada (05/11/2026): el cotizador maneja amortizaciones + periodo 's\u00f3lo inter\u00e9s' (16 + 2 = 18); eso es correcto.
    // Solo se reporta si las amortizaciones del cotizador no coinciden con el plazo del contrato.
    if (Q.plazoOferta?.v != null && plazoC != null && +(Q.plazoBase ?? Q.plazoOferta.v) !== +plazoC) {
      disc("Media", "Plazo del cr\u00e9dito", `Cotizador ${Q.plazoOferta.ubic}: ${Q.plazoOferta.v} meses${Q.soloInteres ? " (" + Q.plazoBase + " + " + Q.soloInteres + " 's\u00f3lo inter\u00e9s')" : ""}.`, `Car\u00e1tula del contrato: ${plazoC} meses; pagar\u00e9 y TA: ${get("pagare", "numPagos")?.v ?? get("ta", "numPagos")?.v ?? "\u2014"} pagos.`, "La oferta presentada no coincide con lo contratado; la comisi\u00f3n de cobranza se calcula sobre el plazo del cotizador.");
    }
    const fc = parseDate(get("contrato", "fechaConst")?.v || get("kyc", "fechaConst")?.v);
    const firma = parseDate(get("contrato", "fechaFirma")?.v);
    if (Q.ventas?.length) {
      report.ventas = Q.ventas.map(v => ({ mes: v.mes, v: v.v }));
      if (fc) {
        const pre = Q.ventas.filter(v => v.mes && dayDiff(v.mes.slice(0, 7) + "-01", fc) > 31 - 0 && new Date(v.mes).getTime() < new Date(fc.slice(0, 7) + "-01").getTime());
        if (pre.length) disc("Alta", "Ventas del cotizador anteriores a la constituci\u00f3n", `Cotizador Inputs: ${pre.length} meses de ventas antes de la constituci\u00f3n (${pre.map(v => v.mes.slice(0, 7)).join(", ")}), suma ${fmtMoney(pre.reduce((a, v) => a + v.v, 0))}.`, `Fecha de constituci\u00f3n: ${fmtDate(fc)}.`, "El monto del cr\u00e9dito se calcul\u00f3 con ventas que no pueden ser de esta persona moral.");
      }
      // Ventas que no son datos capturados sino f\u00f3rmulas (p. ej. =AVERAGE(C3:C4)*0.92): el promedio que define el monto es estimado
      const est = Q.ventas.filter(v => v.formula);
      if (est.length) disc(est.length >= Q.ventas.length / 2 ? "Alta" : "Media", "Ventas del cotizador calculadas con f\u00f3rmula", `Cotizador Inputs: ${est.length} de ${Q.ventas.length} meses son f\u00f3rmulas (${est.slice(0, 3).map(v => v.ubic.replace("Inputs!", "") + " =" + v.formula).join("; ")}${est.length > 3 ? "; \u2026" : ""}).`, `Solo ${Q.ventas.length - est.length} meses traen ventas capturadas (${Q.ventas.filter(v => !v.formula).map(v => (v.mes || "").slice(0, 7)).join(", ")}).`, "El promedio que define el monto del cr\u00e9dito y la comisi\u00f3n SDGC se basa en ventas estimadas, no en ventas reales.");
      const zeros = Q.ventas.filter(v => v.v === 0); if (zeros.length) disc("Alta", "Meses sin ventas en el cotizador", `Cotizador Inputs: ${zeros.map(v => v.mes.slice(0, 7)).join(", ")} = $0.`, "El promedio que define el monto incluye esos meses.", "Puede indicar cambio de procesador, falta de informaci\u00f3n o paro de operaci\u00f3n.");
      const lastMes = Q.ventas[Q.ventas.length - 1]?.mes; if (lastMes && firma && dayDiff(lastMes, firma) > 75) duda("Cotizador", "Inputs", `El mes m\u00e1s reciente de ventas es ${lastMes.slice(0, 7)}.`, "Faltan meses recientes antes de la firma.", "\u00bfHay ventas m\u00e1s recientes para confirmar la tendencia?");
      const pago = get("ta", "pagoMensual")?.v || get("pagare", "pagoMensual")?.v, ret = (get("convenio", "retencion")?.v ?? get("cotizador", "retencion")?.v);
      if (pago && ret) { const low = Q.ventas.filter(v => v.v * ret / 100 < pago); report.ventasInfo = { pago, ret, min: pago / (ret / 100), low: low.map(v => v.mes.slice(0, 7)) }; if (low.length >= 2) duda("Cotizador", "Inputs", `En ${low.length} de ${Q.ventas.length} meses el ${ret}% de las ventas no alcanza la mensualidad de ${fmtMoney(pago)}.`, "La fuente de pago podr\u00eda no cubrir la mensualidad en meses bajos.", "\u00bfSe considera suficiente la cobertura de la retenci\u00f3n?"); }
      const suma = Q.ventas.reduce((a, v) => a + v.v, 0), prom = suma / Q.ventas.length;
      if (Q.mult && get("cotizador", "monto")?.v) ver(`Monto = ${Q.mult} \u00d7 promedio de ventas`, get("cotizador", "monto").v, Math.round(Q.mult * prom), "Cotizador");
      if (K.ventasAnuales) { const va = K.ventasAnuales; if (suma < va.min || suma > va.max) disc("Media", "Ventas anuales declaradas en KYC", `${va.ubic}: '${va.label}'.`, `Cotizador: ventas con tarjeta de ${Q.ventas.length} meses = ${fmtMoney(suma)}.${K.opsMensuales?.v ? " KYC: operaciones mensuales " + fmtMoney(K.opsMensuales.v) + "." : ""}`, "La casilla del KYC no cuadra con las ventas usadas para el cr\u00e9dito (puede haber ventas por otros medios: confirmar)."); }
      if (K.opsMensuales?.v && Math.abs(K.opsMensuales.v - prom) / prom > 0.3) duda("KYC", K.opsMensuales.ubic, `Operaciones mensuales declaradas ${fmtMoney(K.opsMensuales.v)}.`, `El promedio de ventas con tarjeta es ${fmtMoney(prom)} (${Math.round((K.opsMensuales.v / prom - 1) * 100)}% de diferencia).`, "\u00bfLa transaccionalidad declarada en el KYC es correcta?");
      if (Q.sdgcRate && Q.plazoBase) { const base = Q.plazoBase + (Q.soloInteres || 0); const sd = get("contrato", "comSDGC")?.v; if (sd) { ver(`Comisi\u00f3n SDGC s/IVA = ${(Q.sdgcRate / 1.16 * 100).toFixed(2)}% \u00d7 promedio \u00d7 ${base} (m\u00e9todo del cotizador)`, sd, Q.sdgcRate / 1.16 * prom * base, "Cotizador 'Cotizador 18m'!D12"); /* Regla acordada (05/11/2026): el cotizador calcula la comisi\u00f3n \u00d7 (amortizaciones + 2) y es correcto. */ } }
    }


    // ---- Campos variables poblados (regla acordada 06/11/2026) ----
    // Son los campos que en la plantilla vienen resaltados en turquesa. Aunque ya no vengan marcados,
    // se revisa que cada uno tenga valor y que no quede ning\u00fan marcador sin llenar ([*], \u00ab\u2026\u00bb, {\u2026}, [MAY\u00daSCULAS]).
    {
      const CVAR = report.camposVariables = [];
      const yearRe = /(19|20)\d{2}/;
      const pagesOf = k => { const d = one(k); return d ? d.pages.map(p => ({ no: p.no, t: pageFlat(p).replace(/\s+/g, " ") })) : []; };
      const findAll = (pp, re) => { const out = []; pp.forEach(p => { const r = new RegExp(re.source, re.flags.includes("g") ? re.flags : re.flags + "g"); let m; while ((m = r.exec(p.t))) out.push({ v: (m[1] || "").trim(), no: p.no }); }); return out; };
      const reg = (k, campo, ubic, v, extra) => {
        const val = v == null ? "" : String(v).trim();
        const marcador = /\[\*\]|\u00ab[^\u00bb]+\u00bb|\{(?!(?:signature|initials|date|text|check|name|email|company|title)\s*:)[^}]{2,40}\}|\[[A-Z\u00c1\u00c9\u00cd\u00d3\u00da\u00d1$ :]{3,}\]/.test(val);
        const vacio = !val || /^[\s,.\-_]+$/.test(val) || (extra && extra.re && !extra.re.test(val));
        CVAR.push({ doc: LAB[k], campo, ubic, valor: val || "(vac\u00edo)", estado: marcador ? "Marcador sin llenar" : vacio ? "Vac\u00edo o no localizado" : "Poblado" });
      };
      const fv = (k, key) => { const g = get(k, key); return g ? g.v : null; };
      const fu = (k, key, dflt) => get(k, key)?.ubic || dflt;
      const pm = !(C.personaFisica || K.pf || CV.personaFisica);
      // Contrato
      if (one("contrato")) {
        const P = pagesOf("contrato");
        [["monto", "Monto del cr\u00e9dito"], ["total", "Monto total a pagar"], ["cat", "CAT"], ["tasaOrd", "Tasa ordinaria"], ["tasaMor", "Tasa moratoria"], ["plazo", "Plazo (meses)"]].forEach(([key, nm]) => reg("contrato", "Car\u00e1tula \u00b7 " + nm, "Car\u00e1tula (p.1)", fv("contrato", key)));
        const fechas = findAll(P, /de fecha\s+([^()\u201c"]{0,40}?)\s*\((?:el |\u201cel )?.?Contrato/i).concat(findAll(P, /celebrado en fecha\s+([^()\u201c"]{0,40}?)\s*\(/i)).concat(findAll(P, /de fecha\s+([^()\u201c"]{0,40}?)\s+al cual se encuentra adherid/i));
        if (!fechas.length) reg("contrato", "Fecha del contrato (proemio y anexos)", "Proemio y Anexos A, B, D, E, F", "");
        fechas.forEach(f => reg("contrato", "Fecha del contrato", "p." + f.no, f.v, { re: yearRe }));
        const cierre = findAll(P, /AUT[\u00d3O]GRAFA, EL D[\u00cdI]A\s+([^.]{0,40})\./i); reg("contrato", "Fecha de firma (cierre)", cierre[0] ? "Cierre (p." + cierre[0].no + ")" : "Cierre", cierre[0]?.v, { re: yearRe });
        findAll(P, /identificada en la referencia\s+(\S*)\s*del Anexo de Datos Generales,\s*\(el .Obligado/i).forEach(f => reg("contrato", "Referencia del obligado solidario (proemio)", "Proemio (p." + f.no + ")", f.v));
        findAll(P, /contenida en la referencia\s+(\S*)\s*del Anexo de Datos Generales y su Registro Federal de Contribuyentes es el se\u00f1alado en la referencia\s+(\S*)/i).forEach(f => reg("contrato", "Referencias de domicilio y RFC (declaraciones)", "Declaraciones (p." + f.no + ")", f.v));
        reg("contrato", "Anexo A \u00b7 Sucursales", fu("contrato", "sucursales", "Anexo A"), (C.sucursales || []).length ? (C.sucursales || []).map(x => x.nombre || x.ubicacion || x).join(" | ") : fv("contrato", "sucursales"));
        reg("contrato", "Anexo B \u00b7 Procesador", fu("contrato", "procesador", "Anexo B"), fv("contrato", "procesador"));
        reg("contrato", "Anexo B \u00b7 Fecha del contrato con el procesador", fu("contrato", "fechaProcesador", "Anexo B"), fv("contrato", "fechaProcesador"));
        const aD = C.anexoD || {};
        [["pInv", "% comisi\u00f3n por investigaci\u00f3n"], ["inv", "Monto comisi\u00f3n por investigaci\u00f3n"], ["pDisp", "% comisi\u00f3n por disposici\u00f3n"], ["disp", "Monto comisi\u00f3n por disposici\u00f3n"], ["pSDGC", "% comisi\u00f3n SDGC"], ["sdgc", "Monto comisi\u00f3n SDGC"]].forEach(([key, nm]) => reg("contrato", "Anexo D \u00b7 " + nm, aD.ubic || "Anexo D", aD[key]));
        [[/Nombre del Cliente:\s*(.{0,80}?)\s+Cr[\u00e9e]dito:/i, "Nombre del cliente"], [/Cr[\u00e9e]dito:\s*(\S*)\s+Referencia:/i, "Cr\u00e9dito"], [/Referencia:\s*(\S*)\s+Fecha de elaboraci/i, "Referencia"], [/Fecha de elaboraci[\u00f3o]n:\s*(\S*)\s+Producto/i, "Fecha de elaboraci\u00f3n"], [/Monto total a pagar:\s*(\$?[\d,]*\.?\d*)/i, "Monto total a pagar"]].forEach(([re, nm]) => { const f = findAll(P, re)[0]; reg("contrato", "Anexo F \u00b7 " + nm, f ? "Anexo F (p." + f.no + ")" : "Anexo F", f?.v); });
        const adg = pm ? [["razon", "Denominaci\u00f3n"], ["rfc", "RFC"], ["fechaConst", "Fecha de constituci\u00f3n"], ["escrituraConst", "Escritura constitutiva"], ["fedatarioConst", "Fedatario de la constituci\u00f3n"], ["folio", "Folio mercantil"], ["domCP", "C\u00f3digo postal"], ["representante", "Representante legal"], ["escrituraPoder", "Escritura del poder"], ["fechaPoder", "Fecha del poder"], ["referencia", "Referencia de pago"], ["emailNotif", "Correo de notificaciones"]]
          : [["razon", "Nombre"], ["rfc", "RFC"], ["curp", "CURP"], ["domCP", "C\u00f3digo postal"], ["referencia", "Referencia de pago"], ["emailNotif", "Correo de notificaciones"]];
        // CP del acreditado: si el extractor no lo tom\u00f3, se busca el rengl\u00f3n bajo cada etiqueta "C\u00f3digo Postal" del ADG (excepto el de Finsus)
        const adgP = one("contrato").pages.find(p => /Anexo de Datos Generales/i.test(pageFlat(p)) && /Referencia\s*[24]\.\s*Datos Generales/i.test(pageFlat(p)));
        let cpADG = null;
        if (adgP) { const L = buildLines(adgP); L.forEach((l, i) => { if (cpADG || !/C[\u00f3o]digo Postal/i.test(l.text) || !L[i + 1]) return; const c = (L[i + 1].text.match(/\b\d{5}\b/g) || []).find(x => x !== "03230"); if (c) cpADG = c; }); }
        const cpPF = (C.domPF || []).map(d => d.cp).filter(Boolean).join(" / ");
        adg.forEach(([key, nm]) => reg("contrato", "Anexo de Datos Generales \u00b7 " + nm, fu("contrato", key, "Anexo de Datos Generales" + (adgP ? " (p." + adgP.no + ")" : "")), fv("contrato", key) ?? (key === "domCP" ? (cpPF || cpADG) : null)));
      }
      // Pagar\u00e9
      if (one("pagare")) {
        reg("pagare", "Suscriptor", fu("pagare", "razon", "P\u00e1rrafo 1"), fv("pagare", "razon"));
        reg("pagare", "Suma principal (n\u00famero)", fu("pagare", "sumaPrincipal", "P\u00e1rrafo 1"), fv("pagare", "sumaPrincipal"));
        reg("pagare", "Suma principal (letra)", PG.letra?.ubic || "P\u00e1rrafo 1", PG.letra?.v);
        reg("pagare", "Tasa ordinaria", fu("pagare", "tasaOrd", "Intereses ordinarios"), fv("pagare", "tasaOrd"));
        reg("pagare", "Tasa moratoria", fu("pagare", "tasaMor", "Intereses moratorios"), fv("pagare", "tasaMor"));
        const pgT = one("pagare").pages.map(p => pageFlat(p)).join(" ").replace(/\s+/g, " ");
        const domBlk = grab(pgT, /como su domicilio el ubicado en:\s*(.{0,600}?)\s*Este Pagar[\u00e9e] se suscribe/i) || "";
        const domSus = fv("pagare", "domCalle") || grab(domBlk, /^(.*?)(?:\s*(?:Avale?s?:|Ava:|Direcci[\u00f3o]n Aval)|$)/i);
        reg("pagare", "Domicilio del suscriptor", fu("pagare", "domCalle", "Domicilios"), domSus);
        if (PG.conAval) reg("pagare", "Domicilio del aval", fu("pagare", "avalCalle", "Domicilio del aval"), fv("pagare", "avalCalle") || grab(domBlk, /(?:Avale?s?:|Ava:|Direcci[\u00f3o]n Aval 1|Domicilio del Aval:?)\s*(.+)$/i));
        reg("pagare", "Fecha de suscripci\u00f3n", fu("pagare", "fechaFirma", "Lugar y fecha"), fv("pagare", "fechaFirma"), { re: yearRe });
        reg("pagare", "N\u00famero de hojas", "P\u00e1rrafo final", PG.hojasTexto || grab(pgT, /se suscribe en\s+(.{1,40}?)\s+(?:Hojas|fojas)/i), { re: /^\d+/ });
        reg("pagare", "Nombre en la firma del suscriptor", fu("pagare", pm ? "representante" : "razon", "Firmas"), pm ? fv("pagare", "representante") : fv("pagare", "razon"));
        if (PG.conAval) { const avs = [...pgT.matchAll(/Aval 1.?\s*_*\s*([A-Z\u00c1\u00c9\u00cd\u00d3\u00da\u00d1]{2,}(?: [A-Z\u00c1\u00c9\u00cd\u00d3\u00da\u00d1]{2,}){1,5})(?![,0-9A-Za-z])/g)].map(m => m[1].trim()).filter(n => !/^(LA|EL|CALLE|AV|AVE)\b/.test(n)); reg("pagare", "Nombre del aval", "Firma del aval", avs.length ? avs[avs.length - 1] : fv("pagare", "obligado")); }
      }
      // Convenio
      if (one("convenio")) {
        reg("convenio", "Fideicomitente adherente", fu("convenio", "razon", "Proemio"), fv("convenio", "razon"));
        if (pm) reg("convenio", "Representante", fu("convenio", "representante", "Proemio"), fv("convenio", "representante"));
        reg("convenio", "Fecha del contrato de cr\u00e9dito", fu("convenio", "fechaFirma", "Antecedente II"), fv("convenio", "fechaFirma"), { re: yearRe });
        reg("convenio", "Monto del cr\u00e9dito (n\u00famero)", fu("convenio", "monto", "Antecedente II"), fv("convenio", "monto"));
        reg("convenio", "Monto del cr\u00e9dito (letra)", CV.letra?.ubic || "Antecedente II", CV.letra?.v);
        reg("convenio", "Domicilio (Cl\u00e1usula Octava)", "Cl\u00e1usula Octava", CV.domOct || fv("convenio", "domCalle"));
        reg("convenio", "Correo (Cl\u00e1usula Octava)", fu("convenio", "emailNotif", "Cl\u00e1usula Octava"), fv("convenio", "emailNotif"));
        reg("convenio", "Anexo A \u00b7 Procesador", fu("convenio", "procesador", "Anexo A"), fv("convenio", "procesador"));
        reg("convenio", "Anexo A \u00b7 Fecha del contrato con el procesador", fu("convenio", "fechaProcesador", "Anexo A"), fv("convenio", "fechaProcesador"));
        reg("convenio", "Anexo B \u00b7 N\u00famero de cr\u00e9dito", fu("convenio", "credito", "Anexo B"), fv("convenio", "credito"));
        reg("convenio", "Anexo B \u00b7 Fecha", "Anexo B", CV.anexoBFecha, { re: /\d{1,2}\/\d{1,2}\/\d/ });
        reg("convenio", "Anexo B \u00b7 Monto del pagar\u00e9", CV.montoPagare?.ubic || "Anexo B", CV.montoPagare?.v);
        reg("convenio", "Anexo B \u00b7 % retenci\u00f3n", fu("convenio", "retencion", "Anexo B"), fv("convenio", "retencion"));
        reg("convenio", "Anexo B \u00b7 Pago mensual m\u00ednimo", fu("convenio", "pagoMinimo", "Anexo B"), fv("convenio", "pagoMinimo"));
        const cvT = one("convenio").pages.map(p => pageFlat(p)).join(" ").replace(/\s+/g, " ");
        const encs = [...cvT.matchAll(/RE:\s*Notificaci[\u00f3o]n a los Procesadores[^.]*?derivados del Contrato\.?\s*(.{0,140}?)\s*Estimados/gi)].map(m => m[1].replace(/[|+=\-]+/g, " ").trim()).filter(v => v && !/[{}]/.test(v));
        reg("convenio", "Notificaci\u00f3n \u00b7 Afiliaci\u00f3n / encabezado", "Notificaci\u00f3n al procesador", CV.notifEncabezado || encs[encs.length - 1] || (CV.afiliaciones || []).join(", "));
        const lugarF = [...cvT.matchAll(/Ciudad de M[\u00e9e]xico a\s+([^A-Z{]{0,40}?\d{4})/g)].map(m => m[1]);
        reg("convenio", "Notificaci\u00f3n \u00b7 Lugar y fecha", "Notificaci\u00f3n al procesador", lugarF[lugarF.length - 1], { re: yearRe });
      }
      // Marcadores sin llenar en cualquier parte (se excluyen los formatos en blanco por dise\u00f1o:
      // Anexo C del contrato = formato de pagar\u00e9; Anexos C y D del convenio = formatos de notificaci\u00f3n y terminaci\u00f3n)
      // Las etiquetas de firma electr\u00f3nica ({signature:user1___}, {initials:...}) de contratos firmados en l\u00ednea no son marcadores de plantilla
      const marcRe = /\[\*\]|\u00ab[^\u00bb]{2,60}\u00bb|\{(?!(?:signature|initials|date|text|check|name|email|company|title)\s*:)[^{}]{2,60}\}|\[(?:NOMBRE|MONTO|FECHA|DOMICILIO)[^\]]*\]/g;
      ["contrato", "pagare", "convenio"].forEach(k => {
        const d = one(k); if (!d) return;
        let full = "", offs = [];
        d.pages.forEach(p => { offs.push([full.length, p.no]); full += pageFlat(p).replace(/\s+/g, " ") + " "; });
        const pageAt = pos => { let no = offs[0][1]; for (const [o, n] of offs) { if (o <= pos) no = n; else break; } return no; };
        const skip = [];
        if (k === "contrato") { const a = full.search(/ANEXO C\b.{0,60}\[Formato\]/i); if (a >= 0) { const b = full.slice(a + 10).search(/ANEXO D\b/); skip.push([a, b >= 0 ? a + 10 + b : full.length]); } }
        if (k === "convenio") { const a = full.search(/Formato de Notificaci[\u00f3o]n/i); if (a >= 0) { const rest = full.slice(a); const b = rest.search(/ESPACIO DEJADO INTENCIONALMENTE EN BLANCO/i); skip.push([a, b >= 0 ? a + b : full.length]); } }
        const seen = new Set(); let m; marcRe.lastIndex = 0;
        while ((m = marcRe.exec(full))) {
          if (skip.some(([a, b]) => m.index >= a && m.index < b)) continue;
          if (CVAR.some(c => c.doc === LAB[k] && c.estado !== "Poblado" && String(c.valor).includes(m[0].slice(0, 12)))) continue;
          const no = pageAt(m.index), key = m[0] + "|" + no; if (seen.has(key)) continue; seen.add(key);
          const ctx = full.slice(Math.max(0, m.index - 50), m.index + m[0].length + 30);
          CVAR.push({ doc: LAB[k], campo: "Marcador en el texto", ubic: (d.quality === "Word" ? "secci\u00f3n " : "p.") + no, valor: "\u2026 " + ctx + " \u2026", estado: "Marcador sin llenar" });
        }
      });
      CVAR.filter(c => c.estado !== "Poblado" && !(/Procesador/.test(c.campo) && report.discrepancias.some(d => /Procesador.*sin nombre/i.test(d.campo)))).forEach(c => disc(c.doc === "Pagar\u00e9" ? "Alta" : "Media", "Campo sin llenar: " + c.campo, `${c.doc}, ${c.ubic}: ${c.valor}.`, "Debe tener el dato del expediente.", c.estado === "Marcador sin llenar" ? "Qued\u00f3 el marcador de la plantilla." : "El campo est\u00e1 vac\u00edo o no se pudo localizar; conf\u00edrmalo en el documento."));
    }

    // ---- Regla de fechas (acordada 15/11/2026) ----
    // Grupo A: fecha del contrato con el procesador. Debe ser la MISMA en los 3 lugares: Anexo B del contrato,
    //          Anexo A del convenio y la notificaci\u00f3n al procesador ("celebrado el ..."). Puede ser distinta a la firma.
    // Grupo B: todas las dem\u00e1s fechas del contrato, pagar\u00e9 y convenio (proemio, cierre, anexos, domiciliaciones,
    //          encabezado de la notificaci\u00f3n, KYC, TA, cotizador) deben ser la fecha de firma.
    {
      const fmtD = iso => iso ? iso.split("-").reverse().join("/") : "";
      const grpA = [];
      const pushA = (doc, ubic, v) => { const d = parseDate(v); if (d) grpA.push({ doc, ubic, v, d }); };
      const fpC = get("contrato", "fechaProcesador"); if (fpC) [fpC, ...(fpC.more || [])].forEach(e => pushA("Contrato", e.ubic, e.v));
      const fpV = get("convenio", "fechaProcesador"); if (fpV) [fpV, ...(fpV.more || [])].forEach(e => pushA("Convenio de adhesi\u00f3n", e.ubic, e.v));
      if (one("convenio")) {
        const cvT2 = one("convenio").pages.map(p => pageFlat(p)).join(" ").replace(/\s+/g, " ");
        [...cvT2.matchAll(new RegExp("Contrato de Afiliaci[\u00f3o]n del Comercio[^.]{0,80}?celebrado el " + DATE_WORDS, "gi"))].forEach(m => { if (!grpA.some(g => /Notificaci/.test(g.ubic) && g.v === m[1])) pushA("Convenio de adhesi\u00f3n", "Notificaci\u00f3n al procesador ('celebrado el \u2026')", m[1]); });
      }
      const distA = [...new Set(grpA.map(g => g.d))];
      if (grpA.length >= 2 && distA.length > 1) covered.add("fechaProcesador"), disc("Alta", "Fecha del contrato con el procesador (Anexo B, Anexo A del convenio y notificaci\u00f3n)", grpA.map(g => `${g.doc}, ${g.ubic}: ${g.v}`).join(" | "), `Se encontraron ${distA.length} fechas distintas: ${distA.map(fmtD).join(", ")}.`, "Estas tres fechas siempre deben coincidir entre s\u00ed (pueden ser distintas a la fecha de firma).");
      else if (grpA.length === 1 && (one("contrato") && one("convenio"))) duda("Contrato / Convenio", grpA[0].ubic, `Solo se encontr\u00f3 la fecha del procesador en un lugar (${grpA[0].v}).`, "No se pudo leer en los otros dos apartados para confirmar que coincidan.", "\u00bfAnexo B, Anexo A del convenio y notificaci\u00f3n traen la misma fecha?");
      report.fechasProcesador = grpA;
      // Inventario de TODAS las fechas del expediente (contrato, pagar\u00e9, convenio, KYC, TA, cotizador) con documento,
      // p\u00e1gina, apartado y clasificaci\u00f3n; de ah\u00ed sale tambi\u00e9n la verificaci\u00f3n del grupo B.
      const INV = report.inventarioFechas = [];
      {
        const firmaD = firma || null, vencD = parseDate(get("contrato", "vencimiento")?.v || get("pagare", "vencimiento")?.v), primerD = parseDate(get("contrato", "primerPago")?.v || get("pagare", "primerPago")?.v), skipA = new Set(grpA.map(g => g.d));
        const fechaConst = parseDate(get("contrato", "fechaConst")?.v || get("kyc", "fechaConst")?.v), fechaPoder = parseDate(get("contrato", "fechaPoder")?.v || get("kyc", "fechaPoder")?.v);
        const taRows = (PG.rows?.length ? PG.rows : (C.rowsF?.length ? C.rowsF : (X("ta").rows || []))); const calendario = new Set(taRows.map(r => r.fecha).filter(Boolean));
        const seccion = (k, t, p) => {
          if (k === "contrato") {
            if (/Car[\u00e1a]tula/i.test(t) && p.no <= 2) return "Car\u00e1tula";
            const head = t.slice(0, 500);
            if (/Anexo de Datos Generales/i.test(head) && /Referencia\s*[1-9]\.\s*(Datos|Informaci)/i.test(t) && !/CONTRATO DE APERTURA/i.test(head)) return "Anexo de Datos Generales";
            if (/ANEXO .?A\b/i.test(head) && /Sucursal|Listado/i.test(t)) return "Anexo A (sucursales)";
            if (/ANEXO .?B\b/i.test(head) && /procesa/i.test(t)) return "Anexo B (procesadores)";
            if (/ANEXO .?C\b/i.test(head) || /\[Formato\]/.test(head)) return "Anexo C (formato de pagar\u00e9)";
            if (/ANEXO .?D\b/i.test(head) && /Comisi/i.test(t)) return "Anexo D (comisiones)";
            if (/ANEXO .?E\b/i.test(head) && /Preceptos/i.test(t)) return "Anexo E (preceptos)";
            if (/ANEXO .?F\b/i.test(head) || /Tabla de amortizaci/i.test(head)) return "Anexo F (tabla de amortizaci\u00f3n)";
            if (/Domiciliaci[\u00f3o]n|cargo recurrente|AUTORIZACI[\u00d3O]N DE CARGO/i.test(head)) return "Domiciliaci\u00f3n";
            if (/Billpocket|Pocket de Latinoam|BESP-|Contrato de Afiliaci|Afiliaci[\u00f3o]n del Comercio/i.test(t)) return "Contrato con el procesador";
            if (/DOCUMENT COMPLETED|TIMESTAMP|Audit Trail|Constancia de (Firma|Conservaci)/i.test(t)) return "Constancia de firma electr\u00f3nica";
            if (/AUT[\u00d3O]GRAFA, EL D[\u00cdI]A/i.test(t)) return "Cierre y firmas";
            if (p.no <= 3) return "Proemio / antecedentes";
            return "Cl\u00e1usulas";
          }
          if (k === "convenio") {
            if (/Formato de Notificaci[\u00f3o]n|Formato de Terminaci[\u00f3o]n/i.test(t)) return "Anexos C/D (formatos en blanco)";
            if (/RE:\s*Notificaci[\u00f3o]n|Estimados se[\u00f1n]ores/i.test(t)) return "Notificaci\u00f3n al procesador";
            if (/Anexo .?A.{0,40}Convenio|Contratos de Servicios/i.test(t)) return "Anexo A (procesadores)";
            if (/Anexo .?B.{0,40}Convenio|Distribuci[\u00f3o]n de (la )?Cobranza/i.test(t)) return "Anexo B (distribuci\u00f3n de cobranza)";
            if (/Antecedentes|ANTECEDENTES/.test(t) && p.no <= 2) return "Proemio / antecedentes";
            if (/firman el presente|FIDEICOMITENTE ADHERENTE\s*_|Por:/i.test(t) && /Delegad|Fiduciari/i.test(t)) return "Cierre y firmas";
            return "Cl\u00e1usulas";
          }
          if (k === "pagare") { if (/Fecha de pago|No\.\s*Pago|Capital/i.test(t) && /\d{2}\/\d{2}\/\d{4}/.test(t)) return "Calendario de pagos"; return "Texto del pagar\u00e9"; }
          if (k === "kyc") return "KYC p." + p.no;
          return "";
        };
        const tipoDe = (iso, ctxW, raw, sec) => {
          if (/Constancia de firma/.test(sec) || /\b(SIGNED|VIEWED|TIMESTAMP|UTC|COMPLETED BY ALL PARTIES)\b/.test(ctxW)) return "Firma electr\u00f3nica del procesador";
          if (/Fideicomiso|F\/0394|Fiduciari|SOFIMEX/i.test(ctxW) && /2025/.test(raw)) return "Fideicomiso F/0394";
          if (iso === fechaConst || /constituci|acta constitutiva/i.test(ctxW)) return "Constituci\u00f3n de la sociedad";
          if (iso === fechaPoder || /poder|escritura/i.test(ctxW)) return "Escritura / poder";
          if (/nacimiento/i.test(ctxW)) return "Fecha de nacimiento";
          if (/instrumento p[\u00fau]blico|escrituras? p[\u00fau]blicas?|otorgad[ao] ante|ante la fe|protocoli|notari/i.test(ctxW)) return "Escritura (acreditante / acreditado)";
          if (/DOF|Diario Oficial|publicad[ao] en/i.test(ctxW)) return "Publicaci\u00f3n en el DOF";
          if (/Folio Mercantil|inscri|Registro P[\u00fau]blico|RPC|RPP/i.test(ctxW)) return "Inscripci\u00f3n en el Registro P\u00fablico";
          if (/Fecha de afiliaci|celebrado el|Fecha de celebraci|Contratos de Servicio|Nombre del procesador|procesador/i.test(ctxW) || (skipA.has(iso) && /Anexo .?[AB]\b|Contrato con el procesador|Notificaci/i.test(sec))) return "Contrato con el procesador (grupo A)";
          if (iso === vencD || /vencimiento|8\.3\b/i.test(ctxW)) return "Vencimiento del cr\u00e9dito";
          if (iso === primerD) return "Primer pago";
          if (calendario.has(iso) || /Calendario|Anexo F|Tabla de amortizaci/i.test(sec) || /Fecha de pago|Pago\s+\d|No\.\s*\d|Saldo/i.test(ctxW)) return "Calendario de pagos";
          if (/Firmado|signed|Completed|Enviado|Visto|Sello|UTC|TIMESTAMP/i.test(ctxW)) return "Firma electr\u00f3nica del procesador";
          if (/vigencia|expira|v\u00e1lida hasta/i.test(ctxW)) return "Vigencia";
          if (/Fecha de inicio|8\.2\b/i.test(ctxW)) return "Inicio del contrato";
          return "Fecha de firma (grupo B)";
        };
        const scan = (k, docLabel) => {
          const d = one(k); if (!d) return;
          d.pages.forEach(p => {
            const t = pageFlat(p).replace(/\s+/g, " "); const sec = seccion(k, t, p);
            const re = new RegExp(DATE_WORDS + "|\\b(\\d{1,2}\\/\\d{1,2}\\/\\d{4})\\b|\\b(\\d{1,2} [A-Z]{3} \\d{4})\\b", "gi"); let m;
            while ((m = re.exec(t))) {
              const raw = m[1] || m[2] || m[3]; const iso = parseDate(raw); if (!iso) continue;
              const ctx = t.slice(Math.max(0, m.index - 70), m.index + raw.length + 40), ctxW = t.slice(Math.max(0, m.index - 70), m.index + raw.length + 160);
              let tipo = tipoDe(iso, ctxW, raw, sec);
              if (tipo === "Fecha de firma (grupo B)" && firmaD) { const dd = Math.abs(dayDiff(iso, firmaD)); if (dd > 400) tipo = "Otra fecha (anterior / posterior al cr\u00e9dito)"; }
              const enFormato = /Anexo C \(formato|formatos en blanco|Constancia de firma/.test(sec);
              if (enFormato && tipo === "Fecha de firma (grupo B)") tipo = /Constancia/.test(sec) ? "Firma electr\u00f3nica del procesador" : "Formato en blanco";
              if (sec === "Contrato con el procesador" && tipo === "Fecha de firma (grupo B)") tipo = "Contrato con el procesador (texto)";
              const esperado = tipo.startsWith("Fecha de firma") ? firmaD : tipo.startsWith("Contrato con el procesador (grupo A)") ? (grpA[0]?.d || null) : null;
              const estado = tipo.startsWith("Fecha de firma") ? (firmaD ? (iso === firmaD ? "Coincide con la firma" : "No coincide con la firma") : "Sin fecha de firma de referencia") : tipo.startsWith("Contrato con el procesador (grupo A)") ? (grpA.length && distA.length === 1 ? "Coincide (grupo A)" : "Revisar (grupo A)") : "Informativa";
              INV.push({ doc: docLabel, pagina: p.no, apartado: sec, texto: raw, fecha: iso, tipo, estado, contexto: ctx.trim().slice(0, 160) });
            }
          });
        };
        scan("contrato", "Contrato"); scan("pagare", "Pagar\u00e9"); scan("convenio", "Convenio de adhesi\u00f3n"); scan("kyc", "KYC");
        // Excel: TA y cotizador
        const taF = get("ta", "fechaFirma"); if (taF) INV.push({ doc: "Tabla de amortizaci\u00f3n", pagina: "", apartado: taF.ubic, texto: String(taF.v), fecha: parseDate(taF.v), tipo: "Fecha de impresi\u00f3n de la TA", estado: firmaD && parseDate(taF.v) === firmaD ? "Coincide con la firma" : "Informativa", contexto: "" });
        const coF = get("cotizador", "fechaFirma"); if (coF) INV.push({ doc: "Cotizador", pagina: "", apartado: coF.ubic, texto: coF.v instanceof Date ? fmtD(parseDate(coF.v)) : String(coF.v), fecha: parseDate(coF.v), tipo: "Fecha de la oferta", estado: firmaD && parseDate(coF.v) === firmaD ? "Coincide con la firma" : "Informativa", contexto: "" });
        (X("ta").rows || []).forEach((r, i) => { if (r.fecha) INV.push({ doc: "Tabla de amortizaci\u00f3n", pagina: "", apartado: "Pago " + (i + 1), texto: fmtD(r.fecha), fecha: r.fecha, tipo: "Calendario de pagos", estado: "Informativa", contexto: "" }); });
        // Grupo B: discrepancias
        const grpB = INV.filter(x => x.estado === "No coincide con la firma");
        const byD = new Map(); grpB.forEach(g => { (byD.get(g.fecha) || byD.set(g.fecha, []).get(g.fecha)).push(g); });
        byD.forEach((arr, d) => {
          const where = arr.slice(0, 6).map(g => `${g.doc} p.${g.pagina} (${g.apartado}): '${g.texto}'`).join("; ") + (arr.length > 6 ? `; \u2026 (${arr.length} en total)` : "");
          disc("Media", `Fecha distinta a la de firma (${fmtD(d)})`, where, `Fecha de firma del contrato: ${fmtD(firmaD)}.`, "Fuera de las tres fechas del contrato con el procesador, todas las fechas del contrato, pagar\u00e9 y convenio deben ser la fecha de firma.");
        });
        report.fechasFirmaDistintas = grpB;
      }
    }

    // Firmantes como "Representante Legal" en pagar\u00e9 / convenio que no son el apoderado del ADG (ref. 3.1) ni el del KYC
    {
      const repsOk = [get("contrato", "representante")?.v, get("kyc", "representante")?.v].filter(Boolean).flatMap(v => String(v).split(/\s*(?:,|\by\b)\s*/i)).map(norm).filter(n => n.length > 5);
      if (repsOk.length) {
        const extra = new Map();
        for (const k of ["pagare", "convenio"]) {
          const d = one(k); if (!d) continue;
          d.pages.forEach(p => {
            const L = buildLines(p);
            L.forEach((l, i) => {
              if (!/^Representante Legal$/i.test(l.text.trim())) return;
              const nm = (L.slice(Math.max(0, i - 2), i).reverse().find(x => /[A-Z\u00c1\u00c9\u00cd\u00d3\u00da\u00d1]{3,}\s+[A-Z\u00c1\u00c9\u00cd\u00d3\u00da\u00d1]{3,}/.test(x.text) && !/^_+$/.test(x.text.trim())) || {}).text;
              if (!nm) return;
              nm.split(/\s*,\s*/).map(s => s.trim()).filter(s => s.length > 5).forEach(n => {
                if (repsOk.some(r => similar(r, norm(n)) >= 0.85 || r.includes(norm(n)) || norm(n).includes(r))) return;
                const key = norm(n); const e = extra.get(key) || { n, ubic: [] }; e.ubic.push(LAB[k] + " p." + p.no); extra.set(key, e);
              });
            });
          });
        }
        extra.forEach(e => disc("Alta", "Firmante como 'Representante Legal' sin poder en el expediente", `${[...new Set(e.ubic)].join("; ")}: firma ${e.n} como 'Representante Legal'.`, `Anexo de Datos Generales ref. 3.1 y KYC: el apoderado es ${get("contrato", "representante")?.v || get("kyc", "representante")?.v}.`, "No hay datos del poder de esa persona; revisa si debe firmar o quitarla del documento."));
      }
    }
    // pantalla "Datos Captura" del sistema
    const CAP = X("captura");
    if (one("captura")) {
      const cl = "Captura del sistema", L = CAP.labels || {};
      if (CAP.encontrados < 8) duda(cl, "Imagen completa", `Solo se reconocieron ${CAP.encontrados} campos de la pantalla.`, "La imagen es peque\u00f1a o est\u00e1 recortada; el OCR no ley\u00f3 todos los campos.", "\u00bfPuedes enviar la captura completa y a mayor resoluci\u00f3n?");
      const confirmado = k => { const a = CAP.vals[k]; if (a == null) return false; const ref = { pInv: C.anexoD?.pInv, pDisp: C.anexoD?.pDisp, pSDGC: C.anexoD?.pSDGC, ret: get("convenio", "retencion")?.v ?? get("cotizador", "retencion")?.v, tasaOrd: get("contrato", "tasaOrd")?.v, tasaMor: get("contrato", "tasaMor")?.v, cat: get("contrato", "cat")?.v }[k]; return ref != null && Math.abs((/^p/.test(k) ? +ref : pct(ref)) - a) < 0.011; };
      const dud = Object.entries({ ret: 1, pInv: 1, pDisp: 1, pSDGC: 1, tasaOrd: 1, tasaMor: 1, cat: 1 }).map(([k]) => k).filter(k => (CAP.sospechosos || []).some(s => s.startsWith(L[k] + ":")) && !confirmado(k));
      if (dud.length) duda(cl, dud.map(k => "campo '" + L[k] + "'").join(", "), `Lecturas dudosas del OCR: ${CAP.sospechosos.filter(s => dud.some(k => s.startsWith(L[k] + ":"))).join("; ")}.`, "Un d\u00edgito o el punto decimal pueden haberse perdido al leer la imagen; no coinciden con los documentos.", "Confirma esos valores en la pantalla del sistema.");
      if (C.anexoD) [["pInv", "pInv", "Comisi\u00f3n por gastos de investigaci\u00f3n"], ["pDisp", "pDisp", "Comisi\u00f3n por disposici\u00f3n"], ["pSDGC", "pSDGC", "Comisi\u00f3n por gesti\u00f3n de cobranza"]].forEach(([k, ak, nm]) => {
        const a = CAP.vals[k], b = C.anexoD[ak]; if (a == null || b == null) return;
        if (Math.abs(a - b) < 0.001) report.resumen.push({ cat: "Montos y Tasas", campo: nm + " (%)", valor: b + "%", doc: "Contrato", ubic: C.anexoD.ubic, st: "S\u00ed", obs: "" }), report.resumen.push({ cat: "Montos y Tasas", campo: nm + " (%)", valor: a + "%", doc: cl, ubic: "campo '" + L[k] + "'", st: "S\u00ed", obs: "(le\u00eddo con OCR)" });
        else if ((CAP.sospechosos || []).some(s => s.startsWith(L[k] + ":"))) { /* ya va en la duda de lecturas del OCR */ }
        else disc("Media", nm + ": sistema vs. Anexo D", `Captura del sistema, campo '${L[k]}': ${a}%.`, `${C.anexoD.ubic}: ${b}%.`, "Lo capturado en el sistema no coincide con el contrato.");
      });
      if (CAP.vals.clabeAcred) { const ctext = squash(one("contrato") ? one("contrato").pages.map(pageFlat).join(" ") : ""); if (CAP.vals.clabeAcred.length !== 18) duda(cl, "campo 'Cuenta Clabe Acreditante'", `Se ley\u00f3 '${CAP.vals.clabeAcred}'.`, "No tiene 18 d\u00edgitos; lectura incompleta.", "Confirma la CLABE en la pantalla."); else if (ctext && !ctext.includes(CAP.vals.clabeAcred)) disc("Media", "CLABE del acreditante en el sistema", `Captura del sistema: ${CAP.vals.clabeAcred}${CAP.banco ? " (" + CAP.banco + ")" : ""}.`, "No aparece en el contrato (ref. 9 / cuenta del fideicomiso).", "La cuenta capturada no es la del contrato."); else report.resumen.push({ cat: "Cuentas Bancarias", campo: "Cuenta concentradora (CLABE)", valor: CAP.vals.clabeAcred, doc: cl, ubic: "campo 'Cuenta Clabe Acreditante'", st: "S\u00ed", obs: "Aparece tambi\u00e9n en el contrato." }); }
      if (CAP.vals.numProc && one("convenio")) { const cvt = squash(one("convenio").pages.map(pageFlat).join(" ")); if (!cvt.includes(squash(CAP.vals.numProc))) duda(cl, "campo '# Contrato Procesador'", `N\u00famero ${CAP.vals.numProc}.`, "Ese n\u00famero no aparece en el convenio ni en su notificaci\u00f3n al procesador.", "\u00bfCu\u00e1l es el n\u00famero de afiliaci\u00f3n correcto?"); }
      const ocC = get("contrato", "ocupacion")?.v || get("kyc", "ocupacion")?.v;
      if (CAP.ocupacion !== undefined && !String(CAP.ocupacion || "").trim()) disc("Baja", "Ocupaci\u00f3n en el sistema", "Captura del sistema: campo 'Ocupaci\u00f3n' vac\u00edo.", ocC ? `Contrato / KYC: ${ocC}.` : "El contrato tampoco la indica.", "Dato incompleto en el sistema.");
      if (!(C.personaFisica || K.pf || CV.personaFisica) && CAP.tipoSoc !== undefined && !String(CAP.tipoSoc || "").trim()) disc("Baja", "Tipo de sociedad en el sistema", "Captura del sistema: campo 'Tipo Sociedad' vac\u00edo.", "El acreditado es persona moral.", "Dato incompleto en el sistema.");
    }
    const fbf = C.fechaBloqueFirma; if (fbf && firma && parseDate(fbf.v) !== firma) covered.add("fechaFirma_bloque"); // la discrepancia la reporta la regla de fechas (grupo B)
    if (K.anios?.v && fc) { const yrs = yearsBetween(fc, firma || new Date().toISOString().slice(0, 10)); if (+K.anios.v > yrs + 1) disc("Media", "Antig\u00fcedad en el domicilio (KYC)", `${K.anios.ubic}: ${K.anios.v} a\u00f1os.`, `Constituci\u00f3n: ${fmtDate(fc)} (${yrs.toFixed(1)} a\u00f1os).`, "La antig\u00fcedad declarada es mayor que la vida de la sociedad."); }
    const cred = get("contrato", "credito")?.v || get("ta", "credito")?.v;
    // Regla acordada (05/11/2026): el '#\u2026' de 'Fines que pretende dar al fideicomiso' es el ID del caso en la plataforma Anticipa;
    // no corresponde al cr\u00e9dito de Open Fin ni al ID de Analyzer, as\u00ed que no se compara.
    if (K.contrato?.v) report.resumen.push({ cat: "Identificadores", campo: "N\u00famero de caso Anticipa citado en el KYC", valor: "#" + K.contrato.v, doc: "KYC", ubic: K.contrato.ubic, st: "Solo en un documento", obs: "ID del caso en la plataforma Anticipa; no se compara con Open Fin ni con Analyzer." });
    // sucursales vs anexo A
    if (C.sucursales?.length) {
      const seen = {}; C.sucursales.forEach(s => { const k = squash(s.dom.replace(/\b(av|avenida|calle)\b/gi, "")); (seen[k] = seen[k] || []).push(s.nombre); });
      const dups = Object.values(seen).filter(v => v.length > 1);
      if (dups.length) disc("Baja", "Sucursales con el mismo domicilio (Anexo A)", `Anexo A: ${dups.map(d => d.join(" / ")).join("; ")} comparten domicilio.`, `${C.sucursales.length} renglones; ${Object.keys(seen).length} domicilios distintos.`, "Puede ser dos terminales en un local o un error del listado.");
    }
    // montos con letra
    const lets = [["Contrato, ref. 8.1", get("contrato", "monto")?.v, C.montoLetra?.v], ["Pagar\u00e9, suma principal", sp, PG.letra?.v], ["Convenio, antecedente II", get("convenio", "monto")?.v, CV.letra?.v]];
    lets.forEach(([w, n, l]) => { const ok = letraOk(n, l); if (ok === false) disc("Alta", "Cantidad con n\u00famero vs. letra", `${w}: ${fmtMoney(n)}.`, `Con letra: '${l}'.`, "En un pagar\u00e9 prevalece la cantidad escrita con letra."); });
    // verificaciones del calendario
    const rows = PG.rows?.length ? PG.rows : (C.rowsF?.length ? C.rowsF : T.rows || []);
    const rowsSrc = PG.rows?.length ? "pagar\u00e9" : C.rowsF?.length ? "Anexo F" : "TA";
    const monto = get("contrato", "monto")?.v ?? get("ta", "monto")?.v, tasa = get("contrato", "tasaOrd")?.v ?? 22;
    if (rows.length && monto) {
      ver(`Suma de capital del calendario (${rowsSrc}) = monto del cr\u00e9dito`, rows.reduce((a, r) => a + r.cap, 0), monto);
      const tot = get("contrato", "total")?.v; if (tot) ver(`Suma de pagos del calendario (${rowsSrc}) = monto total a pagar`, rows.reduce((a, r) => a + r.pago, 0), tot);
      if (sp) ver("Suma principal del pagar\u00e9 = cr\u00e9dito + comisi\u00f3n SDGC con IVA (cl. 2.5)", sp, monto + rows.reduce((a, r) => a + r.com + r.ivacom, 0));
      // Regla acordada (05/11/2026): la suma principal del pagar\u00e9 = Monto + 'Comisi\u00f3n SGDC Monto' del cotizador ('Cotizador 18m'!D12)
      if (sp && Q.sdgcCIVA) { const esperado = r2(monto + r2(Q.sdgcCIVA)); ver("Suma principal del pagar\u00e9 = Monto + 'Comisi\u00f3n SGDC Monto' del cotizador (D12)", sp, esperado, "Regla acordada"); if (Math.abs(sp - esperado) >= 1) disc("Alta", "Suma principal del pagar\u00e9 vs. cotizador", `Pagar\u00e9, p\u00e1rrafo 1: ${fmtMoney(sp)}.`, `Cotizador: Monto ${fmtMoney(monto)} + 'Comisi\u00f3n SGDC Monto' (D12) ${fmtMoney(Q.sdgcCIVA)} = ${fmtMoney(esperado)}.`, "El pagar\u00e9 debe ser el monto m\u00e1s la comisi\u00f3n SGDC del cotizador."); }
      const sd = get("contrato", "comSDGC")?.v; if (sd) { ver("Comisi\u00f3n SDGC s/IVA = suma del calendario", rows.reduce((a, r) => a + r.com, 0), sd); ver("IVA de la comisi\u00f3n SDGC: calendario vs 16% del Anexo D", rows.reduce((a, r) => a + r.ivacom, 0), sd * 0.16, "Diferencias de centavos = redondeo mensual"); }
      const f0 = firma || rows[0].fecha; let prevD = f0, prevS = monto, recalc = 0, maxDiff = 0;
      rows.forEach(r => { const dias = dayDiff(prevD, r.fecha); const i = prevS * (tasa / 100) / 360 * dias; recalc += i; maxDiff = Math.max(maxDiff, Math.abs(i - r.int)); prevD = r.fecha; prevS = r.saldo; });
      ver(`Intereses del calendario vs rec\u00e1lculo (${tasa}% / 360 \u00d7 d\u00edas)`, rows.reduce((a, r) => a + r.int, 0), recalc, `Diferencia m\u00e1xima por pago: ${fmtMoney(maxDiff)}`);
      ver("Saldo final del calendario = 0", rows[rows.length - 1].saldo, 0);
      if (rows.some(r => r.iva === 0) && C.ivaIntereses) duda("Pagar\u00e9 / Contrato", "Columna IVA de intereses", "IVA de intereses $0.00 en todos los pagos.", "El contrato dice 'm\u00e1s el IVA que en su caso se cause'.", "\u00bfConfirmas que los intereses de este cr\u00e9dito no causan IVA?");
      const pagos = [...new Set(rows.map(r => r.pago))];
      if (pagos.length > 1) { const dom = C.domis?.map(d => d.monto).filter(Boolean) || []; const modo = modeOf(rows.map(r => r.pago)); if (dom.length && dom.some(m => Math.abs(m - modo) >= 0.005)) disc("Baja", "Monto de la domiciliaci\u00f3n", `Domiciliaciones: ${[...new Set(dom)].map(fmtMoney).join(", ")}.`, `Calendario: ${pagos.map(fmtMoney).join(" / ")} (el m\u00e1s frecuente es ${fmtMoney(modo)}).`, "La domiciliaci\u00f3n cita el monto del primer pago, no el de la mayor\u00eda de los pagos."); }
    }
    const tm = get("contrato", "tasaMor")?.v, to = get("contrato", "tasaOrd")?.v; if (tm && to) ver("Tasa moratoria = 1.5 \u00d7 tasa ordinaria (cl. 2.6)", tm, to * 1.5);
    if (C.anexoD && monto) { ver(`Comisi\u00f3n por investigaci\u00f3n = ${C.anexoD.pInv}% \u00d7 monto`, C.anexoD.inv, r2(monto * C.anexoD.pInv / 100)); ver(`Comisi\u00f3n por disposici\u00f3n = ${C.anexoD.pDisp}% \u00d7 monto`, C.anexoD.disp, r2(monto * C.anexoD.pDisp / 100)); if (Q.neto) ver("Neto a dispersar del cotizador = monto \u2212 (investigaci\u00f3n + disposici\u00f3n) \u00d7 1.16", Q.neto, monto - (C.anexoD.inv + C.anexoD.disp) * 1.16, "El cotizador redondea a pesos"); }
    const tq = get("cotizador", "total")?.v, tc = get("contrato", "total")?.v; if (tq && tc) ver("Total del contrato vs total del cotizador", tc, tq);
    if (CV.afiliaciones?.length && C.sucursales?.length && CV.afiliaciones.length !== C.sucursales.length) duda("Convenio / Contrato", "Notificaci\u00f3n al procesador / Anexo A", `${CV.afiliaciones.length} afiliaciones y ${C.sucursales.length} sucursales.`, "El n\u00famero de afiliaciones no coincide con el de sucursales (puede haber varias terminales por sucursal).", "\u00bfCu\u00e1ntas terminales tiene cada sucursal?");

    // cuentas
    const allCl = [...(C.clabes || []), ...(CV.clabes || [])];
    const byClabe = {};
    allCl.forEach(c => { if (!c.clabe) return; const k = c.clabe; if (!byClabe[k]) byClabe[k] = { ...c, ubics: [c.ubic], cuentas: c.cuenta ? [c.cuenta] : [] }; else { byClabe[k].ubics.push(c.ubic); if (c.cuenta) byClabe[k].cuentas.push(c.cuenta); if (/domicili|dispersi|concentradora|Dep\u00f3sito/.test(c.uso) && !/Anexo A/.test(c.uso)) byClabe[k].uso = c.uso; if (!byClabe[k].banco && c.banco) byClabe[k].banco = c.banco; } });
    const idcl = get("contrato", "idcliente")?.v || get("ta", "idcliente")?.v;
    Object.values(byClabe).forEach(c => {
      const k = checkClabe(c.clabe); const cuentas = [...new Set(c.cuentas.map(s => String(s).trim()))];
      const cont = cuentas.map(a => ({ a, ok: contained(a, k.emb) }));
      const bm = bankMatches(k.clave, c.banco);
      let obs = [];
      if (k.len !== 18) { obs.push(`Tiene ${k.len} d\u00edgitos.`); const full = Object.keys(byClabe).find(o => o.length === 18 && k.len === 17 && [...Array(18).keys()].some(i => o.slice(0, i) + o.slice(i + 1) === k.clabe)); if (full) obs.push(`Parece la CLABE ${full} (listada en el expediente) con un d\u00edgito de menos.`); } else if (!k.dvOk) obs.push("D\u00edgito verificador incorrecto.");
      if (bm === false) obs.push(`La clave ${k.clave} (${k.banco}) no corresponde al banco indicado (${c.banco}).`);
      cont.forEach(x => { if (x.ok === false) obs.push(idcl && squash(x.a) === squash(idcl) ? `La 'cuenta' ${x.a} es el n\u00famero de cliente Finsus, no la cuenta de la CLABE.` : `El n\u00famero de cuenta ${x.a} no est\u00e1 contenido en la CLABE.`); });
      report.cuentas.push({ uso: c.uso, banco: c.banco, clabe: c.clabe, cuenta: cuentas.join(" / "), ubic: [...new Set(c.ubics)].join("; "), obs: obs.join(" ") || "Sin observaciones." });
      if (k.len !== 18 || !k.dvOk) disc("Alta", "CLABE inv\u00e1lida", `${c.uso}: CLABE ${c.clabe}.`, obs.join(" "), "Una CLABE inv\u00e1lida no puede recibir ni enviar transferencias.");
      if (bm === false) disc("Alta", "Banco de la CLABE", `${c.uso}: CLABE ${c.clabe} (clave ${k.clave} = ${k.banco}).`, `El documento dice: ${c.banco}.`, "La CLABE no pertenece al banco indicado.");
      cont.forEach(x => { if (x.ok === false) { if (idcl && squash(x.a) === squash(idcl)) duda("Contrato", [...new Set(c.ubics)].join("; "), `Cuenta ${x.a} con CLABE ${c.clabe}.`, "Es el n\u00famero de cliente Finsus, no la cuenta de la CLABE.", "\u00bfLa CLABE es del acreditado o una cuenta referenciada de Finsus?"); else disc("Media", "N\u00famero de cuenta vs. CLABE", `${c.uso}: cuenta ${x.a}.`, `CLABE ${c.clabe}: cuenta embebida ${k.emb}.`, "El n\u00famero de cuenta no se encuentra en la CLABE; puede ser otra cuenta o un error de captura."); } });
    });
    const mult = C.domis?.find(d => d.multiplos);
    if (mult && mult.monto && (mult.monto % money(mult.multiplos)) !== 0) duda("Contrato", "Domiciliaciones", `Cargo de ${fmtMoney(mult.monto)} 'en m\u00faltiplos de $${mult.multiplos}'.`, "El cargo no es m\u00faltiplo de esa cantidad.", "\u00bfEs texto del formato que debe ajustarse?");
    if (C.estadoCivil && /mancomun/i.test(C.estadoCivil)) duda("Contrato", "Anexo de Datos Generales ref. 5", `Obligado solidario '${C.estadoCivil}'.`, "No s\u00e9 si la pol\u00edtica exige consentimiento del c\u00f3nyuge.", "\u00bfSe requiere firma del c\u00f3nyuge del obligado solidario?");
    if (Q.fechaToday) duda("Cotizador", "'PDF Oferta'!F2", "La fecha de la oferta es la f\u00f3rmula TODAY().", "No hay fecha fija de la cotizaci\u00f3n.", "\u00bfExiste la oferta aceptada con su fecha?");
    if (Q.alt12?.monto) report.resumen.push({ cat: "Montos y Tasas", campo: "Oferta alternativa del cotizador (no formalizada)", valor: `${fmtMoney(Q.alt12.monto)} a ${pct(Q.alt12.tasa)}%, pago ${fmtMoney(Q.alt12.pago)}`, doc: "Cotizador (Excel)", ubic: "'Oferta 12m'!D2:D9", st: "Diferencia menor", obs: "Simulaci\u00f3n descartada; no corresponde al cr\u00e9dito formalizado." });
    if (T.totalesFijos?.length) duda("Tabla de amortizaci\u00f3n (Excel)", T.totalesFijos.join(", "), "Totales de capital e intereses capturados como valor fijo.", "No se recalculan si cambia el calendario.", "\u00bfSe pueden dejar como f\u00f3rmulas?");
    if (C.procFirmaTipografica) duda("Contrato", "Contrato Billpocket (p." + C.procSigPage + ")", "El nombre del cliente aparece en tipograf\u00eda manuscrita, sin firma aut\u00f3grafa ni fecha.", "No se puede confirmar que est\u00e9 firmado.", "\u00bfExiste la constancia de firma electr\u00f3nica con fecha?");
    if (one("kyc")) { const lowK = Object.entries(one("kyc").ex.fields).filter(([, e]) => e.ocr && e.conf != null && e.conf < 70); if (lowK.length) duda("KYC", "Varios campos", `Campos le\u00eddos con OCR de baja confianza: ${lowK.map(([k]) => (FIELDS.find(f => f[0] === k) || [, , k])[2]).join(", ")}.`, "El OCR puede confundir 0/O, 1/l, 5/S.", "Confirma esos valores contra el KYC original."); }
    if (one("kyc")) duda("KYC", "Toda la hoja", "El KYC se ley\u00f3 con OCR.", "El OCR no distingue firmas ni casillas marcadas con certeza.", "Revisa que el KYC est\u00e9 firmado por el cliente y por el ejecutivo.");

    /* ---- reglas a\u00f1adidas (persona f\u00edsica, procesador, comisi\u00f3n, ventas, KYC) ---- */
    const now = opts.now ? new Date(opts.now) : new Date();
    const todayIso = new Date(now.getTime() - now.getTimezoneOffset() * 6e4).toISOString().slice(0, 10);
    const pf = !!(C.personaFisica || K.pf || CV.personaFisica);
    if (pf) report.faltantes = report.faltantes.map(f => /INE del representante/.test(f) ? "INE de la acreditada (y del obligado solidario, si lo hay)" : f).filter(f => !/Acta constitutiva/.test(f));
    const nombre = get("contrato", "razon")?.v || get("pagare", "razon")?.v || get("convenio", "razon")?.v || get("cotizador", "razon")?.v;
    // nombre incompleto (subconjunto de palabras)
    const fr = fs("razon");
    if (fr && fr.st === "No") {
      const tk = v => new Set(regimen(v).split(" ").filter(Boolean));
      const full = tk(nombre || fr.entries[0].v);
      const subs = fr.entries.filter(e => { const t = tk(e.v); return [...t].every(w => full.has(w)) && t.size < full.size; });
      const others = fr.entries.filter(e => !subs.includes(e) && [...tk(e.v)].sort().join(" ") !== [...full].sort().join(" "));
      if (subs.length && !others.length) {
        fr.st = "Diferencia menor"; covered.add("razon");
        report.resumen.forEach(r => { if (r.campo === fr.label) { r.st = "Diferencia menor"; r.obs = "Nombre incompleto en " + subs.map(e => e.doc).join(", ") + ". " + (r.obs || ""); } });
        subs.forEach(e => disc("Baja", "Nombre incompleto", `${e.doc}, ${e.ubic}: '${e.v}'.`, `Resto de documentos: '${nombre}'.`, "Falta parte del nombre de la acreditada en ese documento."));
      }
    }
    if (pf) {
      if (K.texto && nombre) { const toks = regimen(nombre).split(" ").filter(w => w.length > 2); const T = norm(K.texto); const ok = toks.filter(w => T.includes(w)).length; report.resumen.push({ cat: "Identificadores", campo: "Nombre de la acreditada en el KYC", valor: ok === toks.length ? nombre : `${ok} de ${toks.length} palabras encontradas`, doc: "KYC", ubic: "KYC p.1", st: ok === toks.length ? "S\u00ed" : "No legible", obs: ok === toks.length ? "El nombre completo aparece en el KYC (le\u00eddo con OCR)." : "El OCR no encontr\u00f3 todas las palabras del nombre: confirma en el original." }); }
      const tel = get("contrato", "telefono")?.v; if (tel && K.telefonos?.length) { const ok = K.telefonos.includes(String(tel).replace(/\D/g, "")); report.resumen.push({ cat: "Personas", campo: "Tel\u00e9fono de la acreditada en el KYC", valor: K.telefonos.join(" \u00b7 "), doc: "KYC", ubic: "KYC (tel\u00e9fonos)", st: ok ? "S\u00ed" : "No", obs: ok ? `Incluye el ${tel} del contrato.` : `El ${tel} del contrato no aparece en el KYC.` }); }
      // domicilios del contrato vs pagar\u00e9, convenio y KYC
      (C.domPF || []).forEach(d => {
        const lab = "Domicilio " + d.tipo.toLowerCase();
        report.resumen.push({ cat: "Domicilios", campo: lab, valor: addrText(d), doc: "Contrato", ubic: d.ubic, st: "", obs: "" });
        const checks = [];
        if (PG.domicilios?.length) checks.push(["Pagar\u00e9", PG.domicilios.map(x2 => x2.texto).join(" | "), PG.domicilios[0].ubic]);
        if (CV.domOct) checks.push(["Convenio de adhesi\u00f3n", CV.domOct, "Cl\u00e1usula Octava"]);
        if (K.texto) checks.push(["KYC", (K.domicilios || []).map(x2 => [x2.calle, x2.colonia, x2.cp].join(" ")).join(" | ") || K.texto, (K.domicilios || [])[0]?.ubic || "KYC"]);
        const res = checks.map(([doc, txt, ub]) => ({ doc, ok: addrIn(d, txt), ub, txt }));
        res.forEach(r => report.resumen.push({ cat: "Domicilios", campo: lab, valor: r.ok ? "Aparece" : "No aparece", doc: r.doc, ubic: r.ub, st: r.ok ? "S\u00ed" : (r.doc === "Convenio de adhesi\u00f3n" ? "Diferencia menor" : "No"), obs: r.ok ? "" : (r.doc === "Convenio de adhesi\u00f3n" ? "El convenio solo se\u00f1ala un domicilio para notificaciones." : `El ${lab} del contrato (${addrText(d)}) no se encontr\u00f3 en ${r.doc}.`) }));
        const status = res.some(r => !r.ok && r.doc !== "Convenio de adhesi\u00f3n") ? "No" : "S\u00ed";
        report.resumen.forEach(r => { if (r.campo === lab && r.doc === "Contrato" && r.st === "") r.st = status; });
        res.filter(r => !r.ok && r.doc !== "Convenio de adhesi\u00f3n").forEach(r => disc("Media", `${lab} en ${r.doc}`, `Contrato, ${d.ubic}: ${addrText(d)}.`, r.doc === "KYC" ? `KYC: ${(K.domicilios || []).length ? (K.domicilios || []).map(x2 => x2.tipo + ": " + [x2.calle, x2.colonia, x2.cp].filter(Boolean).join(", ")).join("; ") : "no se encontr\u00f3 el domicilio"}.` : `${r.doc}: ${r.txt}.`, `El ${lab} del contrato no aparece en ${r.doc === "KYC" ? "el KYC" : "el " + r.doc.toLowerCase()}.`));
      });
      (K.domicilios || []).forEach(d => report.resumen.push({ cat: "Domicilios", campo: "Domicilio en KYC (" + d.tipo + ")", valor: [d.calle, d.colonia, d.cp ? "C.P. " + d.cp : null, d.anios ? d.anios + " a\u00f1os" : null].filter(Boolean).join(", "), doc: "KYC", ubic: d.ubic, st: "Solo en un documento", obs: "Le\u00eddo con OCR." }));
    }
    if (C.sinObligado) duda("Contrato / Pagar\u00e9", "Proemio del contrato y firma del pagar\u00e9", "El obligado solidario es 'N/A' y el pagar\u00e9 no tiene aval.", PG.conAval === false ? "El texto del pagar\u00e9 sigue mencionando al 'Aval'." : "No s\u00e9 si la pol\u00edtica permite este cr\u00e9dito sin aval.", "\u00bfEl cr\u00e9dito se autoriz\u00f3 sin obligado solidario?");
    [C.procesadorVacio && ["Contrato", C.procesadorVacio], CV.procesadorVacio && ["Convenio de adhesi\u00f3n", CV.procesadorVacio]].filter(Boolean).forEach(([doc, pv]) => disc("Alta", "Procesador de pagos sin nombre", `${doc}, ${pv.ubic}: fecha ${pv.fecha} con 'Nombre del procesador' en blanco.`, `${get("contrato", "firmanteProcesador") || C.procSigPage ? "El expediente s\u00ed trae el contrato del procesador" : "No se identific\u00f3 el contrato del procesador"}${CV.notifEncabezado ? "; notificaci\u00f3n: '" + CV.notifEncabezado + "'" : ""}.`, "Los anexos de garant\u00eda deben identificar al procesador cuya cobranza se cede.", "procesador"));
    if (CV.notifSinProcesador) disc("Alta", "Notificaci\u00f3n al procesador sin nombre", `Convenio, ${CV.notifSinProcesador}: el cuerpo omite el nombre del procesador (p. ej. 'instruir a para que\u2026').`, CV.notifEncabezado ? `Encabezado: '${CV.notifEncabezado}'.` : "Sin encabezado identificable.", "La notificaci\u00f3n es la que instruye al procesador a depositar en el fideicomiso.");
    // comisi\u00f3n SDGC: porcentaje escrito en el Anexo D
    const sdC = get("contrato", "comSDGC")?.v;
    const promQ = Q.ventas?.length ? Q.ventas.reduce((a, v) => a + v.v, 0) / Q.ventas.length : Q.promedio;
    const baseQ = (Q.plazoBase || 0) + (Q.soloInteres || 0);
    const sIVA = Q.sinIVA, rateS = sIVA?.rate?.v != null ? +sIVA.rate.v : null, montoS = sIVA?.monto?.v != null ? +sIVA.monto.v : null;
    if (C.anexoD?.pSDGC && rateS != null && montoS != null) {
      // Regla acordada (15/11/2026): el Anexo D se llena con el bloque "Cotizador s/IVA" del cotizador (tasa redondeada y su monto); se compara contra \u00e9l, no contra la tasa c/IVA \u00f7 1.16
      const pctS = r2(rateS * 100 * 100) / 100;
      ver(`% SDGC del Anexo D = tasa del bloque 'Cotizador s/IVA' (${sIVA.rate.ubic})`, C.anexoD.pSDGC, pctS, "Regla acordada");
      ver(`Monto SDGC del Anexo D = 'Comisi\u00f3n SGDC Monto' del bloque s/IVA (${sIVA.monto.ubic})`, sdC, r2(montoS), "Regla acordada");
      if (Math.abs(C.anexoD.pSDGC - pctS) > 0.005) disc("Media", "Porcentaje de la comisi\u00f3n SDGC (Anexo D vs. cotizador s/IVA)", `${C.anexoD.ubic}: ${C.anexoD.pSDGC}%.`, `Cotizador, bloque s/IVA (${sIVA.rate.ubic}): ${pctS}%.`, "El porcentaje del Anexo D debe ser el del bloque sin IVA del cotizador.");
      if (sdC && Math.abs(sdC - montoS) > 1) disc("Media", "Monto de la comisi\u00f3n SDGC (Anexo D vs. cotizador s/IVA)", `${C.anexoD.ubic}: ${fmtMoney(sdC)}.`, `Cotizador, bloque s/IVA (${sIVA.monto.ubic}): ${fmtMoney(montoS)}.`, "El monto del Anexo D debe ser el del bloque sin IVA del cotizador.");
    } else if (C.anexoD?.pSDGC && promQ && sdC && baseQ) {
      const lit = C.anexoD.pSDGC / 100 * promQ * baseQ;
      ver(`Comisi\u00f3n SDGC s/IVA = ${C.anexoD.pSDGC}% (Anexo D) \u00d7 promedio \u00d7 ${baseQ}`, sdC, lit, "Porcentaje tal como est\u00e1 escrito en el Anexo D");
      if (Math.abs(lit - sdC) > 1) disc("Media", "Porcentaje de la comisi\u00f3n SDGC (Anexo D)", `${C.anexoD.ubic}: '${C.anexoD.pSDGC}% del promedio hist\u00f3rico\u2026' = ${fmtMoney(sdC)}.`, `${C.anexoD.pSDGC}% \u00d7 promedio ${fmtMoney(promQ)} \u00d7 ${baseQ} = ${fmtMoney(lit)}.${Q.sdgcRate ? " El cotizador usa " + (Q.sdgcRate * 100).toFixed(3).replace(/0+$/, "") + "% c/IVA \u00f7 1.16 = " + (Q.sdgcRate / 1.16 * 100).toFixed(4) + "%" + (Q.sdgcFormula ? "" : " (capturado a mano)") + "." : ""}`, "El porcentaje escrito no reproduce el monto; adem\u00e1s el texto no menciona el multiplicador de meses.");
    }
    if (Q.mult != null && Math.abs(Q.mult - 2) > 1e-9) { const dec = (String(Q.mult).split(".")[1] || "").length; duda("Cotizador", "'Cotizador 18m'!D4 (Meses ventas)", `Monto = ${dec > 4 ? Q.mult.toFixed(4) : Q.mult} \u00d7 promedio de ventas${Q.multFormula ? "" : " (valor capturado a mano" + (dec > 4 ? " con " + dec + " decimales" : "") + ")"}.`, dec > 4 ? "Parece que el monto se fij\u00f3 primero y el factor se despej\u00f3 para llegar a esa cifra." : "En los dem\u00e1s expedientes el factor suele ser 2.", "\u00bfC\u00f3mo se determin\u00f3 el monto y qui\u00e9n autoriz\u00f3 ese factor?", true); }
    if (Q.ventas?.length) {
      const sin = Q.ventas.filter(v => v.v > 0 && v.v < 1); if (sin.length) disc("Media", "Mes sin ventas registradas en el cotizador", `Cotizador ${sin.map(v => v.ubic + " (" + v.mes.slice(0, 7) + ") = " + fmtMoney(v.v) + (v.nota ? " '" + v.nota + "'" : "")).join("; ")}.`, `El promedio de ${Q.ventas.length} meses (${fmtMoney(promQ)}) incluye ese mes; sin \u00e9l ser\u00eda ${fmtMoney(Q.ventas.filter(v => v.v >= 1).reduce((a, v) => a + v.v, 0) / Q.ventas.filter(v => v.v >= 1).length)}.`, "El monto se calcula sobre un promedio que incluye un mes sin informaci\u00f3n.");
      const fpro = parseDate(get("contrato", "fechaProcesador")?.v || get("convenio", "fechaProcesador")?.v); const firstSale = Q.ventas.find(v => v.v >= 1)?.mes;
      if (fpro && firstSale && dayDiff(firstSale, fpro) > 60) duda("Cotizador / Contrato del procesador", "Inputs!C (ventas) / Anexo B", `Ventas con tarjeta desde ${firstSale.slice(0, 7)}, pero el contrato con el procesador es del ${fmtDate(fpro)}.`, "Las ventas hist\u00f3ricas se procesaron con otro adquirente que no aparece en el Anexo B ni en la notificaci\u00f3n de cesi\u00f3n.", "\u00bfCon qu\u00e9 procesador se hicieron esas ventas? Si sigue activo, sus cobros no estar\u00edan cedidos al fideicomiso.", true);
    }
    if (Q.neto && C.anexoD && monto) { const calc = monto - ((C.anexoD.inv || 0) + (C.anexoD.disp || 0)) * 1.16; if (Math.abs(calc - Q.neto) >= 1) disc("Media", "Comisi\u00f3n de apertura: cotizador vs. Anexo D", `Cotizador 'Cotizador 18m'!D11/H18: apertura ${Q.apertura != null ? (Q.apertura / 1.16 * 100).toFixed(4) + "% s/IVA" : "\u2014"}; neto a dispersar ${fmtMoney(Q.neto)}.`, `Anexo D: ${C.anexoD.pInv}% + ${C.anexoD.pDisp}% = ${fmtMoney((C.anexoD.inv || 0) + (C.anexoD.disp || 0))} m\u00e1s IVA; neto resultante ${fmtMoney(calc)} (${fmtMoney(Q.neto - calc)} de diferencia).`, "Lo que se dispersa con las comisiones del contrato no es lo que se ofreci\u00f3 en el cotizador."); }
    if (Q.ventas?.length > 1) { const gaps = []; for (let i = 1; i < Q.ventas.length; i++) { const a = Q.ventas[i - 1].mes, b = Q.ventas[i].mes; if (!a || !b) continue; const ma = +a.slice(0, 4) * 12 + +a.slice(5, 7), mb = +b.slice(0, 4) * 12 + +b.slice(5, 7); for (let m = ma + 1; m < mb; m++) gaps.push(Math.floor((m - 1) / 12) + "-" + String((m - 1) % 12 + 1).padStart(2, "0")); } if (gaps.length) disc("Media", "Meses faltantes en las ventas del cotizador", `Cotizador Inputs: la serie salta ${gaps.join(", ")} (${Q.ventas[0].mes.slice(0, 7)} a ${Q.ventas[Q.ventas.length - 1].mes.slice(0, 7)} = ${Q.ventas.length} meses capturados).`, "El promedio de '12 meses' no corresponde a 12 meses consecutivos.", "Puede ser un mes omitido a prop\u00f3sito (bajo, cambio de procesador) o un error de captura."); }
    // Regla acordada (05/11/2026): el ejecutivo del KYC (Lic. Rub\u00e9n A. Guill\u00e9n Villarino) es dato fijo y quien certifica la entrevista cambia: no es error.
    if (T.fecha) { const d = parseDate(T.fecha), hm = (T.fecha.match(/\((\d{1,2}):(\d{2})\)/) || []); if (d) { const when = new Date(d + "T" + (hm[1] ? hm[1].padStart(2, "0") + ":" + hm[2] : "00:00") + ":00"); if (when.getTime() > now.getTime() + 5 * 6e4) duda("Tabla de amortizaci\u00f3n (Excel)", "Hoja1!A12", `Impresi\u00f3n '${T.fecha}'.`, "Es una fecha u hora posterior a la de esta revisi\u00f3n.", "\u00bfLa TA se gener\u00f3 con una fecha de sistema distinta? \u00bfEs la versi\u00f3n definitiva?"); } }
    if (firma && dayDiff(firma, todayIso) > 15) duda("Contrato / Pagar\u00e9 / Convenio", "Fecha de firma", `Los documentos est\u00e1n fechados el ${fmtDate(firma)} (hace ${Math.round(dayDiff(firma, todayIso))} d\u00edas).`, "Esta herramienta no puede ver si un PDF trae firmas.", "\u00bfSon las versiones firmadas? Si no, p\u00eddelas.", true);
    const rows0 = PG.rows?.length ? PG.rows : (C.rowsF || []);
    if (rows0.length && C.domis?.length) { const p1 = rows0[0].pago, dm = [...new Set(C.domis.map(d => d.monto).filter(Boolean))]; if (dm.length && dm.every(m => Math.abs(m - p1) >= 0.005) && [...new Set(rows0.map(r => r.pago))].length > 1 && !report.discrepancias.some(d2 => d2.campo === "Monto de la domiciliaci\u00f3n")) disc("Baja", "Pago 1 vs monto domiciliado", `Calendario: pago 1 = ${fmtMoney(p1)}; dem\u00e1s pagos ${fmtMoney(modeOf(rows0.map(r => r.pago)))}.`, `Domiciliaciones: ${dm.map(fmtMoney).join(", ")}.`, "El monto domiciliado no coincide con el primer pago."); }
    if (PG.hojasTexto && /^\d+ Hoja$/i.test(PG.hojasTexto) && +PG.hojasTexto.split(" ")[0] > 1) disc("Baja", "Hojas del pagar\u00e9", `Pagar\u00e9: 'se suscribe en ${PG.hojasTexto}'.`, "Deber\u00eda decir 'Hojas'.", "Error de redacci\u00f3n.");

    // discrepancias de campos no cubiertos
    const PRIO = { "Montos y Tasas": "Alta", Personas: "Alta", Fechas: "Media", Domicilios: "Media", Identificadores: "Media", Sucursales: "Media" };
    Object.entries(fieldStatus).forEach(([key, s]) => {
      if (covered.has(key)) return;
      if (s.st === "No") {
        const groups = {}; s.entries.forEach(e => { const k = String(canon(s.type, e.v)); (groups[k] = groups[k] || []).push(e); });
        const G = Object.values(groups).sort((a, b) => b.length - a.length);
        if (key === "fechaFirma" && fbf && G.length === 2 && G[1].length === 1 && G[1][0].ubic === fbf.ubic) return;
        const describe = g => g.map(e => `${e.doc}, ${e.ubic}`).join("; ") + `: ${display(s.type, g[0].v)}`;
        disc(key === "plazo" ? "Media" : (PRIO[s.cat] || "Media"), s.label, describe(G.slice(1).flat().length ? G.slice(1).flat() : G[0]).replace(/^/, ""), describe(G[0]), s.cat === "Montos y Tasas" ? "Las cifras deben ser iguales en todos los documentos." : "El dato no coincide entre documentos.");
      } else if (s.st === "No legible") {
        duda(s.entries.map(e => e.doc).join(" / "), s.entries.map(e => e.ubic).join("; "), s.entries.map(e => `${e.doc}: ${display(s.type, e.v)}`).join(" \u00b7 "), "Uno de los valores sali\u00f3 de OCR o de texto da\u00f1ado y no coincide.", `Confirma ${s.label.toLowerCase()} contra el original.`);
      }
    });
    const rank = { Alta: 0, Media: 1, Baja: 2 }; report.discrepancias.sort((a, b) => rank[a.prio] - rank[b.prio]);

    report.excel = { sinIVA: Q.sinIVA || null, calendario: rows, calSrc: rowsSrc, cotInt: Q.intRows || [], tasa, firma, monto, anexoD: C.anexoD || null, sumaPrincipal: sp ?? null, montoPagareConv: CV.montoPagare?.v ?? null, totalContrato: get("contrato", "total")?.v ?? null, totalCot: get("cotizador", "total")?.v ?? null, neto: Q.neto ?? null, mult: Q.mult ?? null, sdgcRate: Q.sdgcRate ?? null, base: (Q.plazoBase || 0) + (Q.soloInteres || 0) || null, ret: get("convenio", "retencion")?.v ?? get("cotizador", "retencion")?.v ?? null, tasaMor: get("contrato", "tasaMor")?.v ?? null, ventasMin: (K.ventasAnuales && isFinite(K.ventasAnuales.min) && K.ventasAnuales.min > 0) ? K.ventasAnuales : null, pf };
    // encabezado
    const razon = get("contrato", "razon")?.v || get("pagare", "razon")?.v || get("kyc", "razon")?.v || get("cotizador", "razon")?.v || "Expediente";
    const idexp = get("cotizador", "idExpediente")?.v;
    report.expediente = [razon, idexp ? "Expediente " + idexp : null, cred ? "Cr\u00e9dito " + cred : null].filter(Boolean).join(" \u00b7 ");
    report.datos = { monto, total: get("contrato", "total")?.v, plazo: plazoC, pago: get("ta", "pagoMensual")?.v || get("pagare", "pagoMensual")?.v, tasa: to, cat: get("contrato", "cat")?.v, rfc: get("contrato", "rfc")?.v, constitucion: fc, representante: rep, procesador: get("contrato", "procesador")?.v, retencion: get("convenio", "retencion")?.v ?? get("cotizador", "retencion")?.v, firma };
    const nNo = Object.values(fieldStatus).filter(s => s.st === "No").length, nSi = Object.values(fieldStatus).filter(s => s.st === "S\u00ed").length;
    const top = report.discrepancias.filter(d => d.prio === "Alta").map(d => d.campo.toLowerCase());
    const nA = report.discrepancias.filter(d => d.prio === "Alta").length;
    report.conclusion = `${nSi} datos coinciden entre documentos. Se encontraron ${report.discrepancias.length} discrepancias (${nA} de prioridad alta) y ${report.dudas.length} dudas. ` + (top.length ? `Lo m\u00e1s importante: ${[...new Set(top)].slice(0, 4).join("; ")}.` : "No hay discrepancias de prioridad alta.");
    // Sello de validaci\u00f3n: se muestra en la p\u00e1gina y en la portada del Excel
    {
      const cnt = p => report.discrepancias.filter(d => d.prio === p).length;
      const al = cnt("Alta"), me = cnt("Media"), ba = cnt("Baja"), cv = report.camposVariables || [];
      const okCv = cv.filter(c => c.estado === "Poblado").length;
      const vt = opts.validadoEn ? new Date(opts.validadoEn) : new Date();
      const p2 = n => String(n).padStart(2, "0");
      const fecha = `${p2(vt.getDate())}/${p2(vt.getMonth() + 1)}/${vt.getFullYear()} ${p2(vt.getHours())}:${p2(vt.getMinutes())}`;
      const nivel = al ? "critico" : me ? "observaciones" : "ok";
      const titulo = al ? "Expediente validado \u2014 con observaciones de prioridad alta" : me ? "Expediente validado \u2014 con observaciones" : "Expediente validado \u2014 sin observaciones";
      const detalle = `Revisi\u00f3n terminada el ${fecha}. ${al} alta \u00b7 ${me} media \u00b7 ${ba} baja \u00b7 ${report.dudas.length} dudas` + (cv.length ? ` \u00b7 campos variables poblados: ${okCv} de ${cv.length}` : "") + ".";
      report.validacion = { fecha, nivel, titulo, detalle, alta: al, media: me, baja: ba, dudas: report.dudas.length, camposOk: okCv, camposTotal: cv.length };
    }
    return report;
  };

  /* ============ procesamiento de un documento le\u00eddo ============ */
  E.process = function (doc) {
    doc.kind = E.classify(doc);
    const fn = { contrato: E.extractContrato, pagare: E.extractPagare, convenio: E.extractConvenio, kyc: E.extractKyc, csf: E.extractCsf, ta: E.extractTA, cotizador: E.extractCotizador, captura: E.extractCaptura }[doc.kind];
    doc.ex = fn ? fn(doc) : { fields: {}, x: {} };
    return doc;
  };
  E.needsOcr = function (doc) {
    // p\u00e1ginas a pasar por OCR. Documento escaneado completo: todas (m\u00e1x. 12). Documento con texto: de cada bloque de
    // p\u00e1ginas escaneadas s\u00f3lo las 3 \u00faltimas (comisiones y firmas del contrato del procesador) y las firmas con texto da\u00f1ado.
    if (doc.type !== "pdf") return [];
    const emp = doc.pages.map(p => { const t = pageFlat(p); return (t.match(/[A-Za-z\u00c1\u00c9\u00cd\u00d3\u00da\u00d1\u00e1\u00e9\u00ed\u00f3\u00fa\u00f1]/g) || []).length < 40; });
    if (emp.every(Boolean)) return doc.pages.slice(0, 12).map(p => p.no);
    const out = new Set();
    for (let i = 0; i < emp.length;) { if (!emp[i]) { i++; continue; } let j = i; while (j < emp.length && emp[j]) j++; for (let k = Math.max(i, j - 3); k < j; k++) out.add(k + 1); i = j; }
    doc.pages.forEach(p => { const t = pageFlat(p); if (/REPRESENTADO POR/i.test(t) && garbledRatio(t) > 0.004) out.add(p.no); });
    return [...out].sort((a, b) => a - b);
  };

  /* ============ agrupar documentos por expediente ============ */
  // Une documentos que comparten n\u00famero de cr\u00e9dito, ID del expediente (nombre de archivo o cotizador), RFC o nombre del acreditado.
  E.docKeys = function (doc) {
    const f = doc.ex?.fields || {}, x = doc.ex?.x || {};
    const base = String(doc.name || "").replace(/^[0-9a-f]{8}-/i, "");
    const id = (base.match(/(?:^|[^\d])(\d{5})(?:[^\d]|$)/) || [])[1] || (f.idExpediente?.v ? String(f.idExpediente.v) : null);
    const nm = [f.razon?.v, ...((f.razon?.more) || []).map(m => m.v)].filter(Boolean).map(v => regimen(v).split(" ").filter(w => w.length > 1).sort().join(" ")).filter(Boolean);
    return { credito: f.credito?.v ? squash(f.credito.v) : null, id, rfc: f.rfc?.v ? squash(f.rfc.v).toUpperCase() : null, names: [...new Set(nm)], texto: x.texto ? norm(x.texto) : null };
  };
  E.groupDocs = function (docs) {
    const n = docs.length, par = docs.map((_, i) => i);
    const find = i => par[i] === i ? i : (par[i] = find(par[i]));
    const join = (a, b) => { a = find(a); b = find(b); if (a !== b) par[b] = a; };
    const K = docs.map(E.docKeys);
    const jac = (a, b) => { const A = new Set(a.split(" ")), B = new Set(b.split(" ")); const inter = [...A].filter(w => B.has(w)).length; return inter / (A.size + B.size - inter); };
    for (let i = 0; i < n; i++) for (let j = i + 1; j < n; j++) {
      const a = K[i], b = K[j];
      if ((a.credito && a.credito === b.credito) || (a.id && a.id === b.id) || (a.rfc && a.rfc === b.rfc) || a.names.some(x => b.names.some(y => x === y || jac(x, y) >= 0.75))) { join(i, j); continue; }
      const t = a.texto || b.texto, other = a.texto ? b : a;
      if (t && other.names.some(nm => nm.split(" ").filter(w => w.length > 2).every(w => t.includes(w)))) join(i, j);
    }
    const groups = {}; docs.forEach((d, i) => { const r = find(i); (groups[r] = groups[r] || []).push(d); });
    return Object.values(groups);
  };

  if (typeof module !== "undefined" && module.exports) module.exports = E; else root.CruceEngine = E;
})(typeof self !== "undefined" ? self : this);
