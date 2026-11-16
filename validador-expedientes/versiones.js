/* Comparaci\u00f3n entre versiones de un mismo expediente.
   - Instant\u00e1nea por documento: huella del archivo, l\u00edneas de texto de p\u00e1ginas digitales, imagen de p\u00e1ginas escaneadas
     (900 px de ancho) con las palabras del OCR, y celdas de Excel (f\u00f3rmula o valor).
   - Las versiones se guardan solo en este navegador (IndexedDB) para comparar la siguiente entrega.
   Uso en la p\u00e1gina: CruceVersiones.* ; las funciones de texto y celdas tambi\u00e9n corren en Node para pruebas. */
(function (root) {
  "use strict";
  const V = {};
  const norm = s => String(s == null ? "" : s).replace(/\s+/g, " ").trim();
  const MAXDET = 250;

  /* ---------- huella ---------- */
  V.hash = async function (buf) {
    try { const h = await crypto.subtle.digest("SHA-256", buf); return [...new Uint8Array(h)].map(b => b.toString(16).padStart(2, "0")).join(""); }
    catch (e) { const u = new Uint8Array(buf); let a = 0x811c9dc5; for (let i = 0; i < u.length; i++) { a ^= u[i]; a = Math.imul(a, 16777619) >>> 0; } return "f" + a.toString(16) + "-" + u.length; }
  };

  /* ---------- instant\u00e1neas ---------- */
  V.THUMB_W = 900;
  V.canvasToThumb = function (canvas) {
    const w = V.THUMB_W, h = Math.round(canvas.height * w / canvas.width);
    const c = document.createElement("canvas"); c.width = w; c.height = h; const x = c.getContext("2d");
    x.fillStyle = "#fff"; x.fillRect(0, 0, w, h); x.imageSmoothingQuality = "high"; x.drawImage(canvas, 0, 0, w, h);
    return c.toDataURL("image/jpeg", 0.85);
  };
  V.renderThumb = async function (pdfPage) {
    const vp1 = pdfPage.getViewport({ scale: 1 }); const vp = pdfPage.getViewport({ scale: V.THUMB_W / vp1.width });
    const c = document.createElement("canvas"); c.width = Math.round(vp.width); c.height = Math.round(vp.height);
    const x = c.getContext("2d"); x.fillStyle = "#fff"; x.fillRect(0, 0, c.width, c.height);
    await pdfPage.render({ canvasContext: x, viewport: vp }).promise;
    return c.toDataURL("image/jpeg", 0.85);
  };
  // palabras del OCR normalizadas a 0..1 para ubicarlas en la zona que cambi\u00f3
  V.ocrWords = (items, W, H) => (items || []).map(i => ({ s: i.s, x0: i.x0 / W, x1: i.x1 / W, y: i.y / H }));
  V.excelCells = function (wb) {
    const out = {}; let n = 0;
    for (const sn of wb.SheetNames) {
      const ws = wb.Sheets[sn]; const m = {};
      for (const a in ws) { if (a[0] === "!") continue; const c = ws[a]; const v = c.v instanceof Date ? c.v.toISOString().slice(0, 10) : c.v; m[a] = [c.f ? "=" + c.f : null, v == null ? null : v]; if (++n > 40000) break; }
      out[sn] = m;
    }
    return out;
  };

  /* ---------- diferencias de texto ---------- */
  function lineDiff(A, B) {
    // devuelve bloques {i (\u00edndice en A), j (\u00edndice en B), del:[...], add:[...]}
    let s = 0; while (s < A.length && s < B.length && A[s] === B[s]) s++;
    let ea = A.length, eb = B.length; while (ea > s && eb > s && A[ea - 1] === B[eb - 1]) { ea--; eb--; }
    const a = A.slice(s, ea), b = B.slice(s, eb), n = a.length, m = b.length;
    if (!n && !m) return [];
    if (!n || !m || n * m > 4e6) {
      if (n * m > 4e6) { const SB = new Set(b), SA = new Set(a); return [{ i: s, j: s, del: a.filter(x => !SB.has(x)), add: b.filter(x => !SA.has(x)) }].filter(h => h.del.length || h.add.length); }
      return [{ i: s, j: s, del: a, add: b }];
    }
    const L = Array.from({ length: n + 1 }, () => new Uint16Array(m + 1));
    for (let i = n - 1; i >= 0; i--) for (let j = m - 1; j >= 0; j--) L[i][j] = a[i] === b[j] ? L[i + 1][j + 1] + 1 : Math.max(L[i + 1][j], L[i][j + 1]);
    const hunks = []; let i = 0, j = 0, cur = null;
    const flush = () => { if (cur && (cur.del.length || cur.add.length)) hunks.push(cur); cur = null; };
    while (i < n || j < m) {
      if (i < n && j < m && a[i] === b[j]) { flush(); i++; j++; continue; }
      if (!cur) cur = { i: s + i, j: s + j, del: [], add: [] };
      if (j < m && (i >= n || L[i][j + 1] >= L[i + 1][j])) cur.add.push(b[j++]); else cur.del.push(a[i++]);
    }
    flush();
    return hunks;
  }
  V.lineDiff = lineDiff;
  // recorta lo com\u00fan entre dos textos para mostrar solo la parte que cambi\u00f3 con un poco de contexto
  function focus(a, b) {
    const wa = a.split(" "), wb = b.split(" ");
    let p = 0; while (p < wa.length && p < wb.length && wa[p] === wb[p]) p++;
    let q = 0; while (q < wa.length - p && q < wb.length - p && wa[wa.length - 1 - q] === wb[wb.length - 1 - q]) q++;
    const ctx = 6, from = Math.max(0, p - ctx);
    const cut = (w) => { const to = Math.min(w.length, w.length - q + ctx); return (from > 0 ? "\u2026 " : "") + w.slice(from, to).join(" ") + (to < w.length ? " \u2026" : ""); };
    return [cut(wa), cut(wb)];
  }
  const alnum = arr => arr.join(" ").replace(/[^0-9A-Za-z\u00c1\u00c9\u00cd\u00d3\u00da\u00d1\u00e1\u00e9\u00ed\u00f3\u00fa\u00f1]/g, "").toLowerCase();
  function pushText(out, doc, ubic, del, add) {
    if (alnum(del) === alnum(add)) return; // solo cambi\u00f3 el orden de s\u00edmbolos o espacios
    if (del.length === add.length && del.length) del.forEach((d, k) => { const [x, y] = focus(d, add[k]); out.push({ doc, ubic, antes: x, despues: y }); });
    else out.push({ doc, ubic, antes: del.length ? del.join(" / ") : "(no exist\u00eda)", despues: add.length ? add.join(" / ") : "(se elimin\u00f3)" });
  }
  function diffText(label, a, b, out) {
    const tp = p => p && p.t ? p.t.map(norm).filter(Boolean) : null;
    if (a.pages.length === b.pages.length) {
      a.pages.forEach((pa, k) => {
        const A = tp(pa), B = tp(b.pages[k]); if (!A || !B) return;
        lineDiff(A, B).forEach(h => pushText(out, label, "p." + (k + 1), h.del, h.add));
      });
      return;
    }
    const flat = d => { const L = [], P = []; d.pages.forEach((p, k) => (tp(p) || []).forEach(t => { L.push(t); P.push(k + 1); })); return { L, P }; };
    const A = flat(a), B = flat(b);
    lineDiff(A.L, B.L).forEach(h => pushText(out, label, "p." + (B.P[h.j] || A.P[h.i] || "?") + " (versi\u00f3n nueva)", h.del, h.add));
  }
  V.diffText = diffText;

  /* ---------- diferencias de celdas (Excel) ---------- */
  function diffCells(label, a, b, out) {
    const show = c => c == null ? "(vac\u00eda)" : c[0] ? c[0] + (c[1] != null ? "  \u2192 " + c[1] : "") : String(c[1]);
    const sheets = [...new Set([...Object.keys(a), ...Object.keys(b)])];
    for (const sn of sheets) {
      if (!a[sn]) { out.push({ doc: label, ubic: "Hoja '" + sn + "'", antes: "(no exist\u00eda)", despues: "Hoja nueva" }); continue; }
      if (!b[sn]) { out.push({ doc: label, ubic: "Hoja '" + sn + "'", antes: "Hoja", despues: "(se elimin\u00f3)" }); continue; }
      const addrs = [...new Set([...Object.keys(a[sn]), ...Object.keys(b[sn])])];
      for (const ad of addrs) {
        const x = a[sn][ad], y = b[sn][ad];
        if (x && y && x[0] && x[0] === y[0]) continue; // misma f\u00f3rmula: cambia solo el resultado
        const vx = x ? (x[0] || x[1]) : null, vy = y ? (y[0] || y[1]) : null;
        if (String(vx ?? "") === String(vy ?? "")) continue;
        if (typeof vx === "number" && typeof vy === "number" && Math.abs(vx - vy) < 1e-9) continue;
        out.push({ doc: label, ubic: "'" + sn + "'!" + ad, antes: show(x), despues: show(y) });
      }
    }
  }
  V.diffCells = diffCells;

  /* ---------- diferencias de imagen (p\u00e1ginas escaneadas) ---------- */
  const loadImg = src => new Promise((res, rej) => { const im = new Image(); im.onload = () => res(im); im.onerror = rej; im.src = src; });
  function gray(im, w, h) { const c = document.createElement("canvas"); c.width = w; c.height = h; const x = c.getContext("2d"); x.drawImage(im, 0, 0, w, h); const d = x.getImageData(0, 0, w, h).data; const g = new Uint8Array(w * h); for (let i = 0; i < g.length; i++) g[i] = (d[i * 4] * 3 + d[i * 4 + 1] * 6 + d[i * 4 + 2]) / 10; return { g, c }; }
  function crop(canvas, x0, y0, x1, y1) { const w = x1 - x0, h = y1 - y0; const c = document.createElement("canvas"); const sc = Math.min(1.5, 440 / w); c.width = Math.round(w * sc); c.height = Math.round(h * sc); const x = c.getContext("2d"); x.fillStyle = "#fff"; x.fillRect(0, 0, c.width, c.height); x.drawImage(canvas, x0, y0, w, h, 0, 0, c.width, c.height); return { url: c.toDataURL("image/png"), w: c.width, h: c.height }; }
  const zona = (cx, cy, W, H) => (cy < H / 3 ? "parte superior" : cy < 2 * H / 3 ? "parte media" : "parte inferior") + ", " + (cx < W / 3 ? "izquierda" : cx < 2 * W / 3 ? "centro" : "derecha");
  const wordsIn = (words, bx) => (words || []).filter(w => w.x1 >= bx.x0 && w.x0 <= bx.x1 && w.y >= bx.y0 && w.y <= bx.y1).sort((p, q) => (p.y - q.y) || (p.x0 - q.x0)).map(w => w.s).join(" ").replace(/\s+/g, " ").trim();
  V.diffImage = async function (label, pageNo, pa, pb, out) {
    const [ia, ib] = await Promise.all([loadImg(pa.img), loadImg(pb.img)]);
    const W = V.THUMB_W, Ha = Math.round(ia.height * W / ia.width), Hb = Math.round(ib.height * W / ib.width);
    if (Math.abs(Ha - Hb) / Ha > 0.03) { const A = gray(ia, W, Ha), B = gray(ib, W, Hb); const ca = crop(A.c, 0, 0, W, Ha), cbb = crop(B.c, 0, 0, W, Hb); out.push({ doc: label, ubic: "p." + pageNo, antes: "P\u00e1gina con otras proporciones", despues: "La p\u00e1gina fue reemplazada o re-escaneada; rev\u00edsala completa.", imgA: ca.url, imgB: cbb.url, w: ca.w, h: Math.max(ca.h, cbb.h) }); return; }
    const H = Ha, A = gray(ia, W, H), B = gray(ib, W, H), S = 10, gw = Math.ceil(W / S), gh = Math.ceil(H / S), mark = new Uint8Array(gw * gh);
    let changed = 0;
    for (let by = 0; by < gh; by++) for (let bx = 0; bx < gw; bx++) {
      let cnt = 0; for (let y = by * S; y < Math.min(H, by * S + S); y++) for (let x = bx * S; x < Math.min(W, bx * S + S); x++) { const k = y * W + x; if (Math.abs(A.g[k] - B.g[k]) > 70) cnt++; }
      if (cnt >= 6) { mark[by * gw + bx] = 1; changed++; }
    }
    if (!changed) return;
    if (changed / (gw * gh) > 0.25) { const ca = crop(A.c, 0, 0, W, H), cbb = crop(B.c, 0, 0, W, H); out.push({ doc: label, ubic: "p." + pageNo, antes: "(p\u00e1gina completa)", despues: "La p\u00e1gina cambi\u00f3 casi por completo (re-escaneo o p\u00e1gina reemplazada); rev\u00edsala completa.", imgA: ca.url, imgB: cbb.url, w: ca.w, h: ca.h }); return; }
    // agrupa bloques vecinos (con 2 bloques de holgura)
    const seen = new Uint8Array(gw * gh), boxes = [];
    for (let k = 0; k < mark.length; k++) {
      if (!mark[k] || seen[k]) continue;
      const st = [k]; seen[k] = 1; let x0 = 1e9, y0 = 1e9, x1 = -1, y1 = -1, n = 0;
      while (st.length) { const c = st.pop(), cx = c % gw, cy = (c / gw) | 0; n++; x0 = Math.min(x0, cx); x1 = Math.max(x1, cx); y0 = Math.min(y0, cy); y1 = Math.max(y1, cy);
        for (let dy = -2; dy <= 2; dy++) for (let dx = -2; dx <= 2; dx++) { const nx = cx + dx, ny = cy + dy; if (nx < 0 || ny < 0 || nx >= gw || ny >= gh) continue; const q = ny * gw + nx; if (mark[q] && !seen[q]) { seen[q] = 1; st.push(q); } } }
      boxes.push({ x0: x0 * S, y0: y0 * S, x1: Math.min(W, (x1 + 1) * S), y1: Math.min(H, (y1 + 1) * S), n });
    }
    // une zonas en el mismo rengl\u00f3n y cercanas (p. ej., dos casillas M/F)
    for (let again = true; again;) { again = false; for (let i = 0; i < boxes.length && !again; i++) for (let j = i + 1; j < boxes.length && !again; j++) { const p = boxes[i], q = boxes[j]; const vOver = Math.min(p.y1, q.y1) - Math.max(p.y0, q.y0) > -20, hGap = Math.max(p.x0, q.x0) - Math.min(p.x1, q.x1); if (vOver && hGap < 220) { boxes[i] = { x0: Math.min(p.x0, q.x0), y0: Math.min(p.y0, q.y0), x1: Math.max(p.x1, q.x1), y1: Math.max(p.y1, q.y1), n: p.n + q.n }; boxes.splice(j, 1); again = true; } } }
    boxes.sort((p, q) => q.n - p.n).slice(0, 6).sort((p, q) => p.y0 - q.y0).forEach(bx => {
      const padX = Math.max(60, (300 - (bx.x1 - bx.x0)) / 2), padY = 45;
      const x0 = Math.max(0, Math.round(bx.x0 - padX)), x1 = Math.min(W, Math.round(bx.x1 + padX)), y0 = Math.max(0, Math.round(bx.y0 - padY)), y1 = Math.min(H, Math.round(bx.y1 + padY));
      const ca = crop(A.c, x0, y0, x1, y1), cbb = crop(B.c, x0, y0, x1, y1);
      const nb = { x0: bx.x0 / W, x1: bx.x1 / W, y0: (bx.y0 - 8) / H, y1: (bx.y1 + 8) / H };
      const ta = wordsIn(pa.words, nb), tb = wordsIn(pb.words, nb);
      const txt = (ta || tb) && ta !== tb;
      out.push({ doc: label, ubic: `p.${pageNo}, ${zona((bx.x0 + bx.x1) / 2, (bx.y0 + bx.y1) / 2, W, H)}`, antes: txt ? "Texto (OCR): " + (ta || "\u2014") : "Ver imagen", despues: txt ? "Texto (OCR): " + (tb || "\u2014") : (ta && ta === tb ? "Mismo texto; cambi\u00f3 una marca, casilla, firma o sello (ver imagen)." : "Ver imagen"), imgA: ca.url, imgB: cbb.url, w: ca.w, h: ca.h });
    });
  };

  /* ---------- comparaci\u00f3n completa ---------- */
  V.compare = async function (prev, cur, LAB) {
    const res = { previa: { fecha: prev.fecha, origen: prev.origen || "versi\u00f3n guardada en este navegador" }, archivos: [], detalles: [], disc: { resueltas: [], nuevas: [], siguen: [], sinPrevia: !prev.disc }, dudas: { antes: prev.dudas ?? null, despues: cur.dudas ?? null } };
    const kinds = [...new Set([...prev.docs, ...cur.docs].map(d => d.kind))];
    for (const k of kinds) {
      const a = prev.docs.find(d => d.kind === k), b = cur.docs.find(d => d.kind === k), label = (LAB && LAB[k]) || k;
      if (!a) { res.archivos.push({ doc: label, estado: "Nuevo en esta versi\u00f3n", antes: "\u2014", despues: b.name }); continue; }
      if (!b) { res.archivos.push({ doc: label, estado: "No se envi\u00f3 en esta versi\u00f3n", antes: a.name, despues: "\u2014" }); continue; }
      if (a.hash === b.hash) { res.archivos.push({ doc: label, estado: "Id\u00e9ntico", antes: a.name, despues: b.name }); continue; }
      const before = res.detalles.length;
      try {
        if (a.cells && b.cells) diffCells(label, a.cells, b.cells, res.detalles);
        else if (a.pages && b.pages) {
          if (a.pages.length !== b.pages.length && !/\.docx$/i.test(b.name || "")) res.detalles.push({ doc: label, ubic: "Documento", antes: a.pages.length + " p\u00e1ginas", despues: b.pages.length + " p\u00e1ginas" });
          diffText(label, a, b, res.detalles);
          if (a.pages.length === b.pages.length) for (let i = 0; i < a.pages.length; i++) {
            const pa = a.pages[i], pb = b.pages[i];
            if (pa.img && pb.img) { if (pa.img !== pb.img) await V.diffImage(label, i + 1, pa, pb, res.detalles); }
            else if (!!pa.img !== !!pb.img) res.detalles.push({ doc: label, ubic: "p." + (i + 1), antes: pa.img ? "P\u00e1gina escaneada" : "P\u00e1gina digital", despues: pb.img ? "P\u00e1gina escaneada" : "P\u00e1gina digital" });
            if (res.detalles.length > MAXDET) break;
          }
        }
      } catch (e) { res.detalles.push({ doc: label, ubic: "\u2014", antes: "", despues: "No se pudo comparar este archivo (" + (e && e.message ? e.message : e) + ")." }); }
      const n = res.detalles.length - before;
      res.archivos.push({ doc: label, estado: n ? `Modificado (${n} cambio${n === 1 ? "" : "s"})` : "Archivo distinto, mismo contenido", antes: a.name, despues: b.name });
      if (res.detalles.length > MAXDET) { res.detalles.length = MAXDET; res.detalles.push({ doc: label, ubic: "\u2014", antes: "", despues: "Hay m\u00e1s cambios; se muestran los primeros " + MAXDET + "." }); break; }
    }
    if (prev.disc) {
      const keyOf = d => norm(d.campo).toLowerCase();
      const P = new Map(prev.disc.map(d => [keyOf(d), d])), C = new Map((cur.disc || []).map(d => [keyOf(d), d]));
      P.forEach((d, k) => (C.has(k) ? res.disc.siguen : res.disc.resueltas).push({ prio: (C.get(k) || d).prio, campo: d.campo }));
      C.forEach((d, k) => { if (!P.has(k)) res.disc.nuevas.push({ prio: d.prio, campo: d.campo }); });
    }
    return res;
  };
  V.sig = snap => snap.docs.map(d => d.kind + ":" + d.hash).sort().join("|");

  /* ---------- almacenamiento local (IndexedDB) ---------- */
  const DB = "cruce-versiones", ST = "versiones";
  function idb() { return new Promise((res, rej) => { try { const r = indexedDB.open(DB, 1); r.onupgradeneeded = () => r.result.createObjectStore(ST, { keyPath: "id" }); r.onsuccess = () => res(r.result); r.onerror = () => rej(r.error); } catch (e) { rej(e); } }); }
  const tx = (db, mode, fn) => new Promise((res, rej) => { const t = db.transaction(ST, mode); const s = t.objectStore(ST); const out = fn(s); t.oncomplete = () => res(out && out.result !== undefined ? out.result : out); t.onerror = () => rej(t.error); t.onabort = () => rej(t.error); });
  V.listVersions = async function (keys) {
    try { const db = await idb(); const all = await tx(db, "readonly", s => s.getAll()); db.close(); return (all || []).filter(v => v.keys && v.keys.some(k => keys.includes(k))).sort((a, b) => String(b.fecha).localeCompare(String(a.fecha))); }
    catch (e) { return []; }
  };
  V.saveVersion = async function (keys, snap) {
    try {
      const db = await idb(); const id = keys[0] + "|" + V.sig(snap);
      await tx(db, "readwrite", s => s.put({ id, keys, fecha: snap.fecha, snap }));
      const all = await tx(db, "readonly", s => s.getAll());
      const mine = (all || []).filter(v => v.keys && v.keys.some(k => keys.includes(k))).sort((a, b) => String(b.fecha).localeCompare(String(a.fecha)));
      if (mine.length > 4) await tx(db, "readwrite", s => { mine.slice(4).forEach(v => s.delete(v.id)); });
      db.close(); return true;
    } catch (e) { return false; }
  };
  V.clearAll = async function () { try { const db = await idb(); await tx(db, "readwrite", s => s.clear()); db.close(); return true; } catch (e) { return false; } };
  V.count = async function () { try { const db = await idb(); const n = await tx(db, "readonly", s => s.count()); db.close(); return n || 0; } catch (e) { return 0; } };

  if (typeof module !== "undefined" && module.exports) module.exports = V; else root.CruceVersiones = V;
})(typeof self !== "undefined" ? self : this);
