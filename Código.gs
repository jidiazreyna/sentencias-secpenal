// ====== CONFIG ======
const FOLDER_IN_ID = "17nSFMOO2jCM3NNcic8ISRkDMQ4otROTW";
const FOLDER_OUT_ID = "1oZFdsz60t8eT2fDvMiZRFZVP9LqdlRXA";

const VOCALES = [
  "Aída Tarditti",
  "Domingo Sesín",
  "Luis Enrique Rubio",
  "María Marta Cáceres de Bollati",
  "Sebastián Cruz López Peña",
  "Jessica Valentini"
];

const DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
const GDOC_MIME = "application/vnd.google-apps.document";

// ====== GÉNERO (para “el señor/la señora Vocal …”) ======
const VOCALES_GENERO = {
  "Aída Tarditti": "F",
  "María Marta Cáceres de Bollati": "F",
  "Jessica Valentini": "F",
  "Domingo Sesín": "M",
  "Luis Enrique Rubio": "M",
  "Sebastián Cruz López Peña": "M"
};

// ====== WEB APP ======
function doGet() {
  return HtmlService.createHtmlOutputFromFile("Index")
    .setTitle("Corrector de Sentencias");
}

function listInputFiles() {
  const folder = DriveApp.getFolderById(FOLDER_IN_ID);
  const files = folder.getFiles();
  const out = [];
  while (files.hasNext()) {
    const f = files.next();
    if (f.getMimeType() === DOCX_MIME) {
      out.push({ id: f.getId(), name: f.getName(), mime: f.getMimeType() });
    }
  }
  return out;
}

// FIX #11: eliminado flujo de comparación/comentarios para acelerar proceso.

function copyBodyToCellPreserveFormat_(srcBody, dstCell) {
  const n = srcBody.getNumChildren();

  for (let i = 0; i < n; i++) {
    const el = srcBody.getChild(i);
    const type = el.getType();

    // Párrafos
    if (type === DocumentApp.ElementType.PARAGRAPH) {
      dstCell.appendParagraph(el.asParagraph().copy());
      continue;
    }

    // Ítems de lista
    if (type === DocumentApp.ElementType.LIST_ITEM) {
      dstCell.appendListItem(el.asListItem().copy());
      continue;
    }

    // Tablas
    if (type === DocumentApp.ElementType.TABLE) {
      dstCell.appendTable(el.asTable().copy());
      continue;
    }

    // Saltos / otros: los ignoramos para no romper
  }
}


function correctDocx(fileId, settings) {
  validateSettings_(settings);

  const inFile = DriveApp.getFileById(fileId);
  const outFolder = DriveApp.getFolderById(FOLDER_OUT_ID);

  const meta = Drive.Files.get(fileId);
  if (meta.mimeType !== DOCX_MIME) {
    throw new Error("Drive API detecta que NO es DOCX. MIME: " + meta.mimeType);
  }

  const changeLog = [];
  changeLog.push(makeChange_("DEBUG_STEP", "Inicio", "", "1) Validado DOCX por Drive API", {}));

  // Convertir DOCX -> Google Doc (base)
  const baseGDoc = convertDocxToGoogleDoc_(
    fileId,
    stripExt_(inFile.getName()) + "_BASE_GDoc",
    outFolder,
    changeLog
  );

  // ✅ COPIA DE TRABAJO (la que corregimos)
  const correctedGDoc = DriveApp.getFileById(baseGDoc.getId())
    .makeCopy(stripExt_(inFile.getName()) + "_CORREGIDO", outFolder);

  changeLog.push(makeChange_("DEBUG_STEP", "Copias", "", "2) Creado CORREGIDO", {
    correctedId: correctedGDoc.getId()
  }));

  // Abrir y aplicar reglas sobre CORREGIDO
  const doc = DocumentApp.openById(correctedGDoc.getId());
  applyGeneralNormalizations_(doc, changeLog);
  applyGlobalStyle_(doc, changeLog);
  applyFirstParagraphRules_(doc, settings, changeLog);
  fixSecondParagraphAbiertoElActo_(doc, changeLog);
  applyVotesLine_(doc, settings, changeLog);
  formatQuestionHeadings_(doc, changeLog);
  fixFirstQuestionIntroSentenciaI_(doc, changeLog);
  applyVotersInSections_(doc, settings, changeLog);
  fixResuelve_(doc, changeLog);

  doc.saveAndClose();
  changeLog.push(makeChange_("DEBUG_STEP", "Fin", "", "3) Guardado OK", {}));

  // ✅ Word corregido (DOCX)
  const correctedDocxFile = exportGoogleDocToDocx_(
    correctedGDoc.getId(),
    outFolder,
    stripExt_(inFile.getName()) + "_CORREGIDO"
  );

  // FIX #11: sin documentos extras de comparación/comentarios.
  return {
    correctedDocxUrl: correctedDocxFile.getUrl()
  };

}

// FIX #11: eliminado flujo de comentario/comparación en backend.

function forEachText_(element, fn) {
  const type = element.getType();
  if (type === DocumentApp.ElementType.TEXT) {
    fn(element.asText());
    return;
  }
  if (!element.getNumChildren) return;
  for (let i = 0; i < element.getNumChildren(); i++) {
    forEachText_(element.getChild(i), fn);
  }
}

function forEachTextWithDepth_(element, fn, depth) {
  const type = element.getType();
  depth = depth || 0;
  if (type === DocumentApp.ElementType.TEXT) {
    const textEl = element.asText();
    const txt = textEl.getText() || "";
    const depthArr = buildParenDepthArray_(txt, depth);
    fn(textEl, depthArr);
    return depthArr[depthArr.length - 1];
  }
  if (!element.getNumChildren) return depth;
  let d = depth;
  for (let i = 0; i < element.getNumChildren(); i++) {
    d = forEachTextWithDepth_(element.getChild(i), fn, d);
  }
  return d;
}

function applyGeneralNormalizations_(doc, log) {
  const body = doc.getBody();

  // Helper: reemplazo global (ojo: NO preserva negritas mixtas dentro del mismo Text-run,
  // pero para estas normalizaciones “simples” suele estar OK)
  const R = (pattern, repl) => body.replaceText(pattern, repl);

  // =========================
  // A0) Sr./Sra. -> señor/señora (preserva formato)
  // =========================
  const cSr = normalizeSrSra_(doc);

  // =========================
  // A1) Lic. -> licenciado/licenciada según contexto (preserva formato)
  // =========================
  const cLic = normalizeLicenciadoConditional_(doc);


  // =========================
  // A) Dr./Dra. -> doctor/doctora (y plurales)
  // =========================
  // FIX #8: conversión robusta sin dejar "doctor." colgando.
  const cDr = normalizeDoctorTitles_(doc);

  R("\\bDoctor\\b", "doctor");
  R("\\bDoctora\\b", "doctora");
  R("\\bDoctores\\b", "doctores");
  R("\\bDoctoras\\b", "doctoras");

  // =========================
  // B) Variantes de número -> n° (sin lookahead)
  // =========================
  const cNGrado = normalizeNumeroSymbolNGrado_(doc); // (debe ser la versión sin setText)

  // Si aparece "n °" suelto por conversión rara
  R("\\bn\\s*°\\b", "n°");

  // =========================
  // C) Si antes viene sentencia/auto/decreto/resolución/dictamen -> Capitalizar + "n°"
  // =========================
  R("\\bsentencia\\s+n°\\b", "Sentencia n°");
  R("\\bauto\\s+n°\\b", "Auto n°");
  R("\\bdecreto\\s+n°\\b", "Decreto n°");
  R("\\bresoluci[oó]n\\s+n°\\b", "Resolución n°");
  R("\\bdictamen\\s+n°\\b", "Dictamen n°");

  R("\\bSENTENCIA\\s+n°\\b", "Sentencia n°");
  R("\\bAUTO\\s+n°\\b", "Auto n°");
  R("\\bDECRETO\\s+n°\\b", "Decreto n°");
  R("\\bRESOLUCI[ÓO]N\\s+n°\\b", "Resolución n°");
  R("\\bDICTAMEN\\s+n°\\b", "Dictamen n°");

  // =========================
  // D) "Sala Penal" siempre así
  // =========================
  R("\\bsala\\s+penal\\b", "Sala Penal");
  R("\\bSala\\s+penal\\b", "Sala Penal");

  // =========================
  // E) Siglas sin puntos (CSJN/TSJ/CP/CN/CPP)
  // =========================
  R("\\bC\\s*\\.\\s*S\\s*\\.\\s*J\\s*\\.\\s*N\\s*\\.", "CSJN");
  R("\\bC\\s*\\.\\s*S\\s*\\.\\s*J\\s*\\.\\s*N\\b", "CSJN");

  R("\\bT\\s*\\.\\s*S\\s*\\.\\s*J\\s*\\.", "TSJ");
  R("\\bT\\s*\\.\\s*S\\s*\\.\\s*J\\b", "TSJ");

  R("\\bC\\s*\\.\\s*P\\s*\\.", "CP");
  R("\\bC\\s*\\.\\s*P\\b", "CP");

  R("\\bC\\s*\\.\\s*N\\s*\\.", "CN");
  R("\\bC\\s*\\.\\s*N\\b", "CN");

  R("\\bC\\s*\\.\\s*P\\s*\\.\\s*P\\s*\\.", "CPP");
  R("\\bC\\s*\\.\\s*P\\s*\\.\\s*P\\b", "CPP");
  // FIX #9: quitar punto residual de CPP.
  R("\\bCPP\\.", "CPP");

  // FIX #6: TSJ Córdoba/de Córdoba/de la Provincia de Córdoba -> TSJ.
  const cTSJ = normalizeTSJCordoba_(doc);

  // FIX #4: A./S. no X -> A./S. n° X
  const cANo = normalizeANoNumero_(doc);

  // FIX #5: Ley 456 -> Ley n° 456
  const cLey = normalizeLeyNumero_(doc);

  // FIX #7: Fiscal sólo mayúscula en "Fiscal General".
  const cFiscal = normalizeFiscalCase_(doc);

  // =========================
  // F) Artículos del/de la + siglas CP/CPP/CN/CSJN/TSJ
  // (IMPORTANTE: versiones reescritas con replaceInTextPreserveStyle_ para no romper negritas)
  // =========================
  const cArtOut = normalizeArticlesOutsideParens_(doc); // versión sin setText
  const cArtIn  = normalizeArticlesInsideParens_(doc);  // versión sin setText

  // =========================
  // G) Fechas numéricas -> d/m/yyyy con "/"
  // (IMPORTANTE: usar versión sin setText, con replaceInTextPreserveStyle_)
  // =========================
  const cFechas = normalizeNumericDates_(doc);

  // =========================
  // H) Nuevas normalizaciones (todas deben estar reescritas sin setText)
  // =========================
  const cDec = normalizeDecimoCompuestos_(doc);
  const cMP  = normalizeMinisterioPublico_(doc);
  const cOrd = normalizeOrdinalesAbreviados_(doc);
  const cTri = normalizeTribunalCase_(doc);
  const cVoc = normalizeVocalConditional_(doc);
  // FIX #2 y #3: latinismos en cursiva + vgr/vrg -> v. gr.
  const cLat = italicizeLatinisms_(doc);
  // FIX #14: preservar formato "(SAC ####)" sin "n°".
  const cSAC = normalizeSACOpening_(doc);

  // =========================
  // LOGS
  // =========================
  if (log) {
    log.push(makeChange_(
      "GENERAL_NORMALIZATIONS",
      "Documento completo",
      "",
      `Aplicadas normalizaciones generales. n°=${cNGrado || 0}; artOut=${cArtOut || 0}; artIn=${cArtIn || 0}; fechas=${cFechas || 0}.`,
      {}
    ));

    log.push(makeChange_(
      "NEW_NORMALIZATIONS",
      "Documento completo",
      "",
      `Aplicadas: decimoComp=${cDec || 0}, MP=${cMP || 0}, ordinales=${cOrd || 0}, tribunal=${cTri || 0}, vocal=${cVoc || 0}, sr/sra=${cSr||0}; lic=${cLic||0}; dr=${cDr||0}; TSJ=${cTSJ||0}; A/S no=${cANo||0}; Ley=${cLey||0}; fiscal=${cFiscal||0}; latin=${cLat||0}; SAC=${cSAC||0}.`,
      {}
    ));
  }
}

function normalizeDoctorTitles_(doc) {
  const body = doc.getBody();
  let touched = 0;
  const DELIM = "(\\s|$|[\\.,;:!\\?\\)\\]\\u00BB\\u201D])";
  const rules = [
    [new RegExp("\\bDras?\\.?" + DELIM, "ig"), (m) => (/^dras/i.test(m[0]) ? `doctoras${m[1]||""}` : `doctora${m[1]||""}`)],
    [new RegExp("\\bDrs\\.?" + DELIM, "ig"), (m) => `doctores${m[1]||""}`],
    [new RegExp("\\bDr\\.?" + DELIM, "ig"), (m) => `doctor${m[1]||""}`]
  ];
  forEachText_(body, (textEl) => rules.forEach(([re, fn]) => touched += replaceInTextPreserveStyle_(textEl, re, fn)));
  return touched;
}


function replaceInTextPreserveStyle_(textEl, regex, makeReplacement) {
  const s = textEl.getText() || "";
  if (!s) return 0;

  // RE2 en Apps Script: NO lookahead/lookbehind. Regex normal con flags g/i.
  const flags = (regex.ignoreCase ? "i" : "") + "g";
  const rx = new RegExp(regex.source, flags);

  let m;
  const matches = [];

  while ((m = rx.exec(s)) !== null) {
    const start = m.index;
    const end = start + m[0].length - 1;

    const repl = (typeof makeReplacement === "function")
      ? makeReplacement(m, s)
      : String(makeReplacement);

    if (repl !== m[0]) matches.push({ start, end, repl });
    if (rx.lastIndex === m.index) rx.lastIndex++; // safety
  }

  if (!matches.length) return 0;

  // Aplicar de atrás hacia adelante para no desfasar índices
  for (let i = matches.length - 1; i >= 0; i--) {
    const { start, end, repl } = matches[i];
    try {
      const attrs = textEl.getAttributes(start); // copia estilo del inicio del match
      textEl.deleteText(start, end);
      textEl.insertText(start, repl);
      textEl.setAttributes(start, start + repl.length - 1, attrs); // lo aplica al texto insertado
    } catch (e) {}
  }

  return matches.length;
}

function italicizeLatinisms_(doc) {
  const body = doc.getBody();
  let touched = 0;

  // FIX #2: vgr/vrg -> v. gr. (en cursiva)
  forEachText_(body, (textEl) => {
    touched += replaceInTextPreserveStyle_(textEl, /\b(?:vgr|vrg)\b/ig, "v. gr.");
  });

  // FIX #2 y #3: latinismos/locuciones en cursiva preservando formato existente.
  const terms = [
    "v\\.\\s*gr\\.", "in\\s+re", "in\\s+dubio\\s+pro\\s+reo", "bis", "ter", "quater", "quinquies", "sexies", "septies", "octies", "novies", "nonies", "decies",
    "a\\s+quo", "ad\\s+quem", "onus\\s+probandi", "res\\s+iudicata", "habeas\\s+corpus", "ex\\s+lege", "dura\\s+lex", "sed\\s+lex",
    "non\\s+bis\\s+in\\s+idem", "ad\\s+effectum\\s+videndi", "sine\\s+qua\\s+non", "prima\\s+facie", "ut\\s+supra", "supra", "modus\\s+operandi", "animus\\s+domini", "animus", "ad\\s+hoc"
  ];

  const rx = new RegExp("\\b(?:" + terms.join("|") + ")\\b", "ig");
  forEachText_(body, (textEl) => {
    const s = textEl.getText() || "";
    if (!s) return;
    const matches = [];
    let m;
    rx.lastIndex = 0;
    while ((m = rx.exec(s)) !== null) {
      matches.push({ start: m.index, end: m.index + m[0].length - 1 });
      if (rx.lastIndex === m.index) rx.lastIndex++;
    }
    for (const it of matches) {
      let allItalic = true;
      for (let i = it.start; i <= it.end; i++) {
        const a = textEl.getAttributes(i) || {};
        if (a.ITALIC !== true) { allItalic = false; break; }
      }
      if (!allItalic) {
        try { textEl.setItalic(it.start, it.end, true); touched++; } catch (e) {}
      }
    }
  });

  return touched;
}

function normalizeDecimoCompuestos_(doc) {
  const body = doc.getBody();
  let touched = 0;

  const rules = [
    [/\bDecimo\s+Primera\b/ig, "Decimoprimera"],
    [/\bDécimo\s+Primera\b/ig, "Decimoprimera"],
    [/\bDecimo\s+Segunda\b/ig, "Decimosegunda"],
    [/\bDécimo\s+Segunda\b/ig, "Decimosegunda"],
    [/\bDecimo\s+Tercera\b/ig, "Decimotercera"],
    [/\bDécimo\s+Tercera\b/ig, "Decimotercera"],
    [/\bDecimo\s+Cuarta\b/ig, "Decimocuarta"],
    [/\bDécimo\s+Cuarta\b/ig, "Decimocuarta"],
    [/\bDecimo\s+Quinta\b/ig, "Decimoquinta"],
    [/\bDécimo\s+Quinta\b/ig, "Decimoquinta"],
    [/\bDecimo\s+Sexta\b/ig, "Decimosexta"],
    [/\bDécimo\s+Sexta\b/ig, "Decimosexta"],
    [/\bDecimo\s+Séptima\b/ig, "Decimoséptima"],
    [/\bDécimo\s+Séptima\b/ig, "Decimoséptima"],
    [/\bDecimo\s+Septima\b/ig, "Decimoséptima"],
    [/\bDécimo\s+Septima\b/ig, "Decimoséptima"],
    [/\bDecimo\s+Octava\b/ig, "Decimoctava"],
    [/\bDécimo\s+Octava\b/ig, "Decimoctava"],
    [/\bDecimo\s+Novena\b/ig, "Decimonovena"],
    [/\bDécimo\s+Novena\b/ig, "Decimonovena"],
    [/\bVigésimo\s+Primera\b/ig, "Vigesimoprimera"],
    [/\bVigésimo\s+Segunda\b/ig, "Vigesimosegunda"],
  ];

  forEachText_(body, (textEl) => {
    rules.forEach(([re, repl]) => {
      touched += replaceInTextPreserveStyle_(textEl, re, repl);
    });
  });

  return touched;
}


function normalizeMinisterioPublico_(doc) {
  const body = doc.getBody();
  let touched = 0;

  const reDef = /\bministerio\s+p[uú]blico\s+de\s+la\s+defensa\b/ig;
  const reMPF = /\bministerio\s+p[uú]blico\s+fiscal\b/ig;
  const reMPFF = /\bministerio\s+p[uú]blico\s+fiscal\s+fiscal\b/ig;
  const reMP  = /\bministerio\s+p[uú]blico\b/ig;

  forEachText_(body, (textEl) => {
    // 1) Defensa primero (así no termina convertido a Fiscal)
    touched += replaceInTextPreserveStyle_(textEl, reDef, "Ministerio Público de la Defensa");

    // 2) MP Fiscal explícito
    touched += replaceInTextPreserveStyle_(textEl, reMPF, "Ministerio Público Fiscal");
    // FIX #13: evita duplicación "Fiscal Fiscal".
    touched += replaceInTextPreserveStyle_(textEl, reMPFF, "Ministerio Público Fiscal");

    // 3) “Ministerio Público” suelto -> Fiscal, salvo que en el texto inmediato diga “de la Defensa”
    touched += replaceInTextPreserveStyle_(textEl, reMP, (m, full) => {
      const idx = m.index;
      const tail = (full || "").slice(idx, idx + 60).toLowerCase();
      if (/\bministerio\s+p[uú]blico\s+de\s+la\s+defensa\b/i.test(tail)) return m[0]; // dejar
      return "Ministerio Público Fiscal";
    });
  });

  return touched;
}

function normalizeTSJCordoba_(doc) {
  const body = doc.getBody();
  let touched = 0;
  const rules = [
    /\bTSJ\s+de\s+la\s+Provincia\s+de\s+C[óo]rdoba\b/ig,
    /\bTSJ\s+de\s+C[óo]rdoba\b/ig,
    /\bTSJ\s+C[óo]rdoba\b/ig
  ];
  forEachText_(body, (textEl) => rules.forEach((re) => touched += replaceInTextPreserveStyle_(textEl, re, "TSJ")));
  return touched;
}

function normalizeANoNumero_(doc) {
  const body = doc.getBody();
  let touched = 0;
  const re = /\b([AS]\.)\s*(?:n[°º]?|no|nro)\.?\s*(\d+)\b/ig;
  forEachText_(body, (textEl) => touched += replaceInTextPreserveStyle_(textEl, re, (m) => `${m[1]} n° ${m[2]}`));
  return touched;
}

function normalizeLeyNumero_(doc) {
  const body = doc.getBody();
  let touched = 0;
  const re = /\b[Ll]ey\s+(\d+)\b/g;
  forEachText_(body, (textEl) => touched += replaceInTextPreserveStyle_(textEl, re, (m) => `Ley n° ${m[1]}`));
  return touched;
}

function normalizeFiscalCase_(doc) {
  const body = doc.getBody();
  let touched = 0;
  const re = /\bFiscal\b/g;
  forEachText_(body, (textEl) => {
    touched += replaceInTextPreserveStyle_(textEl, re, (m, full) => {
      const tail = (full || "").slice(m.index, m.index + 30);
      return /^Fiscal\s+General\b/.test(tail) ? "Fiscal" : "fiscal";
    });
  });
  return touched;
}

function normalizeSACOpening_(doc) {
  const body = doc.getBody();
  let touched = 0;
  const re = /\(\s*SAC\s+n\s*[°º]\s*(\d+)\s*\)/ig;
  forEachText_(body, (textEl) => touched += replaceInTextPreserveStyle_(textEl, re, (m) => `(SAC ${m[1]})`));
  return touched;
}


function normalizeOrdinalesAbreviados_(doc) {
  const body = doc.getBody();
  let touched = 0;

  const re = /\b([1-9]\d{0,2})\s*(ro|do|to|mo|no|vo)\.?\b/ig;

  forEachText_(body, (textEl) => {
    touched += replaceInTextPreserveStyle_(textEl, re, (m) => `${m[1]}°`);
  });

  return touched;
}


function normalizeTribunalCase_(doc) {
  const body = doc.getBody();
  let touched = 0;

  const reTribunal = /\bTribunal\b|\bTRIBUNAL\b/g;
  const reEste = /\beste\s+tribunal\b/ig;
  const reAlto = /\balto\s+tribunal\b/ig;
  const reTS   = /\btribunal\s+superior\b/ig;
  const reTSJ  = /\btribunal\s+superior\s+de\s+justicia\b/ig;

  forEachText_(body, (textEl) => {
    // 1) bajar “Tribunal” a minúscula (solo la palabra)
    touched += replaceInTextPreserveStyle_(textEl, reTribunal, "tribunal");

    // 2) restaurar excepciones
    touched += replaceInTextPreserveStyle_(textEl, reEste, "este Tribunal");
    touched += replaceInTextPreserveStyle_(textEl, reAlto, "Alto Tribunal");
    // TSJ antes que TS (por seguridad)
    touched += replaceInTextPreserveStyle_(textEl, reTSJ, "Tribunal Superior de Justicia");
    touched += replaceInTextPreserveStyle_(textEl, reTS, "Tribunal Superior");
  });

  return touched;
}


function normalizeVocalConditional_(doc) {
  const body = doc.getBody();
  let touched = 0;

  const rePreopinante = /\bvocal\s+preopinante\b/ig;
  const rePrimerVoto  = /\bvocal\s+del\s+primer\s+voto\b/ig;
  const reAntecede    = /\bvocal\s+que\s+antecede\b/ig;

  forEachText_(body, (textEl) => {
    // vocal doctor/a + NOMBRE (solo si el nombre está en VOCALES)
    for (const v of VOCALES) {
      const esc = v.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
      touched += replaceInTextPreserveStyle_(textEl, new RegExp(`\\bvocal\\s+doctor\\s+${esc}\\b`, "ig"), `Vocal doctor ${v}`);
      touched += replaceInTextPreserveStyle_(textEl, new RegExp(`\\bvocal\\s+doctora\\s+${esc}\\b`, "ig"), `Vocal doctora ${v}`);
    }

    // frases fijas
    touched += replaceInTextPreserveStyle_(textEl, rePreopinante, "Vocal preopinante");
    touched += replaceInTextPreserveStyle_(textEl, rePrimerVoto,  "Vocal del primer voto");
    touched += replaceInTextPreserveStyle_(textEl, reAntecede,    "Vocal que antecede");
  });

  return touched;
}



function normalizeNumeroSymbolNGrado_(doc) {
  const body = doc.getBody();

  // Captura: n.° / N.° / n° / N° / n º / N . º etc.
  // Sin lookahead: reemplazamos SOLO el match "n...°" y no tocamos lo que sigue.
  const rx = /\b[nN]\s*\.?\s*[°º]/g;

  let touched = 0;

  forEachText_(body, (textEl) => {
    const s = textEl.getText() || "";
    if (!s) return;
    if (s.search(rx) === -1) return;

    // Buscamos todos los matches con índices
    const matches = [];
    rx.lastIndex = 0;
    let m;
    while ((m = rx.exec(s)) !== null) {
      matches.push({ start: m.index, end: m.index + m[0].length - 1 });
      if (rx.lastIndex === m.index) rx.lastIndex++; // safety
    }
    if (!matches.length) return;

    // Aplicamos de atrás para adelante para no desfasar índices
    for (let i = matches.length - 1; i >= 0; i--) {
      const { start, end } = matches[i];

      try {
        // Guardamos atributos del primer char del match para mantener estilo
        const attrs = textEl.getAttributes(start);

        textEl.deleteText(start, end);
        textEl.insertText(start, "n°");

        // Restaurar estilo del reemplazo
        textEl.setAttributes(start, start + 1, attrs);

        touched++;
      } catch (e) {
        // no frenamos toda la ejecución
      }
    }
  });

  return touched;
}

function normalizeNumericDates_(doc) {
  const body = doc.getBody();
  let touched = 0;

  const re = /\b(\d{1,2})\s*[\/-]\s*(\d{1,2})\s*[\/-]\s*(\d{2,4})\b/g;

  forEachText_(body, (textEl) => {
    touched += replaceInTextPreserveStyle_(textEl, re, (m) => {
      const d = parseInt(m[1], 10);
      const mo = parseInt(m[2], 10);
      const yRaw = m[3];
      let y = parseInt(yRaw, 10);

      if (!(d >= 1 && d <= 31 && mo >= 1 && mo <= 12)) return m[0];

      if (yRaw.length === 2) y = (y <= 29) ? (2000 + y) : (1900 + y);
      else if (yRaw.length === 3) return m[0];
      else if (y < 1000 || y > 2999) return m[0];

      return `${d}/${mo}/${y}`;
    });
  });

  return touched;
}



function articleForSigla_(sigla) {
  const s = (sigla || "").toUpperCase();
  if (s === "CN" || s === "CSJN") return "de la";
  // CP / CPP / TSJ (y cualquier otra que quieras sumar)
  return "del";
}

function isInsideParens_(text, index) {
  // Determina si la posición `index` cae dentro de un paréntesis abierto no cerrado
  // (simple y suficiente para casos normales de textos jurídicos)
  let depth = 0;
  for (let i = 0; i < index; i++) {
    const ch = text[i];
    if (ch === "(") depth++;
    else if (ch === ")" && depth > 0) depth--;
  }
  return depth > 0;
}

function buildParenDepthArray_(s, startDepth) {
  s = s || "";
  const depth = new Array(s.length + 1);
  let d = startDepth || 0;
  depth[0] = d;

  for (let i = 0; i < s.length; i++) {
    const ch = s[i];
    if (ch === "(") d++;
    else if (ch === ")" && d > 0) d--;
    depth[i + 1] = d;
  }
  return depth;
}

/**
 * Dentro de paréntesis: elimina "del/de la/de los/de las" antes de CP/CPP/CN/CSJN/TSJ
 * Ej: "(art. 54 del CP)" -> "(art. 54 CP)"
 * Funciona aunque el "(" esté en un Text y el "del CP" en otro.
 */
function normalizeArticlesInsideParens_(doc) {
  const body = doc.getBody();
  let touched = 0;

  // Dentro de paréntesis: "( ... art. ... del CP ... )" -> "( ... art. ... CP ... )"
  // Como no hay lookbehind, matcheamos el tramo y lo rearmamos.
  const re = /\(([^)]*?\b(?:art\.?|arts\.?|número)\b[^)]*?)\s+(del|de la|de los|de las)\s+(CP|CPP|CN|CSJN|TSJ)\b/ig;

  forEachText_(body, (textEl) => {
    touched += replaceInTextPreserveStyle_(textEl, re, (m) => `(${m[1]} ${m[3]}`);
  });

  return touched;
}


function normalizeFirstParagraphFlow_(txt) {
  if (!txt) return txt;

  // 1) "...) . La sentencia se pronuncia con motivo del ..." -> "...), con motivo del ..."
  // (cubre "sentencia" y "resolución")
  txt = txt.replace(
    /(\)\s*)\.\s*La\s+(?:sentencia|resoluci[oó]n)\s+se\s+pronuncia\s+con\s+motivo\s+del\s+/ig,
    "$1, con motivo del "
  );

  // fallback: si no hay ")" justo antes
  txt = txt.replace(
    /\.\s*La\s+(?:sentencia|resoluci[oó]n)\s+se\s+pronuncia\s+con\s+motivo\s+del\s+/ig,
    ", con motivo del "
  );

  // 2) "J. La impugnación se presenta en contra de la/del/..." -> "J., en contra de la/del/..."
  txt = txt.replace(
    /([A-ZÁÉÍÓÚÑ])\.\s*La\s+impugnaci[oó]n\s+se\s+presenta\s+en\s+contra\s+(del|de la|de los|de las)\s+/ig,
    (m, ini, art) => `${ini}., en contra ${art.toLowerCase()} `
  );

  // 3) ". La impugnación se presenta en contra de la/del/..." -> ", en contra de la/del/..."
  // (si ya venía con coma, también)
  txt = txt.replace(
    /[.,]\s*La\s+impugnaci[oó]n\s+se\s+presenta\s+en\s+contra\s+(del|de la|de los|de las)\s+/ig,
    (m, art) => `, en contra ${art.toLowerCase()} `
  );

  // 1) Caso especial: si antes hay iniciales tipo "R.A.M." mantenemos el punto y agregamos coma.
  txt = txt.replace(
    /((?:[A-ZÁÉÍÓÚÑ]\.){2,})\s*Se\s+(?:presenta|interpone|deduce|plantea|articula|formula|promueve|dirige)\s+en\s+contra\s+de\s+/g,
    "$1, en contra de "
  );

  // 2) Caso general: reemplaza el punto por coma.
  txt = txt.replace(
    /\.\s*Se\s+(?:presenta|interpone|deduce|plantea|articula|formula|promueve|dirige)\s+en\s+contra\s+de\s+/ig,
    ", en contra de "
  );


  // Limpieza de signos/espacios
  txt = txt
    .replace(/\s+,/g, ",")
    .replace(/,\s*,/g, ", ")
    .replace(/\s{2,}/g, " ")
    .replace(/,\s*\./g, ".")
    .trim();

  return txt;
}


function normalizeArticlesOutsideParens_(doc) {
  const body = doc.getBody();

  // (art/arts/número ... ) + (opcional artículo) + SIGLA
  const rx = /\b(?:art\.?|arts\.?|número)\b[\s\S]{0,180}?\s+(?:(del|de la|de los|de las)\s+)?(CP|CPP|CN|CSJN|TSJ)\b/ig;

  let touched = 0;

  // FIX #10: usa profundidad de paréntesis acumulada para evitar tocar texto dentro de paréntesis entre runs.
  forEachTextWithDepth_(body, (textEl, depthArr) => {
    const src = textEl.getText() || "";
    if (!src) return;

    // Necesitamos iterar matches y decidir caso por caso (y saltar si está dentro de paréntesis)
    touched += replaceInTextPreserveStyle_(textEl, rx, (m, full) => {
      const matchText = m[0];
      const article = (m[1] || "").toLowerCase();
      const sigla = (m[2] || "").toUpperCase();

      // ¿Este match cae dentro de paréntesis? (usamos el índice del match dentro del string del run)
      const idx = m.index;
      if (isInsideParens_(full, idx) || (depthArr[idx] || 0) > 0) return matchText; // no tocar

      const desired = articleForSigla_(sigla); // "del" o "de la"

      if (!article) {
        // insertar artículo antes de la sigla dentro del match
        return matchText.replace(new RegExp("\\s+" + sigla + "\\b"), " " + desired + " " + sigla);
      }

      if (article !== desired) {
        // corregir artículo equivocado
        return matchText.replace(new RegExp("\\b" + article.replace(/\s+/g, "\\s+") + "\\s+" + sigla + "\\b", "i"), desired + " " + sigla);
      }

      return matchText; // ya estaba bien
    });
  }, 0);

  return touched;
}



function shareAnyoneWithLinkView_(file) {
  // file: DriveApp File
  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (e) {
    // si falla, no frenamos toda la ejecución
  }
}

function exportGoogleDocToDocx_(googleDocFileId, outFolder, outName) {
  const mime = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

  // Drive v2: exportLinks trae la URL correcta (con alt=media)
  const meta = Drive.Files.get(googleDocFileId);
  const exportUrl = meta.exportLinks && meta.exportLinks[mime];

  if (!exportUrl) {
    throw new Error("No se encontró exportLinks para DOCX. ¿Es un Google Doc real? ID=" + googleDocFileId);
  }

  // Descargar el contenido exportado (requiere auth)
  const resp = UrlFetchApp.fetch(exportUrl, {
    headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  });

  const code = resp.getResponseCode();
  if (code < 200 || code >= 300) {
    throw new Error("Falló descarga exportada (" + code + "): " + resp.getContentText());
  }

  const blob = resp.getBlob();
  blob.setName(outName.endsWith(".docx") ? outName : (outName + ".docx"));

  const docxFile = outFolder.createFile(blob);
  shareAnyoneWithLinkView_(docxFile);
  return docxFile;
}



// ====== CONVERSIÓN ESTABLE ======
function convertDocxToGoogleDoc_(fileId, title, outFolder, log) {
  try {
    const copied = Drive.Files.copy(
      { title: title, mimeType: GDOC_MIME },
      fileId,
      { convert: true }
    );

    if (copied.mimeType !== GDOC_MIME) {
      throw new Error("Resultado de conversión inesperado: " + copied.mimeType);
    }

    const gfile = DriveApp.getFileById(copied.id);
    outFolder.addFile(gfile);
    return gfile;

  } catch (e) {
    log.push(makeChange_("ERROR_CONVERT", "Conversión", "", String(e), {}));
    throw e;
  }
}

function validateSettings_(s) {
  if (!s) throw new Error("Faltan settings.");
  if (!VOCALES.includes(s.presidente)) throw new Error("Presidente inválido.");

  if (!Array.isArray(s.ordenVotos) || s.ordenVotos.length !== 3) {
    throw new Error("Debés elegir 3 vocales (orden de votos).");
  }
  s.ordenVotos.forEach(v => { if (!VOCALES.includes(v)) throw new Error("Orden inválido: " + v); });

  const uniqO = [...new Set(s.ordenVotos)];
  if (uniqO.length !== 3) throw new Error("El orden de votos no puede repetir vocales.");

  if (!s.ordenVotos.includes(s.presidente)) {
    throw new Error("La presidencia debe estar entre los 3 vocales.");
  }

  // Unificación: los “vocales” son exactamente los 3 del orden
  s.vocales = [...s.ordenVotos];
}

function applyTextFont12Times_(textEl, options) {
  if (!textEl) return;

  const opts = options || {};
  const clearInlineStyles = !!opts.clearInlineStyles;

  textEl.setFontFamily("Times New Roman");
  textEl.setFontSize(12);

  const len = (textEl.getText() || "").length;
  if (len <= 0) return;

  // Por defecto NO tocamos negrita/cursiva/subrayado para preservar formato del DOCX.
  if (clearInlineStyles) {
    textEl.setBold(0, len - 1, false);
    textEl.setItalic(0, len - 1, false);
    textEl.setUnderline(0, len - 1, false);
  }

  // Evita que sobreviva formato heredado de DOCX (p. ej. versalitas/small caps)
  if (DocumentApp.Attribute && DocumentApp.Attribute.SMALL_CAPS) {
    textEl.setAttributes(0, len - 1, {
      [DocumentApp.Attribute.SMALL_CAPS]: false
    });
  }
}


// ====== ESTILO GLOBAL ======
function applyGlobalStyle_(doc, log) {
  const body = doc.getBody();
  const n = body.getNumChildren();

  let countBody = 0;
  for (let i = 0; i < n; i++) {
    const el = body.getChild(i);
    if (el.getType() !== DocumentApp.ElementType.PARAGRAPH &&
        el.getType() !== DocumentApp.ElementType.LIST_ITEM) continue;

    const p = elementToParagraphOrListItem_(el);

    p.setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY);
    removeAllIndents_(p, el.getType()); // <-- FIX REAL (lista vs párrafo)
    p.setLineSpacing(1.5);
    p.setSpacingBefore(0);
    p.setSpacingAfter(0);

    const t = p.editAsText();
    applyTextFont12Times_(t);


    countBody++;
  }

  // También dentro de tablas
  const tables = body.getTables();
  let countTables = 0;

  for (let ti = 0; ti < tables.length; ti++) {
    const table = tables[ti];
    for (let r = 0; r < table.getNumRows(); r++) {
      const row = table.getRow(r);
      for (let c = 0; c < row.getNumCells(); c++) {
        const cell = row.getCell(c);
        for (let k = 0; k < cell.getNumChildren(); k++) {
          const el = cell.getChild(k);
          if (el.getType() !== DocumentApp.ElementType.PARAGRAPH &&
              el.getType() !== DocumentApp.ElementType.LIST_ITEM) continue;

          const p = elementToParagraphOrListItem_(el);

          p.setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY);
          removeAllIndents_(p, el.getType()); // <-- FIX REAL
          p.setLineSpacing(1.5);
          p.setSpacingBefore(0);
          p.setSpacingAfter(0);

          const t = p.editAsText();
          applyTextFont12Times_(t);


          countTables++;
        }
      }
    }
  }

  log.push(makeChange_(
    "STYLE_GLOBAL",
    "Documento completo",
    "",
    `Aplicado Times New Roman 12 + Justificado + 1,5 + sin espaciado + sin sangrías en ${countBody} párrafos del body (y también en tablas: ${countTables}).`,
    {}
  ));
}

function elementToParagraphOrListItem_(el) {
  if (el.getType() === DocumentApp.ElementType.LIST_ITEM) return el.asListItem();
  return el.asParagraph();
}

// ====== PRIMER PÁRRAFO (APERTURA) ======
function isFirstParagraphCanonical_(txt, settings) {
  const s = (txt || "")
    .replace(/[\t\u00A0]/g, " ")
    .replace(/\s+/g, " ")
    .trim();

  const hasCause = /\b(emitir[aá]|emite)\s+sentencia\s+en\s+(los\s+autos|la\s+causa)\b/i.test(s);
  const hasCaratulaQuotes = /["“”][^"“”]+["“”]/.test(s);
  const hasSac = /\(\s*SAC\s+[^)]+\)/i.test(s);
  const hasResolutionPhrase = /la\s+resoluci[oó]n\s+se\s+pronuncia/i.test(s);

  // Si no hay settings, solo validamos estructura base
  if (!settings || !settings.presidente || !Array.isArray(settings.vocales)) {
    return hasCause && hasCaratulaQuotes && hasSac && hasResolutionPhrase;
  }

  const presidente = settings.presidente;
  const otros = settings.vocales.filter(v => v !== presidente);
  if (otros.length !== 2) return false;

  const presTit = vocalTitulo_(presidente).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const v2 = otros[0].replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const v3 = otros[1].replace(/[.*+?^${}()|[\]\\]/g, "\\$&");

  // Género dinámico: 2 mujeres => "las señoras Vocales doctoras", si no => "los señores Vocales doctores"
  const gens = otros.map(n => (VOCALES_GENERO[n] || "M"));
  const allF = gens.every(g => g === "F");
  const art = allF ? "las" : "los";
  const sen = allF ? "señoras" : "señores";
  const doc = allF ? "doctoras" : "doctores";

  const hasExpectedPresidency = new RegExp(`presidida\\s+por\\s+${presTit}`, "i").test(s);

  // Aceptamos DOS formatos: "e integrada por ..." o "con asistencia de ..."
  const rxIntegrada = new RegExp(
    `integrada\\s+por\\s+${art}\\s+${sen}\\s+Vocales\\s+${doc}\\s+${v2}\\s+y\\s+${v3}`,
    "i"
  );

  const rxAsistencia = new RegExp(
    `con\\s+asistencia\\s+de\\s+${art}\\s+${sen}\\s+Vocales\\s+${doc}\\s+${v2}\\s+y\\s+${v3}`,
    "i"
  );

  const hasExpectedIntegration = rxIntegrada.test(s) || rxAsistencia.test(s);

  return hasCause && hasCaratulaQuotes && hasSac && hasResolutionPhrase && hasExpectedPresidency && hasExpectedIntegration;
}


function applyFirstParagraphRules_(doc, settings, log) {
  const body = doc.getBody();

  let found = findParagraphContaining_(body, /En la\s+(ciudad|Ciudad)\s+de\s+Córdoba/i);
  if (!found) found = findInTables_(body, /En la\s+(ciudad|Ciudad)\s+de\s+Córdoba/i);

  if (!found) {
    log.push(makeChange_("P1_RULES", "Apertura", "No encontré 'En la ciudad de Córdoba' (ni en párrafos ni en tablas).", "", {}));
    return;
  }

  const p = found.paragraph;
  const where = found.where;

  let txt = p.getText() || "";
  const beforeAll = txt;

  log.push(makeChange_("DEBUG_APERTURA", where, txt, "", {}));

  // Normalizaciones puntuales previas
  txt = txt.replace(/^En la\s+Ciudad\s+de\s+Córdoba\b\s*,?/i, "En la ciudad de Córdoba,");
  txt = txt.replace(/a los fines de dictar sentencia en los autos/gi, "emitirá sentencia en los autos");
  txt = txt.replace(
    /Abierto el acto por la señora presidenta, se informa que las cuestiones a resolver son las siguientes:/gi,
    "Las cuestiones a resolver son las siguientes:"
  );

  txt = txt.replace(/\ben contra de la sentencia\b/gi, "en contra de la Sentencia");
  txt = txt.replace(/\ben contra del auto\b/gi, "en contra del Auto");
  txt = txt.replace(/\bcontra\s+la\s+sentencia\b/gi, "en contra de la Sentencia");
  txt = txt.replace(/\bcontra\s+el\s+auto\b/gi, "en contra del Auto");

  txt = txt.replace(
    /(La\s+resoluci[oó]n\s+se\s+pronuncia\s+con\s+motivo\s+del\s+recurso[\s\S]*?)\.\s*(La\s+impugnaci[oó]n\s+se\s+presenta\s+en\s+contra\s+de\s+(la\s+Sentencia|el\s+Auto))/i,
    "$1, $2"
  );

  txt = normalizeResolucionNumeroYFechaEnLetras_(txt);
  txt = normalizeEnContraStructure_(txt);
  txt = normalizeNominacionEnLetras_(txt);
  // ✅ NUEVO: arregla ". La sentencia se pronuncia..." y ". La impugnación..."
  txt = normalizeFirstParagraphFlow_(txt);

  txt = txt.replace(
    /\bTodos\s+los\s+recursos\s+se\s+interponen\s+contra\s+(la\s+Sentencia|el\s+Auto)\b/gi,
    "Todos los recursos se interponen en contra de $1"
  );

  const esModeloLargo =
    /a los\s+.*días?.*siendo.*se constituy[oó].*Sala Penal/i.test(txt) ||
    /se constituy[oó].*audiencia pública.*Sala Penal/i.test(txt);

  const esTSJ = /Sala Penal del Tribunal Superior de Justicia/i.test(txt);

  // ✅ FIX: cubre “emite/emitirá sentencia en la causa/los autos”
  const esPlantillaCruda = /\b(emitir[aá]|emite)\s+sentencia\s+en\s+(los\s+autos|la\s+causa)\b/i.test(txt);

  if (esModeloLargo || (esTSJ && esPlantillaCruda)) {
    // Tomamos la cola desde “emite/emitirá sentencia …”
    let tail = "";
    const mTail = txt.match(/(emitir[aá]|emite)\s+sentencia[\s\S]*/i);
    if (mTail) tail = mTail[0];

    // Fallback por si algo raro:
    if (!tail) {
      const mAutos = txt.match(/(en los autos[\s\S]*)/i);
      if (mAutos) tail = "emitirá sentencia " + mAutos[1].replace(/^\s*emitirá sentencia\s*/i, "");
    }

    // Forzamos “emitirá” al reconstruir (aunque venga “emite”)
    if (tail) {
      tail = tail.replace(/^(emitir[aá]|emite)\s+sentencia\b/i, "emitirá sentencia");
    }

    const presidente = settings.presidente;
    const otros = settings.vocales.filter(v => v !== presidente); // 2 asistentes

    const presTit = vocalTitulo_(presidente);

    // ✅ Género dinámico para “Vocales doctores/doctoras”
    const gens = otros.map(n => (VOCALES_GENERO[n] || "M"));
    const allF = gens.length && gens.every(g => g === "F");
    const art = allF ? "las" : "los";
    const sen = allF ? "señoras" : "señores";
    const docu = allF ? "doctoras" : "doctores";

    // Podés elegir el estilo: “e integrada por …” o “con asistencia de …”.
    // Yo dejo “e integrada por …” como venías, pero con género dinámico:
    const integracion = `${art} ${sen} Vocales ${docu} ${joinWithY_(otros)}`;

    let nuevo =
      `En la ciudad de Córdoba, la Sala Penal del Tribunal Superior de Justicia, ` +
      `presidida por ${presTit}, con asistencia de ${integracion}, ` +
      (tail ? tail : "emitirá sentencia en los autos ");

    nuevo = nuevo.replace(/\s+,/g, ",").replace(/,\s*,/g, ", ").replace(/\s{2,}/g, " ").trim();

    p.setText(nuevo);

    // ✅ FIX NEGRITA: limpiar todo el párrafo y dejar solo carátula entre comillas
    const t = p.editAsText();
    const len = (p.getText() || "").length;
    if (len > 0) t.setBold(0, len - 1, false);
    boldAutosBetweenQuotes_(p);

    // Estilo
    p.setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY);
    removeAllIndents_(p, DocumentApp.ElementType.PARAGRAPH);
    p.setLineSpacing(1.5);
    p.setSpacingBefore(0);
    p.setSpacingAfter(0);
    p.editAsText().setFontFamily("Times New Roman").setFontSize(12);
    clearUnderline_(p);

    log.push(makeChange_("P1_RULES", where, beforeAll, nuevo, {}));
    return;
  }

  // Si NO reconstruimos, pero sí hubo normalizaciones, aplicamos texto y estilos
  if (txt !== beforeAll) {
    p.setText(txt);

    // opcional: también acá podrías limpiar negrita si querés consistencia:
    // const t = p.editAsText();
    // const len = (p.getText() || "").length;
    // if (len > 0) t.setBold(0, len - 1, false);
    // boldAutosBetweenQuotes_(p);

    p.setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY);
    removeAllIndents_(p, DocumentApp.ElementType.PARAGRAPH);
    p.setLineSpacing(1.5);
    p.setSpacingBefore(0);
    p.setSpacingAfter(0);
    p.editAsText().setFontFamily("Times New Roman").setFontSize(12);
    clearUnderline_(p);

    log.push(makeChange_("P1_RULES", where, beforeAll, txt, {
      location: found.container === "BODY"
        ? { container: "BODY", index: found.index }
        : { container: "TABLE", tablePath: found.tablePath },
      highlights: [
        { kind:"literal", text:"a los fines de dictar sentencia en los autos" },
        { kind:"regex", re:"\\b(?:Sentencia|Auto)\\s*(?:n[°ºo]\\.?|nº|n°|nro\\.?|número)\\s*([0-9]{1,4})\\b", group:1 },
        { kind:"regex", re:"\\ba\\s+los[\\s\\S]{0,140}?se\\s+constituy[oó]\\s+en\\s+audiencia\\s+p[úu]blica\\b", group:0 },
        { kind:"regex", re:"[\"“«]([^\"”»]*\\bSAC\\s*\\d+\\b[^\"”»]*)[\"”»]", group:1, onlyIfNotBold:true }
      ]
    }));
  } else {
    log.push(makeChange_("P1_RULES", where, "(sin cambios)", "(ya estaba correcto o no coincidió)", {}));
  }
}


function asistentesArticuloYTitulo_(nombres) {
  // nombres: array de 2 vocales (sin el presidente)
  const gens = (nombres || []).map(v => (VOCALES_GENERO[v] || "M"));
  const allF = gens.length && gens.every(g => g === "F");

  return allF
    ? { art: "las", senores: "señoras", doct: "doctoras" }
    : { art: "los", senores: "señores", doct: "doctores" };
}

function asistentesFrase_(nombres) {
  const a = asistentesArticuloYTitulo_(nombres);
  return `${a.art} ${a.senores} Vocales ${a.doct} ${joinWithY_(nombres)}`;
}


function ordinalFemenino_(n) {
  const map = {
    1:"Primera", 2:"Segunda", 3:"Tercera", 4:"Cuarta", 5:"Quinta",
    6:"Sexta", 7:"Séptima", 8:"Octava", 9:"Novena", 10:"Décima",
    11:"Undécima", 12:"Duodécima", 13:"Decimotercera", 14:"Decimocuarta",
    15:"Decimoquinta", 16:"Decimosexta", 17:"Decimoséptima", 18:"Decimoctava",
    19:"Decimonovena", 20:"Vigésima"
  };
  return map[n] || (numberToWordsEs_(n).replace(/^./, c => c.toUpperCase()));
}

function normalizeNominacionEnLetras_(txt) {
  // Ej: "de 3ª Nominación", "de 3a Nominacion", "de 3ra Nominación", "de 3° Nominación"
  return (txt || "").replace(
    /\b(de|del)\s+(\d{1,2})\s*(?:ª|º|°|a|A|ra|RA)?\s+Nominaci[óo]n\b/g,
    (m, prep, numStr) => {
      const n = parseInt(numStr, 10);
      return `${prep} ${ordinalFemenino_(n)} Nominación`;
    }
  );
}


function joinWithY_(names) {
  const arr = (names || []).filter(Boolean);
  if (arr.length === 0) return "";
  if (arr.length === 1) return arr[0];
  if (arr.length === 2) return `${arr[0]} y ${arr[1]}`;
  return `${arr[0]}, ${arr[1]} y ${arr[2]}`; // en tu caso siempre 3
}

function normalizeInitialsDoubleDot_(txt) {
  // R.A.M.. -> R.A.M.
  return (txt || "").replace(/([A-ZÁÉÍÓÚÑ]\.){2,}\./g, (m) => m.slice(0, -1));
}

function normalizeEnContraStructure_(txt) {
  txt = normalizeInitialsDoubleDot_(txt);

  // Buscamos "Sentencia/Auto" pero SOLO actuamos si hay ANCLA cercana antes
  const rxObj = /\b(la\s+Sentencia|el\s+Auto)\b/i;
  const mObj = txt.match(rxObj);
  if (!mObj) return txt;

  const objIndex = mObj.index;
  const objText  = mObj[1]; // "la Sentencia" / "el Auto"

  // Ventana de seguridad: miramos solo 260 caracteres antes del objeto
  const WINDOW = 260;
  const startWin = Math.max(0, objIndex - WINDOW);
  const win = txt.slice(startWin, objIndex);

  // Anclas “personales” o “colectivas” típicas (tolerantes)
  const anchorRx = new RegExp(
    [
      // personales
      "\\bdefensor\\b",
      "\\babogado\\b",
      "\\bdefensora\\b",
      "\\bdefiende\\b",
      "\\basiste\\b",
      "\\bimputad[oa]\\b",
      "\\bacusad[oa]\\b",
      "\\bencartad[oa]\\b",
      // colectivas típicas de TSJ
      "\\bTodos\\s+los\\s+recursos\\b",
      "\\bLos\\s+recursos\\b",
      "\\bEl\\s+recurso\\b",
      "\\bLa\\s+impugnaci[óo]n\\b"
    ].join("|"),
    "i"
  );

  // Si no hay ancla cerca del objeto, NO tocamos (evita el desastre que te pasó)
  if (!anchorRx.test(win)) return txt;

  // Si ya está bien ("... contra la Sentencia ..." o "... en contra de la Sentencia ..."), NO tocamos
  if (/\b(contra|en\s+contra\s+de)\s+(la\s+Sentencia|el\s+Auto)\b/i.test(win + txt.slice(objIndex, objIndex + 30))) {
    return txt;
  }

  // Patrones “malos” que queremos compactar:
  // "... imputado X. Se presenta/interpone/deduce... contra/en contra de la Sentencia ..."
  // o "... Todos los recursos ... . Se ... contra/en contra de la Sentencia ..."
  txt = txt.replace(
    /(\b(?:defensor|abogado|defensora|defiende|asiste|imputad[oa]|acusad[oa]|encartad[oa]|Todos\s+los\s+recursos|Los\s+recursos|El\s+recurso)[\s\S]{0,220}?)\.\s*(Se\s+\w+|Interpuest[oa]|Deducid[oa]|Plantead[oa]|Promovid[oa]|Formulad[oa])\s+(?:en\s+)?contra\s+de\s+(la\s+Sentencia|el\s+Auto)\b/ig,
    (m, leftPart, _bridge, obj) => {
      // aseguramos coma y estructura fija
      let L = (leftPart || "").replace(/\s+/g, " ").trim();
      L = L.replace(/[.,;:]\s*$/g, "");
      return `${L}, en contra de ${obj}`;
    }
  );

  // Variante: ". Se presenta en contra de la Sentencia ..." sin “contra de”
  txt = txt.replace(
    /(\b(?:defensor|abogado|defensora|defiende|asiste|imputad[oa]|acusad[oa]|encartad[oa]|Todos\s+los\s+recursos|Los\s+recursos|El\s+recurso)[\s\S]{0,220}?)\.\s*(Se\s+\w+|Interpuest[oa]|Deducid[oa]|Plantead[oa]|Promovid[oa]|Formulad[oa])\s+(la\s+Sentencia|el\s+Auto)\b/ig,
    (m, leftPart, _bridge, obj) => {
      let L = (leftPart || "").replace(/\s+/g, " ").trim();
      L = L.replace(/[.,;:]\s*$/g, "");
      return `${L}, en contra de ${obj}`;
    }
  );

  // ====== FIX ROBUSTO: "... X. Se presenta..." -> "... X, en contra de ..." ======
  txt = txt.replace(
    /([\s\S]{0,260}?)\.\s*(Se\s+(?:presenta|interpone|deduce|plantea|articula|formula|promueve|dirige))\s+en\s+contra\s+de\s+(la\s+Sentencia|el\s+Auto)\b/gi,
    (m, leftPart, _verb, obj) => {
      // Guardas: solo aplicamos si el tramo previo tiene señales de defensa/imputado/recursos
      const anchor = /\b(defensor|defensora|defensa|abogado|abogada|asiste|en\s+car[aá]cter\s+de|imputad[oa]|acusad[oa]|encartad[oa]|recurso|recursos)\b/i;
      if (!anchor.test(leftPart)) return m;

      let L = (leftPart || "").replace(/\s+/g, " ").trim();
      L = L.replace(/[.,;:]\s*$/g, "");
      return `${L}, en contra de ${obj}`;
    }
  );


  return txt;
}

function applyVotesLine_(doc, settings, log) {
  const body = doc.getBody();

  const desired =
    `Los señores Vocales emitirán sus votos en el siguiente orden: doctores ${joinWithY_(settings.ordenVotos)}.`;

  const votesAnyRegex =
    /^\s*Los\s+(Sres\.?|señores)\s+Vocales?\s+emitir[aá]n\s+sus\s+votos\s+en\s+el\s+siguiente\s+orden\s*:/i;

  const beforeFirstQuestionRegex = /^\s*A\s+LA\s+PRIMERA\s+CUESTI[ÓO]N\s*:?\s*$/i;

  // 1) Encontrar "A LA PRIMERA CUESTIÓN:" para insertar antes
  let insertAt = -1;
  const n = body.getNumChildren();

  for (let i = 0; i < n; i++) {
    const el = body.getChild(i);
    if (el.getType() !== DocumentApp.ElementType.PARAGRAPH &&
        el.getType() !== DocumentApp.ElementType.LIST_ITEM) continue;

    const p = elementToParagraphOrListItem_(el);
    const txt = (p.getText() || "").trim();
    if (beforeFirstQuestionRegex.test(txt)) {
      insertAt = i;
      break;
    }
  }

  if (insertAt === -1) {
    log.push(makeChange_(
      "VOTES_LINE_INSERT",
      "No insertado",
      "No encontré 'A LA PRIMERA CUESTIÓN:'",
      "No se insertó línea de votos.",
      {}
    ));
    return;
  }

  // 2) Eliminar TODAS las líneas existentes de orden de votos
  let removed = 0;
  for (let i = body.getNumChildren() - 1; i >= 0; i--) {
    const el = body.getChild(i);
    if (el.getType() !== DocumentApp.ElementType.PARAGRAPH &&
        el.getType() !== DocumentApp.ElementType.LIST_ITEM) continue;

    const p = elementToParagraphOrListItem_(el);
    const txt = (p.getText() || "").trim();
    if (votesAnyRegex.test(txt)) {
      body.removeChild(el);
      removed++;
      if (i < insertAt) insertAt--; // si borré antes, corre el índice de inserción
    }
  }

  // 3) Insertar en el lugar correcto
  const newP = body.insertParagraph(insertAt, desired);

  newP.setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY);
  removeAllIndents_(newP, DocumentApp.ElementType.PARAGRAPH);
  newP.setLineSpacing(1.5);
  newP.setSpacingBefore(0);
  newP.setSpacingAfter(0);
  newP.editAsText().setFontFamily("Times New Roman").setFontSize(12);

  log.push(makeChange_(
    "VOTES_LINE_INSERT",
    `Insertado antes de párrafo ${insertAt + 1} (removed=${removed})`,
    "",
    desired,
    { insertBeforeParagraphIndex: insertAt }
  ));
}


// ====== PÁRRAFO INTRODUCTORIO DE CUESTIONES (CANONICALIZACIÓN ROBUSTA) ======
function fixSecondParagraphAbiertoElActo_(doc, log) {
  const body = doc.getBody();
  const CANON = "Las cuestiones a resolver son las siguientes:";

  // 1) Encontrar apertura (para ubicar el “párrafo de cuestiones” cerca del inicio)
  let opening = findParagraphContaining_(body, /En la\s+(ciudad|Ciudad)\s+de\s+Córdoba/i);
  if (!opening) opening = findInTables_(body, /En la\s+(ciudad|Ciudad)\s+de\s+Córdoba/i);

  if (!opening) {
    log.push(makeChange_("QUESTIONS_INTRO", "Segundo párrafo", "No encontré apertura", "No se aplicó.", {}));
    return;
  }

  // 2) Determinar el índice del párrafo de apertura EN EL BODY (no en tablas)
  const n = body.getNumChildren();
  let openingBodyIndex = -1;

  for (let i = 0; i < n; i++) {
    const el = body.getChild(i);
    if (el.getType() !== DocumentApp.ElementType.PARAGRAPH) continue;
    const t = (el.asParagraph().getText() || "");
    if (/En la\s+(ciudad|Ciudad)\s+de\s+Córdoba/i.test(t)) { openingBodyIndex = i; break; }
  }

  // 3) Buscar el “párrafo intro” de cuestiones a partir del párrafo siguiente al de apertura.
  //    Si la apertura estaba en tabla (openingBodyIndex=-1), buscamos en todo el body.
  const start = (openingBodyIndex !== -1) ? (openingBodyIndex + 1) : 0;

  // Regex tolerante para reconocer “intro de cuestiones” (incluye MUCHAS variantes)
  const rxIntro = new RegExp(
    [
      // ya viene “Las cuestiones…”
      "^\\s*Las\\s+cuestiones\\s+a\\s+resolver\\s+son",
      // variantes con “cuestiones a resolver”, “cuestiones a decidir”, etc.
      "^\\s*(?:Seguidamente\\s*,?\\s*)?(?:se\\s+)?(?:informa|hace\\s+saber|señala|manifiesta|expone|indica).{0,120}cuestiones\\s+a\\s+(?:resolver|decidir|tratar|considerar)",
      // “Abierto el acto…” con cola
      "^\\s*Abierto\\s+el\\s+acto\\b[\\s\\S]{0,160}",
      // “A continuación…” / “Luego…” / “Acto seguido…”
      "^\\s*(?:A\\s+continuaci[óo]n|Luego|Acto\\s+seguido|Seguidamente)\\b[\\s\\S]{0,160}(?:cuestiones|puntos)\\s+a\\s+(?:resolver|decidir|tratar|considerar)",
      // “Las cuestiones a dilucidar…”
      "^\\s*Las\\s+cuestiones\\s+a\\s+(?:dilucidar|diluc\u00E1dar|tratar|considerar|decidir)\\b"
    ].join("|"),
    "i"
  );

  // Para recortar “intro largo” si trae enumeración pegada
  const rxCanonLike = /Las\s+cuestiones\s+a\s+resolver\s+son\s+(?:las\s+siguientes\s*)?:/i;

  let target = null;
  let targetIndex = -1;

  for (let i = start; i < n; i++) {
    const el = body.getChild(i);
    if (el.getType() !== DocumentApp.ElementType.PARAGRAPH) continue;

    const p = el.asParagraph();
    const raw = (p.getText() || "");
    const t = raw.trim();
    if (!t) continue;

    // Cortamos la búsqueda si ya entramos a “A LA PRIMERA CUESTIÓN” (ya es tarde)
    if (/^\s*A\s+LA\s+(PRIMERA|SEGUNDA|TERCERA)\s+CUESTI[ÓO]N\b/i.test(t)) break;

    if (rxIntro.test(t)) {
      target = p;
      targetIndex = i;
      break;
    }
  }

  if (!target) {
    log.push(makeChange_("QUESTIONS_INTRO", "Body", "No encontré párrafo intro de cuestiones", "Sin cambios", {}));
    return;
  }

  const before = target.getText() || "";

  // 4) Canonicalizar: SI trae “Las cuestiones…” con basura extra, recortamos desde ahí.
  let after = before;

  const posCanon = after.search(rxCanonLike);
  if (posCanon !== -1) {
    after = after.slice(posCanon);
    after = after.replace(rxCanonLike, CANON);
  } else {
    // Cualquier otra variante -> reemplazo total
    after = CANON;
  }

  if (after !== before) {
    target.setText(after);
  } else {
    // aunque “coincida”, igual normalizamos exactamente el texto canónico
    target.setText(CANON);
    after = CANON;
  }

  // 5) Aplicar estilo consistente
  target.setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY);
  removeAllIndents_(target, DocumentApp.ElementType.PARAGRAPH);
  target.setLineSpacing(1.5);
  target.setSpacingBefore(0);
  target.setSpacingAfter(0);
  target.editAsText().setFontFamily("Times New Roman").setFontSize(12);

  // ✅ NUEVO: dejar el texto NORMAL (sin negrita/cursiva/subrayado)
  {
    const t = target.editAsText();
    const len = (t.getText() || "").length;
    if (len > 0) {
      t.setBold(0, len - 1, false);
      t.setItalic(0, len - 1, false);
      t.setUnderline(0, len - 1, false);
    }
  }

  log.push(makeChange_("QUESTIONS_INTRO", `Body párrafo ${targetIndex + 1}`, before, after, { paragraphIndex: targetIndex }));
}


// ====== VOTANTES EN SECCIONES (FIX: detecta placeholders El/La señor/a... doctor/a...) ======
function applyVotersInSections_(doc, settings, log) {
  const body = doc.getBody();
  const o = settings.ordenVotos; // [v1, v2, v3]
  if (!o || o.length !== 3) return;

  // Caso “normal” (ya lo tenías)
  const voteLineRegexNormal =
    /^(El|La)\s+(señor|señora)\s+Vocal\s+(doctor|doctora)\s+(.+?)\s*,?\s+dijo:\s*$/i;


  // Caso plantilla con placeholders (robusto)
  const voteLineRegexPlaceholder =
    /^El\s*\/\s*La\s+(?:señor|senor)\s*\/\s*a\s+Vocal\s+doctor\s*\/\s*a\b[\s\S]*?(?:,\s*)?dijo\s*:\s*$/i;
  // FIX #1: placeholders con puntos/guiones/espacios: "........ dijo:" / "— dijo:"
  const voteLineRegexDotsPlaceholder = /^\s*(?:[\.•\-–—_\|¦│┃·…\s]{2,})\s*dijo\s*:\s*$/i;
  const voteLineRegexDijoOnly = /^\s*dijo\s*:\s*$/i;

  const sectionRegex = /^A\s+LA\s+(PRIMERA|SEGUNDA|TERCERA)\s+CUESTI[ÓO]N/i;

  const norm = (s) => (s || "")
    .replace(/[\t\u00A0]/g, " ")
    .replace(/\s+/g, " ")
    .trim();

  const n = body.getNumChildren();
  let i = 0;

  while (i < n) {
    const el = body.getChild(i);
    if (el.getType() !== DocumentApp.ElementType.PARAGRAPH &&
        el.getType() !== DocumentApp.ElementType.LIST_ITEM) { i++; continue; }

    const p = elementToParagraphOrListItem_(el);
    const t = (p.getText() || "").trim();

    if (sectionRegex.test(t)) {
      const voteParas = [];
      let j = i + 1;

      while (j < n) {
        const el2 = body.getChild(j);
        if (el2.getType() !== DocumentApp.ElementType.PARAGRAPH &&
            el2.getType() !== DocumentApp.ElementType.LIST_ITEM) { j++; continue; }

        const p2 = elementToParagraphOrListItem_(el2);
        const t2 = (p2.getText() || "").trim();

        if (sectionRegex.test(t2)) break;

        // FIX: matchea normal o placeholder
        if (voteLineRegexNormal.test(t2) || voteLineRegexPlaceholder.test(t2) || voteLineRegexDotsPlaceholder.test(t2) || voteLineRegexDijoOnly.test(t2)) {
          voteParas.push({ index: j, paragraph: p2, elementType: el2.getType(), text: t2 });
          if (voteParas.length === 3) break;
        }

        j++;
      }

      if (voteParas.length > 0) {
        for (let k = 0; k < voteParas.length; k++) {
          const vp = voteParas[k];
          const desiredName = o[Math.min(k, 2)];

          const g = vocalGenero_(desiredName);
          const newLine = `${g.art} ${g.senor} Vocal ${g.doc} ${desiredName} dijo:`;



          const before = vp.paragraph.getText() || "";

          // Reescribe si el nombre actual NO coincide con el deseado (placeholder o normal)
          if (norm(before) !== norm(newLine)) {
            vp.paragraph.setText(newLine);

            vp.paragraph.setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY);
            removeAllIndents_(vp.paragraph, vp.elementType);
            vp.paragraph.setLineSpacing(1.5);
            vp.paragraph.setSpacingBefore(0);
            vp.paragraph.setSpacingAfter(0);
            const txt = vp.paragraph.editAsText();
            applyTextFont12Times_(txt);
            txt.setBold(true);
            txt.setUnderline(true);

            log.push(makeChange_("VOTER_LINE_REWRITE", `Sección ${t} / Párrafo ${vp.index + 1}`, before, newLine, {
              voter: desiredName
            }));
          } else {
            // si ya coincide, solo formato
            vp.paragraph.setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY);
            removeAllIndents_(vp.paragraph, vp.elementType);

            const txt = vp.paragraph.editAsText();
            applyTextFont12Times_(txt);
            txt.setBold(true);
            txt.setUnderline(true);

            log.push(makeChange_("VOTER_LINE_FORMAT", `Sección ${t} / Párrafo ${vp.index + 1}`, "(ya estaba)", "Formato aplicado", {
              voter: desiredName
            }));
          }

        }
      } else {
        log.push(makeChange_(
          "VOTER_LINE_REWRITE",
          `Sección ${t}`,
          "No encontré líneas de votante (ni normal ni placeholder)",
          "Sin cambios",
          { sectionParagraphIndex: i }
        ));
      }

      i = j;
      continue;
    }

    i++;
  }
}

// ====== FORMATO "RESUELVE:" (en negrita+subrayado y en línea separada) ======
function fixResuelve_(doc, log) {
  const body = doc.getBody();
  const n = body.getNumChildren();

  // Detecta "RESUELVE:" al inicio, tolerante a espacios y a "RESUELVE :"
  const rx = /^\s*RESUELVE\s*:\s*/i;

  // Para evitar partir casos donde ya está solo (RESUELVE: y nada más)
  const rxOnly = /^\s*RESUELVE\s*:\s*$/i;

  let fixed = 0;

  for (let i = 0; i < n; i++) {
    const el = body.getChild(i);
    if (el.getType() !== DocumentApp.ElementType.PARAGRAPH &&
        el.getType() !== DocumentApp.ElementType.LIST_ITEM) continue;

    const p = elementToParagraphOrListItem_(el);
    const before = p.getText() || "";

    if (!rx.test(before)) continue;

    // Si ya está solo, solo aseguramos formato
    if (rxOnly.test(before.trim())) {
      const t = p.editAsText();
      const len = before.length;
      if (len > 0) {
        t.setBold(0, len - 1, true);
        t.setUnderline(0, len - 1, true);
      }
      fixed++;
      log.push(makeChange_("RESUELVE_FORMAT", `Párrafo ${i + 1}`, before, "Formato aplicado (ya estaba solo)", { paragraphIndex: i }));
      continue;
    }

    // Caso: "RESUELVE: texto..." -> separar en dos párrafos
    const afterText = before.replace(rx, "").trim();
    const newHeader = "RESUELVE:";

    // 1) Este párrafo queda solo con "RESUELVE:"
    p.setText(newHeader);
    p.setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY);
    removeAllIndents_(p, el.getType());
    p.setLineSpacing(1.5);
    p.setSpacingBefore(0);
    p.setSpacingAfter(0);
    p.editAsText().setFontFamily("Times New Roman").setFontSize(12);

    // Negrita + subrayado a TODO "RESUELVE:"
    {
      const t = p.editAsText();
      const len = newHeader.length;
      t.setBold(0, len - 1, true);
      t.setUnderline(0, len - 1, true);
    }

    // 2) Insertar párrafo debajo con el texto restante
    const newP = body.insertParagraph(i + 1, afterText);

    newP.setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY);
    removeAllIndents_(newP, DocumentApp.ElementType.PARAGRAPH);
    newP.setLineSpacing(1.5);
    newP.setSpacingBefore(0);
    newP.setSpacingAfter(0);
    newP.editAsText().setFontFamily("Times New Roman").setFontSize(12);

    fixed++;
    log.push(makeChange_("RESUELVE_SPLIT", `Párrafo ${i + 1}`, before, `RESUELVE: (separado) + párrafo siguiente`, {
      location: { container:"BODY", index: i },
      highlights: [
        { kind:"regex", re:"^\\s*RESUELVE\\s*:\\s*\\S+", group:0 }
      ],
      insertedParagraphIndex: i + 1
    }));


    // Saltar el párrafo recién insertado para no re-procesarlo
    i++;
  }

  if (fixed === 0) {
    log.push(makeChange_("RESUELVE_SPLIT", "Documento", "No encontré 'RESUELVE:'", "Sin cambios", {}));
  }
}


// ====== FORMATO: ENCABEZADOS "A LA PRIMERA/SEGUNDA/TERCERA CUESTION" ======
function formatQuestionHeadings_(doc, log) {
  const body = doc.getBody();

  // CUESTION/CUESTIÓN, con o sin ":", tolerante a espacios
  const headingRegex = /^\s*A\s+LA\s+(PRIMERA|SEGUNDA|TERCERA)\s+CUESTI[ÓO]N\s*:?\s*$/i;

  const n = body.getNumChildren();
  let count = 0;

  for (let i = 0; i < n; i++) {
    const el = body.getChild(i);

    if (el.getType() !== DocumentApp.ElementType.PARAGRAPH &&
        el.getType() !== DocumentApp.ElementType.LIST_ITEM) continue;

    const p = elementToParagraphOrListItem_(el);
    const txt = (p.getText() || "").trim();

    if (!headingRegex.test(txt)) continue;

    // Asegura estilo base (sin romper el resto)
    p.setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY);
    removeAllIndents_(p, el.getType());
    p.setLineSpacing(1.5);
    p.setSpacingBefore(0);
    p.setSpacingAfter(0);
    p.editAsText().setFontFamily("Times New Roman").setFontSize(12);

    // Negrita + subrayado a TODA la línea
    const t = p.editAsText();
    const len = (p.getText() || "").length;
    if (len > 0) {
      t.setBold(0, len - 1, true);
      t.setUnderline(0, len - 1, true);
    }

    count++;
    log.push(makeChange_(
      "QUESTION_HEADING_FORMAT",
      `Párrafo ${i + 1}`,
      txt,
      "Aplicado negrita + subrayado",
      { paragraphIndex: i }
    ));
  }

  if (count === 0) {
    log.push(makeChange_(
      "QUESTION_HEADING_FORMAT",
      "Documento",
      "No encontré encabezados A LA PRIMERA/SEGUNDA/TERCERA CUESTION",
      "Sin cambios",
      {}
    ));
  }
}

function findParagraphContaining_(body, regex) {
  const n = body.getNumChildren();
  for (let i = 0; i < n; i++) {
    const el = body.getChild(i);
    if (el.getType() !== DocumentApp.ElementType.PARAGRAPH &&
        el.getType() !== DocumentApp.ElementType.LIST_ITEM) continue;

    const p = elementToParagraphOrListItem_(el);
    const t = (p.getText() || "");
    if (regex.test(t)) {
      return { paragraph: p, where: `Body párrafo ${i + 1}`, index: i, container: "BODY" };
    }
  }
  return null;
}

function findInTables_(body, regex) {
  const tables = body.getTables();
  for (let ti = 0; ti < tables.length; ti++) {
    const table = tables[ti];
    for (let r = 0; r < table.getNumRows(); r++) {
      const row = table.getRow(r);
      for (let c = 0; c < row.getNumCells(); c++) {
        const cell = row.getCell(c);
        const cn = cell.getNumChildren();
        for (let k = 0; k < cn; k++) {
          const el = cell.getChild(k);
          if (el.getType() !== DocumentApp.ElementType.PARAGRAPH &&
              el.getType() !== DocumentApp.ElementType.LIST_ITEM) continue;

          const p = elementToParagraphOrListItem_(el);
          const t = (p.getText() || "");
          if (regex.test(t)) {
            return {
              paragraph: p,
              where: `Tabla ${ti + 1} Fila ${r + 1} Col ${c + 1}`,
              container: "TABLE",
              tablePath: { ti, r, c, k } // ruta exacta
            };
          }
        }
      }
    }
  }
  return null;
}



function rangeHasBold_(text, start, end) {
  const s = text.getText() || "";
  const max = Math.min(end, s.length - 1);
  for (let i = start; i <= max; i++) {
    const attrs = text.getAttributes(i);
    if (attrs && attrs.BOLD === true) return true;
  }
  return false;
}

function buildAnchorMap_(corrParas) {
  const map = {
    apertura: -1,
    cuestiones: -1,
    votos: -1,
    resuelve: -1
  };

  for (let i = 0; i < corrParas.length; i++) {
    const t = (corrParas[i].text || "");

    if (map.apertura === -1 && /En la ciudad de Córdoba/i.test(t)) map.apertura = i;
    if (map.cuestiones === -1 && /Las cuestiones a resolver son las siguientes/i.test(t)) map.cuestiones = i;
    if (map.votos === -1 && /Los señores vocales emitirán sus votos en el siguiente orden/i.test(t)) map.votos = i;
    if (map.resuelve === -1 && /^\s*RESUELVE\s*:/i.test(t)) map.resuelve = i;
  }
  return map;
}


// ====== TITULACIÓN PRESIDENTE ======
function vocalTitulo_(nombre) {
  const g = VOCALES_GENERO[nombre] || "M";
  if (g === "F") return `la señora Vocal doctora ${nombre}`;
  return `el señor Vocal doctor ${nombre}`;
}

// ====== NEGRITA EN AUTOS ENTRE COMILLAS ======
function boldAutosBetweenQuotes_(paragraph) {
  const t = paragraph.editAsText();
  const full = paragraph.getText() || "";
  if (!full) return;

  // Evita herencia de negrita en todo el párrafo: primero limpiamos formato bold.
  t.setBold(0, full.length - 1, false);

  // FIX #15: soporta comillas de apertura/cierre mezcladas.
  const openSet = ['"', '“', '«'];
  const closeSet = ['"', '”', '»'];
  let i1 = -1;
  for (let i = 0; i < full.length; i++) {
    if (openSet.indexOf(full[i]) !== -1) { i1 = i; break; }
  }
  if (i1 === -1) return;

  let i2 = -1;
  for (let j = i1 + 1; j < full.length; j++) {
    if (closeSet.indexOf(full[j]) !== -1) { i2 = j; break; }
  }
  if (i2 === -1) return;

  const start = i1 + 1;
  const end = i2 - 1;
  if (end >= start) t.setBold(start, end, true);
}


// ====== NUMERO A LETRAS EN RESOLUCIÓN (solo primer párrafo) ======
function normalizeResolucionNumeroYFechaEnLetras_(txt) {
  txt = txt.replace(/\b(Sentencia|sentencia|Auto|auto)\s*(n[°ºo]\.?|nº|n°|nro\.?|número)?\s*([0-9]{1,4})\b/g,
    (m, tipo, _, num) => {
      const T = (tipo[0].toUpperCase() + tipo.slice(1).toLowerCase());
      const w = numberToWordsEs_(parseInt(num, 10));
      return `${T} número ${w}`;
    });

  txt = txt.replace(/\b(dictad[ao] el|de fecha)\s+(\d{1,2})\/(\d{1,2})\/(\d{4})\b/gi,
    (m, pref, dd, mm, yyyy) => {
      const day = numberToWordsEs_(parseInt(dd, 10));
      const month = monthNameEs_(parseInt(mm, 10));
      const year = yearToWordsEs_(parseInt(yyyy, 10));
      const p = pref.toLowerCase().startsWith("de fecha") ? "de fecha" : "dictada el día";
      return `${p} ${day} de ${month} de ${year}`;
    });

  txt = txt.replace(/\b(dictad[ao] el día|dictad[ao] el|de fecha)\s+(\d{1,2})\s+de\s+([a-záéíóúñ]+)\s+de\s+(\d{4})\b/gi,
    (m, pref, dd, mes, yyyy) => {
      const day = numberToWordsEs_(parseInt(dd, 10));
      const year = yearToWordsEs_(parseInt(yyyy, 10));
      const p = pref.toLowerCase().startsWith("de fecha") ? "de fecha" : "dictada el día";
      return `${p} ${day} de ${mes.toLowerCase()} de ${year}`;
    });

  // ====== NUEVO: "con fecha 3 de julio de dos mil veinticuatro" (año ya en letras) ======
  txt = txt.replace(
    /\b(con\s+fecha)\s+(\d{1,2})\s+de\s+([a-záéíóúñ]+)\s+de\s+([a-záéíóúñ\s]+?)(?=[,.;)]|\s|$)/gi,
    (m, pref, dd, mes, yearWords) => {
      const day = numberToWordsEs_(parseInt(dd, 10));
      return `${pref.toLowerCase()} ${day} de ${mes.toLowerCase()} de ${yearWords.trim()}`;
    }
  );

  // ====== NUEVO: "con fecha 3 de julio de 2024" (año en números) ======
  txt = txt.replace(
    /\b(con\s+fecha)\s+(\d{1,2})\s+de\s+([a-záéíóúñ]+)\s+de\s+(\d{4})\b/gi,
    (m, pref, dd, mes, yyyy) => {
      const day = numberToWordsEs_(parseInt(dd, 10));
      const year = yearToWordsEs_(parseInt(yyyy, 10));
      return `${pref.toLowerCase()} ${day} de ${mes.toLowerCase()} de ${year}`;
    }
  );

  return txt;
}

function monthNameEs_(mm) {
  const map = {
    1: "enero", 2: "febrero", 3: "marzo", 4: "abril", 5: "mayo", 6: "junio",
    7: "julio", 8: "agosto", 9: "septiembre", 10: "octubre", 11: "noviembre", 12: "diciembre"
  };
  return map[mm] || "mes";
}

function yearToWordsEs_(y) {
  if (y === 2000) return "dos mil";
  if (y > 2000 && y < 2100) {
    const rest = y - 2000;
    if (rest === 0) return "dos mil";
    return "dos mil " + numberToWordsEs_(rest);
  }
  if (y >= 1900 && y < 2000) {
    return "mil novecientos " + numberToWordsEs_(y - 1900);
  }
  return String(y);
}

function numberToWordsEs_(n) {
  if (isNaN(n)) return "";
  if (n === 0) return "cero";

  const u = ["", "uno", "dos", "tres", "cuatro", "cinco", "seis", "siete", "ocho", "nueve"];
  const d10 = ["diez", "once", "doce", "trece", "catorce", "quince", "dieciséis", "diecisiete", "dieciocho", "diecinueve"];
  const dec = ["", "", "veinte", "treinta", "cuarenta", "cincuenta", "sesenta", "setenta", "ochenta", "noventa"];
  const cen = ["", "ciento", "doscientos", "trescientos", "cuatrocientos", "quinientos", "seiscientos", "setecientos", "ochocientos", "novecientos"];

  if (n < 10) return u[n];
  if (n < 20) return d10[n - 10];
  if (n < 30) {
    // FIX #12: asegura tildes correctas en veintidós/veintitrés/veintiséis.
    const map20 = {
      21: "veintiuno", 22: "veintidós", 23: "veintitrés", 24: "veinticuatro", 25: "veinticinco",
      26: "veintiséis", 27: "veintisiete", 28: "veintiocho", 29: "veintinueve"
    };
    return (n === 20) ? "veinte" : map20[n];
  }
  if (n < 100) {
    const t = Math.floor(n / 10);
    const r = n % 10;
    return dec[t] + (r ? " y " + u[r] : "");
  }
  if (n === 100) return "cien";
  if (n < 1000) {
    const c = Math.floor(n / 100);
    const r = n % 100;
    return cen[c] + (r ? " " + numberToWordsEs_(r) : "");
  }
  if (n < 2000) {
    const r = n - 1000;
    return "mil" + (r ? " " + numberToWordsEs_(r) : "");
  }
  if (n < 10000) {
    const m = Math.floor(n / 1000);
    const r = n % 1000;
    return numberToWordsEs_(m) + " mil" + (r ? " " + numberToWordsEs_(r) : "");
  }
  return String(n);
}

// =====================
// DIFF -> RESALTADO REAL
// =====================

function applyDiffHighlights_(targetDoc /* DocumentApp abierto */, otherDocId, opts) {
  opts = opts || {};
  const mode = opts.mode || "orig"; // "orig" o "corr"

  const Y = "#fff59d"; // amarillo (lo que sale)
  const G = "#c8f7c5"; // verde (lo que entra)

  const bodyT = targetDoc.getBody();
  const bodyO = DocumentApp.openById(otherDocId).getBody();

  const targetParas = collectTextBlocks_(bodyT);
  const otherParas  = collectTextBlocks_(bodyO);

  // anclas sobre "other" (para que el match no se cruce)
  const anchors = buildAnchorMap_(otherParas);

  const otherIndex = buildNormalizedIndex_(otherParas);

  let total = 0;

  for (let i = 0; i < targetParas.length; i++) {
    const t = targetParas[i];
    const tNorm = normalizeForMatch_(t.text);

    let j = -1;

    // MATCH FORZADO por ancla (igual que antes)
    if (/En la\s+(ciudad|Ciudad)\s+de\s+Córdoba/i.test(t.text) && anchors.apertura !== -1) {
      j = anchors.apertura;
    } else if (/Las\s+cuestiones\s+a\s+resolver\s+son\s+las\s+siguientes/i.test(t.text) && anchors.cuestiones !== -1) {
      j = anchors.cuestiones;
    } else if (/Los\s+señores\s+vocales\s+emitir[aá]n\s+sus\s+votos\s+en\s+el\s+siguiente\s+orden/i.test(t.text) && anchors.votos !== -1) {
      j = anchors.votos;
    } else if (/^\s*RESUELVE\s*:/i.test(t.text) && anchors.resuelve !== -1) {
      j = anchors.resuelve;
    } else {
      j = findBestMatchIndex_(tNorm, otherParas, otherIndex, i);
    }

    if (j === -1) continue;

    const o = otherParas[j];
    const tTxt = t.text || "";
    const oTxt = o.text || "";
    if (tTxt === oTxt) continue;

    const sim = similarityScore_(tNorm, normalizeForMatch_(oTxt));
    if (sim < 0.35) {
      total += highlightWhole_(t.el.editAsText(), mode === "orig" ? Y : G);
      continue;
    }

    // 👉 Acá está el cambio real:
    // - si target es ORIGINAL (mode orig): resalto lo que desaparece o se reemplaza (amarillo)
    // - si target es CORREGIDO (mode corr): resalto lo que aparece o se reemplaza (verde)
    const ranges = (mode === "orig")
      ? diffChangedRangesInOriginal_(tTxt, oTxt)   // “sale” desde target
      : diffInsertedRangesInCorrected_(tTxt, oTxt); // “entra” en target

    if (!ranges.length) continue;

    const text = t.el.editAsText();
    ranges.forEach(r => {
      try {
        text.setBackgroundColor(r.start, r.end, mode === "orig" ? Y : G);
        total++;
      } catch (e) {}
    });
  }

  return total;
}

function diffInsertedRangesInCorrected_(corr, orig) {
  // Queremos rangos de caracteres en el CORREGIDO que son inserciones o reemplazos
  const cTok = tokenizeWithOffsets_(corr);
  const oTok = tokenizePlain_(orig);

  const cWords = cTok.map(x => x.t);
  const oWords = oTok;

  const ops = myersDiffOps_(oWords, cWords); // OJO: a=origTokens, b=corrTokens

  const ranges = [];
  let cIndex = 0; // índice en b (corr)

  for (let k = 0; k < ops.length; k++) {
    const op = ops[k];

    if (op.type === "equal") {
      cIndex += op.count;
      continue;
    }

    if (op.type === "insert") {
      const startTok = cIndex;
      const endTok = cIndex + op.count - 1;
      if (cTok[startTok] && cTok[endTok]) {
        ranges.push({ start: cTok[startTok].start, end: cTok[endTok].end });
      }
      cIndex += op.count;
      continue;
    }

    if (op.type === "replace") {
      // en replace, lo “nuevo” está del lado inserción (corr): marcamos insCount tokens
      const ins = op.insCount || 0;
      if (ins > 0) {
        const startTok = cIndex;
        const endTok = cIndex + ins - 1;
        if (cTok[startTok] && cTok[endTok]) {
          ranges.push({ start: cTok[startTok].start, end: cTok[endTok].end });
        }
      }
      cIndex += ins;
      continue;
    }

    if (op.type === "delete") {
      // deletes existen en original, no avanzan en corregido
      continue;
    }
  }

  return mergeCloseRanges_(ranges, 2);
}


/**
 * Junta bloques de texto “comparables” en orden de lectura:
 * - Párrafos del body
 * - Párrafos dentro de tablas (en el orden en que aparecen)
 * Devuelve: [{el, text}]
 */
function collectTextBlocks_(body) {
  const blocks = [];

  const pushIfText_ = (el) => {
    const t = (el.editAsText().getText() || "");
    if (t.trim()) blocks.push({ el, text: t });
  };

  const n = body.getNumChildren();
  for (let i = 0; i < n; i++) {
    const el = body.getChild(i);

    if (el.getType() === DocumentApp.ElementType.PARAGRAPH) {
      pushIfText_(el.asParagraph());
      continue;
    }

    if (el.getType() === DocumentApp.ElementType.LIST_ITEM) {
      pushIfText_(el.asListItem());
      continue;
    }

    if (el.getType() === DocumentApp.ElementType.TABLE) {
      const table = el.asTable();
      for (let r = 0; r < table.getNumRows(); r++) {
        const row = table.getRow(r);
        for (let c = 0; c < row.getNumCells(); c++) {
          const cell = row.getCell(c);
          for (let k = 0; k < cell.getNumChildren(); k++) {
            const celEl = cell.getChild(k);
            if (celEl.getType() === DocumentApp.ElementType.PARAGRAPH) pushIfText_(celEl.asParagraph());
            if (celEl.getType() === DocumentApp.ElementType.LIST_ITEM) pushIfText_(celEl.asListItem());
          }
        }
      }
    }
  }
  return blocks;
}

function buildNormalizedIndex_(corrParas) {
  const index = {}; // token -> [idx...]
  for (let i = 0; i < corrParas.length; i++) {
    const norm = normalizeForMatch_(corrParas[i].text);
    const key = norm.slice(0, 80); // prefijo como “bucket”
    if (!index[key]) index[key] = [];
    index[key].push(i);
  }
  return index;
}

function findBestMatchIndex_(oNorm, corrParas, corrIndex, iGuess) {
  // candidatos: cerca del índice + bucket por prefijo
  const candidates = new Set();

  // ventana alrededor del índice (reduce errores)
  const W = 10;
  for (let j = Math.max(0, iGuess - W); j <= Math.min(corrParas.length - 1, iGuess + W); j++) {
    candidates.add(j);
  }

  const key = oNorm.slice(0, 80);
  (corrIndex[key] || []).forEach(j => candidates.add(j));

  let bestJ = -1;
  let best = 0;

  candidates.forEach(j => {
    const cNorm = normalizeForMatch_(corrParas[j].text);
    const s = similarityScore_(oNorm, cNorm);
    if (s > best) {
      best = s;
      bestJ = j;
    }
  });

  // umbral: si no pasa, mejor no marcar nada (evita “subrayado fantasma”)
  return (best >= 0.62) ? bestJ : -1;
}

function normalizeForMatch_(s) {
  return (s || "")
    .replace(/[\u00A0\t]+/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
}

// Similaridad simple (Jaccard de palabras)
function similarityScore_(a, b) {
  if (!a || !b) return 0;
  const A = new Set(a.split(" ").filter(Boolean));
  const B = new Set(b.split(" ").filter(Boolean));
  if (!A.size || !B.size) return 0;

  let inter = 0;
  A.forEach(x => { if (B.has(x)) inter++; });
  const union = A.size + B.size - inter;
  return union ? (inter / union) : 0;
}

function highlightWhole_(text, color) {
  const s = text.getText() || "";
  if (!s) return 0;
  try { text.setBackgroundColor(0, s.length - 1, color); return 1; } catch (e) { return 0; }
}

/**
 * Devuelve rangos [start,end] (índices en el string ORIGINAL) que representan
 * tokens “que desaparecen” o “se reemplazan” respecto del corregido.
 *
 * Importante: resalta SOLO el original, por eso evita marcar partes iguales.
 */
function diffChangedRangesInOriginal_(orig, corr) {
  const oTok = tokenizeWithOffsets_(orig);
  const cTok = tokenizePlain_(corr);

  const oWords = oTok.map(x => x.t);
  const cWords = cTok;

  const ops = myersDiffOps_(oWords, cWords); // ops sobre tokens

  // Convertir deletes + replaces a rangos de caracteres en original
  const ranges = [];
  let oIndex = 0;

  for (let k = 0; k < ops.length; k++) {
    const op = ops[k];

    if (op.type === "equal") {
      oIndex += op.count;
      continue;
    }

    if (op.type === "delete") {
      const startTok = oIndex;
      const endTok = oIndex + op.count - 1;

      const startChar = oTok[startTok].start;
      const endChar = oTok[endTok].end;

      ranges.push({ start: startChar, end: endChar });
      oIndex += op.count;
      continue;
    }

    if (op.type === "replace") {
      const startTok = oIndex;
      const endTok = oIndex + op.delCount - 1;

      const startChar = oTok[startTok].start;
      const endChar = oTok[endTok].end;

      ranges.push({ start: startChar, end: endChar });
      oIndex += op.delCount;
      continue;
    }

    if (op.type === "insert") {
      // inserciones existen solo en corregido -> no se pueden “subrayar” en original
      continue;
    }
  }

  // Unir rangos muy pegados (evita “confetti”)
  return mergeCloseRanges_(ranges, 2);
}

function mergeCloseRanges_(ranges, gap) {
  if (!ranges.length) return [];
  ranges.sort((a, b) => a.start - b.start);

  const out = [ranges[0]];
  for (let i = 1; i < ranges.length; i++) {
    const prev = out[out.length - 1];
    const cur = ranges[i];

    if (cur.start <= prev.end + gap) {
      prev.end = Math.max(prev.end, cur.end);
    } else {
      out.push(cur);
    }
  }
  return out;
}

function tokenizePlain_(s) {
  // palabras + signos importantes como tokens separados
  const out = [];
  const rx = /[A-Za-zÁÉÍÓÚÑáéíóúñ0-9]+|[“”"«».,;:()¿?¡!\-–—]/g;
  let m;
  while ((m = rx.exec(s || "")) !== null) out.push(m[0].toLowerCase());
  return out;
}

function tokenizeWithOffsets_(s) {
  const out = [];
  const rx = /[A-Za-zÁÉÍÓÚÑáéíóúñ0-9]+|[“”"«».,;:()¿?¡!\-–—]/g;
  let m;
  while ((m = rx.exec(s || "")) !== null) {
    out.push({ t: m[0].toLowerCase(), start: m.index, end: m.index + m[0].length - 1 });
  }
  return out;
}

/**
 * Myers diff (token-level). Devuelve ops compactadas:
 * equal/delete/insert/replace.
 * Implementación pensada para textos “normales” (párrafos), no libros enteros.
 */
function myersDiffOps_(a, b) {
  const N = a.length, M = b.length;
  const max = N + M;
  const v = {};
  v[1] = 0;
  const trace = [];

  for (let d = 0; d <= max; d++) {
    const vv = {};
    for (let k = -d; k <= d; k += 2) {
      let x;
      if (k === -d || (k !== d && v[k - 1] < v[k + 1])) {
        x = v[k + 1]; // down (insert)
      } else {
        x = v[k - 1] + 1; // right (delete)
      }
      let y = x - k;

      while (x < N && y < M && a[x] === b[y]) { x++; y++; }

      vv[k] = x;
      if (x >= N && y >= M) {
        trace.push(vv);
        return backtrackOps_(trace, a, b);
      }
    }
    trace.push(vv);
    Object.keys(vv).forEach(k => v[k] = vv[k]);
  }

  return backtrackOps_(trace, a, b);
}

function backtrackOps_(trace, a, b) {
  let x = a.length;
  let y = b.length;
  const ops = [];

  for (let d = trace.length - 1; d >= 0; d--) {
    const v = trace[d];
    const k = x - y;

    let prevK;
    if (k === -d || (k !== d && (v[k - 1] == null ? -1 : v[k - 1]) < (v[k + 1] == null ? -1 : v[k + 1]))) {
      prevK = k + 1; // insert
    } else {
      prevK = k - 1; // delete
    }

    const prevX = v[prevK];
    const prevY = prevX - prevK;

    while (x > prevX && y > prevY) {
      ops.push({ type: "equal" });
      x--; y--;
    }

    if (d === 0) break;

    if (x === prevX) {
      ops.push({ type: "insert" });
      y--;
    } else {
      ops.push({ type: "delete" });
      x--;
    }
  }

  ops.reverse();
  return compactOps_(ops);
}

function compactOps_(ops) {
  // primero compactar iguales/delete/insert
  const compact = [];
  let cur = null;

  const pushCur = () => { if (cur) compact.push(cur); cur = null; };

  ops.forEach(op => {
    if (!cur || cur.type !== op.type) {
      pushCur();
      cur = { type: op.type, count: 1 };
    } else {
      cur.count++;
    }
  });
  pushCur();

  // ahora convertir delete+insert adyacentes en replace (más fiel para “cambios”)
  const out = [];
  for (let i = 0; i < compact.length; i++) {
    const a = compact[i];
    const b = compact[i + 1];

    if (a && b && a.type === "delete" && b.type === "insert") {
      out.push({ type: "replace", delCount: a.count, insCount: b.count });
      i++;
      continue;
    }
    out.push(a.type === "delete" ? { type: "delete", count: a.count }
          : a.type === "insert" ? { type: "insert", count: a.count }
          : { type: "equal", count: a.count });
  }
  return out;
}

function collectFromElement_(el, out) {
  const t = el.getType();
  if (t === DocumentApp.ElementType.PARAGRAPH) {
    const p = el.asParagraph();
    const te = p.editAsText();
    out.push({ textEl: te, text: te.getText() || "" });
    return;
  }
  if (t === DocumentApp.ElementType.LIST_ITEM) {
    const li = el.asListItem();
    const te = li.editAsText();
    out.push({ textEl: te, text: te.getText() || "" });
    return;
  }
  if (t === DocumentApp.ElementType.TABLE) {
    // ya lo recorremos arriba, pero no molesta
    return;
  }
  // otros: ignorar
}

function highlightDeletionsAndReplacements_(textEl, originalStr, correctedStr) {
  // colores: eliminaciones/reemplazos (rojo claro)
  const RED = "#ffd6d6";

  const A = tokenizeWords_(originalStr);
  const B = tokenizeWords_(correctedStr);

  const ops = myersDiff_(A.map(x => x.w), B.map(x => x.w));

  // ops: [{type:'equal'|'delete'|'insert', a0,a1,b0,b1}]
  // resaltar deletes en el original
  let marks = 0;

  for (const op of ops) {
    if (op.type !== "delete") continue;
    const startTok = A[op.a0];
    const endTok = A[op.a1 - 1];
    if (!startTok || !endTok) continue;

    // solo resaltar “cosas reales” (no espacios solos)
    const span = originalStr.slice(startTok.s, endTok.e);
    if (!span.trim()) continue;

    try {
      textEl.setBackgroundColor(startTok.s, endTok.e - 1, RED);
      marks++;
    } catch (e) {}
  }

  return marks;
}

function tokenizeWords_(s) {
  // tokens con indices (incluye puntuación pegada como “palabra”)
  const out = [];
  const re = /[A-Za-zÁÉÍÓÚÑáéíóúñ0-9]+|[^\sA-Za-zÁÉÍÓÚÑáéíóúñ0-9]+/g;
  let m;
  while ((m = re.exec(s)) !== null) {
    out.push({ w: m[0], s: m.index, e: m.index + m[0].length });
    if (re.lastIndex === m.index) re.lastIndex++;
  }
  return out;
}

function normForMatch_(s) {
  return (s || "")
    .toLowerCase()
    .replace(/\s+/g, " ")
    .replace(/[“”«»"']/g, "")
    .trim();
}

function bigrams_(s) {
  const x = (s || "");
  if (x.length < 2) return [];
  const arr = [];
  for (let i = 0; i < x.length - 1; i++) arr.push(x.slice(i, i + 2));
  return arr;
}

function diceCoef_(a, b) {
  if (!a || !b) return 0;
  if (a === b) return 1;

  const A = bigrams_(a);
  const B = bigrams_(b);
  if (A.length === 0 || B.length === 0) return 0;

  const map = {};
  for (const bg of A) map[bg] = (map[bg] || 0) + 1;

  let inter = 0;
  for (const bg of B) {
    if (map[bg]) { inter++; map[bg]--; }
  }
  return (2 * inter) / (A.length + B.length);
}

// Myers diff (simple y suficiente para esto)
function myersDiff_(a, b) {
  const N = a.length, M = b.length;
  const max = N + M;
  const v = new Map();
  v.set(1, 0);

  const trace = [];

  for (let d = 0; d <= max; d++) {
    const v2 = new Map(v);
    trace.push(v2);

    for (let k = -d; k <= d; k += 2) {
      let x;
      if (k === -d || (k !== d && (v.get(k - 1) ?? -Infinity) < (v.get(k + 1) ?? -Infinity))) {
        x = v.get(k + 1) ?? 0; // insertion
      } else {
        x = (v.get(k - 1) ?? 0) + 1; // deletion
      }
      let y = x - k;

      while (x < N && y < M && a[x] === b[y]) {
        x++; y++;
      }
      v.set(k, x);

      if (x >= N && y >= M) {
        return backtrackMyers_(trace, a, b);
      }
    }
  }
  return [];
}

function backtrackMyers_(trace, a, b) {
  let x = a.length;
  let y = b.length;
  const ops = [];

  for (let d = trace.length - 1; d >= 0; d--) {
    const v = trace[d];
    const k = x - y;

    let prevK;
    if (k === -d || (k !== d && (v.get(k - 1) ?? -Infinity) < (v.get(k + 1) ?? -Infinity))) {
      prevK = k + 1; // came from insert
    } else {
      prevK = k - 1; // came from delete
    }

    const prevX = v.get(prevK) ?? 0;
    const prevY = prevX - prevK;

    while (x > prevX && y > prevY) {
      // equal
      x--; y--;
    }

    if (d === 0) break;

    if (x === prevX) {
      // insert (en b)
      ops.push({ type: "insert", a0: x, a1: x, b0: y - 1, b1: y });
      y--;
    } else {
      // delete (en a)
      ops.push({ type: "delete", a0: x - 1, a1: x, b0: y, b1: y });
      x--;
    }
  }

  ops.reverse();
  return mergeOps_(ops);
}

function mergeOps_(ops) {
  if (ops.length === 0) return ops;
  const out = [];
  let cur = ops[0];

  for (let i = 1; i < ops.length; i++) {
    const o = ops[i];
    const contig =
      o.type === cur.type &&
      o.a0 === cur.a1 &&
      o.b0 === cur.b1;

    if (contig) {
      cur = { type: cur.type, a0: cur.a0, a1: o.a1, b0: cur.b0, b1: o.b1 };
    } else {
      out.push(cur);
      cur = o;
    }
  }
  out.push(cur);
  return out;
}


// ====== REPORTE ======
// FIX #11: eliminado reporte de comparación.

function formatComment_(ch) {
  const rule = ch.ruleId || "REGLA";
  const scope = ch.scope || "";
  // Un “comentario” útil y corto
  return `${rule}\n${scope}`;
}


function removeAllIndents_(p, elementType) {
  // Indents de párrafo
  p.setIndentStart(0);
  p.setIndentEnd(0);
  p.setIndentFirstLine(0);

  // Si es LIST_ITEM, bajamos nesting level
  if (elementType === DocumentApp.ElementType.LIST_ITEM) {
    try { p.asListItem().setNestingLevel(0); } catch (e) {}
  }

  // ✅ NUEVO: si hay TAB dentro del texto, lo convierte a espacio (evita “sangría falsa”)
  normalizeTabsInParagraph_(p);

  // Borrar tabs/espacios al inicio (incluye NBSP)
  trimLeadingWhitespace_(p);
}

function normalizeSrSra_(doc) {
  const body = doc.getBody();
  let touched = 0;

  const DELIM = "(\\s|$|[\\.,;:!\\?\\)\\]\\u00BB\\u201D])";

  const reSra = new RegExp("\\bSra\\.?"+DELIM, "ig");
  const reSr  = new RegExp("\\bSr\\.?"+DELIM, "ig");

  forEachText_(body, (textEl) => {
    touched += replaceInTextPreserveStyle_(textEl, reSra, (m) => `señora${m[1] || ""}`);
    touched += replaceInTextPreserveStyle_(textEl, reSr,  (m) => `señor${m[1] || ""}`);
  });

  return touched;
}


function normalizeLicenciadoConditional_(doc) {
  const body = doc.getBody();
  let touched = 0;

  const DELIM = "(\\s|$|[\\.,;:!\\?\\)\\]\\u00BB\\u201D])";
  const reLic = new RegExp("\\blic\\.?"+DELIM, "ig");

  forEachText_(body, (textEl) => {
    touched += replaceInTextPreserveStyle_(textEl, reLic, (m, full) => {
      const idx = m.index;
      const before = (full || "").slice(Math.max(0, idx - 40), idx).toLowerCase();

      // tolera coma/punto antes de Lic: “señor, Lic.” / “la Lic.”
      const female = /\b(la|señora|sra)\b[\s,;:.]*$/.test(before);
      const male   = /\b(el|señor|sr)\b[\s,;:.]*$/.test(before);

      const delim = m[1] || "";

      if (female) return `licenciada${delim}`;
      if (male)   return `licenciado${delim}`;

      // Si no hay pista, elegí tu default:
      // (a) conservador: no tocar
      // return m[0];

      // (b) default masculino:
      return `licenciado${delim}`;
    });
  });

  return touched;
}

function stripAccents_(s) {
  try {
    return (s || "").normalize("NFD").replace(/[\u0300-\u036f]/g, "");
  } catch (e) {
    return (s || "");
  }
}

/**
 * Convierte números en palabras (español) a número.
 * Soporta lo que te aparece en sentencias: 1..9999 aprox (incluye "dos mil veinticuatro", "setenta y tres", etc.)
 * Devuelve number o null si no puede parsear.
 */
function wordsToNumberEs_(words) {
  if (!words) return null;

  let s = stripAccents_((words + "").toLowerCase())
    .replace(/[^a-z0-9\s]/g, " ")
    .replace(/\s+/g, " ")
    .trim();

  if (!s) return null;
  if (/^\d+$/.test(s)) return parseInt(s, 10);

  const U = { cero:0, un:1, uno:1, una:1, dos:2, tres:3, cuatro:4, cinco:5, seis:6, siete:7, ocho:8, nueve:9 };
  const T10 = { diez:10, once:11, doce:12, trece:13, catorce:14, quince:15, dieciseis:16, diecisiete:17, dieciocho:18, diecinueve:19 };
  const D = { veinte:20, treinta:30, cuarenta:40, cincuenta:50, sesenta:60, setenta:70, ochenta:80, noventa:90 };
  const C = { cien:100, ciento:100, doscientos:200, trescientos:300, cuatrocientos:400, quinientos:500, seiscientos:600, setecientos:700, ochocientos:800, novecientos:900 };

  const toks = s.split(" ");
  let total = 0;
  let cur = 0;

  for (let i = 0; i < toks.length; i++) {
    const tok = toks[i];
    if (!tok || tok === "y") continue;

    // veintidos / veintitres / veinticuatro...
    if (tok.startsWith("veinti") && tok !== "veinte" && tok !== "veinti" && tok.length > 6) {
      const rest = tok.slice(6);
      const u = (U[rest] != null) ? U[rest] : ((rest === "un" || rest === "uno") ? 1 : null);
      if (u == null) return null;
      cur += 20 + u;
      continue;
    }
    if (tok === "veinti") { cur += 20; continue; } // “veinti cuatro”

    // dieciseis / diecisiete / ...
    if (tok.startsWith("dieci") && tok.length > 5) {
      const rest = tok.slice(5);
      const u = U[rest];
      if (u == null) return null;
      cur += 10 + u;
      continue;
    }

    if (tok === "mil") {
      total += (cur || 1) * 1000;
      cur = 0;
      continue;
    }

    if (C[tok] != null) { cur += C[tok]; continue; }
    if (D[tok] != null) { cur += D[tok]; continue; }
    if (T10[tok] != null) { cur += T10[tok]; continue; }
    if (U[tok] != null) { cur += U[tok]; continue; }

    return null;
  }

  return total + cur;
}

/** Pone en negrita SOLO el prefijo "I." (con punto). */
function boldRomanNumeralIPrefix_(paragraph) {
  const s = (paragraph.getText() || "");
  if (!s) return 0;

  // solo si arranca con I. (permitimos espacios antes)
  if (!/^\s*I\./.test(s)) return 0;

  const start = s.search(/I\./);
  if (start === -1) return 0;

  try {
    const t = paragraph.editAsText();
    t.setBold(start, start + 1, true);       // "I."
    t.setItalic(start, start + 1, false);
    t.setUnderline(start, start + 1, false);
    return 1;
  } catch (e) {
    return 0;
  }
}


function normalizeTabsInParagraph_(p) {
  try {
    const t = p.editAsText();
    const s = t.getText() || "";
    if (!s || s.indexOf("\t") === -1) return 0;

    // Reemplaza cualquier cantidad de tabs por 1 espacio
    return replaceInTextPreserveStyle_(t, /\t+/g, " ");
  } catch (e) {
    return 0;
  }
}

function fixFirstQuestionIntroSentenciaI_(doc, log) {
  const body = doc.getBody();
  const n = body.getNumChildren();

  const rxPrimera = /^\s*A\s+LA\s+PRIMERA\s+CUESTI[ÓO]N\b/i;
  const rxSeccion = /^\s*A\s+LA\s+(PRIMERA|SEGUNDA|TERCERA)\s+CUESTI[ÓO]N\b/i;

  // 1) encontrar el heading de PRIMERA CUESTIÓN
  let idxPrimera = -1;
  for (let i = 0; i < n; i++) {
    const el = body.getChild(i);
    if (el.getType() !== DocumentApp.ElementType.PARAGRAPH &&
        el.getType() !== DocumentApp.ElementType.LIST_ITEM) continue;

    const p = elementToParagraphOrListItem_(el);
    const t = (p.getText() || "").trim();
    if (rxPrimera.test(t)) { idxPrimera = i; break; }
  }

  if (idxPrimera === -1) {
    log && log.push(makeChange_("FIRSTQ_I_SENTENCIA", "Documento", "No encontré A LA PRIMERA CUESTIÓN", "Sin cambios", {}));
    return;
  }

  // 2) buscar el primer párrafo después que empiece con I. y tenga “Sentencia” cerca del inicio
  let targetEl = null;
  let targetP = null;
  let targetIndex = -1;

  for (let i = idxPrimera + 1; i < n; i++) {
    const el = body.getChild(i);

    if (el.getType() !== DocumentApp.ElementType.PARAGRAPH &&
        el.getType() !== DocumentApp.ElementType.LIST_ITEM) continue;

    const p = elementToParagraphOrListItem_(el);
    const raw = (p.getText() || "");
    const t = raw.trim();
    if (!t) continue;

    // si ya entramos a otra sección, cortamos
    if (rxSeccion.test(t)) break;

    // condición: empieza con I. (o I)) y “Sentencia” está “ahí nomás” (en los primeros ~80 chars)
    const startsI = /^\s*I\s*[\.\)]/.test(raw);
    const posSent = raw.toLowerCase().indexOf("sentencia");
    if (startsI && posSent !== -1 && posSent <= 80) {
      targetEl = el;
      targetP = p;
      targetIndex = i;
      break;
    }
  }

  if (!targetP) {
    log && log.push(makeChange_("FIRSTQ_I_SENTENCIA", `Body párrafo ${idxPrimera + 1}`, "No encontré párrafo I. con Sentencia", "Sin cambios", {}));
    return;
  }

  const before = targetP.getText() || "";
  let after = canonicalizeFirstQuestionIText_(before);

  // nada para hacer
  if (!after || after === before) {
    // igual aplicamos formato de comillas (por si venía mal)
    const allLen = (targetP.getText() || "").length;
    if (allLen > 0) {
      const tAll = targetP.editAsText();
      tAll.setBold(0, allLen - 1, false);
      tAll.setItalic(0, allLen - 1, false);
      tAll.setUnderline(0, allLen - 1, false);
    }
    const q = italicizeQuotedTextNoBold_(targetP);
    boldRomanNumeralIPrefix_(targetP);


    log && log.push(makeChange_("FIRSTQ_I_SENTENCIA", `Body párrafo ${targetIndex + 1}`, "(sin cambios)", `Comillas en cursiva/no negrita: ${q}`, {}));
    return;
  }

  // 3) reescribir texto
  targetP.setText(after);

  // 4) estilo consistente
  targetP.setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY);
  removeAllIndents_(targetP, targetEl.getType());
  targetP.setLineSpacing(1.5);
  targetP.setSpacingBefore(0);
  targetP.setSpacingAfter(0);

  const tt = targetP.editAsText();
  tt.setFontFamily("Times New Roman");
  tt.setFontSize(12);
  clearUnderline_(targetP);

  // 5) Dejar el párrafo “normal” (sin negrita/cursiva/subrayado)
  const allLen = (targetP.getText() || "").length;
  if (allLen > 0) {
    const tAll = targetP.editAsText();
    tAll.setBold(0, allLen - 1, false);
    tAll.setItalic(0, allLen - 1, false);
    tAll.setUnderline(0, allLen - 1, false);
  }

  // 6) Comillas: todo lo entre comillas -> cursiva y sin negrita
  const qCount = italicizeQuotedTextNoBold_(targetP);
  boldRomanNumeralIPrefix_(targetP);


}

function canonicalizeFirstQuestionIText_(txt) {
  if (!txt) return txt;

  // Limpieza básica (incluye tabs)
  txt = txt
    .replace(/[\t\u00A0]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();

  // I. (normaliza “I)”, “I .”, etc.)
  txt = txt.replace(/^\s*I\s*[\.\)]\s*/i, "I. ");

  // I. Por Sentencia ...
  txt = txt.replace(/^I\.\s*Por\s+sentencia\b/i, "I. Por Sentencia");

  // Asegurar: Sentencia n° <num>
  // (acepta N°, Nº, nro, número, etc.)
  txt = txt.replace(
    /\bSentencia\s*(?:n[°ºo]\.?|nº|n°|N°|Nº|nro\.?|Nro\.?|número)\s*([0-9]{1,4})\b/i,
    (m, num) => `Sentencia n° ${num}`
  );

  // fallback: “Sentencia 16” -> “Sentencia n° 16”
  txt = txt.replace(
    /\bSentencia\s+([0-9]{1,4})\b/i,
    (m, num) => `Sentencia n° ${num}`
  );

  // Fecha: del/de fecha dd/mm/yyyy -> del/de fecha d de mes de yyyy
  txt = txt.replace(
    /\b(del|de fecha)\s+(\d{1,2})\s*[\/-]\s*(\d{1,2})\s*[\/-]\s*(\d{4})\b/i,
    (m, pref, dd, mm, yyyy) => {
      const d = parseInt(dd, 10);
      const mo = parseInt(mm, 10);
      return `${pref.toLowerCase()} ${d} de ${monthNameEs_(mo)} de ${yyyy}`;
    }
  );

  // Si ya está “del 1 de enero de 2001”, dejamos pero normalizamos minúsculas de mes
  txt = txt.replace(
    /\b(del|de fecha)\s+(\d{1,2})\s+de\s+([a-záéíóúñ]+)\s+de\s+(\d{4})\b/i,
    (m, pref, dd, mes, yyyy) => {
      const d = parseInt(dd, 10);
      return `${pref.toLowerCase()} ${d} de ${mes.toLowerCase()} de ${yyyy}`;
    }
  );

  // Nominación en letras (de/del 3ra Nominación -> de Tercera Nominación)
  txt = normalizeNominacionEnLetras_(txt);

  // Extra: si viene “Cámara 3ra Nominación” sin “de”, lo arreglamos a “Cámara de Tercera Nominación”
  txt = txt.replace(
    /\b(c[aá]mara|juzgado)\s+(\d{1,2})\s*(?:ª|º|°|a|A|ra|RA)?\s+Nominaci[óo]n\b/ig,
    (m, trib, numStr) => {
      const n = parseInt(numStr, 10);
      const T = trib.toLowerCase().startsWith("c") ? "Cámara" : "Juzgado";
      return `${T} de ${ordinalFemenino_(n)} Nominación`;
    }
  );

  // ✅ NUEVO: "Sentencia número setenta y tres" -> "Sentencia n° 73" (y Auto)
  txt = txt.replace(
    /\b(Sentencia|Auto)\s+n[úu]mero\s+([a-záéíóúñ\s]+?)(\s+(?:el|del|de\s+fecha)\b|[.,;:\)\]])/ig,
    (m, tipo, numWords, delim) => {
      const n = wordsToNumberEs_(numWords);
      if (n == null) return m;
      const T = tipo[0].toUpperCase() + tipo.slice(1).toLowerCase();
      return `${T} n° ${n}${delim}`;
    }
  );

  // ✅ NUEVO: "el ocho de octubre de dos mil veinticuatro" -> "el 8 de octubre de 2024"
  txt = txt.replace(
    /\b(el|del|de fecha)\s+([a-záéíóúñ0-9\s]{1,30})\s+de\s+([a-záéíóúñ]+)\s+de\s+([a-záéíóúñ0-9\s]{2,60})\b/ig,
    (m, pref, dayStr, mes, yearStr) => {
      const dRaw = (dayStr || "").trim();
      const yRaw = (yearStr || "").trim();

      const day = /^\d{1,2}$/.test(dRaw) ? parseInt(dRaw, 10) : wordsToNumberEs_(dRaw);
      const year = /^\d{4}$/.test(yRaw) ? parseInt(yRaw, 10) : wordsToNumberEs_(yRaw);

      if (day == null || day < 1 || day > 31) return m;
      if (year == null || year < 1000 || year > 2999) return m;

      return `${pref.toLowerCase()} ${day} de ${mes.toLowerCase()} de ${year}`;
    }
  );


  // Limpieza final
  txt = txt
    .replace(/\s+,/g, ",")
    .replace(/\s+\./g, ".")
    .replace(/\s{2,}/g, " ")
    .trim();

  return txt;
}

function italicizeQuotedTextNoBold_(paragraph) {
  const t = paragraph.editAsText();
  const full = paragraph.getText() || "";
  if (!full) return 0;

  const OPEN  = ['"', '“', '«'];
  const CLOSE = ['"', '”', '»'];

  let from = 0;
  let count = 0;

  while (from < full.length) {
    // buscar próxima comilla de apertura (cualquiera)
    let i1 = -1;
    for (const ch of OPEN) {
      const idx = full.indexOf(ch, from);
      if (idx !== -1 && (i1 === -1 || idx < i1)) i1 = idx;
    }
    if (i1 === -1) break;

    // buscar próximo cierre (cualquiera) después de i1
    let i2 = -1;
    for (const ch of CLOSE) {
      const idx = full.indexOf(ch, i1 + 1);
      if (idx !== -1 && (i2 === -1 || idx < i2)) i2 = idx;
    }
    if (i2 === -1) break;

    const start = i1 + 1;
    const end = i2 - 1;

    if (end >= start) {
      try {
        t.setItalic(start, end, true);
        t.setBold(start, end, false);
      } catch (e) {}
      count++;
    }

    from = i2 + 1;
  }

  return count;
}

function trimLeadingWhitespace_(p) {
  const t = p.editAsText();
  const s = t.getText();
  if (!s) return;

  const m = s.match(/^[\t \u00A0]+/);
  if (!m) return;

  const len = m[0].length;
  t.deleteText(0, len - 1);
}

function clearUnderline_(paragraph) {
  const txt = paragraph.editAsText();
  const s = txt.getText() || "";
  if (!s) return;
  txt.setUnderline(0, s.length - 1, false);
}

// ====== GÉNERO LINEA VOTO ======
function vocalGenero_(nombre) {
  const g = VOCALES_GENERO[nombre] || "M";
  return (g === "F")
    ? { art: "La", senor: "señora", doc: "doctora" }
    : { art: "El", senor: "señor", doc: "doctor" };
}

// ----------------- CORE (recorre párrafos y aplica regex con grupos) -----------------

function highlightByRulesInBody_(body, rules) {
  let found = 0;
  rules.forEach(rule => found += highlightByRuleInBody_(body, rule));
  return found;
}

function highlightByRuleInBody_(body, rule) {
  let found = 0;

  // Body
  const n = body.getNumChildren();
  for (let i = 0; i < n; i++) {
    const el = body.getChild(i);
    found += highlightInElement_(el, rule);
  }

  // Tablas
  const tables = body.getTables();
  for (let ti = 0; ti < tables.length; ti++) {
    const table = tables[ti];
    for (let r = 0; r < table.getNumRows(); r++) {
      const row = table.getRow(r);
      for (let c = 0; c < row.getNumCells(); c++) {
        const cell = row.getCell(c);
        for (let k = 0; k < cell.getNumChildren(); k++) {
          found += highlightInElement_(cell.getChild(k), rule);
        }
      }
    }
  }

  return found;
}

function highlightInElement_(el, rule) {
  const t = el.getType();

  if (t === DocumentApp.ElementType.PARAGRAPH) {
    return highlightInParagraph_(el.asParagraph(), rule);
  }
  if (t === DocumentApp.ElementType.LIST_ITEM) {
    return highlightInListItem_(el.asListItem(), rule);
  }
  // Si querés también dentro de tablas anidadas ya las cubrimos arriba.
  return 0;
}

function highlightInParagraph_(p, rule) {
  const text = p.editAsText();
  return highlightInText_(text, rule);
}

function highlightInListItem_(li, rule) {
  const text = li.editAsText();
  return highlightInText_(text, rule);
}

function highlightInText_(text, rule) {
  const s = text.getText() || "";
  if (!s) return 0;

  // Re-creamos el regex para evitar problemas con lastIndex si viene reutilizado
  const flags = (rule.re.ignoreCase ? "i" : "") + (rule.re.multiline ? "m" : "") + (rule.re.global ? "g" : "g") + (rule.re.unicode ? "u" : "");
  const re = new RegExp(rule.re.source, flags);

  let m;
  let count = 0;

  while ((m = re.exec(s)) !== null) {
    const fullStart = m.index;
    const fullEnd = fullStart + m[0].length - 1;

    let start = fullStart;
    let end = fullEnd;

    // Si piden un grupo específico, calculamos offset del grupo
    if (rule.group && rule.group > 0 && m[rule.group] != null) {
      const gText = m[rule.group];

      // Ojo: buscamos el grupo dentro del match (para ubicarlo)
      const within = m[0].indexOf(gText);
      if (within >= 0) {
        start = fullStart + within;
        end = start + gText.length - 1;
      }
    }

    // Si es “solo si NO está en negrita”, verificamos
    if (rule.onlyIfNotBold) {
      if (rangeHasBold_(text, start, end)) continue;
    }

    try {
      text.setBackgroundColor(start, end, rule.color);
      count++;
    } catch (e) {
      // seguimos, no frenamos
    }

    // Evita loop infinito si el match es vacío
    if (re.lastIndex === m.index) re.lastIndex++;
  }

  return count;
}

// ====== UTIL ======
function makeChange_(ruleId, scope, beforeText, afterText, extra) {
  return {
    changeId: Utilities.getUuid(),
    ruleId,
    scope,
    beforeText,
    afterText,
    extra: extra || {}
  };
}

function stripExt_(name) {
  return name.replace(/\.[^/.]+$/, "");
}
