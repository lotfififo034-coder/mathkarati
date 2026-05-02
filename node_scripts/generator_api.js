/**
 * MathKarati PRO v3 — API Mode
 * Reads JSON payload from stdin, writes PPTX bytes to stdout.
 * Called by Flask as a subprocess.
 *
 * Usage:
 *   echo '<json>' | node generator_api.js
 */

"use strict";

const PptxGenJS = require("pptxgenjs");

const W = 13.33;
const H = 7.5;
const MX = 0.65;
const MY = 0.52;
const GU = 0.28;

const T = {
  display: { sz: 52, bold: true },
  h1:      { sz: 38, bold: true },
  h2:      { sz: 28, bold: true },
  h3:      { sz: 20, bold: true },
  h4:      { sz: 15, bold: true },
  over:    { sz: 9,  bold: true },
  body:    { sz: 13, bold: false },
  bodyS:   { sz: 11, bold: false },
  caption: { sz: 9,  bold: false },
  stat:    { sz: 48, bold: true },
  statS:   { sz: 32, bold: true },
};

// ═══════════════════════════════════════════════════════
// STYLES
// ═══════════════════════════════════════════════════════
const STYLES = {
  noir: {
    id: "noir", name: "Noir Académique",
    ink: "0A0C12", inkMid: "131620", paper: "F5F2EC", paperS: "EDE9E1",
    pl: "F5F2EC", pd: "0A0C12",
    gold: "C9A84C", goldS: "E8C97A", silver: "9BA8B5",
    cardD: "161B26", cardL: "FFFFFF", border: "2A3040", borderL: "E0DDD6",
    textL: "F5F2EC", textD: "0A0C12", textM: "6B7A8D", textMl: "B8AD9E",
    chart: ["C9A84C","9BA8B5","E8C97A","5A6B7A","F5F2EC","3D4E60","8C7A56"],
    hf: "Palatino Linotype", bf: "Calibri", af: "Cairo",
    rule: "C9A84C", accent: "C9A84C",
  },
  atlas: {
    id: "atlas", name: "Atlas Corporate",
    ink: "040D1E", inkMid: "082040", paper: "F8FAFE", paperS: "EEF4FC",
    pl: "F8FAFE", pd: "040D1E",
    gold: "00D4FF", goldS: "7AEAFF", silver: "4A6FA5",
    cardD: "071428", cardL: "FFFFFF", border: "1A3054", borderL: "D4E4F4",
    textL: "F8FAFE", textD: "040D1E", textM: "4A6FA5", textMl: "8BAAD0",
    chart: ["00D4FF","FF6B00","7AEAFF","0A4A8A","FFFFFF","1A3054","FF9A50"],
    hf: "Trebuchet MS", bf: "Calibri", af: "Cairo",
    rule: "00D4FF", accent: "00D4FF",
  },
  sakura: {
    id: "sakura", name: "Sakura Créative",
    ink: "1C0E2B", inkMid: "2D1B45", paper: "FDFAF8", paperS: "F5EFF5",
    pl: "FDFAF8", pd: "1C0E2B",
    gold: "FF5C5C", goldS: "FFAAAA", silver: "6B4F8C",
    cardD: "261540", cardL: "FFFFFF", border: "3D2560", borderL: "E8D8F0",
    textL: "FDFAF8", textD: "1C0E2B", textM: "6B4F8C", textMl: "B09CC0",
    chart: ["FF5C5C","8B5CF6","FFAAAA","4C1D80","FDFAF8","5B3A7A","FF9090"],
    hf: "Georgia", bf: "Calibri", af: "Cairo",
    rule: "FF5C5C", accent: "FF5C5C",
  },
};

// Map UI theme names → style keys
const THEME_MAP = {
  navy_gold:        "noir",
  dark_teal:        "atlas",
  burgundy:         "sakura",
  forest:           "noir",
  midnight_purple:  "sakura",
  charcoal_orange:  "atlas",
  ice_blue:         "atlas",
  sand_gold:        "noir",
  noir:             "noir",
  atlas:            "atlas",
  sakura:           "sakura",
};

// ═══════════════════════════════════════════════════════
// UTILITIES
// ═══════════════════════════════════════════════════════
const safe = (v, fb = "") =>
  (v !== null && v !== undefined && String(v).trim()) ? String(v).trim() : fb;

const mkShadow = (blur = 8, offset = 2, opacity = 0.12) =>
  ({ type: "outer", color: "000000", blur, offset, angle: 135, opacity });

function overline(slide, S, text, x, y, w, align = "left") {
  slide.addText(text.toUpperCase(), {
    x, y, w, h: 0.22,
    fontSize: T.over.sz, fontFace: S.bf, bold: true,
    color: S.accent, charSpacing: 3.5, align, margin: 0,
  });
}

function rule(slide, color, x, y, w, h = 0.025) {
  slide.addShape("rect", { x, y, w, h, fill: { color }, line: { type: "none" } });
}

function decoCircle(slide, color, cx, cy, r, opacity = 0.06) {
  slide.addShape("ellipse", {
    x: cx - r, y: cy - r, w: r * 2, h: r * 2,
    fill: { color, transparency: Math.round((1 - opacity) * 100) },
    line: { type: "none" },
  });
}

function card(slide, x, y, w, h, fillColor, accentColor = null, accentW = 0, shadow = true) {
  slide.addShape("rect", {
    x, y, w, h,
    fill: { color: fillColor },
    line: { type: "none" },
    ...(shadow ? { shadow: mkShadow(10, 2, 0.10) } : {}),
  });
  if (accentColor && accentW > 0) {
    slide.addShape("rect", {
      x, y, w: accentW, h,
      fill: { color: accentColor }, line: { type: "none" },
    });
  }
}

// ═══════════════════════════════════════════════════════
// SLIDE 01 — CINEMATIC COVER
// ═══════════════════════════════════════════════════════
async function slide01_Cover(prs, S, d) {
  const slide = prs.addSlide();
  slide.background = { color: S.ink };

  const panelW = W * 0.38;

  slide.addShape("rect", { x: 0, y: 0, w: panelW, h: H, fill: { color: S.inkMid }, line: { type: "none" } });
  slide.addShape("rect", { x: panelW - 0.055, y: 0, w: 0.055, h: H, fill: { color: S.accent }, line: { type: "none" } });
  decoCircle(slide, S.accent, panelW * 0.5, H * 0.72, 3.2, 0.04);

  slide.addText("MÉMOIRE DE MASTER  ·  رسالة ماستر", {
    x: 0.10, y: 0.6, w: H - 1.2, h: 0.38,
    fontSize: 7.5, fontFace: S.bf, bold: true,
    color: S.textM, charSpacing: 2.5,
    align: "center", rotate: 90, margin: 0,
  });

  slide.addText(safe(d.university, "الجامعة"), {
    x: 0.28, y: H - 2.1, w: panelW - 0.55, h: 0.38,
    fontSize: 12, fontFace: S.af, bold: true,
    color: S.goldS, align: "right", rtlMode: true, margin: 0,
  });

  const fac = [d.faculty, d.department].filter(Boolean).join("  ·  ");
  if (fac) {
    slide.addText(fac, {
      x: 0.28, y: H - 1.68, w: panelW - 0.55, h: 0.5,
      fontSize: 10, fontFace: S.af,
      color: S.textM, align: "right", rtlMode: true, margin: 0,
    });
  }

  slide.addShape("rect", { x: 0.28, y: H - 1.04, w: panelW - 0.55, h: 0.38, fill: { color: S.accent }, line: { type: "none" } });
  slide.addText(safe(d.level, "ماستر  ·  Master"), {
    x: 0.28, y: H - 1.04, w: panelW - 0.55, h: 0.38,
    fontSize: 11, fontFace: S.af, bold: true,
    color: S.ink, align: "center", valign: "middle", margin: 0,
  });

  const rx = panelW + 0.68;
  const rw = W - rx - MX;

  overline(slide, S, safe(d.fieldEn, "Research  ·  Thesis"), rx, MY + 0.05, rw, "left");
  rule(slide, S.accent, rx, MY + 0.32, 1.4);

  slide.addText(safe(d.titleAr, "عنوان الرسالة"), {
    x: rx, y: MY + 0.48, w: rw, h: 3.2,
    fontSize: 30, fontFace: S.af, bold: true,
    color: S.textL, align: "right", rtlMode: true,
    valign: "middle", paraSpaceAfter: 6, margin: 0,
  });

  if (d.titleFr) {
    rule(slide, S.border, rx, MY + 3.82, rw);
    slide.addText(safe(d.titleFr), {
      x: rx, y: MY + 3.96, w: rw, h: 0.50,
      fontSize: 12, fontFace: S.bf, italic: true,
      color: S.textM, align: "left", margin: 0,
    });
  }

  const barY = H - 1.18;
  slide.addShape("rect", { x: panelW + 0.055, y: barY, w: W - panelW - 0.055, h: 1.18, fill: { color: S.cardD }, line: { type: "none" } });
  rule(slide, S.accent, panelW + 0.055, barY, W - panelW - 0.055, 0.04);

  slide.addText("PRÉSENTÉ PAR  ·  إعداد", {
    x: rx, y: barY + 0.12, w: rw / 2, h: 0.20,
    fontSize: 7.5, fontFace: S.bf, bold: true,
    color: S.textM, charSpacing: 2, align: "left", margin: 0,
  });
  slide.addText(safe(d.studentName, "اسم الطالب"), {
    x: rx, y: barY + 0.34, w: rw / 2 - 0.2, h: 0.40,
    fontSize: 15, fontFace: S.af, bold: true,
    color: S.textL, align: "right", rtlMode: true, margin: 0,
  });

  const sCol = rx + rw / 2 + GU;
  slide.addText("ENCADRÉ PAR  ·  إشراف", {
    x: sCol, y: barY + 0.12, w: rw / 2, h: 0.20,
    fontSize: 7.5, fontFace: S.bf, bold: true,
    color: S.textM, charSpacing: 2, align: "left", margin: 0,
  });
  slide.addText(safe(d.supervisor, "اسم المشرف"), {
    x: sCol, y: barY + 0.34, w: rw / 2 - 0.2, h: 0.40,
    fontSize: 15, fontFace: S.af, bold: true,
    color: S.textL, align: "right", rtlMode: true, margin: 0,
  });

  slide.addShape("rect", { x: W - MX - 1.4, y: barY + 0.70, w: 1.4, h: 0.30, fill: { color: S.accent }, line: { type: "none" } });
  slide.addText(safe(d.year, "2024–2025"), {
    x: W - MX - 1.4, y: barY + 0.70, w: 1.4, h: 0.30,
    fontSize: 10, fontFace: S.bf, bold: true,
    color: S.ink, align: "center", valign: "middle", margin: 0,
  });

  if (d.defenseDate) {
    slide.addText("تاريخ المناقشة: " + safe(d.defenseDate), {
      x: rx, y: barY + 0.76, w: rw / 2, h: 0.28,
      fontSize: 9, fontFace: S.af,
      color: S.textM, align: "right", rtlMode: true, margin: 0,
    });
  }

  return slide;
}

// ═══════════════════════════════════════════════════════
// SLIDE 02 — SECTION DIVIDER (Table of Contents)
// ═══════════════════════════════════════════════════════
function slide02_TOC(prs, S, chapters) {
  const slide = prs.addSlide();
  slide.background = { color: S.ink };

  decoCircle(slide, S.accent, MX + 0.6, H / 2, 1.8, 0.07);
  slide.addShape("rect", { x: MX, y: H / 2 - 1.05, w: 0.14, h: 2.1, fill: { color: S.accent }, line: { type: "none" } });

  overline(slide, S, "TABLE DES MATIÈRES  ·  المحتويات", MX + 0.30, MY, 8, "left");
  slide.addText("المحتويات", {
    x: MX + 0.30, y: MY + 0.28, w: W - MX * 2, h: 0.80,
    fontSize: 38, fontFace: S.af, bold: true,
    color: S.textL, align: "right", rtlMode: true, valign: "middle", margin: 0,
  });
  rule(slide, S.accent, MX + 0.30, MY + 1.18, W - MX * 2 - 0.30, 0.04);

  const chs = (chapters || []).slice(0, 6);
  const cols = 2;
  const cw = (W - MX * 2 - GU) / cols - 0.1;
  const startY = MY + 1.38;
  const availH = H - startY - MY * 0.5;
  const rows = Math.ceil(chs.length / cols);
  const ch = (availH - 0.16 * (rows - 1)) / rows;

  chs.forEach((item, i) => {
    const col = i % cols;
    const row = Math.floor(i / cols);
    const cx = MX + 0.30 + col * (cw + GU + 0.2);
    const cy = startY + row * (ch + 0.16);
    const sc = S.chart[i % S.chart.length];

    slide.addShape("rect", { x: cx, y: cy, w: cw, h: ch, fill: { color: S.inkMid }, line: { type: "none" } });
    slide.addShape("rect", { x: cx, y: cy, w: cw, h: 0.04, fill: { color: sc }, line: { type: "none" } });
    slide.addShape("rect", { x: cx, y: cy, w: 0.08, h: ch, fill: { color: sc }, line: { type: "none" } });

    slide.addText(String(i + 1).padStart(2, "0"), {
      x: cx + 0.16, y: cy + 0.06, w: 0.70, h: ch - 0.12,
      fontSize: 28, fontFace: "Georgia", bold: true,
      color: sc, align: "left", valign: "middle", margin: 0,
    });

    slide.addText(safe(item.title || item), {
      x: cx + 0.90, y: cy + 0.08, w: cw - 1.06, h: ch * 0.55,
      fontSize: 13, fontFace: S.af, bold: true,
      color: S.textL, align: "right", rtlMode: true, valign: "middle", margin: 0,
    });

    if (item.sub) {
      slide.addText(safe(item.sub), {
        x: cx + 0.90, y: cy + ch * 0.60, w: cw - 1.06, h: ch * 0.34,
        fontSize: 9, fontFace: S.bf, italic: true,
        color: S.textM, align: "left", valign: "middle", margin: 0,
      });
    }
  });

  return slide;
}

// ═══════════════════════════════════════════════════════
// SLIDE 03 — RESEARCH PROBLEM
// ═══════════════════════════════════════════════════════
function slide03_Problem(prs, S, d) {
  const slide = prs.addSlide();
  slide.background = { color: S.paper };

  const leftW = W * 0.50;

  slide.addShape("rect", { x: 0, y: 0, w: leftW, h: H, fill: { color: S.ink }, line: { type: "none" } });
  slide.addText("\u201C", {
    x: MX, y: 0.1, w: 2.5, h: 2.0,
    fontSize: 140, fontFace: S.hf, bold: false,
    color: S.accent, align: "left", valign: "top", margin: 0,
  });

  overline(slide, S, "PROBLÉMATIQUE CENTRALE  ·  إشكالية البحث", MX, MY + 1.55, leftW - MX - 0.4);

  slide.addText(safe(d.mainProblem || d.problemAr, "نص الإشكالية الرئيسية"), {
    x: MX, y: MY + 1.85, w: leftW - MX * 2 + 0.25, h: H - MY * 2 - 1.85 - 0.6,
    fontSize: 15.5, fontFace: S.af,
    color: S.textL, align: "right", rtlMode: true,
    valign: "top", paraSpaceAfter: 8, margin: 0,
  });

  rule(slide, S.accent, MX, H - 0.52, leftW - MX - 0.2);

  const rx = leftW + 0.55;
  const rw = W - rx - MX;

  overline(slide, S, "SOUS-QUESTIONS  ·  التساؤلات الفرعية", rx, MY, rw, "left");
  rule(slide, S.accent, rx, MY + 0.26, 1.0);

  const subs = (d.subQuestions || []).filter(Boolean).slice(0, 5);
  const avail = H - MY - 0.55;
  const rowH = subs.length ? (avail / subs.length) - 0.1 : avail;

  subs.forEach((q, i) => {
    const ry = MY + 0.48 + i * (rowH + 0.10);
    const sc = S.chart[i % S.chart.length];

    slide.addShape("rect", {
      x: rx - 0.1, y: ry, w: rw + 0.1, h: rowH,
      fill: { color: i % 2 === 0 ? "FFFFFF" : S.paperS },
      line: { color: S.borderL, width: 0.3 },
    });
    slide.addText(String(i + 1).padStart(2, "0"), {
      x: rx + 0.10, y: ry + 0.04, w: 0.55, h: rowH - 0.08,
      fontSize: 26, fontFace: "Georgia", bold: true,
      color: sc, align: "center", valign: "middle", margin: 0,
    });
    slide.addShape("rect", {
      x: rx + 0.72, y: ry + rowH * 0.2, w: 0.025, h: rowH * 0.6,
      fill: { color: sc }, line: { type: "none" },
    });
    slide.addText(safe(q), {
      x: rx + 0.85, y: ry + 0.06, w: rw - 0.90, h: rowH - 0.12,
      fontSize: 12.5, fontFace: S.af,
      color: S.textD, align: "right", rtlMode: true,
      valign: "middle", margin: 0,
    });
  });

  return slide;
}

// ═══════════════════════════════════════════════════════
// SLIDE 04 — OBJECTIVES & HYPOTHESES
// ═══════════════════════════════════════════════════════
function slide04_Objectives(prs, S, d) {
  const slide = prs.addSlide();
  slide.background = { color: S.paperS };

  const headerH = 0.88;
  slide.addShape("rect", { x: 0, y: 0, w: W, h: headerH, fill: { color: S.ink }, line: { type: "none" } });
  overline(slide, S, "OBJECTIFS & HYPOTHÈSES", MX, 0.10, W / 2 - MX, "left");
  slide.addText("الأهداف والفرضيات", {
    x: MX, y: 0.26, w: W - MX * 2, h: 0.52,
    fontSize: 22, fontFace: S.af, bold: true,
    color: S.textL, align: "right", rtlMode: true, valign: "middle", margin: 0,
  });

  const colW = (W - MX * 2 - GU) / 2;
  const contentY = headerH + 0.22;
  const contentH = H - contentY - MY * 0.5;

  const objs = (d.objectives || []).filter(Boolean).slice(0, 5);
  const hypos = (d.hypotheses || []).filter(Boolean).slice(0, 5);
  const objRows = Math.max(objs.length, 1);
  const hypRows = Math.max(hypos.length, 1);
  const objH = (contentH - 0.14 * (objRows - 1)) / objRows;
  const hypH = (contentH - 0.14 * (hypRows - 1)) / hypRows;

  // OBJECTIVES column header
  slide.addShape("rect", { x: MX, y: contentY, w: colW, h: 0.36, fill: { color: S.accent }, line: { type: "none" } });
  slide.addText("الأهداف  ·  Objectifs", {
    x: MX + 0.12, y: contentY, w: colW - 0.24, h: 0.36,
    fontSize: 10.5, fontFace: S.af, bold: true,
    color: S.ink, align: "right", valign: "middle", margin: 0,
  });

  objs.forEach((obj, i) => {
    const sc = S.chart[i % S.chart.length];
    const cy = contentY + 0.44 + i * (objH - 0.44 / objRows + 0.14);
    const ch = objH - (objRows > 1 ? 0.44 / objRows : 0.1);
    card(slide, MX, cy, colW, ch, S.cardL, sc, 0.055, true);
    slide.addText(String(i + 1).padStart(2, "0"), {
      x: MX + 0.08, y: cy + 0.05, w: 0.52, h: ch - 0.10,
      fontSize: 22, fontFace: "Georgia", bold: true,
      color: sc, align: "center", valign: "middle", margin: 0,
    });
    slide.addText(safe(obj), {
      x: MX + 0.68, y: cy + 0.06, w: colW - 0.80, h: ch - 0.12,
      fontSize: 11.5, fontFace: S.af,
      color: S.textD, align: "right", rtlMode: true, valign: "middle", margin: 0,
    });
  });

  // HYPOTHESES column header
  const rx = MX + colW + GU;
  slide.addShape("rect", { x: rx, y: contentY, w: colW, h: 0.36, fill: { color: S.silver }, line: { type: "none" } });
  slide.addText("الفرضيات  ·  Hypothèses", {
    x: rx + 0.12, y: contentY, w: colW - 0.24, h: 0.36,
    fontSize: 10.5, fontFace: S.af, bold: true,
    color: S.textL, align: "right", valign: "middle", margin: 0,
  });

  hypos.forEach((hy, i) => {
    const sc = S.chart[(i + 2) % S.chart.length];
    const cy = contentY + 0.44 + i * (hypH - 0.44 / hypRows + 0.14);
    const ch = hypH - (hypRows > 1 ? 0.44 / hypRows : 0.1);

    slide.addShape("rect", {
      x: rx, y: cy, w: colW, h: ch,
      fill: { color: S.ink },
      line: { color: sc, width: 1 },
      shadow: mkShadow(8, 2, 0.12),
    });
    slide.addShape("rect", { x: rx + colW - 0.62, y: cy + 0.06, w: 0.52, h: 0.32, fill: { color: sc }, line: { type: "none" } });
    slide.addText(`H${i + 1}`, {
      x: rx + colW - 0.62, y: cy + 0.06, w: 0.52, h: 0.32,
      fontSize: 11, fontFace: "Georgia", bold: true,
      color: S.ink, align: "center", valign: "middle", margin: 0,
    });
    slide.addText(safe(hy), {
      x: rx + 0.14, y: cy + 0.06, w: colW - 0.90, h: ch - 0.12,
      fontSize: 11.5, fontFace: S.af,
      color: S.textL, align: "right", rtlMode: true, valign: "middle", margin: 0,
    });
  });

  return slide;
}

// ═══════════════════════════════════════════════════════
// SLIDE 05 — IMPORTANCE & REASONS
// ═══════════════════════════════════════════════════════
function slide05_Importance(prs, S, d) {
  const slide = prs.addSlide();
  slide.background = { color: S.ink };

  decoCircle(slide, S.accent, W - 4, 2, 4, 0.05);

  overline(slide, S, "IMPORTANCE  ·  أهمية البحث", MX, MY, W - MX * 2, "left");
  slide.addText("أهمية البحث وأسباب اختياره", {
    x: MX, y: MY + 0.26, w: W - MX * 2, h: 0.70,
    fontSize: 28, fontFace: S.af, bold: true,
    color: S.textL, align: "right", rtlMode: true, valign: "middle", margin: 0,
  });
  rule(slide, S.accent, MX, MY + 1.02, W - MX * 2, 0.04);

  const items = [
    { label: "الأهمية العلمية والعملية", val: d.importance, icon: "⭐" },
    { label: "أسباب اختيار الموضوع", val: d.reasons, icon: "🔍" },
  ].filter(it => it.val);

  const ph = items.length ? (H - MY - 1.22 - 0.26 * items.length) / items.length : 0;

  items.forEach((it, i) => {
    const py = MY + 1.22 + i * (ph + 0.26);
    const sc = S.chart[i % S.chart.length];

    slide.addShape("rect", { x: MX, y: py, w: W - MX * 2, h: ph, fill: { color: S.inkMid }, line: { type: "none" } });
    slide.addShape("rect", { x: MX, y: py, w: 0.32, h: ph, fill: { color: sc }, line: { type: "none" } });
    slide.addText(`${it.icon}  ${it.label}`, {
      x: MX + 0.46, y: py + 0.12, w: W - MX * 2 - 0.60, h: 0.60,
      fontSize: 15, fontFace: S.af, bold: true,
      color: sc, align: "right", rtlMode: true, margin: 0,
    });
    rule(slide, S.border, MX + 0.46, py + 0.84, W - MX * 2 - 0.60, 0.020);
    slide.addText(safe(it.val), {
      x: MX + 0.46, y: py + 0.96, w: W - MX * 2 - 0.60, h: ph - 1.08,
      fontSize: 13, fontFace: S.af,
      color: S.textL, align: "right", rtlMode: true, valign: "top", margin: 0,
    });
  });

  return slide;
}

// ═══════════════════════════════════════════════════════
// SLIDE 06 — THEORETICAL FRAMEWORK
// ═══════════════════════════════════════════════════════
function slide06_Theory(prs, S, d) {
  const slide = prs.addSlide();
  slide.background = { color: S.paper };

  overline(slide, S, "CADRE THÉORIQUE  ·  الإطار النظري", MX, MY, 5, "left");
  slide.addText("الإطار النظري والمفاهيمي", {
    x: MX, y: MY + 0.26, w: W - MX * 2, h: 0.62,
    fontSize: 24, fontFace: S.af, bold: true,
    color: S.textD, align: "right", rtlMode: true, margin: 0,
  });
  rule(slide, S.accent, MX, MY + 0.94, W - MX * 2, 0.025);

  const concepts = (d.concepts || []).filter(c => c.name).slice(0, 6);
  if (!concepts.length) return slide;

  const n = concepts.length;
  const cols = n <= 2 ? n : n <= 4 ? 2 : 3;
  const rows = Math.ceil(n / cols);
  const gx = 0.28, gy = 0.20;
  const gridY = MY + 1.10;
  const availH = H - gridY - MY * 0.6;
  const availW = W - MX * 2;
  const cw = (availW - gx * (cols - 1)) / cols;
  const ch = (availH - gy * (rows - 1)) / rows;

  concepts.forEach((c, i) => {
    const col = i % cols, row = Math.floor(i / cols);
    const cx = MX + col * (cw + gx);
    const cy = gridY + row * (ch + gy);
    const sc = S.chart[i % S.chart.length];

    slide.addShape("rect", { x: cx, y: cy, w: cw, h: ch, fill: { color: S.cardL }, line: { color: S.borderL, width: 0.4 }, shadow: mkShadow(8, 1.5, 0.08) });
    slide.addShape("rect", { x: cx, y: cy, w: cw, h: 0.055, fill: { color: sc }, line: { type: "none" } });
    slide.addText(safe(c.name), {
      x: cx + 0.16, y: cy + 0.10, w: cw - 0.32, h: 0.45,
      fontSize: 13, fontFace: S.af, bold: true,
      color: sc, align: "right", rtlMode: true, valign: "middle", margin: 0,
    });
    if (c.nameEn) {
      slide.addText(safe(c.nameEn), {
        x: cx + 0.16, y: cy + 0.10, w: cw - 0.32, h: 0.40,
        fontSize: 8.5, fontFace: S.bf, italic: true,
        color: S.textMl, align: "left", margin: 0,
      });
    }
    rule(slide, S.borderL, cx + 0.16, cy + 0.58, cw - 0.32, 0.020);
    slide.addText(safe(c.def || ""), {
      x: cx + 0.16, y: cy + 0.66, w: cw - 0.32, h: ch - 0.78,
      fontSize: 11, fontFace: S.af,
      color: S.textD, align: "right", rtlMode: true, valign: "top", margin: 0,
    });
  });

  return slide;
}

// ═══════════════════════════════════════════════════════
// SLIDE 07 — METHODOLOGY
// ═══════════════════════════════════════════════════════
function slide07_Methodology(prs, S, d) {
  const slide = prs.addSlide();
  slide.background = { color: S.ink };

  const headerH = 0.88;
  slide.addShape("rect", { x: 0, y: 0, w: W, h: headerH, fill: { color: S.inkMid }, line: { type: "none" } });
  overline(slide, S, "MÉTHODOLOGIE", MX, 0.08, 5, "left");
  slide.addText("منهجية وأدوات البحث", {
    x: MX, y: 0.08, w: W - MX * 2, h: 0.72,
    fontSize: 22, fontFace: S.af, bold: true,
    color: S.textL, align: "right", rtlMode: true, valign: "middle", margin: 0,
  });

  // Process steps rail
  const steps = (d.steps || [
    { num: 1, titleAr: "المراجعة النظرية", descAr: "الإطار النظري" },
    { num: 2, titleAr: "جمع البيانات", descAr: d.dataSource || "البيانات" },
    { num: 3, titleAr: "الاختبارات", descAr: (d.statisticalTests || []).join(" · ") || "الاختبارات الإحصائية" },
    { num: 4, titleAr: "التحليل", descAr: d.software || "برنامج التحليل" },
    { num: 5, titleAr: "النتائج", descAr: "التفسير والتوصيات" },
  ]).slice(0, 5);

  const stepsN = steps.length;
  const railY = MY + 1.22;
  const nodeY = railY + 1.35;
  const availW = W - MX * 2;
  const stepSpan = availW / stepsN;
  const nodeR = 0.34;

  slide.addShape("rect", {
    x: MX + stepSpan * 0.5, y: nodeY - 0.022,
    w: availW - stepSpan, h: 0.044,
    fill: { color: S.border }, line: { type: "none" },
  });

  steps.forEach((step, i) => {
    const sc = S.chart[i % S.chart.length];
    const cx = MX + stepSpan * 0.5 + i * stepSpan;
    const isTop = i % 2 === 0;
    const cardH = 1.05;
    const cardW = stepSpan - 0.24;
    const cardX = cx - cardW / 2;

    if (i > 0) {
      slide.addShape("rect", {
        x: MX + stepSpan * 0.5 + (i - 1) * stepSpan, y: nodeY - 0.022,
        w: stepSpan, h: 0.044,
        fill: { color: sc, transparency: 50 }, line: { type: "none" },
      });
    }

    slide.addShape("ellipse", { x: cx - nodeR - 0.07, y: nodeY - nodeR - 0.07, w: (nodeR + 0.07) * 2, h: (nodeR + 0.07) * 2, fill: { color: S.inkMid }, line: { color: sc, width: 2 } });
    slide.addShape("ellipse", { x: cx - nodeR, y: nodeY - nodeR, w: nodeR * 2, h: nodeR * 2, fill: { color: sc }, line: { type: "none" } });
    slide.addText(String(step.num || i + 1), {
      x: cx - nodeR, y: nodeY - nodeR, w: nodeR * 2, h: nodeR * 2,
      fontSize: 18, fontFace: "Georgia", bold: true,
      color: S.ink, align: "center", valign: "middle", margin: 0,
    });

    const cardY = isTop ? nodeY - nodeR - 0.10 - cardH : nodeY + nodeR + 0.10;
    slide.addShape("rect", { x: cardX, y: cardY, w: cardW, h: cardH, fill: { color: S.cardD }, line: { color: sc, width: 0.6 } });
    slide.addText(safe(step.titleAr), {
      x: cardX + 0.10, y: cardY + 0.05, w: cardW - 0.20, h: 0.44,
      fontSize: 11, fontFace: S.af, bold: true,
      color: sc, align: "right", rtlMode: true, valign: "middle", margin: 0,
    });
    slide.addText(safe(step.descAr || ""), {
      x: cardX + 0.10, y: cardY + 0.50, w: cardW - 0.20, h: 0.50,
      fontSize: 9.5, fontFace: S.af,
      color: S.textM, align: "right", rtlMode: true, valign: "top", margin: 0,
    });
  });

  // Metadata row
  const mY = H - 1.82;
  rule(slide, S.border, MX, mY - 0.12, W - MX * 2, 0.020);

  const metaItems = [
    { label: "المنهج", val: safe(d.methodology, "—") },
    { label: "مصدر البيانات", val: safe(d.dataSource, "—") },
    { label: "الفترة", val: safe(d.timePeriod, "—") },
    { label: "البرنامج", val: safe(d.software, "—") },
  ];
  const mW = (W - MX * 2 - GU * 3) / 4;

  metaItems.forEach((m, i) => {
    const mx = MX + i * (mW + GU);
    const sc = S.chart[i % S.chart.length];
    slide.addShape("rect", { x: mx, y: mY, w: mW, h: 1.62, fill: { color: S.inkMid }, line: { type: "none" } });
    slide.addShape("rect", { x: mx, y: mY, w: mW, h: 0.045, fill: { color: sc }, line: { type: "none" } });
    slide.addText(m.label, {
      x: mx + 0.12, y: mY + 0.08, w: mW - 0.24, h: 0.32,
      fontSize: 9, fontFace: S.af, bold: true,
      color: sc, align: "right", valign: "middle", margin: 0,
    });
    slide.addText(m.val, {
      x: mx + 0.12, y: mY + 0.44, w: mW - 0.24, h: 1.10,
      fontSize: 11, fontFace: S.af,
      color: S.textL, align: "right", rtlMode: true, valign: "top", margin: 0,
    });
  });

  return slide;
}

// ═══════════════════════════════════════════════════════
// SLIDE 08 — KPI DASHBOARD
// ═══════════════════════════════════════════════════════
function slide08_KPI(prs, S, d) {
  const slide = prs.addSlide();
  slide.background = { color: S.ink };

  const headerH = 0.78;
  slide.addShape("rect", { x: 0, y: 0, w: W, h: headerH, fill: { color: S.inkMid }, line: { type: "none" } });
  slide.addText("النتائج الكمية — لوحة المؤشرات", {
    x: MX, y: 0, w: W - MX * 2, h: headerH,
    fontSize: 22, fontFace: S.af, bold: true,
    color: S.textL, align: "right", rtlMode: true, valign: "middle", margin: 0,
  });
  overline(slide, S, "KPI DASHBOARD  ·  المؤشرات الإحصائية", MX, 0.08, W / 2, "left");

  const stats = (d.stats || []).filter(s => s.label && s.value).slice(0, 8);
  if (!stats.length) return slide;

  const insightAr = d.insightAr || d.generalConclusion;
  const insightH = insightAr ? 0.72 : 0;
  const n = stats.length;
  const cols = Math.min(n, 4);
  const rows = Math.ceil(n / cols);
  const gx = 0.20, gy = 0.18;
  const gridY = headerH + 0.20;
  const availH = H - gridY - insightH - 0.22;
  const availW = W - MX * 2;
  const cw = (availW - gx * (cols - 1)) / cols;
  const ch = (availH - gy * (rows - 1)) / rows;

  stats.forEach((s, i) => {
    const col = i % cols, row = Math.floor(i / cols);
    const cx = MX + col * (cw + gx);
    const cy = gridY + row * (ch + gy);
    const sc = S.chart[i % S.chart.length];

    slide.addShape("rect", { x: cx, y: cy, w: cw, h: ch, fill: { color: S.cardD }, line: { type: "none" }, shadow: mkShadow(12, 2, 0.15) });
    slide.addShape("rect", { x: cx, y: cy, w: cw, h: 0.06, fill: { color: sc }, line: { type: "none" } });
    slide.addShape("rect", { x: cx, y: cy + ch - 0.04, w: cw, h: 0.04, fill: { color: sc, transparency: 60 }, line: { type: "none" } });

    const valStr = safe(s.value, "—");
    const valSz = valStr.length > 5 ? T.statS.sz : T.stat.sz;
    slide.addText(valStr, {
      x: cx + 0.12, y: cy + 0.12, w: cw - 0.24, h: ch * 0.58,
      fontSize: valSz, fontFace: "Georgia", bold: true,
      color: sc, align: "center", valign: "middle", margin: 0,
    });
    rule(slide, S.border, cx + 0.20, cy + ch * 0.63, cw - 0.40, 0.018);
    slide.addText(safe(s.label), {
      x: cx + 0.12, y: cy + ch * 0.67, w: cw - 0.24, h: ch * 0.24,
      fontSize: 10.5, fontFace: S.af,
      color: S.textM, align: "center", valign: "middle", margin: 0,
    });
    if (s.sub) {
      slide.addText(safe(s.sub), {
        x: cx + 0.12, y: cy + ch * 0.88, w: cw - 0.24, h: ch * 0.12,
        fontSize: 8, fontFace: S.bf, italic: true,
        color: S.border, align: "center", margin: 0,
      });
    }
  });

  if (insightAr) {
    const iy = H - insightH - 0.04;
    slide.addShape("rect", { x: MX, y: iy, w: W - MX * 2, h: insightH, fill: { color: S.inkMid }, line: { type: "none" } });
    slide.addShape("rect", { x: MX, y: iy, w: 0.08, h: insightH, fill: { color: S.accent }, line: { type: "none" } });
    slide.addShape("ellipse", { x: MX + 0.14, y: iy + insightH / 2 - 0.18, w: 0.36, h: 0.36, fill: { color: S.accent }, line: { type: "none" } });
    slide.addText("!", {
      x: MX + 0.14, y: iy + insightH / 2 - 0.18, w: 0.36, h: 0.36,
      fontSize: 15, fontFace: "Georgia", bold: true,
      color: S.ink, align: "center", valign: "middle", margin: 0,
    });
    slide.addText(safe(insightAr), {
      x: MX + 0.62, y: iy + 0.06, w: W - MX * 2 - 0.72, h: insightH - 0.12,
      fontSize: 12, fontFace: S.af,
      color: S.textL, align: "right", rtlMode: true, valign: "middle", margin: 0,
    });
  }

  return slide;
}

// ═══════════════════════════════════════════════════════
// SLIDE 09 — RESULTS
// ═══════════════════════════════════════════════════════
function slide09_Results(prs, S, d) {
  const slide = prs.addSlide();
  slide.background = { color: S.paperS };

  slide.addShape("rect", { x: 0, y: 0, w: W, h: 0.08, fill: { color: S.accent }, line: { type: "none" } });
  overline(slide, S, "RÉSULTATS  ·  نتائج البحث", MX, 0.15, 8, "left");
  slide.addText("نتائج البحث التفصيلية", {
    x: MX, y: 0.32, w: W - MX * 2, h: 0.55,
    fontSize: 22, fontFace: S.af, bold: true,
    color: S.textD, align: "right", rtlMode: true, margin: 0,
  });

  const results = (d.mainResults || []).filter(Boolean).slice(0, 7);
  if (!results.length) return slide;

  const tableY = 0.94;
  const availH = H - tableY - MY;
  const gap = 0.12;
  const rh = (availH - gap * results.length) / results.length;

  results.forEach((res, i) => {
    const ry = tableY + i * (rh + gap);
    const sc = S.chart[i % S.chart.length];

    slide.addShape("rect", {
      x: MX, y: ry, w: W - MX * 2, h: rh,
      fill: { color: i % 2 === 0 ? "FFFFFF" : S.paperS },
      line: { color: S.borderL, width: 0.4 },
      shadow: mkShadow(5, 1, 0.06),
    });
    slide.addShape("rect", { x: MX, y: ry, w: 0.055, h: rh, fill: { color: sc }, line: { type: "none" } });

    slide.addText(String(i + 1).padStart(2, "0"), {
      x: MX + 0.10, y: ry + 0.04, w: 0.60, h: rh - 0.08,
      fontSize: Math.min(22, rh * 18), fontFace: "Georgia", bold: true,
      color: sc, align: "center", valign: "middle", margin: 0,
    });
    slide.addShape("rect", { x: MX + 0.78, y: ry + rh * 0.2, w: 0.025, h: rh * 0.6, fill: { color: sc }, line: { type: "none" } });
    slide.addText(safe(res), {
      x: MX + 0.90, y: ry + 0.06, w: W - MX * 2 - 1.05, h: rh - 0.12,
      fontSize: Math.min(12.5, rh * 10), fontFace: S.af,
      color: S.textD, align: "right", rtlMode: true, valign: "middle", margin: 0,
    });
  });

  return slide;
}

// ═══════════════════════════════════════════════════════
// SLIDE 10 — LITERATURE TABLE
// ═══════════════════════════════════════════════════════
function slide10_Literature(prs, S, d) {
  const slide = prs.addSlide();
  slide.background = { color: S.paperS };

  slide.addShape("rect", { x: 0, y: 0, w: W, h: 0.08, fill: { color: S.accent }, line: { type: "none" } });
  overline(slide, S, "REVUE DE LITTÉRATURE  ·  مراجعة الأدبيات", MX, 0.15, 8, "left");
  slide.addText("مراجعة الدراسات السابقة", {
    x: MX, y: 0.32, w: W - MX * 2, h: 0.55,
    fontSize: 22, fontFace: S.af, bold: true,
    color: S.textD, align: "right", rtlMode: true, margin: 0,
  });

  const lits = (d.literatures || []).filter(l => l.title || l.author).slice(0, 5);
  if (!lits.length) return slide;

  const tableY = 0.94;
  const findingsH = 0.72;
  const tableH = H - tableY - findingsH - MY * 0.6;

  const headers = ["الباحث / المؤلف", "السنة", "عنوان الدراسة", "أبرز النتائج"];
  const rows = lits.map(l => [safe(l.author), safe(l.year), safe(l.title), safe(l.findings)]);

  const headerRow = headers.map((h, i) => ({
    text: safe(h),
    options: {
      fill: { color: i === 0 ? S.ink : S.chart[i % S.chart.length] },
      color: i === 0 ? S.textL : S.ink,
      bold: true, fontSize: 10.5, fontFace: S.af, align: "center", valign: "middle",
    },
  }));

  const dataRows = rows.map((row, ri) =>
    row.map(cell => ({
      text: safe(cell),
      options: {
        fill: { color: ri % 2 === 0 ? "FFFFFF" : S.paperS },
        color: S.textD,
        fontSize: 10.5, fontFace: S.af, align: "center",
      },
    }))
  );

  slide.addTable([headerRow, ...dataRows], {
    x: MX, y: tableY, w: W - MX * 2, h: tableH,
    border: { color: S.borderL.replace("#", ""), pt: 0.4 },
    autoPage: false,
    rowH: Math.min(0.52, tableH / (rows.length + 1.5)),
  });

  // Findings pills
  const findings = (d.findings || []).filter(Boolean).slice(0, 4);
  if (findings.length) {
    const fy = H - findingsH - 0.08;
    rule(slide, S.accent, MX, fy - 0.06, W - MX * 2, 0.020);
    const fw = (W - MX * 2 - GU * (findings.length - 1)) / findings.length;
    findings.forEach((f, i) => {
      const fx = MX + i * (fw + GU);
      const sc = S.chart[i % S.chart.length];
      slide.addShape("rect", { x: fx, y: fy, w: fw, h: findingsH - 0.08, fill: { color: sc }, line: { type: "none" } });
      slide.addText(safe(f), {
        x: fx + 0.12, y: fy + 0.04, w: fw - 0.24, h: findingsH - 0.16,
        fontSize: 10, fontFace: S.af, bold: true,
        color: S.ink, align: "right", rtlMode: true, valign: "middle", margin: 0,
      });
    });
  }

  return slide;
}

// ═══════════════════════════════════════════════════════
// SLIDE 11 — RECOMMENDATIONS
// ═══════════════════════════════════════════════════════
function slide11_Recommendations(prs, S, d) {
  const slide = prs.addSlide();
  slide.background = { color: S.ink };

  decoCircle(slide, S.accent, W - 3, 2, 5, 0.04);
  overline(slide, S, "RECOMMANDATIONS  ·  التوصيات", MX, MY, W - MX * 2, "left");
  slide.addText("التوصيات", {
    x: MX, y: MY + 0.26, w: W - MX * 2, h: 0.70,
    fontSize: 28, fontFace: S.af, bold: true,
    color: S.textL, align: "right", rtlMode: true, valign: "middle", margin: 0,
  });
  rule(slide, S.accent, MX, MY + 1.02, W - MX * 2, 0.04);

  const recs = (d.recommendations || []).filter(Boolean).slice(0, 6);
  if (!recs.length) return slide;

  const recY = MY + 1.22;
  const recH = H - recY - MY * 0.5;
  const cols = recs.length > 3 ? 2 : 1;
  const rows = Math.ceil(recs.length / cols);
  const gapx = 0.40, gapy = 0.22;
  const rw = (W - MX * 2 - gapx * (cols - 1)) / cols;
  const ch = (recH - gapy * (rows - 1)) / rows;

  recs.forEach((rec, i) => {
    const col = i % cols, row = Math.floor(i / cols);
    const rx = MX + col * (rw + gapx);
    const ry = recY + row * (ch + gapy);
    const sc = S.chart[i % S.chart.length];

    slide.addShape("rect", { x: rx, y: ry, w: rw, h: ch, fill: { color: S.inkMid }, line: { type: "none" } });
    slide.addShape("rect", { x: rx, y: ry, w: 0.25, h: ch, fill: { color: sc }, line: { type: "none" } });

    slide.addText(String(i + 1).padStart(2, "0"), {
      x: rx + 0.34, y: ry + 0.08, w: 1.44, h: ch * 0.40,
      fontSize: 26, fontFace: "Georgia", bold: true,
      color: sc, align: "left", valign: "middle", margin: 0,
    });
    rule(slide, S.border, rx + 0.34, ry + ch * 0.42, rw - 0.50, 0.020);
    slide.addText(safe(rec), {
      x: rx + 0.34, y: ry + ch * 0.48, w: rw - 0.50, h: ch * 0.48,
      fontSize: 12, fontFace: S.af,
      color: S.textL, align: "right", rtlMode: true, valign: "top", margin: 0,
    });
  });

  return slide;
}

// ═══════════════════════════════════════════════════════
// SLIDE 12 — CONCLUSION & THANK YOU
// ═══════════════════════════════════════════════════════
function slide12_Conclusion(prs, S, d) {
  const slide = prs.addSlide();
  slide.background = { color: S.paper };

  const topH = H * 0.42;
  slide.addShape("rect", { x: 0, y: 0, w: W, h: topH, fill: { color: S.ink }, line: { type: "none" } });
  slide.addShape("rect", { x: 0, y: 0, w: W, h: 0.10, fill: { color: S.accent }, line: { type: "none" } });
  decoCircle(slide, S.accent, 1.2, topH * 0.5, 2.8, 0.04);

  overline(slide, S, "CONCLUSION  ·  الخاتمة", MX, 0.18, W / 2, "left");
  slide.addText("\u201C", {
    x: MX, y: 0.28, w: 2.0, h: 1.6,
    fontSize: 110, fontFace: S.hf,
    color: S.accent, align: "left", valign: "top", margin: 0,
  });
  slide.addText(safe(d.generalConclusion, "خلاصة البحث"), {
    x: MX + 0.50, y: 0.52, w: W - MX * 2 - 0.50, h: topH - 0.68,
    fontSize: 15.5, fontFace: S.af,
    color: S.textL, align: "right", rtlMode: true,
    valign: "middle", paraSpaceAfter: 5, margin: 0,
  });
  slide.addShape("rect", { x: 0, y: topH - 0.055, w: W, h: 0.055, fill: { color: S.accent }, line: { type: "none" } });

  // Recommendations mid zone
  const recY = topH + 0.22;
  const recH = H - recY - 0.92;
  const recs = (d.recommendations || []).filter(Boolean).slice(0, 4);
  const recN = recs.length || 1;
  const rw = (W - MX * 2 - GU * (recN - 1)) / recN;

  recs.forEach((rec, i) => {
    const rx = MX + i * (rw + GU);
    const sc = S.chart[i % S.chart.length];
    slide.addShape("rect", { x: rx, y: recY, w: rw, h: recH, fill: { color: S.cardL }, line: { color: S.borderL, width: 0.4 }, shadow: mkShadow(8, 1.5, 0.09) });
    slide.addShape("rect", { x: rx, y: recY, w: rw, h: 0.06, fill: { color: sc }, line: { type: "none" } });
    slide.addText(String(i + 1).padStart(2, "0"), {
      x: rx + 0.10, y: recY + 0.10, w: rw - 0.20, h: 0.52,
      fontSize: 30, fontFace: "Georgia", bold: true,
      color: sc, align: "center", valign: "middle", margin: 0,
    });
    rule(slide, S.borderL, rx + 0.16, recY + 0.66, rw - 0.32, 0.018);
    slide.addText(safe(rec), {
      x: rx + 0.12, y: recY + 0.72, w: rw - 0.24, h: recH - 0.82,
      fontSize: 11, fontFace: S.af,
      color: S.textD, align: "right", rtlMode: true, valign: "top", margin: 0,
    });
  });

  // Thank-you bar
  const barY = H - 0.85;
  slide.addShape("rect", { x: 0, y: barY, w: W, h: 0.85, fill: { color: S.ink }, line: { type: "none" } });
  slide.addShape("rect", { x: 0, y: barY, w: W, h: 0.042, fill: { color: S.accent }, line: { type: "none" } });
  slide.addText("شكراً لحسن استماعكم", {
    x: MX, y: barY + 0.08, w: W / 2, h: 0.38,
    fontSize: 13, fontFace: S.af, bold: true,
    color: S.textL, align: "right", rtlMode: true, valign: "middle", margin: 0,
  });
  slide.addText("Merci pour votre attention", {
    x: MX, y: barY + 0.46, w: W / 2, h: 0.28,
    fontSize: 10, fontFace: S.bf, italic: true,
    color: S.textM, align: "left", margin: 0,
  });
  slide.addText(safe(d.studentName) + (d.year ? "  ·  " + safe(d.year) : ""), {
    x: W / 2, y: barY + 0.16, w: W / 2 - MX, h: 0.52,
    fontSize: 13, fontFace: S.af, bold: true,
    color: S.accent, align: "right", rtlMode: true, valign: "middle", margin: 0,
  });

  return slide;
}

// ═══════════════════════════════════════════════════════
// ORCHESTRATOR
// ═══════════════════════════════════════════════════════
async function buildPresentation(data) {
  const styleKey = THEME_MAP[data.theme] || "noir";
  const S = STYLES[styleKey];

  const prs = new PptxGenJS();
  prs.layout = "LAYOUT_WIDE";
  prs.title = `مذكرتي Pro — ${safe(data.titleAr, "مذكرة تخرج")}`;
  prs.author = "مذكرتي Pro v3";

  const chapters = [
    { title: "الإشكالية والتساؤلات", sub: "Research Problem" },
    { title: "الأهداف والفرضيات",    sub: "Objectives & Hypotheses" },
    (data.importance || data.reasons) && { title: "أهمية البحث", sub: "Significance & Motivation" },
    (data.concepts || []).filter(c => c.name).length && { title: "الإطار النظري", sub: "Theoretical Framework" },
    { title: "المنهجية والأدوات",    sub: "Methodology & Tools" },
    (data.stats || []).filter(s => s.value && s.label).length && { title: "النتائج الكمية", sub: "KPI Dashboard" },
    (data.mainResults || []).filter(Boolean).length && { title: "نتائج البحث", sub: "Research Findings" },
    (data.literatures || []).filter(l => l.title || l.author).length && { title: "الدراسات السابقة", sub: "Literature Review" },
    (data.recommendations || []).filter(Boolean).length && { title: "التوصيات", sub: "Recommendations" },
    { title: "الخاتمة", sub: "Conclusion" },
  ].filter(Boolean);

  // Build slides
  await slide01_Cover(prs, S, data);
  slide02_TOC(prs, S, chapters);
  slide03_Problem(prs, S, data);
  slide04_Objectives(prs, S, data);

  if (data.importance || data.reasons) {
    slide05_Importance(prs, S, data);
  }

  const concepts = (data.concepts || []).filter(c => c.name);
  if (concepts.length) {
    slide06_Theory(prs, S, { concepts });
  }

  slide07_Methodology(prs, S, data);

  const stats = (data.stats || []).filter(s => s.value && s.label);
  if (stats.length) {
    slide08_KPI(prs, S, { stats, insightAr: data.generalConclusion, generalConclusion: data.generalConclusion });
  }

  if ((data.mainResults || []).filter(Boolean).length) {
    slide09_Results(prs, S, data);
  }

  const lits = (data.literatures || []).filter(l => l.title || l.author);
  if (lits.length) {
    slide10_Literature(prs, S, {
      literatures: lits,
      findings: data.findings || [],
    });
  }

  if ((data.recommendations || []).filter(Boolean).length) {
    slide11_Recommendations(prs, S, data);
  }

  slide12_Conclusion(prs, S, data);

  // Write to stdout as base64
  const buf = await prs.write({ outputType: "nodebuffer" });
  process.stdout.write(buf);

  const n = prs.slides.length;
  process.stderr.write(`✅  [${S.name}]  ${n} slides  →  stdout\n`);
}

// ═══════════════════════════════════════════════════════
// ENTRY POINT — read JSON from stdin
// ═══════════════════════════════════════════════════════
let raw = "";
process.stdin.setEncoding("utf8");
process.stdin.on("data", chunk => { raw += chunk; });
process.stdin.on("end", () => {
  try {
    const data = JSON.parse(raw);
    buildPresentation(data).catch(e => {
      process.stderr.write("❌ ERROR: " + e.message + "\n" + e.stack + "\n");
      process.exit(1);
    });
  } catch (e) {
    process.stderr.write("❌ JSON parse error: " + e.message + "\n");
    process.exit(1);
  }
});
