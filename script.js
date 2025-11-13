// -----------------------------------------------------------
// script.js - Génération du PowerPoint (PptxGenJS v3.x)
// -----------------------------------------------------------

document.addEventListener("DOMContentLoaded", () => {
  console.log("✅ script.js chargé");

  document.getElementById("exportBtn")?.addEventListener("click", createPowerPoint);
});

function createPowerPoint() {
  const btn = document.getElementById("exportBtn");
  btn?.setAttribute("disabled", "true");
  btn?.setAttribute("aria-busy", "true");

  if (typeof PptxGenJS === "undefined") {
    alert("❌ PptxGenJS n'est pas chargé.");
    btn?.removeAttribute("aria-busy");
    btn?.removeAttribute("disabled");
    return;
  }

  const pptx = new PptxGenJS();
  pptx.layout = "LAYOUT_WIDE"; // 16:9

  // Types de formes
  const RECT    = PptxGenJS.ShapeType?.rect    || "rect";
  const ELLIPSE = PptxGenJS.ShapeType?.ellipse || "ellipse";

  // --- util ---
  const getVal = (id) => document.getElementById(id)?.value || "";

  // --- Champs saisis par l'utilisateur ---
  const clientName    = getVal("clientName");
  const rae           = getVal("rae");
  const power         = getVal("power");
  const commercial    = getVal("commercial");
  const raeClient     = getVal("raeClient");

  const clientAddress = getVal("clientAddress");
  const siret         = getVal("siret");
  const oppoNumber    = getVal("oppoNumber");
  const nbBornes      = getVal("nbBornes");
  const bornesPower   = getVal("bornesPower");

  // ------------------ SLIDE 1 : Couverture avec infos client ------------------
  function addCoverSlide() {
    const slide = pptx.addSlide();
    slide.background = { fill: "363636" };

    const lines = [];

    if (clientName)    lines.push({ text: `Client : ${clientName}\n`,      options: { fontSize: 20, color: "FFFFFF", bold: true } });
    if (rae)           lines.push({ text: `RAE : ${rae}\n`,                options: { fontSize: 16, color: "FFFFFF" } });
    if (power)         lines.push({ text: `Puissance : ${power}\n`,        options: { fontSize: 16, color: "FFFFFF" } });
    if (commercial)    lines.push({ text: `Commercial : ${commercial}\n`,  options: { fontSize: 16, color: "FFFFFF" } });
    if (clientAddress) lines.push({ text: `Adresse : ${clientAddress}\n`,  options: { fontSize: 16, color: "FFFFFF" } });
    if (siret)         lines.push({ text: `SIRET : ${siret}\n`,            options: { fontSize: 16, color: "FFFFFF" } });
    if (oppoNumber)    lines.push({ text: `Numéro Oppo : ${oppoNumber}\n`, options: { fontSize: 16, color: "FFFFFF" } });
    if (nbBornes)      lines.push({ text: `Nombre de bornes : ${nbBornes}\n`,       options: { fontSize: 16, color: "FFFFFF" } });
    if (bornesPower)   lines.push({ text: `Puissance des bornes : ${bornesPower}\n`, options: { fontSize: 16, color: "FFFFFF" } });

    lines.push({
      text: "Projet d’infrastructure de recharge pour véhicules électriques",
      options: { fontSize: 14, color: "FFFFFF", italic: true, breakLine: true }
    });

    slide.addText(lines, {
      x: 0.5,
      y: 0.5,
      w: 5.8,
      h: 6.2
    });
  }

  // --------------- SLIDE 2 : Compléments d'infos ---------------
  function addInfoSlide() {
    const slide = pptx.addSlide();

    slide.addText("Compléments d’informations", {
      x: 0.5,
      y: 0.4,
      w: 9,
      fontSize: 36,
      bold: true,
      color: "0070C0",
      align: "center"
    });

    slide.addText(raeClient || "—", {
      x: 0.5,
      y: 1.5,
      w: "90%",
      h: "70%",
      fontSize: 18,
      fill: { color: "FFFFFF" },
      line: { color: "CCCCCC" }
    });
  }

  // ---------------- Mise en page générale ----------------
  const SLIDE_W = 10.0;
  const SLIDE_H = 5.625;
  const MARGIN  = 0.5;

  // Zone image (gauche)
  const IMG = { x: MARGIN, y: 1.6, w: 6.1, h: 4.3 };
  // Zone texte commentaire (droite)
  const BOX = { x: 6.7,   y: 1.6, w: SLIDE_W - 6.7 - MARGIN, h: 4.3 };

  // Tailles / positions des formes
  const TGBT_RECT = { w: 1.6, h: 1.1, dx: 0.8, dy: 0.6 };  // rouge
  const BORNE_RECT = { w: 1.6, h: 1.1, dx: 2.0, dy: 1.8 }; // bleu

  const GREEN_CIRCLE = {
    w: 1.8,
    h: 1.8,
    stroke: "00FF00",
    strokeWidth: 3,
    x: BOX.x, // à droite (sous le bloc de texte)
    y: Math.min(BOX.y + BOX.h + 0.2, SLIDE_H - MARGIN - 1.8)
  };

  // ---------------- Légende bas-droite ----------------
  function addLegend(slide, texts = []) {
    if (!texts || texts.length === 0) return;

    const LEG_W = 3.2;
    const LEG_H = 0.8;

    const x = SLIDE_W - LEG_W - 0.1;
    const y = SLIDE_H - LEG_H; // coin bas droit

    slide.addText(texts.join("\n"), {
      x,
      y,
      w: LEG_W,
      h: LEG_H,
      fontSize: 12,
      color: "000000",
      valign: "top"
    });
  }

  // ---------------- Image + formes + légende ----------------
  function placeImageAndShapes(slide, title, imgBox, dataUrl) {
    if (dataUrl) {
      slide.addImage({
        data: dataUrl,
        x: imgBox.x,
        y: imgBox.y,
        w: imgBox.w,
        h: imgBox.h,
        sizing: { type: "contain" }
      });
    }

    const low = title.toLowerCase();
    const legendTexts = [];

    // Plan d'implantation : rectangle rouge + cercle vert
    if (low.includes("implantation")) {
      // Rectangle rouge (TGBT)
      slide.addShape(RECT, {
        x: IMG.x + TGBT_RECT.dx,
        y: IMG.y + TGBT_RECT.dy,
        w: TGBT_RECT.w,
        h: TGBT_RECT.h,
        fill: { color: "FF0000" },
        line: { color: "880000" }
      });

      // Cercle vert (place à équiper)
      slide.addShape(ELLIPSE, {
        x: GREEN_CIRCLE.x,
        y: GREEN_CIRCLE.y,
        w: GREEN_CIRCLE.w,
        h: GREEN_CIRCLE.h,
        fill: null,
        line: { color: GREEN_CIRCLE.stroke, width: GREEN_CIRCLE.strokeWidth }
      });

      legendTexts.push(
        "Rectangle rouge = TGBT",
        "Cercle vert = place à équiper",
        "Rectangle bleu = future borne"
      );
    }

    // Places à électrifier : rectangle bleu
    if (low.includes("places à électrifier") || low.includes("places a electrifier")) {
      slide.addShape(RECT, {
        x: IMG.x + BORNE_RECT.dx,
        y: IMG.y + BORNE_RECT.dy,
        w: BORNE_RECT.w,
        h: BORNE_RECT.h,
        fill: { color: "0070C0" },
        line: { color: "003A70" }
      });

      legendTexts.push("Rectangle bleu = future borne");
    }

    addLegend(slide, legendTexts);
  }

  // ---------------- Slides checklist ----------------
  function addChecklistSlides(onAllSlidesReady) {

    // Numéro de section par rubrique
    const sectionNumber = {
      "Plan d'implantation": 1,
      "Places à électrifier": 2,
      "TGBT + disjoncteur de tête": 3,
      "Cheminement": 4,
      "Plan du site": 5,
      "Éléments complémentaires": 6
    };

    const items = [
      { base:"Plan d'implantation",        file:"file1",  comment:"comment1"  },
      { base:"Plan d'implantation",        file:"file1b", comment:"comment1b" },

      { base:"Places à élect
