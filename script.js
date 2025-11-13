// -----------------------------------------------------------
// script.js - Génération du PowerPoint (PptxGenJS v3.x)
// Version avec titres numérotés + style bleu + taille 36
// -----------------------------------------------------------

document.addEventListener("DOMContentLoaded", () => {
  console.log("✅ script.js chargé (titres améliorés)");

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
  pptx.layout = "LAYOUT_WIDE";

  const RECT    = PptxGenJS.ShapeType?.rect    || "rect";
  const ELLIPSE = PptxGenJS.ShapeType?.ellipse || "ellipse";

  // Champs du formulaire
  const getVal = (id) => document.getElementById(id)?.value || "";
  const clientName = getVal("clientName");
  const raeClient  = getVal("raeClient");

  // ---- SLIDE : Couverture ----
  function addCoverSlide() {
    const slide = pptx.addSlide();
    slide.background = { fill: "363636" };

    slide.addText(`Client : ${clientName}`, {
      x: 0.5, y: 0.6, fontSize: 38, color: "FFFFFF", bold: true
    });

    slide.addText("Projet d’infrastructure de recharge", {
      x: 0.5, y: 1.5, fontSize: 26, color: "FFFFFF"
    });
  }

  // ---- SLIDE : Informations complémentaires ----
  function addInfoSlide() {
    const slide = pptx.addSlide();

    slide.addText("Compléments d’informations", {
      x: 0.5, y: 0.4, fontSize: 36, bold: true, color: "0070C0", align: "center", w: 9
    });

    slide.addText(raeClient || "—", {
      x: 0.5, y: 1.5, w: "90%", h: "70%",
      fontSize: 18, fill: { color: "FFFFFF" }, line: { color: "CCCCCC" }
    });
  }

  // ---- Mise en page ----
  const SLIDE_W = 10.0;
  const SLIDE_H = 5.625;
  const MARGIN  = 0.5;

  const IMG = { x: MARGIN, y: 1.6, w: 6.1, h: 4.3 };
  const BOX = { x: 6.7,   y: 1.6, w: SLIDE_W - 6.7 - MARGIN, h: 4.3 };

  // ---- Formes ----
  const TGBT_RECT = { w:1.6, h:1.1, dx:0.8, dy:0.6 };
  const BORNE_RECT = { w:1.6, h:1.1, dx:2.0, dy:1.8 };

  const GREEN_CIRCLE = {
    w: 1.8, h: 1.8,
    stroke: "00FF00",
    strokeWidth: 3,
    x: BOX.x,
    y: Math.min(BOX.y + BOX.h + 0.2, SLIDE_H - MARGIN - 1.8)
  };

  // ---- Légende bas-droite ----
  function addLegend(slide, texts = []) {
    if (!texts || texts.length === 0) return;

    const LEG_W = 3.2;
    const LEG_H = 0.8;

    // Plus bas possible
    const x = SLIDE_W - LEG_W - 0.1;
    const y = SLIDE_H - LEG_H;

    slide.addText(texts.join("\n"), {
      x, y, w: LEG_W, h: LEG_H,
      fontSize: 12, color: "000000", valign: "top"
    });
  }

  // ---- Image + Formes + Légende ----
  function placeImageAndShapes(slide, title, imgBox, dataUrl) {
    if (dataUrl) {
      slide.addImage({
        data: dataUrl,
        x: imgBox.x, y: imgBox.y, w: imgBox.w, h: imgBox.h,
        sizing: { type: "contain" }
      });
    }

    const low = title.toLowerCase();
    const legendTexts = [];

    if (low.includes("implantation")) {
      slide.addShape(RECT, {
        x: IMG.x + TGBT_RECT.dx,
        y: IMG.y + TGBT_RECT.dy,
        w: TGBT_RECT.w, h: TGBT_RECT.h,
        fill: { color: "FF0000" }, line: { color: "880000" }
      });

      slide.addShape(ELLIPSE, {
        x: GREEN_CIRCLE.x,
        y: GREEN_CIRCLE.y,
        w: GREEN_CIRCLE.w, h: GREEN_CIRCLE.h,
        fill: null,
        line: { color: GREEN_CIRCLE.stroke, width: GREEN_CIRCLE.strokeWidth }
      });

      legendTexts.push(
        "Rectangle rouge = TGBT",
        "Cercle vert = place à équiper",
        "Rectangle bleu = future borne"
      );
    }

    if (low.includes("places à électrifier") || low.includes("places a electrifier")) {
      slide.addShape(RECT, {
        x: IMG.x + BORNE_RECT.dx,
        y: IMG.y + BORNE_RECT.dy,
        w: BORNE_RECT.w, h: BORNE_RECT.h,
        fill: { color: "0070C0" }, line: { color: "003A70" }
      });

      legendTexts.push("Rectangle bleu = future borne");
    }

    addLegend(slide, legendTexts);
  }

  // ---- Slides checklist ----
  function addChecklistSlides() {

    // Numérotation des sections :
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

      { base:"Places à électrifier",       file:"file2",  comment:"comment2"  },
      { base:"Places à électrifier",       file:"file2b", comment:"comment2b" },

      { base:"TGBT + disjoncteur de tête", file:"file3",  comment:"comment3"  },
      { base:"TGBT + disjoncteur de tête", file:"file3b", comment:"comment3b" },

      { base:"Cheminement",                file:"file4",  comment:"comment4"  },
      { base:"Cheminement",                file:"file4b", comment:"comment4b" },

      { base:"Plan du site",               file:"file5",  comment:"comment5"  },
      { base:"Plan du site",               file:"file5b", comment:"comment5b" },

      { base:"Éléments complémentaires",   file:"file6",  comment:"comment6"  },
      { base:"Éléments complémentaires",   file:"file6b", comment:"comment6b" }
    ];

    let countInSection = {};

    items.forEach((item) => {

      if (!countInSection[item.base]) countInSection[item.base] = 1;

      const index = countInSection[item.base]++;
      const num = sectionNumber[item.base];
      const title = `${num} – ${item.base} #${index}`;

      const slide = pptx.addSlide();

      // ---- TITRE FORMATÉ ----
      slide.addText(title, {
        x: 0.5,
        y: 0.3,
        w: 9,
        fontSize: 36,
        bold: true,
        color: "0070C0",
        align: "center"
      });

      // ---- Commentaire ----
      const commentText = document.getElementById(item.comment)?.value || "—";
      slide.addText(commentText, {
        x: BOX.x, y: BOX.y, w: BOX.w, h: BOX.h,
        fill: { color: "FFFFFF" }, line: { color: "AAAAAA" },
        fontSize: 18, valign: "top", margin: 0.1
      });

      // ---- Image ----
      const fileInput = document.getElementById(item.file);

      const afterImage = (dataUrl) =>
        placeImageAndShapes(slide, title, IMG, dataUrl);

      if (fileInput?.files.length > 0) {
        const r = new FileReader();
        r.onload = (e) => afterImage(e.target.result);
        r.readAsDataURL(fileInput.files[0]);
      } else {
        afterImage(null);
      }
    });
  }

  // ---- Lancement ----
  addCoverSlide();
  addInfoSlide();
  addChecklistSlides();

  btn?.removeAttribute("aria-busy");
  btn?.removeAttribute("disabled");
}
