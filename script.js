// -----------------------------------------------------------
// script.js - Génération du PowerPoint (PptxGenJS v3.x)
// -----------------------------------------------------------

document.addEventListener("DOMContentLoaded", () => {
  console.log("✅ script.js chargé – PptxGenJS présent ?", typeof PptxGenJS !== "undefined");

  const btn = document.getElementById("exportBtn");
  if (btn) {
    btn.addEventListener("click", createPowerPoint);
  }
});

function createPowerPoint() {
  const btn = document.getElementById("exportBtn");
  btn?.setAttribute("disabled", "true");
  btn?.setAttribute("aria-busy", "true");

  if (typeof PptxGenJS === "undefined") {
    alert("❌ PptxGenJS n'est pas chargé (vérifie bien <script src=\"pptxgen.bundle.js\"></script>).");
    btn?.removeAttribute("aria-busy");
    btn?.removeAttribute("disabled");
    return;
  }

  const pptx = new PptxGenJS();
  pptx.layout = "LAYOUT_WIDE"; // 16:9

  // Types de formes (utilisation de l'instance pptx)
  const RECT    = pptx.shapes.RECTANGLE;
  const ELLIPSE = pptx.shapes.OVAL;

  // --- util ---
  const getVal = (id) => document.getElementById(id)?.value || "";

  // --- Champs formulaire ---
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

  // ------------------ SLIDE 1 : Couverture ------------------
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

  // --------------- SLIDE 2 : Compléments d’informations ---------------
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
  const IMG = { x: MARGIN, y: 1.4, w: 5.8, h: 3.8 };
  // Zone texte commentaire (droite)
  const BOX = { x: 6.6,   y: 1.4, w: SLIDE_W - 6.6 - MARGIN, h: 3.8 };

  // Tailles / positions des formes (relatives à l'image)
  const TGBT_RECT = { w: 1.6, h: 1.1, dx: 1.0, dy: 0.8 };  // rectangle rouge
  const BORNE_RECT = { w: 1.6, h: 1.1, dx: 3.2, dy: 2.0 }; // rectangle bleu

  const GREEN_CIRCLE = {
    w: 1.8,
    h: 1.8,
    stroke: "00FF00",
    strokeWidth: 3,
    x: BOX.x + 0.2,
    y: BOX.y + BOX.h + 0.15   // sous la zone de texte, à droite
  };

  // ---------------- Légende bas-droite ----------------
  function addLegend(slide, texts = []) {
    if (!texts || texts.length === 0) return;

    const LEG_W = 3.5;
    const LEG_H = 1.0;

    const x = SLIDE_W - LEG_W - 0.2; // coin bas droite
    const y = SLIDE_H - LEG_H - 0.2;

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
    // Image
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
      // Rectangle rouge = TGBT
      slide.addShape(RECT, {
        x: IMG.x + TGBT_RECT.dx,
        y: IMG.y + TGBT_RECT.dy,
        w: TGBT_RECT.w,
        h: TGBT_RECT.h,
        fill: { color: "FF0000" },
        line: { color: "880000" }
      });

      // Cercle vert = place à équiper
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
  function addChecklistSlides() {
    // Numéro de section par rubrique
    const sectionNumber = {
      "Plan d'implantation": 1,
      "Places à électrifier": 2,
      "TGBT + disjoncteur de tête": 3,
      "Cheminement": 4,
      "Plan du site": 5,
      "Éléments complémentaires": 6
    };

    // 6 rubriques × 2 photos/commentaires
    const slideDefs = [
      { title: "Plan d'implantation",        fileId: "file1",  commentId: "comment1"  },
      { title: "Plan d'implantation",        fileId: "file1b", commentId: "comment1b" },

      { title: "Places à électrifier",       fileId: "file2",  commentId: "comment2"  },
      { title: "Places à électrifier",       fileId: "file2b", commentId: "comment2b" },

      { title: "TGBT + disjoncteur de tête", fileId: "file3",  commentId: "comment3"  },
      { title: "TGBT + disjoncteur de tête", fileId: "file3b", commentId: "comment3b" },

      { title: "Cheminement",                fileId: "file4",  commentId: "comment4"  },
      { title: "Cheminement",                fileId: "file4b", commentId: "comment4b" },

      { title: "Plan du site",               fileId: "file5",  commentId: "comment5"  },
      { title: "Plan du site",               fileId: "file5b", commentId: "comment5b" },

      { title: "Éléments complémentaires",   fileId: "file6",  commentId: "comment6"  },
      { title: "Éléments complémentaires",   fileId: "file6b", commentId: "comment6b" }
    ];

    let remaining = slideDefs.length;

    function finalizeIfDone() {
      if (remaining === 0) {
        const safeName = (clientName || "Projet")
          .replace(/[^\p{L}\p{N}_\- ]/gu, "")
          .trim()
          .replace(/\s+/g, "_") || "Projet";

        pptx
          .writeFile({ fileName: `Borne_Electrique_${safeName}.pptx` })
          .finally(() => {
            btn?.removeAttribute("aria-busy");
            btn?.removeAttribute("disabled");
          });
      }
    }

    slideDefs.forEach((def) => {
      const slide = pptx.addSlide();

      // Titre de section : N. Titre
      const secNum = sectionNumber[def.title] || "";
      slide.addText(`${secNum}. ${def.title}`, {
        x: 0.5,
        y: 0.3,
        w: SLIDE_W - 1.0,
        fontSize: 36,
        bold: true,
        color: "0070C0",
        align: "center"
      });

      // Commentaire à droite
      const commentText = getVal(def.commentId) || "—";
      slide.addText(commentText, {
        x: BOX.x,
        y: BOX.y,
        w: BOX.w,
        h: BOX.h,
        fill: { color: "FFFFFF" },
        line: { color: "AAAAAA" },
        margin: 0.12,
        fontSize: 18,
        color: "111111",
        align: "left",
        valign: "top"
      });

      const fileInput = document.getElementById(def.fileId);

      const finishSlide = (dataUrl) => {
        placeImageAndShapes(slide, def.title, IMG, dataUrl);
        remaining--;
        finalizeIfDone();
      };

      if (fileInput && fileInput.files && fileInput.files.length > 0) {
        const reader = new FileReader();
        reader.onload = (e) => {
          finishSlide(e.target.result);
        };
        reader.readAsDataURL(fileInput.files[0]);
      } else {
        finishSlide(null);
      }
    });
  }

  // ---------------- Lancement de la génération ----------------
  try {
    addCoverSlide();
    addInfoSlide();
    addChecklistSlides();
  } catch (err) {
    console.error("❌ Erreur lors de la génération PPTX :", err);
    alert("Erreur lors de la génération du PowerPoint. Regarde la console pour plus de détails.");
    btn?.removeAttribute("aria-busy");
    btn?.removeAttribute("disabled");
  }
}
