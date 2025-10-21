// -----------------------------------------------------------
// script.js - Génération du PowerPoint (PptxGenJS v3.x)
// - Commentaires à droite des images (zone large déplaçable)
// - 2 slides par rubrique
// - "Compléments d’informations" au lieu de "RAE du client"
// - Rectangles SANS texte : rouge (TGBT) / bleu (bornes)
// - Cercle vert épais sur "Plan d’implantation"
// -----------------------------------------------------------

document.addEventListener("DOMContentLoaded", () => {
  console.log(
    "Pptx present ?",
    typeof PptxGenJS !== "undefined",
    "(PptxGenJS:",
    typeof PptxGenJS !== "undefined",
    ")"
  );
  if (typeof PptxGenJS === "undefined") {
    alert("❌ PptxGenJS n'est pas chargé. Vérifie que 'pptxgen.bundle.js' est bien à la racine.");
  }
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

  const getVal = (id) => document.getElementById(id)?.value || "";

  // --- Champs (infos couverture & complément) ---
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

  // ---------------- Slide d’accueil (infos client) ----------------
  (function addCoverSlide() {
    const slide = pptx.addSlide();
    slide.background = { fill: "363636" };

    const lines = [];
    if (clientName)     lines.push({ text: `Client : ${clientName}\n`,      options: { fontSize: 20, color: "FFFFFF", bold: true } });
    if (rae)            lines.push({ text: `RAE : ${rae}\n`,                options: { fontSize: 16, color: "FFFFFF" } });
    if (power)          lines.push({ text: `Puissance : ${power}\n`,        options: { fontSize: 16, color: "FFFFFF" } });
    if (commercial)     lines.push({ text: `Commercial : ${commercial}\n`,  options: { fontSize: 16, color: "FFFFFF" } });
    if (clientAddress)  lines.push({ text: `Adresse : ${clientAddress}\n`,  options: { fontSize: 16, color: "FFFFFF" } });
    if (siret)          lines.push({ text: `SIRET : ${siret}\n`,            options: { fontSize: 16, color: "FFFFFF" } });
    if (oppoNumber)     lines.push({ text: `Numéro Oppo : ${oppoNumber}\n`, options: { fontSize: 16, color: "FFFFFF" } });
    if (nbBornes)       lines.push({ text: `Nombre de bornes : ${nbBornes}\n`, options: { fontSize: 16, color: "FFFFFF" } });
    if (bornesPower)    lines.push({ text: `Puissance des bornes : ${bornesPower}\n`, options: { fontSize: 16, color: "FFFFFF" } });
    lines.push({
      text: "Projet d’infrastructure de recharge pour véhicules électriques",
      options: { fontSize: 14, color: "FFFFFF", italic: true, breakLine: true }
    });

    slide.addText(lines, { x: 0.7, y: 0.7, w: 8.6, h: 4.2 });
  })();

  // ---------------- Slide “Compléments d’informations” ----------------
  (function addInfoSlide() {
    const slide = pptx.addSlide();
    slide.addText("Compléments d’informations", { x: 0.7, y: 0.5, fontSize: 24, bold: true });
    slide.addText(raeClient || "—", {
      x: 0.7, y: 1.3, w: 8.6, h: 4.8,
      fontSize: 18, color: "363636",
      fill: { color: "FFFFFF" }, line: { color: "DDDDDD" }, margin: 0.12, valign: "top"
    });
  })();

  // ---------------- Mise en page image/texte ----------------
  const SLIDE_W = 10.0;
  const MARGIN  = 0.5;

  // Image à gauche
  const IMG = { x: MARGIN, y: 1.1, w: 6.1, h: 4.6 };
  // Commentaire à droite (grande zone déplaçable)
  const BOX = { x: 6.7, y: 1.1, w: SLIDE_W - 6.7 - MARGIN, h: 4.6 };

  // Rectangles par défaut (positions initiales relatives à l'image)
  const TGBT_RECT = {  // Rouge (Plan d’implantation)
    w: 1.5, h: 1.0,
    dx: 0.8, dy: 0.6
  };
  const BORNE_RECT = { // Bleu (Places à électrifier)
    w: 1.5, h: 1.0,
    dx: 2.0, dy: 1.8
  };

  // Cercle vert (Plan d’implantation) : position et taille par défaut
  const GREEN_CIRCLE = {
    dx: 3.5,  // offset X par rapport au coin haut-gauche de l'image
    dy: 2.0,  // offset Y
    w: 2.0,   // largeur
    h: 2.0,   // hauteur (égale à w pour un cercle parfait)
    fill: "00FF00",     // vert clair
    stroke: "008000",   // vert foncé
    strokeWidth: 4      // épaisseur visible
  };

  // ---------------- Éléments (2 slides par rubrique) ----------------
  const items = [
    { title: "Plan d'implantation #1", file: "file1",  comment: "comment1"  },
    { title: "Plan d'implantation #2", file: "file1b", comment: "comment1b" },
    { title: "Places à électrifier #1", file: "file2",  comment: "comment2"  },
    { title: "Places à électrifier #2", file: "file2b", comment: "comment2b" },
    { title: "TGBT + disjoncteur de tête #1", file: "file3",  comment: "comment3"  },
    { title: "TGBT + disjoncteur de tête #2", file: "file3b", comment: "comment3b" },
    { title: "Cheminement #1", file: "file4",  comment: "comment4"  },
    { title: "Cheminement #2", file: "file4b", comment: "comment4b" },
    { title: "Plan du site #1", file: "file5",  comment: "comment5"  },
    { title: "Plan du site #2", file: "file5b", comment: "comment5b" },
    { title: "Éléments complémentaires #1", file: "file6",  comment: "comment6"  },
    { title: "Éléments complémentaires #2", file: "file6b", comment: "comment6b" }
  ];

  let done = 0;
  const total = items.length;

  items.forEach((item) => {
    const fileInput = document.getElementById(item.file);
    const comment   = document.getElementById(item.comment)?.value || "—";
    const slide     = pptx.addSlide();

    // Titre
    slide.addText(item.title, { x: MARGIN, y: 0.5, fontSize: 24, bold: true });

    // Zone commentaire (déplaçable)
    slide.addText(comment, {
      x: BOX.x, y: BOX.y, w: BOX.w, h: BOX.h,
      fill: { color: "FFFFFF" },
      line: { color: "AAAAAA" },
      margin: 0.12,
      fontSize: 18, color: "111111",
      align: "left", valign: "top"
    });

    // Ajout de l'image puis des formes si besoin
    const placeImageAndShapes = (dataUrl) => {
      if (dataUrl) {
        slide.addImage({
          data: dataUrl,
          x: IMG.x, y: IMG.y, w: IMG.w, h: IMG.h,
          sizing: { type: "contain", w: IMG.w, h: IMG.h }
        });
      }

      const titleLower = item.title.toLowerCase();

      // Plan d’implantation → rectangle rouge (TGBT) + cercle vert épais
      if (titleLower.includes("implantation")) {
        // 🔴 Rectangle rouge
        slide.addShape(pptx.shapes.RECTANGLE, {
          x: IMG.x + TGBT_RECT.dx,
          y: IMG.y + TGBT_RECT.dy,
          w: TGBT_RECT.w,
          h: TGBT_RECT.h,
          fill: { color: "FF0000" },
          line: { color: "000000", width: 1.5 }
        });

        // 🟢 Cercle vert épais
        slide.addShape(pptx.shapes.ELLIPSE, {
          x: IMG.x + GREEN_CIRCLE.dx,
          y: IMG.y + GREEN_CIRCLE.dy,
          w: GREEN_CIRCLE.w,
          h: GREEN_CIRCLE.h,
          fill: { color: GREEN_CIRCLE.fill },
          line: { color: GREEN_CIRCLE.stroke, width: GREEN_CIRCLE.strokeWidth }
        });
      }

      // Places à électrifier → rectangle bleu (bornes)
      if (titleLower.includes("places")) {
        slide.addShape(pptx.shapes.RECTANGLE, {
          x: IMG.x + BORNE_RECT.dx,
          y: IMG.y + BORNE_RECT.dy,
          w: BORNE_RECT.w,
          h: BORNE_RECT.h,
          fill: { color: "0070C0" },           // bleu
          line: { color: "000000", width: 1.5 }
        });
      }

      checkDone();
    };

    if (fileInput?.files?.length > 0) {
      const reader = new FileReader();
      reader.onload = (e) => placeImageAndShapes(e.target.result);
      reader.readAsDataURL(fileInput.files[0]);
    } else {
      placeImageAndShapes(null);
    }
  });

  function checkDone() {
    done++;
    if (done === total) {
      const safeName = (clientName || "Projet")
        .replace(/[^\p{L}\p{N}_\- ]/gu, "")
        .trim()
        .replace(/\s+/g, "_");

      pptx.writeFile({ fileName: `Borne_Electrique_${safeName}.pptx` })
        .finally(() => {
          btn?.removeAttribute("aria-busy");
          btn?.removeAttribute("disabled");
        });
    }
  }
}
