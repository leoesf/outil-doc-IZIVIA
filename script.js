// -----------------------------------------------------------
// script.js - Génération du PowerPoint (PptxGenJS v3.x)
// - Couverture avec infos client
// - Slide "Compléments d'informations"
// - 3 diapositives par rubrique de checklist
// - Rectangles rouge/bleu + cercle vert + légende
// - Logo EDF en bas à droite de chaque slide (EDF.png)
// -----------------------------------------------------------

document.addEventListener("DOMContentLoaded", () => {
  console.log("✅ script.js chargé – PptxGenJS présent ?", typeof PptxGenJS !== "undefined");
  const btn = document.getElementById("exportBtn");
  if (btn) btn.addEventListener("click", createPowerPoint);
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
  pptx.layout = "LAYOUT_WIDE"; // 10 x 5.625

  const RECT = pptx.shapes.RECTANGLE;
  const ELLIPSE = pptx.shapes.OVAL;

  const getVal = (id) => document.getElementById(id)?.value || "";

  // Champs formulaire
  const clientName      = getVal("clientName");
  const rae             = getVal("rae");
  const powerSubscribed = getVal("powerSubscribed");
  const powerMax        = getVal("powerMax");
  const commercial      = getVal("commercial");
  const raeClient       = getVal("raeClient");
  const clientAddress   = getVal("clientAddress");
  const siret           = getVal("siret");
  const oppoNumber      = getVal("oppoNumber");
  const nbBornes        = getVal("nbBornes");
  const bornesPower     = getVal("bornesPower");

  const commercialPhone = getVal("commercialPhone");
  const commercialEmail = getVal("commercialEmail");

  const contactName     = getVal("contactName");
  const contactPhone    = getVal("contactPhone");
  const contactEmail    = getVal("contactEmail");

  // -----------------------------------------------------------
  // Constantes de mise en page
  // -----------------------------------------------------------
  const SLIDE_W = 10.0;
  const SLIDE_H = 5.625;
  const MARGIN  = 0.5;

  // Zone image (gauche) et zone commentaire (droite)
  const IMG = { x: MARGIN, y: 1.4, w: 5.8, h: 3.8 };
  const BOX = { x: 6.6, y: 1.4, w: 3.0, h: 3.8 };

  // Positions relatives des formes
  const TGBT_RECT = { w: 1.6, h: 1.1, dx: 1.0, dy: 0.8 };
  const BORNE_RECT = { w: 1.6, h: 1.1, dx: 3.2, dy: 2.0 };

  // Cercle vert placé dans le coin bas-droit de l'image
  const GREEN_CIRCLE = {
    w: 1.2,
    h: 1.2,
    x: IMG.x + IMG.w - 1.4,
    y: IMG.y + IMG.h - 1.4,
    stroke: "00FF00",
    strokeWidth: 3
  };

  // -----------------------------------------------------------
  // Logo EDF en bas à droite
  // -----------------------------------------------------------
  function addEDFLogo(slide) {
    slide.addImage({
      path: "EDF.png",   // le fichier doit être présent à la racine du projet
      x: 0.001,
      y: 6.95,
      w: 1.2,
      h: 0.55
    });
  }

  // -----------------------------------------------------------
  // SLIDE 1 – COUVERTURE
  // -----------------------------------------------------------
  function addCoverSlide() {
    const slide = pptx.addSlide();
    slide.background = { fill: "363636" };

    // Titre client
    slide.addText(`Client : ${clientName || ""}`, {
      x: 0.5,
      y: 0.4,
      w: 9,
      fontSize: 44,
      bold: true,
      color: "FFFFFF",
      align: "center"
    });

    // Sous-titre projet
    slide.addText("Projet d’infrastructure de recharge", {
      x: 0.5,
      y: 1.2,
      w: 9,
      fontSize: 22,
      color: "FFFFFF",
      align: "center"
    });

    // Bloc d'infos décalé légèrement vers la droite
    const lines = [];

    if (oppoNumber)    lines.push(`Oppo : ${oppoNumber}`);
    if (clientName)    lines.push(`Client : ${clientName}`);
    if (clientAddress) lines.push(`Adresse : ${clientAddress}`);
    if (siret)         lines.push(`Siret : ${siret}`);
    lines.push(""); // ligne vide

    if (rae)              lines.push(`Rae : ${rae}`);
    if (powerSubscribed)  lines.push(`Puissance souscrite : ${powerSubscribed} kVA`);
    if (powerMax)         lines.push(`Puissance max : ${powerMax} kVA`);
    if (nbBornes)         lines.push(`Bornes : ${nbBornes} bornes`);
    if (bornesPower)      lines.push(`Puissance des bornes : ${bornesPower} kW`);
    lines.push("");

    if (contactName)  lines.push(`Interlocuteur client : ${contactName}`);
    if (contactPhone) lines.push(`Mobile : ${contactPhone}`);
    if (contactEmail) lines.push(`Adresse électronique : ${contactEmail}`);
    lines.push("");

    if (commercial)      lines.push(`Interlocuteur EDF : ${commercial}`);
    if (commercialPhone) lines.push(`Tél. EDF : ${commercialPhone}`);
    if (commercialEmail) lines.push(`Mail EDF : ${commercialEmail}`);

    slide.addText(lines.join("\n"), {
      x: 1.0,       // léger déplacement à droite
      y: 2.0,
      w: 8,
      fontSize: 16,
      color: "FFFFFF",
      align: "left"
    });

    addEDFLogo(slide);
  }

  // -----------------------------------------------------------
  // SLIDE 2 – Compléments d’informations
  // -----------------------------------------------------------
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

    addEDFLogo(slide);
  }

  // -----------------------------------------------------------
  // Légende en bas à droite (fixe, quand il y a des formes)
  // -----------------------------------------------------------
  function addLegend(slide) {
  // Position en bas à droite
  const x = 11.0;
  const y = 6.5;

  const legendText = [
    // Ligne 1 : carré rouge = TGBT
    { text: "■ ", options: { fontSize: 14, color: "FF0000", bold: true } },
    { text: "= TGBT\n", options: { fontSize: 12, color: "000000" } },

    // Ligne 2 : carré bleu = Borne
    { text: "■ ", options: { fontSize: 14, color: "0070C0", bold: true } },
    { text: "= Borne\n", options: { fontSize: 12, color: "000000" } },

    // Ligne 3 : rond vert = Zone à équiper
    { text: "● ", options: { fontSize: 14, color: "00AA00", bold: true } },
    { text: "= Zone à équiper", options: { fontSize: 12, color: "000000" } }
  ];

  slide.addText(legendText, {
    x,
    y,
    w: 3.5,
    h: 1.2,
    valign: "top"
    // pas de bordure, pas de fond -> zone de texte seule, facile à déplacer
  });
}

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
    let hasShapes = false;

    // Plan d'implantation : rectangle rouge + cercle vert
    if (low.includes("implantation")) {
      slide.addShape(RECT, {
        x: IMG.x + TGBT_RECT.dx,
        y: IMG.y + TGBT_RECT.dy,
        w: TGBT_RECT.w,
        h: TGBT_RECT.h,
        fill: { color: "FF0000" },
        line: { color: "880000" }
      });

      slide.addShape(ELLIPSE, {
        x: GREEN_CIRCLE.x,
        y: GREEN_CIRCLE.y,
        w: GREEN_CIRCLE.w,
        h: GREEN_CIRCLE.h,
        fill: null,
        line: { color: GREEN_CIRCLE.stroke, width: GREEN_CIRCLE.strokeWidth }
      });

      hasShapes = true;
    }

    // Places à électrifier : rectangle bleu
    if (low.includes("places à électrifier")) {
      slide.addShape(RECT, {
        x: IMG.x + BORNE_RECT.dx,
        y: IMG.y + BORNE_RECT.dy,
        w: BORNE_RECT.w,
        h: BORNE_RECT.h,
        fill: { color: "0070C0" },
        line: { color: "003A70" }
      });

      hasShapes = true;
    }

    // Si au moins une forme a été ajoutée, on ajoute la légende générale
    if (hasShapes) {
      addLegend(slide);
    }

    addEDFLogo(slide);
  }

  // -----------------------------------------------------------
  // Slides checklist (3 par rubrique)
  // -----------------------------------------------------------
  function addChecklistSlides() {
    const sectionNumber = {
      "Plan d'implantation": 1,
      "Places à électrifier": 2,
      "TGBT + disjoncteur de tête": 3,
      "Cheminement": 4,
      "Plan du site": 5,
      "Éléments complémentaires": 6
    };

    const defs = [
      { base: "Plan d'implantation",           file: "file1",  comment: "comment1" },
      { base: "Plan d'implantation",           file: "file1b", comment: "comment1b" },
      { base: "Plan d'implantation",           file: "file1c", comment: "comment1c" },

      { base: "Places à électrifier",          file: "file2",  comment: "comment2" },
      { base: "Places à électrifier",          file: "file2b", comment: "comment2b" },
      { base: "Places à électrifier",          file: "file2c", comment: "comment2c" },

      { base: "TGBT + disjoncteur de tête",    file: "file3",  comment: "comment3" },
      { base: "TGBT + disjoncteur de tête",    file: "file3b", comment: "comment3b" },
      { base: "TGBT + disjoncteur de tête",    file: "file3c", comment: "comment3c" },

      { base: "Cheminement",                   file: "file4",  comment: "comment4" },
      { base: "Cheminement",                   file: "file4b", comment: "comment4b" },
      { base: "Cheminement",                   file: "file4c", comment: "comment4c" },

      { base: "Plan du site",                  file: "file5",  comment: "comment5" },
      { base: "Plan du site",                  file: "file5b", comment: "comment5b" },
      { base: "Plan du site",                  file: "file5c", comment: "comment5c" },

      { base: "Éléments complémentaires",      file: "file6",  comment: "comment6" },
      { base: "Éléments complémentaires",      file: "file6b", comment: "comment6b" },
      { base: "Éléments complémentaires",      file: "file6c", comment: "comment6c" }
    ];

    let remaining = defs.length;

    defs.forEach((item) => {
      const slide = pptx.addSlide();
      const sec = sectionNumber[item.base];

      // Titre centré, bleu, taille 36, avec numéro
      slide.addText(`${sec}. ${item.base}`, {
        x: 0.5,
        y: 0.3,
        w: SLIDE_W - 1,
        fontSize: 36,
        bold: true,
        color: "0070C0",
        align: "center"
      });

      // Zone commentaire à droite
      slide.addText(getVal(item.comment) || "—", {
        x: BOX.x,
        y: BOX.y,
        w: BOX.w,
        h: BOX.h,
        fill: { color: "FFFFFF" },
        line: { color: "AAAAAA" },
        fontSize: 18,
        valign: "top"
      });

      const fileInput = document.getElementById(item.file);

      const finalizeSlide = (dataUrl) => {
        placeImageAndShapes(slide, item.base, IMG, dataUrl);
        remaining--;
        if (remaining === 0) {
          pptx
            .writeFile({ fileName: "Borne_Electrique_Projet.pptx" })
            .finally(() => {
              btn?.removeAttribute("aria-busy");
              btn?.removeAttribute("disabled");
            });
        }
      };

      if (fileInput?.files?.length > 0) {
        const reader = new FileReader();
        reader.onload = (e) => finalizeSlide(e.target.result);
        reader.readAsDataURL(fileInput.files[0]);
      } else {
        finalizeSlide(null);
      }
    });
  }

  // -----------------------------------------------------------
  // Génération finale
  // -----------------------------------------------------------
  addCoverSlide();
  addInfoSlide();
  addChecklistSlides();
}




