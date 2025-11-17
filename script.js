// -----------------------------------------------------------
// script.js - Génération du PowerPoint (PptxGenJS v3.x)
// - Couverture avec infos client + logo EDF + logo IZIVIA
// - Slide "Compléments d'informations"
// - 3 diapositives par rubrique de checklist
// - Rectangles rouge/bleu + cercle vert + trait rouge + légende
// - Logo EDF en bas à gauche de chaque slide (EDF.png)
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

  const RECT    = pptx.shapes.RECTANGLE;
  const ELLIPSE = pptx.shapes.OVAL;
  const LINE    = pptx.shapes.LINE;

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

  // Zone image (gauche)
  const IMG = { x: MARGIN, y: 1.4, w: 5.8, h: 3.8 };

  // ✅ Zone commentaire décalée vers la droite (~3 cm)
  const BOX = { x: 7.8, y: 1.4, w: 2.2, h: 3.8 };

  // Positions relatives des formes (sur l'image)
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
  // Logo EDF en bas à gauche
  // -----------------------------------------------------------
  function addEDFLogo(slide) {
    slide.addImage({
      path: "EDF.png",  // le fichier doit être présent à la racine du projet
      x: 0.001,
      y: 6.9, // proche du bas
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

    // Bloc d'infos (légèrement à gauche pour laisser la place à l'image IZIVIA à droite)
    const lines = [];

    if (oppoNumber)    lines.push(`Oppo : ${oppoNumber}`);
    if (clientName)    lines.push(`Client : ${clientName}`);
    if (clientAddress) lines.push(`Adresse : ${clientAddress}`);
    if (siret)         lines.push(`Siret : ${siret}`);
    lines.push("");

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
      x: 0.6,      // un peu plus à gauche
      y: 2.0,
      w: 4.8,      // largeur réduite pour laisser une zone libre à droite
      fontSize: 16,
      color: "FFFFFF",
      align: "left"
    });

    // 🖼️ Image IZIVIA à droite (IZIVIA.jpg à la racine du projet)
    slide.addImage({
      path: "IZIVIA.jpg",
      x: 6.9,
      y: 3.2,
      w: 6.1,
      h: 5.8,
      sizing: { type: "contain" }
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
  // Légende : UNE SEULE ZONE DE TEXTE (fond transparent)
  // -----------------------------------------------------------
  function addLegend(slide, { showRedRect, showBlueRect, showGreenCircle, showRedLine }) {
    if (!showRedRect && !showBlueRect && !showGreenCircle && !showRedLine) return;

    const richText = [];

    // Titre "Légende :"
    richText.push({
      text: "Légende :",
      options: { bold: true, fontSize: 12, color: "000000", breakLine: true }
    });

    // Chaque ligne : symbole coloré + texte, mais tout dans UNE SHAPE
    if (showRedRect) {
      richText.push(
        { text: "■ ", options: { fontSize: 12, color: "FF0000" } },
        { text: "TGBT", options: { fontSize: 12, color: "000000", breakLine: true } }
      );
    }

    if (showBlueRect) {
      richText.push(
        { text: "■ ", options: { fontSize: 12, color: "0070C0" } },
        { text: "Borne / place équipée", options: { fontSize: 12, color: "000000", breakLine: true } }
      );
    }

    if (showGreenCircle) {
      richText.push(
        { text: "● ", options: { fontSize: 12, color: "00AA00" } },
        { text: "Zone à équiper", options: { fontSize: 12, color: "000000", breakLine: true } }
      );
    }

    if (showRedLine) {
      richText.push(
        { text: "― ", options: { fontSize: 12, color: "FF0000" } },
        { text: "Chemin de câble", options: { fontSize: 12, color: "000000", breakLine: true } }
      );
    }

    // Une seule zone de texte, FOND TRANSPARENT
    slide.addText(richText, {
      x: 11.5,
      y: 6.8,
      w: 3.0,
      h: 1.5,
      fontSize: 12,
      color: "000000",
      valign: "top"
    });
  }

  // -----------------------------------------------------------
  // Ajout image + formes (selon le type de slide)
  // -----------------------------------------------------------
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

    let showRedRect      = false;
    let showBlueRect     = false;
    let showGreenCircle  = false;
    let showRedLine      = false;

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
      showRedRect = true;

      slide.addShape(ELLIPSE, {
        x: GREEN_CIRCLE.x,
        y: GREEN_CIRCLE.y,
        w: GREEN_CIRCLE.w,
        h: GREEN_CIRCLE.h,
        fill: null,
        line: { color: GREEN_CIRCLE.stroke, width: GREEN_CIRCLE.strokeWidth }
      });
      showGreenCircle = true;
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
      showBlueRect = true;
    }

    // Cheminement : trait rouge (chemin de câble)
    if (low.includes("cheminement")) {
      slide.addShape(LINE, {
        x: IMG.x + 0.3,
        y: IMG.y + IMG.h - 0.6,
        w: IMG.w - 0.6,
        h: 0,
        line: { color: "FF0000", width: 3 }
      });
      showRedLine = true;
    }

    addLegend(slide, { showRedRect, showBlueRect, showGreenCircle, showRedLine });
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




