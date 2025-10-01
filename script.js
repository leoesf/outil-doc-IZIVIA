// -----------------------------------------------------------
// script.js - Génération du PowerPoint (PptxGenJS v3.x)
// -----------------------------------------------------------

document.addEventListener("DOMContentLoaded", () => {
  console.log("PptxGenJS chargé ?", typeof PptxGenJS !== "undefined");
  // pour onclick HTML et pour listener JS
  window.createPowerPoint = createPowerPoint;
  document.getElementById("exportBtn")?.addEventListener("click", createPowerPoint);
});

function createPowerPoint() {
  const btn = document.getElementById("exportBtn");
  btn?.setAttribute("disabled", "true");
  btn?.setAttribute("aria-busy", "true");

  if (typeof PptxGenJS === "undefined") {
    alert("PptxGenJS n'est pas chargé.");
    btn?.removeAttribute("aria-busy");
    btn?.removeAttribute("disabled");
    return;
  }

  const pptx = new PptxGenJS();
  pptx.layout = "LAYOUT_WIDE"; // 16:9 (10" x 5.625")

  // ========= Utilitaire =========
  const getVal = (id) => document.getElementById(id)?.value || "";

  // ========= Champs =========
  const clientName    = getVal("clientName");
  const rae           = getVal("rae");
  const power         = getVal("power");
  const commercial    = getVal("commercial");
  const raeClient     = getVal("raeClient");
  const coverImageInp = document.getElementById("coverImage");

  // Nouveaux champs
  const clientAddress = getVal("clientAddress");
  const siret         = getVal("siret");
  const oppoNumber    = getVal("oppoNumber");
  const nbBornes      = getVal("nbBornes");
  const bornesPower   = getVal("bornesPower");

  // ========= Diapo 1 : Couverture =========
  function addCoverSlide(imageData) {
    const slide = pptx.addSlide();
    slide.background = { fill: "363636" };

    const lines = [];
    if (clientName)   lines.push({ text: `Client : ${clientName}\n`,      options: { fontSize: 20, color: "FFFFFF", bold: true } });
    if (rae)          lines.push({ text: `RAE : ${rae}\n`,                options: { fontSize: 16, color: "FFFFFF" } });
    if (power)        lines.push({ text: `Puissance : ${power}\n`,        options: { fontSize: 16, color: "FFFFFF" } });
    if (commercial)   lines.push({ text: `Commercial : ${commercial}\n`,  options: { fontSize: 16, color: "FFFFFF" } });

    // Ajouts
    if (clientAddress) lines.push({ text: `Adresse : ${clientAddress}\n`,              options: { fontSize: 16, color: "FFFFFF" } });
    if (siret)         lines.push({ text: `SIRET : ${siret}\n`,                        options: { fontSize: 16, color: "FFFFFF" } });
    if (oppoNumber)    lines.push({ text: `Numéro Oppo : ${oppoNumber}\n`,            options: { fontSize: 16, color: "FFFFFF" } });
    if (nbBornes)      lines.push({ text: `Nombre de bornes : ${nbBornes}\n`,         options: { fontSize: 16, color: "FFFFFF" } });
    if (bornesPower)   lines.push({ text: `Puissance des bornes : ${bornesPower}\n`,  options: { fontSize: 16, color: "FFFFFF" } });

    lines.push({
      text: "Projet d’infrastructure de recharge pour véhicules électriques",
      options: { fontSize: 14, color: "FFFFFF", italic: true, breakLine: true }
    });

    // Texte à gauche
    slide.addText(lines, { x: 0.5, y: 0.5, w: 5.8, h: 6.2 });

    // Image de couverture à droite (si fournie)
    if (imageData) {
      slide.addImage({
        data: imageData,
        x: 6.7, y: 0.4, w: 3.8, h: 5.8,
        sizing: { type: "cover", w: 3.8, h: 5.8 }
      });
    }
  }

  // ========= Diapo 2 : RAE Client =========
  function addRAESlide() {
    const slide = pptx.addSlide();
    slide.addText("RAE du client", { x: 0.5, y: 0.5, fontSize: 24, bold: true });
    slide.addText(raeClient || "—", { x: 0.5, y: 1.5, fontSize: 18, w: "90%", h: "70%", color: "363636" });
  }

  // ========= Marqueurs (via addText shape='ellipse'/'rect') =========
  function addMoveableMarkers(slide, imgBox) {
    // On place les marqueurs DANS la zone image (coin haut-droit) pour être visibles et manipulables
    const baseX = imgBox.x + imgBox.w - 1.2;
    let y = imgBox.y + 0.2;

    // Cercle vert (ellipse) contour épais, intérieur transparent
    slide.addText('', {
      shape: 'ellipse',              // <-- compatible partout
      x: baseX, y, w: 0.9, h: 1.6,
      fill: { color: 'FFFFFF', transparency: 100 },
      line: { color: '3A8F2D', width: 6 }
    });

    y += 1.9;

    // Carré jaune
    slide.addText('', {
      shape: 'rect',
      x: baseX, y, w: 0.9, h: 0.9,
      fill: { color: 'FFD24D' },
      line: { color: 'C2A23B', width: 2 }
    });

    y += 1.1;

    // Carré rouge
    slide.addText('', {
      shape: 'rect',
      x: baseX, y, w: 0.9, h: 0.9,
      fill: { color: 'FF2B2B' },
      line: { color: 'B00000', width: 2 }
    });

    // (facultatif) petite étiquette pour confirmer visuellement
    slide.addText('Marqueurs', {
      shape: 'rect',
      x: baseX - 0.2, y: imgBox.y - 0.2, w: 2, h: 0.4,
      fill: { color: 'FFFFFF' }, line: { color: 'DDDDDD' },
      fontSize: 10, color: '111111', align: 'left', valign: 'middle'
    });
  }

  // ========= Diapos Checklist (image gauche / texte droite) =========
  function addChecklistSlides() {
    // Slide 16:9 : 10" x 5.625"
    const SLIDE_W = 10.0;
    const MARGIN  = 0.5;

    // Zone image à gauche
    const IMG = { x: MARGIN, y: 1.1, w: 6.5, h: 4.8 }; // fin à 7.0"
    // Zone texte à droite (dans la slide)
    const BOX = { x: 7.2, y: 1.1, w: SLIDE_W - 7.2 - MARGIN, h: 4.8 }; // 7.2 → 9.5"

    const items = [
      { file: "file1", comment: "comment1", title: "Plan d'implantation" },
      { file: "file2", comment: "comment2", title: "Places à électrifier" },
      { file: "file3", comment: "comment3", title: "TGBT + disjoncteur de tête" },
      { file: "file4", comment: "comment4", title: "Cheminement" },
      { file: "file5", comment: "comment5", title: "Plan du site" },
      { file: "file6", comment: "comment6", title: "Éléments complémentaires" }
    ];

    let done = 0;
    const total = items.length;

    items.forEach((item) => {
      const fileInput = document.getElementById(item.file);
      const comment   = document.getElementById(item.comment)?.value || "—";
      const slide     = pptx.addSlide();

      // Titre
      slide.addText(item.title, { x: MARGIN, y: 0.5, fontSize: 24, bold: true });

      // Zone de texte à droite (forme rect pour la lisibilité et déplaçable)
      slide.addText(comment, {
        x: BOX.x, y: BOX.y, w: BOX.w, h: BOX.h,
        shape: 'rect',
        fill: { color: 'FFFFFF' },   // fond blanc
        line: { color: 'AAAAAA' },   // bordure grise
        margin: 0.12,
        fontSize: 18,
        color: '111111',
        align: 'left',
        valign: 'top',
        bullet: false,
        paraSpaceAfter: 6
      });

      // Image à gauche
      const injectImage = (dataUrl) => {
        if (dataUrl) {
          slide.addImage({
            data: dataUrl,
            x: IMG.x, y: IMG.y, w: IMG.w, h: IMG.h,
            sizing: { type: 'contain', w: IMG.w, h: IMG.h }
          });
        }

        // Marqueurs uniquement pour la diapo "Plan d'implantation"
        if (item.title.toLowerCase().includes("implantation")) {
          addMoveableMarkers(slide, IMG);
        }
        checkDone();
      };

      if (fileInput?.files?.length > 0) {
        const reader = new FileReader();
        reader.onload = (e) => injectImage(e.target.result);
        reader.readAsDataURL(fileInput.files[0]);
      } else {
        injectImage(null);
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

  // ========= Séquencement =========
  if (coverImageInp?.files?.length > 0) {
    const reader = new FileReader();
    reader.onload = (e) => {
      addCoverSlide(e.target.result);
      addRAESlide();
      addChecklistSlides();
    };
    reader.readAsDataURL(coverImageInp.files[0]);
  } else {
    addCoverSlide();
    addRAESlide();
    addChecklistSlides();
  }
}
