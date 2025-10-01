// -----------------------------------------------------------
// script.js - Génération du PowerPoint (PptxGenJS v3.x)
// -----------------------------------------------------------

document.addEventListener("DOMContentLoaded", () => {
  console.log("PptxGenJS chargé ?", typeof PptxGenJS !== "undefined");
  window.createPowerPoint = createPowerPoint; // pour onclick HTML
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
  pptx.layout = "LAYOUT_WIDE"; // 16:9 (10 x 5.625")

  // ========= Utilitaire =========
  const getVal = (id) => document.getElementById(id)?.value || "";

  // ========= Champs =========
  const clientName    = getVal("clientName");
  const rae           = getVal("rae");
  const power         = getVal("power");
  const commercial    = getVal("commercial");
  const raeClient     = getVal("raeClient");
  const coverImageInp = document.getElementById("coverImage");

  // Ajouts
  const clientAddress = getVal("clientAddress");
  const siret         = getVal("siret");
  const oppoNumber    = getVal("oppoNumber");
  const nbBornes      = getVal("nbBornes");
  const bornesPower   = getVal("bornesPower");

  // ========= Couverture =========
  function addCoverSlide(imageData) {
    const slide = pptx.addSlide();
    slide.background = { fill: "363636" };

    const lines = [];
    if (clientName)   lines.push({ text: `Client : ${clientName}\n`,      options: { fontSize: 20, color: "FFFFFF", bold: true } });
    if (rae)          lines.push({ text: `RAE : ${rae}\n`,                options: { fontSize: 16, color: "FFFFFF" } });
    if (power)        lines.push({ text: `Puissance : ${power}\n`,        options: { fontSize: 16, color: "FFFFFF" } });
    if (commercial)   lines.push({ text: `Commercial : ${commercial}\n`,  options: { fontSize: 16, color: "FFFFFF" } });

    if (clientAddress) lines.push({ text: `Adresse : ${clientAddress}\n`,              options: { fontSize: 16, color: "FFFFFF" } });
    if (siret)         lines.push({ text: `SIRET : ${siret}\n`,                        options: { fontSize: 16, color: "FFFFFF" } });
    if (oppoNumber)    lines.push({ text: `Numéro Oppo : ${oppoNumber}\n`,            options: { fontSize: 16, color: "FFFFFF" } });
    if (nbBornes)      lines.push({ text: `Nombre de bornes : ${nbBornes}\n`,         options: { fontSize: 16, color: "FFFFFF" } });
    if (bornesPower)   lines.push({ text: `Puissance des bornes : ${bornesPower}\n`,  options: { fontSize: 16, color: "FFFFFF" } });

    lines.push({
      text: "Projet d’infrastructure de recharge pour véhicules électriques",
      options: { fontSize: 14, color: "FFFFFF", italic: true, breakLine: true }
    });

    slide.addText(lines, { x: 0.5, y: 0.5, w: 5.8, h: 6.2 });

    if (imageData) {
      slide.addImage({
        data: imageData,
        x: 6.7, y: 0.4, w: 3.8, h: 5.8,
        sizing: { type: "cover", w: 3.8, h: 5.8 }
      });
    }
  }

  // ========= RAE =========
  function addRAESlide() {
    const slide = pptx.addSlide();
    slide.addText("RAE du client", { x: 0.5, y: 0.5, fontSize: 24, bold: true });
    slide.addText(raeClient || "—", { x: 0.5, y: 1.5, fontSize: 18, w: "90%", h: "70%", color: "363636" });
  }

  // ========= Marqueurs déplaçables =========
  function addMoveableMarkers(slide, imgBox) {
    const baseX = imgBox.x + imgBox.w - 1.2;
    let y = imgBox.y + 0.2;

    // Ellipse contour vert épais (remplissage transparent)
    slide.addShape({
      shape: PptxGenJS.ShapeType.ellipse,
      x: baseX, y, w: 0.7, h: 1.5,
      line: { color: "3A8F2D", width: 6 },
      fill: { color: "FFFFFF", transparency: 100 }
    });

    y += 1.8;

    // Carré jaune
    slide.addShape({
      shape: PptxGenJS.ShapeType.rect,
      x: baseX, y, w: 0.8, h: 0.8,
      fill: { color: "FFD24D" },
      line: { color: "C2A23B", width: 2 }
    });

    y += 1.0;

    // Carré rouge
    slide.addShape({
      shape: PptxGenJS.ShapeType.rect,
      x: baseX, y, w: 0.8, h: 0.8,
      fill: { color: "FF2B2B" },
      line: { color: "B00000", width: 2 }
    });

    // Petite étiquette pour debug visuel
    slide.addText("Marqueurs", {
      x: baseX - 0.2, y: imgBox.y - 0.2, w: 2, h: 0.4,
      fontSize: 10, color: "111111",
      fill: { color: "FFFFFF" }, line: { color: "DDDDDD" }
    });
  }

  // ========= Diapos Checklist =========
  function addChecklistSlides() {
    const SLIDE_W = 10.0;
    const MARGIN  = 0.5;

    const IMG = { x: MARGIN, y: 1.1, w: 6.5, h: 4.8 };             // image gauche
    const BOX = { x: 7.2, y: 1.1, w: SLIDE_W - 7.2 - MARGIN, h: 4.8 }; // texte droite

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

      slide.addText(item.title, { x: MARGIN, y: 0.5, fontSize: 24, bold: true });

      // Zone texte droite avec shape rect
      slide.addText(comment, {
        x: BOX.x, y: BOX.y, w: BOX.w, h: BOX.h,
        shape: PptxGenJS.ShapeType.rect,
        fill: { color: "FFFFFF" },
        line: { color: "AAAAAA" },
        margin: 0.12,
        fontSize: 18,
        color: "111111",
        align: "left",
        valign: "top",
        bullet: false,
        paraSpaceAfter: 6
      });

      const injectImage = (dataUrl) => {
        if (dataUrl) {
          slide.addImage({
            data: dataUrl,
            x: IMG.x, y: IMG.y, w: IMG.w, h: IMG.h,
            sizing: { type: "contain", w: IMG.w, h: IMG.h }
          });
        }
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

  // ========= Séquence =========
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

