// -----------------------------------------------------------
// script.js - Génération du PowerPoint (PptxGenJS)
// -----------------------------------------------------------

document.addEventListener("DOMContentLoaded", () => {
  console.log("Page prête. PptxGenJS chargé ?", typeof PptxGenJS !== "undefined");

  // Si le bouton a l'ID exportBtn, on branche un écouteur
  const btn = document.getElementById("exportBtn");
  if (btn) btn.addEventListener("click", createPowerPoint);
});

function createPowerPoint() {
  console.log("[PPT] Lancement export…");

  const btn = document.getElementById("exportBtn");
  btn?.setAttribute("disabled", "true");
  btn?.setAttribute("aria-busy", "true");

  if (typeof PptxGenJS === "undefined") {
    console.error("PptxGenJS non chargé.");
    alert("La librairie PptxGenJS n'est pas chargée.");
    btn?.removeAttribute("aria-busy");
    btn?.removeAttribute("disabled");
    return;
  }

  const pptx = new PptxGenJS();
  pptx.layout = "LAYOUT_WIDE"; // 16:9

  // ======== Récupération des champs ========
  const clientName    = document.getElementById("clientName")?.value || "";
  const rae           = document.getElementById("rae")?.value || "";
  const power         = document.getElementById("power")?.value || "";
  const commercial    = document.getElementById("commercial")?.value || "";
  const raeClient     = document.getElementById("raeClient")?.value || "";
  const coverImageInp = document.getElementById("coverImage");

  // Nouveaux champs
  const clientAddress = document.getElementById("clientAddress")?.value || "";
  const siret         = document.getElementById("siret")?.value || "";
  const oppoNumber    = document.getElementById("oppoNumber")?.value || "";
  const nbBornes      = document.getElementById("nbBornes")?.value || "";
  const bornesPower   = document.getElementById("bornesPower")?.value || "";

  // ======== Diapo 1 : Couverture ========
  function addCoverSlide(imageData) {
    const slide = pptx.addSlide();
    slide.background = { fill: "363636" };

    const lines = [];
    if (clientName)   lines.push({ text: `Client : ${clientName}\n`, options: { fontSize: 20, color: "FFFFFF", bold: true } });
    if (rae)          lines.push({ text: `RAE : ${rae}\n`,           options: { fontSize: 16, color: "FFFFFF" } });
    if (power)        lines.push({ text: `Puissance : ${power}\n`,   options: { fontSize: 16, color: "FFFFFF" } });
    if (commercial)   lines.push({ text: `Commercial : ${commercial}\n`, options: { fontSize: 16, color: "FFFFFF" } });

    // Nouveaux champs
    if (clientAddress) lines.push({ text: `Adresse : ${clientAddress}\n`, options: { fontSize: 16, color: "FFFFFF" } });
    if (siret)         lines.push({ text: `SIRET : ${siret}\n`,           options: { fontSize: 16, color: "FFFFFF" } });
    if (oppoNumber)    lines.push({ text: `Numéro Oppo : ${oppoNumber}\n`,options: { fontSize: 16, color: "FFFFFF" } });
    if (nbBornes)      lines.push({ text: `Nombre de bornes : ${nbBornes}\n`, options: { fontSize: 16, color: "FFFFFF" } });
    if (bornesPower)   lines.push({ text: `Puissance des bornes : ${bornesPower}\n`, options: { fontSize: 16, color: "FFFFFF" } });

    lines.push({
      text: "Projet d’infrastructure de recharge pour véhicules électriques",
      options: { fontSize: 14, color: "FFFFFF", italic: true, breakLine: true }
    });

    // Texte à gauche
    slide.addText(lines, { x: 0.5, y: 0.5, w: 5.8, h: 6.2 });

    // Image à droite (si fournie)
    if (imageData) {
      slide.addImage({
        data: imageData,
        x: 6.7, y: 0.4, w: 3.8, h: 5.8,
        sizing: { type: "cover", w: 3.8, h: 5.8 }
      });
    }
  }

  // ======== Diapo 2 : RAE Client ========
  function addRAESlide() {
    const slide = pptx.addSlide();
    slide.addText("RAE du client", { x: 0.5, y: 0.5, fontSize: 24, bold: true });
    slide.addText(raeClient || "—", { x: 0.5, y: 1.5, fontSize: 18, w: "90%", h: "70%", color: "363636" });
  }

  // ======== Diapos Checklist : Image à gauche / texte à droite ========
  function addChecklistSlides() {
    // Layout zones
    const IMG = { x: 0.5, y: 1.2, w: 6.8, h: 4.8 }; // image à gauche
    const BOX = { x: 7.6, y: 1.2, w: 3.4, h: 4.8 }; // zone de texte (commentaire) à droite

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
      slide.addText(item.title, { x: 0.5, y: 0.5, fontSize: 24, bold: true });

      // Zone de texte à droite (textbox avec fond + bordure)
      slide.addText(comment, {
        x: BOX.x, y: BOX.y, w: BOX.w, h: BOX.h,
        shape: pptx.shapes.ROUNDED_RECTANGLE, // ou pptx.shapes.RECTANGLE
        fill: { color: "F3F4F6" },            // fond clair
        line: { color: "D1D5DB" },            // bordure
        fontSize: 16,
        color: "363636",
        align: "left",
        valign: "top"                         // texte en haut de la zone
      });

      // Image à gauche (si sélectionnée)
      if (fileInput?.files?.length > 0) {
        const reader = new FileReader();
        reader.onload = function (e) {
          slide.addImage({
            data: e.target.result,
            x: IMG.x, y: IMG.y, w: IMG.w, h: IMG.h,
            sizing: { type: "contain", w: IMG.w, h: IMG.h }
          });
          checkDone();
        };
        reader.readAsDataURL(fileInput.files[0]);
      } else {
        // Pas d'image : on laisse la zone gauche vide
        checkDone();
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

  // ======== Lancement séquencé ========
  if (coverImageInp?.files?.length > 0) {
    const reader = new FileReader();
    reader.onload = function (e) {
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

// Fallback global pour compatibilité avec onclick="createPowerPoint()"
if (typeof window !== "undefined") {
  window.createPowerPoint = createPowerPoint;
}
