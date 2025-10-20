// -----------------------------------------------------------
// script.js - Génération du PowerPoint (PptxGenJS v3.x)
// - "RAE du client" → "Compléments d’informations"
// - 2 jeux (photo + commentaire) par rubrique → 2 diapositives
// -----------------------------------------------------------

document.addEventListener("DOMContentLoaded", () => {
  console.log("PptxGenJS chargé ?", typeof PptxGenJS !== "undefined");
  window.createPowerPoint = createPowerPoint; // si bouton HTML utilise onclick
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

  // --- util ---
  const getVal = (id) => document.getElementById(id)?.value || "";

  // --- champs couverture ---
  const clientName    = getVal("clientName");
  const rae           = getVal("rae");
  const power         = getVal("power");
  const commercial    = getVal("commercial");
  const infoComplem   = getVal("raeClient"); // zone de texte, renommée côté UI
  const coverImageInp = document.getElementById("coverImage");

  // nouveaux champs
  const clientAddress = getVal("clientAddress");
  const siret         = getVal("siret");
  const oppoNumber    = getVal("oppoNumber");
  const nbBornes      = getVal("nbBornes");
  const bornesPower   = getVal("bornesPower");

  // ---------------- Couverture ----------------
  function addCoverSlide(imageData) {
    const slide = pptx.addSlide();
    slide.background = { color: "363636" };

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

  // ---------------- Diapo Compléments d'informations ----------------
  function addComplementsSlide() {
    const slide = pptx.addSlide();
    slide.background = { color: "FFFFFF" };
    slide.addText("Compléments d’informations", { x: 0.5, y: 0.5, fontSize: 24, bold: true, color: "111827" });
    slide.addText(infoComplem || "—", { x: 0.5, y: 1.1, w: 9.2, h: 4.2, fontSize: 18, color: "363636", valign: "top" });
  }

  // ---------------- Marqueurs en zones de texte (optionnels) ----------------
  function addPlanMarkersAsTextBoxes(slide, imgBox) {
    const baseX = imgBox.x + imgBox.w - 2.0;
    let y = imgBox.y + 0.2;
    const common = {
      x: baseX, w: 1.8, h: 0.7, fontSize: 14, bold: true,
      align: "center", valign: "middle", line: { color: "444444", width: 1.0 }, margin: 0.06
    };
    slide.addText("Zone JAUNE", { ...common, y,   fill: { color: "FFD24D" }, color: "111111" });
    slide.addText("Zone ROUGE", { ...common, y: y+0.85, fill: { color: "FF2B2B" }, color: "FFFFFF" });
    slide.addText("Zone VERTE", { ...common, y: y+1.70, fill: { color: "3A8F2D" }, color: "FFFFFF" });
    slide.addText("Marqueurs (déplaçables)", {
      x: baseX - 0.1, y: imgBox.y - 0.25, w: 2.2, h: 0.35,
      fontSize: 10, color: "111111", fill: { color: "FFFFFF" }, line: { color: "DDDDDD" },
      align: "center", valign: "middle"
    });
  }

  // ---------------- Diapos Checklist (2 slides par rubrique) ----------------
  function addChecklistSlides() {
    const SLIDE_W = 10.0;
    const MARGIN  = 0.5;
    const IMG = { x: MARGIN, y: 1.1, w: 6.5, h: 4.8 };
    const BOX = { x: 7.2, y: 1.1, w: SLIDE_W - 7.2 - MARGIN, h: 4.8 };

    // Chaque rubrique possède 2 entrées (a/b)
    const items = [
      { base: "Plan d'implantation",      pairs: [ ["file1a","comment1a"], ["file1b","comment1b"] ] },
      { base: "Places à électrifier",     pairs: [ ["file2a","comment2a"], ["file2b","comment2b"] ] },
      { base: "TGBT + disjoncteur de tête", pairs: [ ["file3a","comment3a"], ["file3b","comment3b"] ] },
      { base: "Cheminement",              pairs: [ ["file4a","comment4a"], ["file4b","comment4b"] ] },
      { base: "Plan du site",             pairs: [ ["file5a","comment5a"], ["file5b","comment5b"] ] },
      { base: "Éléments complémentaires", pairs: [ ["file6a","comment6a"], ["file6b","comment6b"] ] }
    ];

    let done = 0;
    const totalSlides = items.reduce((acc, it) => acc + it.pairs.length, 0);

    items.forEach((rub) => {
      rub.pairs.forEach(async ([fileId, commentId], idx) => {
        const fileInput = document.getElementById(fileId);
        const comment   = document.getElementById(commentId)?.value || "—";
        const slide     = pptx.addSlide();

        // Titre avec suffixe #1 / #2
        const title = `${rub.base} — ${idx === 0 ? "1" : "2"}`;
        slide.addText(title, { x: MARGIN, y: 0.5, fontSize: 24, bold: true });

        // Zone commentaire à droite
        slide.addText(comment, {
          x: BOX.x, y: BOX.y, w: BOX.w, h: BOX.h,
          fill: { color: "FFFFFF" }, line: { color: "AAAAAA" }, margin: 0.12,
          fontSize: 18, color: "111111", align: "left", valign: "top", bullet: false, paraSpaceAfter: 6
        });

        const injectImage = (dataUrl) => {
          if (dataUrl) {
            slide.addImage({ data: dataUrl, x: IMG.x, y: IMG.y, w: IMG.w, h: IMG.h, sizing: { type: "contain", w: IMG.w, h: IMG.h } });
          }

          // Marqueurs uniquement sur "Plan d'implantation"
          if (rub.base.toLowerCase().includes("implantation")) {
            addPlanMarkersAsTextBoxes(slide, IMG);
          }

          checkDone();
        };

        if (fileInput?.files?.length > 0) {
          // possibilité d’ajouter ici une compression si souhaité
          const reader = new FileReader();
          reader.onload = (e) => injectImage(e.target.result);
          reader.readAsDataURL(fileInput.files[0]);
        } else {
          injectImage(null);
        }
      });
    });

    function checkDone() {
      done++;
      if (done === totalSlides) {
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

  // ---------------- Séquence ----------------
  if (coverImageInp?.files?.length > 0) {
    const reader = new FileReader();
    reader.onload = (e) => {
      addCoverSlide(e.target.result);
      addComplementsSlide();     // 🔁 nouveau nom + slide
      addChecklistSlides();      // 🔁 deux slides par rubrique
    };
    reader.readAsDataURL(coverImageInp.files[0]);
  } else {
    addCoverSlide();
    addComplementsSlide();       // 🔁 nouveau nom + slide
    addChecklistSlides();        // 🔁 deux slides par rubrique
  }
}
