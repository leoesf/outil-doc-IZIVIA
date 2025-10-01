// -----------------------------------------------------------
/* script.js - Génération du PowerPoint (PptxGenJS v3.x)
   — Zones de texte colorées pour les marqueurs sur "Plan d'implantation" */
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
  const raeClient     = getVal("raeClient");
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

  // ---------------- Diapo RAE ----------------
  function addRAESlide() {
    const slide = pptx.addSlide();
    slide.addText("RAE du client", { x: 0.5, y: 0.5, fontSize: 24, bold: true });
    slide.addText(raeClient || "—", { x: 0.5, y: 1.5, fontSize: 18, w: "90%", h: "70%", color: "363636" });
  }

  // ---------------- Marqueurs en zones de texte ----------------
  // On crée 3 zones de texte colorées, vides ou avec un petit libellé, déplaçables dans PowerPoint.
  function addPlanMarkersAsTextBoxes(slide, imgBox) {
    // Position: dans le coin haut-droit de l'image
    const baseX = imgBox.x + imgBox.w - 2.0; // on réserve ~2" de large pour 3 boîtes empilées
    let y = imgBox.y + 0.2;

    // Style commun
    const common = {
      x: baseX,
      w: 1.8,
      h: 0.7,
      fontSize: 14,
      bold: true,
      align: "center",
      valign: "middle",
      line: { color: "444444", width: 1.0 },
      margin: 0.06
    };

    // Boîte JAUNE
    slide.addText("Zone JAUNE", {
      ...common,
      y,
      fill: { color: "FFD24D" }, // jaune
      color: "111111"
    });

    y += 0.85;

    // Boîte ROUGE
    slide.addText("Zone ROUGE", {
      ...common,
      y,
      fill: { color: "FF2B2B" }, // rouge
      color: "FFFFFF"
    });

    y += 0.85;

    // Boîte VERTE
    slide.addText("Zone VERTE", {
      ...common,
      y,
      fill: { color: "3A8F2D" }, // vert
      color: "FFFFFF"
    });

    // Petite étiquette pour te repérer (facultatif)
    slide.addText("Marqueurs (déplaçables)", {
      x: baseX - 0.1, y: imgBox.y - 0.25, w: 2.2, h: 0.35,
      fontSize: 10, color: "111111",
      fill: { color: "FFFFFF" }, line: { color: "DDDDDD" },
      align: "center", valign: "middle"
    });
  }

  // ---------------- Diapos Checklist ----------------
  function addChecklistSlides() {
    // Slide 16:9 : 10" x 5.625"
    const SLIDE_W = 10.0;
    const MARGIN  = 0.5;

    // Zone image gauche
    const IMG = { x: MARGIN, y: 1.1, w: 6.5, h: 4.8 }; // s'arrête vers 7.0"
    // Zone texte droite
    const BOX = { x: 7.2, y: 1.1, w: SLIDE_W - 7.2 - MARGIN, h: 4.8 }; // ~7.2 → 9.5"

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

      // Zone de commentaire à droite
      slide.addText(comment, {
        x: BOX.x, y: BOX.y, w: BOX.w, h: BOX.h,
        // une zone de texte est naturellement un rectangle déplaçable
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

      // Image à gauche
      const injectImage = (dataUrl) => {
        if (dataUrl) {
          slide.addImage({
            data: dataUrl,
            x: IMG.x, y: IMG.y, w: IMG.w, h: IMG.h,
            sizing: { type: "contain", w: IMG.w, h: IMG.h }
          });
        }

        // 👉 Ajout des "marqueurs" (zones de texte colorées) UNIQUEMENT pour Plan d'implantation
        if (item.title.toLowerCase().includes("implantation")) {
          addPlanMarkersAsTextBoxes(slide, IMG);
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

  // ---------------- Séquence ----------------
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
