// -----------------------------------------------------------
// Génération du PowerPoint (PptxGenJS v3.x)
// - Commentaires à DROITE de la photo (textbox large, déplaçable)
// - Rubriques doublées (2 slides par rubrique)
// - "RAE du client" renommé en "Compléments d’informations"
// - PAS d'image de couverture (retirée précédemment)
// -----------------------------------------------------------

document.addEventListener("DOMContentLoaded", () => {
  console.log("PptxGenJS chargé ?", typeof PptxGenJS !== "undefined");
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
  pptx.layout = "LAYOUT_WIDE"; // 16:9

  const getVal = (id) => document.getElementById(id)?.value || "";

  // Champs infos
  const clientName    = getVal("clientName");
  const rae           = getVal("rae");
  const power         = getVal("power");
  const commercial    = getVal("commercial");
  const infoComplem   = getVal("raeClient");
  const clientAddress = getVal("clientAddress");
  const siret         = getVal("siret");
  const oppoNumber    = getVal("oppoNumber");
  const nbBornes      = getVal("nbBornes");
  const bornesPower   = getVal("bornesPower");

  // ---------- Diapo 1 : Informations client ----------
  function addInfoSlide() {
    const slide = pptx.addSlide();
    slide.background = { color: "363636" };

    const lines = [];
    if (clientName)   lines.push({ text: `Client : ${clientName}\n`,      options: { fontSize: 20, color: "FFFFFF", bold: true } });
    if (rae)          lines.push({ text: `RAE : ${rae}\n`,                options: { fontSize: 16, color: "FFFFFF" } });
    if (power)        lines.push({ text: `Puissance : ${power}\n`,        options: { fontSize: 16, color: "FFFFFF" } });
    if (commercial)   lines.push({ text: `Commercial : ${commercial}\n`,  options: { fontSize: 16, color: "FFFFFF" } });
    if (clientAddress) lines.push({ text: `Adresse : ${clientAddress}\n`, options: { fontSize: 16, color: "FFFFFF" } });
    if (siret)         lines.push({ text: `SIRET : ${siret}\n`,           options: { fontSize: 16, color: "FFFFFF" } });
    if (oppoNumber)    lines.push({ text: `Numéro Oppo : ${oppoNumber}\n`,options: { fontSize: 16, color: "FFFFFF" } });
    if (nbBornes)      lines.push({ text: `Nombre de bornes : ${nbBornes}\n`,options: { fontSize: 16, color: "FFFFFF" } });
    if (bornesPower)   lines.push({ text: `Puissance des bornes : ${bornesPower}\n`,options: { fontSize: 16, color: "FFFFFF" } });

    slide.addText(lines, { x: 0.6, y: 0.6, w: 8.8, h: 5.0 });
  }

  // ---------- Diapo 2 : Compléments d’informations ----------
  function addComplementsSlide() {
    const slide = pptx.addSlide();
    slide.addText("Compléments d’informations", { x: 0.6, y: 0.6, fontSize: 24, bold: true });
    slide.addText(infoComplem || "—", {
      x: 0.6, y: 1.2, w: 8.8, h: 4.2,
      fontSize: 18, color: "363636", valign: "top",
      fill: { color: "FFFFFF" }, line: { color: "AAAAAA" }, margin: 0.14
    });
  }

  // ---------- Checklist : photo à gauche, commentaire à droite ----------
  function addChecklistSlides() {
    // Slide 16:9 : 10" x 5.625"
    const SLIDE_W = 10.0;
    const MARGIN  = 0.5;

    // Image à gauche
    const IMG = { x: MARGIN, y: 1.1, w: 6.3, h: 4.6 };

    // Zone de texte à droite (grande, déplaçable)
    const BOX = { x: 7.1, y: 1.1, w: SLIDE_W - 7.1 - MARGIN, h: 4.6 };

    const items = [
      { base: "Plan d'implantation",        pairs: [ ["file1a","comment1a"], ["file1b","comment1b"] ] },
      { base: "Places à électrifier",       pairs: [ ["file2a","comment2a"], ["file2b","comment2b"] ] },
      { base: "TGBT + disjoncteur de tête", pairs: [ ["file3a","comment3a"], ["file3b","comment3b"] ] },
      { base: "Cheminement",                pairs: [ ["file4a","comment4a"], ["file4b","comment4b"] ] },
      { base: "Plan du site",               pairs: [ ["file5a","comment5a"], ["file5b","comment5b"] ] },
      { base: "Éléments complémentaires",   pairs: [ ["file6a","comment6a"], ["file6b","comment6b"] ] }
    ];

    let done = 0;
    const totalSlides = items.reduce((acc, it) => acc + it.pairs.length, 0);

    items.forEach((rub) => {
      rub.pairs.forEach(([fileId, commentId], idx) => {
        const fileInput = document.getElementById(fileId);
        const comment   = document.getElementById(commentId)?.value || "—";
        const slide     = pptx.addSlide();

        // Titre
        const title = `${rub.base} — ${idx === 0 ? "1" : "2"}`;
        slide.addText(title, { x: MARGIN, y: 0.5, fontSize: 24, bold: true });

        // Zone de commentaire à droite (textbox large)
        slide.addText(comment, {
          x: BOX.x, y: BOX.y, w: BOX.w, h: BOX.h,
          fontSize: 18, color: "111111",
          fill: { color: "FFFFFF" },
          line: { color: "AAAAAA", width: 1 },
          margin: 0.14,
          align: "left",
          valign: "top",
          bullet: false,
          autoFit: false
        });

        // Image à gauche (contain)
        const injectImage = (dataUrl) => {
          if (dataUrl) {
            slide.addImage({
              data: dataUrl,
              x: IMG.x, y: IMG.y, w: IMG.w, h: IMG.h,
              sizing: { type: "contain", w: IMG.w, h: IMG.h }
            });
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

  // ---------- Exécution ----------
  addInfoSlide();
  addComplementsSlide();
  addChecklistSlides();
}
