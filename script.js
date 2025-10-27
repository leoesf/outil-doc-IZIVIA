// -----------------------------------------------------------
/* script.js - Génération du PowerPoint (PptxGenJS v3.x)
   - Images à gauche, commentaire à droite (bloc déplaçable)
   - Overlays :
       * Plan d'implantation : rectangle ROUGE + cercle VERT épais
       * Places à électrifier : rectangle BLEU
*/
// -----------------------------------------------------------

document.addEventListener("DOMContentLoaded", () => {
  const ok = typeof PptxGenJS !== "undefined";
  console.log("Pptx present ?", ok, "(PptxGenJS:", ok, ")");
  if (!ok) {
    alert("❌ PptxGenJS n'est pas chargé. Vérifie que `pptxgen.bundle.js` est bien présent et référencé dans borneelectrique.html.");
  }
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

  // === Types de formes : compat maximum (enum OU string) ===
  const RECT    = PptxGenJS?.ShapeType?.rect    || "rect";
  const ELLIPSE = PptxGenJS?.ShapeType?.ellipse || "ellipse";
  console.log("Shapes used => RECT:", RECT, "ELLIPSE:", ELLIPSE);

  // --- util ---
  const getVal = (id) => document.getElementById(id)?.value || "";

  // --- champs / infos ---
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

  // ---------------- Couverture ----------------
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
    if (nbBornes)      lines.push({ text: `Nombre de bornes : ${nbBornes}\n`, options: { fontSize: 16, color: "FFFFFF" } });
    if (bornesPower)   lines.push({ text: `Puissance des bornes : ${bornesPower}\n`, options: { fontSize: 16, color: "FFFFFF" } });
    lines.push({
      text: "Projet d’infrastructure de recharge pour véhicules électriques",
      options: { fontSize: 14, color: "FFFFFF", italic: true, breakLine: true }
    });

    slide.addText(lines, { x: 0.5, y: 0.5, w: 5.8, h: 6.2 });
  }

  // ---------------- Diapo Compléments d'informations ----------------
  function addInfoSlide() {
    const slide = pptx.addSlide();
    slide.addText("Compléments d’informations", { x: 0.5, y: 0.5, fontSize: 24, bold: true });
    slide.addText(raeClient || "—", {
      x: 0.5, y: 1.5, w: "90%", h: "70%",
      fontSize: 18, color: "363636",
      fill: { color: "FFFFFF" }, line: { color: "DDDDDD" },
      margin: 0.12, valign: "top"
    });
  }

  // ---------------- Mise en page standard (image gauche / texte droite) ----------------
  const SLIDE_W = 10.0;
  const MARGIN  = 0.5;
  const IMG = { x: MARGIN, y: 1.1, w: 6.1, h: 4.6 }; // image à gauche
  const BOX = { x: 6.7, y: 1.1, w: SLIDE_W - 6.7 - MARGIN, h: 4.6 }; // texte à droite

  // ---------------- Paramètres des overlays ----------------
  const TGBT_RECT    = { w: 1.6, h: 1.1, dx: 0.8, dy: 0.6 }; // rouge
  const BORNE_RECT   = { w: 1.6, h: 1.1, dx: 2.0, dy: 1.8 }; // bleu
  const GREEN_CIRCLE = { dx: 3.5, dy: 2.0, w: 2.0, h: 2.0, fill: "00FF00", stroke: "008000", strokeWidth: 4 }; // cercle vert épais

  function placeImageAndShapes(slide, title, imgBox, dataUrl) {
    // Image
    if (dataUrl) {
      slide.addImage({
        data: dataUrl,
        x: imgBox.x, y: imgBox.y, w: imgBox.w, h: imgBox.h,
        sizing: { type: "contain", w: imgBox.w, h: imgBox.h }
      });
    }

    const lower = title.toLowerCase();

    if (lower.includes("implantation")) {
      // Rectangle ROUGE (TGBT)
      slide.addShape(RECT, {
        x: IMG.x + TGBT_RECT.dx, y: IMG.y + TGBT_RECT.dy,
        w: TGBT_RECT.w, h: TGBT_RECT.h,
        fill: { color: "FF0000" },
        line: { color: "880000", width: 1 }
      });

      // Cercle VERT épais
      slide.addShape(ELLIPSE, {
        x: IMG.x + GREEN_CIRCLE.dx,
        y: IMG.y + GREEN_CIRCLE.dy,
        w: GREEN_CIRCLE.w,
        h: GREEN_CIRCLE.h,
        fill: { color: GREEN_CIRCLE.fill },
        line: { color: GREEN_CIRCLE.stroke, width: GREEN_CIRCLE.strokeWidth }
      });
    }

    if (
      lower.includes("places à électrifier") ||
      lower.includes("places a electrifier") ||
      lower.includes("place à electrifier") ||
      lower.includes("places elect")
    ) {
      // Rectangle BLEU (bornes)
      slide.addShape(RECT, {
        x: IMG.x + BORNE_RECT.dx, y: IMG.y + BORNE_RECT.dy,
        w: BORNE_RECT.w, h: BORNE_RECT.h,
        fill: { color: "0070C0" },
        line: { color: "004A87", width: 1 }
      });
    }
  }

  // ---------------- Slides de checklist (doublées) ----------------
  function addChecklistSlides() {
    const items = [
      { title: "Plan d'implantation #1",        file: "file1",  comment: "comment1"  },
      { title: "Plan d'implantation #2",        file: "file1b", comment: "comment1b" },

      { title: "Places à électrifier #1",       file: "file2",  comment: "comment2"  },
      { title: "Places à électrifier #2",       file: "file2b", comment: "comment2b" },

      { title: "TGBT + disjoncteur de tête #1", file: "file3",  comment: "comment3"  },
      { title: "TGBT + disjoncteur de tête #2", file: "file3b", comment: "comment3b" },

      { title: "Cheminement #1",                file: "file4",  comment: "comment4"  },
      { title: "Cheminement #2",                file: "file4b", comment: "comment4b" },

      { title: "Plan du site #1",               file: "file5",  comment: "comment5"  },
      { title: "Plan du site #2",               file: "file5b", comment: "comment5b" },

      { title: "Éléments complémentaires #1",   file: "file6",  comment: "comment6"  },
      { title: "Éléments complémentaires #2",   file: "file6b", comment: "comment6b" },
    ];

    let done = 0;
    const total = items.length;

    items.forEach((item) => {
      const fileInput = document.getElementById(item.file);
      const comment   = document.getElementById(item.comment)?.value || "—";
      const slide     = pptx.addSlide();

      // Titre
      slide.addText(item.title, { x: MARGIN, y: 0.5, fontSize: 24, bold: true });

      // Zone commentaire (à DROITE) — bloc déplaçable
      slide.addText(comment, {
        x: BOX.x, y: BOX.y, w: BOX.w, h: BOX.h,
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

      // Image (à GAUCHE) puis formes
      const inject = (dataUrl) => {
        placeImageAndShapes(slide, item.title, IMG, dataUrl);
        checkDone();
      };

      if (fileInput?.files?.length > 0) {
        const reader = new FileReader();
        reader.onload = (e) => inject(e.target.result);
        reader.readAsDataURL(fileInput.files[0]);
      } else {
        inject(null);
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
  addCoverSlide();
  addInfoSlide();
  addChecklistSlides();
}
