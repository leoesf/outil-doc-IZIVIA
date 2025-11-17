// -----------------------------------------------------------
// script.js - Génération du PowerPoint (PptxGenJS v3.x)
// Logo EDF + formes + mise en page couverture
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
  pptx.layout = "LAYOUT_WIDE"; // 16:9

  const RECT = pptx.shapes.RECTANGLE;
  const ELLIPSE = pptx.shapes.OVAL;

  const getVal = (id) => document.getElementById(id)?.value || "";

  // --------- Champs formulaire ---------
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
  // Logo EDF sur chaque diapo
  // -----------------------------------------------------------
  function addEDFLogo(slide) {
    slide.addImage({
      path: "EDF.png",   // fichier à avoir dans le repo
      x: 0.1,
      y: 6.8,
      w: 1.2,
      h: 0.55
    });
  }

  // -----------------------------------------------------------
  // SLIDE 1 – COUVERTURE (mise en forme façon exemple photo)
  // -----------------------------------------------------------
  function addCoverSlide() {
    const slide = pptx.addSlide();
    slide.background = { fill: "363636" };

    // Titre principal
    slide.addText(`Client : ${clientName || ""}`, {
      x: 0.5,
      y: 0.4,
      w: 9,
      fontSize: 44,
      bold: true,
      color: "FFFFFF",
      align: "center"
    });

    // Sous-titre
    slide.addText("Projet d’infrastructure de recharge", {
      x: 0.5,
      y: 1.0,
      w: 9,
      fontSize: 24,
      color: "FFFFFF",
      align: "center"
    });

    const leftX = 1.8;
    let y = 1.6;
    const lineH = 0.4;

    const addLabelValue = (label, value) => {
      if (!value) return;
      slide.addText(
        [
          { text: `${label} : `, options: { bold: true } },
          { text: value }
        ],
        {
          x: leftX,
          y,
          fontSize: 18,
          color: "FFFFFF"
        }
      );
      y += lineH;
    };

    // Bloc infos générales (comme sur la photo)
    addLabelValue("Oppo", oppoNumber);
    addLabelValue("Client", clientName);
    addLabelValue("Adresse", clientAddress);
    addLabelValue("Siret", siret);
    addLabelValue("Rae", rae);
    addLabelValue("Puissance souscrite", powerSubscribed ? `${powerSubscribed} kVA` : "");
    addLabelValue("Puissance max", powerMax ? `${powerMax} kVA` : "");
    addLabelValue(
      "Bornes",
      nbBornes || bornesPower
        ? `${nbBornes || "—"} borne(s) de ${bornesPower || "—"} kW`
        : ""
    );

    // Petit espace avant les blocs interlocuteurs
    y += 0.3;

    // Interlocuteur client
    if (contactName || contactPhone || contactEmail) {
      slide.addText(`Interlocuteur client : ${contactName || ""}`, {
        x: leftX,
        y,
        fontSize: 18,
        bold: true,
        color: "FFFFFF"
      });
      y += lineH;
      if (contactPhone) {
        slide.addText(`Mobile : ${contactPhone}`, {
          x: leftX,
          y,
          fontSize: 18,
          color: "FFFFFF"
        });
        y += lineH;
      }
      if (contactEmail) {
        slide.addText(`Adresse électronique : ${contactEmail}`, {
          x: leftX,
          y,
          fontSize: 18,
          color: "FFFFFF"
        });
        y += lineH;
      }
      y += 0.2;
    }

    // Interlocuteur EDF (Commercial)
    if (commercial || commercialPhone || commercialEmail) {
      slide.addText(`Interlocuteur EDF : ${commercial || ""}`, {
        x: leftX,
        y,
        fontSize: 18,
        bold: true,
        color: "FFFFFF"
      });
      y += lineH;
      if (commercialPhone) {
        slide.addText(`Mobile : ${commercialPhone}`, {
          x: leftX,
          y,
          fontSize: 18,
          color: "FFFFFF"
        });
        y += lineH;
      }
      if (commercialEmail) {
        slide.addText(`Adresse électronique : ${commercialEmail}`, {
          x: leftX,
          y,
          fontSize: 18,
          color: "FFFFFF"
        });
        y += lineH;
      }
    }

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
  // Mise en page commune des slides checklist
  // -----------------------------------------------------------
  const SLIDE_W = 10.0;
  const SLIDE_H = 5.625;
  const MARGIN  = 0.5;

  const IMG = { x: MARGIN, y: 1.4, w: 5.8, h: 3.8 };
  const BOX = { x: 6.6, y: 1.4, w: 3.0, h: 3.8 };

  const TGBT_RECT = { w: 1.6, h: 1.1, dx: 1.0, dy: 0.8 };
  const BORNE_RECT = { w: 1.6, h: 1.1, dx: 3.2, dy: 2.0 };

  const GREEN_CIRCLE = {
    w: 1.8,
    h: 1.8,
    x: BOX.x + 0.2,
    y: BOX.y + BOX.h + 0.15,
    stroke: "00FF00",
    strokeWidth: 3
  };

  function addLegend(slide, items) {
    if (!items.length) return;

    slide.addText(items.join("\n"), {
      x: SLIDE_W - 3.7,
      y: SLIDE_H - 1.2,
      w: 3.5,
      h: 1.2,
      fontSize: 12,
      color: "000000",
      valign: "top"
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
    const legend = [];

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

      legend.push(
        "Rectangle rouge = TGBT",
        "Cercle vert = place à équiper",
        "Rectangle bleu = future borne"
      );
    }

    if (low.includes("plans") || low.includes("places à électrifier")) {
      slide.addShape(RECT, {
        x: IMG.x + BORNE_RECT.dx,
        y: IMG.y + BORNE_RECT.dy,
        w: BORNE_RECT.w,
        h: BORNE_RECT.h,
        fill: { color: "0070C0" },
        line: { color: "003A70" }
      });

      legend.push("Rectangle bleu = future borne");
    }

    addLegend(slide, legend);
    addEDFLogo(slide);
  }

  // -----------------------------------------------------------
  // Slides checklist
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
      { base: "Plan d'implantation", file: "file1",  comment: "comment1"  },
      { base: "Plan d'implantation", file: "file1b", comment: "comment1b" },
      { base: "Places à électrifier", file: "file2",  comment: "comment2"  },
      { base: "Places à électrifier", file: "file2b", comment: "comment2b" },
      { base: "TGBT + disjoncteur de tête", file: "file3",  comment: "comment3"  },
      { base: "TGBT + disjoncteur de tête", file: "file3b", comment: "comment3b" },
      { base: "Cheminement", file: "file4",  comment: "comment4"  },
      { base: "Cheminement", file: "file4b", comment: "comment4b" },
      { base: "Plan du site", file: "file5",  comment: "comment5"  },
      { base: "Plan du site", file: "file5b", comment: "comment5b" },
      { base: "Éléments complémentaires", file: "file6",  comment: "comment6"  },
      { base: "Éléments complémentaires", file: "file6b", comment: "comment6b" }
    ];

    let remaining = defs.length;

    defs.forEach((item) => {
      const slide = pptx.addSlide();
      const sec = sectionNumber[item.base];

      slide.addText(`${sec}. ${item.base}`, {
        x: 0.5,
        y: 0.3,
        w: SLIDE_W - 1,
        fontSize: 36,
        bold: true,
        color: "0070C0",
        align: "center"
      });

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
          pptx.writeFile({ fileName: "Borne_Electrique_Projet.pptx" })
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

