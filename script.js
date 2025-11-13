// -----------------------------------------------------------
// script.js - Génération du PowerPoint (PptxGenJS v3.x)
// -----------------------------------------------------------
// Version de test : affichage "TEST COIN BAS DROITE" pour vérifier la position
// -----------------------------------------------------------

document.addEventListener("DOMContentLoaded", () => {
  console.log("🔥 VERSION TEST LEGEND – vTest1");

  if (typeof PptxGenJS === "undefined") {
    alert("❌ PptxGenJS n'est pas chargé.");
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
  pptx.layout = "LAYOUT_WIDE";

  const RECT    = PptxGenJS.ShapeType?.rect    || "rect";
  const ELLIPSE = PptxGenJS.ShapeType?.ellipse || "ellipse";

  // ---- Récupération valeurs formulaire ----
  const getVal = (id) => document.getElementById(id)?.value || "";
  const clientName = getVal("clientName");
  const raeClient  = getVal("raeClient");

  // ---------------- Slides Informations ----------------
  function addCoverSlide() {
    const slide = pptx.addSlide();
    slide.background = { fill: "363636" };

    slide.addText(`Client : ${clientName}`, {
      x: 0.5, y: 0.5,
      fontSize: 30, color: "FFFFFF", bold: true
    });

    slide.addText("Projet d’infrastructure de recharge", {
      x: 0.5, y: 1.3,
      fontSize: 20, color: "FFFFFF"
    });
  }

  function addInfoSlide() {
    const slide = pptx.addSlide();
    slide.addText("Compléments d’informations", {
      x: 0.5, y: 0.5,
      fontSize: 24, bold: true
    });

    slide.addText(raeClient || "—", {
      x: 0.5, y: 1.3,
      w: "90%", h: "70%",
      fontSize: 18,
      fill: { color: "FFFFFF" },
      line: { color: "CCCCCC" }
    });
  }

  // ---------------- Mise en page ----------------
  const SLIDE_W = 10.0;
  const SLIDE_H = 5.625;
  const MARGIN  = 0.5;

  const IMG = { x: MARGIN, y: 1.1, w: 6.1, h: 4.6 };
  const BOX = { x: 6.7, y: 1.1, w: SLIDE_W - 6.7 - MARGIN, h: 4.6 };

  // ---------------- Formes ----------------
  const TGBT_RECT = { w:1.6, h:1.1, dx:0.8, dy:0.6 };
  const BORNE_RECT = { w:1.6, h:1.1, dx:2.0, dy:1.8 };

  const GREEN_CIRCLE = {
    w: 1.8, h: 1.8,
    stroke: "00FF00",
    strokeWidth: 3,
    x: BOX.x,
    y: Math.min(BOX.y + BOX.h + 0.2, SLIDE_H - MARGIN - 1.8)
  };

  // ---------------- Légende : TEST TEMPORAIRE ----------------
  function addLegend(slide, texts = []) {
    console.log("🟢 TEST POSITION LÉGENDE – fonction active !");

    slide.addText("TEST COIN BAS DROITE", {
      x: 9.5,        // très proche du bord droit
      y: 5.3,        // très proche du bas
      w: 1.5,
      h: 0.5,
      fontSize: 20,
      color: "FF0000",
      bold: true
    });

    // (Le vrai affichage de légende est désactivé volontairement pour le test)
  }

  // ---------------- Placement image + formes + légendes ----------------
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

    const lower = title.toLowerCase();

    // Formes selon titre
    if (lower.includes("implantation")) {
      slide.addShape(RECT, {
        x: IMG.x + TGBT_RECT.dx,
        y: IMG.y + TGBT_RECT.dy,
        w: TGBT_RECT.w,
        h: TGBT_RECT.h,
        fill:{color:"FF0000"},
        line:{color:"880000"}
      });

      slide.addShape(ELLIPSE, {
        x: GREEN_CIRCLE.x,
        y: GREEN_CIRCLE.y,
        w: GREEN_CIRCLE.w,
        h: GREEN_CIRCLE.h,
        fill: null,
        line: { color:GREEN_CIRCLE.stroke, width:GREEN_CIRCLE.strokeWidth }
      });
    }

    if (lower.includes("places à électrifier") || lower.includes("places a electrifier")) {
      slide.addShape(RECT, {
        x: IMG.x + BORNE_RECT.dx,
        y: IMG.y + BORNE_RECT.dy,
        w: BORNE_RECT.w,
        h: BORNE_RECT.h,
        fill:{color:"0070C0"},
        line:{color:"003A70"}
      });
    }

    // → Ajout de la légende TEST
    addLegend(slide);
  }

  // ---------------- Slides checklist ----------------
  function addChecklistSlides() {
    const items = [
      { title:"Plan d'implantation #1", file:"file1",  comment:"comment1"  },
      { title:"Plan d'implantation #2", file:"file1b", comment:"comment1b" },

      { title:"Places à électrifier #1", file:"file2",  comment:"comment2"  },
      { title:"Places à électrifier #2", file:"file2b", comment:"comment2b" },

      { title:"TGBT + disjoncteur de tête #1", file:"file3",  comment:"comment3"  },
      { title:"TGBT + disjoncteur de tête #2", file:"file3b", comment:"comment3b" },

      { title:"Cheminement #1", file:"file4",  comment:"comment4"  },
      { title:"Cheminement #2", file:"file4b", comment:"comment4b" },

      { title:"Plan du site #1", file:"file5",  comment:"comment5"  },
      { title:"Plan du site #2", file:"file5b", comment:"comment5b" },

      { title:"Éléments complémentaires #1", file:"file6",  comment:"comment6"  },
      { title:"Éléments complémentaires #2", file:"file6b", comment:"comment6b" },
    ];

    let done = 0;

    items.forEach((item) => {
      const slide = pptx.addSlide();

      slide.addText(item.title, {
        x: MARGIN,
        y: 0.5,
        fontSize: 24,
        bold: true
      });

      const comment = document.getElementById(item.comment)?.value || "—";
      slide.addText(comment, {
        x: BOX.x,
        y: BOX.y,
        w: BOX.w,
        h: BOX.h,
        fill: { color: "FFFFFF" },
        line: { color: "AAAAAA" },
        fontSize: 18,
        valign: "top",
        margin: 0.1
      });

      const fileInput = document.getElementById(item.file);

      const afterImage = (dataUrl) => {
        placeImageAndShapes(slide, item.title, IMG, dataUrl);

        done++;
        if (done === items.length) {
          const safeName = (clientName || "Projet")
            .replace(/[^\p{L}\p{N}_\- ]/gu, "")
            .trim()
            .replace(/\s+/g, "_");

          pptx.writeFile({ fileName:`Borne_Electrique_${safeName}.pptx` })
            .finally(() => {
              btn?.removeAttribute("aria-busy");
              btn?.removeAttribute("disabled");
            });
        }
      };

      if (fileInput?.files.length > 0) {
        const reader = new FileReader();
        reader.onload = (e) => afterImage(e.target.result);
        reader.readAsDataURL(fileInput.files[0]);
      } else {
        afterImage(null);
      }
    });
  }

  // ---------------- Lancement ----------------
  addCoverSlide();
  addInfoSlide();
  addChecklistSlides();
}
