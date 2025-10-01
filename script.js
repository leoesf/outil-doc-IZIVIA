document.addEventListener("DOMContentLoaded", () => {
  if (typeof PptxGenJS !== "undefined") {
    console.log("PptxGenJS est chargé correctement.");
  }
});

function createPowerPoint() {
  const pptx = new PptxGenJS();

  // Champs existants
  const clientName   = document.getElementById("clientName").value;
  const rae          = document.getElementById("rae").value;
  const power        = document.getElementById("power").value; // Puissance souscrite / max
  const commercial   = document.getElementById("commercial").value;
  const raeClient    = document.getElementById("raeClient").value;
  const coverImageInput = document.getElementById("coverImage");

  // ✅ NOUVEAUX CHAMPS (assure-toi d'avoir ajouté ces <input> dans borneelectrique.html)
  const clientAddress = document.getElementById("clientAddress")?.value || "";
  const siret         = document.getElementById("siret")?.value || "";
  const oppoNumber    = document.getElementById("oppoNumber")?.value || "";
  const nbBornes      = document.getElementById("nbBornes")?.value || "";
  const bornesPower   = document.getElementById("bornesPower")?.value || "";

  const addCoverSlide = (imageData = null) => {
    const slide = pptx.addSlide();
    slide.background = { fill: "363636" };

    // Bloc texte à gauche (on empile les lignes, seulement si renseignées)
    const textBox = [];
    if (clientName)   textBox.push({ text: `Client : ${clientName}\n`, options: { fontSize: 20, color: "FFFFFF", bold: true } });
    if (rae)          textBox.push({ text: `RAE : ${rae}\n`,                    options: { fontSize: 16, color: "FFFFFF" } });
    if (power)        textBox.push({ text: `Puissance : ${power}\n`,            options: { fontSize: 16, color: "FFFFFF" } });
    if (commercial)   textBox.push({ text: `Commercial : ${commercial}\n`,      options: { fontSize: 16, color: "FFFFFF" } });

    // ➕ Lignes ajoutées
    if (clientAddress) textBox.push({ text: `Adresse : ${clientAddress}\n`,     options: { fontSize: 16, color: "FFFFFF" } });
    if (siret)         textBox.push({ text: `SIRET : ${siret}\n`,               options: { fontSize: 16, color: "FFFFFF" } });
    if (oppoNumber)    textBox.push({ text: `Numéro Oppo : ${oppoNumber}\n`,    options: { fontSize: 16, color: "FFFFFF" } });
    if (nbBornes)      textBox.push({ text: `Nombre de bornes : ${nbBornes}\n`, options: { fontSize: 16, color: "FFFFFF" } });
    if (bornesPower)   textBox.push({ text: `Puissance des bornes : ${bornesPower}\n`, options: { fontSize: 16, color: "FFFFFF" } });

    textBox.push({ text: "Projet d’infrastructure de recharges pour véhicule électrique", options: { fontSize: 14, color: "FFFFFF", italic: true, breakLine: true } });

    // Texte à gauche
    slide.addText(textBox, { x: 0.5, y: 0.5, w: 5.5, h: 6.0 });

    // Image à droite
    if (imageData) {
      slide.addImage({ data: imageData, x: 6.5, y: 0.3, w: 3.8, h: 5.6, sizing: { type: "cover", w: 3.8, h: 5.6 } });
    }
  };

  const addRAESlide = () => {
    const slide = pptx.addSlide();
    slide.addText("RAE du client", { x: 0.5, y: 0.5, fontSize: 24, bold: true });
    slide.addText(raeClient, { x: 0.5, y: 1.5, fontSize: 18, w: "90%", h: "70%", color: "363636" });
  };

  const addChecklistSlides = () => {
    const items = [
      { file: "file1", comment: "comment1", title: "Plan d'implantation" },
      { file: "file2", comment: "comment2", title: "Places à électrifier" },
      { file: "file3", comment: "comment3", title: "TGBT + disjoncteur de tête" },
      { file: "file4", comment: "comment4", title: "Cheminement" },
      { file: "file5", comment: "comment5", title: "Plan du site" },
      { file: "file6", comment: "comment6", title: "Éléments complémentaires" }
    ];

    let completed = 0;
    items.forEach((item) => {
      const fileInput = document.getElementById(item.file);
      const comment = document.getElementById(item.comment).value;
      const slide = pptx.addSlide();
      slide.addText(item.title, { x: 0.5, y: 0.5, fontSize: 24 });

      if (fileInput.files.length > 0) {
        const reader = new FileReader();
        reader.onload = function (e) {
          slide.addImage({ data: e.target.result, x: 0.5, y: 1.5, w: 8, h: 4.5 });
          slide.addText(comment, { x: 0.5, y: 6, fontSize: 18 });
          checkCompletion();
        };
        reader.readAsDataURL(fileInput.files[0]);
      } else {
        slide.addText(comment, { x: 0.5, y: 1.5, fontSize: 18 });
        checkCompletion();
      }
    });

    function checkCompletion() {
      completed++;
      if (completed === items.length) {
        pptx.writeFile({ fileName: "Projet_Borne_Electrique.pptx" });
      }
    }
  };

  if (coverImageInput.files.length > 0) {
    const reader = new FileReader();
    reader.onload = function (e) {
      addCoverSlide(e.target.result);
      addRAESlide();
      addChecklistSlides();
    };
    reader.readAsDataURL(coverImageInput.files[0]);
  } else {
    addCoverSlide();
    addRAESlide();
    addChecklistSlides();
  }
}


    function checkCompletion() {
      completed++;
      if (completed === items.length) {
        pptx.writeFile({ fileName: "Projet_Borne_Electrique.pptx" });
      }
    }
  };

  if (coverImageInput.files.length > 0) {
    const reader = new FileReader();
    reader.onload = function (e) {
      addCoverSlide(e.target.result);
      addRAESlide();
      addChecklistSlides();
    };
    reader.readAsDataURL(coverImageInput.files[0]);
  } else {
    addCoverSlide();
    addRAESlide();
    addChecklistSlides();
  }
}

