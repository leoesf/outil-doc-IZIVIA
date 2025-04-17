document.addEventListener("DOMContentLoaded", function() {
    if (typeof PptxGenJS !== 'undefined') {
        console.log("PptxGenJS est chargé correctement.");
    } else {
        console.log("PptxGenJS n'est pas chargé.");
    }
});

function createPowerPoint() {
    console.log("Début de la création du PowerPoint");
    if (typeof PptxGenJS === 'undefined') {
        console.error("PptxGenJS n'est pas chargé.");
        return;
    }
    var pptx = new PptxGenJS();
    var items = [
        { id: 'item1', file: 'file1', comment: 'comment1', title: "Plan d'implantation (image maps, Géoportail)" },
        { id: 'item2', file: 'file2', comment: 'comment2', title: "Places à électrifier" },
        { id: 'item3', file: 'file3', comment: 'comment3', title: "TGBT + disjoncteur de tête" },
        { id: 'item4', file: 'file4', comment: 'comment4', title: "Cheminement" },
        { id: 'item5', file: 'file5', comment: 'comment5', title: "Plan du site" },
        { id: 'item6', file: 'file6', comment: 'comment6', title: "Éléments complémentaires" }
    ];

    var completed = 0;
    items.forEach(function(item, index) {
        var fileInput = document.getElementById(item.file);
        var comment = document.getElementById(item.comment).value;
        var slide = pptx.addSlide();
        slide.addText(item.title, { x: 0.5, y: 0.5, fontSize: 24 });
        console.log("Traitement de l'élément : " + item.title);
        if (fileInput.files.length > 0) {
            var reader = new FileReader();
            reader.onload = function(e) {
                console.log("Lecture du fichier : " + fileInput.files[0].name);
                slide.addImage({ data: e.target.result, x: 0.5, y: 1.5, w: 8, h: 4.5 });
                slide.addText(comment, { x: 0.5, y: 6, fontSize: 18 });
                completed++;
                if (completed === items.length) {
                    console.log("Tous les éléments sont traités, génération du fichier PowerPoint");
                    pptx.writeFile({ fileName: "Projet_Borne_Electrique.pptx" });
                }
            };
            reader.readAsDataURL(fileInput.files[0]);
        } else {
            slide.addText(comment, { x: 0.5, y: 1.5, fontSize: 18 });
            completed++;
            if (completed === items.length) {
                console.log("Tous les éléments sont traités, génération du fichier PowerPoint");
                pptx.writeFile({ fileName: "Projet_Borne_Electrique.pptx" });
            }
        }
    });
}
