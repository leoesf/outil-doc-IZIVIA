document.addEventListener("DOMContentLoaded", function () {
    if (typeof PptxGenJS !== 'undefined') {
        console.log("PptxGenJS est chargé correctement.");
    } else {
        console.log("PptxGenJS n'est pas chargé.");
    }
});

function createPowerPoint() {
    const pptx = new PptxGenJS();

    // 1. Slide de couverture
    const coverSlide = pptx.addSlide();
    coverSlide.background = { fill: '363636' }; // gris foncé

    const coverImageInput = document.getElementById("coverImage");
    if (coverImageInput.files.length > 0) {
        const reader = new FileReader();
        reader.onload = function (e) {
            coverSlide.addImage({ data: e.target.result, x: 0, y: 0, w: '100%', h: '100%' });
            addRAESlide(pptx);
        };
        reader.readAsDataURL(coverImageInput.files[0]);
    } else {
        addRAESlide(pptx);
    }

    function addRAESlide(pptx) {
        // 2. Slide RAE du client
        const raeText = document.getElementById("raeClient").value;
        const raeSlide = pptx.addSlide();
        raeSlide.addText("RAE du client", { x: 0.5, y: 0.5, fontSize: 24, bold: true });
        raeSlide.addText(raeText, { x: 0.5, y: 1.5, fontSize: 18, w: '90%', h: '70%', color: '363636' });

        // 3. Slides checklist
        const items = [
            { id: 'item1', file: 'file1', comment: 'comment1', title: "Plan d'implantation (image maps, Géoportail)" },
            { id: 'item2', file: 'file2', comment: 'comment2', title: "Places à électrifier" },
            { id: 'item3', file: 'file3', comment: 'comment3', title: "TGBT + disjoncteur de tête" },
            { id: 'item4', file: 'file4', comment: 'comment4', title: "Cheminement" },
            { id: 'item5', file: 'file5', comment: 'comment5', title: "Plan du site" },
            { id: 'item6', file: 'file6', comment: 'comment6', title: "Éléments complémentaires" }
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
    }
}
