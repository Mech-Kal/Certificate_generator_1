function log(msg) {
    document.getElementById("log").textContent += msg + "\n";
}

function openModal() {
    document.getElementById("howToUseModal").style.display = "block";
}

function closeModal() {
    document.getElementById("howToUseModal").style.display = "none";
}

window.onclick = function(e) {
    const modal = document.getElementById("howToUseModal");
    if (e.target === modal) modal.style.display = "none";
};

async function generateCertificates() {
    document.getElementById("log").textContent = "";
    log("Starting...");

    const templateFile = document.getElementById("template").files[0];
    const excelFile = document.getElementById("excel").files[0];

    if (!templateFile || !excelFile) {
        alert("Upload both DOCX and Excel files.");
        return;
    }

    if (typeof PizZip === "undefined") {
        log("ERROR: PizZip NOT loaded!");
        return;
    }
    if (typeof docxtemplater === "undefined") {
        log("ERROR: Docxtemplater NOT loaded!");
        return;
    }

    log("Libraries loaded successfully.");

    const templateBuffer = await templateFile.arrayBuffer();

    const excelBuffer = await excelFile.arrayBuffer();
    const workbook = XLSX.read(excelBuffer, { type: "array" });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet);

    log("Excel rows found: " + rows.length);

    rows.forEach((row, i) => {
        const name = row.name;
        const rank = row.rank;

        if (!name) {
            log(`Row ${i+1}: Missing name, skipping`);
            return;
        }

        log(`Generating certificate for: ${name}`);

        const zip = new PizZip(templateBuffer);
        const doc = new docxtemplater(zip, {
            paragraphLoop: true,
            linebreaks: true,
        });

        doc.setData({ name, rank });

        try {
            doc.render();
        } catch (error) {
            log("Render error: " + error);
            return;
        }

        const output = doc.getZip().generate({
            type: "blob",
            mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        });

        const link = document.createElement("a");
        link.href = URL.createObjectURL(output);
        link.download = `${name}_certificate.docx`;
        link.click();
    });

    log("DONE! All certificates generated.");
}
