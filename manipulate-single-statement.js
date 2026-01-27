import fs from "fs";
import PizZip from "pizzip";
import { DOMParser, XMLSerializer } from "xmldom";

const zip = new PizZip(fs.readFileSync("book.docx"));
const xml = zip.file("word/document.xml").asText();

const doc = new DOMParser().parseFromString(xml, "text/xml");
const paragraphs = doc.getElementsByTagName("w:p");

for (let i = 0; i < paragraphs.length; i++) {
    const p = paragraphs[i];

    // Collect visible paragraph text
    const texts = p.getElementsByTagName("w:t");
    let fullText = "";
    for (let j = 0; j < texts.length; j++) {
        fullText += texts[j].textContent;
    }

    if (fullText.trim() !== "Introduction") continue;

    const runs = p.getElementsByTagName("w:r");

    for (let r = 0; r < runs.length; r++) {
        const run = runs[r];

        let rPr = run.getElementsByTagName("w:rPr")[0];
        if (!rPr) {
            rPr = doc.createElement("w:rPr");
            run.insertBefore(rPr, run.firstChild);
        }

        // Sans-serif font
        let rFonts = rPr.getElementsByTagName("w:rFonts")[0];
        if (!rFonts) {
            rFonts = doc.createElement("w:rFonts");
            rPr.appendChild(rFonts);
        }

        rFonts.setAttribute("w:ascii", "DejaVu Sans");
        rFonts.setAttribute("w:hAnsi", "DejaVu Sans");

        // Size: 30pt → 60 half-points
        let sz = rPr.getElementsByTagName("w:sz")[0];
        if (!sz) {
            sz = doc.createElement("w:sz");
            rPr.appendChild(sz);
        }
        sz.setAttribute("w:val", "60");

        // 🎀 Color: pink
        let color = rPr.getElementsByTagName("w:color")[0];
        if (!color) {
            color = doc.createElement("w:color");
            rPr.appendChild(color);
        }
        color.setAttribute("w:val", "FFC0CB");
    }

    break; // stop after first matching heading
}

zip.file("word/document.xml", new XMLSerializer().serializeToString(doc));
fs.writeFileSync("output.docx", zip.generate({ type: "nodebuffer" }));
