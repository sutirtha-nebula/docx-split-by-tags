import fs from "fs";
import PizZip from "pizzip";
import { DOMParser, XMLSerializer } from "xmldom";

/**
 * Checks if paragraph style name means non-content
 */
function isNonContentStyle(styleName) {
    if (!styleName) return false;
    const nonContentPrefixes = [
        "Heading",
        "Title",
        "Subtitle",
        "Caption",
        "TOC",
        "Header",
        "Footer",
        "SourceCode"
    ];
    return nonContentPrefixes.some(prefix => styleName.startsWith(prefix));
}

/**
 * Checks if paragraph has center alignment
 */
function isCenteredParagraph(p) {
    const pPr = p.getElementsByTagName("w:pPr")[0];
    if (!pPr) return false;

    const jc = pPr.getElementsByTagName("w:jc")[0];
    if (!jc) return false;

    const val = jc.getAttribute("w:val");
    return val === "center";
}

/**
 * Checks if paragraph has any run with font size > 16pt (w:sz is half-points, so > 32)
 */
function hasLargeFontSize(p) {
    const runs = p.getElementsByTagName("w:r");
    for (let i = 0; i < runs.length; i++) {
        const run = runs[i];
        const rPr = run.getElementsByTagName("w:rPr")[0];
        if (!rPr) continue;

        const sz = rPr.getElementsByTagName("w:sz")[0];
        if (!sz) continue;

        const val = sz.getAttribute("w:val");
        if (!val) continue;

        const size = parseInt(val, 10);

        if (size > 30) { // 16pt * 2 (half points)
            return true;
        }
    }
    return false;
}

/**
 * Centralized check: Is paragraph non-content?
 */
function isNonContentParagraph(p) {
    const pPr = p.getElementsByTagName("w:pPr")[0];
    let styleName = null;
    if (pPr) {
        const pStyle = pPr.getElementsByTagName("w:pStyle")[0];
        if (pStyle) {
            styleName = pStyle.getAttribute("w:val");
        }
    }

    return isNonContentStyle(styleName) || isCenteredParagraph(p) || hasLargeFontSize(p);
}

// Main processing

const zip = new PizZip(fs.readFileSync("./templated1.doc/final_result.docx"));
const xml = zip.file("word/document.xml").asText();

const doc = new DOMParser().parseFromString(xml, "text/xml");
const paragraphs = doc.getElementsByTagName("w:p");

for (let i = 0; i < paragraphs.length; i++) {
    const p = paragraphs[i];

    if (isNonContentParagraph(p)) {
        // Skip this paragraph (heading, title, centered, or large font)
        continue;
    }

    // Color runs in this paragraph dark blue
    const runs = p.getElementsByTagName("w:r");
    for (let r = 0; r < runs.length; r++) {
        const run = runs[r];

        let rPr = run.getElementsByTagName("w:rPr")[0];
        if (!rPr) {
            rPr = doc.createElement("w:rPr");
            run.insertBefore(rPr, run.firstChild);
        }

        let color = rPr.getElementsByTagName("w:color")[0];
        if (!color) {
            color = doc.createElement("w:color");
            rPr.appendChild(color);
        }
        color.setAttribute("w:val", "00008B"); // Dark blue color
    }
}

zip.file("word/document.xml", new XMLSerializer().serializeToString(doc));
fs.writeFileSync("output.docx", zip.generate({ type: "nodebuffer" }));

console.log("Content paragraph colors updated successfully!");
