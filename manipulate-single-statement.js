import fs from "fs";
import PizZip from "pizzip";
import { DOMParser, XMLSerializer } from "xmldom";

/* ============================
   Helpers
============================ */

function getOrCreate(parent, tag, doc) {
    let el = parent.getElementsByTagName(tag)[0];
    if (!el) {
        el = doc.createElement(tag);
        parent.appendChild(el);
    }
    return el;
}

function getOrCreateRPr(run, doc) {
    return getOrCreate(run, "w:rPr", doc);
}

/* ============================
   Formatting functions
============================ */

function changeFontSize(run, sizePt, doc) {
    const rPr = getOrCreateRPr(run, doc);
    let sz = rPr.getElementsByTagName("w:sz")[0];
    if (!sz) {
        sz = doc.createElement("w:sz");
        rPr.appendChild(sz);
    }
    sz.setAttribute("w:val", String(sizePt * 2));
}

function changeFontFamily(run, fontName, doc) {
    const rPr = getOrCreateRPr(run, doc);
    let rFonts = rPr.getElementsByTagName("w:rFonts")[0];
    if (!rFonts) {
        rFonts = doc.createElement("w:rFonts");
        rPr.appendChild(rFonts);
    }
    rFonts.setAttribute("w:ascii", fontName);
    rFonts.setAttribute("w:hAnsi", fontName);
}

function changeFontColor(run, hexColor, doc) {
    const rPr = getOrCreateRPr(run, doc);
    let color = rPr.getElementsByTagName("w:color")[0];
    if (!color) {
        color = doc.createElement("w:color");
        rPr.appendChild(color);
    }
    color.setAttribute("w:val", hexColor);
}

/* ============================
   Apply formatting to specific text
============================ */

function applyFormattingToText(p, searchText, doc, formatter) {
    const runs = Array.from(p.getElementsByTagName("w:r"));

    let fullText = "";
    const runMap = [];

    // Build text + position map
    for (const run of runs) {
        const t = run.getElementsByTagName("w:t")[0];
        if (!t) continue;

        const text = t.textContent || "";
        const start = fullText.length;
        fullText += text;
        const end = fullText.length;

        runMap.push({ run, start, end });
    }

    const matchIndex = fullText.indexOf(searchText);
    if (matchIndex === -1) return;

    const matchStart = matchIndex;
    const matchEnd = matchIndex + searchText.length;

    // Apply formatting only to overlapping runs
    for (const { run, start, end } of runMap) {
        if (start < matchEnd && end > matchStart) {
            formatter(run);
        }
    }
}

/* ============================
   Main
============================ */

const zip = new PizZip(fs.readFileSync("./templates/2026_01_Precision_AI_UFA_Template.docx"));
const xml = zip.file("word/document.xml").asText();

const doc = new DOMParser().parseFromString(xml, "text/xml");
const paragraphs = Array.from(doc.getElementsByTagName("w:p"));

const TARGET_TEXT = "Name of the book";

for (const p of paragraphs) {
    applyFormattingToText(p, TARGET_TEXT, doc, (run) => {
        changeFontFamily(run, "DejaVu Sans", doc);
        changeFontSize(run, 30, doc);
        changeFontColor(run, "FFC0CB", doc);
    });
}

zip.file("word/document.xml", new XMLSerializer().serializeToString(doc));
fs.writeFileSync("output.docx", zip.generate({ type: "nodebuffer" }));

console.log("✅ Correct runs formatted");
