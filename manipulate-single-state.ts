import * as fs from "fs";
import PizZip from "pizzip";
import { DOMParser, XMLSerializer } from "xmldom";

/* ============================
   Types
============================ */

type FormatterFn = (run: Element) => void;

/* ============================
   Helpers
============================ */

function getOrCreate(parent: Element, tag: string, doc: Document): Element {
    let el = parent.getElementsByTagName(tag)[0];
    if (!el) {
        el = doc.createElement(tag);
        parent.appendChild(el);
    }
    return el;
}

function getOrCreateRPr(run: Element, doc: Document): Element {
    return getOrCreate(run, "w:rPr", doc);
}

/* ============================
   Formatting functions
============================ */

// 1️⃣ Font size (pt → half-points)
function changeFontSize(run: Element, sizePt: number, doc: Document): void {
    const rPr = getOrCreateRPr(run, doc);
    let sz = rPr.getElementsByTagName("w:sz")[0];
    if (!sz) {
        sz = doc.createElement("w:sz");
        rPr.appendChild(sz);
    }
    sz.setAttribute("w:val", String(sizePt * 2));
}

// 2️⃣ Font family
function changeFontFamily(run: Element, fontName: string, doc: Document): void {
    const rPr = getOrCreateRPr(run, doc);
    let rFonts = rPr.getElementsByTagName("w:rFonts")[0];
    if (!rFonts) {
        rFonts = doc.createElement("w:rFonts");
        rPr.appendChild(rFonts);
    }
    rFonts.setAttribute("w:ascii", fontName);
    rFonts.setAttribute("w:hAnsi", fontName);
}

// 3️⃣ Font color (HEX without #)
function changeFontColor(run: Element, hexColor: string, doc: Document): void {
    const rPr = getOrCreateRPr(run, doc);
    let color = rPr.getElementsByTagName("w:color")[0];
    if (!color) {
        color = doc.createElement("w:color");
        rPr.appendChild(color);
    }
    color.setAttribute("w:val", hexColor);
}

/* ============================
   Apply formatting to text inside ONE paragraph
============================ */

function applyFormattingToText(
    paragraph: Element,
    searchText: string,
    doc: Document,
    formatter: FormatterFn
): void {
    const runs = Array.from(paragraph.getElementsByTagName("w:r"));

    let fullText = "";
    const runMap: { run: Element; start: number; end: number }[] = [];

    // Build paragraph text + run position map
    for (const run of runs) {
        const t = run.getElementsByTagName("w:t")[0];
        if (!t) continue;

        const text = t.textContent ?? "";
        const start = fullText.length;
        fullText += text;
        const end = fullText.length;

        runMap.push({ run, start, end });
    }

    // 🔁 Find and process ALL matches
    let searchFrom = 0;

    while (true) {
        const matchIndex = fullText.indexOf(searchText, searchFrom);
        if (matchIndex === -1) break;

        const matchStart = matchIndex;
        const matchEnd = matchIndex + searchText.length;


        // Apply formatter only to overlapping runs
        for (const { run, start, end } of runMap) {
            console.log(start, end, matchIndex, matchEnd);

            if (start < matchEnd && end > matchStart) {
                formatter(run);
            }
        }

        // Move forward to find next occurrence
        searchFrom = matchEnd;
    }
}

/* ============================
   Iterate all paragraphs (NEW FUNCTION)
============================ */

function formatTextInDocument(
    doc: Document,
    targetText: string,
    formatter: FormatterFn
): void {
    const paragraphs = Array.from(doc.getElementsByTagName("w:p"));

    for (const paragraph of paragraphs) {
        applyFormattingToText(paragraph, targetText, doc, formatter);
    }
}

/* ============================
   Main
============================ */

const zip = new PizZip(fs.readFileSync("./templates/2026_01_Precision_AI_UFA_Template.docx"));
const xml = zip.file("word/document.xml")!.asText();

const doc = new DOMParser().parseFromString(xml, "text/xml");

formatTextInDocument(doc, "Lorem Ipsum", (run) => {
    changeFontFamily(run, "DejaVu Sans", doc);
    changeFontSize(run, 12, doc);
    changeFontColor(run, "FFC0CB", doc);
});

/* ============================
   Save
============================ */

zip.file("word/document.xml", new XMLSerializer().serializeToString(doc));
fs.writeFileSync("output.docx", zip.generate({ type: "nodebuffer" }));

console.log("✅ Correct runs formatted");
