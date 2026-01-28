import fs from "fs";
import PizZip from "pizzip";
import { DOMParser, XMLSerializer } from "xmldom";

const zip = new PizZip(fs.readFileSync("./templates/2026_01_Precision_AI_UFA_Template.docx"));
const xml = zip.file("word/document.xml").asText();
const doc = new DOMParser().parseFromString(xml, "text/xml");

// ------------- CONFIG -------------
const searchText = "Lorem Ipsum"; // text to search for
const options = {
  color: "FFFFFF",
  bgColor: "FF0000",
  fontSize: 44,
  fontFamily: "Calibri",
  bold: true
};
// ---------------------------------

const paragraphs = doc.getElementsByTagName("w:p");

function createRun(doc, text, styled = false) {
  const run = doc.createElement("w:r");

  if (styled) {
    const rPr = doc.createElement("w:rPr");

    if (options.color) {
      const color = doc.createElement("w:color");
      color.setAttribute("w:val", options.color);
      rPr.appendChild(color);
    }

    if (options.bgColor) {
      const shd = doc.createElement("w:shd");
      shd.setAttribute("w:val", "clear");
      shd.setAttribute("w:color", "auto");
      shd.setAttribute("w:fill", options.bgColor);
      rPr.appendChild(shd);
    }

    if (options.fontSize) {
      const sz = doc.createElement("w:sz");
      sz.setAttribute("w:val", options.fontSize.toString());
      rPr.appendChild(sz);
    }

    if (options.fontFamily) {
      const rFonts = doc.createElement("w:rFonts");
      rFonts.setAttribute("w:ascii", options.fontFamily);
      rFonts.setAttribute("w:hAnsi", options.fontFamily);
      rPr.appendChild(rFonts);
    }

    if (options.bold) {
      const b = doc.createElement("w:b");
      rPr.appendChild(b);
    }

    run.appendChild(rPr);
  }

  const t = doc.createElement("w:t");
  t.textContent = text;
  run.appendChild(t);

  return run;
}

function getParagraphText(p) {
  const texts = p.getElementsByTagName("w:t");
  let fullText = "";
  for (let i = 0; i < texts.length; i++) {
    fullText += texts[i].textContent || "";
  }
  return fullText;
}

// ------------ Search & Replace ------------
for (let i = 0; i < paragraphs.length; i++) {
  const p = paragraphs[i];
  const fullText = getParagraphText(p);

  if (!fullText.includes(searchText)) continue;

  rebuildParagraphWithMatches(p, doc, fullText, searchText);
}

// -------- Save --------
zip.file("word/document.xml", new XMLSerializer().serializeToString(doc));
fs.writeFileSync("output.docx", zip.generate({ type: "nodebuffer" }));

console.log("Sentence/chunk styled successfully!");

function rebuildParagraphWithMatches(p, doc, fullText, searchText) {
  // remove all existing runs
  const runs = Array.from(p.getElementsByTagName("w:r"));
  runs.forEach(r => p.removeChild(r));

  let remaining = fullText;

  while (remaining.includes(searchText)) {
    const index = remaining.indexOf(searchText);

    const before = remaining.slice(0, index);
    const match = remaining.slice(index, index + searchText.length);

    if (before) {
      p.appendChild(createRun(doc, before, false));
    }

    // styled chunk
    p.appendChild(createRun(doc, match, true));

    remaining = remaining.slice(index + searchText.length);
  }

  // tail text
  if (remaining) {
    p.appendChild(createRun(doc, remaining, false));
  }
}

