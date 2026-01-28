import * as fs from "fs";
import PizZip from "pizzip";
import { DOMParser, XMLSerializer } from "xmldom";

type Mode = "title" | "header" | "footer" | "heading" | "searchReplace" | "replaceContent";


interface Options {
  title?: string;

  prefixText?: string;
  suffixText?: string;

  color?: string;
  fontSize?: number;
  fontFamily?: string;
  underline?: boolean;
  bold?: boolean;

  bgColor?: string;

  borderColor?: string;
  borderSize?: number;
  borderStyle?: string;

  // search mode
  searchText?: string;
}

const zip = new PizZip(fs.readFileSync("./templates/2026_01_Precision_AI_UFA_Template.docx"));

let mode: Mode = "replaceContent";

const options: Options = {
  title: "My Document Title",

  prefixText: "[PRE] ",
  suffixText: " (POST)",

  color: "FF0000",
  fontSize: 20,
  fontFamily: "Arial",
  underline: true,
  bold: true,

  bgColor: "D9EAD3",

  borderColor: "FF0000",
  borderSize: 16,
  borderStyle: "single",

  searchText: "Executive Summary",
};

function addPrefixSuffix(text: string, opts: Options): string {
  return `${opts.prefixText || ""}${text}${opts.suffixText || ""}`;
}

function applyStyleToRuns(runs: any, doc: any, opts: Options) {
  for (let r = 0; r < runs.length; r++) {
    const run = runs[r];

    let rPr = run.getElementsByTagName("w:rPr")[0];
    if (!rPr) {
      rPr = doc.createElement("w:rPr");
      run.insertBefore(rPr, run.firstChild);
    }

    // Bold
    if (opts.bold) {
      let b = rPr.getElementsByTagName("w:b")[0];
      if (!b) {
        b = doc.createElement("w:b");
        rPr.appendChild(b);
      }
    }

    // Color
    if (opts.color) {
      let color = rPr.getElementsByTagName("w:color")[0];
      if (!color) {
        color = doc.createElement("w:color");
        rPr.appendChild(color);
      }
      color.setAttribute("w:val", opts.color);
    }

    // Font size
    if (opts.fontSize) {
      let sz = rPr.getElementsByTagName("w:sz")[0];
      if (!sz) {
        sz = doc.createElement("w:sz");
        rPr.appendChild(sz);
      }
      sz.setAttribute("w:val", opts.fontSize.toString());
    }

    // Font family
    if (opts.fontFamily) {
      let rFonts = rPr.getElementsByTagName("w:rFonts")[0];
      if (!rFonts) {
        rFonts = doc.createElement("w:rFonts");
        rPr.appendChild(rFonts);
      }
      rFonts.setAttribute("w:ascii", opts.fontFamily);
      rFonts.setAttribute("w:hAnsi", opts.fontFamily);
    }

    // Underline
    if (opts.underline) {
      let u = rPr.getElementsByTagName("w:u")[0];
      if (!u) {
        u = doc.createElement("w:u");
        rPr.appendChild(u);
      }
      u.setAttribute("w:val", "single");
    }
  }
}

function applyBorder(p: any, doc: any, opts: Options) {
  let pPr = p.getElementsByTagName("w:pPr")[0];
  if (!pPr) {
    pPr = doc.createElement("w:pPr");
    p.insertBefore(pPr, p.firstChild);
  }

  // Remove existing border if exists
  let oldBdr = pPr.getElementsByTagName("w:pBdr")[0];
  if (oldBdr) pPr.removeChild(oldBdr);

  const pBdr = doc.createElement("w:pBdr");
  pPr.appendChild(pBdr);

  const sides = ["top", "left", "bottom", "right"];

  sides.forEach(side => {
    const node = doc.createElement(`w:${side}`);
    node.setAttribute("w:val", opts.borderStyle || "single");
    node.setAttribute("w:sz", (opts.borderSize || 8).toString());
    node.setAttribute("w:space", "1");
    node.setAttribute("w:color", opts.borderColor || "000000");
    pBdr.appendChild(node);
  });
}

function isHeadingStyle(p: any): boolean {
  const pPr = p.getElementsByTagName("w:pPr")[0];
  if (!pPr) return false;
  const pStyle = pPr.getElementsByTagName("w:pStyle")[0];
  if (!pStyle) return false;
  const styleVal = pStyle.getAttribute("w:val");
  return !!styleVal && styleVal.startsWith("Heading");
}

/* -------- SEARCH / REPLACE HELPERS -------- */

function getParagraphText(p: any): string {
  const texts = p.getElementsByTagName("w:t");
  let fullText = "";
  for (let i = 0; i < texts.length; i++) {
    fullText += texts[i].textContent || "";
  }
  return fullText;
}

function createRun(doc: any, text: string, styled: boolean) {
  const run = doc.createElement("w:r");

  if (styled) {
    const rPr = doc.createElement("w:rPr");
    run.appendChild(rPr);

    applyStyleToRuns([run], doc, options);
  }

  const t = doc.createElement("w:t");
  t.textContent = text;
  run.appendChild(t);

  return run;
}

function rebuildParagraphWithMatches(
  p: any,
  doc: any,
  fullText: string,
  searchText: string
) {
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

    p.appendChild(createRun(doc, match, true));
    remaining = remaining.slice(index + searchText.length);
  }

  if (remaining) {
    p.appendChild(createRun(doc, remaining, false));
  }
}

/* -------- REPLACE CONTENT HELPERS -------- */
function isNonContentStyle(styleName: string | null): boolean {
    if (!styleName) return false;
    const nonContentPrefixes = [
        "Heading",
        "Title",
        "Subtitle",
        "Caption",
        // "TOC",
        "Header",
        "Footer",
        "SourceCode"
    ];
    return nonContentPrefixes.some(prefix => styleName.startsWith(prefix));
}

/*** Checks if paragraph has center alignment ***/
function isCenteredParagraph(p: Element): boolean {
    const pPr = p.getElementsByTagName("w:pPr")[0];
    if (!pPr) return false;

    const jc = pPr.getElementsByTagName("w:jc")[0];
    if (!jc) return false;

    const val = jc.getAttribute("w:val");
    return val === "center";
}

/*** Checks if paragraph has any run with font size > 16pt (w:sz is half-points, so > 32) */
function hasLargeFontSize(p: Element): boolean {
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

        if (size > 32) { // 16pt * 2 (half points)
            return true;
        }
    }
    return false;
}

/*** Centralized check: Is paragraph non-content? */
function isNonContentParagraph(p: Element): boolean {
    const pPr = p.getElementsByTagName("w:pPr")[0];
    let styleName: string | null = null;
    if (pPr) {
        const pStyle = pPr.getElementsByTagName("w:pStyle")[0];
        if (pStyle) {
            styleName = pStyle.getAttribute("w:val");
        }
    }

    return isNonContentStyle(styleName) || isCenteredParagraph(p) || hasLargeFontSize(p);
}

// ---------------- TITLE ----------------
if (mode === ("title" as Mode)) {
  const coreXml = zip.file("docProps/core.xml")!.asText();
  const coreDoc = new DOMParser().parseFromString(coreXml, "text/xml");

  const titleNode = coreDoc.getElementsByTagName("dc:title")[0];
  const newTitleText = addPrefixSuffix(options.title || "", options);

  if (titleNode) {
    titleNode.textContent = newTitleText;
  } else {
    const root = coreDoc.documentElement;
    const newTitle = coreDoc.createElement("dc:title");
    newTitle.textContent = newTitleText;
    root.appendChild(newTitle);
  }

  zip.file("docProps/core.xml", new XMLSerializer().serializeToString(coreDoc));
  console.log("Title updated!");
}

// ---------------- HEADING ----------------
else if (mode === ("heading" as Mode)) {
  const xml = zip.file("word/document.xml")!.asText();
  const doc = new DOMParser().parseFromString(xml, "text/xml");
  const paragraphs = doc.getElementsByTagName("w:p");

  for (let i = 0; i < paragraphs.length; i++) {
    const p = paragraphs[i];
    if (!isHeadingStyle(p)) continue;

    const runs = p.getElementsByTagName("w:r");

    for (let r = 0; r < runs.length; r++) {
      const t = runs[r].getElementsByTagName("w:t")[0];
      if (t && t.textContent) {
        t.textContent = addPrefixSuffix(t.textContent, options);
      }
    }

    applyStyleToRuns(runs, doc, options);

    // shading background
    if (options.bgColor) {
      let pPr = p.getElementsByTagName("w:pPr")[0];
      if (!pPr) {
        pPr = doc.createElement("w:pPr");
        p.insertBefore(pPr, p.firstChild);
      }

      let shd = pPr.getElementsByTagName("w:shd")[0];
      if (!shd) {
        shd = doc.createElement("w:shd");
        pPr.appendChild(shd);
      }
      shd.setAttribute("w:fill", options.bgColor);
    }

    // border
    applyBorder(p, doc, options);
  }

  zip.file("word/document.xml", new XMLSerializer().serializeToString(doc));
  console.log("Headings updated!");
}

// ---------------- SEARCH / REPLACE ----------------
else if (mode === ("searchReplace" as Mode)) {
  if (!options.searchText) {
    throw new Error("searchText is required for searchReplace mode");
  }

  const xml = zip.file("word/document.xml")!.asText();
  const doc = new DOMParser().parseFromString(xml, "text/xml");
  const paragraphs = doc.getElementsByTagName("w:p");

  for (let i = 0; i < paragraphs.length; i++) {
    const p = paragraphs[i];
    const fullText = getParagraphText(p);

    if (!fullText.includes(options.searchText)) continue;

    rebuildParagraphWithMatches(
      p,
      doc,
      fullText,
      options.searchText
    );
  }

  zip.file("word/document.xml", new XMLSerializer().serializeToString(doc));
}

// ---------------- REPLACE CONTENT ----------------
else if (mode === ("replaceContent" as Mode)) {
  const xml = zip.file("word/document.xml")!.asText();
  const doc = new DOMParser().parseFromString(xml, "text/xml");
  const paragraphs = doc.getElementsByTagName("w:p");

  for (let i = 0; i < paragraphs.length; i++) {
    const p = paragraphs[i];
    const runs = p.getElementsByTagName("w:r");

    if (isNonContentParagraph(p)) {
        // Skip this paragraph (heading, title, centered, or large font)
        continue;
    }

    for (let r = 0; r < runs.length; r++) {
      const t = runs[r].getElementsByTagName("w:t")[0];
      if (t && t.textContent) {
        t.textContent = addPrefixSuffix(t.textContent, options);
      }
    }

    applyStyleToRuns(runs, doc, options);
  }

  zip.file("word/document.xml", new XMLSerializer().serializeToString(doc));
  console.log("Content replaced!");
}

fs.writeFileSync("output.docx", zip.generate({ type: "nodebuffer" }));
console.log("File saved!");
