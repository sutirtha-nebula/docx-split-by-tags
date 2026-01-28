import * as fs from "fs";
import PizZip from "pizzip";
import { DOMParser, XMLSerializer } from "xmldom";

/* ============================
   Namespaces
============================ */
const NS: any = {
  w: "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
  rel: "http://schemas.openxmlformats.org/package/2006/relationships",
};

  export function extractFirstDoc(inputFile: string, outputFile: string, splitTag: string): boolean {
    const serializer = new XMLSerializer();
    const parser = new DOMParser();
    const srcZip: any = new PizZip(fs.readFileSync(inputFile));
    const docXml: any = parser.parseFromString(
        srcZip.file("word/document.xml").asText(),
        "text/xml"
    );

    const body: any = docXml.getElementsByTagNameNS(NS.w, "body")[0];
    const nodes: any[] = Array.from(body.childNodes);

    let splitIndex = findSplitIndexByTagString(nodes, splitTag);

    if (splitIndex === -1) {
      fs.writeFileSync(outputFile, fs.readFileSync(inputFile));
      console.log("ℹ️ No split tag found — copied full doc");
      return true;
    }

    const newBody: any = docXml.createElementNS(NS.w, "body");

    // if (splitIndex === nodes.length) {
    //   splitIndex = splitIndex - 1;
    // }

    for (let i = 0; i < splitIndex; i++) {
      newBody.appendChild(nodes[i].cloneNode(true));
    }

    // Append section properties if any
    const sectPrs = docXml.getElementsByTagNameNS(NS.w, "sectPr");
    if (sectPrs.length > 0) {
      const lastPara = docXml.createElementNS(NS.w, "w:p");
      const lastParaPr = docXml.createElementNS(NS.w, "w:pPr");
      lastParaPr.appendChild(sectPrs[sectPrs.length - 1].cloneNode(true));
      lastPara.appendChild(lastParaPr);
      newBody.appendChild(lastPara);
    }

    body.parentNode.replaceChild(newBody, body);

    const outZip = cloneZip(
        srcZip,
        serializer.serializeToString(docXml)
    );

    fs.writeFileSync(outputFile, outZip.generate({ type: "nodebuffer" }));
    console.log(`✅ First doc written → ${outputFile}`);
    return true;
  }

  export function extractRemainingDoc(inputFile: string, outputFile: string, splitTag: string): boolean {
    const serializer = new XMLSerializer();
    const parser = new DOMParser();
    const srcZip: any = new PizZip(fs.readFileSync(inputFile));
    const docXml: any = parser.parseFromString(
        srcZip.file("word/document.xml").asText(),
        "text/xml"
    );

    const body: any = docXml.getElementsByTagNameNS(NS.w, "body")[0];
    const nodes: any[] = Array.from(body.childNodes);

    const splitIndex = findSplitIndexByTagString(nodes, splitTag);

    if (splitIndex === -1 || splitIndex === (nodes.length - 1)) {
      fs.writeFileSync(outputFile, fs.readFileSync(inputFile));
      console.log("ℹ️ No split tag found — copied full doc");
      return true;
    }

    const newBody: any = docXml.createElementNS(NS.w, "body");


    for (let i = splitIndex; i < nodes.length; i++) {
      if (i === splitIndex) {
        // Remove split tag from first node
        const cleanedNode = removeSplitTagByString(nodes[i], splitTag);
        newBody.appendChild(cleanedNode);
      } else {
        newBody.appendChild(nodes[i].cloneNode(true));
      }
    }
    body.parentNode.replaceChild(newBody, body);

    // Fix section properties - remove header/footer refs to prevent duplication
    const sectPrs = docXml.getElementsByTagNameNS(NS.w, "sectPr");
    for (let sc = 0; sc < sectPrs.length; sc++) {
      const sectPr = sectPrs[sc];

      // const headerRefs = sectPr.getElementsByTagNameNS(NS.w, "headerReference");
      // const footerRefs = sectPr.getElementsByTagNameNS(NS.w, "footerReference");
      const titlePg = sectPr.getElementsByTagNameNS(NS.w, "titlePg");

      // Remove header/footer references
      // for (let i = headerRefs.length - 1; i >= 0; i--) {
      //   sectPr.removeChild(headerRefs[i]);
      // }
      // for (let i = footerRefs.length - 1; i >= 0; i--) {
      //   sectPr.removeChild(footerRefs[i]);
      // }
      for (let i = titlePg.length - 1; i >= 0; i--) {
        sectPr.removeChild(titlePg[i]);
      }

      // Page numbering starts at 2
      let pgNum = sectPr.getElementsByTagNameNS(NS.w, "pgNumType")[0];
      if (!pgNum) {
        pgNum = docXml.createElementNS(NS.w, "pgNumType");
        sectPr.appendChild(pgNum);
      }
      pgNum.setAttribute("w:start", "2");

      // newBody.appendChild(sectPr);
    }
/*
    if (sectPrs.length > 0) {
      const sectPr = sectPrs[sectPrs.length - 1].cloneNode(true);

      const headerRefs = sectPr.getElementsByTagNameNS(NS.w, "headerReference");
      const footerRefs = sectPr.getElementsByTagNameNS(NS.w, "footerReference");
      const titlePg = sectPr.getElementsByTagNameNS(NS.w, "titlePg");

      // Remove header/footer references
      for (let i = headerRefs.length - 1; i >= 0; i--) {
        sectPr.removeChild(headerRefs[i]);
      }
      for (let i = footerRefs.length - 1; i >= 0; i--) {
        sectPr.removeChild(footerRefs[i]);
      }
      for (let i = titlePg.length - 1; i >= 0; i--) {
        sectPr.removeChild(titlePg[i]);
      }

      // Page numbering starts at 2
      let pgNum = sectPr.getElementsByTagNameNS(NS.w, "pgNumType")[0];
      if (!pgNum) {
        pgNum = docXml.createElementNS(NS.w, "pgNumType");
        sectPr.appendChild(pgNum);
      }
      pgNum.setAttribute("w:start", "2");

      newBody.appendChild(sectPr);
    }
*/


    const outZip = cloneZip(
        srcZip,
        serializer.serializeToString(docXml)
    );

    fs.writeFileSync(outputFile, outZip.generate({ type: "nodebuffer" }));
    console.log(`✅ Remaining doc written → ${outputFile}`);
    return true;
  }

  function cloneZip(srcZip: any, documentXml: string): any {
    const outZip = new PizZip();

    Object.keys(srcZip.files).forEach((name: string) => {
      if (name === "word/document.xml") return;

      const file = srcZip.file(name);
      if (!file) return;

      if (name.includes("media/")) {
        outZip.file(name, file.asNodeBuffer());
      } else {
        outZip.file(name, file.asText());
      }
    });

    outZip.file("word/document.xml", documentXml);
    return outZip;
  }

  function removeSplitTagByString(node: any, splitTag: string): any {
    const serializer = new XMLSerializer();
    const parser = new DOMParser();
    const nodeXml = serializer.serializeToString(node);
    const cleanedXml = nodeXml.replace(splitTag, "");
    return parser.parseFromString(cleanedXml, "text/xml").documentElement;
  }

  function findSplitIndexByTagString(nodes: any[], splitTag: string): number {
    const serializer = new XMLSerializer();
    for (let i = 0; i < nodes.length; i++) {
      const nodeXml = serializer.serializeToString(nodes[i]);
      if (nodeXml.includes(splitTag)) {
        return i + 1;
      }
    }
    return -1;
  }


/* ============================
   USAGE
============================ */

// const SPLIT_TAG = `<w:br w:type="page"/>`;
// or
// const SPLIT_TAG = `<w:lastRenderedPageBreak/>`;

const SPLIT_TAG = `<w:pageBreakBefore w:val="0"/>`

// const inputFile = "./2026_01_Precision_AI_UFA_Template.docx";
const inputFile = "./templates/Brochure.docx";


extractFirstDoc(inputFile, "main.docx", SPLIT_TAG);
extractRemainingDoc(inputFile, "content.docx", SPLIT_TAG);
