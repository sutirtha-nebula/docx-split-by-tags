"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
var fs = require("fs");
var pizzip_1 = require("pizzip");
var xmldom_1 = require("xmldom");
/* ============================
   Namespaces
============================ */
var NS = {
    w: "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    rel: "http://schemas.openxmlformats.org/package/2006/relationships",
};
function extractFirstDoc(inputFile, outputFile, splitTag) {
    var serializer = new xmldom_1.XMLSerializer();
    var parser = new xmldom_1.DOMParser();
    var srcZip = new pizzip_1.default(fs.readFileSync(inputFile));
    var docXml = parser.parseFromString(srcZip.file("word/document.xml").asText(), "text/xml");
    var body = docXml.getElementsByTagNameNS(NS.w, "body")[0];
    var nodes = Array.from(body.childNodes);
    var splitIndex = findSplitIndexByTagString(nodes, splitTag);
    if (splitIndex === -1) {
        fs.writeFileSync(outputFile, fs.readFileSync(inputFile));
        console.log("ℹ️ No split tag found — copied full doc");
        return true;
    }
    var newBody = docXml.createElementNS(NS.w, "body");
    // if (splitIndex === nodes.length) {
    //   splitIndex = splitIndex - 1;
    // }
    for (var i = 0; i < splitIndex; i++) {
        newBody.appendChild(nodes[i].cloneNode(true));
    }
    // Append section properties if any
    var sectPrs = docXml.getElementsByTagNameNS(NS.w, "sectPr");
    if (sectPrs.length > 0) {
        var lastPara = docXml.createElementNS(NS.w, "w:p");
        var lastParaPr = docXml.createElementNS(NS.w, "w:pPr");
        lastParaPr.appendChild(sectPrs[sectPrs.length - 1].cloneNode(true));
        lastPara.appendChild(lastParaPr);
        newBody.appendChild(lastPara);
    }
    body.parentNode.replaceChild(newBody, body);
    var outZip = cloneZip(srcZip, serializer.serializeToString(docXml));
    fs.writeFileSync(outputFile, outZip.generate({ type: "nodebuffer" }));
    console.log("\u2705 First doc written \u2192 ".concat(outputFile));
    return true;
}
function extractRemainingDoc(inputFile, outputFile, splitTag) {
    var serializer = new xmldom_1.XMLSerializer();
    var parser = new xmldom_1.DOMParser();
    var srcZip = new pizzip_1.default(fs.readFileSync(inputFile));
    var docXml = parser.parseFromString(srcZip.file("word/document.xml").asText(), "text/xml");
    var body = docXml.getElementsByTagNameNS(NS.w, "body")[0];
    var nodes = Array.from(body.childNodes);
    var splitIndex = findSplitIndexByTagString(nodes, splitTag);
    if (splitIndex === -1 || splitIndex === (nodes.length - 1)) {
        fs.writeFileSync(outputFile, fs.readFileSync(inputFile));
        console.log("ℹ️ No split tag found — copied full doc");
        return true;
    }
    var newBody = docXml.createElementNS(NS.w, "body");
    for (var i = splitIndex; i < nodes.length; i++) {
        if (i === splitIndex) {
            // Remove split tag from first node
            var cleanedNode = removeSplitTagByString(nodes[i], splitTag);
            newBody.appendChild(cleanedNode);
        }
        else {
            newBody.appendChild(nodes[i].cloneNode(true));
        }
    }
    body.parentNode.replaceChild(newBody, body);
    // Fix section properties - remove header/footer refs to prevent duplication
    var sectPrs = docXml.getElementsByTagNameNS(NS.w, "sectPr");
    for (var sc = 0; sc < sectPrs.length; sc++) {
        var sectPr = sectPrs[sc];
        // const headerRefs = sectPr.getElementsByTagNameNS(NS.w, "headerReference");
        // const footerRefs = sectPr.getElementsByTagNameNS(NS.w, "footerReference");
        var titlePg = sectPr.getElementsByTagNameNS(NS.w, "titlePg");
        // Remove header/footer references
        // for (let i = headerRefs.length - 1; i >= 0; i--) {
        //   sectPr.removeChild(headerRefs[i]);
        // }
        // for (let i = footerRefs.length - 1; i >= 0; i--) {
        //   sectPr.removeChild(footerRefs[i]);
        // }
        for (var i = titlePg.length - 1; i >= 0; i--) {
            sectPr.removeChild(titlePg[i]);
        }
        // Page numbering starts at 2
        var pgNum = sectPr.getElementsByTagNameNS(NS.w, "pgNumType")[0];
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
    var outZip = cloneZip(srcZip, serializer.serializeToString(docXml));
    fs.writeFileSync(outputFile, outZip.generate({ type: "nodebuffer" }));
    console.log("\u2705 Remaining doc written \u2192 ".concat(outputFile));
    return true;
}
function cloneZip(srcZip, documentXml) {
    var outZip = new pizzip_1.default();
    Object.keys(srcZip.files).forEach(function (name) {
        if (name === "word/document.xml")
            return;
        var file = srcZip.file(name);
        if (!file)
            return;
        if (name.includes("media/")) {
            outZip.file(name, file.asNodeBuffer());
        }
        else {
            outZip.file(name, file.asText());
        }
    });
    outZip.file("word/document.xml", documentXml);
    return outZip;
}
function removeSplitTagByString(node, splitTag) {
    var serializer = new xmldom_1.XMLSerializer();
    var parser = new xmldom_1.DOMParser();
    var nodeXml = serializer.serializeToString(node);
    var cleanedXml = nodeXml.replace(splitTag, "");
    return parser.parseFromString(cleanedXml, "text/xml").documentElement;
}
function findSplitIndexByTagString(nodes, splitTag) {
    var serializer = new xmldom_1.XMLSerializer();
    for (var i = 0; i < nodes.length; i++) {
        var nodeXml = serializer.serializeToString(nodes[i]);
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
var SPLIT_TAG = "<w:pageBreakBefore w:val=\"0\"/>";
// const inputFile = "./2026_01_Precision_AI_UFA_Template.docx";
var inputFile = "./templates/Brochure.docx";
extractFirstDoc(inputFile, "main.docx", SPLIT_TAG);
extractRemainingDoc(inputFile, "content.docx", SPLIT_TAG);
module.exports = {
    extractFirstDoc: extractFirstDoc,
    extractRemainingDoc: extractRemainingDoc,
};
