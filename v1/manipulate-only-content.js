"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.colorContentParagraphs = colorContentParagraphs;
var fs = require("fs");
var pizzip_1 = require("pizzip");
var xmldom_1 = require("xmldom");
/**
 * Checks if paragraph style name means non-content
 */
function isNonContentStyle(styleName) {
    if (!styleName)
        return false;
    var nonContentPrefixes = [
        "Heading",
        "Title",
        "Subtitle",
        "Caption",
        // "TOC",
        'Table of Contents',
        "Header",
        "Footer",
        "SourceCode"
    ];
    return nonContentPrefixes.some(function (prefix) { return styleName.startsWith(prefix); });
}
/**
 * Checks if paragraph has center alignment
 */
function isCenteredParagraph(p) {
    var pPr = p.getElementsByTagName("w:pPr")[0];
    if (!pPr)
        return false;
    var jc = pPr.getElementsByTagName("w:jc")[0];
    if (!jc)
        return false;
    var val = jc.getAttribute("w:val");
    return val === "center";
}
/**
 * Checks if paragraph has any run with font size > 16pt (w:sz is half-points, so > 32)
 */
function hasLargeFontSize(p) {
    var runs = p.getElementsByTagName("w:r");
    for (var i = 0; i < runs.length; i++) {
        var run = runs[i];
        var rPr = run.getElementsByTagName("w:rPr")[0];
        if (!rPr)
            continue;
        var sz = rPr.getElementsByTagName("w:sz")[0];
        if (!sz)
            continue;
        var val = sz.getAttribute("w:val");
        if (!val)
            continue;
        var size = parseInt(val, 10);
        if (size > 32) { // 16pt * 2 (half points)
            return true;
        }
    }
    return false;
}
/**
 * Centralized check: Is paragraph non-content?
 */
function isNonContentParagraph(p) {
    var pPr = p.getElementsByTagName("w:pPr")[0];
    var styleName = null;
    if (pPr) {
        var pStyle = pPr.getElementsByTagName("w:pStyle")[0];
        if (pStyle) {
            styleName = pStyle.getAttribute("w:val");
        }
    }
    return isNonContentStyle(styleName) || isCenteredParagraph(p) || hasLargeFontSize(p);
}
/**
 * Main function to color content paragraphs
 */
function colorContentParagraphs(inputFilePath, outputFilePath, colorHex // default dark blue
) {
    var _a;
    if (colorHex === void 0) { colorHex = "00008B"; }
    var buffer = fs.readFileSync(inputFilePath);
    var zip = new pizzip_1.default(buffer);
    var xml = (_a = zip.file("word/document.xml")) === null || _a === void 0 ? void 0 : _a.asText();
    if (!xml) {
        throw new Error("document.xml not found in the DOCX file");
    }
    var doc = new xmldom_1.DOMParser().parseFromString(xml, "text/xml");
    var paragraphs = doc.getElementsByTagName("w:p");
    for (var i = 0; i < paragraphs.length; i++) {
        var p = paragraphs[i];
        if (isNonContentParagraph(p)) {
            // Skip this paragraph (heading, title, centered, or large font)
            continue;
        }
        // Color runs in this paragraph
        var runs = p.getElementsByTagName("w:r");
        for (var r = 0; r < runs.length; r++) {
            var run = runs[r];
            var rPr = run.getElementsByTagName("w:rPr")[0];
            if (!rPr) {
                rPr = doc.createElement("w:rPr");
                run.insertBefore(rPr, run.firstChild);
            }
            var color = rPr.getElementsByTagName("w:color")[0];
            if (!color) {
                color = doc.createElement("w:color");
                rPr.appendChild(color);
            }
            color.setAttribute("w:val", colorHex);
        }
    }
    zip.file("word/document.xml", new xmldom_1.XMLSerializer().serializeToString(doc));
    fs.writeFileSync(outputFilePath, zip.generate({ type: "nodebuffer" }));
    console.log("Content paragraph colors updated successfully and saved to ".concat(outputFilePath));
}
colorContentParagraphs("./templates/2026_01_Precision_AI_UFA_Template.docx", "output.docx", "00008B");
