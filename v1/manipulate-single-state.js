"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
var fs = require("fs");
var pizzip_1 = require("pizzip");
var xmldom_1 = require("xmldom");
/* ============================
   Helpers
============================ */
function getOrCreate(parent, tag, doc) {
    var el = parent.getElementsByTagName(tag)[0];
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
// 1️⃣ Font size (pt → half-points)
function changeFontSize(run, sizePt, doc) {
    var rPr = getOrCreateRPr(run, doc);
    var sz = rPr.getElementsByTagName("w:sz")[0];
    if (!sz) {
        sz = doc.createElement("w:sz");
        rPr.appendChild(sz);
    }
    sz.setAttribute("w:val", String(sizePt * 2));
}
// 2️⃣ Font family
function changeFontFamily(run, fontName, doc) {
    var rPr = getOrCreateRPr(run, doc);
    var rFonts = rPr.getElementsByTagName("w:rFonts")[0];
    if (!rFonts) {
        rFonts = doc.createElement("w:rFonts");
        rPr.appendChild(rFonts);
    }
    rFonts.setAttribute("w:ascii", fontName);
    rFonts.setAttribute("w:hAnsi", fontName);
}
// 3️⃣ Font color (HEX without #)
function changeFontColor(run, hexColor, doc) {
    var rPr = getOrCreateRPr(run, doc);
    var color = rPr.getElementsByTagName("w:color")[0];
    if (!color) {
        color = doc.createElement("w:color");
        rPr.appendChild(color);
    }
    color.setAttribute("w:val", hexColor);
}
/* ============================
   Apply formatting to text inside ONE paragraph
============================ */
function applyFormattingToText(paragraph, searchText, doc, formatter) {
    var _a;
    var runs = Array.from(paragraph.getElementsByTagName("w:r"));
    var fullText = "";
    var runMap = [];
    // Build paragraph text + run position map
    for (var _i = 0, runs_1 = runs; _i < runs_1.length; _i++) {
        var run = runs_1[_i];
        var t = run.getElementsByTagName("w:t")[0];
        if (!t)
            continue;
        var text = (_a = t.textContent) !== null && _a !== void 0 ? _a : "";
        var start = fullText.length;
        fullText += text;
        var end = fullText.length;
        runMap.push({ run: run, start: start, end: end });
    }
    // 🔁 Find and process ALL matches
    var searchFrom = 0;
    while (true) {
        var matchIndex = fullText.indexOf(searchText, searchFrom);
        if (matchIndex === -1)
            break;
        var matchStart = matchIndex;
        var matchEnd = matchIndex + searchText.length;
        // Apply formatter only to overlapping runs
        for (var _b = 0, runMap_1 = runMap; _b < runMap_1.length; _b++) {
            var _c = runMap_1[_b], run = _c.run, start = _c.start, end = _c.end;
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
function formatTextInDocument(doc, targetText, formatter) {
    var paragraphs = Array.from(doc.getElementsByTagName("w:p"));
    for (var _i = 0, paragraphs_1 = paragraphs; _i < paragraphs_1.length; _i++) {
        var paragraph = paragraphs_1[_i];
        applyFormattingToText(paragraph, targetText, doc, formatter);
    }
}
/* ============================
   Main
============================ */
var zip = new pizzip_1.default(fs.readFileSync("./templates/2026_01_Precision_AI_UFA_Template.docx"));
var xml = zip.file("word/document.xml").asText();
var doc = new xmldom_1.DOMParser().parseFromString(xml, "text/xml");
formatTextInDocument(doc, "Lorem Ipsum", function (run) {
    changeFontFamily(run, "DejaVu Sans", doc);
    changeFontSize(run, 12, doc);
    changeFontColor(run, "FFC0CB", doc);
});
/* ============================
   Save
============================ */
zip.file("word/document.xml", new xmldom_1.XMLSerializer().serializeToString(doc));
fs.writeFileSync("output.docx", zip.generate({ type: "nodebuffer" }));
console.log("✅ Correct runs formatted");
