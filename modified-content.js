"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
var fs = require("fs");
var pizzip_1 = require("pizzip");
var xmldom_1 = require("xmldom");
var zip = new pizzip_1.default(fs.readFileSync("./templates/final_result_main.docx"));
var mode = "replaceContent";

var options = {
    title: "My Document Title",
    prefixText: "[PRE] ",
    suffixText: " (POST)",
    color: "808080",
    fontSize: 16,
    fontFamily: "Calibri",
    underline: true,
    bold: false,
    italic: true,
    bgColor: "D9EAD3",
    borderColor: "FF0000",
    borderSize: 0,
    borderStyle: "single",
    searchText: "Lorem Ipsum is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, when an unknown printer took a galley of type and scrambled it to make a type specimen book. It has survived not only five centuries, but also the leap into electronic typesetting, remaining essentially unchanged. It was popularised in the 1960s with the release of Letraset sheets containing Lorem Ipsum passages, and more recently with desktop publishing software like Aldus PageMaker including versions of Lorem Ipsum.",
};
function addPrefixSuffix(text, opts) {
    return "".concat(opts.prefixText || "").concat(text).concat(opts.suffixText || "");
}
function normalizeText(text) {
    return text
        // normalize unicode
        .normalize("NFKC")

        // convert curly quotes → straight quotes
        .replace(/[\u2018\u2019]/g, "'")
        .replace(/[\u201C\u201D]/g, '"')

        // convert non-breaking spaces → normal spaces
        .replace(/\u00A0/g, " ")

        // collapse whitespace
        .replace(/\s+/g, " ")
        .trim();
}

function applyStyleToRuns(runs, doc, opts) {
    for (var r = 0; r < runs.length; r++) {
        var run = runs[r];
        var rPr = run.getElementsByTagName("w:rPr")[0];
        if (!rPr) {
            rPr = doc.createElement("w:rPr");
            run.insertBefore(rPr, run.firstChild);
        }
        // Bold
        if (opts.bold) {
            var b = rPr.getElementsByTagName("w:b")[0];
            if (!b) {
                b = doc.createElement("w:b");
                rPr.appendChild(b);
            }
        }
        // Italic
if (opts.italic) {
    var i = rPr.getElementsByTagName("w:i")[0];
    if (!i) {
        i = doc.createElement("w:i");
        rPr.appendChild(i);
    }
}

        // Color
        if (opts.color) {
            var color = rPr.getElementsByTagName("w:color")[0];
            if (!color) {
                color = doc.createElement("w:color");
                rPr.appendChild(color);
            }
            color.setAttribute("w:val", opts.color);
        }
        // Font size
        if (opts.fontSize) {
            var sz = rPr.getElementsByTagName("w:sz")[0];
            if (!sz) {
                sz = doc.createElement("w:sz");
                rPr.appendChild(sz);
            }
            sz.setAttribute("w:val", opts.fontSize.toString());
        }
        // Font family
        if (opts.fontFamily) {
            var rFonts = rPr.getElementsByTagName("w:rFonts")[0];
            if (!rFonts) {
                rFonts = doc.createElement("w:rFonts");
                rPr.appendChild(rFonts);
            }
            rFonts.setAttribute("w:ascii", opts.fontFamily);
            rFonts.setAttribute("w:hAnsi", opts.fontFamily);
        }
        // Underline
        if (opts.underline) {
            var u = rPr.getElementsByTagName("w:u")[0];
            if (!u) {
                u = doc.createElement("w:u");
                rPr.appendChild(u);
            }
            u.setAttribute("w:val", "single");
        }
    }
}
function applyBorder(p, doc, opts) {
    var pPr = p.getElementsByTagName("w:pPr")[0];
    if (!pPr) {
        pPr = doc.createElement("w:pPr");
        p.insertBefore(pPr, p.firstChild);
    }
    // Remove existing border if exists
    var oldBdr = pPr.getElementsByTagName("w:pBdr")[0];
    if (oldBdr)
        pPr.removeChild(oldBdr);
    var pBdr = doc.createElement("w:pBdr");
    pPr.appendChild(pBdr);
    var sides = ["top", "left", "bottom", "right"];
    sides.forEach(function (side) {
        var node = doc.createElement("w:".concat(side));
        node.setAttribute("w:val", opts.borderStyle || "single");
        node.setAttribute("w:sz", (opts.borderSize || 8).toString());
        node.setAttribute("w:space", "1");
        node.setAttribute("w:color", opts.borderColor || "000000");
        pBdr.appendChild(node);
    });
}
function isHeadingStyle(p) {
    var pPr = p.getElementsByTagName("w:pPr")[0];
    if (!pPr)
        return false;
    var pStyle = pPr.getElementsByTagName("w:pStyle")[0];
    if (!pStyle)
        return false;
    var styleVal = pStyle.getAttribute("w:val");
    return !!styleVal && styleVal.startsWith("Heading");
}
/* -------- SEARCH / REPLACE HELPERS -------- */
function getParagraphText(p) {
    var texts = p.getElementsByTagName("w:t");
    var fullText = "";
    for (var i = 0; i < texts.length; i++) {
        fullText += texts[i].textContent || "";
    }
    return fullText;
}
function createRun(doc, text, styled) {
    var run = doc.createElement("w:r");
    if (styled) {
        var rPr = doc.createElement("w:rPr");
         rPr.appendChild(doc.createElement("w:b"));
        run.appendChild(rPr);
        applyStyleToRuns([run], doc, options);
    }
    var t = doc.createElement("w:t");
      t.setAttribute("xml:space", "preserve");
    t.textContent = text;
    run.appendChild(t);
    return run;
}
function rebuildParagraphWithMatches(p, doc, fullText, searchText) {
    // Remove everything except paragraph properties
    var children = Array.from(p.childNodes);
    children.forEach(function (node) {
        if (node.nodeName !== "w:pPr") {
            p.removeChild(node);
        }
    });

    var normalizedFull = normalizeText(fullText);
    var normalizedSearch = normalizeText(searchText);

    var originalIndex = 0;

    while (true) {
        var matchIndex = normalizedFull.indexOf(normalizedSearch);
        if (matchIndex === -1) break;

        // Find corresponding slice in ORIGINAL text
        var before = fullText.slice(originalIndex, originalIndex + matchIndex);
        var match = fullText.slice(
            originalIndex + matchIndex,
            originalIndex + matchIndex + searchText.length
        );

        if (before) {
            p.appendChild(createRun(doc, before, false));
        }

        p.appendChild(createRun(doc, match, true));

        // Advance both strings
        originalIndex += matchIndex + searchText.length;
        fullText = fullText.slice(matchIndex + searchText.length);
        normalizedFull = normalizeText(fullText);
    }

    if (fullText) {
        p.appendChild(createRun(doc, fullText, false));
    }
}

/* -------- REPLACE CONTENT HELPERS -------- */
function isNonContentStyle(styleName) {
    if (!styleName)
        return false;
    var nonContentPrefixes = [
        "Heading",
        "Title",
        "Subtitle",
        "Caption",
        "TOC",      // Table of Contents
        "Contents", // Table of Contents
        "Header",
        "Footer",
        "SourceCode"
    ];
    return nonContentPrefixes.some(function (prefix) { return styleName.startsWith(prefix); });
}
/*** Checks if paragraph has center alignment ***/
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
/*** Checks if paragraph has any run with font size > 16pt (w:sz is half-points, so > 32) */
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
/*** Centralized check: Is paragraph non-content? */
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
// ---------------- TITLE ----------------
if (mode === "title") {
    var coreXml = zip.file("docProps/core.xml").asText();
    var coreDoc = new xmldom_1.DOMParser().parseFromString(coreXml, "text/xml");
    var titleNode = coreDoc.getElementsByTagName("dc:title")[0];
    var newTitleText = addPrefixSuffix(options.title || "", options);
    if (titleNode) {
        titleNode.textContent = newTitleText;
    }
    else {
        var root = coreDoc.documentElement;
        var newTitle = coreDoc.createElement("dc:title");
        newTitle.textContent = newTitleText;
        root.appendChild(newTitle);
    }
    zip.file("docProps/core.xml", new xmldom_1.XMLSerializer().serializeToString(coreDoc));
    console.log("Title updated!");
}
// ---------------- HEADING ----------------
else if (mode === "heading") {
    var xml = zip.file("word/document.xml").asText();
    var doc = new xmldom_1.DOMParser().parseFromString(xml, "text/xml");
    var paragraphs = doc.getElementsByTagName("w:p");
    for (var i = 0; i < paragraphs.length; i++) {
        var p = paragraphs[i];
        if (!isHeadingStyle(p))
            continue;
        var runs = p.getElementsByTagName("w:r");
        for (var r = 0; r < runs.length; r++) {
            var t = runs[r].getElementsByTagName("w:t")[0];
            if (t && t.textContent) {
                t.textContent = addPrefixSuffix(t.textContent, options);
            }
        }
        applyStyleToRuns(runs, doc, options);
        // shading background
        if (options.bgColor) {
            var pPr = p.getElementsByTagName("w:pPr")[0];
            if (!pPr) {
                pPr = doc.createElement("w:pPr");
                p.insertBefore(pPr, p.firstChild);
            }
            var shd = pPr.getElementsByTagName("w:shd")[0];
            if (!shd) {
                shd = doc.createElement("w:shd");
                pPr.appendChild(shd);
            }
            shd.setAttribute("w:fill", options.bgColor);
        }
        // border
        applyBorder(p, doc, options);
    }
    zip.file("word/document.xml", new xmldom_1.XMLSerializer().serializeToString(doc));
    console.log("Headings updated!");
}
// ---------------- SEARCH / REPLACE ----------------
else if (mode === "searchReplace") {
    if (!options.searchText) {
    throw new Error("searchText is required for searchReplace mode");
}

var xml = zip.file("word/document.xml").asText();
var doc = new xmldom_1.DOMParser().parseFromString(xml, "text/xml");

// IMPORTANT: make it a static array
var paragraphs = Array.from(doc.getElementsByTagName("w:p"));

var normalizedSearch = normalizeText(options.searchText);

for (var i = 0; i < paragraphs.length; i++) {
    var p = paragraphs[i];
  var originalText = getParagraphText(p);
    var normalizedText = normalizeText(originalText);

    if (!normalizedText.includes(normalizedSearch)) continue;
    // Rebuild paragraph with highlighted matches

    rebuildParagraphWithMatches(p, doc, originalText, options.searchText);

    // border
    if ("borderSize" in options && options.borderSize !== 0) {
        applyBorder(p, doc, options);
    }
}

zip.file(
    "word/document.xml",
    new xmldom_1.XMLSerializer().serializeToString(doc)
);
}
// ---------------- REPLACE CONTENT ----------------
else if (mode === "replaceContent") {
    var xml = zip.file("word/document.xml").asText();
    var doc = new xmldom_1.DOMParser().parseFromString(xml, "text/xml");
    var paragraphs = doc.getElementsByTagName("w:p");
    for (var i = 0; i < paragraphs.length; i++) {
        var p = paragraphs[i];
        var runs = p.getElementsByTagName("w:r");
        if (isNonContentParagraph(p)) {
            // Skip this paragraph (heading, title, centered, or large font)
            continue;
        }
        for (var r = 0; r < runs.length; r++) {
            var t = runs[r].getElementsByTagName("w:t")[0];
            if (t && t.textContent) {
                t.textContent = addPrefixSuffix(t.textContent, options);
            }
        }
        applyStyleToRuns(runs, doc, options);
    }
    zip.file("word/document.xml", new xmldom_1.XMLSerializer().serializeToString(doc));
    console.log("Content replaced!");
}
fs.writeFileSync("output.docx", zip.generate({ type: "nodebuffer" }));
console.log("File saved!");