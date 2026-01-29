"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
var fs = require("fs");
var pizzip_1 = require("pizzip");
var xmldom_1 = require("xmldom");
var zip = new pizzip_1.default(fs.readFileSync("./templates/2026_01_Precision_AI_UFA_Template1.docx"));
var mode = "searchReplace"; // "title" | "heading" | "searchReplace" | "replaceContent"
var options = {
    headingTexts: ["Executive Summary", "Modified"],
    title: "My Document Title",
    prefixText: "",
    suffixText: "",
    color: '49A361',
    fontSize: null,
    fontFamily: "Calibri",
    underline: true,
    bold: true,
    italic: true,
    bgColor: null,
    borderColor: "black",
    borderSize: 0,
    borderStyle: "single",
    searchTexts: [
        "Lorem Ipsum",
    ],
};
function addPrefixSuffix(text, opts) {
    return "".concat(opts.prefixText || "").concat(text).concat(opts.suffixText || "");
}
function normalizeText(text) {
    return text
        .normalize("NFKC")
        .replace(/[\u2018\u2019]/g, "'")
        .replace(/[\u201C\u201D]/g, '"')
        .replace(/\u00A0/g, " ")
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
        if (opts.bold) {
            var b = rPr.getElementsByTagName("w:b")[0];
            if (!b) {
                b = doc.createElement("w:b");
                rPr.appendChild(b);
            }
        }
        if (opts.italic) {
            var i = rPr.getElementsByTagName("w:i")[0];
            if (!i) {
                i = doc.createElement("w:i");
                rPr.appendChild(i);
            }
        }
        if (opts.color) {
            var color = rPr.getElementsByTagName("w:color")[0];
            if (!color) {
                color = doc.createElement("w:color");
                rPr.appendChild(color);
            }
            color.setAttribute("w:val", opts.color);
        }
        if (opts.fontSize) {
            var sz = rPr.getElementsByTagName("w:sz")[0];
            if (!sz) {
                sz = doc.createElement("w:sz");
                rPr.appendChild(sz);
            }
            sz.setAttribute("w:val", opts.fontSize.toString());
        }
        if (opts.fontFamily) {
            var rFonts = rPr.getElementsByTagName("w:rFonts")[0];
            if (!rFonts) {
                rFonts = doc.createElement("w:rFonts");
                rPr.appendChild(rFonts);
            }
            rFonts.setAttribute("w:ascii", opts.fontFamily);
            rFonts.setAttribute("w:hAnsi", opts.fontFamily);
        }
        if (opts.underline) {
            var u = rPr.getElementsByTagName("w:u")[0];
            if (!u) {
                u = doc.createElement("w:u");
                rPr.appendChild(u);
            }
            u.setAttribute("w:val", "single");
        }
        if ("bgColor" in options && options.bgColor !== null) {
            if (options.bgColor) {
                var p = run.parentNode;
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
        }
        if ("borderSize" in options && options.borderSize !== 0) {
            var p = run.parentNode;
            applyBorder(p, doc, options);
        }
    }
}
function applyBorder(p, doc, opts) {
    var pPr = p.getElementsByTagName("w:pPr")[0];
    if (!pPr) {
        pPr = doc.createElement("w:pPr");
        p.insertBefore(pPr, p.firstChild);
    }
    var oldBdr = pPr.getElementsByTagName("w:pBdr")[0];
    if (oldBdr) {
        pPr.removeChild(oldBdr);
    }
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
function getParagraphText(p) {
    var texts = p.getElementsByTagName("w:t");
    var fullText = "";
    for (var i = 0; i < texts.length; i++) {
        fullText += texts[i].textContent || "";
    }
    return fullText;
}
function cloneRunProperties(r) {
    var rPr = r.getElementsByTagName("w:rPr")[0];
    return rPr ? rPr.cloneNode(true) : null;
}
function createRun(doc, text, styled, existingRPr) {
    var run = doc.createElement("w:r");
    if (existingRPr) {
        run.appendChild(existingRPr.cloneNode(true));
    }
    if (styled) {
        applyStyleToRuns([run], doc, options);
    }
    var t = doc.createElement("w:t");
    t.setAttribute("xml:space", "preserve");
    t.textContent = text;
    run.appendChild(t);
    return run;
}
function rebuildParagraphWithMatches(p, doc, fullText, searchText) {
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
        if (matchIndex === -1)
            break;
        var before = fullText.slice(originalIndex, originalIndex + matchIndex);
        var match = fullText.slice(originalIndex + matchIndex, originalIndex + matchIndex + searchText.length);
        if (before) {
            p.appendChild(createRun(doc, before, false, null));
        }
        p.appendChild(createRun(doc, match, true, null));
        originalIndex += matchIndex + searchText.length;
        fullText = fullText.slice(matchIndex + searchText.length);
        normalizedFull = normalizeText(fullText);
    }
    if (fullText) {
        p.appendChild(createRun(doc, fullText, false, null));
    }
}
function isNonContentStyle(styleName) {
    if (!styleName)
        return false;
    var nonContentPrefixes = [
        "Heading",
        "Title",
        "Subtitle",
        "Caption",
        "TOC",
        "Contents",
        "Header",
        "Footer",
        "SourceCode"
    ];
    return nonContentPrefixes.some(function (prefix) { return styleName.startsWith(prefix); });
}
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
        if (size > 32) {
            return true;
        }
    }
    return false;
}
function styleTextInParagraph(p, doc, searchText, options) {
    var runs = Array.from(p.getElementsByTagName("w:r"));
    var runInfo = [];
    var fullText = "";
    runs.forEach(function (r) {
        var t = r.getElementsByTagName("w:t")[0];
        if (!t)
            return;
        var text = t.textContent || "";
        runInfo.push({
            r: r,
            text: text,
            start: fullText.length,
            end: fullText.length + text.length
        });
        fullText += text;
    });
    var matchStart = fullText.indexOf(searchText);
    if (matchStart === -1)
        return;
    var matchEnd = matchStart + searchText.length;
    var affected = runInfo.filter(function (info) {
        return info.end > matchStart && info.start < matchEnd;
    });
    if (affected.length === 0)
        return;
    var parent = affected[0].r.parentNode;
    var anchor = affected[affected.length - 1].r.nextSibling;
    var newRuns = [];
    affected.forEach(function (info) {
        parent.removeChild(info.r);
    });
    affected.forEach(function (info) {
        var beforeLen = Math.max(0, matchStart - info.start);
        var matchLen = Math.min(info.text.length, matchEnd - info.start) - beforeLen;
        var before = info.text.slice(0, beforeLen);
        var match = info.text.slice(beforeLen, beforeLen + matchLen);
        var after = info.text.slice(beforeLen + matchLen);
        var rPr = cloneRunProperties(info.r);
        if (before) {
            parent.insertBefore(createRun(doc, before, false, rPr), anchor);
        }
        if (match) {
            var matchRun = createRun(doc, match, true, rPr);
            parent.insertBefore(matchRun, anchor);
            newRuns.push(matchRun);
        }
        if (after) {
            parent.insertBefore(createRun(doc, after, false, rPr), anchor);
        }
    });
    if (newRuns.length > 0) {
        applyStyleToRuns(newRuns, doc, options);
    }
}
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
function updateTitle() {
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
function updateHeading() {
    var xml = zip.file("word/document.xml").asText();
    var doc = new xmldom_1.DOMParser().parseFromString(xml, "text/xml");
    var paragraphs = doc.getElementsByTagName("w:p");
    var headingTexts = (options.headingTexts || []).map(function (t) {
        return normalizeText(t);
    });
    for (var i = 0; i < paragraphs.length; i++) {
        var p = paragraphs[i];
        if (!isHeadingStyle(p))
            continue;
        var headingText = getParagraphText(p);
        var normalizedHeading = normalizeText(headingText);
        if (headingTexts.length > 0 && !headingTexts.includes(normalizedHeading)) {
            continue;
        }
        var runs = Array.from(p.getElementsByTagName("w:r"));
        // Fix: Type the textRuns array properly
        var textRuns = runs.filter(function (r) {
            var t = r.getElementsByTagName("w:t")[0];
            return t && t.textContent;
        });
        if (textRuns.length > 0) {
            var prefix = options.prefixText || "";
            var suffix = options.suffixText || "";
            // Fix: Cast to any to avoid TypeScript errors
            var firstT = textRuns[0].getElementsByTagName("w:t")[0];
            var lastT = textRuns[textRuns.length - 1].getElementsByTagName("w:t")[0];
            var fullText = getParagraphText(p);
            if (prefix && !fullText.startsWith(prefix)) {
                firstT.textContent = prefix + firstT.textContent;
                console.log("Added prefix to heading:", prefix, firstT.textContent);
            }
            if (suffix && !fullText.endsWith(suffix)) {
                lastT.textContent = lastT.textContent + suffix;
                console.log("Added suffix to heading:", suffix, lastT.textContent);
            }
        }
        var runsToStyle = Array.from(p.getElementsByTagName("w:r"));
        applyStyleToRuns(runsToStyle, doc, options);
    }
    zip.file("word/document.xml", new xmldom_1.XMLSerializer().serializeToString(doc));
    console.log("Headings updated!");
}
function searchUpdate() {
    if (!options.searchTexts || options.searchTexts.length === 0) {
        throw new Error("searchTexts[] is required for searchReplace mode");
    }
    var xml = zip.file("word/document.xml").asText();
    var doc = new xmldom_1.DOMParser().parseFromString(xml, "text/xml");
    var paragraphs = Array.from(doc.getElementsByTagName("w:p"));
    var normalizedSearchTexts = options.searchTexts.map(normalizeText);
    for (var i = 0; i < paragraphs.length; i++) {
        var p = paragraphs[i];
        var normalizedParaText = normalizeText(getParagraphText(p));
        for (var s = 0; s < normalizedSearchTexts.length; s++) {
            var searchText = options.searchTexts[s];
            var normalizedSearch = normalizedSearchTexts[s];
            if (!normalizedParaText.includes(normalizedSearch))
                continue;
            styleTextInParagraph(p, doc, searchText, options);
            if ("borderSize" in options && options.borderSize !== 0) {
                applyBorder(p, doc, options);
            }
            normalizedParaText = normalizeText(getParagraphText(p));
        }
    }
    zip.file("word/document.xml", new xmldom_1.XMLSerializer().serializeToString(doc));
}
function replaceUpdate() {
    var xml = zip.file("word/document.xml").asText();
    var doc = new xmldom_1.DOMParser().parseFromString(xml, "text/xml");
    var paragraphs = doc.getElementsByTagName("w:p");
    for (var i = 0; i < paragraphs.length; i++) {
        var p = paragraphs[i];
        var runs = p.getElementsByTagName("w:r");
        if (isNonContentParagraph(p)) {
            continue;
        }
        for (var r = 0; r < runs.length; r++) {
            var t = runs[r].getElementsByTagName("w:t")[0];
            if (t && t.textContent) {
                t.textContent = addPrefixSuffix(t.textContent, options);
            }
        }
        var runsArray = Array.from(runs);
        applyStyleToRuns(runsArray, doc, options);
    }
    zip.file("word/document.xml", new xmldom_1.XMLSerializer().serializeToString(doc));
    console.log("Content replaced!");
}
// ---------------- TITLE ----------------
if (mode === "title") {
    updateTitle();
}
// ---------------- HEADING ----------------
else if (mode === "heading") {
    updateHeading();
}
else if (mode === "searchReplace") {
    searchUpdate();
}
// ---------------- REPLACE CONTENT ----------------
else if (mode === "replaceContent") {
    replaceUpdate();
}
fs.writeFileSync("output.docx", zip.generate({ type: "nodebuffer" }));
console.log("File saved!");
