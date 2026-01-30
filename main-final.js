"use strict";
var __assign = (this && this.__assign) || function () {
    __assign = Object.assign || function(t) {
        for (var s, i = 1, n = arguments.length; i < n; i++) {
            s = arguments[i];
            for (var p in s) if (Object.prototype.hasOwnProperty.call(s, p))
                t[p] = s[p];
        }
        return t;
    };
    return __assign.apply(this, arguments);
};
Object.defineProperty(exports, "__esModule", { value: true });
var fs = require("fs");
var pizzip_1 = require("pizzip");
var xmldom_1 = require("xmldom");
var mode = "MAIN"; // "TITLE" | "HEADING" | "CONTENT" | "COMPLETE" | "COMPLETE_CONTENT" | "MAIN"
var options = {
    headingTexts: ["Executive \"Summary\"", "Modified"],
    // prefixText: "",
    // suffixText: "",
    // color: '49A361',
    // fontSize: 22,
    // fontFamily: "Calibri",
    underline: true,
    // bold: true,
    // italic: true,
    // bgColor: "FF0000",
    // borderColor: "black",
    borderSize: 6,
    // borderStyle: "single",
    // borderTop: true,
    // borderRight: true,
    // borderBottom: true,
    // // borderLeft: false,
    // highlightColor: "transparent",
    // alignment: 'center',
    // lineHeight: 4,
    contentTexts: [
        "Date: January 2026"
    ]
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
function getOrCreatePPr(p, doc) {
    for (var i = 0; i < p.childNodes.length; i++) {
        var n = p.childNodes[i];
        if (n.nodeName === "w:pPr")
            return n;
    }
    var pPr = doc.createElement("w:pPr");
    p.insertBefore(pPr, p.firstChild);
    return pPr;
}
function removeChildren(pPr, tag) {
    var nodes = Array.from(pPr.getElementsByTagName(tag));
    nodes.forEach(function (n) { return pPr.removeChild(n); });
}
function applyStyleToRuns(runs, doc, opts) {
    for (var r = 0; r < runs.length; r++) {
        var run = runs[r];
        var rPr = run.getElementsByTagName("w:rPr")[0];
        if (!rPr) {
            rPr = doc.createElement("w:rPr");
            run.insertBefore(rPr, run.firstChild);
        }
        if (opts.bold !== null && opts.bold !== undefined) {
            var b = rPr.getElementsByTagName("w:b")[0];
            if (opts.bold === true) {
                if (!b) {
                    b = doc.createElement("w:b");
                    rPr.appendChild(b);
                }
            }
            else {
                if (b) {
                    rPr.removeChild(b);
                }
            }
        }
        if (opts.italic !== null && opts.italic !== undefined) {
            var i = rPr.getElementsByTagName("w:i")[0];
            if (opts.italic === true) {
                if (!i) {
                    i = doc.createElement("w:i");
                    rPr.appendChild(i);
                }
            }
            else {
                if (i) {
                    rPr.removeChild(i);
                }
            }
        }
        if (opts.underline !== null && opts.underline !== undefined) {
            var u = rPr.getElementsByTagName("w:u")[0];
            if (opts.underline === true) {
                if (!u) {
                    u = doc.createElement("w:u");
                    rPr.appendChild(u);
                    u.setAttribute("w:val", "single");
                }
            }
            else {
                rPr.removeChild(u);
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
        if (opts.highlightColor) {
            var shd = rPr.getElementsByTagName("w:shd")[0];
            if (!shd) {
                shd = doc.createElement("w:shd");
                rPr.appendChild(shd);
            }
            if (opts.highlightColor === "transparent") {
                shd.setAttribute("w:val", "clear");
                shd.setAttribute("w:color", "auto");
                shd.setAttribute("w:fill", "auto");
                rPr.appendChild(shd);
            }
            else {
                shd.setAttribute("w:fill", opts.highlightColor);
                rPr.appendChild(shd);
            }
        }
        if ("bgColor" in opts && opts.bgColor !== null) {
            if (opts.bgColor) {
                try {
                    var p = run.parentNode;
                    var pPr = void 0;
                    //console.log(p);
                    try {
                        let pPro = p.getElementsByTagName("w:pPr")[0];
                        pPr=pPro.getElementsByTagName("w:p")[0];
                        console.log("p found in paragraph");
                    }
                    catch (e) {
                        console.log("No pPr in paragraph");
                        pPr = p.getElementsByTagName("w:hyperlink")[0];
                        console.log("pPr found in hyperlink");
                    }
                    if (!pPr) {
                        pPr = doc.createElement("w:pPr");
                        p.insertBefore(pPr, p.firstChild);
                    }
                    var shd = pPr.getElementsByTagName("w:shd")[0];
                    if (!shd) {
                        shd = doc.createElement("w:shd");
                        pPr.appendChild(shd);
                    }
                    shd.setAttribute("w:fill", opts.bgColor);
                }
                catch (e) {
                    var shd = rPr.getElementsByTagName("w:shd")[0];
                    if (!shd) {
                        shd = doc.createElement("w:shd");
                        rPr.appendChild(shd);
                    }
                    if (opts.highlightColor === "transparent") {
                        shd.setAttribute("w:val", "clear");
                        shd.setAttribute("w:color", "auto");
                        shd.setAttribute("w:fill", "auto");
                        rPr.appendChild(shd);
                    }
                    else {
                        shd.setAttribute("w:fill", opts.highlightColor);
                        rPr.appendChild(shd);
                    }
                }
            }
        }
        if ("borderSize" in opts && opts.borderSize !== 0) {
            var _a = resolveBorderSides(opts), top_1 = _a.top, right = _a.right, bottom = _a.bottom, left = _a.left;
            var borderOpts = __assign(__assign({}, opts), { borderTop: top_1, borderRight: right, borderBottom: bottom, borderLeft: left });
            var p = run.parentNode;
            try {
                console.log("Applying border to paragraph catch", p);
                applyBorder(p, doc, borderOpts);
            }
            catch (e) {
                console.log("Applying border to paragraph catch", p);
                applyBorder(rPr, doc, borderOpts);
            }
        }
        if ((opts === null || opts === void 0 ? void 0 : opts.alignment) != null) {
            var p = run.parentNode;
            if (!p)
                return;
            var pPr = getOrCreatePPr(p, doc);
            // Remove existing jc (important!)
            removeChildren(pPr, "w:jc");
            var jc = doc.createElement("w:jc");
            var map = {
                left: "left",
                center: "center",
                right: "right",
                justify: "both"
            };
            jc.setAttribute("w:val", map[opts.alignment] || "left");
            pPr.appendChild(jc);
        }
        if ("lineHeight" in opts && opts.lineHeight != null) {
            try {
                var p = run.parentNode;
                var pPr = p.getElementsByTagName("w:pPr")[0];
                if (!pPr) {
                    pPr = doc.createElement("w:pPr");
                    p.insertBefore(pPr, p.firstChild);
                }
                var spacing = pPr.getElementsByTagName("w:spacing")[0];
                if (!spacing) {
                    spacing = doc.createElement("w:spacing");
                    pPr.appendChild(spacing);
                }
                // opts.lineHeight can be like 1, 1.5, 2
                var lineTwips = Math.round(240 * opts.lineHeight);
                spacing.setAttribute("w:line", String(lineTwips));
                spacing.setAttribute("w:lineRule", "auto");
            }
            catch (e) {
                // optional fallback – usually paragraph spacing never belongs in rPr
            }
        }
    }
}
function resolveBorderSides(opts) {
    var hasAnySide = opts.borderTop !== undefined ||
        opts.borderRight !== undefined ||
        opts.borderBottom !== undefined ||
        opts.borderLeft !== undefined;
    // Default values
    var top = false;
    var right = false;
    var bottom = false;
    var left = false;
    if (opts.borderSize && opts.borderSize !== 0) {
        if (!hasAnySide) {
            // borderSize only → all sides
            top = right = bottom = left = true;
        }
        else {
            // respect provided sides
            top = !!opts.borderTop;
            right = !!opts.borderRight;
            bottom = !!opts.borderBottom;
            left = !!opts.borderLeft;
        }
    }
    return { top: top, right: right, bottom: bottom, left: left };
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
    // If no borderSize or zero → do nothing
    if (!opts.borderSize || opts.borderSize === 0) {
        return;
    }
    var pBdr = doc.createElement("w:pBdr");
    pPr.appendChild(pBdr);
    var borderConfig = [
        { side: "top", enabled: opts.borderTop },
        { side: "left", enabled: opts.borderLeft },
        { side: "bottom", enabled: opts.borderBottom },
        { side: "right", enabled: opts.borderRight }
    ];
    console.log("Applying borders:", borderConfig);
    borderConfig.forEach(function (_a) {
        var side = _a.side, enabled = _a.enabled;
        if (!enabled)
            return;
        var node = doc.createElement("w:".concat(side));
        node.setAttribute("w:val", opts.borderStyle || "single");
        node.setAttribute("w:sz", String(opts.borderSize || 8));
        node.setAttribute("w:space", "1");
        node.setAttribute("w:color", opts.borderColor || "000000");
        pBdr.appendChild(node);
    });
    // If no sides were added, remove empty pBdr
    if (!pBdr.hasChildNodes()) {
        pPr.removeChild(pBdr);
    }
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
function createRun(doc, text, styled, existingRPr, options) {
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
// function rebuildParagraphWithMatches(p: any, doc: any, fullText: string, searchText: string): void {
//     const children = Array.from(p.childNodes);
//     children.forEach((node: any) => {
//         if (node.nodeName !== "w:pPr") {
//             p.removeChild(node);
//         }
//     });
//     let normalizedFull = normalizeText(fullText);
//     const normalizedSearch = normalizeText(searchText);
//     let originalIndex = 0;
//     while (true) {
//         const matchIndex = normalizedFull.indexOf(normalizedSearch);
//         if (matchIndex === -1) break;
//         const before = fullText.slice(originalIndex, originalIndex + matchIndex);
//         const match = fullText.slice(originalIndex + matchIndex, originalIndex + matchIndex + searchText.length);
//         if (before) {
//             p.appendChild(createRun(doc, before, false, null));
//         }
//         p.appendChild(createRun(doc, match, true, null));
//         originalIndex += matchIndex + searchText.length;
//         fullText = fullText.slice(matchIndex + searchText.length);
//         normalizedFull = normalizeText(fullText);
//     }
//     if (fullText) {
//         p.appendChild(createRun(doc, fullText, false, null));
//     }
// }
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
    function isHeadingParagraph(p) {
        var pPr = p.getElementsByTagName("w:pPr")[0];
        if (!pPr)
            return false;
        var pStyle = pPr.getElementsByTagName("w:pStyle")[0];
        if (!pStyle)
            return false;
        var val = pStyle.getAttribute("w:val") || "";
        return /^Heading\d+$/.test(val) || val === "Heading";
    }
    // ---- MODE CHECK (ONLY ADDITION) ----
    if (mode === "CONTENT" && isHeadingParagraph(p)) {
        return;
    }
    function buildRunInfo() {
        var runs = Array.from(p.getElementsByTagName("w:r"));
        var runInfo = [];
        var fullText = "";
        runs.forEach(function (r) {
            var t = r.getElementsByTagName("w:t")[0];
            if (!t)
                return;
            var text = t.textContent || "";
            if (!text)
                return;
            runInfo.push({
                r: r,
                text: text,
                start: fullText.length,
                end: fullText.length + text.length
            });
            fullText += text;
        });
        return { runInfo: runInfo, fullText: fullText };
    }
    var data = buildRunInfo();
    var matches = [];
    var idx = 0;
    while ((idx = data.fullText.indexOf(searchText, idx)) !== -1) {
        matches.push({
            start: idx,
            end: idx + searchText.length
        });
        idx += searchText.length;
    }
    if (matches.length === 0)
        return;
    matches.reverse().forEach(function (m) {
        var rebuilt = buildRunInfo();
        var runInfo = rebuilt.runInfo;
        var matchStart = m.start;
        var matchEnd = m.end;
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
                parent.insertBefore(createRun(doc, before, false, rPr, options), anchor);
            }
            if (match) {
                var matchRun = createRun(doc, match, true, rPr, options);
                parent.insertBefore(matchRun, anchor);
                newRuns.push(matchRun);
            }
            if (after) {
                parent.insertBefore(createRun(doc, after, false, rPr, options), anchor);
            }
        });
        if (newRuns.length > 0) {
            applyStyleToRuns(newRuns, doc, options);
        }
    });
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
// function updateTitle(): void {
//     const coreXml = zip.file("docProps/core.xml")!.asText();
//     const coreDoc = new DOMParser().parseFromString(coreXml, "text/xml");
//     const titleNode = coreDoc.getElementsByTagName("dc:title")[0];
//     //const newTitleText = addPrefixSuffix(options.title || "", options);
//     if (titleNode) {
//         titleNode.textContent = newTitleText;
//     } else {
//         const root = coreDoc.documentElement;
//         const newTitle = coreDoc.createElement("dc:title");
//         newTitle.textContent = newTitleText;
//         root.appendChild(newTitle);
//     }
//     zip.file("docProps/core.xml", new XMLSerializer().serializeToString(coreDoc));
//     console.log("Title updated!");
// }
function updateHeading(inputPath, options) {
    var zip = new pizzip_1.default(fs.readFileSync(inputPath));
    var xml = zip.file("word/document.xml").asText();
    xml = xml.replace(/\u2019/g, "'");
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
        var textRuns = runs.filter(function (r) {
            var t = r.getElementsByTagName("w:t")[0];
            return t && t.textContent;
        });
        if (textRuns.length > 0) {
            var prefix = options.prefixText || "";
            var suffix = options.suffixText || "";
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
    fs.writeFileSync(inputPath, zip.generate({ type: "nodebuffer" }));
    console.log("File saved!");
    console.log("Headings updated!");
}
function searchUpdate(inputPath, options) {
    var zip = new pizzip_1.default(fs.readFileSync(inputPath));
    if (!options.contentTexts || options.contentTexts.length === 0) {
        throw new Error("contentTexts[] is required for searchReplace mode");
    }
    var xml = zip.file("word/document.xml").asText();
    xml = xml.replace(/\u2019/g, "'");
    var doc = new xmldom_1.DOMParser().parseFromString(xml, "text/xml");
    var paragraphs = Array.from(doc.getElementsByTagName("w:p"));
    var normalizedSearchTexts = options.contentTexts.map(normalizeText);
    for (var i = 0; i < paragraphs.length; i++) {
        var p = paragraphs[i];
        var normalizedParaText = normalizeText(getParagraphText(p));
        for (var s = 0; s < normalizedSearchTexts.length; s++) {
            var searchText = options.contentTexts[s];
            var normalizedSearch = normalizedSearchTexts[s];
            if (!normalizedParaText.includes(normalizedSearch))
                continue;
            styleTextInParagraph(p, doc, searchText, options);
            normalizedParaText = normalizeText(getParagraphText(p));
        }
    }
    zip.file("word/document.xml", new xmldom_1.XMLSerializer().serializeToString(doc));
    fs.writeFileSync(inputPath, zip.generate({ type: "nodebuffer" }));
    console.log("File saved!");
}
function searchUpdateMain(inputPath, options) {
    var zip = new pizzip_1.default(fs.readFileSync(inputPath));
    if (!options.contentTexts || options.contentTexts.length === 0) {
        throw new Error("contentTexts[] is required for searchReplace mode");
    }
    var xml = zip.file("word/document.xml").asText();
    xml = xml.replace(/\u2019/g, "'");
    var doc = new xmldom_1.DOMParser().parseFromString(xml, "text/xml");
    var paragraphs = Array.from(doc.getElementsByTagName("w:p"));
    var normalizedSearchTexts = options.contentTexts.map(normalizeText);
    for (var i = 0; i < paragraphs.length; i++) {
        var p = paragraphs[i];
        var normalizedParaText = normalizeText(getParagraphText(p));
        for (var s = 0; s < normalizedSearchTexts.length; s++) {
            var searchText = options.contentTexts[s];
            var normalizedSearch = normalizedSearchTexts[s];
            if (!normalizedParaText.includes(normalizedSearch))
                continue;
            styleTextInParagraph(p, doc, searchText, options);
            normalizedParaText = normalizeText(getParagraphText(p));
        }
    }
    zip.file("word/document.xml", new xmldom_1.XMLSerializer().serializeToString(doc));
    fs.writeFileSync(inputPath, zip.generate({ type: "nodebuffer" }));
    console.log("File saved!");
}
function replaceUpdate(inputPath, options, complete) {
    if (complete === void 0) { complete = false; }
    var zip = new pizzip_1.default(fs.readFileSync(inputPath));
    var xml = zip.file("word/document.xml").asText();
    xml = xml.replace(/\u2019/g, "'");
    var doc = new xmldom_1.DOMParser().parseFromString(xml, "text/xml");
    var paragraphs = doc.getElementsByTagName("w:p");
    for (var i = 0; i < paragraphs.length; i++) {
        var p = paragraphs[i];
        var runs = p.getElementsByTagName("w:r");
        if (!complete) {
            if (isNonContentParagraph(p)) {
                continue;
            }
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
    fs.writeFileSync(inputPath, zip.generate({ type: "nodebuffer" }));
    console.log("File saved!");
    console.log("Content replaced!");
}
// // ---------------- TITLE ----------------
// if (mode === "TITLE") {
//     updateTitle();
// }
// ---------------- HEADING ----------------
// if (mode === "HEADING") {
//     updateHeading();
// }
// // ---------------- SEARCH & REPLACE STYLE ----------------
// else if (mode === "CONTENT") {
//     searchUpdate();
// }
// // ---------------- SEARCH & REPLACE MAIN PAGE STYLE ----------------
// else if (mode === "MAIN") {
//     searchUpdateMain();
// }
// // ---------------- REPLACE CONTENT ----------------
// else if (mode === "COMPLETE") {
//     replaceUpdate();
// }
function processDocument(mode, templatePath, options) {
    switch (mode) {
        case "HEADING":
            updateHeading(templatePath, options);
            break;
        case "CONTENT":
            searchUpdate(templatePath, options);
            break;
        case "COMPLETE":
            replaceUpdate(templatePath, options, true);
            break;
        case "COMPLETE_CONTENT":
            replaceUpdate(templatePath, options);
            break;
        case "MAIN":
            searchUpdateMain(templatePath, options);
            break;
        default:
            break;
    }
}
processDocument(mode, "./final_result_1.docx", options);
// fs.writeFileSync("output.docx", zip.generate({ type: "nodebuffer" }) as any);
// console.log("File saved!");
