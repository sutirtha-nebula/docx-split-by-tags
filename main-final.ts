import * as fs from "fs";
import PizZip from "pizzip";
import { DOMParser, XMLSerializer } from "xmldom";

interface Options {
    headingTexts?: string[];
    prefixText?: string;
    suffixText?: string;
    color?: string;
    fontSize?: any;
    fontFamily?: string;
    underline?: boolean;
    bold?: boolean;
    italic?: boolean;
    bgColor?: any;
    borderColor?: string;
    borderSize?: any;
    borderStyle?: string;
    borderTop?: boolean;
    borderRight?: boolean;
    borderBottom?: boolean;
    borderLeft?: boolean;
    highlightColor?: string;
    contentTexts?: string[];
}

const mode: any = "CONTENT"; // "TITLE" | "HEADING" | "CONTENT" | "COMPLETE" | "MAIN"

// const options: Options = {
//     headingTexts: ["Executive \"Summary\"", "Modified"],
//     prefixText: "",
//     suffixText: "",
//     color: '49A361',
//     fontSize: null,
//     fontFamily: "Calibri",
//     underline: true,
//     bold: true,
//     italic: true,
//     bgColor: null,
//     borderColor: "black",
//     borderSize: 6,
//     borderStyle: "single",
//     borderTop: true,
//     // borderRight: false,
//     // borderBottom: false,
//     // borderLeft: false,
//     highlightColor: "transparent",
//     contentTexts: [
//         "It has survived not only five centuries, but also the leap into electronic typesetting, remaining essentially unchanged. It was popularised in the 1960s with the release of Letraset sheets containing Lorem Ipsum passages, and more recently with desktop publishing software like Aldus PageMaker including versions of Lorem Ipsum.",
//     ],
// };

function addPrefixSuffix(text: string, opts: any): string {
    return `${opts.prefixText || ""}${text}${opts.suffixText || ""}`;
}

function normalizeText(text: string): string {
    return text
        .normalize("NFKC")
        .replace(/[\u2018\u2019]/g, "'")
        .replace(/[\u201C\u201D]/g, '"')
        .replace(/\u00A0/g, " ")
        .replace(/\s+/g, " ")
        .trim();
}

function applyStyleToRuns(runs: any[], doc: any, opts: any): void {
    for (let r = 0; r < runs.length; r++) {
        const run = runs[r];
        let rPr = run.getElementsByTagName("w:rPr")[0];
        if (!rPr) {
            rPr = doc.createElement("w:rPr");
            run.insertBefore(rPr, run.firstChild);
        }
        
       if (opts.bold !== null && opts.bold !== undefined) {
    let b = rPr.getElementsByTagName("w:b")[0];

    if (opts.bold === true) {
        if (!b) {
            b = doc.createElement("w:b");
            rPr.appendChild(b);
        }
    } else {
        if (b) {
            rPr.removeChild(b);
        }
    }
}

        if (opts.italic !== null && opts.italic !== undefined) {
            let i = rPr.getElementsByTagName("w:i")[0];
            if (opts.italic === true) {
                if (!i) {
                    i = doc.createElement("w:i");
                    rPr.appendChild(i);
                }
            } else {
                if (i) {
                    rPr.removeChild(i);
                }
            }
        }
                if (opts.underline!== null && opts.underline!== undefined) {
            let u = rPr.getElementsByTagName("w:u")[0];
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
            let color = rPr.getElementsByTagName("w:color")[0];
            if (!color) {
                color = doc.createElement("w:color");
                rPr.appendChild(color);
            }
            color.setAttribute("w:val", opts.color);
        }
        
        if (opts.fontSize) {
            let sz = rPr.getElementsByTagName("w:sz")[0];
            if (!sz) {
                sz = doc.createElement("w:sz");
                rPr.appendChild(sz);
            }
            sz.setAttribute("w:val", opts.fontSize.toString());
        }
        
        if (opts.fontFamily) {
            let rFonts = rPr.getElementsByTagName("w:rFonts")[0];
            if (!rFonts) {
                rFonts = doc.createElement("w:rFonts");
                rPr.appendChild(rFonts);
            }
            rFonts.setAttribute("w:ascii", opts.fontFamily);
            rFonts.setAttribute("w:hAnsi", opts.fontFamily);
        }
        
        if (opts.highlightColor) {
            let shd = rPr.getElementsByTagName("w:shd")[0];
            if (!shd) {
                shd = doc.createElement("w:shd");
                rPr.appendChild(shd);
            }
            if (opts.highlightColor === "transparent") {
                shd.setAttribute("w:val", "clear");
                shd.setAttribute("w:color", "auto");
                shd.setAttribute("w:fill", "auto");
                rPr.appendChild(shd);
            } else {
                shd.setAttribute("w:fill", opts.highlightColor);
                rPr.appendChild(shd);
            }
        }

        if ("bgColor" in opts && opts.bgColor !== null) {
            if (opts.bgColor) {
                try{
                const p = run.parentNode;
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
                shd.setAttribute("w:fill", opts.bgColor);
            } catch (e) {
                 let shd = rPr.getElementsByTagName("w:shd")[0];
            if (!shd) {
                shd = doc.createElement("w:shd");
                rPr.appendChild(shd);
            }
            if (opts.highlightColor === "transparent") {
                shd.setAttribute("w:val", "clear");
                shd.setAttribute("w:color", "auto");
                shd.setAttribute("w:fill", "auto");
                rPr.appendChild(shd);
            } else {
                shd.setAttribute("w:fill", opts.highlightColor);
                rPr.appendChild(shd);
            }
            }
            }
        }
        if ("borderSize" in opts && opts.borderSize !== 0) {
            const { top, right, bottom, left } = resolveBorderSides(opts);

            const borderOpts = {
                ...opts,
                borderTop: top,
                borderRight: right,
                borderBottom: bottom,
                borderLeft: left
            };

            const p = run.parentNode;

            try {
                applyBorder(p, doc, borderOpts);
            } catch (e) {
                applyBorder(rPr, doc, borderOpts);
            }
        }
    }
}

function resolveBorderSides(opts: Options) {
    const hasAnySide =
        opts.borderTop !== undefined ||
        opts.borderRight !== undefined ||
        opts.borderBottom !== undefined ||
        opts.borderLeft !== undefined;

    // Default values
    let top = false;
    let right = false;
    let bottom = false;
    let left = false;

    if (opts.borderSize && opts.borderSize !== 0) {
        if (!hasAnySide) {
            // borderSize only → all sides
            top = right = bottom = left = true;
        } else {
            // respect provided sides
            top = !!opts.borderTop;
            right = !!opts.borderRight;
            bottom = !!opts.borderBottom;
            left = !!opts.borderLeft;
        }
    }

    return { top, right, bottom, left };
}


function applyBorder(p: any, doc: any, opts: any): void {
    let pPr = p.getElementsByTagName("w:pPr")[0];
    if (!pPr) {
        pPr = doc.createElement("w:pPr");
        p.insertBefore(pPr, p.firstChild);
    }
    
    const oldBdr = pPr.getElementsByTagName("w:pBdr")[0];
    if (oldBdr) {
        pPr.removeChild(oldBdr);
    }

    // If no borderSize or zero → do nothing
    if (!opts.borderSize || opts.borderSize === 0) {
        return;
    }

    const pBdr = doc.createElement("w:pBdr");
    pPr.appendChild(pBdr);

    const borderConfig = [
        { side: "top", enabled: opts.borderTop },
        { side: "left", enabled: opts.borderLeft },
        { side: "bottom", enabled: opts.borderBottom },
        { side: "right", enabled: opts.borderRight }
    ];

    console.log("Applying borders:", borderConfig);

    borderConfig.forEach(({ side, enabled }) => {
        if (!enabled) return;

        const node = doc.createElement(`w:${side}`);
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

function isHeadingStyle(p: any): boolean {
    const pPr = p.getElementsByTagName("w:pPr")[0];
    if (!pPr) return false;
    
    const pStyle = pPr.getElementsByTagName("w:pStyle")[0];
    if (!pStyle) return false;
    
    const styleVal = pStyle.getAttribute("w:val");
    return !!styleVal && styleVal.startsWith("Heading");
}

function getParagraphText(p: any): string {
    const texts = p.getElementsByTagName("w:t");
    let fullText = "";
    
    for (let i = 0; i < texts.length; i++) {
        fullText += texts[i].textContent || "";
    }
    
    return fullText;
}

function cloneRunProperties(r: any): any {
    const rPr = r.getElementsByTagName("w:rPr")[0];
    return rPr ? rPr.cloneNode(true) : null;
}

function createRun(doc: any, text: string, styled: boolean, existingRPr: any, options: any): any {
    const run = doc.createElement("w:r");
    
    if (existingRPr) {
        run.appendChild(existingRPr.cloneNode(true));
    }
    
    if (styled) {
        applyStyleToRuns([run], doc, options);
    }
    
    const t = doc.createElement("w:t");
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

function isNonContentStyle(styleName: any): boolean {
    if (!styleName) return false;
    
    const nonContentPrefixes = [
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
    
    return nonContentPrefixes.some((prefix: any) => styleName.startsWith(prefix));
}

function isCenteredParagraph(p: any): boolean {
    const pPr = p.getElementsByTagName("w:pPr")[0];
    if (!pPr) return false;
    
    const jc = pPr.getElementsByTagName("w:jc")[0];
    if (!jc) return false;
    
    const val = jc.getAttribute("w:val");
    return val === "center";
}

function hasLargeFontSize(p: any): boolean {
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
        if (size > 32) {
            return true;
        }
    }
    
    return false;
}

function styleTextInParagraph(p: any, doc: any, searchText: string, options: any): void {
    function isHeadingParagraph(p: any): boolean {
        const pPr = p.getElementsByTagName("w:pPr")[0];
        if (!pPr) return false;

        const pStyle = pPr.getElementsByTagName("w:pStyle")[0];
        if (!pStyle) return false;

        const val = pStyle.getAttribute("w:val") || "";
        return /^Heading\d+$/.test(val) || val === "Heading";
    }

    // ---- MODE CHECK (ONLY ADDITION) ----
    if (mode === "CONTENT" && isHeadingParagraph(p)) {
        return;
    }
    function buildRunInfo(): any {
        const runs = Array.from(p.getElementsByTagName("w:r"));
        const runInfo: any[] = [];
        let fullText = "";
        
        runs.forEach((r: any) => {
            const t = r.getElementsByTagName("w:t")[0];
            if (!t) return;
            
            const text = t.textContent || "";
            if (!text) return;
            
            runInfo.push({
                r: r,
                text: text,
                start: fullText.length,
                end: fullText.length + text.length
            });
            fullText += text;
        });
        
        return { runInfo, fullText };
    }
    
    const data = buildRunInfo();
    
    const matches: any[] = [];
    let idx = 0;
    
    while ((idx = data.fullText.indexOf(searchText, idx)) !== -1) {
        matches.push({
            start: idx,
            end: idx + searchText.length
        });
        idx += searchText.length;
    }
    
    if (matches.length === 0) return;
    
    matches.reverse().forEach((m: any) => {
        const rebuilt = buildRunInfo();
        const runInfo = rebuilt.runInfo;
        
        const matchStart = m.start;
        const matchEnd = m.end;
        
        const affected = runInfo.filter((info: any) => {
            return info.end > matchStart && info.start < matchEnd;
        });
        
        if (affected.length === 0) return;
        
        const parent = affected[0].r.parentNode;
        const anchor = affected[affected.length - 1].r.nextSibling;
        const newRuns: any[] = [];
        
        affected.forEach((info: any) => {
            parent.removeChild(info.r);
        });
        
        affected.forEach((info: any) => {
            const beforeLen = Math.max(0, matchStart - info.start);
            const matchLen = Math.min(info.text.length, matchEnd - info.start) - beforeLen;
            
            const before = info.text.slice(0, beforeLen);
            const match = info.text.slice(beforeLen, beforeLen + matchLen);
            const after = info.text.slice(beforeLen + matchLen);
            
            const rPr = cloneRunProperties(info.r);
            
            if (before) {
                parent.insertBefore(createRun(doc, before, false, rPr, options), anchor);
            }
            
            if (match) {
                const matchRun = createRun(doc, match, true, rPr, options);
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

function isNonContentParagraph(p: any): boolean {
    const pPr = p.getElementsByTagName("w:pPr")[0];
    let styleName: any = null;
    
    if (pPr) {
        const pStyle = pPr.getElementsByTagName("w:pStyle")[0];
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

function updateHeading(inputPath:any,options:any): void {
    const zip = new PizZip(fs.readFileSync(inputPath));
    const xml = zip.file("word/document.xml")!.asText();
    const doc = new DOMParser().parseFromString(xml, "text/xml");
    const paragraphs = doc.getElementsByTagName("w:p");
    const headingTexts = (options.headingTexts || []).map((t: any) => {
        return normalizeText(t);
    });
    
    for (let i = 0; i < paragraphs.length; i++) {
        const p = paragraphs[i];
        
        if (!isHeadingStyle(p)) continue;
        
        const headingText = getParagraphText(p);
        const normalizedHeading = normalizeText(headingText);
        
        if (headingTexts.length > 0 && !headingTexts.includes(normalizedHeading)) {
            continue;
        }
        
        const runs = Array.from(p.getElementsByTagName("w:r"));
        const textRuns = runs.filter((r: any) => {
            const t = r.getElementsByTagName("w:t")[0];
            return t && t.textContent;
        });
        
        if (textRuns.length > 0) {
            const prefix = options.prefixText || "";
            const suffix = options.suffixText || "";
            
            const firstT = (textRuns[0] as any).getElementsByTagName("w:t")[0];
            const lastT = (textRuns[textRuns.length - 1] as any).getElementsByTagName("w:t")[0];
            
            const fullText = getParagraphText(p);
            
            if (prefix && !fullText.startsWith(prefix)) {
                firstT.textContent = prefix + firstT.textContent;
                console.log("Added prefix to heading:", prefix, firstT.textContent);
            }
            
            if (suffix && !fullText.endsWith(suffix)) {
                lastT.textContent = lastT.textContent + suffix;
                console.log("Added suffix to heading:", suffix, lastT.textContent);
            }
        }
        
        const runsToStyle = Array.from(p.getElementsByTagName("w:r"));
        applyStyleToRuns(runsToStyle, doc, options);
    }
    
    zip.file("word/document.xml", new XMLSerializer().serializeToString(doc));
    fs.writeFileSync(inputPath, zip.generate({ type: "nodebuffer" }) as any);
console.log("File saved!");
    console.log("Headings updated!");
}

function searchUpdate(inputPath:any,options:any): void {
    const zip = new PizZip(fs.readFileSync(inputPath));
    if (!options.contentTexts || options.contentTexts.length === 0) {
        throw new Error("contentTexts[] is required for searchReplace mode");
    }
    
    const xml = zip.file("word/document.xml")!.asText();
    const doc = new DOMParser().parseFromString(xml, "text/xml");
    const paragraphs = Array.from(doc.getElementsByTagName("w:p"));
    const normalizedSearchTexts = options.contentTexts!.map(normalizeText);
    
    for (let i = 0; i < paragraphs.length; i++) {
        const p = paragraphs[i];
        let normalizedParaText = normalizeText(getParagraphText(p));
        
        for (let s = 0; s < normalizedSearchTexts.length; s++) {
            const searchText = options.contentTexts![s];
            const normalizedSearch = normalizedSearchTexts[s];
            
            if (!normalizedParaText.includes(normalizedSearch)) continue;
            
            styleTextInParagraph(p, doc, searchText, options);
            
            normalizedParaText = normalizeText(getParagraphText(p));
        }
    }
    
    zip.file("word/document.xml", new XMLSerializer().serializeToString(doc));
    fs.writeFileSync(inputPath, zip.generate({ type: "nodebuffer" }) as any);
console.log("File saved!");
}
function searchUpdateMain(inputPath:any,options:any): void {
    const zip = new PizZip(fs.readFileSync(inputPath));
    if (!options.contentTexts || options.contentTexts.length === 0) {
        throw new Error("contentTexts[] is required for searchReplace mode");
    }
    
    const xml = zip.file("word/document.xml")!.asText();
    const doc = new DOMParser().parseFromString(xml, "text/xml");
    const paragraphs = Array.from(doc.getElementsByTagName("w:p"));
    const normalizedSearchTexts = options.contentTexts!.map(normalizeText);
    
    for (let i = 0; i < paragraphs.length; i++) {
        const p = paragraphs[i];
        let normalizedParaText = normalizeText(getParagraphText(p));
        
        for (let s = 0; s < normalizedSearchTexts.length; s++) {
            const searchText = options.contentTexts![s];
            const normalizedSearch = normalizedSearchTexts[s];
            
            if (!normalizedParaText.includes(normalizedSearch)) continue;
            
            styleTextInParagraph(p, doc, searchText, options);
            
            normalizedParaText = normalizeText(getParagraphText(p));
        }
    }
    
    zip.file("word/document.xml", new XMLSerializer().serializeToString(doc));
    fs.writeFileSync(inputPath, zip.generate({ type: "nodebuffer" }) as any);
console.log("File saved!");
}

function replaceUpdate(inputPath:any,options:any): void {
    const zip = new PizZip(fs.readFileSync(inputPath));
    const xml = zip.file("word/document.xml")!.asText();
    const doc = new DOMParser().parseFromString(xml, "text/xml");
    const paragraphs = doc.getElementsByTagName("w:p");
    
    for (let i = 0; i < paragraphs.length; i++) {
        const p = paragraphs[i];
        const runs = p.getElementsByTagName("w:r");
        
        if (isNonContentParagraph(p)) {
            continue;
        }
        
        for (let r = 0; r < runs.length; r++) {
            const t = runs[r].getElementsByTagName("w:t")[0];
            if (t && t.textContent) {
                t.textContent = addPrefixSuffix(t.textContent, options);
            }
        }
        
        const runsArray = Array.from(runs);
        applyStyleToRuns(runsArray, doc, options);
    }
    
    zip.file("word/document.xml", new XMLSerializer().serializeToString(doc));
    fs.writeFileSync(inputPath, zip.generate({ type: "nodebuffer" }) as any);
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

function processDocument(
    mode: any,
    templatePath: string,
    options: Options
): void {
    switch (mode) {
        case "HEADING":
            updateHeading(templatePath, options);
            break;

        case "CONTENT":
            searchUpdate(templatePath, options);
            break;

        case "COMPLETE":
            replaceUpdate(templatePath, options);
            break;

        case "MAIN":
            searchUpdateMain(templatePath, options);
            break;

        default:
            break;
    }

}

// processDocument(mode, "./templates/2026_01_Precision_AI_UFA_Template1.docx", options)
// fs.writeFileSync("output.docx", zip.generate({ type: "nodebuffer" }) as any);
console.log("File saved!");