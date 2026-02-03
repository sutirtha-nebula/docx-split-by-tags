import * as fs from "fs";
import PizZip from "pizzip";
import {DOMParser, XMLSerializer} from "xmldom";

type HeaderRule = "firstRow";
type FooterRule = "lastRow" | "containsText";

interface TableTarget {
    tableIndex?: number;
    tableMatchText?: string;
    header?: boolean;
    content?: boolean;
    footer?: boolean;
    headerRule?: HeaderRule;
    footerRule?: FooterRule;
    footerText?: string;
}

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
    alignment?: string;
    lineHeight?: any;
    contentTexts?: string[];
    verticalAlign?: "top" | "center" | "bottom";
    rowHeight?: number;        // twips
    tableTargets?: TableTarget[];
}

const mode: any = "TABLE"; // "TITLE" | "HEADING" | "CONTENT" | "COMPLETE" | "COMPLETE_CONTENT" | "MAIN" | "TABLE"

// const options: Options = {
//     headingTexts: ["Table of Contents", "Introduction", "Executive “Summary”", "Core Definitions", "Battery Technologies", "Structural Architecture"],
//     // prefixText: "",
//     // suffixText: "",
//     color: '#00FF00',
//     // fontSize: 22,
//     fontFamily: "Caveat",
//     underline: true,
//     bold: true,
//     italic: true,
//     bgColor: "#FF0000",
//     borderColor: "#FFFFFF",
//     borderSize: 12,
//     borderStyle: "dotted",
//     borderTop: true,
//     borderRight: true,
//     borderBottom: true,
//     borderLeft: true,
//     // highlightColor: "#000000",
//     alignment: 'center',
//     lineHeight: 2,
//     verticalAlign: "bottom",
//     // rowHeight: 2,
//     tableTargets: [
//         // {
//         //     tableIndex: 0,
//         //     header: true,
//         //     content: false,
//         //     footer: false,
//         //     headerRule: "firstRow"
//         // },
//         // {
//         //     tableIndex: 1,
//         //     header: true,
//         //     content: true,
//         //     footer: false,
//         //     headerRule: "firstRow"
//         // },
//         {
//             tableIndex: 2,
//             header: false,
//             content: true,
//             footer: false,
//             headerRule: "firstRow",
//             footerRule: "lastRow"
//         },
//         // {
//         //     tableIndex: 3,
//         //     header: false,
//         //     content: false,
//         //     footer: true,
//         //     headerRule: "firstRow",
//         //     footerRule: "lastRow"
//         // }
//         // {
//         //     tableIndex: 4,
//         //     header: false,
//         //     content: true,
//         //     footer: false,
//         //     headerRule: "firstRow",
//         //     footerRule: "lastRow"
//         // },
//         // {
//         //     tableIndex: 5,
//         //     header: true,
//         //     content: true,
//         //     footer: true,
//         //     headerRule: "firstRow",
//         //     footerRule: "lastRow"
//         // }
//     ],
//     contentTexts: [
//         "Date: January 2026",
//         "This document contains proprietary information about electric vehicle battery technologies. Distribution is intended for educational and",
//         "The battery represents approximately 30-40% of an electric vehicle's total cost and fundamentally determines its range, charging speed, longevity, and environmental footprint. Understanding the trajectory of battery technology is therefore essential for EV owners and prospective buyers making purchasing decisions, automotive industry professionals planning product strategies, sustainability stakeholders evaluating environmental impacts, researchers and students studying energy storage systems, and policy makers crafting regulations that will shape the industry's future.",
//         "Lithium-ion batteries dominate the current EV market, representing over 95% of electric vehicle energy storage. These batteries store energy through the movement of lithium ions between electrodes during charging and discharging.",
//         "Gravimetric Energy Density (Wh/kg): Energy per unit mass—critical for vehicle weight and efficiency"
//     ]
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

function getOrCreatePPr(p: Element, doc: Document): Element {
    for (let i = 0; i < p.childNodes.length; i++) {
        const n = p.childNodes[i];
        if (n.nodeName === "w:pPr") return n as Element;
    }

    const pPr = doc.createElement("w:pPr");
    p.insertBefore(pPr, p.firstChild);
    return pPr;
}

function removeChildren(pPr: Element, tag: string) {
    const nodes = Array.from(pPr.getElementsByTagName(tag));
    nodes.forEach(n => pPr.removeChild(n));
}

function applyBorderColor(run: any, rPr: any, opts: any, doc: any): void {
    const {top, right, bottom, left} = resolveBorderSides(opts);

    const borderOpts = {
        ...opts,
        borderTop: top,
        borderRight: right,
        borderBottom: bottom,
        borderLeft: left,
        borderStyle: opts.borderStyle || "single"  // default to "single" if not provided
    };

    const p = run.parentNode;

    if(p && paragraphHasText(p)){
        try {
            applyBorder(p, doc, borderOpts);
        } catch (e) {
            applyBorder(rPr, doc, borderOpts);
        }
    }
}

function paragraphHasText(p: any): boolean {
    const runs = Array.from(p.getElementsByTagName("w:r"));
    return runs.some((r: any) =>
        Array.from(r.getElementsByTagName("w:t")).some((t: any) =>
            t.textContent && t.textContent.trim() !== ""
        )
    );
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
        if (opts.underline !== null && opts.underline !== undefined) {
            let u = rPr.getElementsByTagName("w:u")[0];
            if (opts.underline === true) {
                if (!u) {
                    u = doc.createElement("w:u");
                    rPr.appendChild(u);
                    u.setAttribute("w:val", "single");
                }

            } else {
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
                try {
                    const p = run.parentNode;

                    const hasVisibleText =
                        p &&
                        Array.from(p.getElementsByTagName("w:t"))
                            .map((t: any) => (t.textContent || "").replace(/\u00A0/g, ""))
                            .join("")
                            .trim().length > 0;

                    if (!hasVisibleText) {
                        return;
                    }

                    let pPr: any;

                    try {
                        pPr = p.getElementsByTagName("w:pPr")[0];
                        console.log("pPr found in paragraph");
                    } catch (e) {
                        console.log("No pPr in paragraph");
                        pPr = p.getElementsByTagName("w:hyperlink")[0];
                        console.log("pPr found in hyperlink");
                    }

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

                    // Instead of manually setting shading here, call your function:
                    applyBorderColor(run, rPr, {borderColor: opts.bgColor, borderSize: 6}, doc);
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
            applyBorderColor(run, rPr, opts, doc);
        }
        if (opts?.alignment != null) {
            const p = run.parentNode as Element;
            if (!p) return;

            const pPr = getOrCreatePPr(p, doc);

            // Remove existing jc (important!)
            removeChildren(pPr, "w:jc");

            const pStyleNodes = pPr.getElementsByTagName("w:pStyle");
            if (pStyleNodes && pStyleNodes.length > 0) {
                pPr.removeChild(pStyleNodes[0]);
            }

            const jc = doc.createElement("w:jc");

            const map: Record<string, string> = {
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
                const p = run.parentNode as Element;

                let pPr = p.getElementsByTagName("w:pPr")[0];
                if (!pPr) {
                    pPr = doc.createElement("w:pPr");
                    p.insertBefore(pPr, p.firstChild);
                }

                let spacing = pPr.getElementsByTagName("w:spacing")[0];
                if (!spacing) {
                    spacing = doc.createElement("w:spacing");
                    pPr.appendChild(spacing);
                }

                // opts.lineHeight can be like 1, 1.5, 2
                const lineTwips = Math.round(240 * opts.lineHeight);

                spacing.setAttribute("w:line", String(lineTwips));
                spacing.setAttribute("w:lineRule", "auto");
            } catch (e) {
                // optional fallback – usually paragraph spacing never belongs in rPr
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

    return {top, right, bottom, left};
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
        {side: "top", enabled: opts.borderTop},
        {side: "left", enabled: opts.borderLeft},
        {side: "bottom", enabled: opts.borderBottom},
        {side: "right", enabled: opts.borderRight}
    ];

    console.log("Applying borders:", borderConfig);

    borderConfig.forEach(({side, enabled}) => {
        if (!enabled) return;

        const node = doc.createElement(`w:${side}`);
        node.setAttribute("w:val", opts.borderStyle);
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

        return {runInfo, fullText};
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

function updateHeading(inputPath: any, options: any): void {
    const zip = new PizZip(fs.readFileSync(inputPath));
    let xml = zip.file("word/document.xml")!.asText();
    xml = xml.replace(/\u2019/g, "'");
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
    fs.writeFileSync(inputPath, zip.generate({type: "nodebuffer"}) as any);
    console.log("File saved!");
    console.log("Headings updated!");
}

function searchUpdate(inputPath: any, options: any): void {
    const zip = new PizZip(fs.readFileSync(inputPath));
    if (!options.contentTexts || options.contentTexts.length === 0) {
        throw new Error("contentTexts[] is required for searchReplace mode");
    }

    let xml = zip.file("word/document.xml")!.asText();
    xml = xml.replace(/\u2019/g, "'");
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
    fs.writeFileSync(inputPath, zip.generate({type: "nodebuffer"}) as any);
    console.log("File saved!");
}

function searchUpdateMain(inputPath: any, options: any): void {
    const zip = new PizZip(fs.readFileSync(inputPath));
    if (!options.contentTexts || options.contentTexts.length === 0) {
        throw new Error("contentTexts[] is required for searchReplace mode");
    }

    let xml = zip.file("word/document.xml")!.asText();
    xml = xml.replace(/\u2019/g, "'");
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
    fs.writeFileSync(inputPath, zip.generate({type: "nodebuffer"}) as any);
    console.log("File saved!");
}

function replaceUpdate(inputPath: any, options: any, complete: any = false): void {
    const zip = new PizZip(fs.readFileSync(inputPath));
    let xml = zip.file("word/document.xml")!.asText();
    xml = xml.replace(/\u2019/g, "'");
    const doc = new DOMParser().parseFromString(xml, "text/xml");
    const paragraphs = doc.getElementsByTagName("w:p");

    for (let i = 0; i < paragraphs.length; i++) {
        const p = paragraphs[i];
        const runs = p.getElementsByTagName("w:r");

        if (!complete) {
            if (isNonContentParagraph(p)) {
                continue;
            }
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
    fs.writeFileSync(inputPath, zip.generate({type: "nodebuffer"}) as any);
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
            replaceUpdate(templatePath, options, true);
            break;

        case "COMPLETE_CONTENT":
            replaceUpdate(templatePath, options);
            break;

        case "MAIN":
            searchUpdateMain(templatePath, options);
            break;

        case "TABLE":
            searchUpdateTable(templatePath, options);
            break;

        default:
            break;
    }

}

interface SearchItem {
    raw: string;
    normalized: string;
}

function getRowText(tr: Element): string {
    return Array.from(tr.getElementsByTagName("w:t"))
        .map((t) => t.textContent || "")
        .join("");
}

function searchUpdateTable(inputPath: string, options: Options) {
    const zip = new PizZip(fs.readFileSync(inputPath));
    const xml = zip.file("word/document.xml")!.asText();
    const doc = new DOMParser().parseFromString(xml, "text/xml");

    const searches: SearchItem[] = options.contentTexts
        ? options.contentTexts.map((t) => ({
            raw: t,
            normalized: normalizeText(t),
        }))
        : [];

    const tables = Array.from(doc.getElementsByTagName("w:tbl")) as Element[];
    console.log("Total tables found:", tables.length);

    tables.forEach((tbl, tableIndex) => {
        const targets = resolveTableTargets(tbl, tableIndex, options);
        console.log(
            `Table #${tableIndex} - Targets found: ${targets.length}`
        );

        if (targets.length === 0) return;

        console.log(`Processing Table #${tableIndex}`);
        console.log("Searches for text styling:", searches);

        if (searches.length > 0 && options.contentTexts) {
            applyTextStylingInsideTable(tbl, doc, searches, options);
        }

        targets.forEach((target) => {
            applyStyleToTable(tbl, doc, options, target);
        });
    });

    zip.file("word/document.xml", new XMLSerializer().serializeToString(doc));


    fs.writeFileSync(inputPath, zip.generate({type: "nodebuffer"}));


}

function resolveTableTargets(
    tbl: Element,
    index: number,
    opts: Options
): TableTarget[] {
    if (!opts.tableTargets || opts.tableTargets.length === 0) return [];

    return opts.tableTargets.filter((t) => {
        if (t.tableIndex !== undefined) {
            return t.tableIndex === index;
        }
        if (t.tableMatchText) {
            return tableContainsText(tbl, t.tableMatchText);
        }
        return false;
    });
}

function isHeaderRow(tr: Element, rowIndex: number, target: TableTarget): boolean {
    if (target.headerRule === "firstRow") {
        return rowIndex === 0;
    }
    return false;
}

function isFooterRow(
    tr: Element,
    rowIndex: number,
    totalRows: number,
    target: TableTarget
): boolean {
    console.log("Checking footer for row index:", rowIndex, "of", totalRows);
    console.log("Footer target exists, checking rule:", target.footerRule);

    if (target.footerRule === "lastRow") {
        return rowIndex === totalRows - 1;
    }

    if (target.footerRule === "containsText" && target.footerText) {
        return normalizeText(getRowText(tr)).includes(
            normalizeText(target.footerText)
        );
    }

    return false;
}

function applyStyleToTable(
    tbl: Element,
    doc: Document,
    opts: Options,
    target: TableTarget
) {
    const rows = Array.from(tbl.getElementsByTagName("w:tr"));

    rows.forEach((tr, rowIndex) => {
        const isHeader = isHeaderRow(tr, rowIndex, target);
        const isFooter = isFooterRow(tr, rowIndex, rows.length, target);
        const isContent = !isHeader && !isFooter;

        let shouldStyle = false;
        if (isHeader && target.header) shouldStyle = true;
        if (isFooter && target.footer) shouldStyle = true;
        if (isContent && target.content) shouldStyle = true;

        console.log(
            `Row ${rowIndex} → H:${isHeader} C:${isContent} F:${isFooter} | APPLY:${shouldStyle}`
        );

        if (!shouldStyle) return;

        applyStyleToTableRow(tr, doc, opts);
    });
}

function applyStyleToTableRow(tr: Element, doc: Document, opts: Options) {
    let trPr = tr.getElementsByTagName("w:trPr")[0] as Element | undefined;

    if (!trPr) {
        trPr = doc.createElement("w:trPr");
        tr.insertBefore(trPr, tr.firstChild);
    }

    if (opts.rowHeight) {
        let h = trPr.getElementsByTagName("w:trHeight")[0] as Element | undefined;
        if (!h) {
            h = doc.createElement("w:trHeight");
            trPr.appendChild(h);
        }
        h.setAttribute("w:val", opts.rowHeight.toString());
    }

    const cells = Array.from(tr.getElementsByTagName("w:tc"));
    cells.forEach((tc) => applyStyleToTableCell(tc, doc, opts));
}

function applyStyleToTableCell(tc: Element, doc: Document, opts: Options) {
    let tcPr = tc.getElementsByTagName("w:tcPr")[0] as Element | undefined;

    if (!tcPr) {
        tcPr = doc.createElement("w:tcPr");
        tc.insertBefore(tcPr, tc.firstChild);
    }

    if (opts.bgColor) {
        let shd = tcPr.getElementsByTagName("w:shd")[0] as Element | undefined;
        if (!shd) {
            shd = doc.createElement("w:shd");
            tcPr.appendChild(shd);
        }
        shd.setAttribute("w:fill", opts.bgColor);
    }

    if (opts.verticalAlign) {
        let vAlign = tcPr.getElementsByTagName("w:vAlign")[0] as Element | undefined;
        if (!vAlign) {
            vAlign = doc.createElement("w:vAlign");
            tcPr.appendChild(vAlign);
        }
        vAlign.setAttribute("w:val", opts.verticalAlign);
    }

    const paragraphs = Array.from(tc.getElementsByTagName("w:p"));
    paragraphs.forEach((p) => {
        applyParagraphAlignment(p, doc, opts);
        applyStyleToRuns(Array.from(p.getElementsByTagName("w:r")), doc, opts);
    });
}

function applyParagraphAlignment(p: Element, doc: Document, opts: Options) {
    if (!opts.alignment) return;

    let pPr = p.getElementsByTagName("w:pPr")[0] as Element | undefined;

    if (!pPr) {
        pPr = doc.createElement("w:pPr");
        p.insertBefore(pPr, p.firstChild);
    }

    let jc = pPr.getElementsByTagName("w:jc")[0] as Element | undefined;
    if (!jc) {
        jc = doc.createElement("w:jc");
        pPr.appendChild(jc);
    }

    jc.setAttribute("w:val", opts.alignment);
}

function applyTextStylingInsideTable(
    tbl: Element,
    doc: Document,
    searches: SearchItem[],
    opts: Options
) {
    const paragraphs = Array.from(tbl.getElementsByTagName("w:p"));
    console.log("Total paragraphs in table:", paragraphs.length);

    paragraphs.forEach((p) => {
        let paraText = normalizeText(getParagraphText(p));
        console.log("Processing paragraph text in table:", paraText);

        searches.forEach(({raw, normalized}) => {
            if (!paraText.includes(normalized)) return;

            styleTextInParagraph(p, doc, raw, opts);
            paraText = normalizeText(getParagraphText(p));
        });
    });
}

function tableContainsText(tbl: Element, text: string): boolean {
    const fullText = Array.from(tbl.getElementsByTagName("w:t"))
        .map((t) => t.textContent || "")
        .join("");

    return normalizeText(fullText).includes(normalizeText(text));
}

// function removeTableBorders(tbl: Element, doc: Document) {
//     let tblPr = tbl.getElementsByTagName("w:tblPr")[0] as Element | undefined;

//     if (!tblPr) {
//         tblPr = doc.createElement("w:tblPr");
//         tbl.insertBefore(tblPr, tbl.firstChild);
//     }

//     let tblBorders = tblPr.getElementsByTagName("w:tblBorders")[0] as Element | undefined;

//     if (!tblBorders) {
//         tblBorders = doc.createElement("w:tblBorders");
//         tblPr.appendChild(tblBorders);
//     }

//     const borderTags = ["w:top", "w:left", "w:bottom", "w:right", "w:insideH", "w:insideV"];

//     borderTags.forEach((tag) => {
//         let border = tblBorders.getElementsByTagName(tag)[0] as Element | undefined;
//         if (!border) {
//             border = doc.createElement(tag);
//             tblBorders.appendChild(border);
//         }
//         border.setAttribute("w:val", "nil");
//     });
// }

// function addTableBorders(tbl: Element, doc: Document) {
//     let tblPr = tbl.getElementsByTagName("w:tblPr")[0] as Element | undefined;

//     if (!tblPr) {
//         tblPr = doc.createElement("w:tblPr");
//         tbl.insertBefore(tblPr, tbl.firstChild);
//     }

//     let tblBorders = tblPr.getElementsByTagName("w:tblBorders")[0] as Element | undefined;

//     if (!tblBorders) {
//         tblBorders = doc.createElement("w:tblBorders");
//         tblPr.appendChild(tblBorders);
//     }

//     const borderTags = ["w:top", "w:left", "w:bottom", "w:right", "w:insideH", "w:insideV"];

//     borderTags.forEach((tag) => {
//         let border = tblBorders.getElementsByTagName(tag)[0] as Element | undefined;
//         if (!border) {
//             border = doc.createElement(tag);
//             tblBorders.appendChild(border);
//         }

//         border.setAttribute("w:val", "single");
//         border.setAttribute("w:sz", "4");
//         border.setAttribute("w:space", "0");
//         border.setAttribute("w:color", "000000");
//     });
// }

// processDocument(mode, "./final_result_1.docx", options)
// fs.writeFileSync("output.docx", zip.generate({ type: "nodebuffer" }) as any);
// console.log("File saved!");