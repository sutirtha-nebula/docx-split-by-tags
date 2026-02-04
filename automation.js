import fs from "fs";
import path from "path";
import { execSync } from "child_process";
import PizZip from "pizzip";
import { DOMParser, XMLSerializer } from "xmldom";

// ---------------- CONFIG ----------------

const TEMPLATES_DIR = "./test-templates";
const RESULTS_DIR = "./results_layout";
const TMP_DIR = "./tmp";
const DEBUG = true;

// ----------------------------------------

function log(...args) {
    if (DEBUG) console.log(...args);
}

function ensureDir(dir) {
    if (!fs.existsSync(dir)) {
        fs.mkdirSync(dir, { recursive: true });
        log("[mkdir]", dir);
    }
}

function run(cmd) {
    log("[exec]", cmd);
    execSync(cmd, { stdio: "ignore" });
}

// ---------- HEADER / FOOTER STRIP --------

function removeHeadersFooters(inputDocx, outputDocx) {
    const zip = new PizZip(fs.readFileSync(inputDocx));

    Object.keys(zip.files).forEach(name => {
        if (name.startsWith("word/header") || name.startsWith("word/footer")) {
            delete zip.files[name];
        }
    });

    const docXml = zip.file("word/document.xml").asText();
    const doc = new DOMParser().parseFromString(docXml, "text/xml");

    const sectPrs = doc.getElementsByTagName("w:sectPr");
    Array.from(sectPrs).forEach(sectPr => {
        Array.from(sectPr.childNodes).forEach(n => {
            if (
                n.nodeName === "w:headerReference" ||
                n.nodeName === "w:footerReference"
            ) {
                sectPr.removeChild(n);
            }
        });
    });

    zip.file(
        "word/document.xml",
        new XMLSerializer().serializeToString(doc)
    );

    fs.writeFileSync(outputDocx, zip.generate({ type: "nodebuffer" }));
}

// ---------- LibreOffice → PDF ------------

function convertDocxToPdf(docxPath) {
    run(
        `soffice --headless --nologo --nolockcheck ` +
        `--convert-to pdf --outdir "${TMP_DIR}" "${docxPath}"`
    );

    const pdfPath = path.join(
        TMP_DIR,
        path.basename(docxPath, ".docx") + ".pdf"
    );

    if (!fs.existsSync(pdfPath)) {
        throw new Error("PDF conversion failed");
    }

    return pdfPath;
}

// ---------- Extract page-1 text -----------

function extractPdfPage1Text(pdfPath, baseName) {
    log("\n[step] Extract PDF page 1 text");

    const txtPath = path.join(TMP_DIR, `${baseName}_page1.txt`);

    run(`pdftotext -f 1 -l 1 -layout "${pdfPath}" "${txtPath}"`);

    const text = fs.readFileSync(txtPath, "utf8");
    log("[pdf page1 raw length]", text.length);
    log("[pdf page1 txt]", txtPath);

    return text;
}


// ---------- DOCX helpers -----------------

function normalizeText(t) {
    return t
        // non-breaking space
        .replace(/\u00a0/g, " ")

        // ) followed by capital
        .replace(/\)(?=[A-Z])/g, ") ")

        // lowercase followed by uppercase
        .replace(/([a-z])([A-Z])/g, "$1 $2")

        // letter followed by digit (Name4th)
        .replace(/([A-Za-z])(\d)/g, "$1 $2")

        // digit followed by letter (4thSeptember)
        .replace(/(\d)([A-Za-z])/g, "$1 $2")

        // collapse whitespace
        .replace(/\s+/g, " ")
        .trim();
}



function extractDocxBlocks(docxPath) {
    const zip = new PizZip(fs.readFileSync(docxPath));
    const xml = zip.file("word/document.xml").asText();
    const doc = new DOMParser().parseFromString(xml, "text/xml");
    const body = doc.getElementsByTagName("w:body")[0];

    const blocks = Array.from(body.childNodes).filter(
        n =>
            n.nodeType === 1 &&
            (n.nodeName === "w:p" || n.nodeName === "w:tbl")
    );

    return { zip, doc, body, blocks };
}

function paragraphText(p) {
    const texts = p.getElementsByTagName("w:t");
    return normalizeText(
        Array.from(texts).map(t => t.textContent).join("")
    );
}

function normalizeForMatch(t) {
    return t
        .toLowerCase()
        .replace(/\u00a0/g, "")
        .replace(/\s+/g, "")
        .replace(/[^a-z0-9]/g, "");
}


// ---------- Determine split index ----------

function determineSplitIndex(docxPath, page1Text) {
    const { blocks } = extractDocxBlocks(docxPath);
    const pageText = normalizeText(page1Text);

    let lastSafeIndex = -1;

    for (let i = 0; i < blocks.length; i++) {
        const block = blocks[i];

        if (block.nodeName === "w:tbl") break;

        const text = paragraphText(block);

        const blockKey = normalizeForMatch(text);
        const pageKey = normalizeForMatch(pageText);

        console.log("----")


        console.log(pageKey)
        console.log("+++++++")
        console.log(blockKey)

        console.log("----")
        if (!blockKey || pageText.includes(pageKey)) {
            lastSafeIndex = i;
        } else {
            break;
        }
    }

    return lastSafeIndex;
}

// ---------- Split DOCX --------------------

function splitDocxAtIndex(inputPath, splitIndex, outA, outB) {
    const original = extractDocxBlocks(inputPath);
    const serializer = new XMLSerializer();

    function cloneDoc() {
        const zip = new PizZip(fs.readFileSync(inputPath));
        const xml = zip.file("word/document.xml").asText();
        const doc = new DOMParser().parseFromString(xml, "text/xml");
        const body = doc.getElementsByTagName("w:body")[0];
        return { zip, doc, body };
    }

    const docA = cloneDoc();
    const docB = cloneDoc();

    [docA.body, docB.body].forEach(b => {
        while (b.firstChild) b.removeChild(b.firstChild);
    });

    original.blocks.forEach((block, i) => {
        const target = i <= splitIndex ? docA : docB;
        target.body.appendChild(
            target.doc.importNode(block, true)
        );
    });

    const sectPrs = original.body.getElementsByTagName("w:sectPr");
    if (sectPrs.length) {
        const last = sectPrs[sectPrs.length - 1];
        docA.body.appendChild(docA.doc.importNode(last, true));
        docB.body.appendChild(docB.doc.importNode(last, true));
    }

    docA.zip.file(
        "word/document.xml",
        serializer.serializeToString(docA.doc)
    );
    docB.zip.file(
        "word/document.xml",
        serializer.serializeToString(docB.doc)
    );

    fs.writeFileSync(outA, docA.zip.generate({ type: "nodebuffer" }));
    fs.writeFileSync(outB, docB.zip.generate({ type: "nodebuffer" }));
}

// ---------------- MAIN -------------------

function main() {
    ensureDir(TMP_DIR);
    ensureDir(RESULTS_DIR);

    const files = fs
        .readdirSync(TEMPLATES_DIR)
        .filter(f => f.toLowerCase().endsWith(".docx"));

    for (const file of files) {
        const base = path.basename(file, ".docx");
        log("\n==============================");
        log("Processing:", base);

        const input = path.join(TEMPLATES_DIR, file);
        const outDir = path.join(RESULTS_DIR, base);
        ensureDir(outDir);

        const headerless = path.join(TMP_DIR, `${base}_nohf.docx`);
        removeHeadersFooters(input, headerless);

        const pdf = convertDocxToPdf(headerless);
        const page1Text = extractPdfPage1Text(pdf, base);

        const splitIndex = determineSplitIndex(input, page1Text);

        splitDocxAtIndex(
            input,
            splitIndex,
            path.join(outDir, "first.docx"),
            path.join(outDir, "rest.docx")
        );

        log("✔ Done:", base);
    }

    console.log("\n✔ All documents processed");
}

main();
