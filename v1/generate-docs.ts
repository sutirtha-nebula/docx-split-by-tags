import * as fs from "fs";
import * as path from "path";

import { extractFirstDoc } from "./split-by-tags.ts";
import { extractRemainingDoc } from "./extractRemainingDoc";
import { mergeDocx } from "./mergeDocx";

/* ============================
   CONFIG
============================ */

const TEMPLATES_DIR = path.resolve("./templates");
const RESULTS_DIR = path.resolve("./results");

const SPLIT_TAGS: string[] = [
    `<w:br w:type="page"/>`,
    `<w:lastRenderedPageBreak/>`,
    `<w:pageBreakBefore w:val="0"/>`,
];

/* ============================
   HELPERS
============================ */

function ensureDir(dir: string) {
    fs.mkdirSync(dir, { recursive: true });
}

function sanitizeForFolder(name: string): string {
    return name
        .replace(/[<>:"/\\|?*\s]+/g, "_")
        .replace(/_+/g, "_")
        .slice(0, 100);
}

function getDocxFiles(dir: string): string[] {
    return fs
        .readdirSync(dir)
        .filter(f => f.toLowerCase().endsWith(".docx"))
        .map(f => path.join(dir, f));
}

/* ============================
   MAIN RUNNER
============================ */

function run() {
    const files = getDocxFiles(TEMPLATES_DIR);

    if (!files.length) {
        console.warn("⚠️ No DOCX files found in templates/");
        return;
    }

    for (const inputFile of files) {
        const fileName = path.basename(inputFile, ".docx");
        const fileResultRoot = path.join(RESULTS_DIR, fileName);

        for (const splitTag of SPLIT_TAGS) {
            const tagFolder = sanitizeForFolder(splitTag);
            const outputDir = path.join(fileResultRoot, tagFolder);

            ensureDir(outputDir);

            const mainDocPath = path.join(outputDir, "main.docx");
            const contentDocPath = path.join(outputDir, "content.docx");
            const mergedDocPath = path.join(outputDir, "final.docx");

            console.log(`\n📄 File: ${fileName}`);
            console.log(`✂️  Split tag: ${splitTag}`);

            try {
                extractFirstDoc(inputFile, mainDocPath, splitTag);
                extractRemainingDoc(inputFile, contentDocPath, splitTag);

                mergeDocx({
                    baseDocxPath: mainDocPath,
                    appendDocxPath: contentDocPath,
                    outputPath: mergedDocPath,
                });

                console.log(`✅ Done → ${outputDir}`);
            } catch (err) {
                console.error(`❌ Failed for ${fileName} | ${splitTag}`);
                console.error(err);
            }
        }
    }
}

/* ============================
   START
============================ */

run();
