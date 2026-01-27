import * as fs from "fs";
import PizZip from "pizzip";
import { DOMParser, XMLSerializer } from "xmldom";
import * as xpath from "xpath";

/* ===================== TYPES ===================== */

type Zip = InstanceType<typeof PizZip>;

interface MergeDocxOptions {
    baseDocxPath: string;
    appendDocxPath: string;
    outputPath: string;
}

interface ImageRef {
    node: Element;
    attr: string;
    id: string;
}

/* ===================== NAMESPACES ===================== */

const NS = {
    w: "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    a: "http://schemas.openxmlformats.org/drawingml/2006/main",
    r: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    pic: "http://schemas.openxmlformats.org/drawingml/2006/picture",
    v: "urn:schemas-microsoft-com:vml",
    wp: "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    rel: "http://schemas.openxmlformats.org/package/2006/relationships",
};

const select: any = xpath.useNamespaces(NS);

/* ===================== ZIP / XML HELPERS ===================== */

function loadDocx(path: string): Zip {
    return new PizZip(fs.readFileSync(path));
}

function parseXml(zip: Zip, path: string): Document {
    return new DOMParser().parseFromString(
        zip.file(path)!.asText(),
        "text/xml"
    );
}

function writeXml(zip: Zip, path: string, xml: Document): void {
    zip.file(path, new XMLSerializer().serializeToString(xml));
}

function ensureFolder(zip: Zip, folder: string): void {
    if (!zip.files[folder]) {
        zip.file(`${folder}/.placeholder`, "");
    }
}

/* ===================== RELATIONSHIP HELPERS ===================== */

function findAllRelationshipFiles(zip: Zip): string[] {
    return Object.keys(zip.files).filter(
        f => f.includes("_rels") && f.endsWith(".rels")
    );
}

function findRelationship(
    zip: Zip,
    relId: string
): { target: string; type: string } | null {
    for (const relsPath of findAllRelationshipFiles(zip)) {
        try {
            const relsXml = parseXml(zip, relsPath);
            const rels = select("//rel:Relationship", relsXml) as Element[];

            for (const rel of rels) {
                if (rel.getAttribute("Id") === relId) {
                    return {
                        target: rel.getAttribute("Target") || "",
                        type: rel.getAttribute("Type") || "",
                    };
                }
            }
        } catch {
            /* ignore malformed rels */
        }
    }
    return null;
}

function getNextRelId(relsXml: Document): string {
    const rels = select("//rel:Relationship", relsXml) as Element[];
    let max = 0;

    rels.forEach(r => {
        const id = r.getAttribute("Id");
        if (id?.startsWith("rId")) {
            max = Math.max(max, parseInt(id.slice(3), 10));
        }
    });

    return `rId${max + 1}`;
}

/* ===================== IMAGE HANDLING ===================== */

function copyImage(
    srcZip: Zip,
    dstZip: Zip,
    relsXml: Document,
    ref: ImageRef,
    imageCounter: number
): number {
    const rel = findRelationship(srcZip, ref.id);
    if (!rel) return imageCounter;

    const imgPath = rel.target.startsWith("word/")
        ? rel.target
        : `word/${rel.target}`;

    if (!srcZip.file(imgPath)) return imageCounter;

    const ext = imgPath.split(".").pop() || "png";
    const buffer = srcZip.file(imgPath)!.asNodeBuffer();
    const newName = `image${++imageCounter}.${ext}`;

    dstZip.file(`word/media/${newName}`, buffer);

    const newRelId = getNextRelId(relsXml);
    const relNode = relsXml.createElement("Relationship");

    relNode.setAttribute("Id", newRelId);
    relNode.setAttribute(
        "Type",
        rel.type ||
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
    );
    relNode.setAttribute("Target", `media/${newName}`);

    relsXml.documentElement.appendChild(relNode);
    ref.node.setAttribute(ref.attr, newRelId);

    return imageCounter;
}

function processImages(
    node: Node,
    srcZip: Zip,
    dstZip: Zip,
    relsXml: Document,
    imageCounter: number
): number {
    const refs: ImageRef[] = [];

    select(".//a:blip", node).forEach((n: Element) => {
        const id = n.getAttribute("r:embed");
        if (id) refs.push({ node: n, attr: "r:embed", id });
    });

    select(".//v:imagedata", node).forEach((n: Element) => {
        const id = n.getAttribute("r:id");
        if (id) refs.push({ node: n, attr: "r:id", id });
    });

    for (const ref of refs) {
        imageCounter = copyImage(srcZip, dstZip, relsXml, ref, imageCounter);
    }

    return imageCounter;
}

/* ===================== MAIN MERGE FUNCTION ===================== */

export function mergeDocx({
    baseDocxPath,
    appendDocxPath,
    outputPath,
}: MergeDocxOptions): void {
    const baseZip = loadDocx(baseDocxPath);
    const appendZip = loadDocx(appendDocxPath);

    [
        "word/styles.xml",
        "word/numbering.xml",
        "word/_rels/numbering.xml.rels",
    ].forEach(path => {
        if (appendZip.file(path)) {
            baseZip.file(path, appendZip.file(path)!.asText());
        }
    });

    ensureFolder(baseZip, "word/media");

    const baseDocXml = parseXml(baseZip, "word/document.xml");
    const appendDocXml = parseXml(appendZip, "word/document.xml");

    const baseBody = select("//w:body", baseDocXml)[0] as Element;
    const appendBody = select("//w:body", appendDocXml)[0] as Element;

    const relsPath = "word/_rels/document.xml.rels";
    const relsXml = parseXml(baseZip, relsPath);

    let imageCounter = 0;

    Array.from(appendBody.childNodes)
        .filter(n => n.nodeName !== "w:sectPr")
        .forEach(node => {
            const cloned = node.cloneNode(true);
            imageCounter = processImages(
                cloned,
                appendZip,
                baseZip,
                relsXml,
                imageCounter
            );
            baseBody.appendChild(cloned);
        });

    writeXml(baseZip, "word/document.xml", baseDocXml);
    writeXml(baseZip, relsPath, relsXml);

    fs.writeFileSync(
        outputPath,
        baseZip.generate({ type: "nodebuffer" })
    );
}

mergeDocx({
    baseDocxPath: "./main.docx",
    appendDocxPath: "./content.docx",
    outputPath: "./final.docx",
});
