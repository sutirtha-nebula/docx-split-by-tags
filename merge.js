"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.mergeDocx = mergeDocx;
var fs = require("fs");
var pizzip_1 = require("pizzip");
var xmldom_1 = require("xmldom");
var xpath = require("xpath");
/* ===================== NAMESPACES ===================== */
var NS = {
    w: "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    a: "http://schemas.openxmlformats.org/drawingml/2006/main",
    r: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    pic: "http://schemas.openxmlformats.org/drawingml/2006/picture",
    v: "urn:schemas-microsoft-com:vml",
    wp: "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    rel: "http://schemas.openxmlformats.org/package/2006/relationships",
};
var select = xpath.useNamespaces(NS);
/* ===================== ZIP / XML HELPERS ===================== */
function loadDocx(path) {
    return new pizzip_1.default(fs.readFileSync(path));
}
function parseXml(zip, path) {
    return new xmldom_1.DOMParser().parseFromString(zip.file(path).asText(), "text/xml");
}
function writeXml(zip, path, xml) {
    zip.file(path, new xmldom_1.XMLSerializer().serializeToString(xml));
}
function ensureFolder(zip, folder) {
    if (!zip.files[folder]) {
        zip.file("".concat(folder, "/.placeholder"), "");
    }
}
/* ===================== RELATIONSHIP HELPERS ===================== */
function findAllRelationshipFiles(zip) {
    return Object.keys(zip.files).filter(function (f) { return f.includes("_rels") && f.endsWith(".rels"); });
}
function findRelationship(zip, relId) {
    for (var _i = 0, _a = findAllRelationshipFiles(zip); _i < _a.length; _i++) {
        var relsPath = _a[_i];
        try {
            var relsXml = parseXml(zip, relsPath);
            var rels = select("//rel:Relationship", relsXml);
            for (var _b = 0, rels_1 = rels; _b < rels_1.length; _b++) {
                var rel = rels_1[_b];
                if (rel.getAttribute("Id") === relId) {
                    return {
                        target: rel.getAttribute("Target") || "",
                        type: rel.getAttribute("Type") || "",
                    };
                }
            }
        }
        catch (_c) {
            /* ignore malformed rels */
        }
    }
    return null;
}
function getNextRelId(relsXml) {
    var rels = select("//rel:Relationship", relsXml);
    var max = 0;
    rels.forEach(function (r) {
        var id = r.getAttribute("Id");
        if (id === null || id === void 0 ? void 0 : id.startsWith("rId")) {
            max = Math.max(max, parseInt(id.slice(3), 10));
        }
    });
    return "rId".concat(max + 1);
}
/* ===================== IMAGE HANDLING ===================== */
function copyImage(srcZip, dstZip, relsXml, ref, imageCounter) {
    var rel = findRelationship(srcZip, ref.id);
    if (!rel)
        return imageCounter;
    var imgPath = rel.target.startsWith("word/")
        ? rel.target
        : "word/".concat(rel.target);
    if (!srcZip.file(imgPath))
        return imageCounter;
    var ext = imgPath.split(".").pop() || "png";
    var buffer = srcZip.file(imgPath).asNodeBuffer();
    var newName = "image".concat(++imageCounter, ".").concat(ext);
    dstZip.file("word/media/".concat(newName), buffer);
    var newRelId = getNextRelId(relsXml);
    var relNode = relsXml.createElement("Relationship");
    relNode.setAttribute("Id", newRelId);
    relNode.setAttribute("Type", rel.type ||
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image");
    relNode.setAttribute("Target", "media/".concat(newName));
    relsXml.documentElement.appendChild(relNode);
    ref.node.setAttribute(ref.attr, newRelId);
    return imageCounter;
}
function processImages(node, srcZip, dstZip, relsXml, imageCounter) {
    var refs = [];
    select(".//a:blip", node).forEach(function (n) {
        var id = n.getAttribute("r:embed");
        if (id)
            refs.push({ node: n, attr: "r:embed", id: id });
    });
    select(".//v:imagedata", node).forEach(function (n) {
        var id = n.getAttribute("r:id");
        if (id)
            refs.push({ node: n, attr: "r:id", id: id });
    });
    for (var _i = 0, refs_1 = refs; _i < refs_1.length; _i++) {
        var ref = refs_1[_i];
        imageCounter = copyImage(srcZip, dstZip, relsXml, ref, imageCounter);
    }
    return imageCounter;
}
/* ===================== MAIN MERGE FUNCTION ===================== */
function mergeDocx(_a) {
    var baseDocxPath = _a.baseDocxPath, appendDocxPath = _a.appendDocxPath, outputPath = _a.outputPath;
    var baseZip = loadDocx(baseDocxPath);
    var appendZip = loadDocx(appendDocxPath);
    [
        "word/styles.xml",
        "word/numbering.xml",
        "word/_rels/numbering.xml.rels",
    ].forEach(function (path) {
        if (appendZip.file(path)) {
            baseZip.file(path, appendZip.file(path).asText());
        }
    });
    ensureFolder(baseZip, "word/media");
    var baseDocXml = parseXml(baseZip, "word/document.xml");
    var appendDocXml = parseXml(appendZip, "word/document.xml");
    var baseBody = select("//w:body", baseDocXml)[0];
    var appendBody = select("//w:body", appendDocXml)[0];
    var relsPath = "word/_rels/document.xml.rels";
    var relsXml = parseXml(baseZip, relsPath);
    var imageCounter = 0;
    Array.from(appendBody.childNodes)
        .filter(function (n) { return n.nodeName !== "w:sectPr"; })
        .forEach(function (node) {
        var cloned = node.cloneNode(true);
        imageCounter = processImages(cloned, appendZip, baseZip, relsXml, imageCounter);
        baseBody.appendChild(cloned);
    });
    writeXml(baseZip, "word/document.xml", baseDocXml);
    writeXml(baseZip, relsPath, relsXml);
    fs.writeFileSync(outputPath, baseZip.generate({ type: "nodebuffer" }));
}
mergeDocx({
    baseDocxPath: "./main.docx",
    appendDocxPath: "./content.docx",
    outputPath: "./final.docx",
});
