const fs = require("fs");
const path = require("path");
const PizZip = require("pizzip");
const { DOMParser, XMLSerializer } = require("@xmldom/xmldom");
const xpath = require("xpath");

function rebuildDocx(templateDocx, flowDir, outputPath) {
  const zip = new PizZip(fs.readFileSync(templateDocx));
  const parser = new DOMParser();
  const serializer = new XMLSerializer();

  const flow = JSON.parse(fs.readFileSync(path.join(flowDir, "flow.json")));

  const docXml = parser.parseFromString(zip.file("word/document.xml").asText(), "text/xml");
  const relsXml = parser.parseFromString(
    zip.file("word/_rels/document.xml.rels").asText(), 
    "text/xml"
  );

  const select = xpath.useNamespaces({
    w: "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    a: "http://schemas.openxmlformats.org/drawingml/2006/main",
    r: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
  });

  const body = select("//w:body", docXml)[0];
  const sectPr = select("w:sectPr", body)[0];

  // remove old content but keep sectPr
  Array.from(body.childNodes).forEach(n => {
    if (n.nodeName !== "w:sectPr") body.removeChild(n);
  });

  /* ================= RELATIONSHIP COUNTER ================= */
  let maxRel = 0;
  xpath.select("//*[local-name()='Relationship']", relsXml).forEach(r => {
    const id = r.getAttribute("Id");
    if (id?.startsWith("rId")) maxRel = Math.max(maxRel, parseInt(id.slice(3)));
  });

  function newRelId() {
    return `rId${++maxRel}`;
  }

  /* ================= COPY ORIGINAL HYPERLINK RELS (STEP 3) ================= */
  const templateRelsXml = parser.parseFromString(
    zip.file("word/_rels/document.xml.rels").asText(),
    "text/xml"
  );

  copyHyperlinks(templateRelsXml, relsXml);

  /* ================= APPEND BLOCKS ================= */
  const imageFiles = fs.readdirSync(path.join(flowDir, "images"));
  let imgCursor = 0;

  flow.forEach(block => {
    const xml = fs.readFileSync(path.join(flowDir, "nodes", block.file), "utf8");
    const blockDoc = parser.parseFromString(xml, "text/xml");
    const imported = docXml.importNode(blockDoc.documentElement, true);

    if (block.type !== "toc") {

      /* ===== FIX IMAGES ===== */
      const blips = select(".//a:blip", imported);
      blips.forEach(blip => {
        const newId = newRelId();
        blip.setAttribute("r:embed", newId);

        const imgName = imageFiles[imgCursor++];
        if (!imgName) return;

        const imgPath = `media/${imgName}`;

        const relNode = relsXml.createElement("Relationship");
        relNode.setAttribute("Id", newId);
        relNode.setAttribute(
          "Type",
          "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
        );
        relNode.setAttribute("Target", imgPath);
        relsXml.documentElement.appendChild(relNode);

        const buffer = fs.readFileSync(path.join(flowDir, "images", imgName));
        zip.file(`word/${imgPath}`, buffer);
      });
    }

    body.insertBefore(imported, sectPr);
  });

  zip.file("word/document.xml", serializer.serializeToString(docXml));
  zip.file("word/_rels/document.xml.rels", serializer.serializeToString(relsXml));

  enableFieldUpdate(zip, parser, serializer);

  fs.writeFileSync(outputPath, zip.generate({ type: "nodebuffer" }));
}


function copyHyperlinks(srcRelsXml, dstRelsXml) {
  const existingIds = new Set(
    xpath.select("//*[local-name()='Relationship']", dstRelsXml)
      .map(r => r.getAttribute("Id"))
  );

  const links = xpath.select(
    "//*[local-name()='Relationship'][contains(@Type,'hyperlink')]",
    srcRelsXml
  );

  links.forEach(link => {
    const id = link.getAttribute("Id");
    if (existingIds.has(id)) return; // avoid duplicates

    const relNode = dstRelsXml.createElement("Relationship");
    relNode.setAttribute("Id", id);
    relNode.setAttribute("Type", link.getAttribute("Type"));
    relNode.setAttribute("Target", link.getAttribute("Target"));
    relNode.setAttribute("TargetMode", "External");
    dstRelsXml.documentElement.appendChild(relNode);
  });
}


function enableFieldUpdate(zip, parser, serializer) {
  const settingsPath = "word/settings.xml";
  const settingsXml = parser.parseFromString(zip.file(settingsPath).asText(), "text/xml");

  const node = settingsXml.createElement("w:updateFields");
  node.setAttribute("w:val", "true");
  settingsXml.documentElement.appendChild(node);

  zip.file(settingsPath, serializer.serializeToString(settingsXml));
}

rebuildDocx("./templates/Recipe.docx", "flow_output", "rebuilt.docx");
