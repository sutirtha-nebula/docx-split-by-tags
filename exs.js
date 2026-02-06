const fs = require("fs");
const path = require("path");
const PizZip = require("pizzip");
const { DOMParser, XMLSerializer } = require("@xmldom/xmldom");
const xpath = require("xpath");

function extractDocxFlow(inputPath, outDir = "flow_output") {
  const nodesDir = path.join(outDir, "nodes");
  const imagesDir = path.join(outDir, "images");

  [outDir, nodesDir, imagesDir].forEach(d => {
    if (!fs.existsSync(d)) fs.mkdirSync(d, { recursive: true });
  });

  const zip = new PizZip(fs.readFileSync(inputPath));
  const parser = new DOMParser();
  const serializer = new XMLSerializer();

  const docXml = parser.parseFromString(zip.file("word/document.xml").asText(), "text/xml");

  const select = xpath.useNamespaces({
    w: "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    a: "http://schemas.openxmlformats.org/drawingml/2006/main",
    r: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
  });

  const body = select("//w:body", docXml)[0];

  /* ===== IMAGE RELATIONSHIPS ===== */
  const relsXml = parser.parseFromString(
    zip.file("word/_rels/document.xml.rels").asText(), "text/xml"
  );

  const relNodes = xpath.select("//*[local-name()='Relationship']", relsXml);
  const relMap = {};
  relNodes.forEach(r => {
    const id = r.getAttribute("Id");
    const target = r.getAttribute("Target");
    if (target && target.startsWith("media/")) relMap[id] = target;
  });

  let blockIndex = 0;
  let imageIndex = 0;
  const flow = [];

  Array.from(body.childNodes).forEach(node => {
    if (node.nodeType !== 1) return;

    if (node.nodeName === "w:p" || node.nodeName === "w:tbl") {
      const cloned = node.cloneNode(true);

      /* ===== Extract text from this block ===== */
      const textNodes = select(".//w:t", cloned);
      const textContent = textNodes
        .map(t => (t.firstChild ? t.firstChild.data : ""))
        .join(" ");

      /* ===== Extract images inside block ===== */
      const blips = select(".//a:blip", cloned);
      blips.forEach(blip => {
        const rId = blip.getAttribute("r:embed");
        const imgPath = relMap[rId];
        if (imgPath && zip.file(`word/${imgPath}`)) {
          const buffer = zip.file(`word/${imgPath}`).asNodeBuffer();
          const ext = path.extname(imgPath);
          const name = `img_${imageIndex++}${ext}`;
          fs.writeFileSync(path.join(imagesDir, name), buffer);
        }
      });

      /* ===== Save XML ===== */
      const fileName = `block_${blockIndex}.xml`;
      fs.writeFileSync(
        path.join(nodesDir, fileName),
        serializer.serializeToString(cloned)
      );

      /* ===== Save flow info ===== */
      
      const hasTocField =
  select(".//w:instrText", node).some(n =>
    n.firstChild?.data.includes("TOC")
  );

flow.push({
  index: blockIndex,
  type: hasTocField ? "toc" : node.nodeName === "w:p" ? "paragraph" : "table",
  locked: hasTocField,
  file: fileName
});

      blockIndex++;
    }
  });

  fs.writeFileSync(
    path.join(outDir, "flow.json"),
    JSON.stringify(flow, null, 2)
  );

  console.log(`✅ Blocks: ${blockIndex}`);
  console.log(`🖼 Images: ${imageIndex}`);
}

extractDocxFlow("./templates/Recipe.docx");
