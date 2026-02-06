const fs = require("fs");
const path = require("path");
const PizZip = require("pizzip");

/* ============================
   CONFIG
============================ */
const SOURCE_DOCX = "./final_result_2.docx";
const NODES_DIR = "./nodes-b";
const OUTPUT_DOCX = "./rebuild.docx";

/* ============================
   Load node XML fragments (NO SORT)
============================ */
function loadNodeFragments() {
  let combinedXml = "";
  let logLines = [];

  const files = fs
    .readdirSync(NODES_DIR)
    .filter(f => /^node_\d+\.xml$/.test(f))
    .sort((a, b) => {
      const ai = parseInt(a.match(/\d+/)[0], 10);
      const bi = parseInt(b.match(/\d+/)[0], 10);
      return ai - bi;
    });

  for (const file of files) {
    console.log(`Processing node file: ${file}`);
    logLines.push(`Processed: ${file}`);

    const filePath = path.join(NODES_DIR, file);
    const content = fs.readFileSync(filePath, "utf8").trim();

    if (!content) {
      logLines.push(`Skipped empty file: ${file}`);
      continue;
    }

    // Append node XML as-is
    combinedXml += content + "\n";
  }

  // Write processing log (filenames only)
  const logPath = path.join(NODES_DIR, "_node_processing.log");
  fs.writeFileSync(logPath, logLines.join("\n"), "utf8");

  console.log(`Node processing log written to: ${logPath}`);

  return combinedXml;
}




/* ============================
   Replace <w:body> content
============================ */
function replaceBodyAsString(documentXml, bodyContent) {
  const bodyRegex = /<w:body[^>]*>[\s\S]*?<\/w:body>/;

  if (!bodyRegex.test(documentXml)) {
    throw new Error("w:body not found");
  }

  return documentXml.replace(
    bodyRegex,
    match => {
      const openTag = match.match(/<w:body[^>]*>/)[0];
      return `${openTag}\n${bodyContent}\n</w:body>`;
    }
  );
}

/* ============================
   Main
============================ */
function buildDocx() {
  const zip = new PizZip(fs.readFileSync(SOURCE_DOCX));
  const docPath = "word/document.xml";

  const documentXml = zip.file(docPath).asText();
  const bodyContent = loadNodeFragments();

  console.log(bodyContent); // Debug: show loaded fragments

  const updatedXml = replaceBodyAsString(documentXml, bodyContent);

  zip.file(docPath, updatedXml);

  fs.writeFileSync(
    OUTPUT_DOCX,
    zip.generate({ type: "nodebuffer" })
  );

  console.log("✅ DOCX generated:", OUTPUT_DOCX);
}

buildDocx();
