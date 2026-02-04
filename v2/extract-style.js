const fs = require("fs");
const PizZip = require("pizzip");

// Files
const INPUT_DOCX = "./output.docx";
const OUTPUT_XML = "styles.xml";

console.log("Extracting styles.xml...");

// Read DOCX
const content = fs.readFileSync(INPUT_DOCX, "binary");
const zip = new PizZip(content);

// Ensure styles.xml exists
const stylesFile = zip.file("word/styles.xml");
if (!stylesFile) {
  throw new Error("word/styles.xml not found in DOCX");
}

// Extract XML
const stylesXml = stylesFile.asText();

// Save XML
fs.writeFileSync(OUTPUT_XML, stylesXml, "utf8");

console.log("styles.xml extracted successfully");
