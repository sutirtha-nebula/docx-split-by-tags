const fs = require("fs");
const PizZip = require("pizzip");

// Files
const INPUT_DOCX = "./output.docx";
const OUTPUT_XML = "document.xml";

console.log("Extracting document.xml...");

// Read DOCX
const content = fs.readFileSync(INPUT_DOCX, "binary");
const zip = new PizZip(content);

// Ensure document.xml exists
const docFile = zip.file("word/document.xml");
if (!docFile) {
  throw new Error("word/document.xml not found in DOCX");
}

// Extract XML
const documentXml = docFile.asText();

// Save XML
fs.writeFileSync(OUTPUT_XML, documentXml, "utf8");

console.log("document.xml extracted successfully");
