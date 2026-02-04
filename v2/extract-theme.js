const fs = require("fs");
const PizZip = require("pizzip");

// Files
const INPUT_DOCX = "./output.docx";
const OUTPUT_XML = "theme1.xml";

console.log("Extracting theme1.xml...");

// Read DOCX
const content = fs.readFileSync(INPUT_DOCX, "binary");
const zip = new PizZip(content);

// Ensure theme1.xml exists
const themeFile = zip.file("word/theme/theme1.xml");
if (!themeFile) {
  throw new Error("word/theme/theme1.xml not found in DOCX");
}

// Extract XML
const themeXml = themeFile.asText();

// Save XML
fs.writeFileSync(OUTPUT_XML, themeXml, "utf8");

console.log("theme1.xml extracted successfully");
