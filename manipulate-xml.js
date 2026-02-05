import fs from "fs";
import path from "path";
import PizZip from "pizzip";

/**
 * Replace a specific XML file inside a .docx with a modified one and create a new .docx file.
 * @param {string} originalDocxPath Path to original .docx file
 * @param {string} modifiedXmlPath Path to the modified XML file
 * @param {string} xmlToReplacePath Path inside docx zip to the XML file to replace (e.g. 'word/document.xml')
 * @param {string} outputDocxPath Path to save the modified .docx file
 */
async function replaceXmlInDocx(originalDocxPath, modifiedXmlPath, xmlToReplacePath, outputDocxPath) {
  // Read original docx as binary
  const content = fs.readFileSync(originalDocxPath);
  const zip = new PizZip(content);

  // Read modified XML content as UTF-8 text
  const modifiedXml = fs.readFileSync(modifiedXmlPath, "utf8");

  // Replace the XML inside the zip archive
  zip.file(xmlToReplacePath, modifiedXml);

  // Generate new docx file content (binary)
  const newDocxContent = zip.generate({ type: "nodebuffer" });

  // Write new .docx file
  fs.writeFileSync(outputDocxPath, newDocxContent);

  console.log(`New .docx saved at: ${outputDocxPath}`);
}

// Example usage:
const originalDocx = "./templates/book.docx";        // Your original docx file path
const modifiedXml = "./modified_document.xml"; // Your modified XML file path
const xmlInsideDocx = "word/document.xml";     // XML part inside the docx to replace
const outputDocx = "./modified.docx";           // Output modified docx path

replaceXmlInDocx(originalDocx, modifiedXml, xmlInsideDocx, outputDocx);
