import * as fs from "fs";
import PizZip from "pizzip";
import { DOMParser, XMLSerializer } from "@xmldom/xmldom";
import * as xpath from "xpath";

/* ===================== Namespaces ===================== */
const NS = {
  w: "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
};

/* XPath selector with namespaces */
const select = xpath.useNamespaces(NS);

/* ===================== Types ===================== */
type PageBreakRule = {
  name: string;
  xpath: string;
};

type PageBreakMatch = {
  ruleName: string;
  node: Node;
  xml: string;
};

/* ===================== Page-break XPath rules (priority order) ===================== */
const PAGE_BREAK_RULES: PageBreakRule[] =[  
  {
    name: "manual-page-break",
    xpath: "//w:br[@w:type='page']",
  },
  // Page break before paragraph (enabled or default)
  {
    name: "page-break-before",
    // Match any <w:pageBreakBefore> regardless of w:val attribute (present or not)
    xpath: "//w:pageBreakBefore",
  },
  // Section break forcing next page (common in Word)
  {
    name: "section-break-nextPage",
    xpath: "//w:sectPr[w:type/@w:val='nextPage']",
  },
  // Section break forcing odd page
  {
    name: "section-break-oddPage",
    xpath: "//w:sectPr[w:type/@w:val='oddPage']",
  },
  // Section break forcing even page
  {
    name: "section-break-evenPage",
    xpath: "//w:sectPr[w:type/@w:val='evenPage']",
  },
  // Last rendered page break tag (Word optimization)
  {
    name: "last-rendered-page-break",
    xpath: "//w:lastRenderedPageBreak",
  },
];

function loadDocumentXml(docxPath: string): Document {
  console.log(`Loading DOCX from: ${docxPath}`);
  const zip = new PizZip(fs.readFileSync(docxPath));
  const xmlFile = zip.file("word/document.xml");

  if (!xmlFile) {
    throw new Error("word/document.xml not found in DOCX");
  }

  const xmlText = xmlFile.asText();

  // Save extracted XML to result.xml
  fs.writeFileSync("result.xml", xmlText, { encoding: "utf8" });
  console.log("✅ Saved extracted XML to result.xml");

  return new DOMParser().parseFromString(xmlText, "text/xml");
}

function serializeNode(node: Node): string {
  const rawXml = new XMLSerializer().serializeToString(node);
  // Remove the xmlns:w namespace declaration from the string
  return rawXml.replace(/\sxmlns:w="[^"]+"/g, "");
}


function findFirstPageBreak(doc: Document): PageBreakMatch | null {
  for (const rule of PAGE_BREAK_RULES) {
    const nodes = select(rule.xpath, doc) as Node[];

    if (nodes.length > 0) {
      return {
        ruleName: rule.name,
        node: nodes[0],
        xml: serializeNode(nodes[0]),
      };
    }
  }

  return null;
}

function findAllPageBreaks(doc: Document): PageBreakMatch[] {
  const results: PageBreakMatch[] = [];
  const seen = new Set<string>();

  for (const rule of PAGE_BREAK_RULES) {
    const nodes = select(rule.xpath, doc) as Node[];

    for (const node of nodes) {
      const xml = serializeNode(node);

      if (!seen.has(xml)) {
        seen.add(xml);
        results.push({
          ruleName: rule.name,
          node,
          xml,
        });
      }
    }
  }

  return results;
}


function example() {
  try {
    const docXml = loadDocumentXml("./book.docx");

    const firstBreak = findFirstPageBreak(docXml);
    if (firstBreak) {
      // console.log("✅ First page break found:");
      // console.log("Rule:", firstBreak.ruleName);
      // console.log("XML:", firstBreak.xml);
    } else {
      console.log("❌ No page break found");
    }

    const allBreaks = findAllPageBreaks(docXml);
    console.log(`\n🔍 Found ${allBreaks.length} total page break(s):`);
    allBreaks.forEach((b, i) => {
      console.log(`${i + 1}. Rule: ${b.ruleName}`);
      console.log(b.xml);
    });
  } catch (error) {
    console.error("Error reading DOCX:", error);
  }
}

example();
