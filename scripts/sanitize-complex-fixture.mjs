import { mkdirSync, readFileSync, writeFileSync } from "node:fs";
import { dirname, resolve } from "node:path";
import JSZip from "jszip";

const sourcePath =
  process.env.PPTX_COMPLEX_SOURCE ||
  "/Users/sid/Documents/Coding/An Ecosystem of Support for the Fourth Trimester (1).pptx";
const outputPath = resolve("test/fixtures/real/complex-sanitized.pptx");

const transparentPngBase64 =
  "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Wn4n8kAAAAASUVORK5CYII=";
const transparentJpegBase64 =
  "/9j/4AAQSkZJRgABAQAAAQABAAD/2wCEAAkGBxAQEBAQEA8QEA8PEA8QEA8QDw8QEA8PFREWFhURFRUYHSggGBolGxUVITEhJSkrLi4uFx8zODMsNygtLisBCgoKDg0OGhAQGi0lHyUtLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLf/AABEIAAEAAgMBIgACEQEDEQH/xAAXAAEAAwAAAAAAAAAAAAAAAAAAAQID/8QAFBABAAAAAAAAAAAAAAAAAAAAAP/aAAwDAQACEAMQAAAB6A//xAAVEQEBAAAAAAAAAAAAAAAAAAAAEf/aAAgBAQABBQJf/8QAFBEBAAAAAAAAAAAAAAAAAAAAEP/aAAgBAwEBPwF//8QAFBEBAAAAAAAAAAAAAAAAAAAAEP/aAAgBAgEBPwF//8QAFBABAAAAAAAAAAAAAAAAAAAAEP/aAAgBAQAGPwJf/8QAFBABAAAAAAAAAAAAAAAAAAAAEP/aAAgBAQABPyF//9k=";
const placeholderSvg = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 64 64"><rect width="64" height="64" fill="#E5E7EB"/><path d="M10 46l12-14 10 9 10-14 12 19H10z" fill="#9CA3AF"/></svg>`;

const CORE_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>OpenRenderKit Sanitized Regression Deck</dc:title>
  <dc:subject>Sanitized PPTX regression fixture</dc:subject>
  <dc:creator>OpenRenderKit</dc:creator>
  <cp:keywords>sanitized,fixture,regression</cp:keywords>
  <dc:description>Sanitized regression fixture derived from a complex source deck.</dc:description>
  <cp:lastModifiedBy>OpenRenderKit</cp:lastModifiedBy>
</cp:coreProperties>`;

const CUSTOM_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"/>`;

const LABEL_INFO_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<LabelList xmlns="http://schemas.microsoft.com/office/2020/mipLabelMetadata"/>`;

let tokenCounter = 1;

const CURATED_TEXT_BY_PATH = {
  "ppt/slides/slide1.xml": [
    "OpenBridge",
    ":",
    " ",
    "A Connected Support Network for the Care Journey",
    "Team:",
    "Alex Morgan",
    "Priya ",
    "Nair",
    "Jordan Lee ",
    "Mateo",
    " ",
    "Chen",
    " Rivera",
    "Brooks",
    " Quinn ",
    "Rivera",
    " Malik",
  ],
  "ppt/slides/slide2.xml": [
    "The Challenge:",
    "The Care Continuity Gap",
    "– A Coordination Breakdown",
  ],
  "ppt/slides/slide3.xml": [
    "Our Approach:",
    "The OpenBridge Network",
    "™",
    "A connected system for ongoing care:",
    "Companion App:",
    "A conversational guide offering daily check-ins, practical prompts, and reflective support.",
    "Signal Pebble:",
    "A low-friction handheld input for capturing non-verbal wellbeing signals.",
    "Care Summary Agent:",
    "Turns incoming activity into concise updates for practitioners and support teams.",
    "OpenBridge Network™:",
    "A modular platform designed for future integrations and broader care workflows.",
  ],
  "ppt/slides/slide4.xml": [
    "Evaluation Matrix",
    "A single app or gadget only solves part of the experience. The",
    "OpenBridge Network™",
    "was the only option that addressed the full set of user needs:",
    "continuous support, private reflection, objective signals, and handoff into existing care workflows.",
    "Solution Options",
    "Continuous",
    "Support",
    "Reduces Friction",
    "Signal Quality",
    "Team Workflow Integration",
    "Feasibility",
    "Conversation Support App",
    "High",
    "High",
    "Low",
    "Medium",
    "High",
    "Wearable Stress Sensor",
    "High",
    "Medium",
    "High",
    "Low",
    "Medium",
    "Expanded Telehealth Service",
    "Low",
    "Low",
    "Low",
    "High",
    "High",
    "Community Support Network",
    "Medium",
    "High",
    "Low",
    "Low",
    "High",
    "OpenBridge Network (App + Device)",
    "High",
    "High",
    "High",
    "High",
    "Medium",
    "We reviewed multiple solution directions against the key project constraints:",
  ],
  "ppt/slides/slide5.xml": [
    "Landscape Review:",
    " ",
    "Why Existing Offers Miss the Mark",
    "────────────────────",
    "Most products solve one slice of the problem.",
    "OpenBridge combines the missing pieces.",
    "Mindfulness & Tracking Apps (e.g.,",
    "CalmPath",
    ", DailyFocus)",
    "Wearable Signal Devices (e.g.,",
    "Pulse Band, Quiet Ring)",
    "On-Demand Care Platforms (e.g.,",
    "CareLine",
    ",",
    "TheraNow",
    ")",
    "(+) Strength:",
    " ",
    "Easy access to guided exercises and lightweight reflection tools.",
    "(+) Strength:",
    " ",
    "Provides continuous physiological signals such as stress, sleep, or activity.",
    "(+) Strength:",
    " ",
    "Convenient access to licensed professionals when an appointment is needed.",
    "(-) Weakness:",
    " ",
    "Limited workflow integration and mostly dependent on subjective inputs.",
    "(-) Weakness:",
    " ",
    "Signal data arrives without context and without a coordinated response path.",
    "(-) Weakness:",
    " ",
    "Reactive and appointment-based, with limited support between sessions.",
  ],
  "ppt/slides/slide6.xml": [
    "Why OpenBridge is the Stronger Choice.",
    "The",
    "OpenBridge Network™",
    "fits the project constraints because it delivers across the full workflow:",
    "Personalised & Proactive Support:",
    "Reduces",
    "drop-off",
    "with ongoing support instead of only event-based follow-up.",
    "Deeper Signal Insight:",
    "Combines interaction patterns and device signals to",
    "surface subtle changes",
    "that surveys and standalone apps often miss.",
    "Seamless Team Integration:",
    "Care teams receive",
    "clear, action-ready summaries",
    ", saving time and improving continuity.",
    "An Open Ecosystem:",
    "A platform built for future growth,",
    "allowing integration with third-party tools",
    "such as wearables, records, and communication services.",
  ],
  "ppt/slides/slide7.xml": [
    "Road to Prototype:",
    " ",
    "Software Inputs",
  ],
  "ppt/slides/slide8.xml": [
    "Road to Prototype:",
    "Skills & Resources",
    "Skills Required",
    "Resources Required",
    "Applied AI Engineering:",
    " ",
    "Model orchestration",
    "Prompt design, safety rules, and response evaluation.",
    "Hardware:",
    "Prototype board, microphone, speaker, compact display, microcontroller, pressure sensor, rechargeable battery, and enclosure parts.",
    "Working with embedded systems:",
    "Firmware basics, hardware debugging, and sensor integration for a lightweight device.",
    "Software & APIs:",
    "Realtime model API,",
    "hosting platform",
    "for deployment, database storage, email delivery, and report generation services.",
    "Product Design & 3D",
    "Modeling",
    ":",
    "Enclosure design, interaction design, and user-facing flows.",
    "Prototyping:",
    "3D printer and bench tools.",
    "Full-Stack Web Development:",
    "React, Next.js,",
    "TypeScript",
    ".",
    "Domain Expertise:",
    "Research in longitudinal care, emotional wellbeing, and service delivery validation.",
  ],
  "ppt/slides/slide9.xml": [
    "The Key Risk:",
    "Unhealthy Reliance",
  ],
  "ppt/slides/slide10.xml": ["Who benefits?", " "],
  "ppt/slides/slide11.xml": [
    "Build the Network.",
    "OpenBridge",
    "is more than a feature set; it is a coordination layer for a more responsive, more human model of care.",
  ],
  "ppt/slides/slide12.xml": [
    "References",
    "Northfield Research, “Designing continuity in digital care pathways,” Open Systems Review, 2023.",
    "Harbor Institute",
    ", “Support gaps across extended recovery journeys,” 2022.",
    "Center for Responsible Interfaces, “Practical guidance for assistive AI systems,” 2024.",
    "L. Mercer, “When automated guidance becomes over-attached,” The Civic Ledger, Mar. 14, 2024.",
    "R. Kim, et al., “Comparing transitional care models across public health systems,” Journal of Service Design, vol. 18, no. 4, 2024.",
    "T. Vale, “Trust, restraint, and the risks of agreeable assistants,”",
    "Policy Signals",
    ", Feb. 12, 2024.",
  ],
  "ppt/diagrams/data1.xml": [
    "1.",
    "Continuity Gaps in Extended Care:",
    "Affect",
    "many people after high-touch services end",
    "and often create fragmented follow-up.",
    "2.",
    "An Event-Centered Model:",
    "Built around a single milestone",
    "the immediate checkpoint",
    "instead of",
    "the longer transition that follows",
    "when needs continue to change.",
    "3.",
    "The Follow-Through Gap:",
    "Intensive support fades after the initial phase,",
    "leaving many people without clear next steps.",
    "4.",
    "The Outcome:",
    "A reactive system that increases uncertainty",
    "exactly when timely support matters most.",
  ],
  "ppt/diagrams/data2.xml": [
    "The Risk:",
    "An overly deferential,",
    "always-agreeing",
    "assistant can amplify poor judgment and encourage unhealthy dependence.",
    "Our Mitigation Strategy:",
    "A strict human-in-the-loop review model.",
    "The assistant is designed to",
    "redirect",
    ", not simply affirm.",
    "It maintains",
    "clear boundaries",
    "as a support and coordination tool.",
    "A",
    "qualified reviewer",
    "remains accountable for decisions and escalation.",
  ],
  "ppt/diagrams/data3.xml": [
    "Primary Stakeholders:",
    "Individuals & Support Partners",
    "s",
    "eeking timely, low-friction guidance.",
    "Care Teams",
    "n",
    "eeding efficient ways to review patterns and intervene earlier.",
    "Secondary Stakeholders:",
    "Providers & Payers",
    "l",
    "ooking for scalable approaches that improve outcomes and reduce avoidable cost.",
    "Community & Advocacy Groups:",
    "(e.g., local networks and peer-led organizations).",
    "Market Opportunity:",
    "Where transitions are long and fragmented, the need for a coordinated and scalable support network is substantial.",
  ],
};

function nextToken(prefix = "TXT") {
  const token = `${prefix}-${String(tokenCounter).padStart(4, "0")}`;
  tokenCounter += 1;
  return token;
}

function escapeXmlText(value) {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

function sanitizeTextNodes(xml, path) {
  const curated = CURATED_TEXT_BY_PATH[path];
  if (curated) {
    let curatedIndex = 0;
    xml = xml.replace(/(<(?:a:)?t(?: xml:space="preserve")?>)[\s\S]*?(<\/(?:a:)?t>)/g, (_match, start, end) => {
      const text = escapeXmlText(curated[curatedIndex] ?? nextToken("TXT"));
      curatedIndex += 1;
      return `${start}${text}${end}`;
    });
  } else {
    xml = xml.replace(
      /(<(?:a:)?t(?: xml:space="preserve")?>)[\s\S]*?(<\/(?:a:)?t>)/g,
      (_match, start, end) => `${start}${escapeXmlText(nextToken("TXT"))}${end}`,
    );
  }

  return xml
    .replace(/<(?:a:)?fld\b([^>]*)>([\s\S]*?)<\/(?:a:)?fld>/g, (_match, attrs, inner) => {
      const updatedInner = inner.replace(
        /<(?:a:)?t>([\s\S]*?)<\/(?:a:)?t>/g,
        () => `<a:t>${nextToken("FLD")}</a:t>`,
      );
      return `<a:fld${attrs}>${updatedInner}</a:fld>`;
    })
    .replace(/<(dc:title|dc:subject|dc:creator|dc:description|cp:keywords|cp:lastModifiedBy)>([\s\S]*?)<\/\1>/g, (_match, tagName) => {
      return `<${tagName}>${nextToken("META")}</${tagName}>`;
    });
}

function sanitizeCustomXml(xml) {
  return xml.replace(/>([^<]+)</g, (match, value) => {
    if (!value.trim()) return match;
    return `>${nextToken("DATA")}<`;
  });
}

function replaceBinaryForPath(path) {
  const lower = path.toLowerCase();
  if (lower.endsWith(".png")) {
    return Buffer.from(transparentPngBase64, "base64");
  }
  if (lower.endsWith(".jpg") || lower.endsWith(".jpeg")) {
    return Buffer.from(transparentJpegBase64, "base64");
  }
  if (lower.endsWith(".svg")) {
    return placeholderSvg;
  }
  return null;
}

async function main() {
  const input = readFileSync(sourcePath);
  const zip = await JSZip.loadAsync(input);
  const output = new JSZip();

  for (const [path, file] of Object.entries(zip.files)) {
    if (file.dir) {
      continue;
    }

    if (path === "docProps/core.xml") {
      output.file(path, CORE_XML);
      continue;
    }

    if (path === "docProps/custom.xml") {
      output.file(path, CUSTOM_XML);
      continue;
    }

    if (path === "docMetadata/LabelInfo.xml") {
      output.file(path, LABEL_INFO_XML);
      continue;
    }

    if (path.startsWith("customXml/") && path.endsWith(".xml")) {
      output.file(path, sanitizeCustomXml(await file.async("string")));
      continue;
    }

    if (path.endsWith(".xml")) {
      output.file(path, sanitizeTextNodes(await file.async("string"), path));
      continue;
    }

    const replacement = replaceBinaryForPath(path);
    if (replacement !== null) {
      output.file(path, replacement);
      continue;
    }

    output.file(path, await file.async("nodebuffer"));
  }

  mkdirSync(dirname(outputPath), { recursive: true });
  const sanitized = await output.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
  writeFileSync(outputPath, sanitized);
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
