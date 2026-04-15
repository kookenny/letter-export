/**
 * generate_docx.js
 * ────────────────
 * Converts structured JSON (from letter_extract.py) into a formatted Word
 * document using docx-js.
 *
 * Usage: node tools/generate_docx.js <input.json> <output.docx>
 */

const fs = require("fs");
const cheerio = require("cheerio");
const {
  Document,
  Packer,
  Paragraph,
  TextRun,
  Table,
  TableRow,
  TableCell,
  HeadingLevel,
  AlignmentType,
  LevelFormat,
  BorderStyle,
  WidthType,
  ShadingType,
  PageBreak,
  PageOrientation,
  Header,
  Footer,
  PageNumber,
} = require("docx");

// ── CONSTANTS ───────────────────────────────────────────────────────────────

// US Letter landscape: pass portrait dims, docx-js swaps internally
const PAGE_WIDTH = 12240; // short edge (8.5 inches)
const PAGE_HEIGHT = 15840; // long edge (11 inches)
const MARGIN = 1440; // 1 inch
// In landscape, content width uses the long edge minus margins
const CONTENT_WIDTH = PAGE_HEIGHT - 2 * MARGIN; // 12960 DXA

// Track dynamically created numbering configs (one per list instance)
const numberingConfigs = [];
let listCounter = 0;

const COLORS = {
  GREY: "888888",
  PLACEHOLDER_BLUE: "0971AA", // matches CaseWare UI placeholder color
  DARK_NAVY: "1F3864",
  WHITE: "FFFFFF",
  LIGHT_BLUE: "D9E1F2",
  GUIDANCE_BG: "FFF3CD",
  GUIDANCE_BORDER: "FFE69C",
  GUIDANCE_TEXT: "664D03",
};

const THIN_BORDER = {
  style: BorderStyle.SINGLE,
  size: 1,
  color: "CCCCCC",
};
const CELL_BORDERS = {
  top: THIN_BORDER,
  bottom: THIN_BORDER,
  left: THIN_BORDER,
  right: THIN_BORDER,
};
const CELL_MARGINS = { top: 80, bottom: 80, left: 120, right: 120 };

// ── HTML PARSING (cheerio-based) ────────────────────────────────────────────

/**
 * Convert CaseWare HTML to an array of docx-js elements (Paragraph, Table).
 * Uses cheerio for robust DOM parsing instead of regex.
 */
function parseHtml(html) {
  if (!html || !html.trim()) return [];

  const $ = cheerio.load(html, { xmlMode: false, decodeEntities: true });
  const elements = [];

  // Track last numbering ref per list-style-type so split <ol> elements
  // with start > 1 continue instead of restarting
  const listCtx = {}; // { "decimal": "numbers_1", "lower-alpha": "numbers_3", ... }

  // Process top-level block elements
  const body = $.root().children().length > 0 ? $.root() : $("body");

  function processBlock(el) {
    const tag = el.tagName ? el.tagName.toLowerCase() : "";

    if (el.type === "text") {
      const text = $(el).text().trim();
      if (text) {
        elements.push(
          new Paragraph({
            spacing: { after: 120 },
            children: [new TextRun({ text })],
          })
        );
      }
      return;
    }

    if (tag === "p" || tag === "div") {
      const runs = processInlineChildren($(el), $);
      if (runs.length > 0) {
        elements.push(
          new Paragraph({
            spacing: { after: 120 },
            children: runs,
          })
        );
      } else {
        // Empty paragraph for spacing
        elements.push(new Paragraph({ spacing: { after: 60 }, children: [] }));
      }
    } else if (tag === "ul" || tag === "ol") {
      const $listEl = $(el);
      const isBullet = tag !== "ol";

      if (isBullet) {
        // Bullets always get a fresh reference
        const ref = createListRef(true, null);
        elements.push(...processList($listEl, $, ref, 0));
      } else {
        // Ordered list — check start attribute to decide new vs continuation
        const startAttr = parseInt($listEl.attr("start") || "1", 10);
        const styleMatch = ($listEl.attr("style") || "").match(/list-style-type:\s*([^;"]+)/);
        const styleType = styleMatch ? styleMatch[1].trim() : "decimal";

        let ref;
        if (startAttr > 1 && listCtx[styleType]) {
          // Continuation — reuse the previous reference for this style
          ref = listCtx[styleType];
        } else {
          // New list — create fresh reference and track it
          const formatOverride = detectListFormat($listEl);
          ref = createListRef(false, formatOverride);
          listCtx[styleType] = ref;
        }
        elements.push(...processList($listEl, $, ref, 0));
      }
    } else if (tag === "table") {
      elements.push(processTable($(el), $));
    } else if (/^h[1-6]$/.test(tag)) {
      const level = parseInt(tag[1]);
      const runs = processInlineChildren($(el), $);
      const headingLevels = [
        HeadingLevel.HEADING_1, HeadingLevel.HEADING_2,
        HeadingLevel.HEADING_3, HeadingLevel.HEADING_4,
        HeadingLevel.HEADING_5, HeadingLevel.HEADING_6,
      ];
      elements.push(
        new Paragraph({
          heading: headingLevels[Math.min(level - 1, 5)],
          children: runs,
        })
      );
    } else if (tag === "br") {
      // Skip standalone br
    } else {
      // Unknown block — recurse into children
      $(el)
        .contents()
        .each((_, child) => processBlock(child));
    }
  }

  body.contents().each((_, el) => processBlock(el));

  return elements;
}

/**
 * Process children of an element into TextRun objects, preserving inline formatting.
 */
function processInlineChildren($el, $, style = {}, _ctx) {
  // _ctx tracks the last text added across recursive calls to prevent double spaces
  const ctx = _ctx || { lastText: "" };
  const runs = [];

  $el.contents().each((_, node) => {
    if (node.type === "text") {
      // Collapse multiple whitespace into single spaces (like a browser would)
      let text = $(node).text().replace(/\s+/g, " ");
      if (text.trim() !== "") {
        // If last text ended with space and this starts with space, trim leading
        if (ctx.lastText.endsWith(" ") && text.startsWith(" ")) {
          text = text.trimStart();
        }
        if (text) {
          runs.push(createTextRun(text, style));
          ctx.lastText = text;
        }
      } else if (text === " " && ctx.lastText && !ctx.lastText.endsWith(" ")) {
        // Single space between elements — only if not already trailing space
        runs.push(createTextRun(" ", style));
        ctx.lastText = " ";
      }
      return;
    }

    const tag = node.tagName ? node.tagName.toLowerCase() : "";
    const $node = $(node);

    if (tag === "br") {
      runs.push(new TextRun({ break: 1 }));
      ctx.lastText = "";
    } else if (tag === "strong" || tag === "b") {
      runs.push(...processInlineChildren($node, $, { ...style, bold: true }, ctx));
    } else if (tag === "em" || tag === "i") {
      runs.push(...processInlineChildren($node, $, { ...style, italics: true }, ctx));
    } else if (tag === "u") {
      runs.push(...processInlineChildren($node, $, { ...style, underline: {} }, ctx));
    } else if (tag === "s" || tag === "strike" || tag === "del") {
      runs.push(...processInlineChildren($node, $, { ...style, strike: true }, ctx));
    } else if (tag === "sub") {
      runs.push(...processInlineChildren($node, $, { ...style, subScript: true }, ctx));
    } else if (tag === "sup") {
      runs.push(...processInlineChildren($node, $, { ...style, superScript: true }, ctx));
    } else if (tag === "a") {
      runs.push(...processInlineChildren($node, $, { ...style, underline: {}, color: "0563C1" }, ctx));
    } else if (tag === "span") {
      // Check for CaseWare-specific spans
      if ($node.hasClass("cw-placeholder") || $node.attr("placeholder")) {
        const text = $node.text().trim();
        if (text) {
          const display = text.startsWith("((") ? text : `((${text}))`;
          runs.push(createTextRun(display, { ...style, color: COLORS.PLACEHOLDER_BLUE }));
          ctx.lastText = display;
        }
      } else if ($node.hasClass("cw-formula")) {
        const text = $node.text().trim();
        if (text) {
          runs.push(createTextRun(text, { ...style, color: COLORS.GREY }));
          ctx.lastText = text;
        }
      } else if ($node.attr("formula") || $node.hasClass("dynamic-text")) {
        // Original unresolved formula span — skip the space-only content
        // These contain just " " as placeholder content in the source HTML
        const text = $node.text().trim();
        if (text) {
          runs.push(createTextRun(`[[${text}]]`, { ...style, color: COLORS.GREY }));
          ctx.lastText = `[[${text}]]`;
        }
        // If empty/space-only, intentionally skip (don't add extra space)
      } else if ($node.hasClass("firm-logo") || $node.attr("type") === "firm-logo") {
        runs.push(createTextRun("((Firm Logo))", { ...style, color: COLORS.PLACEHOLDER_BLUE }));
        ctx.lastText = "((Firm Logo))";
      } else if ($node.hasClass("caret") || $node.hasClass("hidden-print")) {
        // CaseWare UI helper spans — skip entirely
      } else {
        // Generic span — recurse
        runs.push(...processInlineChildren($node, $, style, ctx));
      }
    } else if (tag === "img") {
      // Skip images (base64 logos etc.) — represent as placeholder
      const title = $node.attr("title") || $node.attr("alt") || "Image";
      runs.push(createTextRun(`((${title}))`, { ...style, color: COLORS.PLACEHOLDER_BLUE }));
      ctx.lastText = `((${title}))`;
    } else {
      // Unknown inline — recurse
      runs.push(...processInlineChildren($node, $, style, ctx));
    }
  });

  return runs;
}

function createTextRun(text, style) {
  const opts = { text };
  if (style.bold) opts.bold = true;
  if (style.italics) opts.italics = true;
  if (style.underline) opts.underline = {};
  if (style.strike) opts.strike = true;
  if (style.color) opts.color = style.color;
  if (style.subScript) opts.subScript = true;
  if (style.superScript) opts.superScript = true;
  if (style.size) opts.size = style.size;
  if (style.font) opts.font = style.font;
  return new TextRun(opts);
}

// ── LIST PARSING ────────────────────────────────────────────────────────────

/**
 * Detect the list format from an <ol> or <ul> element's style attribute.
 * Maps CSS list-style-type to docx-js LevelFormat.
 */
function detectListFormat($list) {
  const style = $list.attr("style") || "";
  const match = style.match(/list-style-type:\s*([^;"]+)/);
  const cssType = match ? match[1].trim() : "";

  switch (cssType) {
    case "lower-alpha":
    case "lower-latin":
      return LevelFormat.LOWER_LETTER;
    case "upper-alpha":
    case "upper-latin":
      return LevelFormat.UPPER_LETTER;
    case "lower-roman":
      return LevelFormat.LOWER_ROMAN;
    case "upper-roman":
      return LevelFormat.UPPER_ROMAN;
    default:
      return null; // use default for the list type
  }
}

/**
 * Create a unique numbering reference for a list instance.
 * Each call creates a new independent list that starts from 1/a/i.
 */
function createListRef(isBullet, formatOverride) {
  listCounter++;
  const ref = isBullet ? `bullets_${listCounter}` : `numbers_${listCounter}`;

  if (isBullet) {
    numberingConfigs.push({
      reference: ref,
      levels: [
        { level: 0, format: LevelFormat.BULLET, text: "\u2022", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } } },
        { level: 1, format: LevelFormat.BULLET, text: "\u25E6", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 1440, hanging: 360 } } } },
        { level: 2, format: LevelFormat.BULLET, text: "\u25AA", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 2160, hanging: 360 } } } },
        { level: 3, format: LevelFormat.BULLET, text: "\u2022", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 2880, hanging: 360 } } } },
      ],
    });
  } else {
    // Determine level 0 format — use override if provided, otherwise decimal
    const l0Format = formatOverride || LevelFormat.DECIMAL;
    const l0Text = l0Format === LevelFormat.LOWER_LETTER ? "%1." :
                   l0Format === LevelFormat.UPPER_LETTER ? "%1." :
                   l0Format === LevelFormat.LOWER_ROMAN ? "%1." :
                   l0Format === LevelFormat.UPPER_ROMAN ? "%1." : "%1.";

    numberingConfigs.push({
      reference: ref,
      levels: [
        { level: 0, format: l0Format, text: l0Text, alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } } },
        { level: 1, format: LevelFormat.LOWER_LETTER, text: "%2.", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 1440, hanging: 360 } } } },
        { level: 2, format: LevelFormat.LOWER_ROMAN, text: "%3.", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 2160, hanging: 360 } } } },
        { level: 3, format: LevelFormat.DECIMAL, text: "%4.", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 2880, hanging: 360 } } } },
      ],
    });
  }

  return ref;
}

function processList($list, $, ref, level) {
  const items = [];

  $list.children("li").each((_, li) => {
    const $li = $(li);
    // Separate nested lists from inline content
    const nestedLists = $li.children("ul, ol");
    const $clone = $li.clone();
    $clone.children("ul, ol").remove();

    const runs = processInlineChildren($clone, $);
    if (runs.length > 0) {
      items.push(
        new Paragraph({
          numbering: { reference: ref, level: Math.min(level, 3) },
          children: runs,
        })
      );
    }

    // Process nested lists — each nested list gets its own reference
    nestedLists.each((_, nested) => {
      const $nested = $(nested);
      const nestedTag = nested.tagName.toLowerCase();
      const isBullet = nestedTag !== "ol";
      const nestedFormat = !isBullet ? detectListFormat($nested) : null;
      const nestedRef = createListRef(isBullet, nestedFormat);
      items.push(...processList($nested, $, nestedRef, level + 1));
    });
  });

  return items;
}

// ── TABLE PARSING ───────────────────────────────────────────────────────────

function processTable($table, $) {
  const rows = [];
  let maxCols = 0;

  // First pass: count columns
  $table.find("tr").each((_, tr) => {
    const cellCount = $(tr).children("td, th").length;
    maxCols = Math.max(maxCols, cellCount);
  });

  if (maxCols === 0) return new Paragraph({ children: [] });

  const colWidth = Math.floor(CONTENT_WIDTH / maxCols);
  const columnWidths = Array(maxCols).fill(colWidth);
  columnWidths[maxCols - 1] = CONTENT_WIDTH - colWidth * (maxCols - 1);

  $table.find("tr").each((_, tr) => {
    const tableCells = [];
    const $cells = $(tr).children("td, th");

    for (let i = 0; i < maxCols; i++) {
      const $cell = $cells.eq(i);
      const runs = $cell.length ? processInlineChildren($cell, $) : [];
      tableCells.push(
        new TableCell({
          borders: CELL_BORDERS,
          margins: CELL_MARGINS,
          width: { size: columnWidths[i], type: WidthType.DXA },
          children: [new Paragraph({ children: runs })],
        })
      );
    }

    rows.push(new TableRow({ children: tableCells }));
  });

  return new Table({
    width: { size: CONTENT_WIDTH, type: WidthType.DXA },
    columnWidths: columnWidths,
    rows: rows,
  });
}

// ── GUIDANCE CALLOUT ────────────────────────────────────────────────────────

function renderGuidance(guidanceHtml) {
  const innerElements = parseHtml(guidanceHtml);
  const children = [
    new Paragraph({
      spacing: { after: 60 },
      children: [
        new TextRun({
          text: "Guidance",
          bold: true,
          color: COLORS.GUIDANCE_TEXT,
          size: 20, // 10pt
        }),
      ],
    }),
    ...innerElements,
  ];

  return new Table({
    width: { size: CONTENT_WIDTH, type: WidthType.DXA },
    columnWidths: [CONTENT_WIDTH],
    rows: [
      new TableRow({
        children: [
          new TableCell({
            borders: {
              top: {
                style: BorderStyle.SINGLE,
                size: 1,
                color: COLORS.GUIDANCE_BORDER,
              },
              bottom: {
                style: BorderStyle.SINGLE,
                size: 1,
                color: COLORS.GUIDANCE_BORDER,
              },
              left: {
                style: BorderStyle.SINGLE,
                size: 6,
                color: COLORS.GUIDANCE_BORDER,
              },
              right: {
                style: BorderStyle.SINGLE,
                size: 1,
                color: COLORS.GUIDANCE_BORDER,
              },
            },
            margins: { top: 120, bottom: 120, left: 200, right: 200 },
            shading: {
              fill: COLORS.GUIDANCE_BG,
              type: ShadingType.CLEAR,
            },
            width: { size: CONTENT_WIDTH, type: WidthType.DXA },
            children: children,
          }),
        ],
      }),
    ],
  });
}

// ── VISIBILITY INLINE ───────────────────────────────────────────────────────

function renderVisibilityInline(visibility) {
  if (!visibility) return [];

  const elements = [];

  // Direction line
  elements.push(
    new Paragraph({
      spacing: { before: 60, after: 40 },
      indent: { left: 360 },
      children: [
        new TextRun({
          text: `Visibility: ${visibility.direction}`,
          italics: true,
          color: COLORS.GREY,
          size: 18, // 9pt
        }),
      ],
    })
  );

  // Condition lines
  for (const cond of visibility.conditions || []) {
    const parts = [];
    if (cond.group) parts.push(cond.group);
    if (cond.name) parts.push(cond.name);
    if (cond.response) parts.push(`= ${cond.response}`);
    const text = parts.join(" > ");

    elements.push(
      new Paragraph({
        spacing: { after: 20 },
        indent: { left: 720 },
        children: [
          new TextRun({
            text: `- ${text}`,
            italics: true,
            color: COLORS.GREY,
            size: 16, // 8pt
          }),
        ],
      })
    );
  }

  // Spacer
  elements.push(new Paragraph({ spacing: { after: 80 }, children: [] }));

  return elements;
}

// ── VISIBILITY SUMMARY TABLE ────────────────────────────────────────────────

function renderVisibilitySummaryTable(sections) {
  // Collect sections with visibility
  const visibleSections = sections.filter(
    (s) => s.visibility && s.visibility.conditions && s.visibility.conditions.length > 0
  );

  if (visibleSections.length === 0) return [];

  const colWidths = [4000, 2500, CONTENT_WIDTH - 6500]; // Section | Setting | Conditions
  const headerBorder = {
    style: BorderStyle.SINGLE,
    size: 1,
    color: COLORS.DARK_NAVY,
  };
  const headerBorders = {
    top: headerBorder,
    bottom: headerBorder,
    left: headerBorder,
    right: headerBorder,
  };

  // Header row
  const headerRow = new TableRow({
    children: ["Section", "Visibility Setting", "Conditions"].map(
      (text, i) =>
        new TableCell({
          borders: headerBorders,
          margins: CELL_MARGINS,
          width: { size: colWidths[i], type: WidthType.DXA },
          shading: {
            fill: COLORS.DARK_NAVY,
            type: ShadingType.CLEAR,
          },
          children: [
            new Paragraph({
              children: [
                new TextRun({ text, bold: true, color: COLORS.WHITE, size: 20 }),
              ],
            }),
          ],
        })
    ),
  });

  // Data rows
  const dataRows = [];
  visibleSections.forEach((section, idx) => {
    const vis = section.visibility;
    const conditions = (vis.conditions || [])
      .map((c) => {
        const parts = [];
        if (c.name) parts.push(c.name);
        if (c.response) parts.push(`= ${c.response}`);
        return parts.join(" ");
      })
      .join("\n");

    const rowShading =
      idx % 2 === 1 ? COLORS.LIGHT_BLUE : COLORS.WHITE;

    dataRows.push(
      new TableRow({
        children: [
          new TableCell({
            borders: CELL_BORDERS,
            margins: CELL_MARGINS,
            width: { size: colWidths[0], type: WidthType.DXA },
            shading: { fill: rowShading, type: ShadingType.CLEAR },
            children: [
              new Paragraph({
                children: [
                  new TextRun({ text: section.title || "", size: 20 }),
                ],
              }),
            ],
          }),
          new TableCell({
            borders: CELL_BORDERS,
            margins: CELL_MARGINS,
            width: { size: colWidths[1], type: WidthType.DXA },
            shading: { fill: rowShading, type: ShadingType.CLEAR },
            children: [
              new Paragraph({
                children: [
                  new TextRun({ text: vis.direction || "", size: 20 }),
                ],
              }),
            ],
          }),
          new TableCell({
            borders: CELL_BORDERS,
            margins: CELL_MARGINS,
            width: { size: colWidths[2], type: WidthType.DXA },
            shading: { fill: rowShading, type: ShadingType.CLEAR },
            children: conditions.split("\n").map(
              (line) =>
                new Paragraph({
                  children: [new TextRun({ text: line, size: 18 })],
                })
            ),
          }),
        ],
      })
    );
  });

  return [
    new Paragraph({
      heading: HeadingLevel.HEADING_1,
      spacing: { before: 240, after: 240 },
      children: [new TextRun({ text: "Visibility Settings Summary" })],
    }),
    new Table({
      width: { size: CONTENT_WIDTH, type: WidthType.DXA },
      columnWidths: colWidths,
      rows: [headerRow, ...dataRows],
    }),
  ];
}

// ── MAIN DOCUMENT BUILDER ───────────────────────────────────────────────────

function buildDocument(data) {
  const { document_name, sections } = data;
  const children = [];

  // Document title
  children.push(
    new Paragraph({
      heading: HeadingLevel.HEADING_1,
      spacing: { after: 240 },
      children: [new TextRun({ text: document_name })],
    })
  );

  // Walk sections
  for (const section of sections) {
    const { type, title, level, html_content, guidance_html, visibility } =
      section;

    // Skip settings/toc (should already be filtered by Python)
    if (type === "settings" || type === "toc") continue;

    // Page break
    if (type === "pagebreak") {
      children.push(new Paragraph({ children: [new PageBreak()] }));
      continue;
    }

    // Grouping/heading — add as Word heading
    if (type === "grouping" || type === "heading") {
      if (title) {
        const headingLevel =
          level <= 0
            ? HeadingLevel.HEADING_1
            : level === 1
              ? HeadingLevel.HEADING_2
              : HeadingLevel.HEADING_3;
        children.push(
          new Paragraph({
            heading: headingLevel,
            spacing: { before: 240, after: 120 },
            children: [new TextRun({ text: title })],
          })
        );
      }
      // Guidance on grouping section
      if (guidance_html) {
        children.push(renderGuidance(guidance_html));
        children.push(new Paragraph({ spacing: { after: 120 }, children: [] }));
      }
      // Visibility on grouping
      if (visibility && visibility.conditions && visibility.conditions.length > 0) {
        children.push(...renderVisibilityInline(visibility));
      }
      continue;
    }

    // Content section
    if (type === "content") {
      // Section title as heading (offset by level)
      if (title) {
        const headingLevel =
          level <= 1
            ? HeadingLevel.HEADING_2
            : level === 2
              ? HeadingLevel.HEADING_3
              : HeadingLevel.HEADING_4;
        children.push(
          new Paragraph({
            heading: headingLevel,
            spacing: { before: 200, after: 100 },
            children: [new TextRun({ text: title })],
          })
        );
      }

      // Body content
      if (html_content) {
        const contentElements = parseHtml(html_content);
        children.push(...contentElements);
      }

      // Guidance callout
      if (guidance_html) {
        children.push(
          new Paragraph({ spacing: { before: 80 }, children: [] })
        );
        children.push(renderGuidance(guidance_html));
        children.push(new Paragraph({ spacing: { after: 120 }, children: [] }));
      }

      // Inline visibility
      if (visibility && visibility.conditions && visibility.conditions.length > 0) {
        children.push(...renderVisibilityInline(visibility));
      }

      continue;
    }

    // Fallback for unknown types: render title + content if present
    if (title) {
      children.push(
        new Paragraph({
          spacing: { before: 120, after: 60 },
          children: [new TextRun({ text: title, bold: true })],
        })
      );
    }
    if (html_content) {
      children.push(...parseHtml(html_content));
    }
  }

  // Visibility summary table at the end
  const summaryElements = renderVisibilitySummaryTable(sections);
  if (summaryElements.length > 0) {
    children.push(new Paragraph({ children: [new PageBreak()] }));
    children.push(...summaryElements);
  }

  // Build the document
  const doc = new Document({
    styles: {
      default: {
        document: {
          run: { font: "Arial", size: 24 }, // 12pt
        },
      },
      paragraphStyles: [
        {
          id: "Heading1",
          name: "Heading 1",
          basedOn: "Normal",
          next: "Normal",
          quickFormat: true,
          run: { size: 32, bold: true, font: "Arial" },
          paragraph: {
            spacing: { before: 240, after: 240 },
            outlineLevel: 0,
          },
        },
        {
          id: "Heading2",
          name: "Heading 2",
          basedOn: "Normal",
          next: "Normal",
          quickFormat: true,
          run: { size: 28, bold: true, font: "Arial" },
          paragraph: {
            spacing: { before: 180, after: 180 },
            outlineLevel: 1,
          },
        },
        {
          id: "Heading3",
          name: "Heading 3",
          basedOn: "Normal",
          next: "Normal",
          quickFormat: true,
          run: { size: 24, bold: true, font: "Arial" },
          paragraph: {
            spacing: { before: 120, after: 120 },
            outlineLevel: 2,
          },
        },
        {
          id: "Heading4",
          name: "Heading 4",
          basedOn: "Normal",
          next: "Normal",
          quickFormat: true,
          run: { size: 22, bold: true, font: "Arial" },
          paragraph: {
            spacing: { before: 120, after: 60 },
            outlineLevel: 3,
          },
        },
      ],
    },
    numbering: {
      config: numberingConfigs,
    },
    sections: [
      {
        properties: {
          page: {
            size: {
              width: PAGE_WIDTH,
              height: PAGE_HEIGHT,
              orientation: PageOrientation.LANDSCAPE,
            },
            margin: {
              top: MARGIN,
              right: MARGIN,
              bottom: MARGIN,
              left: MARGIN,
            },
          },
        },
        children: children,
      },
    ],
  });

  return doc;
}

// ── ENTRY POINT ─────────────────────────────────────────────────────────────

async function main() {
  const [jsonPath, outputPath] = process.argv.slice(2);
  if (!jsonPath || !outputPath) {
    console.error("Usage: node generate_docx.js <input.json> <output.docx>");
    process.exit(1);
  }

  const data = JSON.parse(fs.readFileSync(jsonPath, "utf-8"));
  const doc = buildDocument(data);

  const buffer = await Packer.toBuffer(doc);
  fs.writeFileSync(outputPath, buffer);

  console.log(
    `Generated ${outputPath} (${Math.round(buffer.length / 1024)} KB)`
  );
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
