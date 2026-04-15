/**
 * Vercel serverless function: /api/generate
 *
 * Accepts POST { url, templateName } — extracts letter content from CaseWare
 * Cloud and returns a Word (.docx) file.
 *
 * Combines the Python extraction logic (letter_extract.py) and the Node.js
 * Word generation (generate_docx.js) into a single serverless function.
 */

const fetch = require("node-fetch");
const cheerio = require("cheerio");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  HeadingLevel, AlignmentType, LevelFormat, BorderStyle, WidthType,
  ShadingType, PageBreak, PageOrientation,
} = require("docx");

// ── URL PARSING ─────────────────────────────────────────────────────────────

const CW_URL_PATTERN = /https?:\/\/([^/]+)\/([^/]+)\/e\/eng\/([^/]+)/;
const CW_DOC_PATTERN = /#\/(?:efinancials|letter)\/([^/?\s]+)/;

function parseUrl(url) {
  const match = url.match(CW_URL_PATTERN);
  if (!match) throw new Error("Invalid CaseWare URL");
  const host = `https://${match[1]}`;
  const tenant = match[2];
  const engagementId = match[3];
  const docMatch = url.match(CW_DOC_PATTERN);
  const documentId = docMatch ? docMatch[1] : "";
  return { host, tenant, engagementId, documentId };
}

// ── SESSION / AUTH ──────────────────────────────────────────────────────────

function envPrefixFromHost(host) {
  const hostname = host.replace(/https?:\/\//, "").split("/")[0];
  return hostname.split(".")[0].toUpperCase();
}

async function obtainBearerToken(envPrefix, host, tenant) {
  let clientId = "", clientSecret = "";
  if (envPrefix) {
    clientId = (process.env[`CW_${envPrefix}_CLIENT_ID`] || "").trim();
    clientSecret = (process.env[`CW_${envPrefix}_CLIENT_SECRET`] || "").trim();
  }
  if (!clientId || !clientSecret) {
    clientId = (process.env.CW_CLIENT_ID || "").trim();
    clientSecret = (process.env.CW_CLIENT_SECRET || "").trim();
  }
  if (!clientId || !clientSecret) return null;

  const url = `${host}/${tenant}/ms/caseware-cloud/api/v1/auth/token`;
  const resp = await fetch(url, {
    method: "POST",
    headers: { "Accept": "application/json", "Content-Type": "application/json" },
    body: JSON.stringify({ ClientId: clientId, ClientSecret: clientSecret, Language: "en" }),
  });
  if (!resp.ok) throw new Error(`Auth failed: ${resp.status}`);
  const data = await resp.json();
  return data.Token || null;
}

async function makeHeaders(envPrefix, host, tenant) {
  const headers = { "Accept": "application/json", "Content-Type": "application/json" };
  const token = await obtainBearerToken(envPrefix, host, tenant);
  if (token) {
    headers["Authorization"] = `Bearer ${token}`;
    return headers;
  }
  const cookies = (process.env.CW_COOKIES || "").trim();
  if (!cookies) throw new Error("No auth credentials found.");
  headers["Cookie"] = cookies;
  return headers;
}

// ── API HELPERS ─────────────────────────────────────────────────────────────

function unwrapResponse(data) {
  if (Array.isArray(data)) return data;
  if (data && typeof data === "object") {
    for (const key of ["objects", "sections", "items", "results", "data"]) {
      if (Array.isArray(data[key])) return data[key];
    }
    for (const val of Object.values(data)) {
      if (Array.isArray(val)) return val;
    }
    if (data.object && typeof data.object === "object") return [data.object];
  }
  return [];
}

async function apiPost(headers, url, payload = {}) {
  const resp = await fetch(url, {
    method: "POST", headers,
    body: JSON.stringify(payload),
  });
  if (resp.status === 401) throw new Error("401 Unauthorised — credentials have expired.");
  if (!resp.ok) throw new Error(`${resp.status} from ${url}`);
  return unwrapResponse(await resp.json());
}

// ── DATA FETCHING ───────────────────────────────────────────────────────────

async function fetchDocuments(headers, engagementId, host, tenant) {
  const url = `${host}/${tenant}/e/eng/${engagementId}/api/v1.12.0/document/get`;
  return apiPost(headers, url, {});
}

async function fetchSections(headers, engagementId, documentId, host, tenant) {
  const url = `${host}/${tenant}/e/eng/${engagementId}/api/v1.12.0/section/get`;
  return apiPost(headers, url, {
    filter: { filter: {
      node: "=",
      left: { node: "field", kind: "section", field: "document" },
      right: { node: "string", value: documentId },
    }}
  });
}

async function fetchDocumentLookup(headers, engagementId, host, tenant) {
  try {
    const docs = await fetchDocuments(headers, engagementId, host, tenant);
    const result = {};
    for (const doc of docs) {
      const did = doc.id || "";
      const number = doc.number || "";
      const name = (doc.names || {}).en || doc.name || "";
      const content = doc.content || "";
      if (did) {
        const label = number ? `${number} ${name}`.trim() : name;
        result[did] = label;
        if (content) result[content] = label;
      }
    }
    return result;
  } catch { return {}; }
}

async function fetchProcedureById(headers, engagementId, procedureId, host, tenant) {
  const url = `${host}/${tenant}/e/eng/${engagementId}/api/v1.12.0/procedure/get`;
  try {
    const resp = await fetch(url, {
      method: "POST", headers,
      body: JSON.stringify({ filter: { filter: {
        node: "=",
        left: { node: "field", kind: "procedure", field: "id" },
        right: { node: "string", value: procedureId },
      }}}),
    });
    if (!resp.ok) return null;
    const data = await resp.json();
    const objects = data.objects || [];
    return objects[0] || null;
  } catch { return null; }
}

async function fetchProceduresForChecklist(headers, engagementId, checklistId, host, tenant) {
  const url = `${host}/${tenant}/e/eng/${engagementId}/api/v1.12.0/procedure/get`;
  try {
    const resp = await fetch(url, {
      method: "POST", headers,
      body: JSON.stringify({ filter: { filter: {
        node: "=",
        left: { node: "field", kind: "procedure", field: "checklistId" },
        right: { node: "string", value: checklistId },
      }}}),
    });
    if (!resp.ok) return [];
    const data = await resp.json();
    return (typeof data === "object" && !Array.isArray(data)) ? (data.objects || []) : [];
  } catch { return []; }
}

// ── HTML HELPERS ────────────────────────────────────────────────────────────

function stripHtml(html, formulaMap) {
  if (!html) return "";
  let text = html.replace(/<span[^>]*\bplaceholder="[^"]*"[^>]*>(.*?)<\/span>/gi, "(($1))");
  if (formulaMap) {
    text = text.replace(/<span[^>]*\bformula="([^"]*)"[^>]*>.*?<\/span>/gi, (_, fid) => {
      const val = formulaMap[fid] || "";
      return val ? `[[${val}]]` : "";
    });
  }
  text = text.replace(/<[^>]+>/g, " ");
  text = text.replace(/&amp;/g, "&").replace(/&lt;/g, "<").replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"').replace(/&#39;/g, "'").replace(/&nbsp;/g, " ").replace(/&#160;/g, " ");
  text = text.replace(/\s+/g, " ");
  return text.trim();
}

function buildFormulaMap(section) {
  const result = {};
  for (const a of Object.values(section.attachables || {})) {
    const refId = a.referenceId;
    const calc = (a.calculated || "").trim();
    if (refId && calc) result[refId] = calc;
  }
  return result;
}

function preprocessHtml(html, formulaMap) {
  if (!html) return "";
  let text = html.replace(
    /<span[^>]*\bplaceholder="[^"]*"[^>]*>(.*?)<\/span>/gi,
    '<span class="cw-placeholder">(($1))</span>'
  );
  if (formulaMap) {
    text = text.replace(/<span[^>]*\bformula="([^"]*)"[^>]*>.*?<\/span>/gi, (_, fid) => {
      const val = formulaMap[fid] || "";
      return val ? `<span class="cw-formula">[[${val}]]</span>` : "";
    });
  }
  // Collapse whitespace
  text = text.replace(/>\s{2,}</g, "> <");
  text = text.replace(/>\s{2,}/g, "> ");
  text = text.replace(/\s{2,}</g, " <");
  text = text.replace(/(?<=>)[^<]+(?=<)/g, (m) => m.replace(/ {2,}/g, " "));
  text = text.replace(/^[^<]+/, (m) => m.replace(/ {2,}/g, " "));
  text = text.replace(/>[^<]+$/, (m) => m.replace(/ {2,}/g, " "));
  return text;
}

// ── VISIBILITY ID RESOLUTION ────────────────────────────────────────────────

function collectProcedureIds(cond, ids) {
  const pid = (cond.procedureId || {}).id;
  if (pid) ids.add(pid);
  for (const sub of cond.conditions || []) collectProcedureIds(sub, ids);
}

function collectChecklistIds(cond, ids) {
  const cid = (cond.checklistId || {}).id;
  if (cid) ids.add(cid);
  for (const sub of cond.conditions || []) collectChecklistIds(sub, ids);
}

async function buildIdLookup(headers, engagementId, sections, host, tenant) {
  const procedureIds = new Set();
  const checklistIds = new Set();
  for (const s of sections) {
    for (const cond of (s.visibility || {}).conditions || []) {
      collectProcedureIds(cond, procedureIds);
      collectChecklistIds(cond, checklistIds);
    }
  }

  const lookup = await fetchDocumentLookup(headers, engagementId, host, tenant);

  // Fetch checklist-level response sets
  const clUrl = `${host}/${tenant}/e/eng/${engagementId}/api/v1.12.0/checklist/get`;
  for (const cid of checklistIds) {
    try {
      const resp = await fetch(clUrl, {
        method: "POST", headers,
        body: JSON.stringify({ filter: { filter: {
          node: "=",
          left: { node: "field", kind: "checklist", field: "id" },
          right: { node: "string", value: cid },
        }}}),
      });
      if (resp.ok) {
        for (const cl of (await resp.json()).objects || []) {
          for (const rs of (cl.settings || {}).responseSets || []) {
            for (const ro of rs.responses || []) {
              if (ro.id && (ro.name || (ro.names || {}).en))
                lookup[ro.id] = ro.name || ro.names.en;
            }
          }
        }
      }
    } catch {}
  }

  // Fetch procedures per checklist
  for (const cid of checklistIds) {
    const procs = await fetchProceduresForChecklist(headers, engagementId, cid, host, tenant);
    for (const proc of procs) {
      if (proc.id && !lookup[proc.id]) {
        const name = (proc.summaryNames || {}).en || stripHtml(proc.text || "");
        if (name) lookup[proc.id] = name;
      }
      for (const rs of (proc.settings || {}).responseSets || []) {
        for (const ro of rs.responses || []) {
          if (ro.id && (ro.name || (ro.names || {}).en))
            lookup[ro.id] = ro.name || ro.names.en;
        }
      }
    }
  }

  // Individually resolve remaining procedure IDs
  for (const pid of procedureIds) {
    if (lookup[pid]) continue;
    const proc = await fetchProcedureById(headers, engagementId, pid, host, tenant);
    if (!proc) continue;
    const name = (proc.summaryNames || {}).en || stripHtml(proc.text || "");
    if (name) lookup[pid] = name;
    for (const rs of (proc.settings || {}).responseSets || []) {
      for (const ro of rs.responses || []) {
        if (ro.id && (ro.name || (ro.names || {}).en))
          lookup[ro.id] = ro.name || ro.names.en;
      }
    }
  }

  return lookup;
}

// ── VISIBILITY PARSING ──────────────────────────────────────────────────────

const ORG_TYPE_LABELS = {
  CorporationControlledPrivateCorporation: "Canadian Controlled Private Corporation (CCPC)",
  CorporationControlledPublicCorporation: "Corporation Controlled by a Public Corporation",
  OtherPrivateCorporation: "Other Private Corporation",
  PublicCompany: "Public Company", Individual: "Individual",
  Coownership: "Co-ownership", GeneralPartnership: "General Partnership",
  LimitedPartnership: "Limited Partnership",
  LimitedLiabilityPartnership: "Limited Liability Partnership",
  JointVenture: "Joint Venture", Trust: "Trust", Cooperative: "Cooperative",
  PensionFunds: "Pension Funds", RegisteredCharity: "Registered Charity",
  NotForProfit: "Not For Profit", Government: "Government",
  LimitedLiabilityCompany: "Limited Liability Company (LLC)",
  CCorporation: "C-Corporation", SCorporation: "S-Corporation",
  SoleProprietorship: "Sole Proprietorship", Partnership: "Partnership",
  NotForProfitPrivate: "Not for Profit - Private",
  NotForProfitPublic: "Not for Profit - Public",
  PublicFoundation: "Public Foundation", PrivateFoundation: "Private Foundation",
};

function splitPascalCase(s) {
  return s.replace(/(?<=[a-z])(?=[A-Z])|(?<=[A-Z])(?=[A-Z][a-z])/g, " ");
}

function resolve(lookup, idObj) {
  if (!idObj) return "";
  const rawId = typeof idObj === "object" ? (idObj.id || "") : String(idObj);
  return lookup[rawId] || rawId.slice(0, 12);
}

function resolveOrgType(cond) {
  const customId = cond.customOrganizationTypeId || "";
  if (customId) return ORG_TYPE_LABELS[customId] || splitPascalCase(customId);
  const orgType = cond.organizationType || "";
  if (orgType) return ORG_TYPE_LABELS[orgType] || splitPascalCase(orgType);
  return (cond.id || "unknown").slice(0, 12);
}

function flattenConditions(conditions, lookup, groupLabel = "") {
  const rows = [];
  for (const cond of conditions) {
    const ctype = cond.type || "";
    if (ctype === "response") {
      rows.push({
        group: groupLabel || resolve(lookup, cond.checklistId),
        name: resolve(lookup, cond.procedureId),
        response: resolve(lookup, cond.responseId),
      });
    } else if (ctype === "condition_group") {
      const nested = cond.conditions || [];
      const qualifier = cond.allConditionsNeeded ? "all" : "any";
      const firstCl = nested.length ? resolve(lookup, nested[0].checklistId) : "";
      const label = firstCl ? `${firstCl} (${qualifier})` : `(${qualifier})`;
      rows.push(...flattenConditions(nested, lookup, label));
    } else if (ctype === "organization_type") {
      rows.push({ group: groupLabel, name: "Organization Type", response: resolveOrgType(cond) });
    } else if (ctype === "consolidation") {
      rows.push({ group: groupLabel, name: "Consolidation", response: cond.consolidated ? "Consolidated" : "Not consolidated" });
    } else {
      rows.push({ group: groupLabel, name: ctype, response: JSON.stringify(cond).slice(0, 80) });
    }
  }
  return rows;
}

function parseVisibility(vis, lookup) {
  const conditions = vis.conditions || [];
  const rawOverride = vis.override || "default";
  const normallyVisible = vis.normallyVisible !== false;
  const quantifier = vis.allConditionsNeeded ? "all" : "any";

  let direction;
  if (rawOverride === "show" && !conditions.length) direction = "Show";
  else if (rawOverride === "hide" && !conditions.length) direction = "Hide";
  else if (["show", "hide"].includes(rawOverride) && conditions.length)
    direction = `${rawOverride === "hide" ? "Hide" : "Show"} when ${quantifier}`;
  else if (conditions.length)
    direction = `${normallyVisible ? "Hide" : "Show"} when ${quantifier}`;
  else return null;

  return { direction, conditions: conditions.length ? flattenConditions(conditions, lookup) : [] };
}

function effectiveVisibility(section, byId) {
  let vis = section.visibility || {};
  if ((vis.conditions || []).length) return vis;
  let parentId = section.parent || "";
  const visited = new Set([section.id || ""]);
  while (parentId && byId[parentId]) {
    if (visited.has(parentId)) break;
    visited.add(parentId);
    const parent = byId[parentId];
    const pvis = parent.visibility || {};
    if ((pvis.conditions || []).length) return pvis;
    parentId = parent.parent || "";
  }
  return vis;
}

// ── SECTION TREE ────────────────────────────────────────────────────────────

const SKIP_TYPES = new Set(["settings", "toc"]);

function getTitle(section) {
  let raw = (section.title || (section.titles || {}).en || "").trim();
  let title = raw ? stripHtml(raw) : "";
  if (!title || title === "Note") {
    const spec = section.specification || {};
    const specTitle = spec.title || (spec.titles || {}).en || "";
    if (specTitle) title = stripHtml(specTitle);
  }
  return title;
}

function buildSectionTree(sections, documentId) {
  const byId = {};
  for (const s of sections) byId[s.id] = s;
  const childrenByParent = {};
  for (const s of sections) {
    const p = s.parent || "";
    (childrenByParent[p] = childrenByParent[p] || []).push(s);
  }
  for (const pid of Object.keys(childrenByParent)) {
    childrenByParent[pid].sort((a, b) => (a.order || "") < (b.order || "") ? -1 : (a.order || "") > (b.order || "") ? 1 : 0);
  }

  const result = [];
  function visit(sectionId, level) {
    const section = byId[sectionId];
    if (!section || SKIP_TYPES.has(section.type || "")) return;
    result.push({ section, level });
    for (const child of childrenByParent[sectionId] || []) visit(child.id, level + 1);
  }

  let roots = childrenByParent[documentId] || [];
  if (!roots.length) {
    roots = sections.filter(s => !byId[s.parent || ""]);
    roots.sort((a, b) => (a.order || "") < (b.order || "") ? -1 : (a.order || "") > (b.order || "") ? 1 : 0);
  }
  for (const root of roots) visit(root.id, 0);
  return result;
}

// ── EXTRACT LETTER DATA ─────────────────────────────────────────────────────

async function extractLetterData(url) {
  const { host, tenant, engagementId, documentId } = parseUrl(url);
  if (!documentId) throw new Error("URL must include a document fragment (#/letter/<id>)");

  const envPrefix = envPrefixFromHost(host);
  const headers = await makeHeaders(envPrefix, host, tenant);

  // Find document name
  const documents = await fetchDocuments(headers, engagementId, host, tenant);
  let docName = "Letter";
  for (const d of documents) {
    if (d.id === documentId) {
      docName = (d.names || {}).en || d.name || "Letter";
      break;
    }
  }

  // Fetch sections
  const sections = await fetchSections(headers, engagementId, documentId, host, tenant);
  if (!sections.length) throw new Error(`No sections found for document ${documentId}`);

  const byId = {};
  for (const s of sections) byId[s.id] = s;

  // Resolve visibility IDs
  const lookup = await buildIdLookup(headers, engagementId, sections, host, tenant);

  // Build tree
  const tree = buildSectionTree(sections, documentId);

  // Convert to JSON structure
  const jsonSections = tree.map(({ section, level }) => {
    const spec = section.specification || {};
    const fmap = buildFormulaMap(section);
    const rawHtml = spec.content || "";
    const guidances = section.guidances || {};
    const guidanceHtml = guidances.en || section.guidance || "";
    const vis = effectiveVisibility(section, byId);

    return {
      id: section.id || "",
      type: section.type || "",
      title: getTitle(section),
      level,
      html_content: rawHtml ? preprocessHtml(rawHtml, fmap) : "",
      guidance_html: guidanceHtml ? preprocessHtml(guidanceHtml, fmap) : "",
      visibility: parseVisibility(vis, lookup),
    };
  });

  return { document_name: docName, sections: jsonSections };
}

// ═══════════════════════════════════════════════════════════════════════════
// DOCX GENERATION (ported from tools/generate_docx.js)
// ═══════════════════════════════════════════════════════════════════════════

const PAGE_WIDTH = 12240;
const PAGE_HEIGHT = 15840;
const MARGIN = 1440;
const CONTENT_WIDTH = PAGE_HEIGHT - 2 * MARGIN; // landscape

const COLORS = {
  GREY: "888888", PLACEHOLDER_BLUE: "0971AA", DARK_NAVY: "1F3864",
  WHITE: "FFFFFF", LIGHT_BLUE: "D9E1F2",
  GUIDANCE_BG: "FFF3CD", GUIDANCE_BORDER: "FFE69C", GUIDANCE_TEXT: "664D03",
};

const THIN_BORDER = { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" };
const CELL_BORDERS = { top: THIN_BORDER, bottom: THIN_BORDER, left: THIN_BORDER, right: THIN_BORDER };
const CELL_MARGINS = { top: 80, bottom: 80, left: 120, right: 120 };

// ── HTML → DOCX ────────────────────────────────────────────────────────────

let numberingConfigs = [];
let listCounter = 0;

function resetNumbering() { numberingConfigs = []; listCounter = 0; }

function detectListFormat($list) {
  const style = $list.attr("style") || "";
  const m = style.match(/list-style-type:\s*([^;"]+)/);
  const t = m ? m[1].trim() : "";
  if (t === "lower-alpha" || t === "lower-latin") return LevelFormat.LOWER_LETTER;
  if (t === "upper-alpha" || t === "upper-latin") return LevelFormat.UPPER_LETTER;
  if (t === "lower-roman") return LevelFormat.LOWER_ROMAN;
  if (t === "upper-roman") return LevelFormat.UPPER_ROMAN;
  return null;
}

function createListRef(isBullet, formatOverride) {
  listCounter++;
  const ref = isBullet ? `bullets_${listCounter}` : `numbers_${listCounter}`;
  if (isBullet) {
    numberingConfigs.push({ reference: ref, levels: [
      { level: 0, format: LevelFormat.BULLET, text: "\u2022", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } },
      { level: 1, format: LevelFormat.BULLET, text: "\u25E6", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 1440, hanging: 360 } } } },
      { level: 2, format: LevelFormat.BULLET, text: "\u25AA", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 2160, hanging: 360 } } } },
      { level: 3, format: LevelFormat.BULLET, text: "\u2022", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 2880, hanging: 360 } } } },
    ]});
  } else {
    const l0 = formatOverride || LevelFormat.DECIMAL;
    numberingConfigs.push({ reference: ref, levels: [
      { level: 0, format: l0, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } },
      { level: 1, format: LevelFormat.LOWER_LETTER, text: "%2.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 1440, hanging: 360 } } } },
      { level: 2, format: LevelFormat.LOWER_ROMAN, text: "%3.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 2160, hanging: 360 } } } },
      { level: 3, format: LevelFormat.DECIMAL, text: "%4.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 2880, hanging: 360 } } } },
    ]});
  }
  return ref;
}

function parseHtmlToDocx(html) {
  if (!html || !html.trim()) return [];
  const $ = cheerio.load(html, { xmlMode: false, decodeEntities: true });
  const elements = [];
  const listCtx = {};
  const body = $.root();

  function processBlock(el) {
    const tag = el.tagName ? el.tagName.toLowerCase() : "";
    if (el.type === "text") {
      const t = $(el).text().trim();
      if (t) elements.push(new Paragraph({ spacing: { after: 120 }, children: [new TextRun({ text: t })] }));
      return;
    }
    if (tag === "p" || tag === "div") {
      const runs = inlineChildren($(el), $);
      if (runs.length) elements.push(new Paragraph({ spacing: { after: 120 }, children: runs }));
      else elements.push(new Paragraph({ spacing: { after: 60 }, children: [] }));
    } else if (tag === "ul" || tag === "ol") {
      const $l = $(el);
      if (tag === "ul") {
        elements.push(...processList($l, $, createListRef(true, null), 0));
      } else {
        const startAttr = parseInt($l.attr("start") || "1", 10);
        const sm = ($l.attr("style") || "").match(/list-style-type:\s*([^;"]+)/);
        const st = sm ? sm[1].trim() : "decimal";
        let ref;
        if (startAttr > 1 && listCtx[st]) ref = listCtx[st];
        else { ref = createListRef(false, detectListFormat($l)); listCtx[st] = ref; }
        elements.push(...processList($l, $, ref, 0));
      }
    } else if (tag === "table") {
      elements.push(processTable($(el), $));
    } else if (/^h[1-6]$/.test(tag)) {
      const lvl = parseInt(tag[1]);
      const hl = [HeadingLevel.HEADING_1, HeadingLevel.HEADING_2, HeadingLevel.HEADING_3,
                   HeadingLevel.HEADING_4, HeadingLevel.HEADING_5, HeadingLevel.HEADING_6];
      elements.push(new Paragraph({ heading: hl[Math.min(lvl - 1, 5)], children: inlineChildren($(el), $) }));
    } else if (tag !== "br") {
      $(el).contents().each((_, c) => processBlock(c));
    }
  }

  body.contents().each((_, el) => processBlock(el));
  return elements;
}

function inlineChildren($el, $, style = {}, _ctx) {
  const ctx = _ctx || { lastText: "" };
  const runs = [];

  $el.contents().each((_, node) => {
    if (node.type === "text") {
      let text = $(node).text().replace(/\s+/g, " ");
      if (text.trim() !== "") {
        if (ctx.lastText.endsWith(" ") && text.startsWith(" ")) text = text.trimStart();
        if (text) { runs.push(mkRun(text, style)); ctx.lastText = text; }
      } else if (text === " " && ctx.lastText && !ctx.lastText.endsWith(" ")) {
        runs.push(mkRun(" ", style)); ctx.lastText = " ";
      }
      return;
    }
    const tag = node.tagName ? node.tagName.toLowerCase() : "";
    const $n = $(node);
    if (tag === "br") { runs.push(new TextRun({ break: 1 })); ctx.lastText = ""; }
    else if (tag === "strong" || tag === "b") runs.push(...inlineChildren($n, $, { ...style, bold: true }, ctx));
    else if (tag === "em" || tag === "i") runs.push(...inlineChildren($n, $, { ...style, italics: true }, ctx));
    else if (tag === "u") runs.push(...inlineChildren($n, $, { ...style, underline: {} }, ctx));
    else if (tag === "s" || tag === "strike" || tag === "del") runs.push(...inlineChildren($n, $, { ...style, strike: true }, ctx));
    else if (tag === "sub") runs.push(...inlineChildren($n, $, { ...style, subScript: true }, ctx));
    else if (tag === "sup") runs.push(...inlineChildren($n, $, { ...style, superScript: true }, ctx));
    else if (tag === "a") runs.push(...inlineChildren($n, $, { ...style, underline: {}, color: "0563C1" }, ctx));
    else if (tag === "span") {
      if ($n.hasClass("cw-placeholder") || $n.attr("placeholder")) {
        const t = $n.text().trim();
        if (t) { const d = t.startsWith("((") ? t : `((${t}))`; runs.push(mkRun(d, { ...style, color: COLORS.PLACEHOLDER_BLUE })); ctx.lastText = d; }
      } else if ($n.hasClass("cw-formula")) {
        const t = $n.text().trim();
        if (t) { runs.push(mkRun(t, { ...style, color: COLORS.GREY })); ctx.lastText = t; }
      } else if ($n.attr("formula") || $n.hasClass("dynamic-text")) {
        const t = $n.text().trim();
        if (t) { runs.push(mkRun(`[[${t}]]`, { ...style, color: COLORS.GREY })); ctx.lastText = `[[${t}]]`; }
      } else if ($n.hasClass("firm-logo") || $n.attr("type") === "firm-logo") {
        runs.push(mkRun("((Firm Logo))", { ...style, color: COLORS.PLACEHOLDER_BLUE })); ctx.lastText = "((Firm Logo))";
      } else if ($n.hasClass("caret") || $n.hasClass("hidden-print")) { /* skip */ }
      else runs.push(...inlineChildren($n, $, style, ctx));
    } else if (tag === "img") {
      const title = $n.attr("title") || $n.attr("alt") || "Image";
      runs.push(mkRun(`((${title}))`, { ...style, color: COLORS.PLACEHOLDER_BLUE })); ctx.lastText = `((${title}))`;
    } else runs.push(...inlineChildren($n, $, style, ctx));
  });
  return runs;
}

function mkRun(text, s) {
  const o = { text };
  if (s.bold) o.bold = true; if (s.italics) o.italics = true;
  if (s.underline) o.underline = {}; if (s.strike) o.strike = true;
  if (s.color) o.color = s.color; if (s.subScript) o.subScript = true;
  if (s.superScript) o.superScript = true; if (s.size) o.size = s.size;
  return new TextRun(o);
}

function processList($list, $, ref, level) {
  const items = [];
  $list.children("li").each((_, li) => {
    const $li = $(li);
    const nested = $li.children("ul, ol");
    const $clone = $li.clone(); $clone.children("ul, ol").remove();
    const runs = inlineChildren($clone, $);
    if (runs.length) items.push(new Paragraph({ numbering: { reference: ref, level: Math.min(level, 3) }, children: runs }));
    nested.each((_, n) => {
      const $n = $(n); const nt = n.tagName.toLowerCase();
      const ib = nt !== "ol"; const nf = !ib ? detectListFormat($n) : null;
      items.push(...processList($n, $, createListRef(ib, nf), level + 1));
    });
  });
  return items;
}

function processTable($table, $) {
  let maxCols = 0;
  $table.find("tr").each((_, tr) => { maxCols = Math.max(maxCols, $(tr).children("td, th").length); });
  if (!maxCols) return new Paragraph({ children: [] });
  const cw = Math.floor(CONTENT_WIDTH / maxCols);
  const colWidths = Array(maxCols).fill(cw);
  colWidths[maxCols - 1] = CONTENT_WIDTH - cw * (maxCols - 1);
  const rows = [];
  $table.find("tr").each((_, tr) => {
    const cells = [];
    const $c = $(tr).children("td, th");
    for (let i = 0; i < maxCols; i++) {
      const r = $c.eq(i).length ? inlineChildren($c.eq(i), $) : [];
      cells.push(new TableCell({ borders: CELL_BORDERS, margins: CELL_MARGINS, width: { size: colWidths[i], type: WidthType.DXA }, children: [new Paragraph({ children: r })] }));
    }
    rows.push(new TableRow({ children: cells }));
  });
  return new Table({ width: { size: CONTENT_WIDTH, type: WidthType.DXA }, columnWidths: colWidths, rows });
}

// ── GUIDANCE / VISIBILITY RENDERING ─────────────────────────────────────────

function renderGuidance(html) {
  const inner = parseHtmlToDocx(html);
  return new Table({
    width: { size: CONTENT_WIDTH, type: WidthType.DXA }, columnWidths: [CONTENT_WIDTH],
    rows: [new TableRow({ children: [new TableCell({
      borders: {
        top: { style: BorderStyle.SINGLE, size: 1, color: COLORS.GUIDANCE_BORDER },
        bottom: { style: BorderStyle.SINGLE, size: 1, color: COLORS.GUIDANCE_BORDER },
        left: { style: BorderStyle.SINGLE, size: 6, color: COLORS.GUIDANCE_BORDER },
        right: { style: BorderStyle.SINGLE, size: 1, color: COLORS.GUIDANCE_BORDER },
      },
      margins: { top: 120, bottom: 120, left: 200, right: 200 },
      shading: { fill: COLORS.GUIDANCE_BG, type: ShadingType.CLEAR },
      width: { size: CONTENT_WIDTH, type: WidthType.DXA },
      children: [
        new Paragraph({ spacing: { after: 60 }, children: [new TextRun({ text: "Guidance", bold: true, color: COLORS.GUIDANCE_TEXT, size: 20 })] }),
        ...inner,
      ],
    })]})],
  });
}

function renderVisibilityInline(vis) {
  if (!vis) return [];
  const els = [
    new Paragraph({ spacing: { before: 60, after: 40 }, indent: { left: 360 },
      children: [new TextRun({ text: `Visibility: ${vis.direction}`, italics: true, color: COLORS.GREY, size: 18 })] }),
  ];
  for (const c of vis.conditions || []) {
    const parts = [];
    if (c.group) parts.push(c.group);
    if (c.name) parts.push(c.name);
    if (c.response) parts.push(`= ${c.response}`);
    els.push(new Paragraph({ spacing: { after: 20 }, indent: { left: 720 },
      children: [new TextRun({ text: `- ${parts.join(" > ")}`, italics: true, color: COLORS.GREY, size: 16 })] }));
  }
  els.push(new Paragraph({ spacing: { after: 80 }, children: [] }));
  return els;
}

function renderVisibilitySummaryTable(sections) {
  const vis = sections.filter(s => s.visibility && (s.visibility.conditions || []).length > 0);
  if (!vis.length) return [];
  const cw = [4000, 2500, CONTENT_WIDTH - 6500];
  const hb = { style: BorderStyle.SINGLE, size: 1, color: COLORS.DARK_NAVY };
  const hBorders = { top: hb, bottom: hb, left: hb, right: hb };
  const headerRow = new TableRow({ children: ["Section", "Visibility Setting", "Conditions"].map((t, i) =>
    new TableCell({ borders: hBorders, margins: CELL_MARGINS, width: { size: cw[i], type: WidthType.DXA },
      shading: { fill: COLORS.DARK_NAVY, type: ShadingType.CLEAR },
      children: [new Paragraph({ children: [new TextRun({ text: t, bold: true, color: COLORS.WHITE, size: 20 })] })] })) });
  const dataRows = vis.map((s, idx) => {
    const v = s.visibility;
    const conds = (v.conditions || []).map(c => { const p = []; if (c.name) p.push(c.name); if (c.response) p.push(`= ${c.response}`); return p.join(" "); }).join("\n");
    const bg = idx % 2 === 1 ? COLORS.LIGHT_BLUE : COLORS.WHITE;
    return new TableRow({ children: [
      new TableCell({ borders: CELL_BORDERS, margins: CELL_MARGINS, width: { size: cw[0], type: WidthType.DXA }, shading: { fill: bg, type: ShadingType.CLEAR }, children: [new Paragraph({ children: [new TextRun({ text: s.title || "", size: 20 })] })] }),
      new TableCell({ borders: CELL_BORDERS, margins: CELL_MARGINS, width: { size: cw[1], type: WidthType.DXA }, shading: { fill: bg, type: ShadingType.CLEAR }, children: [new Paragraph({ children: [new TextRun({ text: v.direction || "", size: 20 })] })] }),
      new TableCell({ borders: CELL_BORDERS, margins: CELL_MARGINS, width: { size: cw[2], type: WidthType.DXA }, shading: { fill: bg, type: ShadingType.CLEAR }, children: conds.split("\n").map(l => new Paragraph({ children: [new TextRun({ text: l, size: 18 })] })) }),
    ]});
  });
  return [
    new Paragraph({ heading: HeadingLevel.HEADING_1, spacing: { before: 240, after: 240 }, children: [new TextRun({ text: "Visibility Settings Summary" })] }),
    new Table({ width: { size: CONTENT_WIDTH, type: WidthType.DXA }, columnWidths: cw, rows: [headerRow, ...dataRows] }),
  ];
}

// ── BUILD DOCUMENT ──────────────────────────────────────────────────────────

function buildDocument(data) {
  resetNumbering();
  const { document_name, sections } = data;
  const children = [];
  children.push(new Paragraph({ heading: HeadingLevel.HEADING_1, spacing: { after: 240 }, children: [new TextRun({ text: document_name })] }));

  for (const s of sections) {
    const { type, title, level, html_content, guidance_html, visibility } = s;
    if (type === "settings" || type === "toc") continue;
    if (type === "pagebreak") { children.push(new Paragraph({ children: [new PageBreak()] })); continue; }

    if (type === "grouping" || type === "heading") {
      if (title) {
        const hl = level <= 0 ? HeadingLevel.HEADING_1 : level === 1 ? HeadingLevel.HEADING_2 : HeadingLevel.HEADING_3;
        children.push(new Paragraph({ heading: hl, spacing: { before: 240, after: 120 }, children: [new TextRun({ text: title })] }));
      }
      if (guidance_html) { children.push(renderGuidance(guidance_html)); children.push(new Paragraph({ spacing: { after: 120 }, children: [] })); }
      if (visibility && (visibility.conditions || []).length) children.push(...renderVisibilityInline(visibility));
      continue;
    }

    if (type === "content") {
      if (title) {
        const hl = level <= 1 ? HeadingLevel.HEADING_2 : level === 2 ? HeadingLevel.HEADING_3 : HeadingLevel.HEADING_4;
        children.push(new Paragraph({ heading: hl, spacing: { before: 200, after: 100 }, children: [new TextRun({ text: title })] }));
      }
      if (html_content) children.push(...parseHtmlToDocx(html_content));
      if (guidance_html) {
        children.push(new Paragraph({ spacing: { before: 80 }, children: [] }));
        children.push(renderGuidance(guidance_html));
        children.push(new Paragraph({ spacing: { after: 120 }, children: [] }));
      }
      if (visibility && (visibility.conditions || []).length) children.push(...renderVisibilityInline(visibility));
      continue;
    }

    if (title) children.push(new Paragraph({ spacing: { before: 120, after: 60 }, children: [new TextRun({ text: title, bold: true })] }));
    if (html_content) children.push(...parseHtmlToDocx(html_content));
  }

  const summary = renderVisibilitySummaryTable(sections);
  if (summary.length) { children.push(new Paragraph({ children: [new PageBreak()] })); children.push(...summary); }

  return new Document({
    styles: {
      default: { document: { run: { font: "Arial", size: 24 } } },
      paragraphStyles: [
        { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true, run: { size: 32, bold: true, font: "Arial" }, paragraph: { spacing: { before: 240, after: 240 }, outlineLevel: 0 } },
        { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true, run: { size: 28, bold: true, font: "Arial" }, paragraph: { spacing: { before: 180, after: 180 }, outlineLevel: 1 } },
        { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true, run: { size: 24, bold: true, font: "Arial" }, paragraph: { spacing: { before: 120, after: 120 }, outlineLevel: 2 } },
        { id: "Heading4", name: "Heading 4", basedOn: "Normal", next: "Normal", quickFormat: true, run: { size: 22, bold: true, font: "Arial" }, paragraph: { spacing: { before: 120, after: 60 }, outlineLevel: 3 } },
      ],
    },
    numbering: { config: numberingConfigs },
    sections: [{
      properties: { page: {
        size: { width: PAGE_WIDTH, height: PAGE_HEIGHT, orientation: PageOrientation.LANDSCAPE },
        margin: { top: MARGIN, right: MARGIN, bottom: MARGIN, left: MARGIN },
      }},
      children,
    }],
  });
}

// ═══════════════════════════════════════════════════════════════════════════
// VERCEL HANDLER
// ═══════════════════════════════════════════════════════════════════════════

module.exports = async (req, res) => {
  if (req.method !== "POST") {
    return res.status(405).json({ error: "Method not allowed" });
  }

  const { url, templateName } = req.body || {};
  if (!url) return res.status(400).json({ error: "URL is required." });

  const urlMatch = url.match(CW_URL_PATTERN);
  if (!urlMatch) return res.status(400).json({ error: "Invalid CaseWare URL." });

  const docMatch = url.match(CW_DOC_PATTERN);
  if (!docMatch) return res.status(400).json({ error: "URL must include a document fragment (e.g. #/letter/<documentId>)" });

  try {
    const data = await extractLetterData(url);
    const doc = buildDocument(data);
    const buffer = await Packer.toBuffer(doc);

    const safeName = (templateName || "Letter").replace(/[^\w\s-]/g, "").trim().replace(/ /g, "_");
    const filename = safeName ? `${safeName}_letter.docx` : "letter.docx";

    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
    res.setHeader("Content-Disposition", `attachment; filename="${filename}"`);
    return res.status(200).send(buffer);
  } catch (e) {
    console.error("Generate error:", e);
    const status = e.message.includes("401") ? 502 : e.message.includes("No sections") ? 422 : 500;
    return res.status(status).json({ error: e.message });
  }
};
