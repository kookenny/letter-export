"""
letter_extract.py
─────────────────
Extracts letter content from CaseWare Cloud SE author templates.
Outputs structured JSON that is consumed by generate_docx.js to produce
a formatted Word document.

Usage:
    python tools/letter_extract.py --url "<caseware-url>"            # generate .docx
    python tools/letter_extract.py --url "<caseware-url>" --discover # dump raw JSON
"""

import argparse
import json
import logging
import os
import re
import subprocess
import sys
from html import unescape
from pathlib import Path
from typing import Optional

import requests
from dotenv import load_dotenv

# Load .env from project root
PROJECT_ROOT = Path(__file__).resolve().parent.parent
load_dotenv(PROJECT_ROOT / ".env")

logging.basicConfig(
    level=os.environ.get("CW_LOG_LEVEL", "INFO").upper(),
    format="%(levelname)s  %(message)s",
)
log = logging.getLogger(__name__)


# ── URL PARSING ──────────────────────────────────────────────────────────────

CW_URL_PATTERN = re.compile(r"https?://([^/]+)/([^/]+)/e/eng/([^/]+)")
CW_DOC_PATTERN = re.compile(r"#/(?:efinancials|letter)/([^/?\s]+)")


def parse_url(url: str) -> tuple[str, str, str, str]:
    """Extract (host, tenant, engagement_id, document_id) from a CaseWare URL."""
    match = CW_URL_PATTERN.search(url)
    if not match:
        raise ValueError(f"Invalid CaseWare URL: {url}")
    host = f"https://{match.group(1)}"
    tenant = match.group(2)
    engagement_id = match.group(3)
    doc_match = CW_DOC_PATTERN.search(url)
    document_id = doc_match.group(1) if doc_match else ""
    return host, tenant, engagement_id, document_id


# ── HTML STRIPPING ───────────────────────────────────────────────────────────

def strip_html(html: str, formula_map: dict[str, str] | None = None) -> str:
    """Convert HTML to plain text. Placeholders → (( )), formulas → [[ ]]."""
    if not html:
        return ""
    text = re.sub(
        r'<span[^>]*\bplaceholder="[^"]*"[^>]*>(.*?)</span>',
        r"((\1))",
        html,
    )
    if formula_map:
        def _resolve_formula(m: re.Match) -> str:
            fid = m.group(1)
            val = formula_map.get(fid, "")
            return f"[[{val}]]" if val else ""
        text = re.sub(
            r'<span[^>]*\bformula="([^"]*)"[^>]*>.*?</span>',
            _resolve_formula,
            text,
        )
    text = re.sub(r"<[^>]+>", " ", text)
    text = unescape(text)
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def build_formula_map(section: dict) -> dict[str, str]:
    """Build a {referenceId: calculated_value} map from a section's attachables."""
    result = {}
    for a in (section.get("attachables") or {}).values():
        ref_id = a.get("referenceId")
        calc = (a.get("calculated") or "").strip()
        if ref_id and calc:
            result[ref_id] = calc
    return result


def preprocess_html(html: str, formula_map: dict[str, str] | None = None) -> str:
    """Preprocess CaseWare HTML for Word conversion.

    Resolves placeholder and formula spans to marked-up text while
    preserving the surrounding HTML structure for the docx converter.
    """
    if not html:
        return ""
    # Wrap placeholder spans: keep as styled text
    text = re.sub(
        r'<span[^>]*\bplaceholder="[^"]*"[^>]*>(.*?)</span>',
        r'<span class="cw-placeholder">((\1))</span>',
        html,
    )
    # Resolve formula spans
    if formula_map:
        def _resolve_formula(m: re.Match) -> str:
            fid = m.group(1)
            val = formula_map.get(fid, "")
            return f'<span class="cw-formula">[[{val}]]</span>' if val else ""
        text = re.sub(
            r'<span[^>]*\bformula="([^"]*)"[^>]*>.*?</span>',
            _resolve_formula,
            text,
        )
    # Collapse runs of whitespace to a single space everywhere
    # (prevents double spaces when formula/placeholder spans resolve to empty)
    # Step 1: between/around tags
    text = re.sub(r'>\s{2,}<', '> <', text)
    text = re.sub(r'>\s{2,}', '> ', text)
    text = re.sub(r'\s{2,}<', ' <', text)
    # Step 2: within text content segments (between tags)
    text = re.sub(r'(?<=>)[^<]+(?=<)', lambda m: re.sub(r' {2,}', ' ', m.group(0)), text)
    # Step 3: at the very start/end (text not between tags)
    text = re.sub(r'^[^<]+', lambda m: re.sub(r' {2,}', ' ', m.group(0)), text)
    text = re.sub(r'>[^<]+$', lambda m: re.sub(r' {2,}', ' ', m.group(0)), text)
    return text


# ── SESSION ──────────────────────────────────────────────────────────────────

def _env_prefix_from_host(host: str) -> str:
    """Derive env-var prefix from hostname: 'ca.cwcloudpartner.com' → 'CA'."""
    hostname = host.replace("https://", "").replace("http://", "").split("/")[0]
    return hostname.split(".")[0].upper()


def _obtain_bearer_token(env_prefix: str = "",
                         host: str | None = None,
                         tenant: str | None = None) -> str | None:
    """Exchange CW_CLIENT_ID + CW_CLIENT_SECRET for a Bearer token via OAuth."""
    client_id, client_secret = "", ""
    if env_prefix:
        client_id = os.environ.get(f"CW_{env_prefix}_CLIENT_ID", "").strip()
        client_secret = os.environ.get(f"CW_{env_prefix}_CLIENT_SECRET", "").strip()
    if not client_id or not client_secret:
        client_id = os.environ.get("CW_CLIENT_ID", "").strip()
        client_secret = os.environ.get("CW_CLIENT_SECRET", "").strip()
    if not client_id or not client_secret:
        return None
    url = f"{host}/{tenant}/ms/caseware-cloud/api/v1/auth/token"
    resp = requests.post(url, json={
        "ClientId": client_id, "ClientSecret": client_secret, "Language": "en",
    }, headers={"Accept": "application/json", "Content-Type": "application/json"},
       timeout=15)
    resp.raise_for_status()
    token = resp.json().get("Token")
    if token:
        log.info("Authenticated via OAuth (Bearer token)")
    return token


def make_session(env_prefix: str = "",
                 host: str | None = None,
                 tenant: str | None = None) -> requests.Session:
    """Build a requests.Session using OAuth (preferred) or browser cookies."""
    session = requests.Session()
    session.headers.update({
        "Accept": "application/json",
        "Content-Type": "application/json",
    })
    token = _obtain_bearer_token(env_prefix, host=host, tenant=tenant)
    if token:
        session.headers["Authorization"] = f"Bearer {token}"
        return session
    cookies = os.environ.get("CW_COOKIES", "").strip()
    if not cookies:
        raise RuntimeError(
            "No auth credentials found. "
            "Set CW_CLIENT_ID + CW_CLIENT_SECRET for OAuth, "
            "or CW_COOKIES for cookie auth."
        )
    session.headers["Cookie"] = cookies
    log.info("Authenticated via browser cookies")
    return session


# ── API HELPERS ──────────────────────────────────────────────────────────────

def _unwrap_response(data) -> list:
    """Handle inconsistent API response wrappers."""
    if isinstance(data, list):
        return data
    if isinstance(data, dict):
        for key in ("objects", "sections", "items", "results", "data"):
            if key in data and isinstance(data[key], list):
                return data[key]
        for key, val in data.items():
            if isinstance(val, list):
                return val
        if "object" in data and isinstance(data["object"], dict):
            return [data["object"]]
    return []


def _api_post(session: requests.Session, url: str, payload: dict,
              timeout: int = 30) -> list:
    """POST to a CaseWare API endpoint and unwrap the response."""
    resp = session.post(url, json=payload, timeout=timeout)
    if resp.status_code == 401:
        raise RuntimeError("401 Unauthorised — credentials have expired.")
    if not resp.ok:
        raise RuntimeError(f"{resp.status_code} from {url}\n{resp.text[:500]}")
    return _unwrap_response(resp.json())


# ── DATA FETCHING ────────────────────────────────────────────────────────────

def fetch_documents(session: requests.Session,
                    engagement_id: str,
                    host: str, tenant: str) -> list[dict]:
    """Fetch all documents in the engagement."""
    url = f"{host}/{tenant}/e/eng/{engagement_id}/api/v1.12.0/document/get"
    return _api_post(session, url, {})


def fetch_sections(session: requests.Session,
                   engagement_id: str,
                   document_id: str,
                   host: str, tenant: str) -> list[dict]:
    """Fetch all sections belonging to a document."""
    url = f"{host}/{tenant}/e/eng/{engagement_id}/api/v1.12.0/section/get"
    payload = {
        "filter": {
            "filter": {
                "node": "=",
                "left": {"node": "field", "kind": "section", "field": "document"},
                "right": {"node": "string", "value": document_id},
            }
        }
    }
    log.info("Fetching sections from %s", url)
    return _api_post(session, url, payload)


def fetch_document_lookup(session: requests.Session,
                          engagement_id: str,
                          host: str, tenant: str) -> dict[str, str]:
    """Fetch all documents and return {id: "number name", content: "number name"} map."""
    url = f"{host}/{tenant}/e/eng/{engagement_id}/api/v1.12.0/document/get"
    try:
        documents = _api_post(session, url, {})
        result = {}
        for doc in documents:
            did = doc.get("id", "")
            number = doc.get("number", "")
            names = doc.get("names") or {}
            name = names.get("en", "") or doc.get("name", "")
            content = doc.get("content", "")
            if did:
                label = f"{number} {name}".strip() if number else name
                result[did] = label
                if content:
                    result[content] = label
        return result
    except Exception as exc:
        log.warning("Could not fetch document list: %s", exc)
        return {}


def fetch_procedures_for_checklist(session: requests.Session,
                                   engagement_id: str,
                                   checklist_id: str,
                                   host: str, tenant: str) -> list[dict]:
    """Fetch all procedures belonging to a checklist."""
    url = f"{host}/{tenant}/e/eng/{engagement_id}/api/v1.12.0/procedure/get"
    payload = {"filter": {"filter": {
        "node": "=",
        "left": {"node": "field", "kind": "procedure", "field": "checklistId"},
        "right": {"node": "string", "value": checklist_id},
    }}}
    try:
        resp = session.post(url, json=payload, timeout=30)
        if not resp.ok:
            return []
        return resp.json().get("objects", []) if isinstance(resp.json(), dict) else []
    except Exception:
        return []


def fetch_procedure_by_id(session: requests.Session,
                          engagement_id: str,
                          procedure_id: str,
                          host: str, tenant: str) -> Optional[dict]:
    """Fetch a single procedure by its ID."""
    url = f"{host}/{tenant}/e/eng/{engagement_id}/api/v1.12.0/procedure/get"
    payload = {"filter": {"filter": {
        "node": "=",
        "left": {"node": "field", "kind": "procedure", "field": "id"},
        "right": {"node": "string", "value": procedure_id},
    }}}
    try:
        resp = session.post(url, json=payload, timeout=15)
        if not resp.ok:
            return None
        data = resp.json()
        objects = data.get("objects", [])
        return objects[0] if objects else None
    except Exception:
        return None


# ── VISIBILITY ID RESOLUTION ─────────────────────────────────────────────────

def _collect_procedure_ids_from_cond(cond: dict, ids: set) -> None:
    pid = (cond.get("procedureId") or {}).get("id")
    if pid:
        ids.add(pid)
    for sub in cond.get("conditions") or []:
        _collect_procedure_ids_from_cond(sub, ids)


def _collect_checklist_ids_from_cond(cond: dict, ids: set) -> None:
    cid = (cond.get("checklistId") or {}).get("id")
    if cid:
        ids.add(cid)
    for sub in cond.get("conditions") or []:
        _collect_checklist_ids_from_cond(sub, ids)


def build_id_lookup(session: requests.Session,
                    engagement_id: str,
                    sections: list[dict],
                    host: str, tenant: str) -> dict[str, str]:
    """Resolve procedure/response/checklist IDs referenced in visibility conditions."""
    procedure_ids: set = set()
    checklist_ids: set = set()
    for s in sections:
        for cond in (s.get("visibility") or {}).get("conditions") or []:
            _collect_procedure_ids_from_cond(cond, procedure_ids)
            _collect_checklist_ids_from_cond(cond, checklist_ids)

    log.info("  %d procedure ID(s), %d checklist ID(s) to resolve",
             len(procedure_ids), len(checklist_ids))

    # Seed with document names
    lookup = dict(fetch_document_lookup(session, engagement_id, host, tenant))

    # Fetch checklist-level response sets
    cl_url = f"{host}/{tenant}/e/eng/{engagement_id}/api/v1.12.0/checklist/get"
    for cid in checklist_ids:
        payload = {"filter": {"filter": {
            "node": "=",
            "left": {"node": "field", "kind": "checklist", "field": "id"},
            "right": {"node": "string", "value": cid},
        }}}
        try:
            resp = session.post(cl_url, json=payload, timeout=30)
            if resp.ok:
                for cl in resp.json().get("objects") or []:
                    for rs in (cl.get("settings") or {}).get("responseSets") or []:
                        for resp_opt in rs.get("responses") or []:
                            rid = resp_opt.get("id", "")
                            rname = resp_opt.get("name", "") or (resp_opt.get("names") or {}).get("en", "")
                            if rid and rname:
                                lookup[rid] = rname
        except Exception:
            pass

    # Fetch procedures per checklist for response sets + names
    for cid in checklist_ids:
        procs = fetch_procedures_for_checklist(session, engagement_id, cid, host, tenant)
        for proc in procs:
            pid = proc.get("id", "")
            if pid and pid not in lookup:
                summary = proc.get("summaryNames") or {}
                name = summary.get("en", "") or next(iter(summary.values()), "")
                if not name:
                    name = strip_html(proc.get("text", ""))
                if name:
                    lookup[pid] = name
            for rs in (proc.get("settings") or {}).get("responseSets") or []:
                for resp_opt in rs.get("responses") or []:
                    rid = resp_opt.get("id", "")
                    rname = resp_opt.get("name", "") or (resp_opt.get("names") or {}).get("en", "")
                    if rid and rname:
                        lookup[rid] = rname

    # Individually resolve any remaining procedure IDs
    for pid in procedure_ids:
        if pid in lookup:
            continue
        proc = fetch_procedure_by_id(session, engagement_id, pid, host, tenant)
        if not proc:
            continue
        summary = proc.get("summaryNames") or {}
        proc_text = summary.get("en", "") or next(iter(summary.values()), "")
        if not proc_text:
            proc_text = strip_html(proc.get("text", ""))
        if pid and proc_text:
            lookup[pid] = proc_text
        for rs in (proc.get("settings") or {}).get("responseSets") or []:
            for resp_opt in rs.get("responses") or []:
                rid = resp_opt.get("id", "")
                rname = resp_opt.get("name", "") or (resp_opt.get("names") or {}).get("en", "")
                if rid and rname:
                    lookup[rid] = rname

    log.info("  %d names resolved", len(lookup))
    return lookup


# ── VISIBILITY PARSING ───────────────────────────────────────────────────────

# Organization type labels
_ORG_TYPE_LABELS = {
    "CorporationControlledPrivateCorporation": "Canadian Controlled Private Corporation (CCPC)",
    "CorporationControlledPublicCorporation": "Corporation Controlled by a Public Corporation",
    "OtherPrivateCorporation": "Other Private Corporation",
    "PublicCompany": "Public Company",
    "Individual": "Individual",
    "Coownership": "Co-ownership",
    "GeneralPartnership": "General Partnership",
    "LimitedPartnership": "Limited Partnership",
    "LimitedLiabilityPartnership": "Limited Liability Partnership",
    "JointVenture": "Joint Venture",
    "Trust": "Trust",
    "Cooperative": "Cooperative",
    "PensionFunds": "Pension Funds",
    "RegisteredCharity": "Registered Charity",
    "NotForProfit": "Not For Profit",
    "Government": "Government",
    "LimitedLiabilityCompany": "Limited Liability Company (LLC)",
    "CCorporation": "C-Corporation",
    "SCorporation": "S-Corporation",
    "SoleProprietorship": "Sole Proprietorship",
    "Partnership": "Partnership",
    "NotForProfitPrivate": "Not for Profit - Private",
    "NotForProfitPublic": "Not for Profit - Public",
    "PublicFoundation": "Public Foundation",
    "PrivateFoundation": "Private Foundation",
}


def _split_pascal_case(s: str) -> str:
    return re.sub(r"(?<=[a-z])(?=[A-Z])|(?<=[A-Z])(?=[A-Z][a-z])", " ", s)


def _resolve(lookup: dict, id_obj) -> str:
    if not id_obj:
        return ""
    raw_id = id_obj.get("id", "") if isinstance(id_obj, dict) else str(id_obj)
    return lookup.get(raw_id, raw_id[:12])


def _resolve_org_type(cond: dict) -> str:
    custom_id = cond.get("customOrganizationTypeId", "")
    if custom_id:
        return _ORG_TYPE_LABELS.get(custom_id, _split_pascal_case(custom_id))
    org_type = cond.get("organizationType", "")
    if org_type:
        return _ORG_TYPE_LABELS.get(org_type, _split_pascal_case(org_type))
    return cond.get("id", "unknown")[:12]


def _flatten_conditions(conditions: list, lookup: dict,
                        group_label: str = "") -> list[dict]:
    """Flatten conditions into a list of {group, name, response} dicts."""
    rows = []
    for cond in conditions:
        ctype = cond.get("type", "")
        if ctype == "response":
            checklist = _resolve(lookup, cond.get("checklistId"))
            procedure = _resolve(lookup, cond.get("procedureId"))
            response = _resolve(lookup, cond.get("responseId"))
            rows.append({"group": group_label or checklist,
                         "name": procedure, "response": response})
        elif ctype == "condition_group":
            nested = cond.get("conditions") or []
            group_all = cond.get("allConditionsNeeded", False)
            qualifier = "all" if group_all else "any"
            first_checklist = ""
            if nested:
                first_checklist = _resolve(lookup, nested[0].get("checklistId"))
            label = f"{first_checklist} ({qualifier})" if first_checklist else f"({qualifier})"
            rows.extend(_flatten_conditions(nested, lookup, group_label=label))
        elif ctype == "organization_type":
            org_name = _resolve_org_type(cond)
            rows.append({"group": group_label, "name": "Organization Type",
                         "response": org_name})
        elif ctype == "consolidation":
            label = "Consolidated" if cond.get("consolidated", False) else "Not consolidated"
            rows.append({"group": group_label, "name": "Consolidation",
                         "response": label})
        else:
            rows.append({"group": group_label, "name": ctype,
                         "response": json.dumps(cond)[:80]})
    return rows


def parse_visibility(vis: dict, lookup: dict) -> dict | None:
    """Parse a visibility dict into a structured object for JSON output.

    Returns None if there are no meaningful visibility settings.
    """
    conditions = vis.get("conditions") or []
    raw_override = vis.get("override", "default")
    normally_visible = vis.get("normallyVisible", True)
    all_needed = vis.get("allConditionsNeeded", False)
    quantifier = "all" if all_needed else "any"

    if raw_override == "show" and not conditions:
        direction = "Show"
    elif raw_override == "hide" and not conditions:
        direction = "Hide"
    elif raw_override in ("show", "hide") and conditions:
        direction = f"{'Hide' if raw_override == 'hide' else 'Show'} when {quantifier}"
    elif conditions:
        direction = f"{'Show' if not normally_visible else 'Hide'} when {quantifier}"
    else:
        return None  # default with no conditions — nothing to show

    flat = _flatten_conditions(conditions, lookup) if conditions else []
    return {"direction": direction, "conditions": flat}


def _effective_visibility(section: dict, by_id: dict) -> dict:
    """Walk up the parent chain to find the nearest visibility with conditions."""
    vis = section.get("visibility") or {}
    if vis.get("conditions"):
        return vis
    parent_id = section.get("parent", "")
    visited = {section.get("id", "")}
    while parent_id and parent_id in by_id:
        if parent_id in visited:
            break
        visited.add(parent_id)
        parent = by_id[parent_id]
        pvis = parent.get("visibility") or {}
        if pvis.get("conditions"):
            return pvis
        parent_id = parent.get("parent", "")
    return vis


# ── SECTION TREE ─────────────────────────────────────────────────────────────

_SKIP_TYPES = {"settings", "toc"}
# pagebreak is kept in the tree — the docx generator inserts a Word page break


def get_title(section: dict) -> str:
    """Return display title, checking specification.title for note sections."""
    raw = (section.get("title") or section.get("titles", {}).get("en", "") or "").strip()
    title = strip_html(raw) if raw else ""
    if not title or title == "Note":
        spec = section.get("specification") or {}
        spec_title = spec.get("title") or (spec.get("titles") or {}).get("en", "") or ""
        if spec_title:
            title = strip_html(spec_title)
    return title


def build_section_tree(sections: list[dict], document_id: str) -> list[dict]:
    """Build an ordered tree of sections for the letter, depth-first."""
    by_id = {s["id"]: s for s in sections}
    children_by_parent: dict[str, list[dict]] = {}
    for s in sections:
        parent = s.get("parent", "")
        children_by_parent.setdefault(parent, []).append(s)
    for pid in children_by_parent:
        children_by_parent[pid].sort(key=lambda s: s.get("order", ""))

    result = []

    def visit(section_id: str, level: int):
        section = by_id.get(section_id)
        if not section:
            return
        sec_type = section.get("type", "")
        if sec_type in _SKIP_TYPES:
            return
        result.append((section, level))
        for child in children_by_parent.get(section_id, []):
            visit(child["id"], level + 1)

    # Find root sections (parent == document_id or parent not in by_id)
    roots = children_by_parent.get(document_id, [])
    if not roots:
        # Fallback: sections whose parent is not in by_id
        roots = [s for s in sections if s.get("parent", "") not in by_id]
        roots.sort(key=lambda s: s.get("order", ""))

    for root in roots:
        visit(root["id"], 0)

    return result


# ── EXTRACT TO JSON ──────────────────────────────────────────────────────────

def extract_letter_data(engagement_id: str,
                        document_id: str,
                        host: str,
                        tenant: str) -> dict:
    """Extract letter data and return structured JSON for the docx generator."""
    env_prefix = _env_prefix_from_host(host)
    session = make_session(env_prefix, host=host, tenant=tenant)

    # Find the document name
    documents = fetch_documents(session, engagement_id, host, tenant)
    doc_name = "Letter"
    for d in documents:
        if d.get("id") == document_id:
            names = d.get("names") or {}
            doc_name = names.get("en", "") or d.get("name", "") or "Letter"
            break

    log.info("Document: %s", doc_name)

    # Fetch sections
    sections = fetch_sections(session, engagement_id, document_id, host, tenant)
    log.info("Fetched %d sections", len(sections))

    if not sections:
        raise ValueError(f"No sections found for document {document_id}")

    by_id = {s["id"]: s for s in sections}

    # Build ID lookup for visibility resolution
    log.info("Resolving visibility IDs...")
    lookup = build_id_lookup(session, engagement_id, sections, host, tenant)

    # Build ordered tree
    tree = build_section_tree(sections, document_id)
    log.info("Section tree: %d entries", len(tree))

    # Convert to JSON
    json_sections = []
    for section, level in tree:
        sec_type = section.get("type", "")
        title = get_title(section)
        spec = section.get("specification") or {}

        # Content HTML
        raw_html = spec.get("content", "")
        fmap = build_formula_map(section)
        html_content = preprocess_html(raw_html, fmap) if raw_html else ""

        # Guidance
        guidances = section.get("guidances") or {}
        guidance_html = guidances.get("en", "") or section.get("guidance", "") or ""
        if guidance_html:
            guidance_html = preprocess_html(guidance_html, fmap)

        # Visibility
        vis = _effective_visibility(section, by_id)
        visibility = parse_visibility(vis, lookup)

        json_sections.append({
            "id": section.get("id", ""),
            "type": sec_type,
            "title": title,
            "level": level,
            "html_content": html_content,
            "guidance_html": guidance_html,
            "visibility": visibility,
        })

    return {"document_name": doc_name, "sections": json_sections}


# ── DOCX GENERATION (shells out to Node.js) ──────────────────────────────────

def generate_report_bytes(engagement_id: str,
                          document_id: str,
                          host: str,
                          tenant: str) -> bytes:
    """Extract letter data and generate a Word document. Returns .docx bytes."""
    data = extract_letter_data(engagement_id, document_id, host, tenant)

    # Write JSON to temp file
    tmp_dir = PROJECT_ROOT / ".tmp"
    tmp_dir.mkdir(exist_ok=True)

    json_path = tmp_dir / "letter_data.json"
    output_path = tmp_dir / "letter_output.docx"

    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    log.info("Wrote JSON to %s (%d sections)", json_path, len(data["sections"]))

    # Run Node.js generator
    gen_script = PROJECT_ROOT / "tools" / "generate_docx.js"
    result = subprocess.run(
        ["node", str(gen_script), str(json_path), str(output_path)],
        capture_output=True, text=True, timeout=60,
    )
    if result.returncode != 0:
        log.error("generate_docx.js failed:\n%s", result.stderr)
        raise RuntimeError(f"Word generation failed: {result.stderr[:500]}")

    log.info("Generated %s", output_path)
    return output_path.read_bytes()


# ── DISCOVERY MODE ───────────────────────────────────────────────────────────

def discover(engagement_id: str, document_id: str, host: str, tenant: str):
    """Dump raw API data for letter documents to .tmp/ for inspection."""
    env_prefix = _env_prefix_from_host(host)
    session = make_session(env_prefix, host=host, tenant=tenant)

    tmp_dir = PROJECT_ROOT / ".tmp"
    tmp_dir.mkdir(exist_ok=True)

    # Fetch all documents
    documents = fetch_documents(session, engagement_id, host, tenant)
    docs_path = tmp_dir / "documents.json"
    with open(docs_path, "w", encoding="utf-8") as f:
        json.dump(documents, f, ensure_ascii=False, indent=2)
    log.info("Wrote %d documents to %s", len(documents), docs_path)

    # Show letter-type documents
    print("\n-- Documents --")
    for d in documents:
        dtype = d.get("type", "?")
        did = d.get("id", "?")
        names = d.get("names") or {}
        name = names.get("en", "") or d.get("name", "")
        print(f"  [{dtype}] {did[:16]}... - {name}")

    if not document_id:
        print("\nNo document ID provided. Add #/letter/<id> to the URL to inspect a specific letter.")
        return

    # Fetch sections for the specific document
    sections = fetch_sections(session, engagement_id, document_id, host, tenant)
    secs_path = tmp_dir / "sections.json"
    with open(secs_path, "w", encoding="utf-8") as f:
        json.dump(sections, f, ensure_ascii=False, indent=2)
    log.info("Wrote %d sections to %s", len(sections), secs_path)

    # Summary
    print(f"\n-- Sections ({len(sections)}) --")
    types = {}
    has_guidance = 0
    has_visibility = 0
    for s in sections:
        stype = s.get("type", "?")
        types[stype] = types.get(stype, 0) + 1
        if (s.get("guidances") or {}).get("en") or s.get("guidance"):
            has_guidance += 1
        if (s.get("visibility") or {}).get("conditions"):
            has_visibility += 1

    for t, count in sorted(types.items()):
        print(f"  {t}: {count}")
    print(f"  With guidance: {has_guidance}")
    print(f"  With visibility conditions: {has_visibility}")

    # Show first section with guidance
    print("\n-- Sample section with guidance --")
    for s in sections:
        guidance = (s.get("guidances") or {}).get("en") or s.get("guidance")
        if guidance:
            print(json.dumps({
                "id": s["id"][:16],
                "type": s.get("type"),
                "title": get_title(s),
                "guidance_field": "guidances.en" if (s.get("guidances") or {}).get("en") else "guidance",
                "guidance_preview": guidance[:200],
            }, indent=2))
            break

    # Show first section with content
    print("\n-- Sample content section --")
    for s in sections:
        content = (s.get("specification") or {}).get("content", "")
        if content and len(content) > 50:
            print(json.dumps({
                "id": s["id"][:16],
                "type": s.get("type"),
                "title": get_title(s),
                "content_preview": content[:300],
                "has_attachables": bool(s.get("attachables")),
            }, indent=2))
            break

    # Show section with visibility
    print("\n-- Sample section with visibility --")
    for s in sections:
        if (s.get("visibility") or {}).get("conditions"):
            print(json.dumps({
                "id": s["id"][:16],
                "type": s.get("type"),
                "title": get_title(s),
                "visibility": s["visibility"],
            }, indent=2))
            break

    print(f"\nFull data written to {tmp_dir}/")


# ── CLI ──────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="Extract letters from CaseWare Cloud")
    parser.add_argument("--url", required=True, help="CaseWare engagement or document URL")
    parser.add_argument("--discover", action="store_true",
                        help="Dump raw API data for inspection")
    parser.add_argument("--json-only", action="store_true",
                        help="Output JSON only (skip Word generation)")
    args = parser.parse_args()

    host, tenant, engagement_id, document_id = parse_url(args.url)

    if args.discover:
        discover(engagement_id, document_id, host, tenant)
        return

    if not document_id:
        raise ValueError("URL must include a document fragment (#/letter/<id>)")

    if args.json_only:
        data = extract_letter_data(engagement_id, document_id, host, tenant)
        out_path = PROJECT_ROOT / ".tmp" / "letter_data.json"
        out_path.parent.mkdir(exist_ok=True)
        with open(out_path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        print(f"JSON written to {out_path}")
        return

    docx_bytes = generate_report_bytes(engagement_id, document_id, host, tenant)
    out_path = PROJECT_ROOT / ".tmp" / "letter_output.docx"
    out_path.write_bytes(docx_bytes)
    print(f"Word document written to {out_path}")


if __name__ == "__main__":
    main()
