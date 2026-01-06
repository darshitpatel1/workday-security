function normalize(s) {
  return String(s || "").replace(/\s+/g, " ").trim();
}

function normalizeHeader(s) {
  return normalize(s)
    .toLowerCase()
    .replace(/\s+/g, " ")
    .replace(/[:()\-]/g, "");
}

function opToSections(operationText) {
  const op = normalize(operationText).toLowerCase();

  const sections = new Set();

  // examples: "View Only", "View and Modify", "Get and Put", "Get Only", "Put Only"
  if (op.includes("view")) sections.add("VIEW");
  if (op.includes("modify")) sections.add("MODIFY");
  if (op.includes("get")) sections.add("GET");
  if (op.includes("put")) sections.add("PUT");

  return [...sections];
}

function isKnownOperationValue(v) {
  const s = normalize(v).toLowerCase();
  if (!s) return false;

  const known = new Set([
    "view only",
    "modify only",
    "view and modify",
    "get only",
    "put only",
    "get and put",
  ]);
  if (known.has(s)) return true;

  if (s.includes("view") && s.includes("modify")) return true;
  if (s.includes("get") && s.includes("put")) return true;
  if (s.endsWith(" only") && (s.includes("view") || s.includes("modify") || s.includes("get") || s.includes("put"))) {
    return true;
  }

  return false;
}

function collectKeys(rows, { maxRows = 25 } = {}) {
  const set = new Set();
  for (let i = 0; i < Math.min(rows.length, maxRows); i++) {
    const r = rows[i];
    if (!r || typeof r !== "object") continue;
    for (const k of Object.keys(r)) set.add(k);
  }
  return [...set];
}

function columnStats(rows, key, { maxRows = 200 } = {}) {
  let nonEmpty = 0;
  let opMatches = 0;

  for (let i = 0; i < Math.min(rows.length, maxRows); i++) {
    const r = rows[i];
    const val = r?.[key];
    const s = normalize(val);
    if (!s) continue;
    nonEmpty += 1;
    if (isKnownOperationValue(s)) opMatches += 1;
  }

  return {
    key,
    nonEmpty,
    opMatches,
    opRatio: nonEmpty ? opMatches / nonEmpty : 0,
  };
}

function extractOperationAndPolicy(row, { opKey, policyKey } = {}) {
  const r = row && typeof row === "object" ? row : {};

  const entries = Object.entries(r)
    .map(([k, v]) => ({ k, raw: v, s: normalize(v) }))
    .filter((x) => x.s.length > 0);

  const byKey = (key) => {
    if (!key) return null;
    const hit = entries.find((e) => e.k === key);
    return hit || null;
  };

  let opEntry = byKey(opKey) || entries.find((e) => isKnownOperationValue(e.s)) || null;

  // Policy should not look like an operation; pick "most descriptive" candidate.
  const policyCandidates = entries.filter((e) => e !== opEntry && !isKnownOperationValue(e.s));
  let policyEntry =
    byKey(policyKey) ||
    (policyCandidates.length
      ? policyCandidates.reduce((best, cur) => (cur.s.length > best.s.length ? cur : best))
      : null);

  // If we picked a "policy" that is actually an operation, fall back to a non-operation column.
  if (policyEntry && isKnownOperationValue(policyEntry.s) && policyCandidates.length) {
    policyEntry = policyCandidates.reduce((best, cur) => (cur.s.length > best.s.length ? cur : best));
  }

  // If opKey/policyKey were swapped (common when headers shift), try to recover.
  if (opEntry && policyEntry && !isKnownOperationValue(opEntry.s) && isKnownOperationValue(policyEntry.s)) {
    const tmp = opEntry;
    opEntry = policyEntry;
    policyEntry = tmp;
  }

  return {
    operation: opEntry?.raw ?? "",
    policy: policyEntry?.raw ?? "",
  };
}

function pickColumnKey(keys, { exact = [], includes = [] } = {}) {
  const normalized = keys.map((k) => ({ raw: k, n: normalizeHeader(k) }));

  for (const e of exact) {
    const ne = normalizeHeader(e);
    const hit = normalized.find((x) => x.n === ne);
    if (hit) return hit.raw;
  }

  for (const inc of includes) {
    const ninc = normalizeHeader(inc);
    const hit = normalized.find((x) => x.n.includes(ninc));
    if (hit) return hit.raw;
  }

  return null;
}

function pickOperationKeyByContent(rows, keys) {
  const stats = keys.map((k) => columnStats(rows, k));
  stats.sort((a, b) => {
    // primary: more matches, secondary: higher ratio, tertiary: more non-empty
    if (b.opMatches !== a.opMatches) return b.opMatches - a.opMatches;
    if (b.opRatio !== a.opRatio) return b.opRatio - a.opRatio;
    return b.nonEmpty - a.nonEmpty;
  });
  return stats[0]?.key || null;
}

function pickPolicyKeyByContent(rows, keys, opKey) {
  const candidates = keys.filter((k) => k && k !== opKey);
  const stats = candidates.map((k) => columnStats(rows, k));

  // Prefer columns with lots of non-empty values and very low operation-likeness.
  stats.sort((a, b) => {
    const aBad = a.opRatio;
    const bBad = b.opRatio;
    if (aBad !== bBad) return aBad - bBad;
    return b.nonEmpty - a.nonEmpty;
  });

  return stats[0]?.key || null;
}

function buildBlockFromRows(rows) {
  // rows = [{ Operation, Domain Security Policy }, ...]
  const buckets = { MODIFY: [], VIEW: [], PUT: [], GET: [] };

  const safeRows = Array.isArray(rows) ? rows : [];
  const keys = collectKeys(safeRows);

  let opKey =
    pickColumnKey(keys, {
      exact: ["Operation"],
      includes: ["operation", "access", "permission"],
    }) || null;

  let policyKey =
    pickColumnKey(keys, {
      exact: ["Domain Security Policy", "Domain Security Policies"],
      includes: ["domain security policy", "domain security policies", "security policy", "domain policy"],
    }) || null;

  // Content-based fallback (handles renamed/unknown headers).
  if (!opKey) opKey = pickOperationKeyByContent(safeRows, keys);
  if (!policyKey) policyKey = pickPolicyKeyByContent(safeRows, keys, opKey);

  // Validate header-based picks: the policy column should NOT be mostly operation values.
  if (opKey) {
    const opStats = columnStats(safeRows, opKey);
    if (opStats.opRatio < 0.5) {
      opKey = pickOperationKeyByContent(safeRows, keys);
    }
  }

  if (policyKey) {
    const policyStats = columnStats(safeRows, policyKey);
    if (policyStats.opRatio > 0.3) {
      policyKey = pickPolicyKeyByContent(safeRows, keys, opKey);
    }
  }

  for (const r of safeRows) {
    const { operation, policy } = extractOperationAndPolicy(r, { opKey, policyKey });
    const policyText = normalize(policy);
    if (!policyText) continue;
    if (isKnownOperationValue(policyText)) continue;

    const targets = opToSections(operation);
    for (const t of targets) {
      buckets[t].push(policyText);
    }
  }

  // de-dupe while keeping order
  for (const k of Object.keys(buckets)) {
    const seen = new Set();
    buckets[k] = buckets[k].filter((x) => {
      const key = x.toLowerCase();
      if (seen.has(key)) return false;
      seen.add(key);
      return true;
    });
  }

  // format
  const out = [];
  for (const k of ["MODIFY", "VIEW", "PUT", "GET"]) {
    if (buckets[k].length) {
      out.push(`${k}:`);
      out.push(...buckets[k]);
      out.push(""); // blank line
    }
  }

  return out.join("\n").trim() + "\n";
}

window.addEventListener("message", (event) => {
  const msg = event.data;
  if (!msg) return;

  if (msg.type === "XLSX_SANDBOX_PING") {
    event.source?.postMessage(
      { type: "XLSX_SANDBOX_PONG", requestId: msg.requestId, ok: true },
      "*"
    );
    return;
  }

  if (msg.type !== "PARSE_XLSX") return;

  const { requestId, arrayBuffer } = msg || {};

  try {
    if (typeof XLSX === "undefined") throw new Error("XLSX not available in sandbox");

    const wb = XLSX.read(arrayBuffer, { type: "array" });

    // Use first sheet by default
    const sheetName = wb.SheetNames[0];
    const ws = wb.Sheets[sheetName];

    // Convert to JSON rows
    const json = XLSX.utils.sheet_to_json(ws, { defval: "" });

    // Expect columns:
    // Operation | Domain Security Policy | Domain Security Policies | Functional Areas ...
    // We only need first two.
    const text = buildBlockFromRows(json);

    event.source?.postMessage(
      { type: "PARSE_XLSX_RESULT", requestId, ok: true, text },
      "*"
    );
  } catch (e) {
    event.source?.postMessage(
      { type: "PARSE_XLSX_RESULT", requestId, ok: false, error: e?.message || String(e) },
      "*"
    );
  }
});
