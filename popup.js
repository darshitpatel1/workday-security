async function ensureXLSXLoaded() {

  if (typeof XLSX !== "undefined") return XLSX;
  if (globalThis.XLSX) return globalThis.XLSX;
  if (window.XLSX) return window.XLSX;

  await new Promise((resolve, reject) => {
    const s = document.createElement("script");
    s.src = chrome.runtime.getURL("xlsx.full.min.js");
    s.onload = resolve;
    s.onerror = () => reject(new Error("Failed to load xlsx.full.min.js"));
    document.head.appendChild(s);
  });

  // Check again after load
  if (typeof XLSX !== "undefined") return XLSX;
  if (globalThis.XLSX) return globalThis.XLSX;
  if (window.XLSX) return window.XLSX;

  throw new Error("XLSX loaded but still not available (not on window/globalThis and no global XLSX binding)");
}


const $ = (id) => document.getElementById(id);

const bulkInput = $("bulkInput");
const runBtn = $("runBtn");
const stopBtn = $("stopBtn");
const statusBox = $("statusBox");

const advancedToggle = $("advancedToggle");
const advancedPanel = $("advancedPanel");

const targetSelect = $("targetSelect");
const delayMsInput = $("delayMs");
const skipExistingCb = $("skipExisting");
const stopOnErrorCb = $("stopOnError");

const xlsxFile = $("xlsxFile");
const clearBtn = $("clearBtn");

function setStatus(text) {
  statusBox.textContent = `Status: ${text}`;
}

function normalizeHeader(s) {
  return (s || "")
    .toString()
    .trim()
    .toLowerCase()
    .replace(/\s+/g, " ")
    .replace(/[:()\-]/g, "");
}

function uniqPush(mapSet, key, value) {
  const v = (value || "").toString().trim();
  if (!v) return;
  mapSet[key].add(v);
}

function opToSections(opRaw) {
  const op = (opRaw || "").toString().trim().toLowerCase();

  // common Workday export patterns
  if (op === "view only") return ["VIEW"];
  if (op === "modify only") return ["MODIFY"];
  if (op === "view and modify") return ["VIEW", "MODIFY"];

  if (op === "get only") return ["GET"];
  if (op === "put only") return ["PUT"];
  if (op === "get and put") return ["GET", "PUT"];

  // extra safety: handle variants
  if (op.includes("view") && op.includes("modify")) return ["VIEW", "MODIFY"];
  if (op.includes("get") && op.includes("put")) return ["GET", "PUT"];
  if (op.includes("view")) return ["VIEW"];
  if (op.includes("modify")) return ["MODIFY"];
  if (op.includes("get")) return ["GET"];
  if (op.includes("put")) return ["PUT"];

  return []; // unknown operation
}

function buildBlockText(sections) {
  const order = ["MODIFY", "VIEW", "PUT", "GET"];
  const lines = [];

  for (const k of order) {
    const arr = Array.from(sections[k] || []).sort((a, b) => a.localeCompare(b));
    if (!arr.length) continue;

    lines.push(`${k}:`);
    for (const v of arr) lines.push(v);
    lines.push(""); // blank line between blocks
  }

  return lines.join("\n").trim() + "\n";
}

function wait(ms) {
  return new Promise((r) => setTimeout(r, ms));
}

function getSandboxFrame() {
  return document.getElementById("xlsxSandbox");
}

let SANDBOX_READY_PROMISE = null;
async function ensureSandboxReady() {
  if (SANDBOX_READY_PROMISE) return SANDBOX_READY_PROMISE;

  SANDBOX_READY_PROMISE = (async () => {
    const frame = getSandboxFrame();
    if (!frame) throw new Error("Sandbox iframe not found (xlsxSandbox)");

    // Wait for a contentWindow to exist
    for (let i = 0; i < 50; i++) {
      if (frame.contentWindow) break;
      await wait(50);
    }
    if (!frame.contentWindow) throw new Error("Sandbox iframe not ready");

    // Handshake with retries in case the iframe isn't fully loaded yet.
    const requestId = `ping_${Date.now()}_${Math.random().toString(16).slice(2)}`;

    return await new Promise((resolve, reject) => {
      let done = false;
      let attempts = 0;

      const onMsg = (event) => {
        const msg = event.data;
        if (!msg || msg.type !== "XLSX_SANDBOX_PONG" || msg.requestId !== requestId) return;
        done = true;
        window.removeEventListener("message", onMsg);
        if (msg.xlsxOk === false) {
          reject(
            new Error(
              "XLSX not loaded in sandbox. Reload the extension and ensure `xlsx.full.min.js` is present (not empty)."
            )
          );
          return;
        }
        resolve(true);
      };

      window.addEventListener("message", onMsg);

      const tick = () => {
        if (done) return;
        attempts += 1;

        try {
          frame.contentWindow.postMessage({ type: "XLSX_SANDBOX_PING", requestId }, "*");
        } catch {
          // ignore and retry
        }

        if (attempts >= 60) {
          window.removeEventListener("message", onMsg);
          reject(new Error("Sandbox did not respond (XLSX_SANDBOX_PONG)."));
          return;
        }

        setTimeout(tick, 100);
      };

      tick();
    });
  })();

  return SANDBOX_READY_PROMISE;
}

async function parseExcelViaSandbox(file) {
  const frame = getSandboxFrame();
  if (!frame) throw new Error("Sandbox iframe not found (xlsxSandbox)");

  await ensureSandboxReady();

  const arrayBuffer = await file.arrayBuffer();
  const requestId = `${Date.now()}_${Math.random().toString(16).slice(2)}`;

  const result = await new Promise((resolve, reject) => {
    let timeout = null;
    const onMsg = (event) => {
      const msg = event.data;
      if (!msg || msg.type !== "PARSE_XLSX_RESULT" || msg.requestId !== requestId) return;

      if (timeout) clearTimeout(timeout);
      window.removeEventListener("message", onMsg);

      if (msg.ok) resolve(msg.text);
      else reject(new Error(msg.error || "Excel parse failed"));
    };

    window.addEventListener("message", onMsg);
    timeout = setTimeout(() => {
      window.removeEventListener("message", onMsg);
      reject(new Error("Timed out waiting for Excel parse result."));
    }, 20000);

    // Send buffer to sandbox (transferable)
    frame.contentWindow.postMessage(
      { type: "PARSE_XLSX", requestId, arrayBuffer },
      "*",
      [arrayBuffer]
    );
  });

  return result;
}


// Advanced toggle
advancedToggle.addEventListener("click", () => {
  advancedPanel.classList.toggle("hidden");
});

// Clear
clearBtn.addEventListener("click", () => {
  bulkInput.value = "";
  xlsxFile.value = "";
  setStatus("idle");
});

// Excel upload -> populate textarea
xlsxFile.addEventListener("change", async () => {
  try {
    const file = xlsxFile.files?.[0];
    if (!file) return;

    const txt = await parseExcelViaSandbox(file);
    if (!txt || !txt.trim()) {
      throw new Error(
        "No policies found in the file. Verify the sheet has columns like 'Operation' and 'Domain Security Policy'."
      );
    }
    bulkInput.value = txt;
    setStatus("excel imported ✓");
  } catch (e) {
    console.error(e);
    setStatus("excel import failed");
    alert(e?.message || String(e));
  }
});

// Parse textarea into sections (same logic you already use)
function parseInputToSections(text, defaultTarget) {
  const out = { MODIFY: [], VIEW: [], PUT: [], GET: [] };
  const lines = (text || "").split(/\r?\n/).map((l) => l.trim()).filter((l) => l.length > 0);

  let current = null;

  for (const line of lines) {
    const upper = line.toUpperCase();

    if (upper === "MODIFY:" || upper === "VIEW:" || upper === "PUT:" || upper === "GET:") {
      current = upper.replace(":", "");
      continue;
    }

    if (!current) {
      current = defaultTarget;
    }

    out[current].push(line);
  }

  // de-dupe while keeping order
  for (const k of Object.keys(out)) {
    const seen = new Set();
    out[k] = out[k].filter((v) => {
      const key = v.toLowerCase();
      if (seen.has(key)) return false;
      seen.add(key);
      return true;
    });
  }

  return out;
}

async function sendToActiveTab(message) {
  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
  if (!tab?.id) throw new Error("No active tab found.");
  await chrome.tabs.sendMessage(tab.id, message);
}

// Run
runBtn.addEventListener("click", async () => {
  try {
    const text = bulkInput.value || "";
    if (!text.trim()) {
      alert("Paste list or upload Excel first.");
      return;
    }

    const sections = parseInputToSections(text, targetSelect.value);

    const payload = {
      sections,
      delayMs: Number(delayMsInput.value || 250),
      skipExisting: !!skipExistingCb.checked,
      stopOnError: !!stopOnErrorCb.checked,
    };

    setStatus("sent to page…");
    await sendToActiveTab({ type: "WD_BULK_PERM_RUN", payload });
  } catch (e) {
    console.error(e);
    setStatus("error");
    alert(e?.message || String(e));
  }
});

// Stop
stopBtn.addEventListener("click", async () => {
  try {
    setStatus("stop requested…");
    await sendToActiveTab({ type: "WD_BULK_PERM_STOP" });
    setStatus("idle");
  } catch (e) {
    console.error(e);
    setStatus("error");
    alert(e?.message || String(e));
  }
});

// Initial status
setStatus("idle");
