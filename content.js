// Workday Bulk Domain Permissions (Manifest V3 content script)

let SHOULD_STOP = false;
let RUN_TOKEN = 0; // cancels previous run if user presses Run again

chrome.runtime.onMessage.addListener((msg) => {
  if (msg?.type === "WD_BULK_PERM_STOP") {
    SHOULD_STOP = true;
    console.log("[WD_BULK_PERM] Stop requested.");
  }

  if (msg?.type === "WD_BULK_PERM_RUN") {
    SHOULD_STOP = false;
    RUN_TOKEN += 1;
    const token = RUN_TOKEN;
    console.log(`[WD_BULK_PERM] Run requested. token=${token}`);

    runBulk(msg.payload, token).catch((err) => {
      if (token === RUN_TOKEN && !SHOULD_STOP) {
        console.error("[WD_BULK_PERM] Fatal:", err);
      }
    });
  }
});

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

async function waitFor(fn, { timeout = 15000, interval = 150 } = {}) {
  const start = Date.now();
  while (Date.now() - start < timeout) {
    const v = fn();
    if (v) return v;
    await sleep(interval);
  }
  return null;
}

function normalize(s) {
  return (s || "").replace(/\s+/g, " ").trim().toLowerCase();
}

function isOnCorrectPage() {
  const h1 = document.querySelector("h1");
  const t = h1?.textContent?.trim() || "";
  return t.toLowerCase().includes("maintain domain permissions for security group");
}

function setNativeValue(el, value) {
  const proto = Object.getPrototypeOf(el);
  const desc = Object.getOwnPropertyDescriptor(proto, "value");
  if (desc?.set) desc.set.call(el, value);
  else el.value = value;

  el.dispatchEvent(new Event("input", { bubbles: true }));
  el.dispatchEvent(new Event("change", { bubbles: true }));
}

function dispatchKey(el, key) {
  const codes = { Enter: 13, ArrowDown: 40, Escape: 27, Tab: 9 };
  const keyCode = codes[key] || 0;

  const evt = {
    key,
    code: key,
    keyCode,
    which: keyCode,
    bubbles: true,
    cancelable: true,
  };

  el.dispatchEvent(new KeyboardEvent("keydown", evt));
  el.dispatchEvent(new KeyboardEvent("keypress", evt));
  el.dispatchEvent(new KeyboardEvent("keyup", evt));
}

function getFieldByLabelText(labelText) {
  const labels = [...document.querySelectorAll('label[data-automation-id="formLabel"]')];
  const label = labels.find((l) => (l.textContent || "").trim() === labelText);
  if (!label) return null;

  const li = label.closest("li");
  if (!li) return null;

  const widget = li.querySelector('[data-automation-id="responsiveMonikerInput"]');
  const input = li.querySelector('input[data-uxi-widget-type="selectinput"][id$="-input"]');

  return { li, widget, input };
}

function getSelectedPillsText(containerEl) {
  const pills = [
    ...containerEl.querySelectorAll(
      'ul[data-automation-id="selectedItemList"] [data-automation-id="promptOption"],' +
        '[data-automation-id="selectedItemList"] [data-automation-id="promptOption"]'
    ),
  ];
  return pills.map((p) => normalize(p.textContent));
}

function isAlreadySelectedInLi(li, value) {
  const target = normalize(value);
  return getSelectedPillsText(li).some((t) => t === target);
}

function isSelectedInPopup(popupEl, value) {
  if (!popupEl) return false;
  const target = normalize(value);
  return getSelectedPillsText(popupEl).some((t) => t === target);
}

// IMPORTANT: if widgetId is missing, fallback to "any active popup"
function getActivePopupForWidgetId(widgetId) {
  if (!widgetId) {
    return document.querySelector('[data-automation-activepopup="true"]');
  }

  const anchor = document.querySelector(
    `[data-automation-activepopup="true"] [data-associated-widget="${CSS.escape(widgetId)}"]`
  );
  if (!anchor) {
    return document.querySelector('[data-automation-activepopup="true"]');
  }

  return anchor.closest('[data-automation-activepopup="true"]') || anchor;
}

async function openPrompt(widget, input) {
  // Focus/click input first (more stable)
  if (input) {
    input.click();
    input.focus({ preventScroll: true });
    await sleep(80);
  }

  // Click ONLY real prompt/search opener (never a random button)
  const icon =
    widget.querySelector('[data-automation-id="promptIcon"]') ||
    widget.querySelector('[data-automation-id="promptSearchButton"]') ||
    widget.querySelector('button[aria-label*="Prompt"], button[aria-label*="prompt"]') ||
    widget.querySelector('button[aria-label*="Search"], button[aria-label*="search"]');

  if (icon) {
    icon.click();
    await sleep(250);
    return;
  }

  // Fallback: open dropdown via keyboard
  if (input) {
    dispatchKey(input, "ArrowDown");
    await sleep(250);
  }
}

function rowLabelText(row) {
  const opt =
    row.querySelector('[data-automation-id="promptOption"]') ||
    row.querySelector('[data-automation-id="menuItemLabel"]') ||
    row;
  return normalize(opt.innerText || opt.textContent || "");
}

function isNavMenuRowText(t) {
  if (!t) return true;
  return (
    t === "all" ||
    t.startsWith("partial list") ||
    t.startsWith("create security policy for domain")
  );
}

function getAllPopupRows(popupEl) {
  if (!popupEl) return [];
  return [
    ...popupEl.querySelectorAll('[data-automation-id="menuItem"]'),
    ...popupEl.querySelectorAll('[role="option"]'),
  ];
}

function getSearchResultRows(popupEl) {
  if (!popupEl) return [];

  const rows = getAllPopupRows(popupEl);

  const checkboxRows = rows.filter((r) =>
    r.querySelector('input[type="checkbox"], [role="checkbox"], [data-automation-id*="checkbox"]')
  );

  const cleanCheckboxRows = checkboxRows.filter((r) => !isNavMenuRowText(rowLabelText(r)));
  if (cleanCheckboxRows.length) return cleanCheckboxRows;

  const clean = rows.filter((r) => !isNavMenuRowText(rowLabelText(r)));
  return clean;
}

function findMenuItemByExactText(popupEl, exactText) {
  const target = normalize(exactText);
  const rows = getAllPopupRows(popupEl);
  return rows.find((r) => rowLabelText(r) === target) || null;
}

function findBestRow(rows, value) {
  const target = normalize(value);

  const exact = rows.find((r) => rowLabelText(r) === target);
  if (exact) return exact;

  const prefix = rows.find((r) => rowLabelText(r).startsWith(target));
  if (prefix) return prefix;

  if (rows.length === 1) return rows[0];

  const contains = rows.find((r) => rowLabelText(r).includes(target));
  if (contains) return contains;

  return null;
}

/** NEW: detect if a result row is already checked (prevents toggle-off) */
function isRowChecked(row) {
  const inputCb = row.querySelector('input[type="checkbox"]');
  if (inputCb) return !!inputCb.checked;

  const ariaCb = row.querySelector('[role="checkbox"]');
  if (ariaCb) {
    const v = ariaCb.getAttribute("aria-checked");
    if (v === "true") return true;
    if (v === "false") return false;
  }

  // some Workday rows put aria-checked on the row itself
  const v2 = row.getAttribute?.("aria-checked");
  if (v2 === "true") return true;

  return false;
}

/** NEW: treat selection as true if the matching row is checked */
function isValueCheckedInPopup(popupEl, value) {
  if (!popupEl) return false;
  const rows = getSearchResultRows(popupEl);
  const row = findBestRow(rows, value);
  if (!row) return false;
  return isRowChecked(row);
}

function clickCheckboxInRow(row) {
  const cb =
    row.querySelector('input[type="checkbox"]') ||
    row.querySelector('[role="checkbox"]') ||
    row.querySelector('[data-automation-id*="checkbox"]');

  if (cb) {
    const clickable = cb.closest("label") || cb.closest('[role="checkbox"]') || cb;
    clickable.click();
    return true;
  }

  row.click();
  return false;
}

// PLUS: popup-only settle wait (hadPopup=true)
async function closePopup(input, widgetId, hadPopup = false) {
  // close by clicking outside (no ESC)
  document.body.click();
  await sleep(120);

  let still = getActivePopupForWidgetId(widgetId);
  if (still) {
    document.body.click();
    await sleep(120);
    still = getActivePopupForWidgetId(widgetId);
  }

  // popup-only: ensure popup is fully gone + tiny settle delay
  if (hadPopup) {
    await waitFor(() => !getActivePopupForWidgetId(widgetId), {
      timeout: 6000,
      interval: 120,
    });
    await sleep(180);
  }

  // re-focus the same search box (Workday sometimes needs focus to accept next Enter)
  const focusEl =
    input && document.contains(input) ? input : input?.id ? document.getElementById(input.id) : null;

  if (focusEl) {
    focusEl.focus({ preventScroll: true });
    await sleep(60);
  }
}

/**
 * KEY FIX:
 * After typing, force "search" mode by pressing Enter.
 * If still stuck in menu mode (All / Partial List), click All then press Enter again.
 */
async function forceSearchMode({ popup, widgetId, labelText, value }) {
  let field = getFieldByLabelText(labelText);
  if (!field?.input) return { popup, input: null };

  const input = field.input;

  input.focus();
  await sleep(60);
  dispatchKey(input, "Enter");
  await sleep(250);

  const rowsNow = getAllPopupRows(popup);
  const nonEmpty = rowsNow.length > 0;
  const allNav = nonEmpty && rowsNow.every((r) => isNavMenuRowText(rowLabelText(r)));

  if (allNav) {
    const allRow = findMenuItemByExactText(popup, "All");
    const partialRow = rowsNow.find((r) => rowLabelText(r).startsWith("partial list"));

    if (allRow) {
      allRow.click();
      await sleep(300);
    } else if (partialRow) {
      partialRow.click();
      await sleep(300);
    }

    popup =
      (await waitFor(() => getActivePopupForWidgetId(widgetId), {
        timeout: 6000,
        interval: 120,
      })) || popup;

    field = getFieldByLabelText(labelText);
    if (!field?.input) return { popup, input: null };

    field.input.focus();
    await sleep(60);

    setNativeValue(field.input, value);
    await sleep(120);

    dispatchKey(field.input, "Enter");
    await sleep(300);
    return { popup, input: field.input };
  }

  return { popup, input };
}

async function addValueToField(labelText, value, { delayMs, skipExisting, stopOnError, token }) {
  const field = await waitFor(() => getFieldByLabelText(labelText), { timeout: 15000 });
  if (!field?.input || !field?.widget) {
    throw new Error(`Field not found for label: "${labelText}"`);
  }

  const { li, widget, input } = field;
  const widgetId = widget?.id || input?.id || "";

  // if already selected in main list, do nothing
  if (skipExisting && isAlreadySelectedInLi(li, value)) {
    console.log(`[WD_BULK_PERM] Skip (already selected): ${value}`);
    return { skipped: true };
  }

  widget.scrollIntoView({ block: "center" });
  await sleep(80);

  await openPrompt(widget, input);

  // popup open (with fallback forcing open if it didn't appear)
  let popup = await waitFor(() => getActivePopupForWidgetId(widgetId), {
    timeout: 5000,
    interval: 120,
  });

  if (!popup) {
    const ref = getFieldByLabelText(labelText);
    const inp = ref?.input || input;

    if (inp) {
      inp.click();
      await sleep(80);
      dispatchKey(inp, "ArrowDown");
      await sleep(250);
    }

    popup = await waitFor(() => getActivePopupForWidgetId(widgetId), {
      timeout: 5000,
      interval: 120,
    });
  }

  if (!popup) throw new Error(`Popup did not open for "${labelText}"`);

  // Type value
  const freshField = getFieldByLabelText(labelText);
  const activeInput = freshField?.input || input;

  activeInput.focus();
  await sleep(60);
  setNativeValue(activeInput, "");
  await sleep(60);
  setNativeValue(activeInput, value);
  await sleep(140);

  const forced = await forceSearchMode({
    popup,
    widgetId,
    labelText,
    value,
  });

  popup = forced.popup;
  const usedInput = forced.input || activeInput;

  // IMPORTANT CHANGE:
  // Auto-selected can be either:
  //  - pill is present (first 5 visible)
  //  - OR the matching result row checkbox is already checked (when pills collapse under MORE)
  const autoSelected = await waitFor(
    () =>
      isAlreadySelectedInLi(li, value) ||
      isSelectedInPopup(popup, value) ||
      isValueCheckedInPopup(popup, value),
    { timeout: 2500, interval: 120 }
  );

  if (autoSelected) {
    console.log(`[WD_BULK_PERM] Auto-selected by Workday (pill/checked): ${value}`);

    // For auto-select, don't do popup-settle wait; just close normally
    await closePopup(usedInput, widgetId, false);

    // Verification must also accept "checked row" (because pills may be hidden under MORE)
    const ok = await waitFor(
      () =>
        isAlreadySelectedInLi(li, value) ||
        (getActivePopupForWidgetId(widgetId) &&
          (isSelectedInPopup(getActivePopupForWidgetId(widgetId), value) ||
            isValueCheckedInPopup(getActivePopupForWidgetId(widgetId), value))),
      { timeout: 6000, interval: 150 }
    );

    if (!ok) {
      throw new Error(`Selection not verified (pill not found) for "${value}" in "${labelText}"`);
    }

    await sleep(delayMs);
    return { skipped: false, auto: true };
  }

  // Not auto-selected => must pick from Search Results
  const rows = await waitFor(() => {
    popup = getActivePopupForWidgetId(widgetId) || popup;
    const r = getSearchResultRows(popup);
    return r.length ? r : null;
  }, { timeout: 15000, interval: 150 });

  if (!rows) {
    throw new Error(`Search results did not load for "${value}" in "${labelText}"`);
  }

  const row = findBestRow(rows, value);
  if (!row) {
    const sample = rows.slice(0, 6).map((r) => (r.innerText || r.textContent || "").trim());
    console.warn("[WD_BULK_PERM] Could not match. Sample rows:", sample);
    throw new Error(`No matching option found for "${value}" in "${labelText}"`);
  }

  // CRITICAL FIX:
  // if it's already checked, DO NOT click (clicking toggles OFF)
  if (isRowChecked(row)) {
    console.log(`[WD_BULK_PERM] Already checked in results (no click): ${value}`);
    await closePopup(usedInput, widgetId, true);

    const ok = await waitFor(
      () =>
        isAlreadySelectedInLi(li, value) ||
        (getActivePopupForWidgetId(widgetId) &&
          (isSelectedInPopup(getActivePopupForWidgetId(widgetId), value) ||
            isValueCheckedInPopup(getActivePopupForWidgetId(widgetId), value))),
      { timeout: 8000, interval: 150 }
    );

    if (!ok) {
      throw new Error(`Selection not verified (pill not found) for "${value}" in "${labelText}"`);
    }

    await sleep(delayMs);
    return { skipped: false, auto: false, alreadyChecked: true };
  }

  row.scrollIntoView({ block: "center" });
  await sleep(80);

  clickCheckboxInRow(row);

  await waitFor(
    () =>
      isRowChecked(row) ||
      isSelectedInPopup(popup, value) ||
      isAlreadySelectedInLi(li, value) ||
      isValueCheckedInPopup(popup, value),
    { timeout: 4000, interval: 120 }
  );

  await closePopup(usedInput, widgetId, true);

  const ok = await waitFor(
    () =>
      isAlreadySelectedInLi(li, value) ||
      (getActivePopupForWidgetId(widgetId) &&
        (isSelectedInPopup(getActivePopupForWidgetId(widgetId), value) ||
          isValueCheckedInPopup(getActivePopupForWidgetId(widgetId), value))),
    { timeout: 12000, interval: 150 }
  );

  if (!ok) {
    throw new Error(`Selection not verified (pill not found) for "${value}" in "${labelText}"`);
  }

  await sleep(delayMs);
  return { skipped: false, auto: false };
}

async function runSection(title, labelText, values, opts) {
  if (!values?.length) return { added: 0, skipped: 0 };

  console.log(`[WD_BULK_PERM] Running ${title} (${values.length})…`);
  let added = 0;
  let skipped = 0;

  for (let i = 0; i < values.length; i++) {
    if (opts.token !== RUN_TOKEN) {
      console.warn(`[WD_BULK_PERM] Cancelled (new run started). token=${opts.token}`);
      return { added, skipped, cancelled: true };
    }
    if (SHOULD_STOP) throw new Error("Stopped by user.");

    const v = values[i];
    console.log(`[WD_BULK_PERM] ${title} (${i + 1}/${values.length}): ${v}`);

    try {
      const r = await addValueToField(labelText, v, opts);
      if (r.skipped) skipped++;
      else added++;
    } catch (e) {
      console.error(`[WD_BULK_PERM] ERROR in ${title} for "${v}":`, e);
      if (opts.stopOnError) throw e;
    }

    await sleep(opts.delayMs);
  }

  return { added, skipped };
}

async function runBulk(payload, token) {
  const { sections, delayMs = 250, skipExisting = true, stopOnError = false } = payload || {};

  if (!isOnCorrectPage()) {
    alert(
      "Workday Bulk Permissions: Please open 'Maintain Domain Permissions for Security Group' task page first."
    );
    return;
  }

  const LABELS = {
    MODIFY: "Domain Security Policies permitting Modify access",
    VIEW: "Domain Security Policies permitting View access",
    PUT: "Domain Security Policies permitting Put access",
    GET: "Domain Security Policies permitting Get access",
  };

  const opts = { delayMs, skipExisting, stopOnError, token };

  console.log("[WD_BULK_PERM] Start ✅", {
    token,
    counts: {
      MODIFY: sections?.MODIFY?.length || 0,
      VIEW: sections?.VIEW?.length || 0,
      PUT: sections?.PUT?.length || 0,
      GET: sections?.GET?.length || 0,
    },
    delayMs,
    skipExisting,
    stopOnError,
  });

  const summary = { MODIFY: null, VIEW: null, PUT: null, GET: null };

  summary.MODIFY = await runSection("MODIFY", LABELS.MODIFY, sections?.MODIFY || [], opts);
  summary.VIEW = await runSection("VIEW", LABELS.VIEW, sections?.VIEW || [], opts);
  summary.PUT = await runSection("PUT", LABELS.PUT, sections?.PUT || [], opts);
  summary.GET = await runSection("GET", LABELS.GET, sections?.GET || [], opts);

  if (token === RUN_TOKEN && !SHOULD_STOP) {
    console.log("[WD_BULK_PERM] Done ✅ Summary:", summary);
    alert(`Workday Bulk Permissions done.\n\n${JSON.stringify(summary, null, 2)}`);
  }
}
