const state = {
    snapshot: null,
    draft: null,
    dirty: false,
    flashTimer: null,
    pollTimer: null
};

const ui = {};

document.addEventListener("DOMContentLoaded", init);

async function init() {
    wireUi();
    await refreshState(true);
    state.pollTimer = window.setInterval(() => refreshState(false), 500);
}

function wireUi() {
    ui.saveButton = document.getElementById("saveButton");
    ui.refreshVmixButton = document.getElementById("refreshVmixButton");
    ui.addRuleButton = document.getElementById("addRuleButton");
    ui.flash = document.getElementById("flash");
    ui.rulesContainer = document.getElementById("rulesContainer");
    ui.settingsPanel = document.getElementById("settingsPanel");
    ui.portList = document.getElementById("portList");

    ui.liveFrame = document.getElementById("liveFrame");
    ui.liveRule = document.getElementById("liveRule");
    ui.liveDecimals = document.getElementById("liveDecimals");
    ui.liveTarget = document.getElementById("liveTarget");
    ui.liveUpdated = document.getElementById("liveUpdated");
    ui.actionBadge = document.getElementById("actionBadge");
    ui.serialSummary = document.getElementById("serialSummary");
    ui.vmixSummary = document.getElementById("vmixSummary");
    ui.listenerUrl = document.getElementById("listenerUrl");
    ui.rawSerial = document.getElementById("rawSerial");
    ui.logs = document.getElementById("logs");

    ui.saveButton.addEventListener("click", saveSettings);
    ui.refreshVmixButton.addEventListener("click", refreshVmix);
    ui.addRuleButton.addEventListener("click", addRule);
    ui.settingsPanel.addEventListener("input", markDirty);
    ui.settingsPanel.addEventListener("change", syncDraftFromDom);
    ui.rulesContainer.addEventListener("input", markDirty);
    ui.rulesContainer.addEventListener("change", syncDraftFromDom);
    ui.rulesContainer.addEventListener("click", handleRuleButtons);
}

async function refreshState(resetDraft) {
    try {
        const snapshot = await getJson("/api/state");
        applySnapshot(snapshot, resetDraft);
    } catch (error) {
        showFlash(error.message, "error", 5000);
    }
}

function applySnapshot(snapshot, resetDraft) {
    state.snapshot = snapshot;

    if (!state.draft || resetDraft || !state.dirty) {
        state.draft = deepClone(snapshot.settings);
        state.dirty = false;
        renderEditor();
    } else {
        renderPortList();
        updateRuleRuntimeBadges();
    }

    updateSaveButton();
    renderLive();
    renderRawSerial();
    renderLogs();
}

function renderEditor() {
    if (!state.draft || !state.snapshot) {
        return;
    }

    document.getElementById("serialEnabled").checked = !!state.draft.serial.enabled;
    document.getElementById("serialPort").value = state.draft.serial.port || "";
    document.getElementById("serialBaud").value = state.draft.serial.baud ?? 1200;
    document.getElementById("minSendIntervalMs").value = state.draft.serial.minSendIntervalMs ?? 0;
    document.getElementById("vmixBaseUrl").value = state.draft.vmix.baseUrl || "";
    document.getElementById("openBrowser").checked = !!state.draft.openBrowser;

    renderPortList();
    renderRules();
}

function renderPortList() {
    const options = (state.snapshot?.ports || []).map((port) => `<option value="${escapeHtml(port)}"></option>`).join("");
    ui.portList.innerHTML = options;
}

function renderRules() {
    const rules = state.draft?.rules || [];
    if (!rules.length) {
        ui.rulesContainer.innerHTML = `<div class="rule-card"><strong>No rules yet.</strong><div class="rule-hint">Add one and set the decimal-place match for each timer format.</div></div>`;
        return;
    }

    ui.rulesContainer.innerHTML = rules.map((rule, index) => ruleMarkup(rule, index)).join("");
    updateRuleRuntimeBadges();
}

function ruleMarkup(rule, index) {
    const selectedInput = getResolvedInputObject(rule);
    const fields = selectedInput?.fields || [];
    const targetHint = getRuleTargetHint(rule, selectedInput);
    const runtime = getRuleRuntime(rule.id);
    const fieldListId = `field-list-${rule.id}`;
    const useRunningQuote = !!rule.useRunningQuote;

    return `
        <article class="rule-card" data-index="${index}" data-rule-id="${escapeHtml(rule.id)}">
            <div class="rule-card-header">
                <div>
                    <h3>${escapeHtml(rule.name || `Rule ${index + 1}`)}</h3>
                    <div class="rule-meta">Decimal match ${escapeHtml(String(rule.decimalPlaces))}</div>
                </div>
                <span class="badge ${badgeClassForRule(runtime)}" data-role="rule-badge">${escapeHtml(ruleBadgeText(runtime))}</span>
            </div>

            <div class="rule-grid">
                ${fieldHtml("Enabled", `<input class="rule-enabled" type="checkbox" ${rule.enabled ? "checked" : ""}>`, true)}
                ${fieldHtml("Name", `<input class="rule-name" type="text" value="${escapeAttr(rule.name || "")}" placeholder="Thousandths timer">`)}
                ${fieldHtml("Decimal places", `<input class="rule-decimals" type="number" value="${escapeAttr(String(rule.decimalPlaces ?? 0))}" step="1">`)}
                ${fieldHtml("Target", buildInputSelect(rule.targetInput))}
                ${fieldHtml("Preferred Title", `<input class="rule-title" type="text" value="${escapeAttr(rule.preferredTitle || "")}" placeholder="Leave blank to use rule name">`)}
                ${fieldHtml("Field", `<input class="rule-field" list="${fieldListId}" type="text" value="${escapeAttr(rule.field || "Time.Text")}" placeholder="Time.Text"><datalist id="${fieldListId}">${fields.map((field) => `<option value="${escapeHtml(field.name || "")}"></option>`).join("")}</datalist>`)}
                ${fieldHtml("Overlay", buildOverlaySelect(rule.overlayNumber))}
                ${fieldHtml("Auto overlay", `<input class="rule-auto-overlay" type="checkbox" ${rule.autoOverlay ? "checked" : ""}>`, true)}
                ${fieldHtml("Quote means running", `<input class="rule-use-running-quote" type="checkbox" ${useRunningQuote ? "checked" : ""}>`, true)}
                ${useRunningQuote
                    ? fieldHtml("Flash when stopped", `<input class="rule-flash-when-stopped" type="checkbox" ${rule.flashWhenStopped ? "checked" : ""}>`, true)
                    : fieldHtml("Zero threshold", `<input class="rule-zero-threshold" type="number" step="0.01" value="${escapeAttr(String(rule.zeroThreshold ?? 0.05))}">`)}
                ${useRunningQuote
                    ? fieldHtml("Flash interval (ms)", `<input class="rule-flash-interval-ms" type="number" min="120" step="10" value="${escapeAttr(String(rule.flashIntervalMs ?? 350))}">`)
                    : fieldHtml("Rearm above", `<input class="rule-rearm-above" type="number" step="0.01" value="${escapeAttr(String(rule.rearmAbove ?? 2.0))}">`)}
                ${useRunningQuote
                    ? fieldHtml("Flash duration (ms)", `<input class="rule-flash-duration-ms" type="number" min="60" step="10" value="${escapeAttr(String(rule.flashDurationMs ?? 140))}">`)
                    : fieldHtml("Rearm below", `<input class="rule-rearm-below" type="number" step="0.01" value="${escapeAttr(String(rule.rearmBelow ?? 0.3))}">`)}
                ${useRunningQuote
                    ? `<div class="field field-note">Quote mode removes threshold rearm. Stopped flashing uses vMix input alpha: visible at 255, hidden at 0.</div>`
                    : ``}
            </div>

            <div class="rule-hint" data-role="target-hint">${escapeHtml(targetHint)}</div>
            <div class="runtime-line" data-role="runtime-line">
                <strong>${escapeHtml(runtimeLine(runtime))}</strong>
            </div>

            <div class="rule-actions">
                <button class="small-button" data-action="overlay">Overlay now</button>
                <button class="small-button" data-action="rearm">Rearm</button>
                <button class="small-button" data-action="move-up">Move up</button>
                <button class="small-button" data-action="move-down">Move down</button>
                <button class="small-button" data-action="delete">Delete</button>
            </div>
        </article>
    `;
}

function fieldHtml(label, control, checkbox = false) {
    return `<label class="field ${checkbox ? "checkbox-field" : ""}"><span>${escapeHtml(label)}</span>${control}</label>`;
}

function buildInputSelect(selected) {
    const inputs = state.snapshot?.vmix?.inputs || [];
    const options = ['<option value="">Smart match</option>']
        .concat(inputs.map((input) => {
            const value = input.key || input.number || "";
            const chosen = value === selected ? "selected" : "";
            return `<option value="${escapeHtml(value)}" ${chosen}>${escapeHtml(inputLabel(input))}</option>`;
        }))
        .join("");

    return `<select class="rule-target-input">${options}</select>`;
}

function buildOverlaySelect(selected) {
    return `<select class="rule-overlay-number">
        ${[1, 2, 3, 4].map((number) => `<option value="${number}" ${Number(selected) === number ? "selected" : ""}>Overlay ${number}</option>`).join("")}
    </select>`;
}

function renderLive() {
    const live = state.snapshot?.live || {};
    const serial = state.snapshot?.serialStatus || {};
    const vmix = state.snapshot?.vmix || {};
    const resolved = live.resolvedTarget?.label || live.resolvedTarget?.input || "--";
    const runState = live.isRunning === true ? "running" : live.isRunning === false ? "stopped" : "";

    ui.liveFrame.textContent = live.lastFrame || "--";
    ui.liveRule.textContent = live.matchedRuleName || "--";
    ui.liveDecimals.textContent = live.decimalPlaces ?? "--";
    ui.liveTarget.textContent = resolved;
    ui.liveUpdated.textContent = formatDateTime(live.updatedAt);
    ui.actionBadge.textContent = live.lastAction || "idle";
    ui.actionBadge.className = `badge ${badgeClass(live.lastAction)}`;
    ui.serialSummary.textContent = serial.message ? `${serial.message}${serial.port ? ` | ${serial.port}` : ""}${runState ? ` | ${runState}` : ""}` : "--";
    ui.vmixSummary.textContent = vmix.message || "--";
    ui.listenerUrl.textContent = state.snapshot?.listener?.url || "--";
}

function renderRawSerial() {
    const entries = state.snapshot?.rawSerial || [];
    if (!entries.length) {
        ui.rawSerial.innerHTML = `<div class="raw-line"><div class="raw-line-body">No serial lines captured yet.</div></div>`;
        return;
    }

    ui.rawSerial.innerHTML = entries.slice(0, 40).map((entry) => `
        <div class="raw-line">
            <div class="raw-line-head">
                <span>${escapeHtml(formatTime(entry.timestamp))}</span>
                <span>${escapeHtml(`${entry.length} bytes`)}</span>
            </div>
            <div class="raw-line-body"><span class="raw-line-label">Display</span>${escapeHtml(entry.displayText || "")}</div>
            <div class="raw-line-hex"><span class="raw-line-label">HEX</span>${escapeHtml(entry.hex || "")}</div>
        </div>
    `).join("");
}

function renderLogs() {
    const logs = state.snapshot?.logs || [];
    if (!logs.length) {
        ui.logs.innerHTML = `<div class="log-line"><span class="log-level info">idle</span><span>--</span><span>No events yet.</span></div>`;
        return;
    }

    ui.logs.innerHTML = logs.slice(0, 25).map((entry) => `
        <div class="log-line">
            <span class="log-level ${escapeHtml(entry.level || "info")}">${escapeHtml(entry.level || "info")}</span>
            <span>${escapeHtml(formatTime(entry.timestamp))}</span>
            <span>${escapeHtml(entry.message || "")}</span>
        </div>
    `).join("");
}

function updateRuleRuntimeBadges() {
    document.querySelectorAll(".rule-card").forEach((card) => {
        const ruleId = card.dataset.ruleId;
        const rule = (state.draft?.rules || []).find((item) => item.id === ruleId);
        const runtime = getRuleRuntime(ruleId);
        const resolved = rule ? getResolvedInputObject(rule) : null;

        const badge = card.querySelector('[data-role="rule-badge"]');
        const runtimeLineEl = card.querySelector('[data-role="runtime-line"] strong');
        const hint = card.querySelector('[data-role="target-hint"]');

        if (badge) {
            badge.textContent = ruleBadgeText(runtime);
            badge.className = `badge ${badgeClassForRule(runtime)}`;
        }

        if (runtimeLineEl) {
            runtimeLineEl.textContent = runtimeLine(runtime);
        }

        if (hint && rule) {
            hint.textContent = getRuleTargetHint(rule, resolved);
        }
    });
}

function getRuleRuntime(ruleId) {
    return (state.snapshot?.ruleStatus || []).find((item) => item.id === ruleId) || {};
}

function getResolvedInputObject(rule) {
    const inputs = state.snapshot?.vmix?.inputs || [];
    if (rule.targetInput) {
        return inputs.find((input) => input.key === rule.targetInput || input.number === rule.targetInput) || null;
    }

    const term = (rule.preferredTitle || rule.name || "").trim().toLowerCase();
    if (!term) {
        return null;
    }

    let best = null;
    let bestScore = -1;
    for (const input of inputs) {
        const candidates = [input.number, input.title, input.shortTitle].filter(Boolean).map((value) => value.toLowerCase());
        let score = -1;
        for (const candidate of candidates) {
            if (candidate === term) score = Math.max(score, 100);
            else if (candidate.startsWith(term)) score = Math.max(score, 85);
            else if (candidate.includes(term)) score = Math.max(score, 70);
        }
        if (score > bestScore) {
            best = input;
            bestScore = score;
        }
    }
    return best;
}

function getRuleTargetHint(rule, selectedInput) {
    const routingHint = rule.targetInput
        ? (selectedInput ? `Explicit target: ${inputLabel(selectedInput)}` : `Explicit target: ${rule.targetInput}`)
        : (selectedInput ? `Smart match: ${inputLabel(selectedInput)}` : "Smart match: no vMix title match yet");

    if (!rule.useRunningQuote) {
        return routingHint;
    }

    const stoppedHint = rule.flashWhenStopped ? "Stopped result flashes." : "Stopped result stays steady.";
    return `${routingHint} Quote at the end means running. ${stoppedHint}`;
}

function runtimeLine(runtime) {
    const stateText = runtime.isRunning === true
        ? "Running"
        : runtime.isRunning === false
            ? "Stopped"
            : "State unknown";
    const overlay = runtime.overlayShown ? "Overlay used for this run" : "Overlay ready";
    const lastFrame = runtime.lastFrame ? `Last frame ${runtime.lastFrame}` : "No frame yet";
    const lastSeen = runtime.lastMatchedAt ? `Seen ${formatTime(runtime.lastMatchedAt)}` : "Never matched";
    return `${stateText} | ${overlay} | ${lastFrame} | ${lastSeen}`;
}

function ruleBadgeText(runtime) {
    if (runtime.isRunning === true) return "running";
    if (runtime.isRunning === false) return "stopped";
    if (runtime.overlayShown) return "overlay on";
    if (runtime.lastFrame) return "ready";
    return "waiting";
}

function badgeClassForRule(runtime) {
    if (runtime.isRunning === true) return "ok";
    if (runtime.isRunning === false) return "warn";
    if (runtime.overlayShown) return "ok";
    if (runtime.lastFrame) return "warn";
    return "";
}

function badgeClass(action) {
    if (["sent", "overlay"].includes(action)) return "ok";
    if (["vmix-error", "invalid", "no-target"].includes(action)) return "error";
    if (["duplicate", "throttled", "no-rule"].includes(action)) return "warn";
    return "";
}

function markDirty() {
    state.dirty = true;
    updateSaveButton();
}

function syncDraftFromDom() {
    if (!state.snapshot) {
        return;
    }

    state.draft = readSettingsFromDom();
    state.dirty = true;
    renderRules();
    updateSaveButton();
}

function readSettingsFromDom() {
    const rules = Array.from(document.querySelectorAll(".rule-card")).map((card) => {
        const useRunningQuote = card.querySelector(".rule-use-running-quote").checked;
        return {
            id: card.dataset.ruleId,
            enabled: card.querySelector(".rule-enabled").checked,
            name: card.querySelector(".rule-name").value.trim(),
            decimalPlaces: toInt(card.querySelector(".rule-decimals").value, 0),
            targetInput: card.querySelector(".rule-target-input").value.trim(),
            preferredTitle: card.querySelector(".rule-title").value.trim(),
            field: card.querySelector(".rule-field").value.trim() || "Time.Text",
            overlayNumber: toInt(card.querySelector(".rule-overlay-number").value, 1),
            autoOverlay: card.querySelector(".rule-auto-overlay").checked,
            useRunningQuote,
            flashWhenStopped: card.querySelector(".rule-flash-when-stopped")?.checked ?? false,
            flashIntervalMs: toInt(card.querySelector(".rule-flash-interval-ms")?.value, 350),
            flashDurationMs: toInt(card.querySelector(".rule-flash-duration-ms")?.value, 140),
            zeroThreshold: toNumber(card.querySelector(".rule-zero-threshold")?.value, 0.05),
            rearmAbove: toNumber(card.querySelector(".rule-rearm-above")?.value, 2.0),
            rearmBelow: toNumber(card.querySelector(".rule-rearm-below")?.value, 0.3)
        };
    });

    return {
        uiPort: state.snapshot.settings.uiPort,
        openBrowser: document.getElementById("openBrowser").checked,
        serial: {
            enabled: document.getElementById("serialEnabled").checked,
            port: document.getElementById("serialPort").value.trim(),
            baud: toInt(document.getElementById("serialBaud").value, 1200),
            minSendIntervalMs: toInt(document.getElementById("minSendIntervalMs").value, 0)
        },
        vmix: {
            baseUrl: document.getElementById("vmixBaseUrl").value.trim()
        },
        rules
    };
}

async function saveSettings() {
    try {
        state.draft = readSettingsFromDom();
        const result = await postJson("/api/settings", state.draft);
        applySnapshot(result.state, true);
        showFlash(result.restartRequired ? "Settings saved. Restart the script to move the web UI to a new port." : "Settings saved.", "ok");
    } catch (error) {
        showFlash(error.message, "error", 5000);
    }
}

async function refreshVmix() {
    try {
        ui.refreshVmixButton.disabled = true;
        const result = await getJson("/api/vmix/inputs?refresh=1");
        state.snapshot.vmix.inputs = result.inputs || [];
        state.snapshot.vmix.reachable = result.status.reachable;
        state.snapshot.vmix.message = result.status.message;
        state.snapshot.vmix.lastError = result.status.lastError;
        state.snapshot.vmix.lastDiscoveryAt = result.status.lastDiscoveryAt;
        renderEditor();
        renderLive();
        showFlash("vMix inputs refreshed.", "ok");
    } catch (error) {
        showFlash(error.message, "error", 5000);
    } finally {
        ui.refreshVmixButton.disabled = false;
    }
}

function addRule() {
    state.draft = readSettingsFromDom();
    state.draft.rules.push({
        id: createId(),
        name: "New Timer",
        enabled: true,
        decimalPlaces: 2,
        targetInput: "",
        preferredTitle: "",
        field: "Time.Text",
        overlayNumber: 1,
        autoOverlay: true,
        useRunningQuote: true,
        flashWhenStopped: false,
        flashIntervalMs: 350,
        flashDurationMs: 140,
        zeroThreshold: 0.05,
        rearmAbove: 2.0,
        rearmBelow: 0.3
    });
    state.dirty = true;
    renderRules();
    updateSaveButton();
}

async function handleRuleButtons(event) {
    const button = event.target.closest("[data-action]");
    if (!button) {
        return;
    }

    const card = button.closest(".rule-card");
    const index = Number(card.dataset.index);
    const action = button.dataset.action;
    const draft = readSettingsFromDom();
    const rule = draft.rules[index];

    if (action === "delete") {
        draft.rules.splice(index, 1);
        state.draft = draft;
        state.dirty = true;
        renderRules();
        updateSaveButton();
        return;
    }

    if (action === "move-up" || action === "move-down") {
        const target = action === "move-up" ? index - 1 : index + 1;
        if (target < 0 || target >= draft.rules.length) {
            return;
        }
        const [moved] = draft.rules.splice(index, 1);
        draft.rules.splice(target, 0, moved);
        state.draft = draft;
        state.dirty = true;
        renderRules();
        updateSaveButton();
        return;
    }

    try {
        const result = await postJson("/api/rule-action", { ruleId: rule.id, action: action === "overlay" ? "overlay" : "rearm" });
        applySnapshot(result.state, false);
        showFlash(action === "overlay" ? "Manual overlay sent." : "Rule rearmed.", "ok");
    } catch (error) {
        showFlash(error.message, "error", 5000);
    }
}

function updateSaveButton() {
    ui.saveButton.textContent = state.dirty ? "Save settings*" : "Save settings";
}

function showFlash(message, type = "ok", timeout = 3200) {
    ui.flash.hidden = false;
    ui.flash.className = `flash ${type}`;
    ui.flash.textContent = message;

    window.clearTimeout(state.flashTimer);
    state.flashTimer = window.setTimeout(() => {
        ui.flash.hidden = true;
    }, timeout);
}

async function getJson(url) {
    const response = await fetch(url, { cache: "no-store" });
    const data = await response.json();
    if (!response.ok || data.ok === false) {
        throw new Error(data.error || `Request failed: ${response.status}`);
    }
    return data;
}

async function postJson(url, payload) {
    const response = await fetch(url, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload)
    });
    const data = await response.json();
    if (!response.ok || data.ok === false) {
        throw new Error(data.error || `Request failed: ${response.status}`);
    }
    return data;
}

function inputLabel(input) {
    const number = input.number || "";
    const title = input.title || input.shortTitle || "";
    const shortTitle = input.shortTitle && input.shortTitle !== title ? ` (${input.shortTitle})` : "";
    return title ? `#${number} ${title}${shortTitle}` : `#${number}`;
}

function formatDateTime(value) {
    if (!value) return "--";
    return new Date(value).toLocaleString();
}

function formatTime(value) {
    if (!value) return "--";
    return new Date(value).toLocaleTimeString();
}

function deepClone(value) {
    return JSON.parse(JSON.stringify(value));
}

function escapeHtml(value) {
    return String(value ?? "")
        .replaceAll("&", "&amp;")
        .replaceAll("<", "&lt;")
        .replaceAll(">", "&gt;")
        .replaceAll('"', "&quot;")
        .replaceAll("'", "&#39;");
}

function escapeAttr(value) {
    return escapeHtml(value);
}

function toInt(value, fallback) {
    const parsed = Number.parseInt(value, 10);
    return Number.isFinite(parsed) ? parsed : fallback;
}

function toNumber(value, fallback) {
    const parsed = Number.parseFloat(value);
    return Number.isFinite(parsed) ? parsed : fallback;
}

function createId() {
    if (window.crypto && typeof window.crypto.randomUUID === "function") {
        return window.crypto.randomUUID();
    }
    return `rule-${Date.now()}-${Math.random().toString(16).slice(2)}`;
}
