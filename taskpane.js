/**
 * Send to Power Automate — Outlook Add-in
 * Replace the URL below with your Power Automate HTTP trigger URL.
 */

const POWER_AUTOMATE_URL = "https://defaultfac31a4a31d74bfa8f95dc3432fb68.48.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/f329136f04bb45b1907ccd79f40285b7/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=QNYfLxIX3JJmo_UasCXnUB07f6RN677x7fkr8vKopdQ";
const MAX_NOTE_CHARS = 1000;

// ── DOM refs ─────────────────────────────────────────────────────
const sendBtn       = document.getElementById("sendBtn");
const statusBox     = document.getElementById("statusBox");
const statusMsg     = document.getElementById("statusMsg");
const metaFrom      = document.getElementById("metaFrom");
const metaSubject   = document.getElementById("metaSubject");
const metaDate      = document.getElementById("metaDate");
const attachList    = document.getElementById("attachList");
const attachCount   = document.getElementById("attachCount");
const selectAllRow  = document.getElementById("selectAllRow");
const selectAllChk  = document.getElementById("selectAll");
const notesArea     = document.getElementById("notesArea");
const charCountEl   = document.getElementById("charCount");

// Stores the full attachment list from Office.js
let allAttachments = [];

// ── Office init ──────────────────────────────────────────────────
Office.onReady((info) => {
  if (info.host !== Office.HostType.Outlook) return;
  populatePreview();
  sendBtn.addEventListener("click", sendToFlow);

  // Character counter for notes
  notesArea.addEventListener("input", () => {
    const len = notesArea.value.length;
    charCountEl.textContent = len;
    charCountEl.parentElement.classList.toggle("over", len >= MAX_NOTE_CHARS);
  });

  // Select-all checkbox
  selectAllChk.addEventListener("change", () => {
    document.querySelectorAll(".att-checkbox").forEach(cb => {
      cb.checked = selectAllChk.checked;
    });
    updateSendLabel();
  });
});

// ── Populate email preview + attachment list ─────────────────────
function populatePreview() {
  const item = Office.context.mailbox.item;

  metaFrom.textContent    = item.from
    ? `${item.from.displayName} <${item.from.emailAddress}>` : "(unknown)";
  metaSubject.textContent = item.subject || "(no subject)";
  metaDate.textContent    = item.dateTimeCreated
    ? new Date(item.dateTimeCreated).toLocaleString() : "(unknown)";

  // Filter to file attachments only
  allAttachments = (item.attachments || []).filter(
    a => a.attachmentType === Office.MailboxEnums.AttachmentType.File
  );

  attachList.innerHTML = "";

  if (allAttachments.length === 0) {
    attachList.innerHTML = '<li style="border:none;padding:4px 0"><span class="no-attach">No file attachments</span></li>';
    attachCount.textContent = "";
  } else {
    selectAllRow.style.display = "flex";
    attachCount.textContent = `${allAttachments.length} file${allAttachments.length > 1 ? "s" : ""}`;

    allAttachments.forEach((att, i) => {
      const li = document.createElement("li");
      const cbId = `att-${i}`;
      li.innerHTML = `
        <input type="checkbox" class="att-checkbox" id="${cbId}" data-index="${i}" checked/>
        <label for="${cbId}">
          <svg width="14" height="14" viewBox="0 0 14 14" fill="none">
            <path d="M3 1h6l3 3v9H3V1z" stroke="#0078d4" stroke-width="1.2" fill="none"/>
            <path d="M9 1v3h3" stroke="#0078d4" stroke-width="1.2"/>
          </svg>
          <span class="fname" title="${escHtml(att.name)}">${escHtml(att.name)}</span>
          <span class="fsize">${formatBytes(att.size)}</span>
        </label>
      `;
      // Clicking the row toggles the checkbox
      li.addEventListener("click", (e) => {
        if (e.target.tagName !== "INPUT") {
          const cb = li.querySelector("input");
          cb.checked = !cb.checked;
        }
        updateSelectAllState();
        updateSendLabel();
      });
      li.querySelector("input").addEventListener("change", () => {
        updateSelectAllState();
        updateSendLabel();
      });
      attachList.appendChild(li);
    });
  }

  sendBtn.disabled = false;
  updateSendLabel();
}

// Keep "select all" checkbox in sync with individual checkboxes
function updateSelectAllState() {
  const all  = document.querySelectorAll(".att-checkbox");
  const checked = document.querySelectorAll(".att-checkbox:checked");
  selectAllChk.indeterminate = checked.length > 0 && checked.length < all.length;
  selectAllChk.checked = checked.length === all.length;
}

// Update button label to show how many attachments will be sent
function updateSendLabel() {
  const checked = document.querySelectorAll(".att-checkbox:checked").length;
  const total   = allAttachments.length;
  if (total === 0) {
    sendBtn.querySelector("span") && (sendBtn.lastChild.textContent = " Send to Flow");
    sendBtn.childNodes[1] && (sendBtn.childNodes[1].textContent = " Send to Flow");
    // Simple approach: just set innerHTML keeping the icon
    sendBtn.innerHTML = `
      <svg width="16" height="16" viewBox="0 0 16 16" fill="none">
        <path d="M1 1l14 7-14 7V9.5l10-1.5-10-1.5V1z" fill="currentColor"/>
      </svg>
      Send to Flow
    `;
  } else {
    sendBtn.innerHTML = `
      <svg width="16" height="16" viewBox="0 0 16 16" fill="none">
        <path d="M1 1l14 7-14 7V9.5l10-1.5-10-1.5V1z" fill="currentColor"/>
      </svg>
      Send to Flow ${checked > 0 ? `(${checked} attachment${checked > 1 ? "s" : ""})` : "(no attachments)"}
    `;
  }
}

// ── Main send ────────────────────────────────────────────────────
async function sendToFlow() {
  sendBtn.disabled = true;
  setStatus("info", '<span class="dots">Collecting email data</span>');

  const item = Office.context.mailbox.item;

  try {
    // 1. Metadata
    const metadata = buildMetadata(item);

    // 2. Notes from the textarea
    const notes = notesArea.value.trim();

    // 3. HTML body
    setStatus("info", '<span class="dots">Reading email body</span>');
    const htmlBody = await getBodyAsync();

    // 4. Only checked attachments
    const checkedBoxes = [...document.querySelectorAll(".att-checkbox:checked")];
    const selectedAttachments = checkedBoxes.map(cb => allAttachments[+cb.dataset.index]);

    const attachments = [];
    for (let i = 0; i < selectedAttachments.length; i++) {
      const att = selectedAttachments[i];
      setStatus("info", `<span class="dots">Reading attachment ${i + 1}/${selectedAttachments.length}: ${escHtml(att.name)}</span>`);
      let content = null;
      let error   = null;
      try {
        content = await getAttachmentContentAsync(att.id);
      } catch (e) {
        error = e.message;
      }
      attachments.push({
        name:        att.name,
        contentType: att.contentType,
        sizeBytes:   att.size,
        content,
        error,
      });
    }

    // 5. Build payload — includes notes field
    const payload = { metadata, notes, htmlBody, attachments };

    setStatus("info", '<span class="dots">Sending to Power Automate</span>');

    const response = await fetch(POWER_AUTOMATE_URL, {
      method:  "POST",
      headers: { "Content-Type": "application/json" },
      body:    JSON.stringify(payload),
    });

    if (!response.ok) {
      const body = await response.text().catch(() => "");
      throw new Error(`HTTP ${response.status}${body ? ": " + body.slice(0, 200) : ""}`);
    }

    setStatus("ok", `Sent successfully! (${attachments.length} attachment${attachments.length !== 1 ? "s" : ""} included)`);
    sendBtn.disabled = false;

  } catch (err) {
    console.error("sendToFlow error:", err);
    setStatus("fail", `Failed: ${escHtml(err.message)}`);
    sendBtn.disabled = false;
  }
}

// ── Helpers ──────────────────────────────────────────────────────

function buildMetadata(item) {
  return {
    subject:           item.subject || "",
    from:              { displayName: item.from?.displayName || "", emailAddress: item.from?.emailAddress || "" },
    to:                (item.to  || []).map(r => ({ displayName: r.displayName, emailAddress: r.emailAddress })),
    cc:                (item.cc  || []).map(r => ({ displayName: r.displayName, emailAddress: r.emailAddress })),
    receivedDateTime:  item.dateTimeCreated ? new Date(item.dateTimeCreated).toISOString() : null,
    internetMessageId: item.internetMessageId || "",
    conversationId:    item.conversationId    || "",
  };
}

function getBodyAsync() {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.body.getAsync(Office.CoercionType.Html, (result) => {
      result.status === Office.AsyncResultStatus.Succeeded
        ? resolve(result.value)
        : reject(new Error(result.error?.message || "Failed to read email body"));
    });
  });
}

function getAttachmentContentAsync(id) {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.getAttachmentContentAsync(id, (result) => {
      result.status === Office.AsyncResultStatus.Succeeded
        ? resolve(result.value.content)
        : reject(new Error(result.error?.message || "Failed to read attachment"));
    });
  });
}

function setStatus(type, html) {
  statusBox.className = "show " + { info: "info", ok: "ok", fail: "fail" }[type];
  statusMsg.innerHTML = html;
}

function formatBytes(b) {
  if (!b)          return "";
  if (b < 1024)    return b + " B";
  if (b < 1048576) return (b / 1024).toFixed(1) + " KB";
  return (b / 1048576).toFixed(1) + " MB";
}

function escHtml(s) {
  return String(s)
    .replace(/&/g,"&amp;").replace(/</g,"&lt;")
    .replace(/>/g,"&gt;").replace(/"/g,"&quot;");
}
