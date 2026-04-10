/**
 * Send to Power Automate — Outlook Add-in
 * ─────────────────────────────────────────────────────────────────
 * Replace POWER_AUTOMATE_URL below with your HTTP-trigger URL from
 * Power Automate (When an HTTP request is received → HTTP POST URL).
 * ─────────────────────────────────────────────────────────────────
 */

const POWER_AUTOMATE_URL = "https://defaultfac31a4a31d74bfa8f95dc3432fb68.48.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/f329136f04bb45b1907ccd79f40285b7/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=QNYfLxIX3JJmo_UasCXnUB07f6RN677x7fkr8vKopdQ";

// ── DOM refs ────────────────────────────────────────────────────
const sendBtn = document.getElementById("sendBtn");
const statusBox = document.getElementById("statusBox");
const statusIcon = document.getElementById("statusIcon");
const statusMsg = document.getElementById("statusMsg");
const metaFrom = document.getElementById("metaFrom");
const metaSubject = document.getElementById("metaSubject");
const metaDate = document.getElementById("metaDate");
const attachList = document.getElementById("attachList");

// ── Office initialisation ───────────────────────────────────────
Office.onReady((info) => {
  if (info.host !== Office.HostType.Outlook) return;
  populatePreview();
  sendBtn.addEventListener("click", sendToFlow);
});

// ── Populate the email preview card ────────────────────────────
function populatePreview() {
  const item = Office.context.mailbox.item;

  // From
  const from = item.from;
  metaFrom.textContent = from
    ? `${from.displayName} <${from.emailAddress}>`
    : "(unknown)";

  // Subject
  metaSubject.textContent = item.subject || "(no subject)";

  // Date
  const d = item.dateTimeCreated;
  metaDate.textContent = d
    ? new Date(d).toLocaleString()
    : "(unknown)";

  // Attachments
  const atts = (item.attachments || []).filter(
    (a) => a.attachmentType === Office.MailboxEnums.AttachmentType.File
  );

  attachList.innerHTML = "";

  if (atts.length === 0) {
    attachList.innerHTML = '<li class="no-attach">No file attachments</li>';
  } else {
    atts.forEach((att) => {
      const li = document.createElement("li");
      li.innerHTML = `
        <svg width="14" height="14" viewBox="0 0 14 14" fill="none">
          <path d="M3 1h6l3 3v9H3V1z" stroke="#0078d4" stroke-width="1.2" fill="none"/>
          <path d="M9 1v3h3" stroke="#0078d4" stroke-width="1.2"/>
        </svg>
        <span class="fname">${escHtml(att.name)}</span>
        <span class="fsize">${formatBytes(att.size)}</span>
      `;
      attachList.appendChild(li);
    });
  }

  // Enable button now that preview is loaded
  sendBtn.disabled = false;
}

// ── Main send function ──────────────────────────────────────────
async function sendToFlow() {
  sendBtn.disabled = true;
  setStatus("info", '<span class="dots">Collecting email data</span>');

  const item = Office.context.mailbox.item;

  try {
    // 1. Metadata
    const metadata = buildMetadata(item);

    // 2. HTML body
    setStatus("info", '<span class="dots">Reading email body</span>');
    const htmlBody = await getBodyAsync();

    // 3. Attachments as base64
    const fileAttachments = (item.attachments || []).filter(
      (a) => a.attachmentType === Office.MailboxEnums.AttachmentType.File
    );

    const attachments = [];
    for (let i = 0; i < fileAttachments.length; i++) {
      const att = fileAttachments[i];
      setStatus(
        "info",
        `<span class="dots">Reading attachment ${i + 1}/${fileAttachments.length}: ${escHtml(att.name)}</span>`
      );
      let content = null;
      let errorMsg = null;
      try {
        content = await getAttachmentContentAsync(att.id);
      } catch (e) {
        errorMsg = e.message;
      }
      attachments.push({
        name: att.name,
        contentType: att.contentType,
        sizeBytes: att.size,
        content,          // base64 string, or null on error
        error: errorMsg,
      });
    }

    // 4. Build final payload
    const payload = {
      metadata,
      htmlBody,
      attachments,
    };

    // 5. POST to Power Automate
    setStatus("info", '<span class="dots">Sending to Power Automate</span>');

    const response = await fetch(POWER_AUTOMATE_URL, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });

    if (!response.ok) {
      const body = await response.text().catch(() => "");
      throw new Error(`HTTP ${response.status}${body ? ": " + body.slice(0, 200) : ""}`);
    }

    setStatus("ok", "✅ Email successfully sent to your Power Automate flow!");
    sendBtn.disabled = false;

  } catch (err) {
    console.error("sendToFlow error:", err);
    setStatus("fail", `❌ ${escHtml(err.message)}`);
    sendBtn.disabled = false;
  }
}

// ── Helpers ─────────────────────────────────────────────────────

function buildMetadata(item) {
  return {
    subject: item.subject || "",
    from: {
      displayName: item.from?.displayName || "",
      emailAddress: item.from?.emailAddress || "",
    },
    to: (item.to || []).map(mapRecipient),
    cc: (item.cc || []).map(mapRecipient),
    receivedDateTime: item.dateTimeCreated
      ? new Date(item.dateTimeCreated).toISOString()
      : null,
    internetMessageId: item.internetMessageId || "",
    conversationId: item.conversationId || "",
  };
}

function mapRecipient(r) {
  return { displayName: r.displayName, emailAddress: r.emailAddress };
}

function getBodyAsync() {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.body.getAsync(
      Office.CoercionType.Html,
      { asyncContext: null },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve(result.value);
        } else {
          reject(new Error(result.error?.message || "Failed to read email body"));
        }
      }
    );
  });
}

function getAttachmentContentAsync(attachmentId) {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.getAttachmentContentAsync(
      attachmentId,
      { asyncContext: null },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          // result.value.content is base64 for file attachments
          resolve(result.value.content);
        } else {
          reject(new Error(result.error?.message || "Failed to read attachment"));
        }
      }
    );
  });
}

function setStatus(type, html) {
  const icons = { info: "🔄", ok: "✅", fail: "❌" };
  statusBox.className = "show " + { info: "info", ok: "ok", fail: "fail" }[type];
  statusIcon.textContent = "";
  statusMsg.innerHTML = html;
}

function formatBytes(bytes) {
  if (!bytes) return "";
  if (bytes < 1024) return bytes + " B";
  if (bytes < 1048576) return (bytes / 1024).toFixed(1) + " KB";
  return (bytes / 1048576).toFixed(1) + " MB";
}

function escHtml(str) {
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}
