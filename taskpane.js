/**
 * Send to Power Automate — Outlook Add-in
 * Password is entered by the user and used as the API key sent to Cloudflare.
 */

const PROXY_URL = "https://outlook-addin-proxy.baokhanh3041975.workers.dev/send";

// ── DOM refs ─────────────────────────────────────────────────────
const sendBtn      = document.getElementById("sendBtn");
const statusBox    = document.getElementById("statusBox");
const statusMsg    = document.getElementById("statusMsg");
const metaFrom     = document.getElementById("metaFrom");
const metaSubject  = document.getElementById("metaSubject");
const metaDate     = document.getElementById("metaDate");
const fileList     = document.getElementById("fileList");
const fileControls = document.getElementById("fileControls");
const imgList      = document.getElementById("imgList");
const imgCount     = document.getElementById("imgCount");
const imageSection = document.getElementById("imageSection");
const imgToggle    = document.getElementById("imgToggle");
const imgBody      = document.getElementById("imgBody");
const commentBox   = document.getElementById("commentBox");
const paFunction   = document.getElementById("paFunction");
const pwOverlay    = document.getElementById("pwOverlay");
const pwInput      = document.getElementById("pwInput");
const pwError      = document.getElementById("pwError");
const pwSubmit     = document.getElementById("pwSubmit");
const pwCancel     = document.getElementById("pwCancel");

// ── Office init ──────────────────────────────────────────────────
Office.onReady((info) => {
  if (info.host !== Office.HostType.Outlook) return;
  populatePreview();

  sendBtn.addEventListener("click", () => openPasswordModal());

  // Image dropdown toggle
  imgToggle.addEventListener("click", () => {
    imgToggle.classList.toggle("open");
    imgBody.classList.toggle("open");
  });

  // Select / deselect all — files
  document.getElementById("selectAllFiles").addEventListener("click", () => setAllChecked("file-cb", true));
  document.getElementById("deselectAllFiles").addEventListener("click", () => setAllChecked("file-cb", false));

  // Select / deselect all — images
  document.getElementById("selectAllImgs").addEventListener("click", () => setAllChecked("img-cb", true));
  document.getElementById("deselectAllImgs").addEventListener("click", () => setAllChecked("img-cb", false));

  // Password modal
  pwSubmit.addEventListener("click", handlePasswordSubmit);
  pwCancel.addEventListener("click", closePasswordModal);
  pwInput.addEventListener("keydown", (e) => { if (e.key === "Enter") handlePasswordSubmit(); });
});

// ── Populate preview ─────────────────────────────────────────────
function populatePreview() {
  const item = Office.context.mailbox.item;

  metaFrom.textContent    = item.from ? `${item.from.displayName} <${item.from.emailAddress}>` : "(unknown)";
  metaSubject.textContent = item.subject || "(no subject)";
  metaDate.textContent    = item.dateTimeCreated ? new Date(item.dateTimeCreated).toLocaleString() : "(unknown)";

  const allAtts  = (item.attachments || []).filter(a => a.attachmentType === Office.MailboxEnums.AttachmentType.File);
  const imageExt = /\.(jpg|jpeg|png|gif|webp|bmp|svg|tiff|ico)$/i;
  const isImage  = (a) => (a.contentType && a.contentType.startsWith("image/")) || imageExt.test(a.name);

  const files  = allAtts.filter(a => !isImage(a));
  const images = allAtts.filter(a =>  isImage(a));

  // ── File attachments ──────────────────────────────────────────
  fileList.innerHTML = "";
  if (files.length === 0) {
    fileList.innerHTML = '<li class="no-attach">No file attachments</li>';
  } else {
    fileControls.style.display = "flex";
    files.forEach((att, i) => {
      const li = buildAttachmentItem(att, `file-cb`, i);
      fileList.appendChild(li);
    });
  }

  // ── Image attachments (collapsible) ──────────────────────────
  if (images.length === 0) {
    imageSection.style.display = "none";
  } else {
    imageSection.style.display = "block";
    imgCount.textContent = images.length;
    imgList.innerHTML = "";
    images.forEach((att, i) => {
      const li = buildAttachmentItem(att, `img-cb`, i);
      imgList.appendChild(li);
    });
  }

  sendBtn.disabled = false;
}

function buildAttachmentItem(att, cbClass, i) {
  const li = document.createElement("li");
  const id = `${cbClass}-${i}`;
  li.innerHTML = `
    <label for="${id}">
      <input type="checkbox" id="${id}" class="${cbClass}" checked data-id="${escHtml(att.id)}" data-name="${escHtml(att.name)}" data-type="${escHtml(att.contentType || '')}">
      <svg width="14" height="14" viewBox="0 0 14 14" fill="none">
        <path d="M3 1h6l3 3v9H3V1z" stroke="#0078d4" stroke-width="1.2" fill="none"/>
        <path d="M9 1v3h3" stroke="#0078d4" stroke-width="1.2"/>
      </svg>
      <span class="fname">${escHtml(att.name)}</span>
    </label>
    <span class="fsize">${formatBytes(att.size)}</span>
  `;
  return li;
}

function setAllChecked(cbClass, checked) {
  document.querySelectorAll(`.${cbClass}`).forEach(cb => cb.checked = checked);
}

// ── Password modal ───────────────────────────────────────────────
function openPasswordModal() {
  pwInput.value = "";
  pwError.textContent = "";
  pwOverlay.classList.add("show");
  setTimeout(() => pwInput.focus(), 100);
}

function closePasswordModal() {
  pwOverlay.classList.remove("show");
}

async function handlePasswordSubmit() {
  const password = pwInput.value.trim();
  if (!password) {
    pwError.textContent = "Please enter a password.";
    return;
  }
  pwError.textContent = "";
  pwSubmit.disabled = true;
  pwSubmit.textContent = "Sending…";

  closePasswordModal();
  await sendToFlow(password);

  pwSubmit.disabled = false;
  pwSubmit.textContent = "Confirm & Send";
}

// ── Main send ────────────────────────────────────────────────────
async function sendToFlow(password) {
  sendBtn.disabled = true;
  setStatus("info", '<span class="dots">Collecting email data</span>');

  const item = Office.context.mailbox.item;

  try {
    const metadata = buildMetadata(item);

    setStatus("info", '<span class="dots">Reading email body</span>');
    const htmlBody = await getBodyAsync();

    // Collect selected attachments only
    const selectedIds = new Set(
      [...document.querySelectorAll(".file-cb:checked, .img-cb:checked")]
        .map(cb => cb.dataset.id)
    );

    const allAtts = (item.attachments || []).filter(
      a => a.attachmentType === Office.MailboxEnums.AttachmentType.File && selectedIds.has(a.id)
    );

    const attachments = [];
    for (let i = 0; i < allAtts.length; i++) {
      const att = allAtts[i];
      setStatus("info", `<span class="dots">Reading attachment ${i + 1}/${allAtts.length}: ${escHtml(att.name)}</span>`);
      let content = null, errorMsg = null;
      try { content = await getAttachmentContentAsync(att.id); }
      catch (e) { errorMsg = e.message; }
      attachments.push({ name: att.name, contentType: att.contentType, sizeBytes: att.size, content, error: errorMsg });
    }

    const payload = {
      metadata,
      htmlBody,
      attachments,
      comment:    commentBox.value.trim(),
      paFunction: paFunction.value,
    };

    setStatus("info", '<span class="dots">Sending to Power Automate</span>');

    const response = await fetch(PROXY_URL, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "x-api-key": password,
      },
      body: JSON.stringify(payload),
    });

    const data = await response.json().catch(() => ({}));

    if (response.status === 401) {
      setStatus("fail", "❌ Incorrect password. Request was not sent to Power Automate.");
      sendBtn.disabled = false;
      return;
    }

    if (!response.ok) {
      throw new Error(data.error || `HTTP ${response.status}`);
    }

    const timeStr = data.receivedAt
      ? new Date(data.receivedAt).toLocaleString()
      : new Date().toLocaleString();

    setStatus("ok", `✅ Power Automate received your request at <strong>${timeStr}</strong>`);
    sendBtn.disabled = false;

  } catch (err) {
    console.error("sendToFlow error:", err);
    setStatus("fail", `❌ ${escHtml(err.message)}`);
    sendBtn.disabled = false;
  }
}

// ── Helpers ──────────────────────────────────────────────────────
function buildMetadata(item) {
  return {
    subject:           item.subject || "",
    from:              { displayName: item.from?.displayName || "", emailAddress: item.from?.emailAddress || "" },
    to:                (item.to || []).map(r => ({ displayName: r.displayName, emailAddress: r.emailAddress })),
    cc:                (item.cc || []).map(r => ({ displayName: r.displayName, emailAddress: r.emailAddress })),
    receivedDateTime:  item.dateTimeCreated ? new Date(item.dateTimeCreated).toISOString() : null,
    internetMessageId: item.internetMessageId || "",
    conversationId:    item.conversationId || "",
  };
}

function getBodyAsync() {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.body.getAsync(Office.CoercionType.Html, { asyncContext: null }, (result) => {
      result.status === Office.AsyncResultStatus.Succeeded
        ? resolve(result.value)
        : reject(new Error(result.error?.message || "Failed to read email body"));
    });
  });
}

function getAttachmentContentAsync(attachmentId) {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.getAttachmentContentAsync(attachmentId, { asyncContext: null }, (result) => {
      result.status === Office.AsyncResultStatus.Succeeded
        ? resolve(result.value.content)
        : reject(new Error(result.error?.message || "Failed to read attachment"));
    });
  });
}

function setStatus(type, html) {
  statusBox.className = "show " + type;
  statusMsg.innerHTML = html;
}

function formatBytes(bytes) {
  if (!bytes) return "";
  if (bytes < 1024)    return bytes + " B";
  if (bytes < 1048576) return (bytes / 1024).toFixed(1) + " KB";
  return (bytes / 1048576).toFixed(1) + " MB";
}

function escHtml(str) {
  return String(str)
    .replace(/&/g, "&amp;").replace(/</g, "&lt;")
    .replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}
