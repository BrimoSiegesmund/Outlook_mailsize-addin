/* Mailgröße Anzeige - Live-Berechnung der aktuellen Nachrichtengröße
   inkl. Anhänge während des Verfassens in Outlook (Compose Mode). */

const WARN_THRESHOLD_MB = 20; // Standard-Limit vieler M365-Tenants, ggf. anpassen
const BASE64_OVERHEAD = 1.37; // Anhänge werden beim Versand ca. 37% größer (Base64-Kodierung)
const HEADER_OVERHEAD_BYTES = 2048; // grobe Pauschale für Mail-Header/Metadaten

let refreshTimer = null;

Office.onReady(() => {
  document.getElementById("refreshBtn").addEventListener("click", updateSize);
  updateSize();
  // Automatische Aktualisierung alle 3 Sekunden
  refreshTimer = setInterval(updateSize, 3000);

  // Zusätzlich sofort reagieren, wenn sich Anhänge ändern
  if (Office.context.mailbox.item.addHandlerAsync) {
    Office.context.mailbox.item.addHandlerAsync(
      Office.EventType.AttachmentsChanged,
      updateSize
    );
  }
});

function formatBytes(bytes) {
  if (bytes < 1024) return bytes + " B";
  const kb = bytes / 1024;
  if (kb < 1024) return kb.toFixed(1) + " KB";
  const mb = kb / 1024;
  return mb.toFixed(2) + " MB";
}

function getBodySizeBytes() {
  return new Promise((resolve) => {
    Office.context.mailbox.item.body.getAsync(
      Office.CoercionType.Html,
      (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          const text = result.value || "";
          // UTF-8 Byte-Länge des HTML-Inhalts
          const bytes = new TextEncoder().encode(text).length;
          resolve(bytes);
        } else {
          resolve(0);
        }
      }
    );
  });
}

function getAttachments() {
  return new Promise((resolve) => {
    const item = Office.context.mailbox.item;
    if (!item.attachments) {
      resolve([]);
      return;
    }
    // Im Compose-Modus liefert item.attachments direkt ein Array
    resolve(item.attachments || []);
  });
}

async function updateSize() {
  try {
    const [bodyBytes, attachments] = await Promise.all([
      getBodySizeBytes(),
      getAttachments(),
    ]);

    let attachBytes = 0;
    const listEl = document.getElementById("attachList");
    listEl.innerHTML = "";

    attachments.forEach((att) => {
      const size = att.size || 0;
      attachBytes += size;
      const row = document.createElement("div");
      row.innerHTML = `<span>${escapeHtml(att.name || "Anhang")}</span><span>${formatBytes(size)}</span>`;
      listEl.appendChild(row);
    });

    const totalRaw = bodyBytes + attachBytes + HEADER_OVERHEAD_BYTES;
    const estimatedSend = bodyBytes + HEADER_OVERHEAD_BYTES + attachBytes * BASE64_OVERHEAD;

    document.getElementById("bodySize").textContent = formatBytes(bodyBytes);
    document.getElementById("attachSize").textContent = formatBytes(attachBytes);
    document.getElementById("totalSize").textContent = formatBytes(totalRaw);
    document.getElementById("estimateSize").textContent = formatBytes(estimatedSend);

    const warnEl = document.getElementById("warning");
    const thresholdBytes = WARN_THRESHOLD_MB * 1024 * 1024;
    if (estimatedSend > thresholdBytes) {
      warnEl.textContent = `Achtung: Die geschätzte Sendegröße überschreitet ${WARN_THRESHOLD_MB} MB. Der Versand könnte fehlschlagen, abhängig vom Limit deines Postfachs.`;
      warnEl.classList.remove("hidden");
    } else {
      warnEl.classList.add("hidden");
    }
  } catch (err) {
    // still fail silently in the UI but log for debugging
    console.error("Fehler bei der Größenberechnung:", err);
  }
}

function escapeHtml(str) {
  const div = document.createElement("div");
  div.textContent = str;
  return div.innerHTML;
}
