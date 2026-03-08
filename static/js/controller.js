(function () {
  const form = document.getElementById("movement-form");
  if (!form) return;

  const feedback = document.getElementById("form-feedback");
  const lastRecord = document.getElementById("last-record");
  const scanToggle = document.getElementById("scan-toggle");
  const scannerWrap = document.getElementById("scanner-wrap");
  const eanInput = document.getElementById("ean_code");

  const csrfToken = document.querySelector('meta[name="csrf-token"]')?.content || "";

  let scanner = null;
  let scannerRunning = false;

  function setFeedback(message, type) {
    feedback.textContent = message;
    feedback.classList.remove("ok", "error");
    if (type) feedback.classList.add(type);
  }

  function withSiteLoader(promise, options = {}) {
    if (!window.SiteLoader?.trackPromise) return promise;
    return window.SiteLoader.trackPromise(promise, options);
  }

  async function startScanner() {
    if (scannerRunning) return;
    if (typeof Html5Qrcode !== "function") {
      setFeedback("Scanner library not loaded. Use manual EAN input.", "error");
      return;
    }

    scannerWrap.classList.remove("hidden");
    scanner = new Html5Qrcode("reader");
    try {
      await scanner.start(
        { facingMode: "environment" },
        {
          fps: 10,
          qrbox: { width: 260, height: 150 },
          rememberLastUsedCamera: true,
          formatsToSupport: [
            Html5QrcodeSupportedFormats.EAN_13,
            Html5QrcodeSupportedFormats.EAN_8,
            Html5QrcodeSupportedFormats.UPC_A,
            Html5QrcodeSupportedFormats.UPC_E,
            Html5QrcodeSupportedFormats.CODE_128
          ]
        },
        (decodedText) => {
          const cleaned = String(decodedText || "").replace(/\D+/g, "");
          if (cleaned.length >= 8 && cleaned.length <= 14) {
            eanInput.value = cleaned;
            setFeedback("EAN scanned successfully.", "ok");
            stopScanner();
          }
        },
        () => {}
      );
      scannerRunning = true;
      scanToggle.textContent = "Stop Scanner";
    } catch (_error) {
      setFeedback("Camera access failed. Check permissions or use manual entry.", "error");
      scannerWrap.classList.add("hidden");
    }
  }

  async function stopScanner() {
    if (!scanner || !scannerRunning) {
      scannerWrap.classList.add("hidden");
      scanToggle.textContent = "Scanner EAN";
      return;
    }
    try {
      await scanner.stop();
      await scanner.clear();
    } catch (_error) {
      // Ignore scanner stop errors.
    } finally {
      scannerRunning = false;
      scannerWrap.classList.add("hidden");
      scanToggle.textContent = "Scanner EAN";
    }
  }

  scanToggle.addEventListener("click", () => {
    if (scannerRunning) {
      stopScanner();
    } else {
      startScanner();
    }
  });

  form.addEventListener("submit", async (event) => {
    event.preventDefault();
    setFeedback("Saving...", "");

    const payload = {
      support_number: form.support_number.value.trim(),
      ean_code: form.ean_code.value.trim(),
      product_code: form.product_code.value.trim(),
      diff_plus: form.diff_plus.value,
      diff_minus: form.diff_minus.value
    };

    try {
      const response = await withSiteLoader(
        fetch("/api/movements", {
          method: "POST",
          headers: {
            "Content-Type": "application/json",
            "X-CSRF-Token": csrfToken
          },
          credentials: "same-origin",
          body: JSON.stringify(payload)
        }),
        {
          title: "Loading...",
          step: "Enregistrement du mouvement en cours..."
        }
      );

      const data = await response.json();
      if (!response.ok || !data.ok) {
        setFeedback(data.error || "Failed to save movement.", "error");
        return;
      }

      setFeedback("Movement saved successfully.", "ok");
      form.reset();
      form.diff_plus.value = "0";
      form.diff_minus.value = "0";

      lastRecord.classList.remove("empty");
      lastRecord.innerHTML = `
        <strong>ID #${data.movement_id}</strong><br>
        Date: ${new Date(data.movement_date).toLocaleString()}<br>
        N Support: ${payload.support_number}<br>
        EAN: ${payload.ean_code}<br>
        Code produit: ${payload.product_code}<br>
        Ecart +: ${payload.diff_plus} / Ecart -: ${payload.diff_minus}
      `;
    } catch (_error) {
      setFeedback("Network error. Please retry.", "error");
    }
  });

  window.addEventListener("beforeunload", () => {
    stopScanner();
  });
})();
