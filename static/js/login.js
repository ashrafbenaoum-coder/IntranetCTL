(function () {
  const form = document.getElementById("login-form");
  if (!form) return;

  const submitButton = form.querySelector('button[type="submit"]');
  const feedback = document.getElementById("login-feedback");
  const loginUrl = new URL(form.getAttribute("action") || window.location.href, window.location.origin);

  function setFeedback(message, type) {
    if (!feedback) return;
    feedback.textContent = message || "";
    feedback.classList.remove("ok", "error");
    if (type) feedback.classList.add(type);
  }

  form.addEventListener("submit", async (event) => {
    event.preventDefault();
    if (!form.reportValidity()) return;

    setFeedback("", "");
    if (submitButton) submitButton.disabled = true;

    try {
      const response = await fetch(loginUrl, {
        method: "POST",
        body: new FormData(form),
        credentials: "same-origin",
        redirect: "follow"
      });

      const finalUrl = new URL(response.url, window.location.origin);
      if (response.redirected && finalUrl.pathname !== loginUrl.pathname) {
        window.SiteLoader?.show({
          immediate: true,
          title: "Loading...",
          step: "Chargement du hub logistique..."
        });
        window.setTimeout(() => {
          window.location.href = response.url;
        }, 650);
        return;
      }

      const html = await response.text();
      document.open();
      document.write(html);
      document.close();
    } catch (_error) {
      if (submitButton) submitButton.disabled = false;
      setFeedback("Network error. Retry login.", "error");
    }
  });
})();
