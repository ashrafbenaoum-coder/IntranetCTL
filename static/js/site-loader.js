(function () {
  const overlay = document.getElementById("site-loader");
  const titleNode = document.getElementById("site-loader-title");
  const stepNode = document.getElementById("site-loader-step");
  if (!overlay || !titleNode || !stepNode) return;

  const DEFAULT_TITLE = "Loading...";
  const DEFAULT_STEP = "Connexion en cours...";

  let activeCount = 0;
  let showTimer = 0;
  let hideTimer = 0;
  let visibleSince = 0;

  function setCopy(options = {}) {
    titleNode.textContent = options.title || DEFAULT_TITLE;
    stepNode.textContent = options.step || DEFAULT_STEP;
  }

  function reveal() {
    showTimer = 0;
    if (overlay.classList.contains("is-visible")) return;
    visibleSince = Date.now();
    overlay.classList.add("is-visible");
    overlay.setAttribute("aria-hidden", "false");
    document.body.classList.add("site-loading");
  }

  function conceal() {
    hideTimer = 0;
    overlay.classList.remove("is-visible");
    overlay.setAttribute("aria-hidden", "true");
    document.body.classList.remove("site-loading");
  }

  function show(options = {}) {
    activeCount += 1;
    setCopy(options);
    window.clearTimeout(hideTimer);

    if (overlay.classList.contains("is-visible")) return;

    const delay = options.immediate ? 0 : Number(options.delay || 180);
    window.clearTimeout(showTimer);
    showTimer = window.setTimeout(reveal, delay);
  }

  function hide(force = false) {
    if (force) {
      activeCount = 0;
      window.clearTimeout(showTimer);
      window.clearTimeout(hideTimer);
      conceal();
      return;
    }

    activeCount = Math.max(0, activeCount - 1);
    if (activeCount > 0) return;

    window.clearTimeout(showTimer);
    if (!overlay.classList.contains("is-visible")) return;

    const elapsed = Date.now() - visibleSince;
    const minVisible = 420;
    hideTimer = window.setTimeout(conceal, Math.max(0, minVisible - elapsed));
  }

  function trackPromise(promise, options = {}) {
    show(options);
    return Promise.resolve(promise).finally(() => hide());
  }

  window.SiteLoader = {
    show,
    hide,
    trackPromise,
    setCopy
  };

  window.addEventListener("pageshow", () => hide(true));

  document.addEventListener("click", (event) => {
    const link = event.target.closest("a[href]");
    if (!link) return;
    if (event.defaultPrevented || event.button !== 0 || event.metaKey || event.ctrlKey || event.shiftKey || event.altKey) return;
    if (link.dataset.noLoader === "true" || link.target === "_blank" || link.hasAttribute("download")) return;

    const href = link.getAttribute("href") || "";
    if (!href || href.startsWith("#") || href.startsWith("javascript:")) return;

    let url;
    try {
      url = new URL(link.href, window.location.href);
    } catch (_error) {
      return;
    }
    if (url.origin !== window.location.origin) return;

    show({
      immediate: true,
      title: "Loading...",
      step: "Ouverture de la page en cours..."
    });
  });

  document.addEventListener("submit", (event) => {
    const form = event.target;
    if (!(form instanceof HTMLFormElement)) return;
    if (event.defaultPrevented) return;
    if (form.dataset.noLoader === "true") return;

    show({
      immediate: true,
      title: "Loading...",
      step: "Transmission des donnees en cours..."
    });
  });
})();
