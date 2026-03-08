(function () {
  const shell = document.getElementById("admin-shell");
  const tableBody = document.getElementById("movements-body");
  if (!shell || !tableBody) return;

  const csrfToken = document.querySelector('meta[name="csrf-token"]')?.content || "";
  const isAdmin = shell.dataset.isAdmin === "1";
  const currentUserId = Number(shell.dataset.userId || "0");
  const currentUsername = shell.dataset.username || "";
  const permissionSet = new Set(
    (shell.dataset.permissions || "")
      .split(",")
      .map((permission) => permission.trim())
      .filter(Boolean)
  );

  const SECTION_PERMISSIONS = {
    home: "home",
    library: "library",
    settings: "settings",
    analyse: "analyse",
    donnes: "donnes",
    users: "user_management"
  };

  const state = {
    selectedMovementId: null,
    selectedLibraryId: null,
    selectedLibraryName: "",
    selectedLibraryUser: null,
    selectedMonth: null,
    userMonths: [],
    libraryUsers: [],
    libraries: [],
    settingsLibraries: [],
    settingsLibraryId: null,
    settingsUsers: [],
    settingsAssignments: [],
    userManagementUsers: [],
    permissionsCatalog: [],
    expandedUserId: null,
    analyseType: "fiabilite",
    analyseView: "day",
    analyseMonth: "",
    analyseSelectedKey: "",
    analyseStats: null,
    donnesConnectionType: null,
    donnesConnectionValues: {},
    donnesDrivers: [],
    donnesPyodbcReady: null,
    scanner: null,
    scannerRunning: false
  };

  const navItems = Array.from(document.querySelectorAll(".nav-item"));
  const sections = {
    home: document.getElementById("section-home"),
    library: document.getElementById("section-library"),
    users: document.getElementById("section-users"),
    settings: document.getElementById("section-settings"),
    analyse: document.getElementById("section-analyse"),
    donnes: document.getElementById("section-donnes")
  };

  const usernameFilter = document.getElementById("username-filter");
  const form = document.getElementById("filters-form");
  const resetBtn = document.getElementById("reset-filters");
  const homeFeedback = document.getElementById("admin-feedback");

  const drawer = document.getElementById("details-drawer");
  const closeDrawerBtn = document.getElementById("close-drawer");
  const detailsList = document.getElementById("details-list");
  const deleteMovementBtn = document.getElementById("delete-movement-btn");

  const libraryFeedback = document.getElementById("library-feedback");
  const libraryPath = document.getElementById("library-screen-path");
  const libraryList = document.getElementById("library-list");
  const libraryCreateForm = document.getElementById("library-create-form");
  const libraryNameInput = document.getElementById("library-name");

  const libraryScreens = {
    root: document.getElementById("library-view-root"),
    users: document.getElementById("library-view-users"),
    months: document.getElementById("library-view-months"),
    days: document.getElementById("library-view-days"),
    entry: document.getElementById("library-view-entry")
  };

  const usersTitle = document.getElementById("library-users-title");
  const monthsTitle = document.getElementById("library-months-title");
  const daysTitle = document.getElementById("library-days-title");
  const entryTitle = document.getElementById("library-entry-title");
  const usersGrid = document.getElementById("library-users-grid");
  const monthsGrid = document.getElementById("library-months-grid");
  const daysWrap = document.getElementById("library-days-wrap");

  const backRoot = document.getElementById("back-library-root");
  const backUsers = document.getElementById("back-library-users");
  const backMonths = document.getElementById("back-library-months");
  const backEntry = document.getElementById("back-library-entry-root");

  const libraryForm = document.getElementById("library-data-form");
  const libraryFormFeedback = document.getElementById("library-data-feedback");
  const libraryLastRecord = document.getElementById("library-last-record");
  const supportInput = document.getElementById("library-support-number");
  const eanInput = document.getElementById("library-ean-code");
  const productInput = document.getElementById("library-product-code");
  const plusInput = document.getElementById("library-diff-plus");
  const minusInput = document.getElementById("library-diff-minus");
  const scanToggle = document.getElementById("library-scan-toggle");
  const scannerWrap = document.getElementById("library-scanner-wrap");
  const submitBtn = libraryForm.querySelector('button[type="submit"]');

  const themeSelect = document.getElementById("theme-select");
  const themeChoices = Array.from(document.querySelectorAll("[data-theme-choice]"));
  const themeStatus = document.getElementById("theme-status");
  const settingsLibrarySelect = document.getElementById("settings-library-select");
  const settingsUsersList = document.getElementById("settings-users-list");
  const settingsSaveAccessBtn = document.getElementById("settings-save-access");
  const settingsAccessFeedback = document.getElementById("settings-access-feedback");
  const donnesExtractionChoices = document.getElementById("donnes-extraction-choices");
  const donnesConnectionChoices = document.getElementById("donnes-connection-choices");
  const donnesFeedback = document.getElementById("donnes-feedback");
  const donnesConnectionPanel = document.getElementById("donnes-connection-panel");
  const donnesConnectionTitle = document.getElementById("donnes-connection-title");
  const donnesConnectionForm = document.getElementById("donnes-connection-form");
  const donnesConnectionLabel = document.getElementById("donnes-connection-label");
  const donnesConnectionInput = document.getElementById("donnes-connection-input");
  const donnesConnectionHint = document.getElementById("donnes-connection-hint");
  const donnesConnectionHelper = document.getElementById("donnes-connection-helper");
  const donnesOdbcDrivers = document.getElementById("donnes-odbc-drivers");
  const analyseTabs = Array.from(document.querySelectorAll("[data-analyse-type]"));
  const analyseViewButtons = Array.from(document.querySelectorAll("[data-analyse-view]"));
  const analyseMonthFilter = document.getElementById("analyse-month-filter");
  const analyseRefreshBtn = document.getElementById("analyse-refresh");
  const analyseDemarqueTags = document.getElementById("analyse-demarque-tags");
  const analyseSummary = document.getElementById("analyse-summary");
  const analyseChartWrap = document.getElementById("analyse-chart-wrap");
  const analyseTableWrap = document.getElementById("analyse-table-wrap");
  const analyseFeedback = document.getElementById("analyse-feedback");
  const usersCreateForm = document.getElementById("users-create-form");
  const usersCreateUsername = document.getElementById("users-create-username");
  const usersCreatePassword = document.getElementById("users-create-password");
  const usersCreateRole = document.getElementById("users-create-role");
  const usersCreatePerms = document.getElementById("users-create-perms");
  const usersCreateFeedback = document.getElementById("users-create-feedback");
  const usersAdminList = document.getElementById("users-admin-list");
  const usersAdminFeedback = document.getElementById("users-admin-feedback");
  const systemThemeMedia = typeof window.matchMedia === "function"
    ? window.matchMedia("(prefers-color-scheme: light)")
    : null;
  const ANALYSE_OBJECTIVE = 99.55;

  const DEFAULT_PERMISSION_CATALOG = [
    { id: "home", label: "Home" },
    { id: "library", label: "Library" },
    { id: "settings", label: "Settings" },
    { id: "analyse", label: "Analyse" },
    { id: "donnes", label: "Donnes" },
    { id: "user_management", label: "Gestion Users" }
  ];

  const DONNES_EXTRACTION_OPTIONS = [
    {
      id: "excel",
      label: "Excel",
      icon: "\u{1F4CA}",
      caption: "Export",
      note: "Tableur .xls avec les lignes filtrees.",
      accent: "excel"
    },
    {
      id: "word",
      label: "Word",
      icon: "\u{1F4DD}",
      caption: "Export",
      note: "Document RTF rapide a partager.",
      accent: "word"
    },
    {
      id: "text",
      label: "Text",
      icon: "\u{1F9FE}",
      caption: "Export",
      note: "Texte brut lisible en un clic.",
      accent: "text"
    },
    {
      id: "pdf",
      label: "PDF",
      icon: "\u{1F4D5}",
      caption: "Export",
      note: "Version PDF prete a imprimer.",
      accent: "pdf"
    }
  ];

  const DONNES_CONNECTION_OPTIONS = [
    {
      id: "odbc",
      label: "ODBC",
      icon: "\u{1F50C}",
      caption: "Windows",
      note: "Ouvre ODBC Administrator et affiche les drivers detectes.",
      accent: "odbc"
    },
    {
      id: "excel",
      label: "Excel",
      icon: "\u{1F4D7}",
      caption: "Source",
      note: "Choisit un fichier .xlsx, .xls ou .csv.",
      accent: "excel"
    },
    {
      id: "access",
      label: "Access",
      icon: "\u{1F5C4}",
      caption: "Source",
      note: "Choisit une base .mdb ou .accdb.",
      accent: "access"
    }
  ];

  const DONNES_CONNECTION_FIELDS = {
    odbc: {
      title: "Connexion ODBC",
      label: "Connection string ODBC",
      placeholder: "Driver={ODBC Driver 17 for SQL Server};Server=...;Database=...;UID=...;PWD=...;",
      hint: "1. Clique sur <code>Ouvrir ODBC Administrator</code>. 2. Verifie le driver installe sur le PC. 3. Colle ici la connection string finale puis teste.",
      helperLabel: "Ouvrir ODBC Administrator"
    },
    excel: {
      title: "Source Excel",
      label: "Chemin fichier Excel",
      placeholder: "C:\\\\data\\\\source.xlsx",
      hint: "Clique sur <code>Choisir fichier Windows</code> pour pointer <code>Interception.xlsx</code>. Les headers attendus pour Analyse sont <code>DATE CONTROLE</code>, <code>UVC CONTROLE</code>, <code>UVC ECART</code> et <code>Type de demarque</code>.",
      helperLabel: "Choisir fichier Windows"
    },
    access: {
      title: "Source Access",
      label: "Chemin fichier Access (.mdb/.accdb)",
      placeholder: "C:\\\\data\\\\base.accdb",
      hint: "Choisis la base Access via Windows, puis valide la connexion. Le driver Access doit etre installe sur le PC.",
      helperLabel: "Choisir fichier Windows"
    }
  };

  function escapeHtml(text) {
    return String(text)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#039;");
  }

  function setFeedback(node, message, type) {
    if (!node) return;
    node.textContent = message || "";
    node.classList.remove("ok", "error");
    if (type) node.classList.add(type);
  }

  function formatDate(value) {
    const date = new Date(value);
    return Number.isNaN(date.getTime()) ? value : date.toLocaleString();
  }

  function formatMonth(monthKey) {
    const date = new Date(`${monthKey}-01T00:00:00`);
    return Number.isNaN(date.getTime())
      ? monthKey
      : date.toLocaleDateString("fr-FR", { month: "long", year: "numeric" });
  }

  function basenamePath(value) {
    const normalized = String(value || "").trim();
    if (!normalized) return "";
    const parts = normalized.split(/[\\/]+/).filter(Boolean);
    return parts[parts.length - 1] || normalized;
  }

  function withSiteLoader(promise, options = {}) {
    if (!window.SiteLoader?.trackPromise) return promise;
    return window.SiteLoader.trackPromise(promise, options);
  }

  async function api(url, options = {}) {
    const fetchOptions = {
      credentials: "same-origin",
      ...options,
      headers: { ...(options.headers || {}) }
    };
    const loaderOptions = fetchOptions.loader || {};
    delete fetchOptions.loader;
    if (fetchOptions.body && typeof fetchOptions.body !== "string") {
      fetchOptions.headers["Content-Type"] = "application/json";
      fetchOptions.body = JSON.stringify(fetchOptions.body);
    }
    const res = await withSiteLoader(
      fetch(url, fetchOptions),
      {
        title: loaderOptions.title || "Loading...",
        step: loaderOptions.step || (fetchOptions.method === "GET" ? "Recuperation des donnees..." : "Transmission en cours...")
      }
    );
    const data = await res.json();
    if (!res.ok || data?.ok === false) throw new Error(data?.error || "Request failed.");
    return data;
  }

  function canAccessSection(section) {
    const permission = SECTION_PERMISSIONS[section];
    if (!permission) return true;
    return permissionSet.has(permission);
  }

  function firstAllowedSection() {
    const visibleItem = navItems.find((item) => canAccessSection(item.dataset.section));
    return visibleItem?.dataset.section || "home";
  }

  function applySectionPermissions() {
    navItems.forEach((item) => {
      item.classList.toggle("hidden", !canAccessSection(item.dataset.section));
    });
  }

  function showSection(section) {
    const targetSection = canAccessSection(section) ? section : firstAllowedSection();
    navItems.forEach((item) => item.classList.toggle("is-active", item.dataset.section === targetSection));
    Object.entries(sections).forEach(([name, node]) => {
      if (!node) return;
      node.classList.toggle("hidden", name !== targetSection || !canAccessSection(name));
    });
  }

  function showLibraryScreen(screen) {
    Object.entries(libraryScreens).forEach(([name, node]) => {
      if (!node) return;
      node.classList.toggle("hidden", name !== screen);
    });

    if (screen === "root") {
      libraryPath.textContent = "Library";
      return;
    }
    if (!isAdmin && screen === "entry") {
      libraryPath.textContent = `Library / ${state.selectedLibraryName} / Data`;
      return;
    }
    if (screen === "users") {
      libraryPath.textContent = `Library / ${state.selectedLibraryName} / Users`;
      return;
    }
    if (screen === "months") {
      libraryPath.textContent = `Library / ${state.selectedLibraryName} / Users / ${state.selectedLibraryUser.username}`;
      return;
    }
    if (screen === "days") {
      libraryPath.textContent = `Library / ${state.selectedLibraryName} / Users / ${state.selectedLibraryUser.username} / ${formatMonth(state.selectedMonth.key)}`;
    }
  }

  async function loadUsersFilter() {
    if (!isAdmin) {
      usernameFilter.innerHTML = "";
      const opt = document.createElement("option");
      opt.value = currentUsername;
      opt.textContent = currentUsername;
      usernameFilter.appendChild(opt);
      usernameFilter.value = currentUsername;
      usernameFilter.disabled = true;
      return;
    }
    const users = await api("/api/admin/users");
    usernameFilter.innerHTML = '<option value="">Tous</option>';
    users.forEach((u) => {
      const opt = document.createElement("option");
      opt.value = u.username;
      opt.textContent = u.username;
      usernameFilter.appendChild(opt);
    });
  }

  function movementsQuery() {
    const q = new URLSearchParams();
    if (isAdmin && form.username.value) q.set("username", form.username.value);
    if (form.date_from.value) q.set("date_from", form.date_from.value);
    if (form.date_to.value) q.set("date_to", form.date_to.value);
    const raw = q.toString();
    return raw ? `?${raw}` : "";
  }

  async function loadHomeMovements() {
    setFeedback(homeFeedback, "Loading data...", "");
    try {
      const rows = await api(`/api/admin/movements${movementsQuery()}`);
      tableBody.innerHTML = "";
      if (!rows.length) {
        tableBody.innerHTML = '<tr><td colspan="8">Aucun mouvement trouve.</td></tr>';
        setFeedback(homeFeedback, "No data for selected filters.", "");
        return;
      }
      rows.forEach((row) => {
        const tr = document.createElement("tr");
        tr.dataset.movementId = row.id;
        tr.innerHTML = `
          <td>${row.id}</td>
          <td>${escapeHtml(row.username)}</td>
          <td>${escapeHtml(formatDate(row.movement_date))}</td>
          <td>${escapeHtml(row.support_number)}</td>
          <td>${escapeHtml(row.ean_code)}</td>
          <td>${escapeHtml(row.product_code)}</td>
          <td>${row.diff_plus}</td>
          <td>${row.diff_minus}</td>
        `;
        tableBody.appendChild(tr);
      });
      setFeedback(homeFeedback, `${rows.length} mouvement(s) charges.`, "ok");
    } catch (error) {
      setFeedback(homeFeedback, error.message, "error");
    }
  }

  async function openMovementDetails(movementId) {
    const row = await api(`/api/admin/movements/${movementId}`);
    state.selectedMovementId = row.id;
    detailsList.innerHTML = `
      <dt>ID</dt><dd>${row.id}</dd>
      <dt>Controleur</dt><dd>${escapeHtml(row.username)}</dd>
      <dt>N Support</dt><dd>${escapeHtml(row.support_number)}</dd>
      <dt>Code EAN</dt><dd>${escapeHtml(row.ean_code)}</dd>
      <dt>Code produit</dt><dd>${escapeHtml(row.product_code)}</dd>
      <dt>Nb colis ecart +</dt><dd>${row.diff_plus}</dd>
      <dt>Nb colis ecart -</dt><dd>${row.diff_minus}</dd>
      <dt>Date mouvement</dt><dd>${escapeHtml(formatDate(row.movement_date))}</dd>
      <dt>Date creation</dt><dd>${escapeHtml(formatDate(row.created_at))}</dd>
    `;
    deleteMovementBtn.classList.remove("hidden");
    drawer.classList.remove("hidden");
  }

  async function deleteMovement(id) {
    await api(`/api/admin/movements/${id}`, { method: "DELETE", headers: { "X-CSRF-Token": csrfToken } });
  }

  async function deleteLibrary(libraryId) {
    await api(`/api/admin/libraries/${libraryId}`, { method: "DELETE", headers: { "X-CSRF-Token": csrfToken } });
  }

  function closeDrawer() {
    state.selectedMovementId = null;
    detailsList.innerHTML = "";
    deleteMovementBtn.classList.add("hidden");
    drawer.classList.add("hidden");
  }

  function groupMonths(days) {
    if (!Array.isArray(days) || !days.length) return [];
    const map = new Map();
    days.forEach((d) => {
      const key = String(d.day).slice(0, 7);
      if (!map.has(key)) map.set(key, { key, label: formatMonth(key), count: 0, days: [] });
      const month = map.get(key);
      const items = Array.isArray(d.items) ? d.items : [];
      const day = { ...d, items };
      month.count += Number(d.count || items.length || 0);
      month.days.push(day);
    });
    return Array.from(map.values()).sort((a, b) => b.key.localeCompare(a.key));
  }

  async function loadLibraries() {
    try {
      state.libraries = await api("/api/admin/libraries");
      if (!state.libraries.length) {
        libraryList.innerHTML = "<p class='hint-text'>Aucune library pour le moment.</p>";
        return;
      }
      libraryList.innerHTML = state.libraries
        .map((l) => {
          const openBtn = `<button type="button" class="btn btn-ghost" data-open-library-id="${l.id}">Ouvrir library</button>`;
          const deleteBtn = isAdmin
            ? `<button type="button" class="btn btn-danger btn-icon-delete" data-delete-library-id="${l.id}" title="Supprimer library ${escapeHtml(l.name)}" aria-label="Supprimer library ${escapeHtml(l.name)}">&#128465;</button>`
            : "";
          return `
            <div class="library-item">
              <div>
                <strong>${escapeHtml(l.name)}</strong><br>
                <span class="hint-text">${l.users_count} user(s)</span>
              </div>
              <div class="library-item-actions">
                ${openBtn}
                ${deleteBtn}
              </div>
            </div>
          `;
        })
        .join("");
      setFeedback(libraryFeedback, "", "");
    } catch (error) {
      setFeedback(libraryFeedback, error.message, "error");
    }
  }

  async function openLibraryAdmin(libraryId) {
    state.selectedLibraryId = libraryId;
    state.selectedLibraryName = state.libraries.find((l) => l.id === libraryId)?.name || "";
    const payload = await api(`/api/admin/libraries/${libraryId}/users`);
    state.libraryUsers = (payload.assigned_users || []).map((u) => ({
      ...u,
      id: String(u.id),
      role: u.role || "controller"
    }));
    usersTitle.textContent = `Library ${state.selectedLibraryName} / Users`;
    if (!state.libraryUsers.length) {
      usersGrid.innerHTML = "<p class='hint-text'>Aucun user assigne a cette library.</p>";
    } else {
      usersGrid.innerHTML = state.libraryUsers
        .map(
          (u) => `
            <article class="user-node">
              <h4>${escapeHtml(u.username)}</h4>
              <p class="hint-text">Role: ${escapeHtml(u.role)}</p>
              <button class="btn btn-primary" data-open-user-id="${escapeHtml(String(u.id))}">Ouvrir user</button>
            </article>
          `
        )
        .join("");
    }
    showLibraryScreen("users");
  }

  async function openLibraryUser(userId) {
    const user = state.libraryUsers.find((u) => String(u.id) === String(userId));
    if (!user) return;
    state.selectedLibraryUser = user;
    state.selectedMonth = null;
    const payload = await api(`/api/admin/libraries/${state.selectedLibraryId}/users/${encodeURIComponent(user.id)}/movements`);
    state.userMonths = groupMonths(payload.days || []);
    monthsTitle.textContent = `Library ${state.selectedLibraryName} / Users / ${user.username}`;
    if (!state.userMonths.length) {
      monthsGrid.innerHTML = "<p class='hint-text'>Aucune data pour ce user.</p>";
    } else {
      monthsGrid.innerHTML = state.userMonths
        .map((m) => `<article class="folder-item"><h4>${escapeHtml(m.label)}</h4><p class="hint-text">${m.count} data</p><button class="btn btn-ghost" data-open-month="${m.key}">Ouvrir dossier</button></article>`)
        .join("");
    }
    showLibraryScreen("months");
  }

  function renderMonthDays(monthKey) {
    const month = state.userMonths.find((m) => m.key === monthKey);
    if (!month) return;
    state.selectedMonth = month;
    daysTitle.textContent = `Library ${state.selectedLibraryName} / Users / ${state.selectedLibraryUser.username} / ${month.label}`;
    if (!month.days.length) {
      daysWrap.innerHTML = "<p class='hint-text'>Aucune data pour ce mois.</p>";
      showLibraryScreen("days");
      return;
    }
    daysWrap.innerHTML = month.days
      .map((d) => {
        const rows = d.items
          .map((item) => {
            const actions = `<button class="mini-btn" data-open-movement-id="${item.id}">Detail</button><button class="mini-btn" data-delete-movement-id="${item.id}">Delete</button>`;
            return `<tr><td>${escapeHtml(item.id)}</td><td>${escapeHtml(item.support_number)}</td><td>${escapeHtml(item.ean_code)}</td><td>${escapeHtml(item.product_code)}</td><td>${item.diff_plus}</td><td>${item.diff_minus}</td><td>${escapeHtml(formatDate(item.movement_date))}</td><td class="row-actions">${actions}</td></tr>`;
          })
          .join("");
        return `<article class="day-block"><div class="day-head"><strong>${escapeHtml(d.day)}</strong><span class="hint-text">${d.count} enregistrement(s)</span></div><div class="table-wrap"><table class="table-compact"><thead><tr><th>ID</th><th>Support</th><th>EAN</th><th>Produit</th><th>Ecart +</th><th>Ecart -</th><th>Date</th><th>Action</th></tr></thead><tbody>${rows}</tbody></table></div></article>`;
      })
      .join("");
    showLibraryScreen("days");
  }

  function openLibraryUserMode(libraryId) {
    state.selectedLibraryId = libraryId;
    state.selectedLibraryName = state.libraries.find((l) => l.id === libraryId)?.name || "";
    entryTitle.textContent = `Library ${state.selectedLibraryName} / Data`;
    showLibraryScreen("entry");
    setFormEnabled(true, "");
  }

  function setFormEnabled(enabled, message) {
    [supportInput, eanInput, productInput, plusInput, minusInput, scanToggle, submitBtn].forEach((n) => {
      n.disabled = !enabled;
    });
    if (!enabled) stopScanner();
    if (message) setFeedback(libraryFormFeedback, message, enabled ? "" : "error");
  }

  async function saveLibraryData(payload) {
    return api(`/api/admin/libraries/${state.selectedLibraryId}/users/${currentUserId}/movements`, {
      method: "POST",
      headers: { "X-CSRF-Token": csrfToken },
      body: payload
    });
  }

  async function startScanner() {
    if (state.scannerRunning || !state.selectedLibraryId) return;
    if (typeof Html5Qrcode !== "function") {
      setFeedback(libraryFormFeedback, "Scanner non disponible.", "error");
      return;
    }
    scannerWrap.classList.remove("hidden");
    state.scanner = new Html5Qrcode("library-reader");
    try {
      await state.scanner.start(
        { facingMode: "environment" },
        { fps: 10, qrbox: { width: 260, height: 150 }, formatsToSupport: [Html5QrcodeSupportedFormats.EAN_13, Html5QrcodeSupportedFormats.EAN_8, Html5QrcodeSupportedFormats.UPC_A, Html5QrcodeSupportedFormats.UPC_E, Html5QrcodeSupportedFormats.CODE_128] },
        (decoded) => {
          const cleaned = String(decoded || "").replace(/\D+/g, "");
          if (cleaned.length >= 8 && cleaned.length <= 14) {
            eanInput.value = cleaned;
            setFeedback(libraryFormFeedback, "EAN scanned successfully.", "ok");
            stopScanner();
          }
        },
        () => {}
      );
      state.scannerRunning = true;
      scanToggle.textContent = "Stop Scanner";
    } catch (_e) {
      scannerWrap.classList.add("hidden");
      setFeedback(libraryFormFeedback, "Camera access failed.", "error");
    }
  }

  async function stopScanner() {
    if (!state.scanner || !state.scannerRunning) {
      scannerWrap.classList.add("hidden");
      scanToggle.textContent = "Scanner EAN";
      return;
    }
    try {
      await state.scanner.stop();
      await state.scanner.clear();
    } catch (_e) {
      // ignore
    } finally {
      state.scanner = null;
      state.scannerRunning = false;
      scannerWrap.classList.add("hidden");
      scanToggle.textContent = "Scanner EAN";
    }
  }

  function resolveThemeMode(theme) {
    if (theme === "light" || theme === "dark") return theme;
    return systemThemeMedia?.matches ? "light" : "dark";
  }

  function updateThemeUi(preference, resolved) {
    document.body.classList.toggle("theme-light", resolved === "light");
    document.body.dataset.themePreference = preference;
    document.body.dataset.themeResolved = resolved;
    if (themeSelect) themeSelect.value = preference;

    themeChoices.forEach((choice) => {
      const isActive = choice.dataset.themeChoice === preference;
      choice.classList.toggle("is-active", isActive);
      choice.setAttribute("aria-checked", String(isActive));
    });

    if (!themeStatus) return;
    if (preference === "system") {
      themeStatus.textContent = `Theme synchronise avec Windows. Mode actif: ${resolved === "light" ? "Light" : "Dark"}.`;
      return;
    }
    themeStatus.textContent = `Theme fixe active: ${preference === "light" ? "Light" : "Dark"}.`;
  }

  function applyTheme(theme, options = {}) {
    const preference = theme === "light" || theme === "dark" ? theme : "system";
    const resolved = resolveThemeMode(preference);
    updateThemeUi(preference, resolved);
    if (options.persist === false) return;
    try {
      window.localStorage.setItem("admin_theme", preference);
    } catch (_e) {
      // ignore
    }
  }

  function resetDonnesConnectionPanel() {
    if (!donnesConnectionPanel || !donnesConnectionInput) return;
    donnesConnectionPanel.classList.add("hidden");
    donnesConnectionInput.value = "";
    state.donnesConnectionType = null;
    state.donnesDrivers = [];
    state.donnesPyodbcReady = null;
    if (donnesConnectionHelper) {
      donnesConnectionHelper.textContent = "Ouvrir Windows";
      donnesConnectionHelper.classList.add("hidden");
    }
    if (donnesConnectionHint) {
      donnesConnectionHint.innerHTML = "Analyse source headers: <code>date</code>/<code>jour</code>, <code>total_colis_prepare</code>, <code>taux_fiabilite</code>.";
    }
    if (donnesOdbcDrivers) {
      donnesOdbcDrivers.innerHTML = "";
      donnesOdbcDrivers.classList.add("hidden");
    }
  }

  function renderDonnesDrivers() {
    if (!donnesOdbcDrivers) return;
    if (state.donnesConnectionType !== "odbc") {
      donnesOdbcDrivers.innerHTML = "";
      donnesOdbcDrivers.classList.add("hidden");
      return;
    }

    const pills = [];
    if (state.donnesPyodbcReady === false) {
      pills.push('<span class="driver-pill is-warning">\u26A0 pyodbc non installe</span>');
    }
    if (Array.isArray(state.donnesDrivers) && state.donnesDrivers.length) {
      state.donnesDrivers.forEach((driver) => {
        pills.push(`<span class="driver-pill">\u{1F9E9} ${escapeHtml(driver)}</span>`);
      });
    } else {
      pills.push('<span class="driver-pill">\u2139 Aucun driver ODBC detecte automatiquement.</span>');
    }

    donnesOdbcDrivers.innerHTML = pills.join("");
    donnesOdbcDrivers.classList.remove("hidden");
  }

  function openDonnesConnection(type) {
    if (!donnesConnectionPanel || !donnesConnectionTitle || !donnesConnectionLabel || !donnesConnectionInput) return;
    const config = DONNES_CONNECTION_FIELDS[type];
    if (!config) return;
    state.donnesConnectionType = type;
    donnesConnectionTitle.textContent = config.title;
    donnesConnectionLabel.textContent = config.label;
    donnesConnectionInput.placeholder = config.placeholder;
    donnesConnectionInput.value = state.donnesConnectionValues[type] || "";
    if (donnesConnectionHint) donnesConnectionHint.innerHTML = config.hint;
    if (donnesConnectionHelper) {
      donnesConnectionHelper.textContent = config.helperLabel;
      donnesConnectionHelper.classList.remove("hidden");
    }
    if (type !== "odbc") {
      state.donnesDrivers = [];
      state.donnesPyodbcReady = null;
    }
    donnesConnectionPanel.classList.remove("hidden");
    renderDonnesDrivers();
    donnesConnectionInput.focus();
  }

  function buildDonnesExtractionPayload(format) {
    return {
      format,
      date_from: form?.date_from?.value || "",
      date_to: form?.date_to?.value || "",
      username: isAdmin ? form?.username?.value || "" : ""
    };
  }

  function buildDonnesExtractionUrl(format) {
    const query = new URLSearchParams();
    const payload = buildDonnesExtractionPayload(format);
    query.set("format", payload.format);
    if (payload.date_from) query.set("date_from", payload.date_from);
    if (payload.date_to) query.set("date_to", payload.date_to);
    if (payload.username) query.set("username", payload.username);
    return `/api/admin/donnes/extract?${query.toString()}`;
  }

  function fallbackBrowserExtraction(format, label) {
    const url = buildDonnesExtractionUrl(format);
    const a = document.createElement("a");
    a.href = url;
    a.target = "_blank";
    a.rel = "noopener";
    document.body.appendChild(a);
    a.click();
    a.remove();
    setFeedback(donnesFeedback, `Fenetre Windows indisponible. Extraction ${label} lancee dans le navigateur.`, "ok");
  }

  async function triggerDonnesExtraction(format, label) {
    setFeedback(donnesFeedback, `Ouverture de la fenetre Windows pour ${label}...`, "");
    try {
      const result = await api("/api/admin/donnes/extract/save", {
        method: "POST",
        headers: { "X-CSRF-Token": csrfToken },
        body: buildDonnesExtractionPayload(format)
      });
      if (result.cancelled) {
        setFeedback(donnesFeedback, `Extraction ${label} annulee.`, "");
        return;
      }
      const savedName = basenamePath(result.saved_path);
      setFeedback(
        donnesFeedback,
        `\u2705 ${label} enregistre dans Windows: ${savedName} (${result.rows_count || 0} ligne(s)).`,
        "ok"
      );
    } catch (error) {
      if (/Windows desktop actions|PowerShell/i.test(error.message || "")) {
        fallbackBrowserExtraction(format, label);
        return;
      }
      setFeedback(donnesFeedback, error.message, "error");
    }
  }

  async function submitDonnesConnectionTest(forcedType = null, options = {}) {
    if (!donnesConnectionInput) return;
    const connectionType = forcedType || state.donnesConnectionType;
    if (!connectionType) return;
    const value = donnesConnectionInput.value.trim();
    if (!value) {
      setFeedback(donnesFeedback, "Valeur de connexion requise.", "error");
      return;
    }
    state.donnesConnectionType = connectionType;
    state.donnesConnectionValues[connectionType] = value;
    const payload = { type: connectionType };
    if (connectionType === "odbc") payload.connection_string = value;
    else payload.path = value;

    setFeedback(donnesFeedback, options.loadingMessage || "Test connection en cours...", "");
    try {
      const result = await api("/api/admin/donnes/connections/test", {
        method: "POST",
        headers: { "X-CSRF-Token": csrfToken },
        body: payload
      });
      setFeedback(donnesFeedback, result.message || "Connection valide.", "ok");
    } catch (error) {
      setFeedback(donnesFeedback, error.message, "error");
    }
  }

  async function pickDonnesSourceFile(type, autoTest = true) {
    openDonnesConnection(type);
    setFeedback(donnesFeedback, "Ouverture de la selection Windows...", "");
    try {
      const result = await api("/api/admin/donnes/connections/pick-file", {
        method: "POST",
        headers: { "X-CSRF-Token": csrfToken },
        body: { type }
      });
      if (result.cancelled) {
        setFeedback(donnesFeedback, `Selection ${type} annulee.`, "");
        return;
      }
      const selectedPath = String(result.path || "");
      donnesConnectionInput.value = selectedPath;
      state.donnesConnectionValues[type] = selectedPath;
      setFeedback(donnesFeedback, `\u{1F4C2} ${basenamePath(selectedPath)} selectionne.`, "ok");
      if (autoTest) {
        await submitDonnesConnectionTest(type, {
          loadingMessage: `Validation ${String(type).toUpperCase()} en cours...`
        });
      }
    } catch (error) {
      setFeedback(donnesFeedback, error.message, "error");
    }
  }

  async function launchOdbcAdministrator() {
    openDonnesConnection("odbc");
    setFeedback(donnesFeedback, "Ouverture de ODBC Administrator...", "");
    try {
      const result = await api("/api/admin/donnes/connections/open-odbc-admin", {
        method: "POST",
        headers: { "X-CSRF-Token": csrfToken }
      });
      state.donnesDrivers = Array.isArray(result.drivers) ? result.drivers : [];
      state.donnesPyodbcReady = result.pyodbc_installed !== false;
      renderDonnesDrivers();
      const runtimeNote = result.pyodbc_installed === false
        ? " pyodbc n'est pas installe sur ce poste serveur."
        : "";
      setFeedback(donnesFeedback, `${result.message || "ODBC Administrator ouvert."}${runtimeNote}`, "ok");
    } catch (error) {
      state.donnesDrivers = [];
      state.donnesPyodbcReady = null;
      renderDonnesDrivers();
      setFeedback(donnesFeedback, error.message, "error");
    }
  }

  function renderDonnesCard(item, group) {
    return `
          <button
            type="button"
            class="donnes-choice"
            data-donnes-group="${escapeHtml(group)}"
            data-donnes-item="${escapeHtml(item.id)}"
            data-donnes-label="${escapeHtml(item.label)}"
            data-donnes-accent="${escapeHtml(item.accent || item.id)}"
          >
            <span class="donnes-choice-icon" aria-hidden="true">${escapeHtml(item.icon || "\u2022")}</span>
            <span class="donnes-choice-copy">
              <span class="donnes-choice-caption">${escapeHtml(item.caption || group)}</span>
              <span class="donnes-choice-label">${escapeHtml(item.label)}</span>
              <span class="donnes-choice-note">${escapeHtml(item.note || "")}</span>
            </span>
            <span class="donnes-choice-arrow" aria-hidden="true">\u279C</span>
          </button>
        `;
  }

  function renderDonnes() {
    if (donnesExtractionChoices) {
      donnesExtractionChoices.innerHTML = DONNES_EXTRACTION_OPTIONS.map((item) => renderDonnesCard(item, "extraction")).join("");
    }
    if (donnesConnectionChoices) {
      donnesConnectionChoices.innerHTML = DONNES_CONNECTION_OPTIONS.map((item) => renderDonnesCard(item, "connection")).join("");
    }
    setFeedback(donnesFeedback, "", "");
  }

  function assignedUsersSetForLibrary(libraryId) {
    const selectedId = Number(libraryId || 0);
    const set = new Set();
    state.settingsAssignments.forEach((entry) => {
      if (Number(entry.library_id) !== selectedId) return;
      set.add(Number(entry.user_id));
    });
    return set;
  }

  function checkedSettingsUsers() {
    if (!settingsUsersList) return [];
    return Array.from(settingsUsersList.querySelectorAll("input[data-settings-user-id]:checked"))
      .map((node) => Number(node.dataset.settingsUserId || "0"))
      .filter((id) => Number.isInteger(id) && id > 0);
  }

  function renderSettingsUsers() {
    if (!settingsUsersList) return;
    const libraryId = Number(state.settingsLibraryId || 0);
    if (!libraryId) {
      settingsUsersList.innerHTML = "<p class='hint-text'>Choisis une library.</p>";
      if (settingsSaveAccessBtn) settingsSaveAccessBtn.disabled = true;
      return;
    }
    if (!state.settingsUsers.length) {
      settingsUsersList.innerHTML = "<p class='hint-text'>Aucun user disponible.</p>";
      if (settingsSaveAccessBtn) settingsSaveAccessBtn.disabled = true;
      return;
    }

    const assignedSet = assignedUsersSetForLibrary(libraryId);
    settingsUsersList.innerHTML = state.settingsUsers
      .map((user) => {
        const userId = Number(user.id || 0);
        const role = String(user.role || "").toLowerCase();
        const assignable = role === "controller";
        const checked = assignable && assignedSet.has(userId) ? "checked" : "";
        const disabled = assignable ? "" : "disabled";
        const roleClass = role === "admin" ? "role-admin" : "role-controller";
        const roleLabel = role || "unknown";
        const note = assignable ? "" : "<span class='hint-text'>Acces global admin</span>";
        return `
          <label class="settings-user-item">
            <span class="settings-user-main">
              <strong>${escapeHtml(user.username)}</strong>
              <span class="role-badge ${roleClass}">${escapeHtml(roleLabel)}</span>
            </span>
            <span class="settings-user-main">
              ${note}
              <input type="checkbox" data-settings-user-id="${userId}" ${checked} ${disabled}>
            </span>
          </label>
        `;
      })
      .join("");
    if (settingsSaveAccessBtn) settingsSaveAccessBtn.disabled = false;
  }

  function renderSettingsLibraries() {
    if (!settingsLibrarySelect || !settingsUsersList) return;
    if (!state.settingsLibraries.length) {
      settingsLibrarySelect.innerHTML = "";
      settingsLibrarySelect.disabled = true;
      state.settingsLibraryId = null;
      settingsUsersList.innerHTML = "<p class='hint-text'>Aucune library. Cree une library dans l'onglet Library.</p>";
      if (settingsSaveAccessBtn) settingsSaveAccessBtn.disabled = true;
      return;
    }

    const current = Number(state.settingsLibraryId || 0);
    const hasCurrent = state.settingsLibraries.some((library) => Number(library.id) === current);
    state.settingsLibraryId = hasCurrent ? current : Number(state.settingsLibraries[0].id);
    settingsLibrarySelect.innerHTML = state.settingsLibraries
      .map((library) => `<option value="${library.id}">${escapeHtml(library.name)} (${Number(library.users_count || 0)} user(s))</option>`)
      .join("");
    settingsLibrarySelect.value = String(state.settingsLibraryId);
    settingsLibrarySelect.disabled = false;
    renderSettingsUsers();
  }

  async function loadSettingsAccess(silent = false) {
    if (!isAdmin || !settingsLibrarySelect || !settingsUsersList) return;
    if (!silent) setFeedback(settingsAccessFeedback, "Chargement des acces...", "");
    try {
      const payload = await api("/api/admin/settings/library-access");
      state.settingsLibraries = payload.libraries || [];
      state.settingsUsers = payload.users || [];
      state.settingsAssignments = payload.assignments || [];
      renderSettingsLibraries();
      if (!silent) setFeedback(settingsAccessFeedback, "", "");
    } catch (error) {
      if (!silent) setFeedback(settingsAccessFeedback, error.message, "error");
    }
  }

  async function saveSettingsAccess() {
    const libraryId = Number(state.settingsLibraryId || 0);
    if (!libraryId) return;
    const userIds = checkedSettingsUsers();
    setFeedback(settingsAccessFeedback, "Sauvegarde en cours...", "");
    try {
      const result = await api(`/api/admin/settings/libraries/${libraryId}/users`, {
        method: "PUT",
        headers: { "X-CSRF-Token": csrfToken },
        body: { user_ids: userIds }
      });
      const assignedUsers = result.assigned_users || [];
      state.settingsAssignments = state.settingsAssignments.filter((entry) => Number(entry.library_id) !== libraryId);
      assignedUsers.forEach((user) => state.settingsAssignments.push({ library_id: libraryId, user_id: user.id }));
      state.settingsLibraries = state.settingsLibraries.map((library) =>
        Number(library.id) === libraryId ? { ...library, users_count: assignedUsers.length } : library
      );
      renderSettingsLibraries();
      setFeedback(settingsAccessFeedback, `Acces sauvegarde (${assignedUsers.length} user(s)).`, "ok");
      await loadLibraries();
    } catch (error) {
      setFeedback(settingsAccessFeedback, error.message, "error");
    }
  }

  function currentPermissionCatalog() {
    return state.permissionsCatalog.length ? state.permissionsCatalog : DEFAULT_PERMISSION_CATALOG;
  }

  function defaultPermissionsForRole(role) {
    const defaults = role === "admin"
      ? DEFAULT_PERMISSION_CATALOG.map((perm) => perm.id)
      : ["home", "library", "settings"];
    const available = new Set(currentPermissionCatalog().map((perm) => perm.id));
    return defaults.filter((permissionId) => available.has(permissionId));
  }

  function selectedCreatePermissions() {
    if (!usersCreatePerms) return [];
    return Array.from(usersCreatePerms.querySelectorAll("[data-create-permission]:checked"))
      .map((node) => String(node.value || "").trim().toLowerCase())
      .filter(Boolean);
  }

  function selectedUserPermissions(card) {
    return Array.from(card.querySelectorAll("[data-user-permission]:checked"))
      .map((node) => String(node.value || "").trim().toLowerCase())
      .filter(Boolean);
  }

  function permissionsMarkup(selectedPermissions, disabled = false) {
    const selected = new Set((selectedPermissions || []).map((perm) => String(perm || "").toLowerCase()));
    const disabledAttr = disabled ? "disabled" : "";
    return currentPermissionCatalog()
      .map(
        (perm) => `
          <label class="perm-check">
            <input type="checkbox" value="${escapeHtml(perm.id)}" data-user-permission ${selected.has(perm.id) ? "checked" : ""} ${disabledAttr}>
            ${escapeHtml(perm.label)}
          </label>
        `
      )
      .join("");
  }

  function renderCreatePermissions() {
    if (!usersCreatePerms || !usersCreateRole) return;
    const defaults = defaultPermissionsForRole(usersCreateRole.value || "controller");
    const selected = new Set(defaults);
    usersCreatePerms.innerHTML = currentPermissionCatalog()
      .map(
        (perm) => `
          <label class="perm-check">
            <input type="checkbox" value="${escapeHtml(perm.id)}" data-create-permission ${selected.has(perm.id) ? "checked" : ""}>
            ${escapeHtml(perm.label)}
          </label>
        `
      )
      .join("");
  }

  function renderUserManagementList() {
    if (!usersAdminList) return;
    if (!state.userManagementUsers.length) {
      usersAdminList.innerHTML = "<p class='hint-text'>Aucun user pour le moment.</p>";
      return;
    }
    usersAdminList.innerHTML = state.userManagementUsers
      .map((user) => {
        const userId = Number(user.id);
        const isSelf = userId === currentUserId;
        const role = String(user.role || "controller").toLowerCase();
        const roleClass = role === "admin" ? "role-admin" : "role-controller";
        const blocked = Boolean(user.is_blocked);
        const permissions = Array.isArray(user.permissions) ? user.permissions : [];
        const isExpanded = Number(state.expandedUserId) === userId;
        const disabledAttr = isSelf ? "disabled" : "";
        const roleIcon = role === "admin" ? "&#128737;" : "&#128100;";
        const statusIcon = blocked ? "&#128274;" : "&#128275;";
        const createdAt = user.created_at ? formatDate(user.created_at) : "-";
        return `
          <article class="user-admin-card" data-user-card data-user-id="${user.id}">
            <button type="button" class="user-admin-toggle" data-toggle-user-id="${user.id}">
              <span class="user-admin-id">
                <span class="user-icon">${roleIcon}</span>
                <span>
                  <strong>${escapeHtml(user.username)}</strong>
                  <span class="user-admin-meta">${isSelf ? "Session actuelle" : escapeHtml(role)}</span>
                </span>
              </span>
              <span class="user-admin-badges">
                <span class="icon-pill" title="Role">${roleIcon}</span>
                <span class="icon-pill" title="Status">${statusIcon}</span>
                <span class="icon-pill" title="Permissions">&#128273; ${permissions.length}</span>
                <span class="icon-pill chevron ${isExpanded ? "is-open" : ""}" title="Options">&#9662;</span>
              </span>
            </button>

            <div class="user-admin-details ${isExpanded ? "" : "hidden"}">
              <div class="user-admin-meta">
                ${blocked ? "Bloque" : "Actif"} | Cree: ${escapeHtml(createdAt)}
              </div>

              <div class="user-admin-grid">
                <div>
                  <label>Role</label>
                  <select data-user-role ${disabledAttr}>
                    <option value="controller" ${role === "controller" ? "selected" : ""}>Controller</option>
                    <option value="admin" ${role === "admin" ? "selected" : ""}>Admin</option>
                  </select>
                </div>
                <div>
                  <label>Bloquer le compte</label>
                  <input type="checkbox" data-user-blocked ${blocked ? "checked" : ""} ${disabledAttr}>
                </div>
              </div>

              <div>
                <label>Permissions</label>
                <div class="users-perm-grid">
                  ${permissionsMarkup(permissions, isSelf)}
                </div>
              </div>

              <div class="user-admin-actions">
                <button type="button" class="btn btn-primary" data-save-user-id="${user.id}" ${disabledAttr}>Sauvegarder</button>
                <button type="button" class="btn btn-secondary" data-reset-user-id="${user.id}">Reset password</button>
                <button type="button" class="btn btn-danger" data-delete-user-id="${user.id}" ${disabledAttr}>Supprimer</button>
              </div>
            </div>
          </article>
        `;
      })
      .join("");
  }

  async function loadUserManagement(silent = false) {
    if (!isAdmin || !usersAdminList) return;
    if (!silent) setFeedback(usersAdminFeedback, "Chargement users...", "");
    try {
      const payload = await api("/api/admin/user-management/users");
      const catalog = Array.isArray(payload.permissions_catalog) ? payload.permissions_catalog : [];
      state.permissionsCatalog = catalog.length
        ? catalog
            .map((perm) => ({
              id: String(perm.id || "").trim().toLowerCase(),
              label: String(perm.label || perm.id || "").trim()
            }))
            .filter((perm) => perm.id)
        : [...DEFAULT_PERMISSION_CATALOG];
      state.userManagementUsers = Array.isArray(payload.users) ? payload.users : [];
      if (!state.userManagementUsers.some((user) => Number(user.id) === Number(state.expandedUserId))) {
        state.expandedUserId = null;
      }
      renderCreatePermissions();
      renderUserManagementList();
      if (!silent) setFeedback(usersAdminFeedback, `${state.userManagementUsers.length} user(s) charges.`, "ok");
    } catch (error) {
      if (!silent) setFeedback(usersAdminFeedback, error.message, "error");
    }
  }

  function dayKey(value) {
    return String(value || "").slice(0, 10);
  }

  function monthKey(value) {
    return String(value || "").slice(0, 7);
  }

  function formatDayLabel(key) {
    const parsed = new Date(`${key}T00:00:00`);
    if (Number.isNaN(parsed.getTime())) return key;
    return parsed.toLocaleDateString("fr-FR", { day: "2-digit", month: "2-digit" });
  }

  function formatFullDayLabel(key) {
    const parsed = new Date(`${key}T00:00:00`);
    if (Number.isNaN(parsed.getTime())) return key;
    return parsed.toLocaleDateString("fr-FR");
  }

  function formatMonthLabel(key) {
    const parsed = new Date(`${key}-01T00:00:00`);
    if (Number.isNaN(parsed.getTime())) return key;
    return parsed.toLocaleDateString("fr-FR", { month: "long", year: "numeric" });
  }

  function clampPercent(value) {
    const numeric = Number(value || 0);
    if (!Number.isFinite(numeric)) return 0;
    if (numeric < 0) return 0;
    if (numeric > 100) return 100;
    return numeric;
  }

  function analyzeNumber(value) {
    const numeric = Number(value || 0);
    return Number.isFinite(numeric) ? numeric : 0;
  }

  function analyseQuery() {
    const q = new URLSearchParams();
    if (form?.date_from?.value) q.set("date_from", form.date_from.value);
    if (form?.date_to?.value) q.set("date_to", form.date_to.value);
    const raw = q.toString();
    return raw ? `?${raw}` : "";
  }

  function normalizeAnalysePayload(payload) {
    const source = payload?.source || {};
    const totals = payload?.totals || {};
    const filters = payload?.filters || {};
    const normalizePoints = (points, keyName) => (Array.isArray(points) ? points : [])
      .map((point) => ({
        key: keyName === "month" ? monthKey(point?.month) : dayKey(point?.day),
        value: clampPercent(point?.value)
      }))
      .filter((point) => point.key);
    const normalizeDetails = (details, keyName) => (Array.isArray(details) ? details : [])
      .map((detail) => ({
        key: keyName === "month" ? monthKey(detail?.month) : dayKey(detail?.day),
        month: monthKey(detail?.month || detail?.day),
        uvcControle: analyzeNumber(detail?.uvc_controle),
        uvcEcart: analyzeNumber(detail?.uvc_ecart),
        uvcEcartAll: analyzeNumber(detail?.uvc_ecart_all),
        uvcLivre: analyzeNumber(detail?.uvc_livre),
        uvcLivreDem: analyzeNumber(detail?.uvc_livre_dem),
        rowsCount: analyzeNumber(detail?.rows_count),
        tauxAvecDem: clampPercent(detail?.taux_avec_dem),
        tauxCorrige: clampPercent(detail?.taux_corrige),
        tauxEcarts: clampPercent(detail?.taux_ecarts),
        ecartObjectif: analyzeNumber(detail?.ecart_objectif)
      }))
      .filter((detail) => detail.key);

    return {
      totalUvcControle: analyzeNumber(totals.total_uvc_controle),
      totalUvcEcart: analyzeNumber(totals.total_uvc_ecart),
      totalUvcEcartAll: analyzeNumber(totals.total_uvc_ecart_all),
      totalUvcLivre: analyzeNumber(totals.total_uvc_livre),
      totalUvcLivreDem: analyzeNumber(totals.total_uvc_livre_dem),
      totalRows: analyzeNumber(totals.rows_count),
      tauxAvecDem: clampPercent(totals.taux_avec_dem),
      tauxFiabilite: clampPercent(totals.taux_corrige),
      tauxEcarts: clampPercent(totals.taux_ecarts),
      ecartObjectif: analyzeNumber(totals.ecart_objectif),
      dailyFiabilite: normalizePoints(payload?.daily_fiabilite, "day"),
      dailyEcarts: normalizePoints(payload?.daily_ecarts, "day"),
      monthlyFiabilite: normalizePoints(payload?.monthly_fiabilite, "month"),
      monthlyEcarts: normalizePoints(payload?.monthly_ecarts, "month"),
      dailyDetails: normalizeDetails(payload?.daily_details, "day"),
      monthlyDetails: normalizeDetails(payload?.monthly_details, "month"),
      availableMonths: (Array.isArray(payload?.available_months) ? payload.available_months : [])
        .map((value) => monthKey(value))
        .filter(Boolean),
      demarqueTypes: (Array.isArray(filters?.demarque_types) ? filters.demarque_types : [])
        .map((value) => String(value || "").trim())
        .filter(Boolean),
      sourceLabel: String(source.label || ""),
      sourceType: String(source.type || "")
    };
  }

  function filterAnalysePoints(points) {
    const list = Array.isArray(points) ? points : [];
    if (!state.analyseMonth) return list;
    return list.filter((point) => String(point.key || "").startsWith(state.analyseMonth));
  }

  function filterAnalyseDetails(details) {
    const list = Array.isArray(details) ? details : [];
    if (!state.analyseMonth) return list;
    return list.filter((detail) => String(detail.month || detail.key || "").startsWith(state.analyseMonth));
  }

  function summarizeAnalyseDetails(details) {
    const summary = {
      totalUvcControle: 0,
      totalUvcEcart: 0,
      totalUvcEcartAll: 0,
      totalUvcLivre: 0,
      totalUvcLivreDem: 0,
      totalRows: 0
    };
    details.forEach((detail) => {
      summary.totalUvcControle += analyzeNumber(detail?.uvcControle);
      summary.totalUvcEcart += analyzeNumber(detail?.uvcEcart);
      summary.totalUvcEcartAll += analyzeNumber(detail?.uvcEcartAll);
      summary.totalUvcLivre += analyzeNumber(detail?.uvcLivre);
      summary.totalUvcLivreDem += analyzeNumber(detail?.uvcLivreDem);
      summary.totalRows += analyzeNumber(detail?.rowsCount);
    });
    summary.tauxAvecDem = summary.totalUvcControle > 0
      ? clampPercent((summary.totalUvcLivreDem / summary.totalUvcControle) * 100)
      : 0;
    summary.tauxFiabilite = summary.totalUvcControle > 0
      ? clampPercent((summary.totalUvcLivre / summary.totalUvcControle) * 100)
      : 0;
    summary.tauxEcarts = clampPercent(100 - summary.tauxFiabilite);
    summary.ecartObjectif = ANALYSE_OBJECTIVE - summary.tauxFiabilite;
    summary.tauxMoyenne = details.length
      ? details.reduce((sum, detail) => sum + clampPercent(detail?.tauxCorrige), 0) / details.length
      : 0;
    return summary;
  }

  function formatAnalyseNumber(value) {
    return String(Math.round(analyzeNumber(value)));
  }

  function formatAnalysePercent(value) {
    return `${Number(value || 0).toFixed(2).replace(".", ",")} %`;
  }

  function buildObjectiveBadgeSafe(value) {
    const delta = Number(value || 0);
    const isBelow = delta < 0;
    const arrow = isBelow ? "&#9660;" : "&#9650;";
    const css = isBelow ? "is-down" : "is-up";
    return `<span class="analyse-objective-badge ${css}">${arrow} ${formatAnalysePercent(Math.abs(delta))}</span>`;
  }

  function buildObjectiveBadge(value) {
    const delta = Number(value || 0);
    const isBelow = delta < 0;
    const arrow = isBelow ? "▼" : "▲";
    const css = isBelow ? "is-down" : "is-up";
    return `<span class="analyse-objective-badge ${css}">${arrow} ${formatAnalysePercent(Math.abs(delta))}</span>`;
  }

  function buildAnalyseTable(details, summary, viewMode, selectedKey = "") {
    const isMonthly = viewMode === "month";
    const title = isMonthly ? "Tableau mensuel" : "Tableau journalier";
    const headerDate = isMonthly ? "Mois" : "DATE CONTROLE";
    const selectedLabel = selectedKey
      ? (isMonthly ? formatMonthLabel(selectedKey) : formatFullDayLabel(selectedKey))
      : "";
    const titleLabel = selectedLabel ? `${title} - ${selectedLabel}` : title;
    if (!details.length) {
      return `<h3>${titleLabel}</h3><p class="hint-text">Aucune ligne a afficher pour ce filtre.</p>`;
    }

    const rowsHtml = details
      .map((detail) => `
        <tr>
          <td>${escapeHtml(isMonthly ? formatMonthLabel(detail.key) : formatFullDayLabel(detail.key))}</td>
          <td>${formatAnalyseNumber(detail.uvcControle)}</td>
          <td>${formatAnalyseNumber(detail.uvcLivre)}</td>
          <td>${formatAnalyseNumber(detail.uvcEcart)}</td>
          <td>${formatAnalysePercent(detail.tauxAvecDem)}</td>
          <td>${formatAnalysePercent(detail.tauxCorrige)}</td>
          <td>${formatAnalysePercent(detail.tauxCorrige)}</td>
          <td>${formatAnalysePercent(detail.ecartObjectif)}</td>
        </tr>
      `)
      .join("");

    return `
      <h3>${titleLabel}</h3>
      <div class="analyse-table-shell">
        <table class="analyse-table">
          <thead>
            <tr>
              <th>${headerDate}</th>
              <th>UVC CONTROLE</th>
              <th>UVCLIV</th>
              <th>UVCECART</th>
              <th>Taux avec DEM</th>
              <th>Taux Moyenne</th>
              <th>Taux corrige</th>
              <th>%Ecart</th>
            </tr>
          </thead>
          <tbody>
            ${rowsHtml}
          </tbody>
          <tfoot>
            <tr>
              <th>Total</th>
              <th>${formatAnalyseNumber(summary.totalUvcControle)}</th>
              <th>${formatAnalyseNumber(summary.totalUvcLivre)}</th>
              <th>${formatAnalyseNumber(summary.totalUvcEcart)}</th>
              <th>${formatAnalysePercent(summary.tauxAvecDem)}</th>
              <th>${formatAnalysePercent(summary.tauxMoyenne)}</th>
              <th>${formatAnalysePercent(summary.tauxFiabilite)}</th>
              <th>${formatAnalysePercent(ANALYSE_OBJECTIVE - summary.tauxMoyenne)}</th>
            </tr>
          </tfoot>
        </table>
      </div>
    `;
  }

  function renderAnalyseToolbar() {
    if (analyseMonthFilter) {
      const current = state.analyseMonth || "";
      analyseMonthFilter.innerHTML = [
        '<option value="">Tous les mois</option>',
        ...((state.analyseStats?.availableMonths || []).map((value) => {
          const selected = current === value ? " selected" : "";
          return `<option value="${escapeHtml(value)}"${selected}>${escapeHtml(formatMonthLabel(value))}</option>`;
        }))
      ].join("");
      analyseMonthFilter.value = current;
      analyseMonthFilter.disabled = !(state.analyseStats?.availableMonths || []).length;
    }

    analyseViewButtons.forEach((button) => {
      button.classList.toggle("is-active", button.dataset.analyseView === state.analyseView);
    });

    if (analyseDemarqueTags) {
      const tags = state.analyseStats?.demarqueTypes || [];
      analyseDemarqueTags.innerHTML = tags.length
        ? tags.map((label) => `<span class="analyse-tag">${escapeHtml(label)}</span>`).join("")
        : "";
    }
  }

  function buildAnalyseChart(points, objective, objectiveMode, subtitle, labelMode) {
    const isMonthly = labelMode === "month";
    const chartTitle = isMonthly ? "Synthese mensuelle" : "Synthese journaliere";
    if (!points.length) {
      return `<h3>${chartTitle}</h3><p class="hint-text">${escapeHtml(subtitle)}</p><p class="hint-text">Aucune donnee disponible.</p>`;
    }

    const width = 920;
    const height = 320;
    const padLeft = 50;
    const padRight = 16;
    const padTop = 22;
    const padBottom = 52;
    const plotWidth = width - padLeft - padRight;
    const plotHeight = height - padTop - padBottom;
    const step = points.length > 1 ? plotWidth / (points.length - 1) : 0;
    const yFromValue = (value) => padTop + ((100 - clampPercent(value)) / 100) * plotHeight;

    const coords = points.map((point, index) => {
      const x = padLeft + (points.length > 1 ? step * index : plotWidth / 2);
      const y = yFromValue(point.value);
      const bad = objectiveMode === "min" ? point.value < objective : point.value > objective;
      return { ...point, x, y, bad };
    });

    const polyline = coords.map((point) => `${point.x.toFixed(2)},${point.y.toFixed(2)}`).join(" ");
    const objectiveY = yFromValue(objective);
    const ticks = [0, 25, 50, 75, 100];
    const yGuides = ticks
      .map((tick) => {
        const y = yFromValue(tick);
        return `<line x1="${padLeft}" y1="${y.toFixed(2)}" x2="${(width - padRight).toFixed(2)}" y2="${y.toFixed(2)}" class="analyse-grid-line"></line><text x="${padLeft - 8}" y="${(y + 4).toFixed(2)}" class="analyse-axis-text">${tick}%</text>`;
      })
      .join("");

    const xLabelStep = Math.max(1, Math.ceil(coords.length / 8));
    const xLabels = coords
      .filter((_, index) => index % xLabelStep === 0 || index === coords.length - 1)
      .map(
        (point) => `
          <text x="${point.x.toFixed(2)}" y="${(height - 22).toFixed(2)}" text-anchor="middle" class="analyse-axis-text">${escapeHtml(isMonthly ? formatMonthLabel(point.key) : formatDayLabel(point.key))}</text>
        `
      )
      .join("");

    const pointDots = coords
      .map(
        (point, index) => {
          if (point.bad && objectiveMode === "min") {
            const arrowStartY = Math.max(8, padTop - 10);
            const arrowHeadY = Math.max(arrowStartY + 10, point.y - 3);
            return `
              <g class="analyse-arrow-marker analyse-arrow-marker-down">
                <line x1="${point.x.toFixed(2)}" y1="${arrowStartY.toFixed(2)}" x2="${point.x.toFixed(2)}" y2="${(arrowHeadY - 7).toFixed(2)}"></line>
                <path d="M ${(point.x - 6).toFixed(2)} ${(arrowHeadY - 7).toFixed(2)} L ${point.x.toFixed(2)} ${arrowHeadY.toFixed(2)} L ${(point.x + 6).toFixed(2)} ${(arrowHeadY - 7).toFixed(2)} Z"></path>
              </g>
            `;
          }

          if (point.bad && objectiveMode === "max") {
            return `
              <g transform="translate(${point.x.toFixed(2)} ${point.y.toFixed(2)})" class="analyse-arrow-marker analyse-arrow-marker-up">
                <line x1="0" y1="7" x2="0" y2="-3"></line>
                <path d="M -5 -1 L 0 -8 L 5 -1 Z"></path>
              </g>
            `;
          }

          const radius = index === coords.length - 1 ? 5.8 : 4.4;
          const pointClass = index === coords.length - 1 ? "analyse-point-good analyse-point-current" : "analyse-point-good";
          return `<circle cx="${point.x.toFixed(2)}" cy="${point.y.toFixed(2)}" r="${radius.toFixed(1)}" class="${pointClass}"></circle>`;
        }
      )
      .join("");

    const badLegend = objectiveMode === "min"
      ? `<span><i class="legend-arrow legend-arrow-down"></i> ${isMonthly ? "Mois sous objectif" : "Jour sous objectif"}</span>`
      : `<span><i class="legend-arrow legend-arrow-up"></i> ${isMonthly ? "Mois au-dessus du max" : "Jour au-dessus du max"}</span>`;

    return `
      <h3>${chartTitle}</h3>
      <p class="hint-text">${escapeHtml(subtitle)}</p>
      <div class="analyse-chart-shell">
        <svg viewBox="0 0 ${width} ${height}" role="img" aria-label="${isMonthly ? "Analyse mensuelle" : "Analyse journaliere"}">
          <g>
            ${yGuides}
            <line x1="${padLeft}" y1="${objectiveY.toFixed(2)}" x2="${(width - padRight).toFixed(2)}" y2="${objectiveY.toFixed(2)}" class="analyse-objective-line"></line>
            <polyline points="${polyline}" class="analyse-main-line"></polyline>
            ${pointDots}
            ${xLabels}
          </g>
        </svg>
      </div>
      <div class="analyse-legend">
        <span><i class="legend-dot legend-good"></i> ${isMonthly ? "Mois conforme" : "Jour conforme"}</span>
        ${badLegend}
        <span><i class="legend-line"></i> Objectif ${objective.toFixed(2)}%</span>
      </div>
    `;
  }

  function buildAnalyseObjectiveChart(details, objective, labelMode) {
    const isMonthly = labelMode === "month";
    const chartTitle = isMonthly ? "Synthese mensuelle %Ecart" : "Synthese journaliere %Ecart";
    const points = (Array.isArray(details) ? details : [])
      .map((detail) => ({
        key: detail.key,
        value: Number(detail.ecartObjectif || 0),
        bad: Number(detail.ecartObjectif || 0) > 0
      }))
      .filter((point) => point.key);

    if (!points.length) {
      return `<h3>${chartTitle}</h3><p class="hint-text">%Ecart = 99,55% - Taux corrige.</p><p class="hint-text">Aucune donnee disponible.</p>`;
    }

    const width = 920;
    const height = 320;
    const padLeft = 56;
    const padRight = 18;
    const padTop = 22;
    const padBottom = 52;
    const plotWidth = width - padLeft - padRight;
    const plotHeight = height - padTop - padBottom;
    const stepX = points.length > 1 ? plotWidth / (points.length - 1) : 0;

    const values = points.map((point) => point.value);
    const minValue = Math.min(...values, 0);
    const maxValue = Math.max(...values, 0);
    const rawRange = maxValue - minValue;
    const visualPad = Math.max(0.08, rawRange * 0.18);
    let axisMin = minValue - visualPad;
    let axisMax = maxValue + visualPad;
    if (axisMax - axisMin < 0.35) {
      axisMin -= 0.18;
      axisMax += 0.18;
    }

    const roughStep = (axisMax - axisMin) / 4;
    const stepMag = 10 ** Math.floor(Math.log10(Math.max(Math.abs(roughStep), 0.01)));
    const normalized = roughStep / stepMag;
    const niceBase = normalized <= 1 ? 1 : normalized <= 2 ? 2 : normalized <= 5 ? 5 : 10;
    const tickStep = Math.max(0.05, niceBase * stepMag);
    axisMin = Math.floor(axisMin / tickStep) * tickStep;
    axisMax = Math.ceil(axisMax / tickStep) * tickStep;
    if (axisMin === axisMax) axisMax = axisMin + tickStep;

    const yFromDelta = (value) => padTop + ((axisMax - value) / (axisMax - axisMin)) * plotHeight;
    const zeroY = yFromDelta(0);

    const coords = points.map((point, index) => ({
      ...point,
      x: padLeft + (points.length > 1 ? stepX * index : plotWidth / 2),
      y: yFromDelta(point.value)
    }));

    const linePoints = coords.map((point) => `${point.x.toFixed(2)},${point.y.toFixed(2)}`).join(" ");
    const areaPath = [
      `M ${coords[0].x.toFixed(2)} ${zeroY.toFixed(2)}`,
      ...coords.map((point) => `L ${point.x.toFixed(2)} ${point.y.toFixed(2)}`),
      `L ${coords[coords.length - 1].x.toFixed(2)} ${zeroY.toFixed(2)}`,
      "Z"
    ].join(" ");

    const ticks = [];
    for (let tick = axisMin; tick <= axisMax + tickStep / 2; tick += tickStep) {
      ticks.push(Number(tick.toFixed(2)));
    }
    if (!ticks.some((tick) => Math.abs(tick) < 0.001)) {
      ticks.push(0);
      ticks.sort((a, b) => a - b);
    }

    const yGuides = ticks
      .map((tick) => {
        const y = yFromDelta(tick);
        const css = Math.abs(tick) < 0.001 ? "analyse-zero-line" : "analyse-grid-line";
        return `<line x1="${padLeft}" y1="${y.toFixed(2)}" x2="${(width - padRight).toFixed(2)}" y2="${y.toFixed(2)}" class="${css}"></line><text x="${padLeft - 10}" y="${(y + 4).toFixed(2)}" class="analyse-axis-text">${tick.toFixed(1).replace(".", ",")}%</text>`;
      })
      .join("");

    const xLabelStep = Math.max(1, Math.ceil(coords.length / 8));
    const xLabels = coords
      .filter((_, index) => index % xLabelStep === 0 || index === coords.length - 1)
      .map(
        (point) => `<text x="${point.x.toFixed(2)}" y="${(height - 22).toFixed(2)}" text-anchor="middle" class="analyse-axis-text">${escapeHtml(isMonthly ? formatMonthLabel(point.key) : formatDayLabel(point.key))}</text>`
      )
      .join("");

    const pointLabelStep = coords.length <= 14 ? 1 : coords.length <= 22 ? 2 : 3;
    const pointLabels = coords
      .filter((point, index) => point.bad || index % pointLabelStep === 0 || index === coords.length - 1)
      .map((point) => {
        const direction = point.value > 0 ? -10 : 16;
        return `<text x="${point.x.toFixed(2)}" y="${(point.y + direction).toFixed(2)}" text-anchor="middle" class="analyse-point-text">${Number(point.value).toFixed(2).replace(".", ",")} %</text>`;
      })
      .join("");

    const markers = coords
      .map((point, index) => {
        const radius = index === coords.length - 1 ? 5.8 : 4.4;
        const isSelected = state.analyseSelectedKey === point.key;
        let pointClass = point.bad ? "analyse-point-bad" : "analyse-point-good";
        if (index === coords.length - 1 && !point.bad) {
          pointClass = "analyse-point-good analyse-point-current";
        }
        if (isSelected) {
          pointClass += " analyse-point-selected";
        }
        pointClass += " analyse-point-interactive";
        return `<circle cx="${point.x.toFixed(2)}" cy="${point.y.toFixed(2)}" r="${radius.toFixed(1)}" class="${pointClass}" data-analyse-point="1" data-key="${escapeHtml(point.key)}" data-view="${escapeHtml(labelMode)}"></circle>`;
      })
      .join("");

    return `
      <h3>${chartTitle}</h3>
      <div class="analyse-chart-shell">
        <svg viewBox="0 0 ${width} ${height}" role="img" aria-label="${isMonthly ? "Synthese mensuelle vs objectif" : "Synthese journaliere vs objectif"}">
          <g>
            ${yGuides}
            <path d="${areaPath}" class="analyse-area-fill"></path>
            <polyline points="${linePoints}" class="analyse-delta-line"></polyline>
            ${markers}
            ${pointLabels}
            ${xLabels}
          </g>
        </svg>
      </div>
      <div class="analyse-legend">
        <span><i class="legend-dot legend-good"></i> Au-dessus de 99,55%</span>
        <span><i class="legend-dot legend-bad"></i> Sous 99,55%</span>
        <span><i class="legend-line"></i> Ligne 0 = objectif 99,55%</span>
      </div>
    `;
  }

  function renderAnalyse() {
    if (!analyseSummary || !analyseChartWrap || !analyseTableWrap || !analyseTabs.length || !state.analyseStats) return;

    const stats = state.analyseStats;
    const objectifFiabilite = ANALYSE_OBJECTIVE;
    const objectifEcarts = 100 - objectifFiabilite;
    const isMonthly = state.analyseView === "month";
    const details = filterAnalyseDetails(isMonthly ? stats.monthlyDetails : stats.dailyDetails);
    if (state.analyseSelectedKey && !details.some((detail) => detail.key === state.analyseSelectedKey)) {
      state.analyseSelectedKey = "";
    }
    const tableDetails = state.analyseSelectedKey
      ? details.filter((detail) => detail.key === state.analyseSelectedKey)
      : details;
    const summary = summarizeAnalyseDetails(details);
    const tableSummary = summarizeAnalyseDetails(tableDetails);
    const ecartPoints = filterAnalysePoints(isMonthly ? stats.monthlyEcarts : stats.dailyEcarts);
    const objectifDelta = summary.tauxFiabilite - ANALYSE_OBJECTIVE;
    const ecartVsObjectif = summary.ecartObjectif;
    const fiabiliteClass = objectifDelta < 0 ? "is-down" : "is-up";
    const ecartClass = ecartVsObjectif > 0 ? "is-down" : "is-up";

    analyseTabs.forEach((tab) => {
      tab.classList.toggle("is-active", tab.dataset.analyseType === state.analyseType);
    });
    renderAnalyseToolbar();

    if (state.analyseType === "ecarts") {
      analyseSummary.innerHTML = `
        <article class="card-block analyse-summary-card analyse-kpi-card">
          <div class="analyse-kpi-head">
            <h3>Taux des ecarts</h3>
            ${buildObjectiveBadgeSafe(objectifEcarts - summary.tauxEcarts)}
          </div>
          <p class="analyse-metric-value ${summary.tauxEcarts > objectifEcarts ? "is-down" : "is-up"}"><strong>${summary.tauxEcarts.toFixed(2)}%</strong></p>
        </article>
        <article class="card-block analyse-summary-card">
          <h3>Total UVC Controle</h3>
          <p class="analyse-metric-value"><strong>${summary.totalUvcControle}</strong></p>
        </article>
        <article class="card-block analyse-summary-card">
          <h3>Total UVC Ecart</h3>
          <p class="analyse-metric-value"><strong>${summary.totalUvcEcart}</strong></p>
        </article>
      `;
      analyseChartWrap.innerHTML = buildAnalyseChart(
        ecartPoints,
        objectifEcarts,
        "max",
        "Taux des ecarts = SUM(UVC ECART filtre Manquant/Surplus) / SUM(UVC CONTROLE).",
        isMonthly ? "month" : "day"
      );
      analyseTableWrap.innerHTML = buildAnalyseTable(
        tableDetails,
        tableSummary,
        isMonthly ? "month" : "day",
        state.analyseSelectedKey
      );
      return;
    }

    analyseSummary.innerHTML = `
      <article class="card-block analyse-summary-card">
        <h3>Total UVC Controle</h3>
        <p class="analyse-metric-value"><strong>${summary.totalUvcControle}</strong></p>
      </article>
      <article class="card-block analyse-summary-card">
        <h3>UVCLIV</h3>
        <p class="analyse-metric-value"><strong>${summary.totalUvcLivre}</strong></p>
      </article>
      <article class="card-block analyse-summary-card analyse-kpi-card">
        <div class="analyse-kpi-head">
          <h3>Taux Corrige</h3>
          ${buildObjectiveBadgeSafe(objectifDelta)}
        </div>
        <p class="analyse-metric-value ${fiabiliteClass}"><strong>${summary.tauxFiabilite.toFixed(2)}%</strong></p>
      </article>
      <article class="card-block analyse-summary-card analyse-kpi-card">
        <div class="analyse-kpi-head">
          <h3>%Ecart%</h3>
          ${buildObjectiveBadgeSafe(-ecartVsObjectif)}
        </div>
        <p class="analyse-metric-value ${ecartClass}"><strong>${formatAnalysePercent(ecartVsObjectif)}</strong></p>
      </article>
    `;
    analyseChartWrap.innerHTML = buildAnalyseObjectiveChart(
      details,
      objectifFiabilite,
      isMonthly ? "month" : "day"
    );
    analyseTableWrap.innerHTML = buildAnalyseTable(
      tableDetails,
      tableSummary,
      isMonthly ? "month" : "day",
      state.analyseSelectedKey
    );
  }

  async function loadAnalyse(force = false) {
    if (!isAdmin || !analyseSummary || !analyseChartWrap || !analyseTableWrap) return;
    if (!force && state.analyseStats) {
      renderAnalyse();
      return;
    }
    setFeedback(analyseFeedback, "Loading analyse...", "");
    try {
      const payload = await api(`/api/admin/analyse${analyseQuery()}`);
      state.analyseStats = normalizeAnalysePayload(payload);
      if (state.analyseMonth && !state.analyseStats.availableMonths.includes(state.analyseMonth)) {
        state.analyseMonth = "";
      }
      state.analyseSelectedKey = "";
      renderAnalyse();
      setFeedback(
        analyseFeedback,
        `Analyse chargee (${state.analyseStats.sourceType || "source"}: ${state.analyseStats.sourceLabel || "-"})`,
        "ok"
      );
    } catch (error) {
      setFeedback(analyseFeedback, error.message, "error");
    }
  }

  navItems.forEach((item) => {
    item.addEventListener("click", () => {
      const section = item.dataset.section;
      if (!canAccessSection(section)) return;
      showSection(section);
      if (section === "library") {
        showLibraryScreen("root");
        loadLibraries();
      }
      if (section === "users") loadUserManagement();
      if (section === "analyse") loadAnalyse();
      if (section === "donnes") renderDonnes();
      if (section === "settings") loadSettingsAccess();
    });
  });

  form.addEventListener("submit", (e) => {
    e.preventDefault();
    loadHomeMovements();
  });
  resetBtn.addEventListener("click", () => {
    form.reset();
    if (!isAdmin) usernameFilter.value = currentUsername;
    loadHomeMovements();
  });

  tableBody.addEventListener("click", async (e) => {
    const row = e.target.closest("tr[data-movement-id]");
    if (!row) return;
    try {
      await openMovementDetails(row.dataset.movementId);
    } catch (error) {
      setFeedback(homeFeedback, error.message, "error");
    }
  });

  closeDrawerBtn.addEventListener("click", closeDrawer);
  drawer.addEventListener("click", (e) => {
    if (e.target === drawer) closeDrawer();
  });
  deleteMovementBtn.addEventListener("click", async () => {
    if (!state.selectedMovementId || !window.confirm("Supprimer cette data ?")) return;
    try {
      await deleteMovement(state.selectedMovementId);
      closeDrawer();
      setFeedback(homeFeedback, "Data supprimee.", "ok");
      loadHomeMovements();
    } catch (error) {
      setFeedback(homeFeedback, error.message, "error");
    }
  });

  if (isAdmin && libraryCreateForm) {
    libraryCreateForm.addEventListener("submit", async (e) => {
      e.preventDefault();
      const name = libraryNameInput.value.trim();
      if (!name) return;
      try {
        await api("/api/admin/libraries", { method: "POST", headers: { "X-CSRF-Token": csrfToken }, body: { name } });
        libraryNameInput.value = "";
        setFeedback(libraryFeedback, "Library creee.", "ok");
        await loadLibraries();
        await loadSettingsAccess(true);
      } catch (error) {
        setFeedback(libraryFeedback, error.message, "error");
      }
    });
  }

  if (isAdmin && usersCreateRole) {
    usersCreateRole.addEventListener("change", () => renderCreatePermissions());
  }

  if (isAdmin && usersCreateForm) {
    usersCreateForm.addEventListener("submit", async (e) => {
      e.preventDefault();
      const username = usersCreateUsername?.value.trim() || "";
      const password = usersCreatePassword?.value || "";
      const role = usersCreateRole?.value || "controller";
      const permissions = selectedCreatePermissions();
      if (!permissions.length) {
        setFeedback(usersCreateFeedback, "Selectionne au moins une permission.", "error");
        return;
      }
      try {
        const result = await api("/api/admin/user-management/users", {
          method: "POST",
          headers: { "X-CSRF-Token": csrfToken },
          body: { username, password, role, permissions }
        });
        usersCreateForm.reset();
        if (usersCreateRole) usersCreateRole.value = "controller";
        renderCreatePermissions();
        setFeedback(usersCreateFeedback, `User ${result.user?.username || username} cree.`, "ok");
        await loadUserManagement(true);
        await loadUsersFilter();
        await loadSettingsAccess(true);
      } catch (error) {
        setFeedback(usersCreateFeedback, error.message, "error");
      }
    });
  }

  if (isAdmin && usersAdminList) {
    usersAdminList.addEventListener("click", async (e) => {
      const toggleBtn = e.target.closest("[data-toggle-user-id]");
      if (toggleBtn) {
        const userId = Number(toggleBtn.dataset.toggleUserId);
        state.expandedUserId = Number(state.expandedUserId) === userId ? null : userId;
        renderUserManagementList();
        return;
      }

      const saveBtn = e.target.closest("[data-save-user-id]");
      if (saveBtn) {
        const userId = Number(saveBtn.dataset.saveUserId);
        const card = saveBtn.closest("[data-user-card]");
        if (!card) return;
        const role = card.querySelector("[data-user-role]")?.value || "controller";
        const isBlocked = Boolean(card.querySelector("[data-user-blocked]")?.checked);
        const permissions = selectedUserPermissions(card);
        if (!permissions.length) {
          setFeedback(usersAdminFeedback, "Selectionne au moins une permission.", "error");
          return;
        }
        try {
          await api(`/api/admin/user-management/users/${userId}`, {
            method: "PUT",
            headers: { "X-CSRF-Token": csrfToken },
            body: {
              role,
              is_blocked: isBlocked,
              permissions
            }
          });
          setFeedback(usersAdminFeedback, "User mis a jour.", "ok");
          state.expandedUserId = userId;
          await loadUserManagement(true);
          await loadUsersFilter();
          await loadSettingsAccess(true);
          await loadLibraries();
        } catch (error) {
          setFeedback(usersAdminFeedback, error.message, "error");
        }
        return;
      }

      const resetBtn = e.target.closest("[data-reset-user-id]");
      if (resetBtn) {
        const userId = Number(resetBtn.dataset.resetUserId);
        try {
          const result = await api(`/api/admin/user-management/users/${userId}/reset-password`, {
            method: "POST",
            headers: { "X-CSRF-Token": csrfToken }
          });
          setFeedback(usersAdminFeedback, `Password reset pour ${result.username}.`, "ok");
          window.alert(`Temporary password for ${result.username}: ${result.temporary_password}`);
        } catch (error) {
          setFeedback(usersAdminFeedback, error.message, "error");
        }
        return;
      }

      const deleteBtn = e.target.closest("[data-delete-user-id]");
      if (!deleteBtn) return;
      const userId = Number(deleteBtn.dataset.deleteUserId);
      const card = deleteBtn.closest("[data-user-card]");
      const username = card?.querySelector("strong")?.textContent?.trim() || `#${userId}`;
      if (!window.confirm(`Supprimer user "${username}" ?`)) return;
      try {
        const result = await api(`/api/admin/user-management/users/${userId}`, {
          method: "DELETE",
          headers: { "X-CSRF-Token": csrfToken }
        });
        const deletedMovements = Number(result.deleted_movements || 0);
        const deletedLinks = Number(result.deleted_library_links || 0);
        setFeedback(
          usersAdminFeedback,
          `User ${username} supprime definitivement. ${deletedMovements} mouvement(s) et ${deletedLinks} acces library supprimes.`,
          "ok"
        );
        await loadUserManagement(true);
        await loadUsersFilter();
        await loadSettingsAccess(true);
        await loadLibraries();
      } catch (error) {
        setFeedback(usersAdminFeedback, error.message, "error");
      }
    });
  }

  libraryList.addEventListener("click", async (e) => {
    const deleteBtn = e.target.closest("[data-delete-library-id]");
    if (deleteBtn) {
      if (!isAdmin) return;
      const libraryId = Number(deleteBtn.dataset.deleteLibraryId);
      const library = state.libraries.find((item) => Number(item.id) === libraryId);
      const libraryName = library?.name || `#${libraryId}`;
      if (!window.confirm(`Supprimer la library "${libraryName}" ?`)) return;
      try {
        await deleteLibrary(libraryId);
        if (state.selectedLibraryId === libraryId) {
          state.selectedLibraryId = null;
          state.selectedLibraryName = "";
          state.selectedLibraryUser = null;
          state.selectedMonth = null;
          showLibraryScreen("root");
        }
        setFeedback(libraryFeedback, `Library "${libraryName}" supprimee.`, "ok");
        await loadLibraries();
        await loadSettingsAccess(true);
      } catch (error) {
        setFeedback(libraryFeedback, error.message, "error");
      }
      return;
    }

    const btn = e.target.closest("[data-open-library-id]");
    if (!btn) return;
    const libraryId = Number(btn.dataset.openLibraryId);
    try {
      if (isAdmin) await openLibraryAdmin(libraryId);
      else openLibraryUserMode(libraryId);
    } catch (error) {
      setFeedback(libraryFeedback, error.message, "error");
    }
  });

  usersGrid.addEventListener("click", async (e) => {
    const btn = e.target.closest("[data-open-user-id]");
    if (!btn) return;
    try {
      await openLibraryUser(btn.dataset.openUserId);
    } catch (error) {
      setFeedback(libraryFeedback, error.message, "error");
    }
  });

  monthsGrid.addEventListener("click", (e) => {
    const btn = e.target.closest("[data-open-month]");
    if (!btn) return;
    renderMonthDays(btn.dataset.openMonth);
  });

  daysWrap.addEventListener("click", async (e) => {
    const openBtn = e.target.closest("[data-open-movement-id]");
    if (openBtn) {
      try {
        await openMovementDetails(openBtn.dataset.openMovementId);
      } catch (error) {
        setFeedback(libraryFeedback, error.message, "error");
      }
      return;
    }
    const delBtn = e.target.closest("[data-delete-movement-id]");
    if (!delBtn || !window.confirm("Supprimer cette data ?")) return;
    try {
      const previousMonthKey = state.selectedMonth?.key || "";
      await deleteMovement(delBtn.dataset.deleteMovementId);
      setFeedback(libraryFeedback, "Data supprimee.", "ok");
      loadHomeMovements();
      if (state.selectedLibraryUser) {
        await openLibraryUser(state.selectedLibraryUser.id);
        if (previousMonthKey) renderMonthDays(previousMonthKey);
      }
    } catch (error) {
      setFeedback(libraryFeedback, error.message, "error");
    }
  });

  if (libraryForm) {
    libraryForm.addEventListener("submit", async (e) => {
      e.preventDefault();
      if (isAdmin || !state.selectedLibraryId) return;
      const payload = {
        support_number: supportInput.value.trim(),
        ean_code: eanInput.value.trim(),
        product_code: productInput.value.trim(),
        diff_plus: plusInput.value,
        diff_minus: minusInput.value
      };
      try {
        setFeedback(libraryFormFeedback, "Saving...", "");
        const result = await saveLibraryData(payload);
        setFeedback(libraryFormFeedback, "Data saved.", "ok");
        libraryForm.reset();
        plusInput.value = "0";
        minusInput.value = "0";
        await stopScanner();
        libraryLastRecord.classList.remove("empty");
        libraryLastRecord.innerHTML = `<strong>ID #${result.movement_id}</strong><br>User: ${escapeHtml(result.username)}<br>Library: ${escapeHtml(state.selectedLibraryName)}<br>Date: ${escapeHtml(formatDate(result.movement_date))}`;
        loadHomeMovements();
      } catch (error) {
        setFeedback(libraryFormFeedback, error.message, "error");
      }
    });
  }

  scanToggle.addEventListener("click", () => {
    if (isAdmin) return;
    if (state.scannerRunning) stopScanner();
    else startScanner();
  });

  if (backRoot) backRoot.addEventListener("click", () => showLibraryScreen("root"));
  if (backUsers) backUsers.addEventListener("click", () => showLibraryScreen("users"));
  if (backMonths) backMonths.addEventListener("click", () => showLibraryScreen("months"));
  if (backEntry) backEntry.addEventListener("click", () => showLibraryScreen("root"));

  if (isAdmin && settingsLibrarySelect) {
    settingsLibrarySelect.addEventListener("change", () => {
      state.settingsLibraryId = Number(settingsLibrarySelect.value || "0") || null;
      renderSettingsUsers();
      setFeedback(settingsAccessFeedback, "", "");
    });
  }
  if (isAdmin && settingsSaveAccessBtn) {
    settingsSaveAccessBtn.addEventListener("click", () => saveSettingsAccess());
  }
  analyseTabs.forEach((tab) => {
    tab.addEventListener("click", () => {
      state.analyseType = tab.dataset.analyseType === "ecarts" ? "ecarts" : "fiabilite";
      state.analyseSelectedKey = "";
      renderAnalyse();
    });
  });
  analyseViewButtons.forEach((button) => {
    button.addEventListener("click", () => {
      state.analyseView = button.dataset.analyseView === "month" ? "month" : "day";
      state.analyseSelectedKey = "";
      renderAnalyse();
    });
  });
  if (analyseMonthFilter) {
    analyseMonthFilter.addEventListener("change", () => {
      state.analyseMonth = analyseMonthFilter.value || "";
      state.analyseSelectedKey = "";
      renderAnalyse();
    });
  }
  if (analyseChartWrap) {
    analyseChartWrap.addEventListener("click", (e) => {
      const point = e.target.closest("[data-analyse-point]");
      if (!point) return;
      const key = String(point.dataset.key || "");
      if (!key) return;
      state.analyseSelectedKey = state.analyseSelectedKey === key ? "" : key;
      renderAnalyse();
    });
  }
  if (analyseRefreshBtn) {
    analyseRefreshBtn.addEventListener("click", () => loadAnalyse(true));
  }
  if (donnesExtractionChoices) {
    donnesExtractionChoices.addEventListener("click", (e) => {
      const btn = e.target.closest("[data-donnes-item]");
      if (!btn) return;
      const itemId = String(btn.dataset.donnesItem || "").toLowerCase();
      const itemLabel = String(btn.dataset.donnesLabel || itemId);
      triggerDonnesExtraction(itemId, itemLabel);
    });
  }
  if (donnesConnectionChoices) {
    donnesConnectionChoices.addEventListener("click", async (e) => {
      const btn = e.target.closest("[data-donnes-item]");
      if (!btn) return;
      const itemId = String(btn.dataset.donnesItem || "").toLowerCase();
      if (itemId === "odbc") {
        await launchOdbcAdministrator();
        return;
      }
      await pickDonnesSourceFile(itemId);
    });
  }
  if (donnesConnectionForm && donnesConnectionInput) {
    donnesConnectionInput.addEventListener("input", () => {
      if (state.donnesConnectionType) {
        state.donnesConnectionValues[state.donnesConnectionType] = donnesConnectionInput.value;
      }
    });
    donnesConnectionForm.addEventListener("submit", async (e) => {
      e.preventDefault();
      await submitDonnesConnectionTest();
    });
  }
  if (donnesConnectionHelper) {
    donnesConnectionHelper.addEventListener("click", async () => {
      if (state.donnesConnectionType === "odbc") {
        await launchOdbcAdministrator();
        return;
      }
      if (state.donnesConnectionType === "excel" || state.donnesConnectionType === "access") {
        await pickDonnesSourceFile(state.donnesConnectionType, false);
      }
    });
  }

  themeChoices.forEach((choice) => {
    choice.addEventListener("click", () => applyTheme(choice.dataset.themeChoice || "system"));
  });
  if (systemThemeMedia) {
    const syncSystemTheme = () => {
      if ((themeSelect?.value || "system") === "system") {
        applyTheme("system", { persist: false });
      }
    };
    if (typeof systemThemeMedia.addEventListener === "function") {
      systemThemeMedia.addEventListener("change", syncSystemTheme);
    } else if (typeof systemThemeMedia.addListener === "function") {
      systemThemeMedia.addListener(syncSystemTheme);
    }
  }
  window.addEventListener("beforeunload", () => stopScanner());

  try {
    applyTheme(window.localStorage.getItem("admin_theme") || "system");
  } catch (_e) {
    applyTheme("system");
  }

  applySectionPermissions();
  showSection(firstAllowedSection());
  showLibraryScreen("root");
  setFormEnabled(!isAdmin, "");
  if (canAccessSection("home")) {
    loadUsersFilter().then(loadHomeMovements).catch((error) => setFeedback(homeFeedback, error.message, "error"));
  }
  if (canAccessSection("library")) loadLibraries();
  if (isAdmin && canAccessSection("settings")) loadSettingsAccess(true);
  if (isAdmin && canAccessSection("users")) loadUserManagement(true);
  if (isAdmin && canAccessSection("donnes")) renderDonnes();
})();
