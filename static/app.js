(() => {
  const typeNames = ["DATE", "BOOLEAN", "MONEY", "PERCENT", "INTEGER", "NUMBER", "CATEGORY", "IDENTIFIER", "TEXT"];
  const numericTypes = new Set(["MONEY", "PERCENT", "INTEGER", "NUMBER"]);
  const uploadLimitMB = Number(document.querySelector("#upload-limit").textContent.replace(/\D/g, "")) || 100;
  const elements = {
    fileInput: document.querySelector("#file-input"),
    dropzone: document.querySelector("#dropzone"),
    browseButton: document.querySelector("#browse-button"),
    sampleButton: document.querySelector("#sample-button"),
    fileStatus: document.querySelector("#file-status"),
    workspace: document.querySelector("#data-workspace"),
    results: document.querySelector("#results"),
    columns: document.querySelector("#columns-list"),
    count: document.querySelector("#detected-count"),
    previewCount: document.querySelector("#row-count"),
    columnCount: document.querySelector("#column-count"),
    fileName: document.querySelector("#file-name"),
    fileSubtitle: document.querySelector("#file-subtitle"),
    fileIcon: document.querySelector("#file-type-icon"),
    inputTable: document.querySelector("#input-table"),
    resultTable: document.querySelector("#result-table"),
    sheetRow: document.querySelector("#sheet-row"),
    sheetSelect: document.querySelector("#sheet-select"),
    measure: document.querySelector("#measure-select"),
    dimension: document.querySelector("#dimension-select"),
    optionNote: document.querySelector(".option-note"),
    filterList: document.querySelector("#filter-list"),
    validationList: document.querySelector("#validation-list"),
    addFilter: document.querySelector("#add-filter"),
    clearFilters: document.querySelector("#clear-filters"),
    applyFilters: document.querySelector("#apply-filters"),
    addValidation: document.querySelector("#add-validation"),
    filterSummary: document.querySelector("#filter-summary"),
    filterPreview: document.querySelector("#filter-preview"),
    filteredRowCount: document.querySelector("#filtered-row-count"),
    filterTable: document.querySelector("#filter-table"),
    normalizeCategories: document.querySelector("#normalize-categories"),
    normalizeBooleans: document.querySelector("#normalize-booleans"),
    deriveRevenue: document.querySelector("#derive-revenue"),
    deriveRevenueContainer: document.querySelector("#derive-revenue-container"),
    revenueSuggestionLabel: document.querySelector("#revenue-suggestion-label"),
    analyze: document.querySelector("#analyze-button"),
    progress: document.querySelector("#analysis-progress"),
    toast: document.querySelector("#toast"),
    themeToggle: document.querySelector("#theme-toggle"),
  };
  let sessionId = null;
  let filename = "";
  let ambiguousDates = {};
  let activeToast;
  let availableColumns = [];
  const hero = document.querySelector(".hero");
  document.querySelector(".hero-secondary")?.remove();

  function setHeroMotion(event) {
    if (!hero || window.matchMedia("(prefers-reduced-motion: reduce)").matches) return;
    const bounds = hero.getBoundingClientRect();
    const x = ((event.clientX - bounds.left) / bounds.width - 0.5) * 2;
    const y = ((event.clientY - bounds.top) / bounds.height - 0.5) * 2;
    hero.style.setProperty("--pointer-x", `${(x * 8).toFixed(2)}px`);
    hero.style.setProperty("--pointer-y", `${(y * 6).toFixed(2)}px`);
  }

  function showToast(message, error = false) {
    elements.toast.textContent = message;
    elements.toast.classList.toggle("error", error);
    elements.toast.classList.add("visible");
    clearTimeout(activeToast);
    activeToast = setTimeout(() => elements.toast.classList.remove("visible"), 4200);
  }

  async function requestJson(url, options) {
    const response = await fetch(url, options);
    const result = await response.json().catch(() => ({}));
    if (!response.ok) throw new Error(result.error || `Request failed (${response.status}).`);
    return result;
  }

  function setFileStatus(message, error = false) {
    elements.fileStatus.textContent = message;
    elements.fileStatus.classList.toggle("error-text", error);
  }

  function validateFile(file) {
    if (!file) return false;
    if (!/\.(csv|tsv|tab|txt|xlsx|xlsm|xls)$/i.test(file.name)) {
      showToast("Choose a CSV, Excel, TSV, or text file.", true);
      return false;
    }
    if (file.size > uploadLimitMB * 1024 * 1024) {
      showToast(`This file is larger than the ${uploadLimitMB} MB upload limit.`, true);
      return false;
    }
    return true;
  }

  function selectFile(file) {
    if (validateFile(file)) {
      setFileStatus(`Ready to inspect ${file.name}`);
      uploadFile(file);
    }
  }

  async function uploadFile(file) {
    const bytes = new Uint8Array(await file.arrayBuffer());
    let binary = "";
    const chunkSize = 0x8000;
    for (let index = 0; index < bytes.length; index += chunkSize) {
      binary += String.fromCharCode(...bytes.subarray(index, index + chunkSize));
    }
    await loadDataset(
      "/api/upload-base64",
      {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ filename: file.name, content: btoa(binary) }),
      },
      `Uploading ${file.name}…`,
    );
  }

  async function loadSample() {
    await loadDataset("/api/sample", { method: "POST" }, "Loading the included sample dataset…");
  }

  async function loadDataset(url, options, statusMessage) {
    sessionId = null;
    elements.workspace.classList.add("hidden");
    elements.results.classList.add("hidden");
    setFileStatus(statusMessage);
    elements.dropzone.classList.add("busy");
    elements.sampleButton.disabled = true;
    try {
      const data = await requestJson(url, options);
      sessionId = data.session_id;
      filename = data.filename;
      showWorkspace(data);
      setFileStatus("");
    } catch (error) {
      setFileStatus(error.message, true);
      showToast(error.message, true);
    } finally {
      elements.dropzone.classList.remove("busy");
      elements.sampleButton.disabled = false;
      elements.fileInput.value = "";
    }
  }

  function setCurrentStep(step) {
    const steps = ["upload", "review", "results"];
    document.querySelectorAll(".step-link").forEach((link) => {
      const index = steps.indexOf(link.dataset.step);
      const activeIndex = steps.indexOf(step);
      link.classList.toggle("active", index === activeIndex);
      link.classList.toggle("complete", index < activeIndex);
    });
  }

  function showWorkspace(data) {
    elements.workspace.classList.remove("hidden");
    elements.results.classList.add("hidden");
    elements.fileName.textContent = filename;
    elements.fileSubtitle.textContent = "Ready for review";
    elements.fileIcon.textContent = (filename.split(".").pop() || "FILE").slice(0, 4).toUpperCase();
    elements.previewCount.textContent = Number(data.rows).toLocaleString();
    elements.columnCount.textContent = Number(data.column_count).toLocaleString();
    elements.count.textContent = `${data.columns.length} ${data.columns.length === 1 ? "field" : "fields"} found`;
    availableColumns = data.columns;
    elements.filterList?.replaceChildren();
    elements.validationList?.replaceChildren();
    elements.filterPreview?.classList.add("hidden");
    if (elements.filterSummary) elements.filterSummary.textContent = "Showing all rows";
    ambiguousDates = {};
    for (const column of data.columns) {
      if (column.date_order_ambiguous) ambiguousDates[column.name] = true;
    }
    renderColumns(data.columns);
    renderTable(elements.inputTable, data.preview);
    const hasSheets = Array.isArray(data.sheets) && data.sheets.length > 0;
    elements.sheetRow.classList.toggle("hidden", !hasSheets);
    if (hasSheets) {
      elements.sheetSelect.replaceChildren();
      for (const sheet of data.sheets) elements.sheetSelect.append(new Option(sheet, sheet));
    }
    const suggestion = data.revenue_suggestion;
    elements.deriveRevenue.checked = false;
    elements.deriveRevenueContainer.classList.toggle("hidden", !suggestion);
    if (suggestion) {
      const formula = `${suggestion.quantity} × ${suggestion.price}`;
      const discount = suggestion.discount ? ` × (1 − ${suggestion.discount})` : "";
      elements.revenueSuggestionLabel.textContent = `Add optional computed revenue (${formula}${discount})`;
    }
    refreshFocusOptions();
    setCurrentStep("review");
    elements.workspace.scrollIntoView({ behavior: "smooth", block: "start" });
    document.querySelector("#workspace-title").focus({ preventScroll: true });
  }

  function renderColumns(columns) {
    elements.columns.replaceChildren();
    for (const column of columns) {
      const row = document.createElement("tr");
      const nameCell = document.createElement("td");
      nameCell.textContent = column.name;
      nameCell.title = column.name;
      const typeCell = document.createElement("td");
      const select = document.createElement("select");
      select.className = "type-select";
      select.dataset.column = column.name;
      select.setAttribute("aria-label", `Detected type for ${column.name}`);
      for (const type of typeNames) {
        const option = new Option(type.charAt(0) + type.slice(1).toLowerCase(), type);
        option.selected = type === column.type;
        select.append(option);
      }
      select.addEventListener("change", refreshFocusOptions);
      typeCell.append(select);
      const dateCell = document.createElement("td");
      const dateSelect = document.createElement("select");
      dateSelect.className = "type-select date-order-select";
      dateSelect.dataset.dateColumn = column.name;
      dateSelect.setAttribute("aria-label", `Date order for ${column.name}`);
      dateSelect.append(new Option("Automatic", ""));
      dateSelect.append(new Option("Day / month / year", "day_first"));
      dateSelect.append(new Option("Month / day / year", "month_first"));
      dateSelect.disabled = column.type !== "DATE";
      if (column.date_order_ambiguous && column.type === "DATE") dateSelect.classList.add("date-ambiguous");
      const dateNote = document.createElement("span");
      dateNote.className = "date-ambiguous";
      dateNote.textContent = column.date_order_ambiguous ? "Check order" : "";
      dateCell.append(dateSelect);
      if (column.date_order_ambiguous) dateCell.append(dateNote);
      row.append(nameCell, typeCell, dateCell);
      elements.columns.append(row);
    }
  }

  function ruleOperators(kind, type) {
    const numeric = numericTypes.has(type);
    if (kind === "filter") {
      return numeric
        ? [["equals", "is"], ["not_equals", "is not"], ["greater_than", "greater than"], ["less_than", "less than"], ["is_empty", "is empty"], ["is_not_empty", "is not empty"]]
        : [["equals", "is"], ["not_equals", "is not"], ["contains", "contains"], ["is_empty", "is empty"], ["is_not_empty", "is not empty"]];
    }
    return numeric
      ? [["required", "is required"], ["min", "minimum value"], ["max", "maximum value"]]
      : [["required", "is required"], ["equals", "must equal"], ["not_equals", "must not equal"], ["contains", "must contain"], ["one_of", "must be one of"]];
  }

  function ruleRow(kind, initial = {}) {
    const row = document.createElement("div");
    row.className = "rule-row";
    const column = document.createElement("select");
    column.className = "rule-column";
    column.setAttribute("aria-label", kind === "filter" ? "Filter column" : "Validation column");
    for (const item of availableColumns) column.append(new Option(item.name, item.name));
    column.value = availableColumns.some((item) => item.name === initial.column) ? initial.column : (availableColumns[0]?.name || "");
    const operator = document.createElement("select");
    operator.className = "rule-operator";
    operator.setAttribute("aria-label", kind === "filter" ? "Filter condition" : "Validation condition");
    const value = document.createElement("input");
    value.className = "rule-value";
    value.type = "text";
    value.value = initial.value || "";
    const remove = document.createElement("button");
    remove.type = "button";
    remove.className = "rule-remove";
    remove.setAttribute("aria-label", "Remove rule");
    remove.textContent = "×";
    const refresh = () => {
      const type = availableColumns.find((item) => item.name === column.value)?.type || "TEXT";
      const choices = ruleOperators(kind, type);
      const selected = choices.some(([key]) => key === operator.value) ? operator.value : (choices[0]?.[0] || "");
      operator.replaceChildren(...choices.map(([key, label]) => new Option(label, key)));
      operator.value = selected;
      const valueRequired = !["required", "is_empty", "is_not_empty"].includes(operator.value);
      value.classList.toggle("hidden", !valueRequired);
      value.placeholder = operator.value === "one_of" ? "Comma-separated values" : operator.value === "min" || operator.value === "max" ? "Enter a number" : "Enter a value";
    };
    column.addEventListener("change", refresh);
    operator.addEventListener("change", refresh);
    remove.addEventListener("click", () => { row.remove(); if (kind === "filter") updateFilterSummary(); });
    row.append(column, operator, value, remove);
    refresh();
    return row;
  }

  function addRule(kind, initial = {}) {
    const list = kind === "filter" ? elements.filterList : elements.validationList;
    list.append(ruleRow(kind, initial));
  }

  function collectRules(kind) {
    const list = kind === "filter" ? elements.filterList : elements.validationList;
    if (!list) return [];
    return [...list.querySelectorAll(".rule-row")].map((row) => ({
      column: row.querySelector(".rule-column").value,
      operator: row.querySelector(".rule-operator").value,
      value: row.querySelector(".rule-value").value,
    }));
  }

  function updateFilterSummary(text = null) {
    if (!elements.filterSummary || !elements.filterList) return;
    if (text) {
      elements.filterSummary.textContent = text;
      return;
    }
    const count = elements.filterList.querySelectorAll(".rule-row").length;
    elements.filterSummary.textContent = count ? `${count} filter${count === 1 ? "" : "s"} ready` : "Showing all rows";
  }

  async function applyFilters() {
    if (!sessionId) return;
    const filters = collectRules("filter");
    if (!filters.length) {
      elements.filterPreview.classList.add("hidden");
      updateFilterSummary("Showing all rows");
      return;
    }
    try {
      const result = await requestJson("/api/explore", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ session_id: sessionId, filters }),
      });
      elements.filterPreview.classList.remove("hidden");
      elements.filteredRowCount.textContent = `${Number(result.filtered_rows).toLocaleString()} of ${Number(result.total_rows).toLocaleString()} rows`;
      updateFilterSummary(`${Number(result.filtered_rows).toLocaleString()} rows match`);
      renderTable(elements.filterTable, result.preview);
    } catch (error) {
      showToast(error.message, true);
    }
  }

  function currentTypes() {
    const result = {};
    for (const select of elements.columns.querySelectorAll(".type-select[data-column]")) {
      result[select.dataset.column] = select.value;
    }
    for (const select of elements.columns.querySelectorAll(".type-select[data-date-column]")) {
      select.disabled = result[select.dataset.dateColumn] !== "DATE";
    }
    return result;
  }

  function dateOrders() {
    const result = {};
    for (const select of elements.columns.querySelectorAll("[data-date-column]")) {
      if (select.value) result[select.dataset.dateColumn] = select.value;
    }
    return result;
  }

  function setOptions(select, columns, label, emptyLabel) {
    const previous = select.value;
    select.replaceChildren(new Option(columns.length ? "Automatic" : emptyLabel, ""));
    for (const column of columns) select.append(new Option(column, column));
    select.value = columns.includes(previous) ? previous : "";
    select.disabled = columns.length === 0;
    select.setAttribute("aria-label", label);
    select.title = columns.length ? label : emptyLabel;
  }

  function refreshFocusOptions() {
    const types = currentTypes();
    const numericColumns = Object.keys(types).filter((column) => numericTypes.has(types[column]));
    const dimensionColumns = Object.keys(types).filter((column) => types[column] === "CATEGORY" || types[column] === "BOOLEAN");
    setOptions(
      elements.measure,
      numericColumns,
      "Numeric measure for the report",
      "No numeric fields",
    );
    setOptions(
      elements.dimension,
      dimensionColumns,
      "Category or Yes/No field to group by",
      "No groupable fields",
    );
    const hasCategories = Object.values(types).includes("CATEGORY");
    const hasBooleans = Object.values(types).includes("BOOLEAN");
    elements.normalizeCategories.disabled = !hasCategories;
    elements.normalizeBooleans.disabled = !hasBooleans;
    elements.normalizeCategories.closest(".check-option")?.classList.toggle("option-disabled", !hasCategories);
    elements.normalizeBooleans.closest(".check-option")?.classList.toggle("option-disabled", !hasBooleans);
    if (elements.optionNote) {
      elements.optionNote.textContent = hasCategories || hasBooleans
        ? "All source rows and suspicious values remain available; flags never delete negative or unusual data."
        : "This file has no category or Yes/No fields. The report will focus on numeric and date patterns."
    }
  }

  function renderTable(table, rows) {
    table.replaceChildren();
    if (!rows || rows.length === 0) return;
    const columns = Object.keys(rows[0]);
    const thead = document.createElement("thead");
    const header = document.createElement("tr");
    for (const column of columns) {
      const cell = document.createElement("th");
      cell.scope = "col";
      cell.textContent = column;
      header.append(cell);
    }
    thead.append(header);
    const tbody = document.createElement("tbody");
    for (const row of rows) {
      const tr = document.createElement("tr");
      for (const column of columns) {
        const cell = document.createElement("td");
        cell.textContent = row[column] === null ? "" : String(row[column]);
        cell.title = cell.textContent;
        tr.append(cell);
      }
      tbody.append(tr);
    }
    table.append(thead, tbody);
  }

  async function changeSheet() {
    if (!sessionId) return;
    elements.sheetSelect.disabled = true;
    try {
      const data = await requestJson("/api/sheet", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ session_id: sessionId, sheet: elements.sheetSelect.value }),
      });
      elements.previewCount.textContent = Number(data.rows).toLocaleString();
      elements.columnCount.textContent = Number(data.column_count).toLocaleString();
      elements.count.textContent = `${data.columns.length} ${data.columns.length === 1 ? "field" : "fields"} found`;
      availableColumns = data.columns;
    elements.filterList?.replaceChildren();
    elements.validationList?.replaceChildren();
    elements.filterPreview?.classList.add("hidden");
    if (elements.filterSummary) elements.filterSummary.textContent = "Showing all rows";
      ambiguousDates = {};
      renderColumns(data.columns);
      renderTable(elements.inputTable, data.preview);
      const suggestion = data.revenue_suggestion;
      elements.deriveRevenueContainer.classList.toggle("hidden", !suggestion);
      elements.deriveRevenue.checked = false;
      if (suggestion) elements.revenueSuggestionLabel.textContent = `Add optional computed revenue (${suggestion.quantity} × ${suggestion.price})`;
      refreshFocusOptions();
      showToast(`Loaded worksheet “${elements.sheetSelect.value}”.`);
    } catch (error) {
      showToast(error.message, true);
    } finally {
      elements.sheetSelect.disabled = false;
    }
  }

  async function analyze() {
    if (!sessionId) {
      showToast("Choose and load a data file first.", true);
      return;
    }
    const button = elements.analyze;
    button.disabled = true;
    button.classList.add("busy");
    elements.progress.classList.remove("hidden");
    button.querySelector("span:first-child").textContent = "Making sense of your data…";
    try {
      const data = await requestJson("/api/analyze", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          session_id: sessionId,
          types: currentTypes(),
          date_orders: dateOrders(),
          normalize_categories: elements.normalizeCategories.checked,
          normalize_booleans: elements.normalizeBooleans.checked,
          derive_revenue: elements.deriveRevenue.checked,
          measure: elements.measure.value || null,
          dimension: elements.dimension.value || null,
          filters: collectRules("filter"),
          validation_rules: collectRules("validation"),
        }),
      });
      showResults(data);
      showToast("Your cleaned data and report are ready.");
    } catch (error) {
      showToast(error.message, true);
    } finally {
      button.disabled = false;
      button.classList.remove("busy");
      elements.progress.classList.add("hidden");
      button.querySelector("span:first-child").textContent = "Clean & create report";
    }
  }

  function renderQuality(data) {
    const list = document.querySelector("#quality-list");
    list.replaceChildren();
    const issues = data.quality_issues || [];
    document.querySelector("#quality-count").textContent = `${Number(data.quality_issue_count || 0).toLocaleString()} flag(s)`;
    if (!issues.length) {
      const item = document.createElement("li");
      item.textContent = "No data-quality flags were detected.";
      list.append(item);
      return;
    }
    for (const issue of issues) {
      const item = document.createElement("li");
      const title = document.createElement("strong");
      title.textContent = `${issue.kind.replaceAll("_", " ")}${issue.column ? ` · ${issue.column}` : ""}: `;
      item.append(title, document.createTextNode(issue.message));
      if (issue.rows && issue.rows.length) {
        item.append(document.createTextNode(` Source row(s): ${issue.rows.join(", ")}${issue.rows.length === 10 ? "…" : ""}.`));
      }
      list.append(item);
    }
    const merges = Object.entries(data.category_merges || {});
    for (const [column, groups] of merges) {
      for (const group of groups) {
        const item = document.createElement("li");
        item.textContent = `${column}: merged ${group.merged.join(", ")} into ${group.canonical}.`;
        list.append(item);
      }
    }
  }

  function showResults(data) {
    document.querySelector("#clean-row-result").textContent = Number(data.clean_rows).toLocaleString();
    document.querySelector("#duplicate-result").textContent = Number(data.duplicates_removed).toLocaleString();
    document.querySelector("#blank-result").textContent = Number(data.blank_rows_removed).toLocaleString();
    document.querySelector("#field-result").textContent = Number(data.columns).toLocaleString();
    document.querySelector("#report-link").href = `/download/${encodeURIComponent(data.report)}`;
    document.querySelector("#csv-link").href = `/download/${encodeURIComponent(data.clean_csv)}`;
    renderQuality(data);
    renderTable(elements.resultTable, data.preview);
    elements.results.classList.remove("hidden");
    setCurrentStep("results");
    elements.results.scrollIntoView({ behavior: "smooth", block: "start" });
    document.querySelector("#results-title").focus({ preventScroll: true });
  }

  function setTheme(theme) {
    document.documentElement.dataset.theme = theme;
    elements.themeToggle.setAttribute("aria-label", theme === "dark" ? "Switch to light theme" : "Switch to dark theme");
    try {
      localStorage.setItem("clearview-theme", theme);
    } catch {
      // The interface remains usable when browser storage is disabled.
    }
  }

  elements.browseButton.addEventListener("click", (event) => {
    event.stopPropagation();
    elements.fileInput.click();
  });
  elements.dropzone.addEventListener("click", (event) => {
    if (event.target.closest("button")) return;
    elements.fileInput.click();
  });
  elements.dropzone.addEventListener("keydown", (event) => {
    if (event.key === "Enter" || event.key === " ") {
      event.preventDefault();
      elements.fileInput.click();
    }
  });
  elements.fileInput.addEventListener("change", () => selectFile(elements.fileInput.files[0]));
  elements.dropzone.addEventListener("dragover", (event) => {
    event.preventDefault();
    elements.dropzone.classList.add("drag-over");
  });
  elements.dropzone.addEventListener("dragenter", (event) => {
    event.preventDefault();
    elements.dropzone.classList.add("drag-over");
  });
  elements.dropzone.addEventListener("dragleave", (event) => {
    if (!event.relatedTarget || !elements.dropzone.contains(event.relatedTarget)) {
      elements.dropzone.classList.remove("drag-over");
    }
  });
  elements.dropzone.addEventListener("drop", (event) => {
    event.preventDefault();
    elements.dropzone.classList.remove("drag-over");
    selectFile(event.dataTransfer.files[0]);
  });
  elements.sheetSelect.addEventListener("change", changeSheet);
  elements.addFilter?.addEventListener("click", () => addRule("filter"));
  elements.applyFilters?.addEventListener("click", applyFilters);
  elements.clearFilters?.addEventListener("click", () => {
    elements.filterList?.replaceChildren();
    elements.filterPreview?.classList.add("hidden");
    updateFilterSummary("Showing all rows");
  });
  elements.addValidation?.addEventListener("click", () => addRule("validation"));
  elements.analyze.addEventListener("click", analyze);
  elements.sampleButton.addEventListener("click", loadSample);
  document.querySelector("#change-file").addEventListener("click", () => elements.fileInput.click());
  document.querySelector("#restart-button").addEventListener("click", () => {
    sessionId = null;
    elements.workspace.classList.add("hidden");
    elements.results.classList.add("hidden");
    elements.fileInput.value = "";
    setCurrentStep("upload");
    document.querySelector("#upload-title").focus({ preventScroll: true });
    elements.fileInput.click();
  });
  elements.themeToggle.addEventListener("click", () => {
    setTheme(document.documentElement.dataset.theme === "dark" ? "light" : "dark");
  });
  hero?.addEventListener("pointermove", setHeroMotion);
  hero?.addEventListener("pointerleave", () => {
    hero.style.setProperty("--pointer-x", "0px");
    hero.style.setProperty("--pointer-y", "0px");
  });

  try {
    const storedTheme = localStorage.getItem("clearview-theme");
    const preferredTheme = window.matchMedia("(prefers-color-scheme: dark)").matches ? "dark" : "light";
    setTheme(storedTheme || preferredTheme);
  } catch {
    setTheme("light");
  }
})();
