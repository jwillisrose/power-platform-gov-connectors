const state = {
  records: [],
  search: "",
  cloud: "",
  product: "",
  tier: "",
  status: "",
  sort: "name-asc",
};

const elements = {
  form: document.querySelector("#filters"),
  search: document.querySelector("#search"),
  cloud: document.querySelector("#cloud"),
  product: document.querySelector("#product"),
  tier: document.querySelector("#tier"),
  status: document.querySelector("#status"),
  sort: document.querySelector("#sort"),
  results: document.querySelector("#connector-results"),
  resultCount: document.querySelector("#result-count"),
  emptyState: document.querySelector("#empty-state"),
  emptyClear: document.querySelector("#empty-clear"),
  lastUpdated: document.querySelector("#last-updated"),
};

function getLogoUrl(record) {
  const fileName = record.logo_path?.split(/[\\/]/).pop();
  return fileName ? `assets/logos/${encodeURIComponent(fileName)}` : record.icon_url;
}

function createBadge(text, className = "") {
  const badge = document.createElement("span");
  badge.className = `badge ${className}`.trim();
  badge.textContent = text;
  return badge;
}

function createConnectorCard(record) {
  const card = document.createElement("article");
  card.className = "connector-card";

  const titleRow = document.createElement("div");
  titleRow.className = "connector-title-row";

  const logo = document.createElement("img");
  logo.className = "connector-logo";
  logo.src = getLogoUrl(record);
  logo.alt = "";
  logo.loading = "lazy";
  logo.width = 48;
  logo.height = 48;

  const titleContainer = document.createElement("div");
  titleContainer.style.minWidth = "0";

  const heading = document.createElement("h3");
  heading.textContent = record.name;

  const publisher = document.createElement("p");
  publisher.className = "publisher";
  publisher.textContent = record.publisher || "Publisher not listed";
  publisher.title = publisher.textContent;

  titleContainer.append(heading, publisher);
  titleRow.append(logo, titleContainer);

  const badges = document.createElement("div");
  badges.className = "badges";
  badges.append(
    createBadge(record.tier, record.tier === "Premium" ? "premium" : ""),
    createBadge(record.status, record.status === "Preview" ? "preview" : ""),
  );
  record.availability_products.forEach((product) => badges.append(createBadge(product)));

  const availability = document.createElement("div");
  availability.className = "availability";
  for (const cloud of ["GCC", "GCC High", "DoD"]) {
    const cloudState = document.createElement("span");
    const isAvailable = record.availability[cloud];
    cloudState.className = `cloud-state${isAvailable ? " available" : ""}`;
    cloudState.textContent = `${isAvailable ? "✓ " : ""}${cloud}`;
    cloudState.setAttribute(
      "aria-label",
      `${cloud}: ${isAvailable ? "available" : "not available"}`,
    );
    availability.append(cloudState);
  }

  const link = document.createElement("a");
  link.className = "connector-link";
  link.href = record.learn_url;
  link.rel = "noreferrer";
  link.textContent = "View Microsoft Learn →";

  card.append(titleRow, badges, availability, link);
  return card;
}

function matchesFilters(record) {
  const searchable = `${record.name} ${record.publisher} ${record.publisher_raw}`.toLocaleLowerCase();
  return (
    (!state.search || searchable.includes(state.search)) &&
    (!state.cloud || record.availability[state.cloud]) &&
    (!state.product || record.availability_products.includes(state.product)) &&
    (!state.tier || record.tier === state.tier) &&
    (!state.status || record.status === state.status)
  );
}

function render() {
  const records = state.records
    .filter(matchesFilters)
    .sort((a, b) => {
      const direction = state.sort === "name-desc" ? -1 : 1;
      return a.name.localeCompare(b.name) * direction;
    });

  const fragment = document.createDocumentFragment();
  records.forEach((record) => fragment.append(createConnectorCard(record)));
  elements.results.replaceChildren(fragment);
  elements.results.setAttribute("aria-busy", "false");

  elements.resultCount.textContent =
    records.length === state.records.length
      ? `${records.length} connectors`
      : `${records.length} of ${state.records.length} connectors`;
  elements.emptyState.hidden = records.length !== 0;
}

function syncStateFromForm() {
  state.search = elements.search.value.trim().toLocaleLowerCase();
  state.cloud = elements.cloud.value;
  state.product = elements.product.value;
  state.tier = elements.tier.value;
  state.status = elements.status.value;
  state.sort = elements.sort.value;
  render();
}

function clearFilters() {
  elements.form.reset();
  state.search = "";
  state.cloud = "";
  state.product = "";
  state.tier = "";
  state.status = "";
  render();
  elements.search.focus();
}

function renderSummary(data) {
  document.querySelector("#total-count").textContent = data.records.length;
  document.querySelector("#gcc-count").textContent = data.records.filter(
    (record) => record.availability.GCC,
  ).length;
  document.querySelector("#gcch-count").textContent = data.records.filter(
    (record) => record.availability["GCC High"],
  ).length;
  document.querySelector("#dod-count").textContent = data.records.filter(
    (record) => record.availability.DoD,
  ).length;

  const dates = Object.values(data.source_dates)
    .filter(Boolean)
    .map((date) => new Date(date))
    .filter((date) => !Number.isNaN(date.valueOf()));
  const latestDate = new Date(Math.max(...dates));
  elements.lastUpdated.textContent = `Microsoft Learn sources updated ${latestDate.toLocaleDateString(
    "en-US",
    { year: "numeric", month: "long", day: "numeric", timeZone: "UTC" },
  )}`;
}

async function initialize() {
  try {
    const response = await fetch("data/connectors.json");
    if (!response.ok) {
      throw new Error(`Connector data request failed with status ${response.status}.`);
    }
    const data = await response.json();
    if (!Array.isArray(data.records)) {
      throw new Error("Connector data does not contain a records array.");
    }

    state.records = data.records;
    renderSummary(data);
    render();
  } catch (error) {
    console.error(error);
    elements.results.setAttribute("aria-busy", "false");
    elements.resultCount.textContent = "Connector data could not be loaded.";
    elements.emptyState.hidden = false;
    elements.emptyState.querySelector("h3").textContent = "Unable to load connector data";
    elements.emptyState.querySelector("p").textContent =
      "Refresh the page or try again later. If the problem continues, contact the site owner.";
    elements.emptyState.querySelector("button").hidden = true;
  }
}

elements.form.addEventListener("input", syncStateFromForm);
elements.form.addEventListener("change", syncStateFromForm);
elements.form.addEventListener("reset", () => window.setTimeout(clearFilters, 0));
elements.sort.addEventListener("change", syncStateFromForm);
elements.emptyClear.addEventListener("click", clearFilters);

initialize();
