/* Shared presentation for the existing three SKUs. Serve over HTTP(S), not file://. */
(() => {
  "use strict";
  const siteBase = new URL("../", document.currentScript.src);
  const app = document.getElementById("app");
  const knownSkus = ["110-38650", "400-90753", "110-40359"];

  function element(tag, text, className) {
    const node = document.createElement(tag);
    if (text !== undefined) node.textContent = text;
    if (className) node.className = className;
    return node;
  }
  function link(path, text, className) {
    const node = element("a", text, className);
    node.href = new URL(path, siteBase).href;
    return node;
  }
  function normalize(value) { return value.trim().replace(/[\s-]/g, ""); }
  function fact(label, value) {
    const box = element("div", undefined, "fact");
    box.append(element("dt", label, "label"), element("dd", value, "value"));
    return box;
  }
  function detailList(items) {
    const list = element("dl", undefined, "facts-list");
    items.forEach(([label, value]) => {
      const row = element("div");
      row.append(element("dt", label), element("dd", value));
      list.append(row);
    });
    return list;
  }
  function section(title, children) {
    const node = element("section", undefined, "card section-card");
    node.append(element("h2", title), ...children);
    return node;
  }
  function endpointIsConfigured(endpoint) {
    return typeof endpoint === "string" && /^https:\/\/script\.google\.com\/macros\/s\/[A-Za-z0-9_-]+\/exec$/.test(endpoint);
  }

  function renderCatalog(data) {
    const intro = element("div", undefined, "card section-card");
    intro.append(element("h1", "Juki part-number lookup"), element("p", "Start with the exact part number. Then check the machine model and variant, compatibility, OEM or compatible options, and availability.", "sub"));
    const search = element("form");
    const label = element("label", "Exact part number");
    label.htmlFor = "part-search";
    const input = element("input");
    Object.assign(input, { id: "part-search", name: "part_search", type: "search", placeholder: "e.g. 400-90753", maxLength: 80, required: true });
    const button = element("button", "Find part", "cta");
    button.type = "submit";
    const row = element("div", undefined, "search-row");
    row.append(input, button);
    const result = element("p", undefined, "search-result");
    result.setAttribute("role", "status");
    search.append(label, row, result);
    search.addEventListener("submit", (event) => {
      event.preventDefault();
      result.replaceChildren();
      const query = normalize(input.value);
      const part = data.parts.find((item) => normalize(item.sku) === query);
      if (part) {
        result.append(link(`juki/${part.sku}.html`, `${part.sku} — ${part.name}: view details and request a fit check`));
      } else if (query === "11096500") {
        result.append(element("span", "110-96500 is a historical / alternate reference only. Supersession and interchangeability are unverified. "), link("juki/400-90753.html", "View the reference note for 400-90753"));
      } else {
        result.textContent = "No exact part-number match in this catalog. Machine-model compatibility is unknown; do not infer a fit from a similar number.";
      }
    });
    intro.append(search);
    const catalog = element("div", undefined, "catalog");
    data.parts.forEach((part) => {
      const card = element("article", undefined, "card");
      card.append(element("h2", part.sku, "sku"), element("h3", part.name), element("p", "Machine fit: unknown. OEM / compatible options and availability: unknown.", "sub"), link(`juki/${part.sku}.html`, "View part & request fit check", "cta"));
      catalog.append(card);
    });
    app.replaceChildren(intro, catalog);
  }

  function renderForm(part, endpoint) {
    const side = element("aside", undefined, "card side");
    side.append(element("span", "Fit & availability unconfirmed", "status"), element("h2", "Request an exact fit check"), element("p", "Provide your full machine model and the number on the existing part. A request does not confirm fit, reserve stock or place an order.", "sub"));
    const configured = endpointIsConfigured(endpoint);
    const status = element("p", configured ? "" : "Fit requests are not enabled yet. This form cannot send or store your details until submissions are configured.", "notice");
    status.hidden = configured;
    status.setAttribute("role", "status");
    const form = element("form");
    form.method = "post";
    form.target = "_top";
    form.acceptCharset = "UTF-8";
    if (configured) form.action = endpoint;
    const fields = element("fieldset");
    fields.disabled = !configured;
    // Fixed markup only; all catalog data is assigned through DOM properties.
    fields.innerHTML = `
      <label for="machine-model">Machine model</label>
      <input id="machine-model" name="machine_model" required maxlength="120" autocomplete="off" placeholder="Full model and variant, or Unknown" aria-describedby="machine-help">
      <p id="machine-help" class="help">Include every suffix. If unknown, enter Unknown; fit remains unconfirmed.</p>
      <label for="part-number">Existing part number</label>
      <input id="part-number" name="part_number" required maxlength="80" autocomplete="off" placeholder="Number on the old part, or Unknown" aria-describedby="part-help">
      <p id="part-help" class="help">Read the old part marking; the page SKU is not proof of a match.</p>
      <label for="preference">OEM / compatible preference</label>
      <select id="preference" name="preference" required>
        <option value="">Choose a preference</option>
        <option value="oem">Genuine OEM only</option>
        <option value="compatible">Compatible is acceptable</option>
        <option value="either">Either — please explain the options</option>
      </select>
      <label for="email">Email</label>
      <input id="email" name="email" type="email" required maxlength="254" autocomplete="email" placeholder="you@example.com">
      <label for="notes">Notes (optional)</label>
      <textarea id="notes" name="notes" maxlength="2000" placeholder="Old part markings, machine variant, symptoms or questions. Do not include sensitive or payment information."></textarea>
      <input type="hidden" name="sku">
      <input type="hidden" name="page_url">
    `;
    fields.querySelector('[name="sku"]').value = part.sku;
    const pageUrl = new URL(`juki/${part.sku}.html`, siteBase).href;
    fields.querySelector('[name="page_url"]').value = pageUrl;
    const privacy = element("p", undefined, "tiny");
    privacy.append(element("span", "Submitting sends these details to the site operator via Google Sheets and Gmail to review your request. "), link("privacy.html", "Read the privacy notice."));
    const submit = element("button", "Request fit & availability check", "cta");
    submit.type = "submit";
    submit.disabled = !configured;
    let submitting = false;
    form.addEventListener("submit", (event) => {
      if (!configured || submitting) { event.preventDefault(); return; }
      // Native POST: the backend confirms persistence before returning success.
      // Never treat an opaque no-cors response or iframe load as confirmation.
      fields.querySelector('[name="page_url"]').value = pageUrl;
      submitting = true;
      submit.disabled = true;
      submit.textContent = "Sending request…";
    });
    window.addEventListener("pageshow", () => {
      submitting = false;
      submit.disabled = !configured;
      submit.textContent = "Request fit & availability check";
    });
    form.append(fields, privacy, submit);
    side.append(status, form);
    return side;
  }

  function renderPart(data) {
    const part = data.parts.find((item) => item.sku === app.dataset.sku);
    if (!part) throw new Error("SKU not in catalog");
    const crumbs = element("nav", undefined, "crumbs");
    crumbs.setAttribute("aria-label", "Breadcrumb");
    crumbs.append(link("index.html", "Juki parts"), element("span", ` › ${part.sku}`));
    const product = element("article", undefined, "card product");
    const heading = element("h1");
    heading.append(element("span", part.sku, "sku"), element("span", part.name));
    const facts = element("dl", undefined, "partbox");
    facts.append(fact("Exact part number", `${part.sku} / ${part.sku.replace(/-/g, "")}`), fact("Part type", part.part_type), fact("Machine model / variant", "Unknown — provide your exact model"), fact("Compatibility", "Unknown — fit needs verification"));
    product.append(heading, element("p", "Match the existing part number first, then check the full machine model and variant. A matching search result alone does not establish compatibility.", "sub"), facts);
    const compatibility = section("Machine model & compatibility", [detailList([
      ["Confirmed machine models", "Unknown. No manufacturer fit documentation has been verified here."],
      ["Exact variant and part marking", "Required for review. Similar machine names or part numbers do not establish fit."],
      ["Dimensions and specifications", "Unknown. Compare manufacturer documentation and the existing part before any purchase."]
    ])]);
    if (part.references.length) {
      const reference = element("div", undefined, "reference");
      reference.append(element("h3", "Historical / alternate reference"));
      part.references.forEach((item) => reference.append(element("p", `${item.part_number}: ${item.note}`)));
      compatibility.append(reference);
    }
    const options = section("OEM vs compatible", [detailList([
      ["Genuine OEM option", "Availability and provenance unknown. State whether genuine OEM is required."],
      ["Compatible option", "Availability and fit unknown. An acceptable preference is not a compatibility guarantee."]
    ])]);
    const availability = section("Availability", [detailList([
      ["Stock, price and lead time", "Unknown. Requires a current availability check."],
      ["Quote and fulfillment", "Unconfirmed. A fit request is an inquiry; no purchase or reservation is created."]
    ])]);
    const grid = element("div", undefined, "grid2");
    grid.append(options, availability);
    const details = element("div");
    details.append(product, compatibility, grid);
    const hero = element("div", undefined, "hero");
    hero.append(details, renderForm(part, data.lead_endpoint));
    const steps = element("ol");
    ["Exact existing part number, including all digits and markings.", "Full machine model and variant.", "Compatibility review against the machine documentation and old part.", "OEM or compatible preference and provenance check.", "Current availability, price and lead time check."].forEach((text) => steps.append(element("li", text)));
    app.replaceChildren(crumbs, hero, section("What a fit request needs", [steps, element("p", "Manufacturer documentation for this listing has not yet been verified.", "tiny")]));
  }

  async function start() {
    try {
      const response = await fetch(new URL("data/parts.json", siteBase), { cache: "no-cache" });
      if (!response.ok) throw new Error("Catalog unavailable");
      const data = await response.json();
      if (data.system_version !== "0.2" || !Array.isArray(data.parts) || data.parts.length !== knownSkus.length || knownSkus.some((sku) => data.parts.filter((part) => part.sku === sku).length !== 1)) throw new Error("Invalid catalog");
      if (app.dataset.view === "catalog") renderCatalog(data);
      else renderPart(data);
    } catch (_) {
      app.replaceChildren(element("h1", "Part information is unavailable"), element("p", "Please reload later. Machine fit and availability remain unknown. No request has been submitted.", "notice"), link("index.html", "Return to part-number lookup"));
    } finally {
      app.setAttribute("aria-busy", "false");
    }
  }
  if (app) start();
})();
