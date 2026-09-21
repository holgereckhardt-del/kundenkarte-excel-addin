/* global Office, Excel */

const ui = {
  customerCode: document.getElementById("customerCode"),
  customerName: document.getElementById("customerName"),
  customerAddress: document.getElementById("customerAddress"),
  mapFrame: document.getElementById("mapFrame"),
  mapState: document.getElementById("mapState"),
  mapsButton: document.getElementById("mapsButton"),
  routeButton: document.getElementById("routeButton"),
  refreshButton: document.getElementById("refreshButton"),
  status: document.getElementById("status"),
};

let activeAddress = "";
let refreshTimer;
let requestNumber = 0;

Office.onReady(async (info) => {
  if (info.host !== Office.HostType.Excel) {
    showError("Dieses Add-in kann nur in Excel verwendet werden.");
    return;
  }

  ui.refreshButton.addEventListener("click", refreshCustomer);
  ui.mapsButton.addEventListener("click", () => openExternal(
    `https://www.google.com/maps/search/?api=1&query=${encodeURIComponent(activeAddress)}`,
  ));
  ui.routeButton.addEventListener("click", () => openExternal(
    `https://www.google.com/maps/dir/?api=1&destination=${encodeURIComponent(activeAddress)}`,
  ));

  try {
    await registerWorkbookEvents();
    await refreshCustomer();
  } catch (error) {
    showError(readableError(error));
  }
});

async function registerWorkbookEvents() {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItemOrNullObject("OrderSP");
    sheet.load("isNullObject");
    await context.sync();

    if (sheet.isNullObject) {
      throw new Error("Das Tabellenblatt ‚OrderSP‘ wurde nicht gefunden.");
    }

    sheet.onChanged.add((event) => {
      if (/^W[3-7](?::W[3-7])?$/i.test(event.address) || /^W3:/i.test(event.address)) {
        scheduleRefresh(250);
      }
    });
    sheet.onCalculated.add(() => scheduleRefresh(120));
    await context.sync();
  });
}

function scheduleRefresh(delay) {
  window.clearTimeout(refreshTimer);
  refreshTimer = window.setTimeout(refreshCustomer, delay);
}

async function refreshCustomer() {
  const currentRequest = ++requestNumber;
  setLoading("Kundendaten werden gelesen …");

  try {
    const customer = await readCustomerFromWorkbook();
    if (currentRequest !== requestNumber) return;

    ui.customerCode.textContent = customer.code ? `Kundennummer ${customer.code}` : "Keine Kundennummer";
    ui.customerName.textContent = customer.name || "Kunde auswählen";
    ui.customerAddress.textContent = customer.address || "Wähle den Kunden über die vorhandene Kundensuche aus.";

    if (!customer.address) {
      activeAddress = "";
      showEmpty("Für diesen Kunden ist keine vollständige Anschrift hinterlegt.");
      return;
    }

    activeAddress = customer.address;
    setActionState(true);
    ui.status.textContent = "Standort wird gesucht …";

    const point = await geocode(customer.address);
    if (currentRequest !== requestNumber) return;

    if (!point) {
      showEmpty("Die Anschrift konnte auf der Karte nicht gefunden werden.");
      return;
    }

    showMap(point);
    ui.status.textContent = "Die Karte wird bei einem Kundenwechsel automatisch aktualisiert.";
  } catch (error) {
    if (currentRequest === requestNumber) showError(readableError(error));
  }
}

async function readCustomerFromWorkbook() {
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem("OrderSP");
    const range = sheet.getRange("W3:W7");
    range.load("text");
    await context.sync();

    const rows = range.text.map((row) => (row[0] || "").trim());
    const [code, name, street, postcodeCity, country] = rows;
    const address = [street, postcodeCity, country].filter(Boolean).join(", ");
    return { code, name, address };
  });
}

async function geocode(address) {
  const key = `kundenkarte:${address.toLocaleLowerCase("de-DE")}`;
  const cached = localStorage.getItem(key);
  if (cached) return JSON.parse(cached);

  const query = new URLSearchParams({
    q: address,
    format: "jsonv2",
    limit: "1",
    "accept-language": "de",
  });
  const response = await fetch(`https://nominatim.openstreetmap.org/search?${query.toString()}`, {
    headers: { Accept: "application/json" },
  });

  if (!response.ok) throw new Error("Der Kartendienst ist momentan nicht erreichbar.");
  const results = await response.json();
  if (!results.length) return null;

  const point = { lat: Number(results[0].lat), lon: Number(results[0].lon) };
  localStorage.setItem(key, JSON.stringify(point));
  return point;
}

function showMap({ lat, lon }) {
  const latitudeDelta = 0.008;
  const longitudeDelta = 0.015;
  const bbox = [lon - longitudeDelta, lat - latitudeDelta, lon + longitudeDelta, lat + latitudeDelta]
    .map((value) => value.toFixed(6))
    .join(",");
  const params = new URLSearchParams({
    bbox,
    layer: "mapnik",
    marker: `${lat.toFixed(6)},${lon.toFixed(6)}`,
  });
  ui.mapFrame.src = `https://www.openstreetmap.org/export/embed.html?${params.toString()}`;
  ui.mapState.hidden = true;
  ui.mapFrame.hidden = false;
}

function setLoading(message) {
  ui.mapFrame.hidden = true;
  ui.mapState.hidden = false;
  ui.mapState.className = "map-state";
  ui.mapState.innerHTML = '<div class="spinner" aria-hidden="true"></div><p>Karte wird geladen …</p>';
  ui.status.textContent = message;
  setActionState(false);
}

function showEmpty(message) {
  ui.mapFrame.hidden = true;
  ui.mapState.hidden = false;
  ui.mapState.className = "map-state empty";
  ui.mapState.innerHTML = `<span class="pin" aria-hidden="true">⌖</span><p>${escapeHtml(message)}</p>`;
  ui.status.textContent = message;
  setActionState(Boolean(activeAddress));
}

function showError(message) {
  ui.mapFrame.hidden = true;
  ui.mapState.hidden = false;
  ui.mapState.className = "map-state error";
  ui.mapState.innerHTML = `<p>${escapeHtml(message)}</p>`;
  ui.status.textContent = "Bitte prüfe die Arbeitsmappe und versuche es erneut.";
  setActionState(false);
}

function setActionState(enabled) {
  ui.mapsButton.disabled = !enabled;
  ui.routeButton.disabled = !enabled;
}

function openExternal(url) {
  if (!activeAddress) return;
  if (Office.context.ui && Office.context.ui.openBrowserWindow) {
    Office.context.ui.openBrowserWindow(url);
  } else {
    window.open(url, "_blank", "noopener,noreferrer");
  }
}

function readableError(error) {
  return error && error.message ? error.message : "Die Kundenkarte konnte nicht geladen werden.";
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}
