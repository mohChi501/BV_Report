let summary = {};
let incompleteTrips = [];
let plateSummary = {};

document.getElementById('uploadForm').addEventListener('submit', async function (e) {
  e.preventDefault();

  const bvFile = document.getElementById('bvReport').files[0];
  const taabFile = document.getElementById('taabRegister').files[0];

  const bvData = await parseFile(bvFile);
  const taabData = await parseFile(taabFile);

  const bvHeaders = Object.keys(bvData[0] || {});
  const taabHeaders = Object.keys(taabData[0] || {});

  const bvMap = buildHeaderMap(bvHeaders);
  const taabMap = buildHeaderMap(taabHeaders);

  const taabLookup = {};
  taabData.forEach(entry => {
    const rawId = entry[taabMap.cardId];
    const normalizedId = normalizeCardId(rawId);
    taabLookup[normalizedId] = entry;
  });

  // Reset summaries
  summary = {};
  incompleteTrips = [];
  plateSummary = {};
  const categorySummary = {};
  let totalFare = 0;
  let totalTrips = 0;
  let totalIncomplete = 0;

  bvData.forEach(entry => {
    const rawCardNo = entry[bvMap.cardNo];
    const cardNo = normalizeCardId(rawCardNo);
    const fare = parseFloat(entry[bvMap.fare]) || 0;
    const boarding = entry[bvMap.boardingTime];
    const departure = entry[bvMap.departureTime];
    const date = entry[bvMap.operationDate];
    const toFro = entry[bvMap.toFro];
    const plate = entry[bvMap.plateNumber];
    const route = entry[bvMap.routeId];
    const station = entry[bvMap.boardingStation];

    if (!summary[cardNo]) {
      const tag = taabLookup[cardNo] || {};
      summary[cardNo] = {
        name: tag[taabMap.name] || "",
        phone: tag[taabMap.phone] || "",
        address: tag[taabMap.address] || "",
        category: tag[taabMap.category] || "Unregistered",
        company: tag[taabMap.brandedCompany] || "Unknown",
        trips: 0,
        fare: 0,
        completed: 0,
        incomplete: 0,
        start: boarding,
        end: departure,
        dates: new Set()
      };
    }

    const card = summary[cardNo];
    card.trips++;
    card.fare += fare;
    card.dates.add(date);
    totalFare += fare;
    totalTrips++;

    const isIncomplete = toFro === "2" || (!departure && boarding);
    if (isIncomplete) {
      card.incomplete++;
      totalIncomplete++;
      incompleteTrips.push({
        cardNo,
        name: card.name,
        date,
        boarding,
        departure,
        plate,
        route,
        station
      });
    } else {
      card.completed++;
    }

    if (boarding && (!card.start || boarding < card.start)) card.start = boarding;
    if (departure && (!card.end || departure > card.end)) card.end = departure;

    // Category summary
    const cat = card.category;
    if (!categorySummary[cat]) categorySummary[cat] = { trips: 0, fare: 0 };
    categorySummary[cat].trips++;
    categorySummary[cat].fare += fare;

    // Plate summary
    if (!plateSummary[plate]) plateSummary[plate] = { trips: 0, fare: 0 };
    plateSummary[plate].trips++;
    plateSummary[plate].fare += fare;
  });

  renderBubbles(totalTrips, totalFare, totalIncomplete, categorySummary, plateSummary);
  renderTable(summary);
  renderIncompleteTable(incompleteTrips);
});


/**
 * Parses .csv or .xlsx, auto-locating the real header row.
 */
async function parseFile(file) {
  const ext = file.name.split('.').pop().toLowerCase();
  let workbook;

  // Read workbook
  if (ext === 'csv') {
    const text = await file.text();
    workbook = XLSX.read(text, { type: 'string' });
  } else {
    const data = await file.arrayBuffer();
    workbook = XLSX.read(data, { type: 'array' });
  }

  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  // Get all rows as arrays
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, blankrows: false });

  // Identify header row (must contain card + fare + boarding)
  const headerIdx = rows.findIndex(r => {
    const cells = r.map(c => (c||'').toString().toLowerCase());
    return cells.some(c => c.includes('card')) &&
           cells.some(c => c.includes('fare')) &&
           cells.some(c => c.includes('boarding'));
  });

  // If we found it, slice below; otherwise fallback to default JSON
  if (headerIdx >= 0) {
    const header = rows[headerIdx];
    const dataRows = rows.slice(headerIdx + 1);
    return dataRows.map(r => {
      const obj = {};
      header.forEach((h, i) => { obj[h] = r[i]; });
      return obj;
    });
  } else {
    return XLSX.utils.sheet_to_json(sheet);
  }
}


function normalizeHeaders(headers) {
  return headers.map(h => h.trim().toLowerCase().replace(/[\s_]+/g, ''));
}

function buildHeaderMap(headers) {
  const normalized = normalizeHeaders(headers);
  const map = {};

  normalized.forEach((h, i) => {
    if (h.includes("cardno"))         map.cardNo        = headers[i];
    if (h.includes("cardid"))         map.cardId        = headers[i];
    if (h.includes("name"))           map.name          = headers[i];
    if (h.includes("phone"))          map.phone         = headers[i];
    if (h.includes("address"))        map.address       = headers[i];
    if (h.includes("category"))       map.category      = headers[i];
    if (h.includes("brandedcompany")) map.brandedCompany= headers[i];
    if (h.includes("fare"))           map.fare          = headers[i];
    if (h.includes("boardingtime"))   map.boardingTime  = headers[i];
    if (h.includes("departuretime"))  map.departureTime = headers[i];
    if (h.includes("operationdate"))  map.operationDate = headers[i];
    if (h.includes("tofro"))          map.toFro         = headers[i];
    if (h.includes("platenumber"))    map.plateNumber   = headers[i];
    if (h.includes("routeid"))        map.routeId       = headers[i];
    if (h.includes("boardingstation"))map.boardingStation= headers[i];
  });

  return map;
}

function normalizeCardId(id) {
  return id ? id.toString().toLowerCase().replace(/[^a-f0-9]/gi, '') : '';
}

function renderBubbles(trips, fare, missed, categories, plates) {
  const c = document.getElementById('overallSummary');
  c.innerHTML = `
    <div class="bubble"><h3>Total Trips</h3><p>${trips}</p></div>
    <div class="bubble"><h3>Total Fare</h3><p>$${fare.toFixed(2)}</p></div>
    <div class="bubble"><h3>Incomplete Trips</h3><p>${missed}</p></div>
    ${Object.entries(categories).map(([cat,data])=>`
      <div class="bubble">
        <h3>${cat}</h3>
        <p>Trips: ${data.trips}</p>
        <p>Fare: $${data.fare.toFixed(2)}</p>
      </div>`).join('')}
    ${Object.entries(plates).map(([pl,data])=>`
      <div class="bubble">
        <h3>Bus ${pl}</h3>
        <p>Trips: ${data.trips}</p>
        <p>Fare: $${data.fare.toFixed(2)}</p>
      </div>`).join('')}
  `;
}

function renderTable(summary) {
  const sec = document.getElementById('summaryTable');
  sec.innerHTML = '';
  const tbl = document.createElement('table');
  tbl.innerHTML = `
    <thead><tr>
      <th>Card No</th><th>Name</th><th>Phone</th><th>Address</th>
      <th>Category</th><th>Company</th><th>Trips</th><th>Fare</th>
      <th>Completed</th><th>Incomplete</th><th>Start</th><th>End</th>
    </tr></thead>
    <tbody>
      ${Object.entries(summary).map(([cn,cd])=>`
      <tr>
        <td>${cn}</td><td>${cd.name}</td><td>${cd.phone}</td><td>${cd.address}</td>
        <td>${cd.category}</td><td>${cd.company}</td><td>${cd.trips}</td>
        <td>$${cd.fare.toFixed(2)}</td><td>${cd.completed}</td><td>${cd.incomplete}</td>
        <td>${cd.start}</td><td>${cd.end}</td>
      </tr>`).join('')}
    </tbody>`;
  sec.appendChild(tbl);
}

function renderIncompleteTable(incompleteTrips) {
  const sec = document.getElementById('incompleteTable');
  sec.innerHTML = '';
  const tbl = document.createElement('table');
  tbl.innerHTML = `
    <thead><tr>
      <th>Card No</th><th>Name</th><th>Date</th>
      <th>Boarding</th><th>Departure</th>
      <th>Plate</th><th>Route</th><th>Station</th>
    </tr></thead>
    <tbody>
      ${incompleteTrips.map(t=>`
      <tr>
        <td>${t.cardNo}</td><td>${t.name}</td><td>${t.date}</td>
        <td>${t.boarding}</td><td>${t.departure||'—'}</td>
        <td>${t.plate}</td><td>${t.route}</td><td>${t.station}</td>
      </tr>`).join('')}
    </tbody>`;
  sec.appendChild(tbl);
}

document.getElementById('exportBtn').addEventListener('click', () => {
  const wb = XLSX.utils.book_new();

  // Card Summary
  const cards = [["Card No","Name","Phone","Address","Category","Company","Trips","Fare","Completed","Incomplete","Start","End"]];
  Object.entries(summary).forEach(([cn,cd])=>{
    cards.push([cn,cd.name,cd.phone,cd.address,cd.category,cd.company,cd.trips,cd.fare.toFixed(2),cd.completed,cd.incomplete,cd.start,cd.end]);
  });
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(cards), "Card Summary");

  // Incomplete Trips
  const inc = [["Card No","Name","Date","Boarding","Departure","Plate","Route","Station"]];
  incompleteTrips.forEach(t=>{
    inc.push([t.cardNo,t.name,t.date,t.boarding,t.departure||'',t.plate,t.route,t.station]);
  });
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(inc), "Incomplete Trips");

  // Fare by Bus
  const bus = [["Plate","Trips","Fare"]];
  Object.entries(plateSummary).forEach(([pl,data])=>{
    bus.push([pl,data.trips,data.fare.toFixed(2)]);
  });
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(bus), "Fare by Bus");

  XLSX.writeFile(wb, "B-MohBel_Usage_Summary.xlsx");
});
