let summary = {};
let incompleteTrips = [];
let plateSummary = {};

document.getElementById('uploadForm').addEventListener('submit', async function (e) {
  e.preventDefault();

  const bvFile = document.getElementById('bvReport').files[0];
  const taabFile = document.getElementById('taabRegister').files[0];

  const bvData = await parseFile(bvFile);
  const taabData = await parseFile(taabFile);

  const bvHeaders = Object.keys(bvData[0]);
  const taabHeaders = Object.keys(taabData[0]);

  const bvMap = buildHeaderMap(bvHeaders);
  const taabMap = buildHeaderMap(taabHeaders);

  const taabLookup = {};
  taabData.forEach(entry => {
    const rawId = entry[taabMap.cardId];
    const normalizedId = normalizeCardId(rawId);
    taabLookup[normalizedId] = entry;
  });

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
    card.trips += 1;
    card.fare += fare;
    card.dates.add(date);
    totalFare += fare;
    totalTrips += 1;

    const isIncomplete = toFro === "2" || (!departure && boarding);
    if (isIncomplete) {
      card.incomplete += 1;
      totalIncomplete += 1;
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
      card.completed += 1;
    }

    if (boarding && (!card.start || boarding < card.start)) card.start = boarding;
    if (departure && (!card.end || departure > card.end)) card.end = departure;

    const cat = card.category;
    if (!categorySummary[cat]) categorySummary[cat] = { trips: 0, fare: 0 };
    categorySummary[cat].trips += 1;
    categorySummary[cat].fare += fare;

    if (!plateSummary[plate]) plateSummary[plate] = { fare: 0, trips: 0 };
    plateSummary[plate].fare += fare;
    plateSummary[plate].trips += 1;
  });

  renderBubbles(totalTrips, totalFare, totalIncomplete, categorySummary, plateSummary);
  renderTable(summary);
  renderIncompleteTable(incompleteTrips);
});

async function parseFile(file) {
  const ext = file.name.split('.').pop().toLowerCase();
  let workbook;

  if (ext === 'csv') {
    // 1. Read raw text
    const text = await file.text();
    const lines = text.split(/\r\n|\n/);

    // 2. Define keywords that must appear in your header row
    const headerKeywords = [
      'card',     // covers Card No, Card ID
      'fare',     // covers Fare, Deduction
      'boarding', // covers Boarding time, Boarding station
      'departure',// covers Departure time, Departure station
      'operation' // covers Operation Date
    ];

    // 3. Find the first line containing ≥3 of those keywords
    const headerIndex = lines.findIndex(line => {
      const low = line.toLowerCase();
      const matches = headerKeywords.filter(kw => low.includes(kw)).length;
      return matches >= 3;
    });

    // 4. If nothing found, assume line 0; otherwise slice from header
    const csvContent = lines
      .slice(headerIndex >= 0 ? headerIndex : 0)
      .join('\n');

    // 5. Parse the cleaned CSV into a workbook
    workbook = XLSX.read(csvContent, { type: 'string' });

  } else {
    // .xlsx fallback
    const data = await file.arrayBuffer();
    workbook = XLSX.read(data, { type: 'array' });
  }

  // Convert first sheet to JSON
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  return XLSX.utils.sheet_to_json(sheet);
}

function normalizeHeaders(headers) {
  return headers.map(h => h.trim().toLowerCase().replace(/[\s_]+/g, ''));
}

function buildHeaderMap(headers) {
  const normalized = normalizeHeaders(headers);
  const map = {};

  normalized.forEach((h, i) => {
    if (h.includes("cardno")) map.cardNo = headers[i];
    if (h.includes("cardid")) map.cardId = headers[i];
    if (h.includes("name")) map.name = headers[i];
    if (h.includes("phone")) map.phone = headers[i];
    if (h.includes("address")) map.address = headers[i];
    if (h.includes("category")) map.category = headers[i];
    if (h.includes("brandedcompany")) map.brandedCompany = headers[i];
    if (h.includes("fare")) map.fare = headers[i];
    if (h.includes("boardingtime")) map.boardingTime = headers[i];
    if (h.includes("departuretime")) map.departureTime = headers[i];
    if (h.includes("operationdate")) map.operationDate = headers[i];
    if (h.includes("tofro")) map.toFro = headers[i];
    if (h.includes("platenumber")) map.plateNumber = headers[i];
    if (h.includes("routeid")) map.routeId = headers[i];
    if (h.includes("boardingstation")) map.boardingStation = headers[i];
  });

  return map;
}

function normalizeCardId(id) {
  return id ? id.toLowerCase().replace(/[^a-f0-9]/gi, '') : '';
}

function renderBubbles(trips, fare, missed, categories, plates) {
  const container = document.getElementById('overallSummary');
  container.innerHTML = `
    <div class="bubble"><h3>Total Trips</h3><p>${trips}</p></div>
    <div class="bubble"><h3>Total Fare</h3><p>$${fare.toFixed(2)}</p></div>
    <div class="bubble"><h3>Incomplete Trips</h3><p>${missed}</p></div>
    ${Object.entries(categories).map(([cat, data]) => `
      <div class="bubble">
        <h3>${cat}</h3>
        <p>Trips: ${data.trips}</p>
        <p>Fare: $${data.fare.toFixed(2)}</p>
      </div>
    `).join('')}
    ${Object.entries(plates).map(([plate, data]) => `
      <div class="bubble">
        <h3>Bus ${plate}</h3>
        <p>Trips: ${data.trips}</p>
        <p>Fare: $${data.fare.toFixed(2)}</p>
      </div>
    `).join('')}
  `;
}

function renderTable(summary) {
  const section = document.getElementById('summaryTable');
  section.innerHTML = '';

  const table = document.createElement('table');
  table.innerHTML = `
    <thead>
      <tr>
        <th>Card No</th>
        <th>Name</th>
        <th>Phone</th>
        <th>Address</th>
        <th>Category</th>
        <th>Company</th>
        <th>Trips</th>
        <th>Fare</th>
        <th>Completed</th>
        <th>Incomplete</th>
        <th>Start</th>
        <th>End</th>
      </tr>
    </thead>
    <tbody>
      ${Object.entries(summary).map(([cardNo, card]) => `
        <tr>
          <td>${cardNo}</td>
          <td>${card.name}</td>
          <td>${card.phone}</td>
          <td>${card.address}</td>
          <td>${card.category}</td>
          <td>${card.company}</td>
          <td>${card.trips}</td>
          <td>$${card.fare.toFixed(2)}</td>
          <td>${card.completed}</td>
          <td>${card.incomplete}</td>
          <td>${card.start}</td>
          <td>${card.end}</td>
        </tr>
      `).join('')}
    </tbody>
  `;
  section.appendChild(table);
}

function renderIncompleteTable(incompleteTrips) {
  const section = document.getElementById('incompleteTable');
  section.innerHTML = '';

  const table = document.createElement('table');
  table.innerHTML = `
    <thead>
      <tr>
        <th>Card No</th>
        <th>Name</th>
        <th>Operation Date</th>
        <th>Boarding Time</th>
        <th>Departure Time</th>
        <th>Plate Number</th>
        <th>Route ID</th>
        <th>Boarding Station</th>
      </tr>
    </thead>
    <tbody>
      ${incompleteTrips.map(trip => `
        <tr>
          <td>${trip.cardNo}</td>
          <td>${trip.name}</td>
          <td>${trip.date}</td>
          <td>${trip.boarding}</td>
          <td>${trip.departure || '—'}</td>
          <td>${trip.plate}</td>
          <td>${trip.route}</td>
          <td>${trip.station}</td>
        </tr>
      `).join('')}
    </tbody>
  `;
  section.appendChild(table);
}

document.getElementById('exportBtn').addEventListener('click', () => {
  const workbook = XLSX.utils.book_new();

  // Sheet 1: Card Summary
  const cardRows = [["Card No", "Name", "Phone", "Address", "Category", "Company", "Trips", "Fare", "Completed", "Incomplete", "Start", "End"]];
  Object.entries(summary).forEach(([cardNo, card]) => {
    cardRows.push([
      cardNo, card.name, card.phone, card.address,
      card.category, card.company, card.trips,
      card.fare.toFixed(2), card.completed, card.incomplete,
      card.start, card.end
    ]);
  });
  const cardSheet = XLSX.utils.aoa_to_sheet(cardRows);
  XLSX.utils.book_append_sheet(workbook, cardSheet, "Card Summary");

  // Sheet 2: Incomplete Trips
  const incompleteRows = [["Card No", "Name", "Operation Date", "Boarding Time", "Departure Time", "Plate Number", "Route ID", "Boarding Station"]];
  incompleteTrips.forEach(trip => {
    incompleteRows.push([
      trip.cardNo, trip.name, trip.date, trip.boarding,
      trip.departure || '', trip.plate, trip.route, trip.station
    ]);
  });
  const incompleteSheet = XLSX.utils.aoa_to_sheet(incompleteRows);
  XLSX.utils.book_append_sheet(workbook, incompleteSheet, "Incomplete Trips");

  // Sheet 3: Fare by Bus
  const plateRows = [["Plate Number", "Trips", "Total Fare"]];
  Object.entries(plateSummary).forEach(([plate, data]) => {
    plateRows.push([plate, data.trips, data.fare.toFixed(2)]);
  });
  const plateSheet = XLSX.utils.aoa_to_sheet(plateRows);
  XLSX.utils.book_append_sheet(workbook, plateSheet, "Fare by Bus");

  XLSX.writeFile(workbook, "B-MohBel_Usage_Summary.xlsx");
});
