let summary = {};
let incompleteTrips = [];

document.getElementById('uploadForm').addEventListener('submit', async function (e) {
  e.preventDefault();

  const bvFile = document.getElementById('bvReport').files[0];
  const taabFile = document.getElementById('taabRegister').files[0];

  const bvData = await parseFile(bvFile);
  const taabData = await parseFile(taabFile);

  const taabMap = {};
  taabData.forEach(entry => {
    taabMap[entry["Card ID"]] = entry;
  });

  summary = {};
  incompleteTrips = [];
  const categorySummary = {};
  let totalFare = 0;
  let totalTrips = 0;
  let totalIncomplete = 0;

  bvData.forEach(entry => {
    const cardNo = entry["Card No"];
    const fare = parseFloat(entry["Fare"]) || 0;
    const boarding = entry["Boarding time"];
    const departure = entry["Departure time"];
    const date = entry["Operation Date"];
    const completed = entry["to & fro"] === "1";

    if (!summary[cardNo]) {
      const tag = taabMap[cardNo] || {};
      summary[cardNo] = {
        name: tag["Name"] || "",
        phone: tag["Phone"] || "",
        address: tag["Address"] || "",
        category: tag["Category"] || "Unregistered",
        company: tag["Branded Company"] || "Unknown",
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

    const isIncomplete = entry["to & fro"] === "2" || (!departure && boarding);
    if (isIncomplete) {
      card.incomplete += 1;
      totalIncomplete += 1;
      incompleteTrips.push({
        cardNo,
        name: card.name,
        date,
        boarding,
        departure
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
  });

  renderBubbles(totalTrips, totalFare, totalIncomplete, categorySummary);
  renderTable(summary);
  renderIncompleteTable(incompleteTrips);
});

async function parseFile(file) {
  const data = await file.arrayBuffer();
  const workbook = XLSX.read(data, { type: "array" });
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  return XLSX.utils.sheet_to_json(sheet);
}

function renderBubbles(trips, fare, missed, categories) {
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
  `;
}

function renderTable(summary) {
  const section = document.getElementById('summaryTable');
  const table = document.createElement('table');
  table.innerHTML = `
    <thead>
      <tr>
        <th>Card No</th><th>Name</th><th>Phone</th><th>Address</th>
        <th>Category</th><th>Company</th><th>Trips</th><th>Fare</th>
        <th>Completed</th><th>Incomplete</th><th>Start</th><th>End</th>
      </tr>
    </thead>
    <tbody>
      ${Object.entries(summary).map(([cardNo, card]) => `
        <tr>
          <td>${cardNo}</td><td>${card.name}</td><td>${card.phone}</td><td>${card.address}</td>
          <td>${card.category}</td><td>${card.company}</td><td>${card.trips}</td>
          <td>$${card.fare.toFixed(2)}</td><td>${card.completed}</td><td>${card.incomplete}</td>
          <td>${card.start}</td><td>${card.end}</td>
        </tr>
      `).join('')}
    </tbody>
  `;
  section.innerHTML = '';
  section.appendChild(table);
}

function renderIncompleteTable(incompleteTrips) {
  const section = document.getElementById('incompleteTable');
  const table = document.createElement('table');
  table.innerHTML = `
    <thead>
      <tr><th>Card No</th><th>Name</th><th>Operation Date</th><th>Boarding Time</th><th>Departure Time</th></tr>
    </thead>
    <tbody>
      ${incompleteTrips.map(trip => `
        <tr>
          <td>${trip.cardNo}</td><td>${trip.name}</td><td>${trip.date}</td>
          <td>${trip.boarding}</td><td>${trip.departure || '—'}</td>
        </tr>
      `).join('')}
    </tbody>
  `;
  section.appendChild(table);
}

document.getElementById('exportBtn').addEventListener('click', () => {
  const rows = [["Card No", "Name", "Phone", "Address", "Category", "Company", "Trips", "Fare", "Completed", "Incomplete", "Start", "End"]];
  Object.entries(summary).forEach(([cardNo, card]) => {
    rows.push([
      cardNo, card.name, card.phone, card.address, card.category, card.company,
      card.trips, card.fare.toFixed(2), card.completed, card.incomplete,
      card.start, card.end
    ]);
  });

  const worksheet = XLSX.utils.aoa_to_sheet(rows);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Card Summary");

  const incompleteRows = [["Card No", "Name", "Operation Date", "Boarding Time", "Departure Time"]];
  incompleteTrips.forEach(trip => {
    incompleteRows.push([
      trip.cardNo, trip.name, trip.date, trip.boarding, trip.departure || ''
    ]);
  });

  const incompleteSheet = XLSX.utils.aoa_to_sheet(incompleteRows);
  XLSX.utils.book_append_sheet(workbook, incompleteSheet, "Incomplete Trips");

  XLSX.writeFile(workbook, "B-MohBel_Card_Usage_Summary.xlsx");
});
