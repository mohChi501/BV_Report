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

  const summary = {};
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

    if (completed && departure) {
      card.completed += 1;
    } else {
      card.incomplete += 1;
    }

    if (boarding && (!card.start || boarding < card.start)) card.start = boarding;
    if (departure && (!card.end || departure > card.end)) card.end = departure;
  });

  renderCards(summary);
  renderTable(summary);
});

async function parseFile(file) {
  const data = await file.arrayBuffer();
  const workbook = XLSX.read(data, { type: "array" });
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  return XLSX.utils.sheet_to_json(sheet);
}

function renderCards(summary) {
  const container = document.getElementById('summaryCards');
  container.innerHTML = '';
  Object.entries(summary).forEach(([cardNo, card]) => {
    const div = document.createElement('div');
    div.className = 'card';
    div.innerHTML = `
      <h3>${card.name || 'Unnamed Card'}</h3>
      <p><strong>Card No:</strong> ${cardNo}</p>
      <p><strong>Category:</strong> ${card.category}</p>
      <p><strong>Company:</strong> ${card.company}</p>
      <p><strong>Trips:</strong> ${card.trips}</p>
      <p><strong>Total Fare:</strong> $${card.fare.toFixed(2)}</p>
      <p><strong>Completed:</strong> ${card.completed}</p>
      <p><strong>Incomplete:</strong> ${card.incomplete}</p>
      <p><strong>Start:</strong> ${card.start}</p>
      <p><strong>End:</strong> ${card.end}</p>
    `;
    container.appendChild(div);
  });
}

function renderTable(summary) {
  const section = document.getElementById('summaryTable');
  const table = document.createElement('table');
  table.innerHTML = `
    <thead>
      <tr>
        <th>Card No</th>
        <th>Name</th>
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
  section.innerHTML = '';
  section.appendChild(table);
}

document.getElementById('exportBtn').addEventListener('click', () => {
  const rows = [["Card No", "Name", "Category", "Company", "Trips", "Fare", "Completed", "Incomplete", "Start", "End"]];
  Object.entries(summary).forEach(([cardNo, card]) => {
    rows.push([
      cardNo, card.name, card.category, card.company,
      card.trips, card.fare.toFixed(2), card.completed,
      card.incomplete, card.start, card.end
    ]);
  });
  const worksheet = XLSX.utils.aoa_to_sheet(rows);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Summary");
  XLSX.writeFile(workbook, "Card_Usage_Summary.xlsx");
});
