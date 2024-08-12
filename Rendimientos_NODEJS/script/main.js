const socket = io();

let closedCounts = {};
let userTimes = {};
let packagingSpeed = {};

socket.on('updateCount1', (counts) => {
  updateCounts(counts, 'counter-container-1'); // Asegúrate de que el ID del contenedor sea correcto
});

socket.on('updateCount2', (counts) => {
  updateCounts(counts, 'counter-container-2'); // Asegúrate de que el ID del contenedor sea correcto
});

socket.on('updateClosedCounts', (counts) => {
  closedCounts = counts;
  updateClosedOrdersTable();
  updateTotalClosedCount();
});

socket.on('updateUserTimes', (times) => {
  userTimes = times;
  updateClosedOrdersTable();
});

socket.on('updatePackagingSpeed', (speeds) => {
  packagingSpeed = speeds;
  updateClosedOrdersTable();
});

function updateCounts(counts, containerId) {
  const container = document.getElementById(containerId);

  // Guardar las claves actuales para los contadores actualizados
  const existingKeys = new Set(Object.keys(counts));

  for (const [key, count] of Object.entries(counts)) {
    let card = document.getElementById(`card-${key}`);
    if (!card) {
      card = document.createElement('div');
      card.className = 'counter';
      card.id = `card-${key}`;

      const header = document.createElement('h2');
      header.textContent = key;
      card.appendChild(header);

      const subheader = document.createElement('div');
      subheader.textContent = 'Total de pendientes';
      card.appendChild(subheader);

      const counterDiv = document.createElement('div');
      counterDiv.id = `counter-${key}`;
      counterDiv.className = 'count';
      card.appendChild(counterDiv);

      container.appendChild(card);
    }

    document.getElementById(`counter-${key}`).textContent = count;
  }

  // Eliminar tarjetas que no están en los datos actualizados
  container.querySelectorAll('.counter').forEach(card => {
    const cardKey = card.id.replace('card-', '');
    if (!existingKeys.has(cardKey)) {
      card.remove();
    }
  });
}

function updateClosedOrdersTable() {
  const tableBody = document.querySelector('#closed-orders-table tbody');
  tableBody.innerHTML = '';

  const sortedUsers = Object.entries(closedCounts).sort((a, b) => b[1] - a[1]);

  sortedUsers.forEach(([user, count], index) => {
    const row = document.createElement('tr');

    const userCell = document.createElement('td');
    userCell.textContent = user;
    row.appendChild(userCell);

    const countCell = document.createElement('td');
    countCell.textContent = count;
    row.appendChild(countCell);

    const timeCell = document.createElement('td');
    const speed = packagingSpeed[user];
    timeCell.textContent = speed ? speed.toFixed(2) : 'N/A';
    row.appendChild(timeCell);

    const medalCell = document.createElement('td');
    const medalImage = document.createElement('img');
    medalImage.className = 'medal';
    if (index === 0) {
      medalImage.src = '/images/medalla-oro.png';
    } else if (index === 1) {
      medalImage.src = '/images/medalla-plata.png';
    } else {
      medalImage.src = '/images/medalla-bronce.png';
    }
    medalCell.appendChild(medalImage);
    row.appendChild(medalCell);

    tableBody.appendChild(row);
  });
}

function updateTotalClosedCount() {
  const totalClosed = Object.values(closedCounts).reduce((sum, count) => sum + count, 0);
  document.getElementById('total-closed').textContent = totalClosed;
}
