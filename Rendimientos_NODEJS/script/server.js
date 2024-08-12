const express = require('express');
const path = require('path');
const http = require('http');
const { Server } = require('socket.io');
const { google } = require('googleapis');

// Configuración de identificadores de hojas de cálculo y nombres de hojas
const SERVICE_ACCOUNT_FILE = path.join(__dirname, 'service-account.json');
const SPREADSHEET_ID_1 = '1dCd4QLXt8WMuAKjlQ88JrLOqFxMAex5_mh_ykxLq3FQ'; // PEDIDOS MENSAJERIA
const SPREADSHEET_ID_2 = '1hJIk3bpR5zLzEux5lEb1eUevh21Z2cVhB4ySpWEKZNU'; // PLANILLA COLECTA/CORREOS
const SPREADSHEET_ID_3 = '1MxqTFSBUL8UA0HeABTMI-6Sk8tH5WgK-HvDfqGimll8'; // CERRADOR 1
const SPREADSHEET_ID_4 = '1RRKrOiq6VuKtNx2MMWkBMBQ5H8KBuBs9MtKN-o1zH5c'; // Tabla dinámica 5
const SHEET_NAME_1 = 'Presentacion y cierre';
const SHEET_NAME_2 = 'Presentacion y cierre';
const SHEET_NAME_3 = 'CERRADOR 1';
const SHEET_NAME_4 = 'Tabla dinámica 5';

// Autenticación de Google API
const auth = new google.auth.GoogleAuth({
  keyFile: SERVICE_ACCOUNT_FILE,
  scopes: ['https://www.googleapis.com/auth/spreadsheets.readonly'],
});

const app = express();
const server = http.createServer(app);
const io = new Server(server);
const port = 3000;

// Servir archivos estáticos
app.use(express.static(path.join(__dirname, '../')));

// Función para convertir tiempo en formato HH:MM:SS a minutos
function convertTimeToMinutes(timeStr) {
  const [hours, minutes, seconds] = timeStr.split(':').map(parseFloat);
  return (hours * 60) + minutes + (seconds / 60);
}

// Obtener pedidos pendientes de una hoja de cálculo específica
async function fetchPendingOrders(spreadsheetId, sheetName) {
  try {
    const client = await auth.getClient();
    const sheets = google.sheets({ version: 'v4', auth: client });

    const result = await sheets.spreadsheets.values.get({
      spreadsheetId: spreadsheetId,
      range: `${sheetName}!C:D`,
    });

    const rows = result.data.values;

    if (rows && rows.length > 1) {
      const pendingOrders = rows.slice(1).reduce((acc, row) => {
        const service = row[0]?.trim();
        const pedido = row[1]?.trim();
        if (pedido && pedido !== '') {
          acc[service] = (acc[service] || 0) + 1;
        }
        return acc;
      }, {});

      console.log(`Pending orders from ${sheetName}:`, pendingOrders); // Log para verificar datos

      return pendingOrders;
    } else {
      return {};
    }
  } catch (error) {
    console.error('Error fetching pending orders:', error);
    return {};
  }
}

// Obtener pedidos cerrados de una hoja de cálculo específica
async function fetchClosedOrders(spreadsheetId, sheetName) {
  try {
    const client = await auth.getClient();
    const sheets = google.sheets({ version: 'v4', auth: client });

    const result = await sheets.spreadsheets.values.get({
      spreadsheetId: spreadsheetId,
      range: `${sheetName}!A:D`,
    });

    const rows = result.data.values;

    if (rows && rows.length > 1) {
      const closedOrders = rows.slice(1).reduce((acc, row) => {
        const user = row[3]?.trim();
        if (user) {
          if (!acc[user]) {
            acc[user] = 0;
          }
          acc[user]++;
        }
        return acc;
      }, {});

      return closedOrders;
    } else {
      return {};
    }
  } catch (error) {
    console.error('Error fetching closed orders:', error);
    return {};
  }
}

// Obtener tiempos de usuario de una hoja de cálculo específica
async function fetchUserTimes(spreadsheetId, sheetName) {
  try {
    const client = await auth.getClient();
    const sheets = google.sheets({ version: 'v4', auth: client });

    const result = await sheets.spreadsheets.values.get({
      spreadsheetId: spreadsheetId,
      range: `${sheetName}!A:E`,
    });

    const rows = result.data.values;

    if (rows && rows.length > 1) {
      const userTimes = rows.slice(1).reduce((acc, row) => {
        const user = row[0]?.trim();
        const concept = row[4]?.trim();
        const timeStr = row[3]?.trim();

        if (user && (concept === 'TIEMPO_ENTRE_PAQUETES' || concept === 'EMPAQUETANDO_PEDIDO' || concept === 'REVISANDO_PEDIDO')) {
          const time = convertTimeToMinutes(timeStr);
          if (!acc[user]) {
            acc[user] = { tiempoEntrePaquetes: 0, empaquetando: 0, revisando: 0 };
          }
          if (concept === 'TIEMPO_ENTRE_PAQUETES') {
            acc[user].tiempoEntrePaquetes += time;
          } else if (concept === 'EMPAQUETANDO_PEDIDO') {
            acc[user].empaquetando += time;
          } else if (concept === 'REVISANDO_PEDIDO') {
            acc[user].revisando += time;
          }
        }
        return acc;
      }, {});

      return userTimes;
    } else {
      return {};
    }
  } catch (error) {
    console.error('Error fetching user times:', error);
    return {};
  }
}

// Calcular la velocidad de empaquetado
async function calculatePackagingSpeed(userTimes, closedCounts) {
  const packagingSpeed = {};

  for (const user in userTimes) {
    const times = userTimes[user];
    const totalPackagingTime = times.tiempoEntrePaquetes + times.empaquetando + times.revisando;
    const closedOrders = closedCounts[user] || 1;
    const speed = totalPackagingTime / closedOrders;

    packagingSpeed[user] = speed;
  }

  return packagingSpeed;
}

// Obtener IDs con números de pedido
async function fetchIDsWithOrderNumbers(spreadsheetId, sheetName) {
  try {
    const client = await auth.getClient();
    const sheets = google.sheets({ version: 'v4', auth: client });

    const result = await sheets.spreadsheets.values.get({
      spreadsheetId: spreadsheetId,
      range: `${sheetName}!A:F`,
    });

    const rows = result.data.values;

    if (rows && rows.length > 1) {
      const data = rows.slice(1).reduce((acc, row) => {
        const id = row[0]?.trim();
        const date = row[1]?.trim();
        const service = row[2]?.trim();
        const orderNumber = row[3]?.trim();
        const url = row[5]?.trim();
        if (id && orderNumber) {
          acc.push({ id, date, service, orderNumber, url });
        }
        return acc;
      }, []);

      return data;
    } else {
      return [];
    }
  } catch (error) {
    console.error('Error fetching IDs with order numbers:', error);
    return [];
  }
}

// Conexión de sockets
io.on('connection', async (socket) => {
  const count1 = await fetchPendingOrders(SPREADSHEET_ID_1, SHEET_NAME_1);
  const count2 = await fetchPendingOrders(SPREADSHEET_ID_2, SHEET_NAME_2);
  const closedCounts = await fetchClosedOrders(SPREADSHEET_ID_3, SHEET_NAME_3);
  const userTimes = await fetchUserTimes(SPREADSHEET_ID_4, SHEET_NAME_4);
  const packagingSpeed = await calculatePackagingSpeed(userTimes, closedCounts);

  const idsWithOrders1 = await fetchIDsWithOrderNumbers(SPREADSHEET_ID_1, SHEET_NAME_1);
  const idsWithOrders2 = await fetchIDsWithOrderNumbers(SPREADSHEET_ID_2, SHEET_NAME_2);

  socket.emit('updateCount1', count1);
  socket.emit('updateCount2', count2);
  socket.emit('updateClosedCounts', closedCounts);
  socket.emit('updateUserTimes', userTimes);
  socket.emit('updatePackagingSpeed', packagingSpeed);

  socket.emit('updateIDsWithOrders1', idsWithOrders1);
  socket.emit('updateIDsWithOrders2', idsWithOrders2);
});

// Intervalo para actualizar datos periódicamente
setInterval(async () => {
  const count1 = await fetchPendingOrders(SPREADSHEET_ID_1, SHEET_NAME_1);
  const count2 = await fetchPendingOrders(SPREADSHEET_ID_2, SHEET_NAME_2);
  const closedCounts = await fetchClosedOrders(SPREADSHEET_ID_3, SHEET_NAME_3);
  const userTimes = await fetchUserTimes(SPREADSHEET_ID_4, SHEET_NAME_4);
  const packagingSpeed = await calculatePackagingSpeed(userTimes, closedCounts);

  const idsWithOrders1 = await fetchIDsWithOrderNumbers(SPREADSHEET_ID_1, SHEET_NAME_1);
  const idsWithOrders2 = await fetchIDsWithOrderNumbers(SPREADSHEET_ID_2, SHEET_NAME_2);

  io.emit('updateCount1', count1);
  io.emit('updateCount2', count2);
  io.emit('updateClosedCounts', closedCounts);
  io.emit('updateUserTimes', userTimes);
  io.emit('updatePackagingSpeed', packagingSpeed);

  io.emit('updateIDsWithOrders1', idsWithOrders1);
  io.emit('updateIDsWithOrders2', idsWithOrders2);
}, 15000);

// Iniciar el servidor
server.listen(port, () => {
  console.log(`Servidor escuchando en http://localhost:${port}`);
});
