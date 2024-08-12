const socket = io();

let openServices = new Set();

socket.on('updateIDsWithOrders1', (data) => {
/*     console.log('Datos de Planilla 1:', data); */
    updateOrders(data, 'Planilla 1');
});

socket.on('updateIDsWithOrders2', (data) => {
/*     console.log('Datos de Planilla 2:', data); */
    updateOrders(data, 'Planilla 2');
});

function updateOrders(data, planillaName) {
    const container = document.getElementById('orders-container');
    
    // Save the state of open services
    document.querySelectorAll('.service-header').forEach(header => {
        if (!header.nextElementSibling.classList.contains('hidden')) {
            openServices.add(header.querySelector('span').textContent.split(' ')[0]);
        } else {
            openServices.delete(header.querySelector('span').textContent.split(' ')[0]);
        }
    });
    
    // Clear previous content if this is the first update
    if (planillaName === 'Planilla 1') {
        container.innerHTML = '';
    }

    // Group orders by service
    const ordersByService = data.reduce((acc, order) => {
        const service = order.service || 'Sin Servicio';
        if (!acc[service]) {
            acc[service] = [];
        }
        acc[service].push(order);
        return acc;
    }, {});

    // Create dropdown menus for each service
    Object.keys(ordersByService).forEach(service => {
        const serviceDiv = document.createElement('div');
        serviceDiv.classList.add('service-container');

        const serviceHeader = document.createElement('div');
        serviceHeader.classList.add('service-header');
        
        const serviceSpan = document.createElement('span');
        serviceSpan.textContent = `${service} (${ordersByService[service].length})`;
        serviceHeader.appendChild(serviceSpan);

        serviceHeader.addEventListener('click', () => {
            serviceContent.classList.toggle('hidden');
            if (serviceContent.classList.contains('hidden')) {
                openServices.delete(service);
            } else {
                openServices.add(service);
            }
        });

        const serviceContent = document.createElement('div');
        serviceContent.classList.add('service-content', 'hidden');

        ordersByService[service].forEach(order => {
            const orderDiv = document.createElement('div');
            orderDiv.classList.add('order-card');
            
            const id = document.createElement('span');
            id.textContent = `ID: ${order.id}`;
            orderDiv.appendChild(id);

            const date = document.createElement('span');
            date.textContent = `Fecha: ${order.date}`;
            orderDiv.appendChild(date);

            const orderService = document.createElement('span'); 
            orderService.textContent = `Servicio: ${order.service}`;
            orderDiv.appendChild(orderService);

            const orderNumber = document.createElement('span');
            orderNumber.textContent = `Número de Pedido: ${order.orderNumber}`;
            orderDiv.appendChild(orderNumber);

            const url = document.createElement('a');
            url.href = order.url;
            url.textContent = 'Ver Venta';
            url.target = '_blank';
            orderDiv.appendChild(url);

            serviceContent.appendChild(orderDiv);
        });

        serviceDiv.appendChild(serviceHeader);
        serviceDiv.appendChild(serviceContent);
        container.appendChild(serviceDiv);

        // Restore the open/closed state
        if (openServices.has(service)) {
            serviceContent.classList.remove('hidden');
        }
    });
}
