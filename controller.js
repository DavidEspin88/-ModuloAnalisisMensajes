// Estado global del módulo
let workbook = null; 
let rawData = [];
let filteredData = [];
let chartInstance = null;
let geoChartInstance = null; 

// Referencias al DOM
const dom = {
    fileInput: document.getElementById('excelInput'),
    sheetSelector: document.getElementById('sheetSelector'), 
    filterStart: document.getElementById('filterStart'),
    filterTimeStart: document.getElementById('filterTimeStart'),
    filterEnd: document.getElementById('filterEnd'),
    filterTimeEnd: document.getElementById('filterTimeEnd'),
    filterCanton: document.getElementById('filterCanton'),
    btnFilter: document.getElementById('btnFilter'),
    btnExport: document.getElementById('btnExport'),
    btnMessage: document.getElementById('btnMessage'),
    tableBody: document.getElementById('tableBody'),
    stats: {
        total: document.getElementById('statTotalOps'),
        pmp: document.getElementById('statTotalPmp'),
        effectiveness: document.getElementById('statEfectividad')
    }
};

// 1. Manejo de Carga de Archivos
dom.fileInput.addEventListener('change', (e) => {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();

    reader.onload = (e) => {
        try {
            const data = new Uint8Array(e.target.result);
            // cellDates: true ayuda a identificar fechas en CSV y formatos texto
            workbook = XLSX.read(data, { type: 'array', cellDates: true });

            rawData = [];
            dom.sheetSelector.innerHTML = '<option value="ALL">-- Ver Todas las Hojas (Consolidado) --</option>';
            
            workbook.SheetNames.forEach(name => {
                const opt = document.createElement('option');
                opt.value = name;
                opt.textContent = name;
                dom.sheetSelector.appendChild(opt);
            });
            dom.sheetSelector.disabled = false;

            loadAllSheets();
        
        } catch (error) {
            console.error(error);
            alert("Error procesando archivo: " + error.message);
        }
    };
    reader.readAsArrayBuffer(file);
});

// Función para hacer el modal de Bootstrap arrastrable
function initDraggableModal() {
    const modalEl = document.getElementById('editModal');
    const header = modalEl.querySelector('.modal-header');
    const dialog = modalEl.querySelector('.modal-dialog');
    
    let isDragging = false;
    let offset = { x: 0, y: 0 };

    header.addEventListener('mousedown', (e) => {
        isDragging = true;
        const rect = dialog.getBoundingClientRect();
        offset.x = e.clientX - rect.left;
        offset.y = e.clientY - rect.top;
        header.style.cursor = 'grabbing';
    });

    document.addEventListener('mousemove', (e) => {
        if (!isDragging) return;
        
        dialog.style.margin = '0';
        dialog.style.position = 'absolute';
        dialog.style.left = (e.clientX - offset.x) + 'px';
        dialog.style.top = (e.clientY - offset.y) + 'px';
    });

    document.addEventListener('mouseup', () => {
        isDragging = false;
        header.style.cursor = 'move';
    });

    // Resetear posición al cerrar el modal
    modalEl.addEventListener('hidden.bs.modal', () => {
        dialog.style.left = '';
        dialog.style.top = '';
        dialog.style.margin = '';
        dialog.style.position = '';
    });
}

dom.sheetSelector.addEventListener('change', (e) => {
    const sheetName = e.target.value;
    if (!workbook) return;
    if (sheetName === 'ALL') loadAllSheets();
    else loadSingleSheet(sheetName);
});

function loadAllSheets() {
    rawData = [];
    let sheetsLoaded = 0;
    workbook.SheetNames.forEach(sheetName => {
        const worksheet = workbook.Sheets[sheetName];
        const sheetRows = extractDataFromSheet(worksheet, sheetName);
        rawData = rawData.concat(sheetRows);
        if (sheetRows.length > 0) sheetsLoaded++;
    });
    finalizeLoad();
}

function loadSingleSheet(sheetName) {
    rawData = [];
    const worksheet = workbook.Sheets[sheetName];
    rawData = extractDataFromSheet(worksheet, sheetName);
    finalizeLoad();
}

function finalizeLoad() {
    if (rawData.length === 0) {
        alert("No se encontraron datos válidos.");
        return;
    }
    populateCantonFilter();
    applyFilters();
}

function extractDataFromSheet(worksheet, sheetName) {
    const matrix = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
    if (!matrix || matrix.length === 0) return [];

    const colMap = {
        ord: -1, fecha: -1, horaInicio: -1, horaFin: -1, tipoOp: -1, operaciones: -1,
        provincia: -1, canton: -1, parroquia: -1, resultados: -1, pmp: -1,
        ofi: -1, aerot: -1, res: -1,
        camioneta: -1, camion: -1, bus: -1
    };

    let dataStartRow = 0;

    for (let r = 0; r < Math.min(20, matrix.length); r++) {
        const row = matrix[r];
        if (!row) continue;
        const rowStr = Array.from(row, cell => (cell ? String(cell).toUpperCase().trim() : ""));
        
        const findIdx = (keywords) => {
            let idx = rowStr.findIndex(cell => cell && keywords.some(k => cell === k));
            if (idx === -1) idx = rowStr.findIndex(cell => cell && keywords.some(k => cell.startsWith(k)));
            if (idx === -1) idx = rowStr.findIndex(cell => cell && keywords.some(k => cell.includes(k)));
            return idx;
        };

        if (colMap.ord === -1) colMap.ord = findIdx(['ORD.', 'ORD', 'NRO', 'NO.', 'NUM']);
        if (colMap.fecha === -1) colMap.fecha = findIdx(['FECHA', 'DÍA', 'DIA']);
        if (colMap.horaInicio === -1) colMap.horaInicio = findIdx(['HORA', 'INICIO', 'H. INI']);
        if (colMap.horaFin === -1) colMap.horaFin = findIdx(['FIN', 'TERMINO', 'H. FIN']);
        if (colMap.horaInicio !== -1 && colMap.horaFin === -1) colMap.horaFin = colMap.horaInicio + 1;
        if (colMap.tipoOp === -1) {
            let idx = findIdx(['TIPO DE OP', 'TIPO OP', 'ACTIVIDAD', 'CLASE', 'DETALLE']);
            if (idx !== -1 && idx !== colMap.ord) colMap.tipoOp = idx;
            else if (idx === -1) colMap.tipoOp = findIdx(['TIPO']);
        }
        if (colMap.operaciones === -1) colMap.operaciones = findIdx(['OPERACIONES', 'ESTADO']);
        if (colMap.provincia === -1) colMap.provincia = findIdx(['PROVINCIA', 'PROV']);
        if (colMap.canton === -1) colMap.canton = findIdx(['CANTÓN', 'CANTON', 'JURISDICCION']);
        if (colMap.parroquia === -1) colMap.parroquia = findIdx(['PARROQUIA', 'PARR']);
        if (colMap.resultados === -1) colMap.resultados = findIdx(['RESULTADOS', 'NOVEDAD', 'RESULTADO']);
        if (colMap.ofi === -1) colMap.ofi = findIdx(['OFI', 'OFIC']);
        if (colMap.aerot === -1) colMap.aerot = findIdx(['AEROT', 'AEROT.', 'TROPA']);
        if (colMap.res === -1) colMap.res = findIdx(['RESV', 'RESER', 'RSV']);
        if (colMap.pmp === -1) colMap.pmp = findIdx(['PMP', 'PERS', 'PERSONAL']);
        if (colMap.camioneta === -1) colMap.camioneta = findIdx(['CAMIONETA', 'LUV', 'DMAX']);
        if (colMap.camion === -1) colMap.camion = findIdx(['CAMION', 'MERCEDES', 'HINO', 'UNI']);
        if (colMap.bus === -1) colMap.bus = findIdx(['BUS', 'BUSES', 'TRANSPORTE']);

        if (colMap.ord !== -1 && colMap.fecha !== -1) dataStartRow = r + 1;
    }

    const sheetData = [];
    for (let i = dataStartRow; i < matrix.length; i++) {
        const row = matrix[i];
        if (!row || row.length === 0) continue;

        const get = (idx) => (idx !== -1 && row[idx] !== undefined && row[idx] !== null) ? row[idx] : "";
        const getNum = (idx) => {
            const val = get(idx);
            if (val === "" || val === null || val === undefined) return 0;
            const cleanStr = String(val).replace(/[^0-9]/g, '');
            if (cleanStr === "") return 0;
            return parseInt(cleanStr, 10) || 0;
        };

        const valProv = String(get(colMap.provincia)).toUpperCase().trim();
        const valCant = String(get(colMap.canton)).toUpperCase().trim();
        if (valProv.includes("PROVINCIA") || valProv === "PROV" || valProv.includes("TOTAL") || valCant.includes("TOTAL")) continue;

        // Filtrado por NO CUMPLIDO (Búsqueda en toda la fila para mayor seguridad)
        const rowString = row.join(" ").toUpperCase();
        const cancelKeywords = ["NO CUMPLIO", "NO SE CUMPLE", "NO SE EJECUTO", "CANCELADA", "SUSPENDIDA", "NO REALIZADA", "NO SE REALIZO"];
        if (cancelKeywords.some(k => rowString.includes(k))) continue;

        // Validación de Tipo de Operación (Indispensable)

        const cleanTipo = String(get(colMap.tipoOp) || "").trim().toUpperCase();
        if (cleanTipo === "" || cleanTipo === "S/T" || cleanTipo === "0") continue;

        let parsedDate = null;
        let dateVal = get(colMap.fecha);
        
        if (dateVal) {
            if (typeof dateVal === 'number') {
                // Serial de Excel
                parsedDate = new Date(Math.round((dateVal - 25569) * 86400 * 1000));
            } else {
                // Texto (CSV o Excel con formato texto)
                const dStr = String(dateVal).trim();
                // Intentar detectar formato DD/MM/YYYY o DD-MM-YYYY
                const parts = dStr.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
                
                if (parts) {
                    // Es formato Latino/Español DD/MM/YYYY
                    parsedDate = new Date(parts[3], parts[2] - 1, parts[1]);
                } else {
                    // Intento estándar (ISO o formato US)
                    parsedDate = new Date(dateVal);
                }
            }
        }
        if ((!parsedDate || isNaN(parsedDate.getTime())) && sheetName) {
            try {
                const cleanName = sheetName.trim().replace(/-/g, ' ');
                const d = new Date(`${cleanName} 2026`);
                if (!isNaN(d.getTime())) parsedDate = d;
            } catch(e) {}
        }

        const formatHora = (val) => {
            if (!val) return "0000";
            if (typeof val === 'number' && val < 1) {
                const totalMinutes = Math.round(val * 24 * 60);
                return String(Math.floor(totalMinutes / 60)).padStart(2, '0') + String(totalMinutes % 60).padStart(2, '0');
            }
            return String(val).replace(/[^0-9]/g, '').padStart(4, '0');
        };

        const hIni = formatHora(get(colMap.horaInicio));
        const idxFin = colMap.horaFin !== -1 ? colMap.horaFin : (colMap.horaInicio !== -1 ? colMap.horaInicio + 1 : -1);
        const hFin = formatHora(get(idxFin));

        let startDate = null;
        let endDate = null;
        
        try {
            if (parsedDate && !isNaN(parsedDate.getTime())) {
                startDate = new Date(parsedDate);
                startDate.setHours(parseInt(hIni.substring(0,2)) || 0, parseInt(hIni.substring(2,4)) || 0, 0);
                endDate = new Date(parsedDate);
                endDate.setHours(parseInt(hFin.substring(0,2)) || 0, parseInt(hFin.substring(2,4)) || 0, 0);
                if (endDate < startDate) endDate.setDate(endDate.getDate() + 1);
            }
        } catch (err) {
            console.warn("Error calculando fechas fila " + i, err);
            startDate = null; 
            endDate = null;
        }

        const valOfi = getNum(colMap.ofi);
        const valAerot = getNum(colMap.aerot);
        const valRes = getNum(colMap.res);

        sheetData.push({
            id: rawData.length + sheetData.length,
            ord: get(colMap.ord),
            fecha: parsedDate,
            startDate, endDate,
            horaMilitar: `${hIni} - ${hFin}`,
            fuerza: "AÉREA",
            tipoOp: cleanTipo,
            operaciones: get(colMap.operaciones) || 'EJECUTADA',
            provincia: get(colMap.provincia),
            canton: get(colMap.canton),
            parroquia: get(colMap.parroquia),
            resultados: get(colMap.resultados) || "0",
            pmp: (valOfi + valAerot + valRes) > 0 ? (valOfi + valAerot + valRes) : getNum(colMap.pmp),
            detPmp: { ofi: valOfi, aerot: valAerot, res: valRes },
            medios: {
                camioneta: getNum(colMap.camioneta),
                camion: getNum(colMap.camion),
                bus: getNum(colMap.bus)
            }
        });
    }
    return sheetData;
}

function populateCantonFilter() {
    const cantones = [...new Set(rawData.map(item => item.canton).filter(c => c))].sort();
    dom.filterCanton.innerHTML = '<option value="TODOS">-- Todos --</option>';
    cantones.forEach(c => {
        const opt = document.createElement('option');
        opt.value = c; opt.textContent = c;
        dom.filterCanton.appendChild(opt);
    });
    populateTipoFilter();
}

function populateTipoFilter() {
    const selector = document.getElementById('filterTipo');
    if(!selector) return;
    const tipos = [...new Set(rawData.map(item => item.tipoOp).filter(t => t))].sort();
    selector.innerHTML = '<option value="TODOS">-- Todos los Tipos --</option>';
    tipos.forEach(t => {
        const opt = document.createElement('option');
        opt.value = t; opt.textContent = t;
        selector.appendChild(opt);
    });
}

dom.btnFilter.addEventListener('click', applyFilters);

function applyFilters() {
    const startDateStr = dom.filterStart.value;
    const endDateStr = dom.filterEnd.value;
    const startTimeStr = dom.filterTimeStart.value || "00:00";
    const endTimeStr = dom.filterTimeEnd.value || "23:59";
    const selectedCanton = dom.filterCanton.value;
    const selectedTipo = document.getElementById('filterTipo').value;

    let filterStart = startDateStr ? new Date(`${startDateStr}T${startTimeStr}`) : null;
    let filterEnd = endDateStr ? new Date(`${endDateStr}T${endTimeStr}`) : null;

    const processedData = rawData.map(item => {
        let matchesTime = true;
        if (filterStart && filterEnd) {
            if (!item.startDate || !item.endDate) {
                matchesTime = false;
            } else {
                // REQUERIMIENTO: La operación debe estar ESTRICTAMENTE dentro del rango (Contención total)
                // Si la operación empieza antes o termina después del filtro, no se considera EJECUTADA en este rango.
                matchesTime = (item.startDate >= filterStart && item.endDate <= filterEnd);
            }
        }
        return { ...item, validTime: matchesTime };
    });

    const groups = {};
    const baseFiltered = processedData.filter(item => {
        let cV = (selectedCanton === 'TODOS' || String(item.canton).toUpperCase().trim() === String(selectedCanton).toUpperCase().trim());
        let tV = (selectedTipo === 'TODOS' || String(item.tipoOp).toUpperCase().trim() === String(selectedTipo).toUpperCase().trim());
        
        // Filtro por fecha (Días completos): Queremos ver lo planificado para esos días
        let matchesDateRange = true;
        if (filterStart && filterEnd) {
            const dStart = new Date(filterStart); dStart.setHours(0,0,0,0);
            const dEnd = new Date(filterEnd); dEnd.setHours(23,59,59,999);
            
            if (!item.startDate || item.startDate < dStart || item.startDate > dEnd) {
                matchesDateRange = false;
            }
        }

        return cV && tV && matchesDateRange;
    });

    baseFiltered.forEach(item => {
        // Incluimos la fecha formateada en la clave de agrupación
        const dateKey = item.fecha ? item.fecha.toISOString().split('T')[0] : 'SIN_FECHA';
        const key = `${item.tipoOp}|${item.provincia}|${item.canton}|${item.parroquia}|${dateKey}`.toUpperCase();
        
        if (!groups[key]) {
            groups[key] = { ...item, sumPlanif: 0, sumEjecut: 0, sumPmp: 0, originalOps: [], resultsList: [] };
        }
        
        // Siempre se cuenta como Planificada si entró en baseFiltered
        groups[key].sumPlanif += 1;
        
        // Solo se cuenta como Ejecutada y se suma PMP si cumple estrictamente el horario
        if (item.validTime) {
            groups[key].sumEjecut += 1;
            groups[key].sumPmp += (item.pmp || 0);
            if (item.resultados && item.resultados !== "0" && item.resultados !== "") {
                groups[key].resultsList.push(item.resultados);
            }
        }
        
        groups[key].originalOps.push(item);
    });

    // Consolidar resultados únicos para la vista
    Object.values(groups).forEach(g => {
        if (g.resultsList.length > 0) {
            g.resultados = [...new Set(g.resultsList)].join(" / ");
        } else {
            g.resultados = "0";
        }
    });

    filteredData = Object.values(groups);

    // Definir prioridad de cantones
    const cantonPriority = {
        "MANTA": 1,
        "MONTECRISTI": 2,
        "JIPIJAPA": 3,
        "PUERTO LÓPEZ": 4,
        "PUERTO LOPEZ": 4 // Por si acaso no tiene tilde
    };

    filteredData.sort((a, b) => {
        // 1. Comparar Fechas
        const dateA = a.fecha ? a.fecha.getTime() : 0;
        const dateB = b.fecha ? b.fecha.getTime() : 0;
        if (dateA !== dateB) return dateA - dateB;

        // 2. Comparar Prioridad de Cantón (dentro de la misma fecha)
        const cantonA = String(a.canton).toUpperCase().trim();
        const cantonB = String(b.canton).toUpperCase().trim();
        
        const prioA = cantonPriority[cantonA] || 999;
        const prioB = cantonPriority[cantonB] || 999;

        if (prioA !== prioB) return prioA - prioB;

        // 3. Orden alfabético para otros cantones no prioritarios
        return cantonA.localeCompare(cantonB);
    });

    updateDashboard();
}

// Configuración Global de Chart.js para Estilo Power BI
if (typeof Chart !== 'undefined') {
    Chart.defaults.font.family = "'Segoe UI', 'Helvetica', 'Arial', sans-serif";
    Chart.defaults.color = '#605e5c';
    Chart.defaults.plugins.tooltip.backgroundColor = 'rgba(255, 255, 255, 0.95)';
    Chart.defaults.plugins.tooltip.titleColor = '#201f1e';
    Chart.defaults.plugins.tooltip.bodyColor = '#201f1e';
    Chart.defaults.plugins.tooltip.borderColor = '#edebe9';
    Chart.defaults.plugins.tooltip.borderWidth = 1;
    Chart.defaults.plugins.tooltip.padding = 10;
    Chart.defaults.plugins.tooltip.cornerRadius = 4;
}

const corporatePalette = ['#001f3f', '#004e8c', '#0078d4', '#2b88d8', '#71afe5', '#a6d8ff', '#c7e0f4'];

// ... (resto de funciones de carga y procesamiento iguales)

function updateDashboard() {
    const searchTerm = document.getElementById('searchInput').value.toLowerCase().trim();
    dom.tableBody.innerHTML = '';
    
    const displayData = filteredData.filter(item => {
        if (!searchTerm) return true;
        const rowText = `${item.fuerza} ${item.tipoOp} ${item.provincia} ${item.canton} ${item.parroquia} ${item.resultados}`.toLowerCase();
        return rowText.includes(searchTerm);
    });

    if (displayData.length === 0) {
        dom.tableBody.innerHTML = `<tr><td colspan="11" class="text-center py-4">${filteredData.length > 0 ? 'Sin resultados para la búsqueda.' : 'Cargue un archivo...'}</td></tr>`;
        if (document.getElementById('tableTotalPlanif')) document.getElementById('tableTotalPlanif').textContent = '0';
    } else {
        let totalPlanifTabla = 0;
        displayData.forEach((item) => {
            totalPlanifTabla += item.sumPlanif;
            const formatEjecut = `${String(item.sumEjecut).padStart(2, '0')}/${String(item.sumPlanif).padStart(2, '0')}`;
            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td>${item.fuerza}</td>
                <td>${item.tipoOp}</td>
                <td class="text-center">${item.sumPlanif}</td>
                <td class="text-center"><strong>${formatEjecut}</strong></td>
                <td>${item.provincia}</td>
                <td>${item.canton}</td>
                <td>${item.parroquia}</td>
                <td class="text-center">${item.fecha ? item.fecha.toLocaleDateString() : 'S/F'}</td>
                <td>${item.resultados}</td>
                <td class="text-center"><strong>${item.sumPmp}</strong></td>
                <td class="text-center hide-on-print">
                    <button class="btn btn-sm btn-outline-warning p-1" onclick="openEditModal(${item.id})"><span class="material-icons" style="font-size:16px;">edit</span></button>
                    <button class="btn btn-sm btn-outline-danger p-1" onclick="deleteItem(${item.id})"><span class="material-icons" style="font-size:16px;">delete</span></button>
                </td>
            `;
            dom.tableBody.appendChild(tr);
        });
        if (document.getElementById('tableTotalPlanif')) document.getElementById('tableTotalPlanif').textContent = totalPlanifTabla;
    }

    const tP = filteredData.reduce((a, c) => a + c.sumPlanif, 0);
    const tE = filteredData.reduce((a, c) => a + c.sumEjecut, 0);
    const tPmp = filteredData.reduce((a, c) => a + c.sumPmp, 0);
    const perc = tP > 0 ? ((tE / tP) * 100).toFixed(1) : 0;

    dom.stats.total.textContent = tP;
    dom.stats.pmp.textContent = tPmp;
    dom.stats.effectiveness.textContent = `${tE} (${perc}%)`;

    updateChart();
    updateGeoChart();
    renderSummaryTable();
}

function renderSummaryTable() {
    const body = document.getElementById('summaryBody');
    const total = document.getElementById('summaryTotal');
    if (!body) return;
    const counts = {};
    let gT = 0;
    filteredData.forEach(item => {
        if (item.sumEjecut === 0) return;
        const t = String(item.tipoOp).toUpperCase().trim();
        counts[t] = (counts[t] || 0) + item.sumEjecut;
        gT += item.sumEjecut;
    });
    body.innerHTML = '';
    Object.keys(counts).sort().forEach(t => {
        const tr = document.createElement('tr');
        tr.innerHTML = `<td>${t}</td><td class="text-end"><strong>${counts[t]}</strong></td>`;
        body.appendChild(tr);
    });
    if (total) total.textContent = gT;
}

function updateGeoChart() {
    const ctx = document.getElementById('geoChart').getContext('2d');
    const stats = {};
    const allTypes = new Set();
    
    filteredData.forEach(item => {
        if (item.sumEjecut === 0) return;
        const c = (item.canton || "S/J").toUpperCase().trim();
        const t = (item.tipoOp || "S/T").toUpperCase().trim();
        if (!stats[c]) stats[c] = {};
        stats[c][t] = (stats[c][t] || 0) + item.sumEjecut;
        allTypes.add(t);
    });

    const labels = Object.keys(stats).sort();
    const sortedTypes = Array.from(allTypes).sort();
    
    const datasets = sortedTypes.map((type, i) => ({
        label: type,
        data: labels.map(canton => stats[canton][type] || 0),
        backgroundColor: corporatePalette[i % corporatePalette.length],
        stacked: true,
        borderRadius: 4
    }));

    if (geoChartInstance) geoChartInstance.destroy();
    geoChartInstance = new Chart(ctx, {
        type: 'bar',
        data: { labels, datasets },
        options: {
            indexAxis: 'y',
            responsive: true,
            maintainAspectRatio: false,
            scales: {
                x: { stacked: true, grid: { display: false } },
                y: { stacked: true, grid: { display: false } }
            },
            plugins: {
                legend: { position: 'bottom', labels: { boxWidth: 10, usePointStyle: true } }
            },
            onClick: (e, el) => {
                if (el.length > 0) {
                    const cS = labels[el[0].index];
                    const sel = document.getElementById('filterCanton');
                    for(let i=0; i<sel.options.length; i++) {
                        if(sel.options[i].value.toUpperCase() === cS) { sel.selectedIndex = i; applyFilters(); break; }
                    }
                }
            }
        }
    });
}

function updateChart() {
    const ctx = document.getElementById('opsChart').getContext('2d');
    const typeSeries = {};
    filteredData.forEach(item => {
        if (item.sumEjecut === 0) return;
        const t = (item.tipoOp || "S/C").toUpperCase().trim();
        let sH = parseInt(item.horaMilitar.replace(/[^0-9]/g, '').substring(0,2)) || 0;
        if (!typeSeries[t]) typeSeries[t] = new Array(24).fill(0);
        typeSeries[t][sH] += item.sumEjecut;
    });

    const datasets = Object.keys(typeSeries).sort().map((t, i) => ({
        label: t,
        data: typeSeries[t],
        backgroundColor: corporatePalette[i % corporatePalette.length],
        borderColor: 'white',
        borderWidth: 1,
        stacked: true,
        borderRadius: 2
    }));

    if (chartInstance) chartInstance.destroy();
    chartInstance = new Chart(ctx, {
        type: 'bar',
        data: { labels: Array.from({length:24}, (_,i)=>String(i).padStart(2,'0')+":00"), datasets },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            scales: {
                x: { stacked: true, grid: { display: false } },
                y: { stacked: true, grid: { color: '#f3f2f1' } }
            },
            plugins: {
                legend: { position: 'bottom', labels: { boxWidth: 10, usePointStyle: true } }
            }
        }
    });
}

dom.btnExport.addEventListener('click', () => {
    try {
        if (filteredData.length === 0) {
            alert("No hay datos para exportar.");
            return;
        }

        // Mapear los datos al formato de Excel con los nombres de columnas correctos
        const exportData = filteredData.map(item => ({
            "FUERZA": item.fuerza,
            "TIPO DE OP.": item.tipoOp,
            "PLANIF.": item.sumPlanif,
            "EJECUT.": `${item.sumEjecut}/${item.sumPlanif}`,
            "PROVINCIA": item.provincia,
            "CANTÓN": item.canton,
            "PARROQUIA": item.parroquia,
            "FECHA": item.fecha ? item.fecha.toLocaleDateString() : 'S/F',
            "HORA": item.horaMilitar,
            "RESULTADOS": item.resultados,
            "PMP EMPLEADO": item.sumPmp
        }));

        const headerOrder = [
            "FUERZA", "TIPO DE OP.", "PLANIF.", "EJECUT.", 
            "PROVINCIA", "CANTÓN", "PARROQUIA", "FECHA", "HORA", "RESULTADOS", "PMP EMPLEADO"
        ];

        const ws = XLSX.utils.json_to_sheet(exportData, { header: headerOrder });
        
        // Ajustar ancho de columnas automáticamente
        const wscols = headerOrder.map(h => ({ wch: Math.max(h.length + 5, 15) }));
        ws['!cols'] = wscols;

        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "Reporte Operativo");
        XLSX.writeFile(wb, `Reporte_Operaciones_${new Date().toISOString().split('T')[0]}.xlsx`);

    } catch (error) {
        console.error("Error exportando:", error);
        alert("Error al exportar: " + error.message);
    }
});

dom.btnMessage.addEventListener('click', () => {
    if (rawData.length === 0) {
        alert("Cargue un archivo primero.");
        return;
    }

    let defaultStart = dom.filterStart.value;
    let defaultEnd = dom.filterEnd.value;

    // Si no hay filtros seleccionados, detectar rango del archivo Excel (rawData)
    if ((!defaultStart || !defaultEnd) && rawData.length > 0) {
        let minT = Infinity;
        let maxT = -Infinity;
        let hasDates = false;

        rawData.forEach(d => {
            if (d.fecha && !isNaN(d.fecha.getTime())) {
                const t = d.fecha.getTime();
                if (t < minT) minT = t;
                if (t > maxT) maxT = t;
                hasDates = true;
            }
        });

        if (hasDates) {
            if (!defaultStart) defaultStart = new Date(minT).toISOString().split('T')[0];
            if (!defaultEnd) defaultEnd = new Date(maxT).toISOString().split('T')[0];
        } else {
            // Fallback si no se detectan fechas válidas
            const today = new Date().toISOString().split('T')[0];
            if (!defaultStart) defaultStart = today;
            if (!defaultEnd) defaultEnd = today;
        }
    }

    document.getElementById('reportStartDate').value = defaultStart;
    document.getElementById('reportStartTime').value = dom.filterTimeStart.value || "00:00";
    document.getElementById('reportEndDate').value = defaultEnd;
    document.getElementById('reportEndTime').value = dom.filterTimeEnd.value || "23:59";

    // Limpiar área de texto anterior
    document.getElementById('generatedMessage').value = "Seleccione el rango de fechas y haga clic en 'Actualizar Reporte'.";
    
    const modalEl = document.getElementById('messageModal');
    const modal = new bootstrap.Modal(modalEl);
    modal.show();
});

window.generateCustomReport = function() {
    const startDateStr = document.getElementById('reportStartDate').value;
    const startTimeStr = document.getElementById('reportStartTime').value || "00:00";
    const endDateStr = document.getElementById('reportEndDate').value;
    const endTimeStr = document.getElementById('reportEndTime').value || "23:59";

    if (!startDateStr || !endDateStr) {
        alert("Por favor seleccione fecha de inicio y fin.");
        return;
    }

    const reportStart = new Date(`${startDateStr}T${startTimeStr}`);
    const reportEnd = new Date(`${endDateStr}T${endTimeStr}`);

    // Filtrar rawData basado en el rango específico del reporte
    // Nota: Usamos la misma lógica de "Operación Ejecutada" (validTime) que en el filtro principal
    // pero aplicada al rango específico del reporte.
    
    // 1. Filtrar items relevantes por fecha amplia (día completo)
    const rangeItems = rawData.filter(item => {
        if (!item.startDate || !item.endDate) return false;
        // Modificado a intersección de fechas para no excluir operaciones que cruzan límites de días o son de larga duración
        const dStart = new Date(reportStart); dStart.setHours(0,0,0,0);
        const dEnd = new Date(reportEnd); dEnd.setHours(23,59,59,999);
        return item.startDate <= dEnd && item.endDate >= dStart;
    });

    // 2. Agrupar y Calcular (Lógica similar a applyFilters pero solo para el reporte)
    const groups = {};
    
    rangeItems.forEach(item => {
        // Verificar si cuenta como ejecutada en este rango (INTERSECCIÓN / TRASLAPE)
        // Se considera ejecutada si existe cualquier solapamiento de tiempo entre la operación y el rango del reporte.
        // CAMBIO: Uso de <= y >= para ser inclusivo con los bordes (ej: op termina 17:00, reporte inicia 17:00)
        const timeOverlap = (item.startDate <= reportEnd && item.endDate >= reportStart);
        
        // Excepción 1: ARS 229-230 en ANEGADO
        const isArsSpecific = String(item.tipoOp).toUpperCase().includes("229-230") && String(item.parroquia).toUpperCase().includes("ANEGADO");
        
        // Excepción 2: Operaciones de 24 horas o tiempo indefinido (se asumen ejecutadas en el día)
        const cleanHora = String(item.horaMilitar).replace(/[^0-9]/g, '');
        const is24Hours = cleanHora === "00002359" || cleanHora === "00002400" || cleanHora === "00000000";

        const isEjecutada = timeOverlap || isArsSpecific || is24Hours;

        const dateKey = item.fecha ? item.fecha.toISOString().split('T')[0] : 'SIN_FECHA';
        const key = `${item.tipoOp}|${item.provincia}|${item.canton}|${item.parroquia}|${dateKey}`.toUpperCase();

        if (!groups[key]) {
            groups[key] = { 
                ...item, 
                sumPlanif: 0, sumEjecut: 0, sumPmp: 0, resultsList: [],
                sumOfi: 0, sumAerot: 0, sumRes: 0,
                sumMedios: { camioneta: 0, camion: 0, bus: 0 }
            };
        }

        groups[key].sumPlanif += 1; // Siempre cuenta como planificada si está en los días del rango

        if (isEjecutada) {
            groups[key].sumEjecut += 1;
            groups[key].sumPmp += (item.pmp || 0);

            if (item.detPmp) {
                groups[key].sumOfi += (item.detPmp.ofi || 0);
                groups[key].sumAerot += (item.detPmp.aerot || 0);
                groups[key].sumRes += (item.detPmp.res || 0);
            }

            if (item.medios) {
                groups[key].sumMedios.camioneta += (item.medios.camioneta || 0);
                groups[key].sumMedios.camion += (item.medios.camion || 0);
                groups[key].sumMedios.bus += (item.medios.bus || 0);
            }

            if (item.resultados && item.resultados !== "0" && item.resultados !== "") {
                groups[key].resultsList.push(item.resultados);
            }
        }
    });

    // Consolidar resultados
    const reportData = Object.values(groups).map(g => {
        if (g.resultsList.length > 0) g.resultados = [...new Set(g.resultsList)].join(" / ");
        else g.resultados = "0";
        return g;
    });

    // Ordenar (Fecha -> Cantón)
    reportData.sort((a, b) => {
        const dateA = a.fecha ? a.fecha.getTime() : 0;
        const dateB = b.fecha ? b.fecha.getTime() : 0;
        if (dateA !== dateB) return dateA - dateB;
        return String(a.canton).localeCompare(String(b.canton));
    });

    if (reportData.length === 0) {
        document.getElementById('generatedMessage').value = "No se encontraron operaciones en el rango especificado.";
        return;
    }

    // Generar Texto
    let header = "";
    const h = reportStart.getHours();

    // Lógica de Saludo Protocolario basada en el turno (Hora de Inicio)
    // 17:00 - 05:00 -> Turno Noche -> Reporte "Buenos Días"
    // 05:00 - 17:00 -> Turno Día -> Reporte "Buenas Tardes"
    if (h >= 17 || h < 5) {
        header = "*BUENOS DÍAS MI GENERAL, ME PERMITO DAR PARTE DE LAS OPERACIONES EJECUTADAS POR EL GOMAI “MANABÍ” EN LOS CANTONES:*\n";
    } else {
        header = "*BUENAS TARDES MI GENERAL, ME PERMITO DAR PARTE DE LAS OPERACIONES EJECUTADAS POR EL GOMAI “MANABÍ” EN LOS CANTONES:*\n";
    }

    // Listado de cantones únicos con ejecuciones
    const cantonesEjecutados = [...new Set(reportData.filter(op => op.sumEjecut > 0).map(op => String(op.canton).toUpperCase().trim()))].sort();
    
    let message = header;
    cantonesEjecutados.forEach(c => {
        message += `*-${c}*\n`;
    });
    message += "\n";

    // --- BLOQUE ARS (PROTECCIÓN DE ZONAS) ---
    const arsItems = reportData.filter(i => (String(i.tipoOp).includes("ARS") || String(i.tipoOp).includes("AREAS RESERVADAS")) && i.sumEjecut > 0);
    const totalArs = arsItems.reduce((acc, curr) => acc + curr.sumEjecut, 0);

    if (totalArs > 0) {
        const formatTotal = String(totalArs).padStart(2, '0');
        
        message += `*PROTECCIÓN DE LAS ZONAS DE SEGURIDAD DEL ESTADO*\n\n`;
        message += `*A) PROTECCIÓN DE LAS ÁREAS RESERVADAS DE SEGURIDAD TERRESTRE ARS (${formatTotal})*\n`;
        message += `*1. ARS: (${formatTotal})*\n`;
        
        // Cantones ARS
        message += `*CANTÓN:*\n`;
        const cSet = new Set(arsItems.map(i => i.canton));
        [...cSet].sort().forEach(c => message += `-${c}\n`);

        // Sectores ARS
        message += `*SECTOR:*\n`;
        const sSet = new Set(arsItems.map(i => i.parroquia));
        [...sSet].sort().forEach(s => message += `-${s}\n`);

        // Personal ARS
        message += `*PERSONAL EMPLEADO:*\n`;
        const tOfi = arsItems.reduce((a,b) => a + b.sumOfi, 0);
        const tAerot = arsItems.reduce((a,b) => a + b.sumAerot, 0);
        const tRes = arsItems.reduce((a,b) => a + b.sumRes, 0);

        if (tOfi > 0) message += tOfi === 1 ? `Oficial: ${tOfi}\n` : `Oficiales: ${tOfi}\n`;
        if (tAerot > 0) message += `Aerotécnicos: ${tAerot}\n`;
        if (tRes > 0) message += `Reservistas: ${tRes}\n`;

        // Medios ARS
        const tCamioneta = arsItems.reduce((a,b) => a + b.sumMedios.camioneta, 0);
        const tCamion = arsItems.reduce((a,b) => a + b.sumMedios.camion, 0);
        const tBus = arsItems.reduce((a,b) => a + b.sumMedios.bus, 0);

        let mediosStr = "";
        if (tCamioneta > 0) mediosStr += `Camionetas: ${String(tCamioneta).padStart(2,'0')}\n`;
        if (tCamion > 0) mediosStr += `Camiones: ${String(tCamion).padStart(2,'0')}\n`;
        if (tBus > 0) mediosStr += `Buses: ${String(tBus).padStart(2,'0')}\n`;

        if (mediosStr) {
            message += `*MEDIOS:*\n${mediosStr}`;
        }
        message += "\n";
    }

    // --- BLOQUE B) COMPETENCIA LEGAL DE FUERZAS ARMADAS ---
    const filterOp = (k) => reportData.filter(i => String(i.tipoOp).includes(k) && i.sumEjecut > 0);
    
    // Categorías Específicas
    const opRastrillaje = filterOp("RASTRILLAJE");
    const opCombate = filterOp("COMBATE URBANO");
    const opViales = filterOp("VIALES"); // "CONTROL DE EJES VIALES" o "EJES VIALES"
    // Separar CAMEX normal de CAMEX P.N. si es necesario, o agrupar por string
    const opCamex = reportData.filter(i => String(i.tipoOp).includes("CAMEX") && !String(i.tipoOp).includes("P.N") && !String(i.tipoOp).includes("POLICIA") && i.sumEjecut > 0);
    const opCamexPn = reportData.filter(i => String(i.tipoOp).includes("CAMEX") && (String(i.tipoOp).includes("P.N") || String(i.tipoOp).includes("POLICIA")) && i.sumEjecut > 0);

    const totalB = [opRastrillaje, opCombate, opViales, opCamex, opCamexPn].flat().reduce((a, b) => a + b.sumEjecut, 0);

    if (totalB > 0) {
        message += `*B) COMPETENCIA LEGAL DE FUERZAS ARMADAS (${String(totalB).padStart(2, '0')})*\n\n`;

        const formatBlock = (title, index, data) => {
            if (data.length === 0) return "";
            
            const tCount = data.reduce((a, b) => a + b.sumEjecut, 0);
            let txt = `*${index}. ${title} (${String(tCount).padStart(2, '0')})*\n`;

            // Agrupación Canton -> Sectores (Parroquias)
            const cantonMap = {};
            data.forEach(d => {
                const c = String(d.canton).trim();
                if (!cantonMap[c]) cantonMap[c] = new Set();
                // Usamos parroquia como sector. Si está vacía, usar 'S/N'
                cantonMap[c].add(d.parroquia || "S/N");
            });

            txt += `*CANTÓN:*\n`;
            Object.keys(cantonMap).sort().forEach(c => {
                txt += `*${c}*\n`; // Cantón en una línea
                // Sectores en la SIGUIENTE línea separados por /
                const sectores = [...cantonMap[c]].filter(s => s).join(" / ");
                txt += `${sectores}\n`; 
            });

            // Personal
            txt += `*PERSONAL EMPLEADO:*\n`;
            const tOfi = data.reduce((a,b) => a + b.sumOfi, 0);
            const tAerot = data.reduce((a,b) => a + b.sumAerot, 0);
            const tRes = data.reduce((a,b) => a + b.sumRes, 0);

            if (tOfi > 0) txt += tOfi === 1 ? `Oficial: ${tOfi}\n` : `Oficiales: ${tOfi}\n`;
            if (tAerot > 0) txt += `Aerotécnicos: ${tAerot}\n`;
            if (tRes > 0) txt += `Reservistas: ${tRes}\n`;

            // Medios
            const med = {
                camioneta: data.reduce((a,b) => a + b.sumMedios.camioneta, 0),
                camion: data.reduce((a,b) => a + b.sumMedios.camion, 0),
                bus: data.reduce((a,b) => a + b.sumMedios.bus, 0)
            };

            let mTxt = "";
            if (med.camioneta > 0) mTxt += `Camionetas: ${String(med.camioneta).padStart(2,'0')}\n`;
            if (med.camion > 0) mTxt += `Camiones: ${String(med.camion).padStart(2,'0')}\n`;
            if (med.bus > 0) mTxt += `Buses: ${String(med.bus).padStart(2,'0')}\n`;

            if (mTxt) txt += `*MEDIOS:*\n${mTxt}`;
            
            return txt + "\n";
        };

        let idx = 1;
        message += formatBlock("RASTRILLAJE", idx++, opRastrillaje);
        message += formatBlock("COMBATE URBANO", idx++, opCombate);
        message += formatBlock("EJES VIALES", idx++, opViales);
        message += formatBlock("CAMEX", idx++, opCamex);
        message += formatBlock("CAMEX COORD. P.N.", idx++, opCamexPn);
    }

    // --- BLOQUE C) APOYO A OTRAS ENTIDADES DEL ESTADO SIN EE ---
    const opMineduc = reportData.filter(i => String(i.tipoOp).includes("MINEDUC") && i.sumEjecut > 0);
    const totalC = opMineduc.reduce((a, b) => a + b.sumEjecut, 0);

    if (totalC > 0) {
        message += `*C) APOYO A OTRAS ENTIDADES DEL ESTADO SIN EE (${String(totalC).padStart(2, '0')})*\n\n`;
        
        // Reutilizamos la lógica de formateo manual para este bloque específico
        const tCount = String(totalC).padStart(2, '0');
        message += `*1. APOYO AL MINEDUC (${tCount})*\n`;

        // Agrupación Canton -> Sectores
        const cantonMap = {};
        opMineduc.forEach(d => {
            const c = String(d.canton).trim();
            if (!cantonMap[c]) cantonMap[c] = new Set();
            cantonMap[c].add(d.parroquia || "S/N");
        });

        message += `*CANTÓN:*\n`;
        Object.keys(cantonMap).sort().forEach(c => {
            message += `*${c}*\n`;
            const sectores = [...cantonMap[c]].filter(s => s).join(" / ");
            message += `${sectores}\n`;
        });

        // Personal
        message += `*PERSONAL EMPLEADO:*\n`;
        const tOfi = opMineduc.reduce((a,b) => a + b.sumOfi, 0);
        const tAerot = opMineduc.reduce((a,b) => a + b.sumAerot, 0);
        const tRes = opMineduc.reduce((a,b) => a + b.sumRes, 0);

        if (tOfi > 0) message += tOfi === 1 ? `Oficial: ${tOfi}\n` : `Oficiales: ${tOfi}\n`;
        if (tAerot > 0) message += `Aerotécnicos: ${tAerot}\n`;
        if (tRes > 0) message += `Reservistas: ${tRes}\n`;

        // Medios
        const med = {
            camioneta: opMineduc.reduce((a,b) => a + b.sumMedios.camioneta, 0),
            camion: opMineduc.reduce((a,b) => a + b.sumMedios.camion, 0),
            bus: opMineduc.reduce((a,b) => a + b.sumMedios.bus, 0)
        };

        let mTxt = "";
        if (med.camioneta > 0) mTxt += `Camionetas: ${String(med.camioneta).padStart(2,'0')}\n`;
        if (med.camion > 0) mTxt += `Camiones: ${String(med.camion).padStart(2,'0')}\n`;
        if (med.bus > 0) mTxt += `Buses: ${String(med.bus).padStart(2,'0')}\n`;

        if (mTxt) message += `*MEDIOS:*\n${mTxt}`;
        message += "\n";
    }

    // --- BLOQUE D) PROTECCIÓN ALTAS AUTORIDADES NAC, VISIT INTERNACIONALES ---
    const opVip = reportData.filter(i => (
        String(i.tipoOp).includes("AUTORIDADES") || 
        String(i.tipoOp).includes("VISIT") || 
        String(i.tipoOp).includes("SEGURIDAD DE VUELO") || 
        String(i.tipoOp).includes("VIP") ||
        String(i.tipoOp).includes("PROTEC Y SEG PMI") ||
        String(i.tipoOp).includes("FUNCIONARIOS")
    ) && i.sumEjecut > 0);
    const totalVip = opVip.reduce((a, b) => a + b.sumEjecut, 0);

    if (totalVip > 0) {
        message += `*D) PROTECCIÓN ALTAS AUTORIDADES NAC, VISIT INTERNACIONALES (${String(totalVip).padStart(2, '0')})*\n\n`;

        const vipTypes = {};
        opVip.forEach(item => {
            const t = item.tipoOp;
            if (!vipTypes[t]) vipTypes[t] = [];
            vipTypes[t].push(item);
        });

        let idx = 1;
        Object.keys(vipTypes).sort().forEach(type => {
            const data = vipTypes[type];
            const tCount = data.reduce((a, b) => a + b.sumEjecut, 0);
            
            message += `*${idx++}. ${type} (${String(tCount).padStart(2, '0')})*\n`;

            const cantonMap = {};
            data.forEach(d => {
                const c = String(d.canton).trim();
                if (!cantonMap[c]) cantonMap[c] = new Set();
                cantonMap[c].add(d.parroquia || "S/N");
            });

            message += `*CANTÓN:*\n`;
            Object.keys(cantonMap).sort().forEach(c => {
                message += `*${c}*\n`;
                const sectores = [...cantonMap[c]].filter(s => s).join(" / ");
                message += `${sectores}\n`; 
            });

            message += `*PERSONAL EMPLEADO:*\n`;
            const tOfi = data.reduce((a,b) => a + b.sumOfi, 0);
            const tAerot = data.reduce((a,b) => a + b.sumAerot, 0);
            const tRes = data.reduce((a,b) => a + b.sumRes, 0);

            if (tOfi > 0) message += tOfi === 1 ? `Oficial: ${tOfi}\n` : `Oficiales: ${tOfi}\n`;
            if (tAerot > 0) message += `Aerotécnicos: ${tAerot}\n`;
            if (tRes > 0) message += `Reservistas: ${tRes}\n`;

            const med = {
                camioneta: data.reduce((a,b) => a + b.sumMedios.camioneta, 0),
                camion: data.reduce((a,b) => a + b.sumMedios.camion, 0),
                bus: data.reduce((a,b) => a + b.sumMedios.bus, 0)
            };

            let mTxt = "";
            if (med.camioneta > 0) mTxt += `Camionetas: ${String(med.camioneta).padStart(2,'0')}\n`;
            if (med.camion > 0) mTxt += `Camiones: ${String(med.camion).padStart(2,'0')}\n`;
            if (med.bus > 0) mTxt += `Buses: ${String(med.bus).padStart(2,'0')}\n`;

            if (mTxt) message += `*MEDIOS:*\n${mTxt}`;
            message += "\n";
        });
    }

    // --- BLOQUE TOTALES DE OPERACIONES ---
    const grandTotal = (typeof totalArs !== 'undefined' ? totalArs : 0) + (typeof totalB !== 'undefined' ? totalB : 0) + (typeof totalC !== 'undefined' ? totalC : 0) + totalVip;

    if (grandTotal > 0) {
        message += `*TOTAL DE OPERACIONES (${String(grandTotal).padStart(2, '0')})*\n\n`;

        if (typeof totalArs !== 'undefined' && totalArs > 0) {
            message += `(${String(totalArs).padStart(2, '0')}) PROTECCIÓN DE LAS ÁREAS RESERVADAS DE SEGURIDAD TERRESTRE ARS\n`;
        }
        if (typeof totalB !== 'undefined' && totalB > 0) {
            message += `(${String(totalB).padStart(2, '0')}) COMPETENCIA LEGAL DE FUERZAS ARMADAS\n`;
        }
        if (typeof totalC !== 'undefined' && totalC > 0) {
            message += `(${String(totalC).padStart(2, '0')}) APOYO A OTRAS ENTIDADES DEL ESTADO SIN EE\n`;
        }
        if (totalVip > 0) {
            message += `(${String(totalVip).padStart(2, '0')}) PROTECCIÓN ALTAS AUTORIDADES NAC, VISIT INTERNACIONALES\n`;
        }
        message += "\n";
    }

    // --- BLOQUE FINAL: TOTALES DE PERSONAL Y MEDIOS ---
    const finalOfi = reportData.reduce((a, b) => a + (b.sumOfi || 0), 0);
    const finalAerot = reportData.reduce((a, b) => a + (b.sumAerot || 0), 0);
    const finalRes = reportData.reduce((a, b) => a + (b.sumRes || 0), 0);

    const finalMed = {
        camioneta: reportData.reduce((a, b) => a + (b.sumMedios.camioneta || 0), 0),
        camion: reportData.reduce((a, b) => a + (b.sumMedios.camion || 0), 0),
        bus: reportData.reduce((a, b) => a + (b.sumMedios.bus || 0), 0)
    };

    message += `*TOTAL, PERSONAL MILITAR PROFESIONAL:*\n`;
    if (finalOfi > 0) message += finalOfi === 1 ? `- Oficial: ${finalOfi}\n` : `- Oficiales: ${finalOfi}\n`;
    if (finalAerot > 0) message += `- Aerotécnicos: ${finalAerot}\n`;
    if (finalRes > 0) message += `- Reservistas: ${finalRes}\n`;

    message += `\n*MEDIOS:*\n`;
    let hasMedios = false;
    if (finalMed.camioneta > 0) { message += `- Camionetas: ${String(finalMed.camioneta).padStart(2, '0')}\n`; hasMedios = true; }
    if (finalMed.camion > 0) { message += `- Camiones: ${String(finalMed.camion).padStart(2, '0')}\n`; hasMedios = true; }
    if (finalMed.bus > 0) { message += `- Buses: ${String(finalMed.bus).padStart(2, '0')}\n`; hasMedios = true; }
    
    document.getElementById('generatedMessage').value = message.toUpperCase();
};

window.copyMessage = function() {
    const txt = document.getElementById('generatedMessage');
    txt.select();
    document.execCommand('copy');
    alert("Reporte copiado al portapapeles.");
};

let editModalInstance = null;
window.openEditModal = function(id) {
    const groupItem = filteredData.find(i => i.id === id);
    if (!groupItem || !groupItem.originalOps) return;

    document.getElementById('editId').value = id;
    document.getElementById('editTipo').value = groupItem.tipoOp;

    const container = document.getElementById('opsEditContainer');
    container.innerHTML = '<h6 class="mb-3 text-primary border-bottom pb-2">Registros por Horario</h6>';

    groupItem.originalOps.forEach((op, index) => {
        const dateVal = op.fecha ? new Date(op.fecha).toISOString().split('T')[0] : "";
        const div = document.createElement('div');
        div.className = "card mb-3 p-3 shadow-sm border-start border-primary border-4 bg-light";
        div.innerHTML = `
            <div class="row g-2">
                <div class="col-12 d-flex justify-content-between">
                    <span class="badge bg-primary">REGISTRO #${index + 1}</span>
                    <small class="text-muted">ID: ${op.id}</small>
                </div>
                <div class="col-md-6">
                    <label class="small fw-bold d-block">Fecha</label>
                    <input type="date" class="form-control form-control-sm edit-op-date" data-opid="${op.id}" value="${dateVal}">
                </div>
                <div class="col-md-6">
                    <label class="small fw-bold d-block">Rango Horario</label>
                    <input type="text" class="form-control form-control-sm edit-op-time" data-opid="${op.id}" value="${op.horaMilitar}" placeholder="HHMM-HHMM">
                </div>
                <div class="col-md-9">
                    <label class="small fw-bold d-block">Resultados de este Horario</label>
                    <textarea class="form-control form-control-sm edit-op-results" data-opid="${op.id}" rows="2">${op.resultados || ""}</textarea>
                </div>
                <div class="col-md-3">
                    <label class="small fw-bold d-block">PMP</label>
                    <input type="number" class="form-control form-control-sm edit-op-pmp" data-opid="${op.id}" value="${op.pmp || 0}">
                </div>
            </div>
        `;
        container.appendChild(div);
    });

    if (!editModalInstance) {
        editModalInstance = new bootstrap.Modal(document.getElementById('editModal'));
    }
    editModalInstance.show();
};

window.saveChanges = function() {
    const id = parseInt(document.getElementById('editId').value);
    const groupItem = filteredData.find(i => i.id === id);
    if (!groupItem) return;

    const newTipo = document.getElementById('editTipo').value.toUpperCase().trim();

    // Capturar todos los inputs de las sub-operaciones
    const dateInputs = document.querySelectorAll('.edit-op-date');
    const timeInputs = document.querySelectorAll('.edit-op-time');
    const resultsInputs = document.querySelectorAll('.edit-op-results');
    const pmpInputs = document.querySelectorAll('.edit-op-pmp');

    dateInputs.forEach((input, i) => {
        const opId = parseInt(input.dataset.opid);
        const rawIdx = rawData.findIndex(item => item.id === opId);
        
        if (rawIdx !== -1) {
            // Actualizar Tipo (común a todo el grupo editado)
            rawData[rawIdx].tipoOp = newTipo;
            
            // Actualizar Fecha
            if(input.value) {
                const [year, month, day] = input.value.split('-');
                rawData[rawIdx].fecha = new Date(year, month - 1, day);
            }

            // Actualizar Hora
            const newTime = timeInputs[i].value;
            rawData[rawIdx].horaMilitar = newTime;
            
            // Actualizar Resultados y PMP específicos
            rawData[rawIdx].resultados = resultsInputs[i].value.trim();
            rawData[rawIdx].pmp = parseInt(pmpInputs[i].value) || 0;
            
            // Reprocesar startDate/endDate
            const parts = newTime.split('-').map(p => p.trim().replace(/[^0-9]/g, ''));
            const hIni = parts[0].padStart(4, '0');
            const hFin = (parts[1] || parts[0]).padStart(4, '0');
            
            if (rawData[rawIdx].fecha) {
                const d = rawData[rawIdx].fecha;
                const start = new Date(d);
                start.setHours(parseInt(hIni.substring(0,2)) || 0, parseInt(hIni.substring(2,4)) || 0, 0);
                const end = new Date(d);
                end.setHours(parseInt(hFin.substring(0,2)) || 0, parseInt(hFin.substring(2,4)) || 0, 0);
                if (end < start) end.setDate(end.getDate() + 1);
                rawData[rawIdx].startDate = start;
                rawData[rawIdx].endDate = end;
            }
        }
    });

    applyFilters();
    if(editModalInstance) editModalInstance.hide();
};

window.renderReportCrud = function() {
    const startDateStr = document.getElementById('reportStartDate').value;
    const startTimeStr = document.getElementById('reportStartTime').value || "00:00";
    const endDateStr = document.getElementById('reportEndDate').value;
    const endTimeStr = document.getElementById('reportEndTime').value || "23:59";

    if (!startDateStr || !endDateStr) return;

    const rStart = new Date(`${startDateStr}T${startTimeStr}`);
    const rEnd = new Date(`${endDateStr}T${endTimeStr}`);
    
    // Filtro idéntico al del reporte (Lógica unificada)
    let items = rawData.filter(item => {
        if (!item.startDate || !item.endDate) return false;
        
        // 1. Verificar solapamiento de tiempo (Intersección)
        const timeOverlap = (item.startDate <= rEnd && item.endDate >= rStart);
        
        // 2. Excepción ARS Anegado
        const isArsSpecific = String(item.tipoOp).toUpperCase().includes("229-230") && String(item.parroquia).toUpperCase().includes("ANEGADO");
        
        // 3. Excepción 24 Horas
        const cleanHora = String(item.horaMilitar).replace(/[^0-9]/g, '');
        const is24Hours = cleanHora === "00002359" || cleanHora === "00002400" || cleanHora === "00000000";

        return timeOverlap || isArsSpecific || is24Hours;
    });

    // --- LÓGICA DE FILTRO POR TIPO ---
    const filterSelect = document.getElementById('crudFilterTipo');
    const currentFilter = filterSelect.value;
    
    // 1. Obtener tipos únicos presentes en el rango actual
    const uniqueTypes = [...new Set(items.map(i => i.tipoOp))].sort();
    
    // 2. Repoblar el select (manteniendo la selección si es posible)
    // Guardamos las opciones actuales para no redibujar si no cambia (opcional, pero aquí redibujamos simple)
    filterSelect.innerHTML = '<option value="TODOS">-- Todos los Tipos --</option>';
    uniqueTypes.forEach(t => {
        const opt = document.createElement('option');
        opt.value = t;
        opt.textContent = t;
        if (t === currentFilter) opt.selected = true;
        filterSelect.appendChild(opt);
    });

    // 3. Filtrar la lista si hay selección
    if (currentFilter && currentFilter !== "TODOS") {
        items = items.filter(i => i.tipoOp === currentFilter);
    }

    const tbody = document.getElementById('crudTableBody');
    tbody.innerHTML = '';

    items.sort((a,b) => (b.startDate || 0) - (a.startDate || 0));

    items.forEach(item => {
        const tr = document.createElement('tr');
        
        // Calcular PMP localmente si existe detalle
        const ofi = (item.detPmp ? item.detPmp.ofi : 0) || 0;
        const aerot = (item.detPmp ? item.detPmp.aerot : 0) || 0;
        const res = (item.detPmp ? item.detPmp.res : 0) || 0;
        const pmpStr = `Of:${ofi} Ae:${aerot} Res:${res}`;

        // Calcular Medios
        const m = item.medios || {};
        const medArr = [];
        if(m.camioneta) medArr.push(`Camioneta:${m.camioneta}`);
        if(m.camion) medArr.push(`Camion:${m.camion}`);
        if(m.bus) medArr.push(`Bus:${m.bus}`);
        const medStr = medArr.length > 0 ? medArr.join(', ') : '-';

        tr.innerHTML = `
            <td>${item.fecha ? item.fecha.toLocaleDateString() : ''}<br><small>${item.horaMilitar}</small></td>
            <td>${item.canton}<br><small class="text-muted">${item.parroquia}</small></td>
            <td>${item.tipoOp}</td>
            <td><small>${pmpStr}</small></td>
            <td><small>${medStr}</small></td>
            <td class="text-center">
                <button class="btn btn-outline-primary btn-sm p-0 px-1" onclick="openCrudForm(${item.id})"><span class="material-icons" style="font-size:16px">edit</span></button>
                <button class="btn btn-outline-danger btn-sm p-0 px-1" onclick="deleteCrudItem(${item.id})"><span class="material-icons" style="font-size:16px">delete</span></button>
            </td>
        `;
        tbody.appendChild(tr);
    });
};

window.openCrudForm = function(id = null) {
    document.getElementById('crudFormContainer').classList.remove('d-none');
    document.getElementById('crudId').value = id !== null ? id : '';
    document.getElementById('crudFormTitle').textContent = id !== null ? 'Editar Operación' : 'Nueva Operación';

    if (id !== null) {
        const item = rawData.find(i => i.id === id);
        if (!item) return;
        
        document.getElementById('crudFecha').value = item.fecha ? item.fecha.toISOString().split('T')[0] : '';
        
        // Separar hora militar HHMM-HHMM para inputs HH:MM
        const parts = item.horaMilitar.split('-');
        const h1 = parts[0].trim();
        const h2 = parts[1] ? parts[1].trim() : h1;
        
        const fmtH = (str) => str.length >= 4 ? `${str.substring(0,2)}:${str.substring(2,4)}` : "00:00";
        document.getElementById('crudHoraIni').value = fmtH(h1);
        document.getElementById('crudHoraFin').value = fmtH(h2);

        document.getElementById('crudTipo').value = item.tipoOp;
        document.getElementById('crudCanton').value = item.canton;
        document.getElementById('crudParroquia').value = item.parroquia;
        document.getElementById('crudResultados').value = item.resultados;

        document.getElementById('crudOfi').value = item.detPmp ? item.detPmp.ofi : 0;
        document.getElementById('crudAerot').value = item.detPmp ? item.detPmp.aerot : 0;
        document.getElementById('crudRes').value = item.detPmp ? item.detPmp.res : 0;

        const m = item.medios || {};
        document.getElementById('crudCamioneta').value = m.camioneta || 0;
        document.getElementById('crudCamion').value = m.camion || 0;
        document.getElementById('crudBus').value = m.bus || 0;

    } else {
        // Limpiar para nuevo
        document.getElementById('crudId').value = '';
        document.getElementById('crudFecha').value = document.getElementById('reportStartDate').value;
        document.getElementById('crudHoraIni').value = "08:00";
        document.getElementById('crudHoraFin').value = "17:00";
        document.getElementById('crudTipo').value = "";
        document.getElementById('crudCanton').value = "";
        document.getElementById('crudParroquia').value = "";
        document.getElementById('crudResultados').value = "0";
        
        ['crudOfi','crudAerot','crudRes','crudCamioneta','crudCamion','crudBus'].forEach(id => document.getElementById(id).value = 0);
    }
};

window.closeCrudForm = function() {
    document.getElementById('crudFormContainer').classList.add('d-none');
};

window.saveCrudData = function() {
    const idVal = document.getElementById('crudId').value;
    const isNew = idVal === '';
    
    // Valores Básicos
    const fechaStr = document.getElementById('crudFecha').value;
    const horaIniStr = document.getElementById('crudHoraIni').value.replace(':',''); // HHMM
    const horaFinStr = document.getElementById('crudHoraFin').value.replace(':',''); // HHMM
    const tipo = document.getElementById('crudTipo').value.toUpperCase().trim();
    const canton = document.getElementById('crudCanton').value.toUpperCase().trim();
    const parr = document.getElementById('crudParroquia').value.toUpperCase().trim();
    const resTxt = document.getElementById('crudResultados').value.trim();

    // Valores PMP
    const ofi = parseInt(document.getElementById('crudOfi').value) || 0;
    const aerot = parseInt(document.getElementById('crudAerot').value) || 0;
    const res = parseInt(document.getElementById('crudRes').value) || 0;

    // Valores Medios
    const medios = {
        camioneta: parseInt(document.getElementById('crudCamioneta').value) || 0,
        camion: parseInt(document.getElementById('crudCamion').value) || 0,
        bus: parseInt(document.getElementById('crudBus').value) || 0
    };

    if (!fechaStr || !tipo || !canton) {
        alert("Complete Fecha, Tipo y Cantón");
        return;
    }

    // Construir Objeto
    let newItem = {};
    if (!isNew) {
        const idx = rawData.findIndex(i => i.id == idVal);
        if (idx === -1) return;
        newItem = rawData[idx];
    } else {
        newItem.id = Date.now(); // ID temporal único
        newItem.fuerza = "AÉREA";
        newItem.sumPlanif = 1; // Default
    }

    // Actualizar campos
    newItem.fecha = new Date(fechaStr + "T00:00:00");
    newItem.horaMilitar = `${horaIniStr} - ${horaFinStr}`;
    newItem.tipoOp = tipo;
    newItem.canton = canton;
    newItem.parroquia = parr;
    newItem.resultados = resTxt;
    
    newItem.pmp = ofi + aerot + res;
    newItem.detPmp = { ofi, aerot, res };
    newItem.medios = medios;

    // Recalcular Start/End Date para filtros
    const [y,m,d] = fechaStr.split('-').map(Number);
    const start = new Date(y, m-1, d);
    start.setHours(parseInt(horaIniStr.substring(0,2))||0, parseInt(horaIniStr.substring(2,4))||0, 0);
    const end = new Date(y, m-1, d);
    end.setHours(parseInt(horaFinStr.substring(0,2))||0, parseInt(horaFinStr.substring(2,4))||0, 0);
    if(end < start) end.setDate(end.getDate() + 1);

    newItem.startDate = start;
    newItem.endDate = end;

    if (isNew) {
        rawData.push(newItem);
    } 
    // Si era edit, ya se modificó la referencia en rawData

    closeCrudForm();
    renderReportCrud();
    applyFilters(); // Actualizar dashboard principal también
    alert("Guardado correctamente.");
};

window.deleteCrudItem = function(id) {
    if (confirm("¿Eliminar operación de la base de datos?")) {
        rawData = rawData.filter(i => i.id !== id);
        renderReportCrud();
        applyFilters();
    }
};

document.addEventListener('DOMContentLoaded', () => {
    initDraggableModal();
    // No action needed for tabs, bootstrap handles it
});