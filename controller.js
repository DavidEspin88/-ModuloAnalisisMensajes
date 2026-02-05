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
    filterProvincia: document.getElementById('filterProvincia'),
    filterCanton: document.getElementById('filterCanton'),
    filterTipo: document.getElementById('filterTipo'),
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
            workbook = XLSX.read(data, { type: 'array', cellDates: true });
            rawData = [];
            dom.sheetSelector.innerHTML = '<option value="ALL">-- Ver Todas las Hojas --</option>';
            workbook.SheetNames.forEach(name => {
                const opt = document.createElement('option');
                opt.value = name; opt.textContent = name;
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

// Event listener para el selector de Periodo (Día/Pestaña)
dom.sheetSelector.addEventListener('change', (e) => {
    const selectedSheet = e.target.value;
    console.log('Pestaña seleccionada:', selectedSheet);

    if (selectedSheet === 'ALL') {
        applyFilters();
    } else {
        // Encontrar la fecha de esta pestaña en rawData para sincronizar los filtros de fecha
        const firstItem = rawData.find(it => it.nombreHoja === selectedSheet);
        if (firstItem) {
            const sheetDate = firstItem.fechaPlanificacion.toISOString().split('T')[0];
            dom.filterStart.value = sheetDate;
            dom.filterEnd.value = sheetDate;
            // Resetear tiempos para ver todo lo planificado del día como ejecutado por defecto
            dom.filterTimeStart.value = "00:00";
            dom.filterTimeEnd.value = "23:59";
        }
        applyFiltersForSheet(selectedSheet);
    }
});

// --- MOTOR DE EXTRACCIÓN Y NORMALIZACIÓN ---
function extractDataFromSheet(worksheet, sheetName) {
    const matrix = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
    if (!matrix || matrix.length === 0) return [];

    const colMap = {
        ord: -1, fecha: -1, horaInicio: -1, horaFin: -1, tipoOp: -1, operaciones: -1,
        provincia: -1, canton: -1, parroquia: -1, sector: -1, resultados: -1, pmp: -1,
        ofi: -1, aerot: -1, res: -1  // ← ACTUALIZADO: Columnas de personal
    };

    // Identificar cabeceras (Soporte para multi-fila)
    let dataStartRow = 0;
    let lastHeaderRow = -1;

    for (let r = 0; r < Math.min(20, matrix.length); r++) {
        const row = matrix[r];
        if (!row) continue;

        const rowStr = row.map(c => (c !== null && c !== undefined) ? String(c).toUpperCase().trim() : "");
        const find = (keys) => rowStr.findIndex(c => c && keys.some(k => c.includes(k)));

        let foundInRow = false;

        // Búsqueda de columnas principales y anclas
        const check = (mapKey, keys) => {
            const idx = find(keys);
            if (idx !== -1) {
                if (colMap[mapKey] === -1) colMap[mapKey] = idx;
                foundInRow = true;
            }
        };

        check('ord', ['ORD.', 'NRO', 'NUM', 'ORD']);
        check('fecha', ['FECHA', 'DIA']);
        check('horaInicio', ['INICIO', 'H. INI', 'HORA']);
        check('horaFin', ['FIN', 'H. FIN', 'TERMINO']);
        check('tipoOp', ['TIPO DE OP', 'TIPO OP', 'ACTIVIDAD']);
        check('provincia', ['PROVINCIA', 'PROV']);
        check('canton', ['CANTÓN', 'CANTON', 'JURISDICCION']);
        check('parroquia', ['PARROQUIA']);
        check('sector', ['SECTOR']);
        check('resultados', ['RESULTADOS', 'NOVEDAD']);
        check('pmp', ['PMP', 'PERS', 'PERSONAL']);
        check('ofi', ['OFI', 'OFIC']);
        check('aerot', ['AEROT', 'AEROT.', 'TROPA']);
        check('res', ['RESV', 'RESER', 'RSV', 'RESVISTA']);

        // Ancla especial: COMANDANTE
        if (find(['COMANDANTE']) !== -1) foundInRow = true;

        if (foundInRow) {
            lastHeaderRow = r;
        }

        // Si ya tenemos lo básico y la fila actual no tiene keywords, probablemente terminaron las cabeceras
        if (!foundInRow && lastHeaderRow !== -1 && colMap.tipoOp !== -1) {
            dataStartRow = lastHeaderRow + 1;
            break;
        }

        // Si llegamos a la fila 19 y no hemos roto, fijar el inicio
        dataStartRow = lastHeaderRow + 1;
    }

    // FORZAR inicio desde fila 4 (índice 3) para capturar todas las operaciones
    dataStartRow = 3;

    // Determinar Fecha Base de la Pestaña
    let baseDate = null;
    try {
        const months = { 'ene': 0, 'feb': 1, 'mar': 2, 'abr': 3, 'may': 4, 'jun': 5, 'jul': 6, 'ago': 7, 'sep': 8, 'oct': 9, 'nov': 10, 'dic': 11 };
        let clean = sheetName.toLowerCase().replace(/de/g, '').trim();
        let dMatch = clean.match(/\d+/);
        let day = dMatch ? parseInt(dMatch[0]) : 1;
        let month = 0;
        for (let m in months) { if (clean.includes(m)) { month = months[m]; break; } }
        baseDate = new Date(2026, month, day);
    } catch (e) { baseDate = new Date(); }

    const sheetData = [];
    for (let i = dataStartRow; i < matrix.length; i++) {
        const row = matrix[i];
        if (!row || row.length === 0) continue;

        const get = (idx) => (idx !== -1 && row[idx] !== undefined && row[idx] !== null) ? row[idx] : "";
        const cleanTipo = String(get(colMap.tipoOp)).trim().toUpperCase();

        // ========== VALIDACIÓN MEJORADA DEL TIPO DE OPERACIÓN ==========
        // EXCLUIR: Encabezados, "NO CUMPLIÓ", filas vacías, etc.

        // 1. Excluir filas vacías o sin tipo
        if (cleanTipo === "" || cleanTipo === "0" || cleanTipo === "S/T") {
            console.log(`⏭️ Fila ${i}: Tipo vacío o S/T - IGNORADA`);
            continue;
        }

        // 2. EXCLUIR "NO CUMPLIÓ" y variaciones
        if (cleanTipo.includes("NO CUMPLIO") ||
            cleanTipo.includes("NO CUMPLIÓ") ||
            cleanTipo.includes("INCUMPLIDO") ||
            cleanTipo.includes("NO SE CUMPLIO")) {
            console.log(`⚠️ Fila ${i}: "NO CUMPLIÓ" detectado: ${cleanTipo} - IGNORADA`);
            continue;
        }

        // 3. EXCLUIR encabezados comunes (son títulos, no operaciones)
        const encabezados = [
            "TIPO DE OPERACION", "TIPO DE OP", "TIPO OP",
            "OPERACIONES", "ACTIVIDADES", "PLANIFICADAS",
            "TOTAL", "SUBTOTAL", "RESUMEN"
        ];
        if (encabezados.some(enc => cleanTipo === enc)) {
            console.log(`⚠️ Fila ${i}: Encabezado detectado: ${cleanTipo} - IGNORADA`);
            continue;
        }

        // DEBUG: Mostrar operaciones que SÍ pasan la validación
        if (cleanTipo.includes("SEG") || cleanTipo.includes("ARS")) {
            console.log(`✅ Fila ${i}: Operación SEG/ARS detectada: "${cleanTipo}" - ACEPTADA`);
        }

        // 4. MEJORAR validación de hora (más flexible para no perder filas válidas)
        const horaInicioRaw = String(get(colMap.horaInicio)).trim();

        // Solo excluir si la hora está completamente vacía o es un guión
        // NO excluir "0" porque puede ser 00:00 (medianoche)
        if (horaInicioRaw === "" || horaInicioRaw === "-") {
            // Verificar si la fila tiene al menos cantón o parroquia (puede ser operación válida sin hora aún)
            const canton = String(get(colMap.canton)).trim();
            const parroquia = String(get(colMap.parroquia)).trim();

            // Si tiene cantón o parroquia, es probablemente una operación válida
            if (canton === "" && parroquia === "") {
                console.log(`⚠️ Ignorando fila sin hora ni ubicación: ${cleanTipo}`);
                continue;
            }
            // Si tiene ubicación pero no hora, asignar hora por defecto (00:00)
            console.log(`ℹ️ Fila sin hora pero con ubicación válida: ${cleanTipo} en ${canton || parroquia} - asignando hora por defecto`);
        }

        const formatH = (v) => {
            const s = String(v).replace(/[^0-9]/g, '').padStart(4, '0');
            return s.slice(-4);
        };

        const hIni = formatH(get(colMap.horaInicio));
        const hFin = colMap.horaFin !== -1 ? formatH(get(colMap.horaFin)) : formatH(get(colMap.horaInicio + 1));

        // NORMALIZACIÓN DE TIEMPOS (Lógica de Medianoche)
        let startDate = new Date(baseDate);
        startDate.setHours(parseInt(hIni.slice(0, 2)), parseInt(hIni.slice(2, 4)), 0, 0);

        let endDate = new Date(baseDate);
        // Agregar 59 segundos para cubrir hasta el final del minuto (ej: 23:59:59)
        endDate.setHours(parseInt(hFin.slice(0, 2)), parseInt(hFin.slice(2, 4)), 59, 999);

        if (parseInt(hFin) < parseInt(hIni)) {
            endDate.setDate(endDate.getDate() + 1);
        }

        // ========== CÁLCULO DE PERSONAL (PMP) ==========
        // Intentar obtener valores individuales de OFI, AEROT, RESVITA
        const parseNum = (val) => {
            const num = parseInt(String(val).replace(/[^0-9]/g, ''));
            return isNaN(num) ? 0 : num;
        };

        const valOfi = parseNum(get(colMap.ofi));
        const valAerot = parseNum(get(colMap.aerot));
        const valRes = parseNum(get(colMap.res));

        // Suma las tres categorías
        const sumaDesglosada = valOfi + valAerot + valRes;

        // Decide qué guardar en la base de datos (Prioridad suma > 0)
        const pmpTotal = (sumaDesglosada > 0) ? sumaDesglosada : parseNum(get(colMap.pmp));

        // Obtener provincia (si existe la columna, sino usar MANABÍ por defecto)
        const provincia = colMap.provincia !== -1
            ? String(get(colMap.provincia)).trim().toUpperCase()
            : "MANABÍ";

        // Obtener sector (si existe)
        const sector = colMap.sector !== -1
            ? String(get(colMap.sector)).trim().toUpperCase()
            : "";

        sheetData.push({
            id: rawData.length + sheetData.length,
            nombreHoja: sheetName,
            fechaPlanificacion: baseDate,
            startDate, endDate,
            horaMilitar: `${hIni} - ${hFin}`,
            fuerza: "AÉREA",
            tipoOp: cleanTipo,
            provincia: provincia,
            canton: String(get(colMap.canton)).trim().toUpperCase(),
            parroquia: String(get(colMap.parroquia)).trim().toUpperCase(),
            sector: sector,
            resultados: String(get(colMap.resultados)) || "0",
            pmp: pmpTotal,
            // ← ACTUALIZADO: Valores individuales de personal
            personal: {
                oficiales: valOfi,
                aerotecnicos: valAerot,
                reservistas: valRes
            },
            medios: {
                camionetas: 0,
                buses: 0,
                camiones: 0
            }
        });
    }

    console.log(`📄 Pestaña "${sheetName}": ${sheetData.length} operaciones extraídas (de ${matrix.length - dataStartRow} filas procesadas)`);
    return sheetData;
}

// Función auxiliar para crear fechas en tiempo local de forma robusta
function parseLocalDate(dateStr, timeStr) {
    if (!dateStr) return null;
    try {
        const [y, m, d] = dateStr.split('-').map(Number);
        const [hh, mm] = (timeStr || "00:00").split(':').map(Number);
        return new Date(y, m - 1, d, hh, mm, 0, 0);
    } catch (e) {
        return null;
    }
}

// --- MOTOR DE FILTRADO SENIOR (Intersección de Intervalos) ---
function applyFilters() {
    const dStart = dom.filterStart.value;
    const tStart = dom.filterTimeStart.value || "00:00";
    const dEnd = dom.filterEnd.value;
    const tEnd = dom.filterTimeEnd.value || "23:59";
    const selProvincia = dom.filterProvincia.value;
    const selCanton = dom.filterCanton.value;
    const selTipo = dom.filterTipo.value;
    const searchVal = document.getElementById('searchInput')?.value.toUpperCase().trim() || "";

    const filterStart = parseLocalDate(dStart, tStart);
    let filterEnd = parseLocalDate(dEnd, tEnd);
    // Asegurar que filterEnd incluya hasta el final del minuto
    if (filterEnd) {
        filterEnd.setSeconds(59, 999);
    }

    const groups = {};

    rawData.forEach(item => {
        // 1. Filtro de Provincia
        if (selProvincia !== 'TODOS' && item.provincia !== selProvincia) return;

        // 2. Filtro de Jurisdicción (Cantón)
        if (selCanton !== 'TODOS' && item.canton !== selCanton) return;

        // 3. Filtro de Tipo de Operación
        if (selTipo !== 'TODOS' && item.tipoOp !== selTipo) return;

        // 4. Búsqueda Rápida (Soporta múltiples campos)
        if (searchVal) {
            const rowContent = `${item.tipoOp} ${item.provincia} ${item.canton} ${item.parroquia} ${item.resultados}`.toUpperCase();
            if (!rowContent.includes(searchVal)) return;
        }

        // 3. FILTRO DE TIEMPO (Planificadas vs Ejecutadas)
        // Planificada: La fecha de la operación está dentro del rango de FECHAS del filtro
        let isPlanned = true;
        if (filterStart && filterEnd) {
            const itemDate = new Date(item.fechaPlanificacion);
            itemDate.setHours(0, 0, 0, 0);
            const startD = new Date(filterStart);
            startD.setHours(0, 0, 0, 0);
            const endD = new Date(filterEnd);
            endD.setHours(23, 59, 59, 999);

            isPlanned = (itemDate >= startD && itemDate <= endD);
        }

        if (!isPlanned) return;

        // Ejecutada: Intersecta con el rango de FECHA+HORA exacto del filtro
        let isExecuted = true;
        if (filterStart && filterEnd) {
            isExecuted = (item.startDate <= filterEnd && item.endDate >= filterStart);
        }

        // 4. AGRUPACIÓN (Por Tipo, Provincia, Cantón y Parroquia)
        const dateKey = item.fechaPlanificacion.toISOString().split('T')[0];
        const key = `${item.tipoOp}|${item.provincia}|${item.canton}|${item.parroquia}|${dateKey}`.toUpperCase();

        if (!groups[key]) {
            groups[key] = {
                ...item,
                sumPlanif: 0,
                sumEjecut: 0,
                sumPmp: 0,
                items: []
            };
        }
        groups[key].sumPlanif += 1;
        if (isExecuted) {
            groups[key].sumEjecut += 1;
        }
        groups[key].sumPmp += item.pmp;
        groups[key].items.push(item);
    });

    filteredData = Object.values(groups).map(g => {
        // Consolidar resultados de todos los items del grupo
        const results = g.items
            .map(it => it.resultados)
            .filter(r => r && r !== "0");
        g.resultados = results.length > 0 ? [...new Set(results)].join(" / ") : "0";
        return g;
    });

    // Ordenar por Fecha y Prioridad
    const priority = { "MANTA": 1, "MONTECRISTI": 2, "JIPIJAPA": 3 };
    filteredData.sort((a, b) => {
        if (a.fechaPlanificacion.getTime() !== b.fechaPlanificacion.getTime()) return a.fechaPlanificacion - b.fechaPlanificacion;
        return (priority[a.canton] || 99) - (priority[b.canton] || 99);
    });

    updateDashboard();
}

// --- NUEVA FUNCIÓN: Filtrar por Pestaña Específica (Día) ---
function applyFiltersForSheet(sheetName) {
    console.log(`Filtrando operaciones de la pestaña: ${sheetName}`);

    const dStart = dom.filterStart.value;
    const tStart = dom.filterTimeStart.value || "00:00";
    const dEnd = dom.filterEnd.value;
    const tEnd = dom.filterTimeEnd.value || "23:59";
    const filterStart = parseLocalDate(dStart, tStart);
    let filterEnd = parseLocalDate(dEnd, tEnd);
    // Asegurar que filterEnd incluya hasta el final del minuto
    if (filterEnd) {
        filterEnd.setSeconds(59, 999);
    }

    // Log del rango de filtrado
    if (filterStart && filterEnd) {
        console.log(`🔍 FILTRO CONFIGURADO:`);
        console.log(`   Desde: ${filterStart.toLocaleString()}`);
        console.log(`   Hasta: ${filterEnd.toLocaleString()}`);
    }

    const selProvincia = dom.filterProvincia.value;
    const selCanton = dom.filterCanton.value;
    const selTipo = dom.filterTipo.value;
    const groups = {};

    // Filtrar solo los datos de la pestaña seleccionada
    rawData.forEach(item => {
        if (item.nombreHoja !== sheetName) return;

        // Filtro opcional de provincia
        if (selProvincia !== 'TODOS' && item.provincia !== selProvincia) return;

        // Filtro opcional de jurisdicción
        if (selCanton !== 'TODOS' && item.canton !== selCanton) return;

        // Filtro opcional de tipo
        if (selTipo !== 'TODOS' && item.tipoOp !== selTipo) return;

        // ========== NUEVA LÓGICA DE FILTRADO OPTIMIZADA ==========
        // Criterio 1: VISIBILIDAD - ¿Se debe mostrar la operación?
        // Ocultar operaciones que terminan ANTES del punto "Desde"
        let shouldShow = true;
        if (filterStart && filterEnd) {
            // Si la operación termina antes del inicio del filtro, ocultarla
            shouldShow = (item.endDate.getTime() >= filterStart.getTime());
        }

        // Si no debe mostrarse, saltar esta operación
        if (!shouldShow) {
            console.log(`⏭️ Operación oculta (termina antes del filtro): ${item.tipoOp} | Fin: ${item.endDate.toLocaleString()}`);
            return;
        }

        // Criterio 2: CUMPLIMIENTO - ¿Se ejecutó en el rango?
        // Una operación se considera ejecutada si:
        // - Su hora de inicio es antes o igual al fin del filtro Y
        // - Su hora de fin es después o igual al inicio del filtro
        // Esto incluye operaciones que "cierran" exactamente en el punto "Desde"
        let isExecuted = true;
        if (filterStart && filterEnd) {
            const iStart = item.startDate.getTime();
            const iEnd = item.endDate.getTime();
            const fStart = filterStart.getTime();
            const fEnd = filterEnd.getTime();

            isExecuted = (iStart <= fEnd && iEnd >= fStart);

            console.log(`📊 ${item.tipoOp} | Inicio: ${item.startDate.toLocaleTimeString()} | Fin: ${item.endDate.toLocaleTimeString()} | Ejecutada: ${isExecuted ? '✅' : '❌'}`);
        }
        // Si NO hay filtro de fecha/hora, todas se consideran ejecutadas por defecto

        // Criterio 3: PROYECCIÓN - Todas las operaciones del día son planificadas
        // (se cuenta en sumPlanif más abajo)

        // AGRUPACIÓN (Por Tipo, Provincia, Cantón y Parroquia)
        const dateKey = item.fechaPlanificacion.toISOString().split('T')[0];
        const key = `${item.tipoOp}|${item.provincia}|${item.canton}|${item.parroquia}|${dateKey}`.toUpperCase();

        if (!groups[key]) {
            groups[key] = {
                ...item,
                sumPlanif: 0,
                sumEjecut: 0,
                sumPmp: 0,
                items: []
            };
        }
        groups[key].sumPlanif += 1;
        if (isExecuted) {
            groups[key].sumEjecut += 1;
            groups[key].sumPmp += item.pmp; // PMP solo se suma si se ejecutó
        }
        groups[key].items.push(item);
    });

    filteredData = Object.values(groups).map(g => {
        const results = g.items
            .map(it => it.resultados)
            .filter(r => r && r !== "0");
        g.resultados = results.length > 0 ? [...new Set(results)].join(" / ") : "0";
        return g;
    });

    // Ordenar por Fecha y Prioridad
    const priority = { "MANTA": 1, "MONTECRISTI": 2, "JIPIJAPA": 3 };
    filteredData.sort((a, b) => {
        if (a.fechaPlanificacion.getTime() !== b.fechaPlanificacion.getTime()) return a.fechaPlanificacion - b.fechaPlanificacion;
        return (priority[a.canton] || 99) - (priority[b.canton] || 99);
    });

    // DEBUG: Mostrar conteo total de operaciones
    const totalPlanif = filteredData.reduce((sum, f) => sum + f.sumPlanif, 0);
    const totalEjecut = filteredData.reduce((sum, f) => sum + f.sumEjecut, 0);
    const noEjecutadas = totalPlanif - totalEjecut;

    console.log(`✅ PESTAÑA ${sheetName}:`);
    console.log(`   📋 Total Planificadas: ${totalPlanif}`);
    console.log(`   ✔️  Total Ejecutadas: ${totalEjecut}`);
    if (noEjecutadas > 0) {
        console.log(`   ⚠️  No contadas como ejecutadas: ${noEjecutadas}`);
        // Mostrar cuáles operaciones no fueron contadas
        filteredData.filter(g => g.sumEjecut < g.sumPlanif).forEach(g => {
            console.log(`      - ${g.tipoOp} en ${g.canton}: ${g.sumEjecut}/${g.sumPlanif} (Horario: ${g.horaMilitar})`);
        });
    }

    updateDashboard();
}

function updateDashboard() {
    dom.tableBody.innerHTML = '';
    let tPlan = 0, tEjec = 0, tPmp = 0;

    filteredData.forEach(item => {
        tPlan += item.sumPlanif;
        tEjec += item.sumEjecut;
        // Sumar PMP solo de operaciones ejecutadas (proporcional)
        const pmpEjecutado = item.sumEjecut > 0 ? Math.round((item.sumPmp / item.sumPlanif) * item.sumEjecut) : 0;
        tPmp += pmpEjecutado;
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td>${item.fuerza}</td>
            <td>${item.tipoOp}</td>
            <td class="text-center">${item.sumPlanif}</td>
            <td class="text-center"><strong>${item.sumEjecut}/${item.sumPlanif}</strong></td>
            <td>${item.provincia}</td>
            <td>${item.canton}</td>
            <td>${item.parroquia}</td>
            <td class="text-center">${item.fechaPlanificacion.toLocaleDateString()}</td>
            <td>${item.resultados}</td>
            <td class="text-center"><strong>${item.sumPmp}</strong></td>
            <td class="text-center hide-on-print">
                <button class="btn btn-sm btn-outline-warning p-1" onclick="openEditModal(${item.id})">
                    <span class="material-icons" style="font-size:16px;">edit</span>
                </button>
            </td>
        `;
        dom.tableBody.appendChild(tr);
    });

    dom.stats.total.textContent = tPlan;
    dom.stats.pmp.textContent = tPmp;
    const perc = tPlan > 0 ? ((tEjec / tPlan) * 100).toFixed(1) : 0;
    dom.stats.effectiveness.textContent = `${tEjec} (${perc}%)`;

    // Actualizar el total en el footer de la tabla
    const tableTotalPlanif = document.getElementById('tableTotalPlanif');
    if (tableTotalPlanif) {
        tableTotalPlanif.textContent = tPlan;
    }

    console.log(`📈 KPIs actualizados: Planificadas=${tPlan}, Ejecutadas=${tEjec}, Eficacia=${perc}%, PMP=${tPmp}`);
    renderSummaryTable();
    renderCharts(); // ← AGREGADO: Renderizar gráficos
}

function renderSummaryTable() {
    const body = document.getElementById('summaryBody');
    const totalElement = document.getElementById('summaryTotal');
    if (!body) return;

    const counts = {};
    filteredData.forEach(i => counts[i.tipoOp] = (counts[i.tipoOp] || 0) + i.sumEjecut);

    body.innerHTML = Object.keys(counts).sort().map(t => `<tr><td>${t}</td><td class="text-end"><strong>${counts[t]}</strong></td></tr>`).join('');

    // Calcular y mostrar el TOTAL de operaciones
    const total = Object.values(counts).reduce((sum, val) => sum + val, 0);
    if (totalElement) {
        totalElement.textContent = total;
    }

    console.log('📊 Resumen por Tipo:', counts, 'Total:', total);
}

function populateTipoFilter() {
    const tipos = [...new Set(rawData.map(i => i.tipoOp).filter(x => x))].sort();
    dom.filterTipo.innerHTML = '<option value="TODOS">-- Todos --</option>';
    tipos.forEach(x => {
        if (x !== "TIPO DE OPERACIÓN" && x !== "TIPO OP") {
            dom.filterTipo.innerHTML += `<option value="${x}">${x}</option>`;
        }
    });
}

function populateProvinciaFilter() {
    const p = [...new Set(rawData.map(i => i.provincia).filter(x => x))].sort();
    dom.filterProvincia.innerHTML = '<option value="TODOS">-- Todas --</option>';
    p.forEach(x => dom.filterProvincia.innerHTML += `<option value="${x}">${x}</option>`);
}

function populateCantonFilter() {
    const c = [...new Set(rawData.map(i => i.canton).filter(x => x))].sort();
    dom.filterCanton.innerHTML = '<option value="TODOS">-- Todos --</option>';
    c.forEach(x => dom.filterCanton.innerHTML += `<option value="${x}">${x}</option>`);
}

function populateFilters() {
    populateProvinciaFilter();
    populateCantonFilter();
    populateTipoFilter();
}

function loadAllSheets() {
    rawData = []; // Limpiar antes de cargar nuevo archivo
    workbook.SheetNames.forEach(name => {
        rawData = rawData.concat(extractDataFromSheet(workbook.Sheets[name], name));
    });

    if (rawData.length > 0) {
        // Ajustar fechas del filtro al rango de datos cargados para que se vea todo al inicio
        const dates = rawData.map(d => d.fechaPlanificacion.getTime());
        const minDate = new Date(Math.min(...dates)).toISOString().split('T')[0];
        const maxDate = new Date(Math.max(...dates)).toISOString().split('T')[0];

        dom.filterStart.value = minDate;
        dom.filterEnd.value = maxDate;
    }

    populateFilters();
    applyFilters();
}

// ========== RENDERIZADO DE GRÁFICOS ==========
function renderCharts() {
    renderHourlyChart();
    renderGeoChart();
}

// Colores para tipos de operación (consistentes en todos los gráficos)
const tipoColors = {
    'RASTRILLAJE': 'rgba(0, 120, 212, 0.8)',
    'COMBATE URBANO': 'rgba(220, 53, 69, 0.8)',
    'CONTROL DE ARMAS (CAMEX)': 'rgba(40, 167, 69, 0.8)',
    'PROTECCIÓN DE ÁREAS RESERVADAS (ARS)': 'rgba(255, 193, 7, 0.8)',
    'PATRULLAJE': 'rgba(23, 162, 184, 0.8)',
    'CONTROL VEHICULAR': 'rgba(111, 66, 193, 0.8)',
    'PUNTO DE CONTROL': 'rgba(253, 126, 20, 0.8)',
    'REGISTRO': 'rgba(232, 62, 140, 0.8)',
    'DEFAULT': 'rgba(108, 117, 125, 0.8)'
};

// Variable global para filtro activo por tipo
let activeTypeFilter = null;

function getColorForTipo(tipo) {
    return tipoColors[tipo] || tipoColors['DEFAULT'];
}

// Gráfico 1: Distribución Horaria (Barras Apiladas por Tipo de Operación)
function renderHourlyChart() {
    const canvas = document.getElementById('opsChart');
    if (!canvas) return;

    // Destruir gráfico anterior
    if (chartInstance) {
        chartInstance.destroy();
    }

    // Obtener todos los tipos únicos y horas
    const tiposUnicos = [...new Set(filteredData.map(item => item.tipoOp))].sort();
    const horasSet = new Set();

    // Crear estructura: { hora: { tipo: cantidad } }
    const hourTypeCounts = {};

    filteredData.forEach(item => {
        const hour = parseInt(item.horaMilitar.split('-')[0].trim().substring(0, 2));
        horasSet.add(hour);

        if (!hourTypeCounts[hour]) {
            hourTypeCounts[hour] = {};
        }
        hourTypeCounts[hour][item.tipoOp] = (hourTypeCounts[hour][item.tipoOp] || 0) + item.sumEjecut;
    });

    // Ordenar horas
    const horas = [...horasSet].sort((a, b) => a - b);
    const labels = horas.map(h => `${String(h).padStart(2, '0')}:00`);

    // Crear datasets para cada tipo de operación
    const datasets = tiposUnicos.map((tipo, index) => {
        const data = horas.map(hora => hourTypeCounts[hora]?.[tipo] || 0);
        const color = getColorForTipo(tipo);

        return {
            label: tipo,
            data: data,
            backgroundColor: color,
            borderColor: color.replace('0.8', '1'),
            borderWidth: 1,
            borderRadius: 4,
            // Si hay filtro activo, atenuar los demás
            hidden: activeTypeFilter && activeTypeFilter !== tipo
        };
    });

    // Crear gráfico
    const ctx = canvas.getContext('2d');
    chartInstance = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: labels,
            datasets: datasets
        },
        options: {
            responsive: true,
            maintainAspectRatio: true,
            interaction: {
                mode: 'index',
                intersect: false
            },
            plugins: {
                legend: {
                    display: true,
                    position: 'top',
                    labels: {
                        boxWidth: 12,
                        padding: 8,
                        font: { size: 10 },
                        usePointStyle: true
                    },
                    onClick: (e, legendItem, legend) => {
                        // Filtrar por tipo al hacer clic en la leyenda
                        const tipo = legendItem.text;
                        if (activeTypeFilter === tipo) {
                            activeTypeFilter = null; // Quitar filtro
                        } else {
                            activeTypeFilter = tipo; // Activar filtro
                        }
                        // Re-renderizar ambos gráficos con el filtro
                        renderHourlyChart();
                        renderGeoChart();
                        console.log(`🔍 Filtro de tipo: ${activeTypeFilter || 'TODOS'}`);
                    }
                },
                title: {
                    display: activeTypeFilter !== null,
                    text: `Filtrado: ${activeTypeFilter || ''}`,
                    font: { size: 11 },
                    color: '#dc3545'
                },
                tooltip: {
                    callbacks: {
                        title: (context) => `Hora: ${context[0].label}`,
                        label: (context) => {
                            return `${context.dataset.label}: ${context.raw} operaciones`;
                        },
                        footer: (context) => {
                            const total = context.reduce((sum, c) => sum + c.raw, 0);
                            return `Total: ${total} operaciones`;
                        }
                    }
                }
            },
            scales: {
                x: {
                    stacked: true,
                    grid: { display: false },
                    ticks: {
                        font: { size: 11 },
                        color: '#605e5c'
                    }
                },
                y: {
                    stacked: true,
                    beginAtZero: true,
                    ticks: {
                        stepSize: 1,
                        font: { size: 11 },
                        color: '#605e5c'
                    },
                    grid: {
                        color: 'rgba(0, 0, 0, 0.05)'
                    }
                }
            },
            onClick: (e, elements) => {
                if (elements.length > 0) {
                    const datasetIndex = elements[0].datasetIndex;
                    const tipo = chartInstance.data.datasets[datasetIndex].label;
                    console.log(`📌 Clic en tipo: ${tipo}`);
                    // Opcional: filtrar la tabla al hacer clic
                }
            }
        },
        plugins: [{
            // Plugin personalizado para mostrar totales en las barras
            id: 'barTotals',
            afterDatasetsDraw: (chart) => {
                const ctx = chart.ctx;
                chart.data.labels.forEach((label, index) => {
                    let total = 0;
                    chart.data.datasets.forEach(dataset => {
                        if (!dataset.hidden) {
                            total += dataset.data[index] || 0;
                        }
                    });

                    if (total > 0) {
                        const meta = chart.getDatasetMeta(chart.data.datasets.length - 1);
                        const bar = meta.data[index];

                        if (bar) {
                            ctx.save();
                            ctx.font = 'bold 11px Segoe UI';
                            ctx.fillStyle = '#201f1e';
                            ctx.textAlign = 'center';
                            ctx.textBaseline = 'bottom';
                            ctx.fillText(total, bar.x, bar.y - 5);
                            ctx.restore();
                        }
                    }
                });
            }
        }]
    });

    console.log('📊 Gráfico horario apilado renderizado con', tiposUnicos.length, 'tipos');
}

// Gráfico 2: Por Jurisdicción (Doughnut) - VINCULADO al filtro de tipo
function renderGeoChart() {
    const canvas = document.getElementById('geoChart');
    if (!canvas) return;

    // Destruir gráfico anterior
    if (geoChartInstance) {
        geoChartInstance.destroy();
    }

    // Filtrar datos según el tipo activo (vinculación con gráfico horario)
    const dataToUse = activeTypeFilter
        ? filteredData.filter(item => item.tipoOp === activeTypeFilter)
        : filteredData;

    // Agrupar operaciones por cantón
    const cantonCounts = {};
    dataToUse.forEach(item => {
        cantonCounts[item.canton] = (cantonCounts[item.canton] || 0) + item.sumEjecut;
    });

    // Ordenar por cantidad (descendente)
    const sortedCantones = Object.entries(cantonCounts)
        .sort((a, b) => b[1] - a[1])
        .slice(0, 10); // Top 10 cantones

    const labels = sortedCantones.map(c => c[0]);
    const data = sortedCantones.map(c => c[1]);

    // Colores dinámicos
    const colors = [
        'rgba(0, 120, 212, 0.8)',
        'rgba(40, 167, 69, 0.8)',
        'rgba(255, 193, 7, 0.8)',
        'rgba(220, 53, 69, 0.8)',
        'rgba(23, 162, 184, 0.8)',
        'rgba(108, 117, 125, 0.8)',
        'rgba(111, 66, 193, 0.8)',
        'rgba(253, 126, 20, 0.8)',
        'rgba(232, 62, 140, 0.8)',
        'rgba(13, 110, 253, 0.8)'
    ];

    // Crear gráfico
    const ctx = canvas.getContext('2d');
    geoChartInstance = new Chart(ctx, {
        type: 'doughnut',
        data: {
            labels: labels,
            datasets: [{
                label: 'Operaciones',
                data: data,
                backgroundColor: colors,
                borderColor: '#fff',
                borderWidth: 2
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: true,
            plugins: {
                title: {
                    display: activeTypeFilter !== null,
                    text: `Filtrado: ${activeTypeFilter || ''}`,
                    font: { size: 11 },
                    color: '#dc3545'
                },
                legend: {
                    position: 'right',
                    labels: {
                        boxWidth: 12,
                        font: { size: 11 },
                        generateLabels: (chart) => {
                            const dataset = chart.data.datasets[0];
                            return chart.data.labels.map((label, i) => ({
                                text: `${label} (${dataset.data[i]})`,
                                fillStyle: dataset.backgroundColor[i],
                                strokeStyle: dataset.borderColor,
                                lineWidth: dataset.borderWidth,
                                hidden: false,
                                index: i
                            }));
                        }
                    }
                },
                tooltip: {
                    callbacks: {
                        label: (context) => {
                            const total = context.dataset.data.reduce((a, b) => a + b, 0);
                            const percentage = ((context.raw / total) * 100).toFixed(1);
                            return `${context.label}: ${context.raw} (${percentage}%)`;
                        }
                    }
                }
            }
        }
    });

    const tipoMsg = activeTypeFilter ? ` [Filtro: ${activeTypeFilter}]` : '';
    console.log(`🗺️ Gráfico geográfico renderizado${tipoMsg}`);
}

// ========== GESTIÓN DE REPORTES Y MENSAJES ==========
function generateCustomReport() {
    const dStart = document.getElementById('reportStartDate').value;
    const tStart = document.getElementById('reportStartTime').value || "00:00";
    const dEnd = document.getElementById('reportEndDate').value;
    const tEnd = document.getElementById('reportEndTime').value || "23:59";

    const filterStart = parseLocalDate(dStart, tStart);
    const filterEnd = parseLocalDate(dEnd, tEnd);

    // Función auxiliar para singular/plural
    const pluralize = (count, singular, plural) => {
        return count === 1 ? singular : (plural || singular + 's');
    };

    // Determinar encabezado según hora
    let encabezado = "";
    if (filterEnd) {
        const hora = filterEnd.getHours();
        if ((hora >= 5) && (hora < 17)) {
            encabezado = "*BUENOS DÍAS MI GENERAL, ME PERMITO DAR PARTE DE LAS OPERACIONES EJECUTADAS POR EL GOMAI \"MANABÍ\" EN LOS CANTONES:*";
        } else {
            encabezado = "*BUENAS TARDES MI GENERAL, ME PERMITO DAR PARTE DE LAS OPERACIONES EJECUTADAS POR EL GOMAI \"MANABÍ\" EN LOS CANTONES:*";
        }
    }

    // Filtrar datos ejecutados en el rango de fecha/hora especificado
    const reportData = rawData.filter(item => {
        if (!filterStart || !filterEnd) return false;

        // Verificar solapamiento de horarios:
        // Una operación se ejecuta en el rango si:
        // - Su hora de inicio es antes o igual al fin del filtro Y
        // - Su hora de fin es después o igual al inicio del filtro
        const isExecuted = (item.startDate.getTime() <= filterEnd.getTime() && item.endDate.getTime() >= filterStart.getTime());

        return isExecuted;
    });

    console.log(`📊 Reporte generado:`);
    console.log(`   Rango: ${filterStart.toLocaleString()} - ${filterEnd.toLocaleString()}`);
    console.log(`   Operaciones encontradas: ${reportData.length}`);

    // Obtener cantones únicos
    const cantones = [...new Set(reportData.map(item => item.canton))].filter(c => c).sort();

    // Agrupar por tipo de operación
    const operacionesPorTipo = {};
    reportData.forEach(item => {
        const tipo = item.tipoOp.toUpperCase();
        if (!operacionesPorTipo[tipo]) {
            operacionesPorTipo[tipo] = [];
        }
        operacionesPorTipo[tipo].push(item);
    });

    // Mapeo de tipos a categorías
    const tipoARS = Object.keys(operacionesPorTipo).filter(t =>
        t.includes('SEG') || t.includes('ARS') || t.includes('PROTECCIÓN') || t.includes('PROTECCION')
    );

    const tipoRastrillaje = Object.keys(operacionesPorTipo).filter(t => t.includes('RASTRILLAJE'));
    const tipoCombate = Object.keys(operacionesPorTipo).filter(t => t.includes('COMBATE'));
    const tipoEjes = Object.keys(operacionesPorTipo).filter(t => t.includes('EJE') || t.includes('VIAL'));
    const tipoCAMEX = Object.keys(operacionesPorTipo).filter(t => t.includes('CAMEX') || t.includes('ARMAS'));

    // Función para generar sección de operación
    const generarSeccionOperacion = (tipos, nombreOperacion) => {
        if (tipos.length === 0) return "";

        let ops = [];
        tipos.forEach(tipo => {
            ops = ops.concat(operacionesPorTipo[tipo]);
        });

        if (ops.length === 0) return "";

        let texto = `\n${nombreOperacion} (${String(ops.length).padStart(2, '0')})\n`;

        // Recopilar todos los cantones y sectores
        const cantones = new Set();
        const sectores = new Set();
        let totalOficiales = 0;
        let totalAerotecnicos = 0;
        let totalReservistas = 0;
        let totalCamionetas = 0;
        let totalCamiones = 0;
        let totalBuses = 0;

        ops.forEach(op => {
            if (op.canton) cantones.add(op.canton);
            if (op.sector) sectores.add(op.sector);
            totalOficiales += (op.personal?.oficiales || 0);
            totalAerotecnicos += (op.personal?.aerotecnicos || 0);
            totalReservistas += (op.personal?.reservistas || 0);
            totalCamionetas += (op.medios?.camionetas || 0);
            totalCamiones += (op.medios?.camiones || 0);
            totalBuses += (op.medios?.buses || 0);
        });


        // CANTÓN:
        if (cantones.size > 0) {
            texto += `*CANTÓN:*\n`;
            [...cantones].sort().forEach(canton => {
                texto += `${canton}\n`;
            });
        }

        // SECTOR:
        if (sectores.size > 0) {
            texto += `*SECTOR:*\n`;
            texto += `${[...sectores].sort().join(' / ')}\n`;
        }

        // Personal empleado
        texto += `*PERSONAL EMPLEADO:*\n`;
        if (totalOficiales > 0) {
            texto += `${pluralize(totalOficiales, 'Oficial', 'Oficiales')}: ${String(totalOficiales).padStart(2, '0')}\n`;
        }
        if (totalAerotecnicos > 0) {
            texto += `${pluralize(totalAerotecnicos, 'Aerotécnico', 'Aerotécnicos')}: ${String(totalAerotecnicos).padStart(2, '0')}\n`;
        }
        if (totalReservistas > 0) {
            texto += `${pluralize(totalReservistas, 'Reservista', 'Reservistas')}: ${String(totalReservistas).padStart(2, '0')}\n`;
        }

        // Medios (si existen)
        const hayMedios = totalCamionetas > 0 || totalCamiones > 0 || totalBuses > 0;
        if (hayMedios) {
            texto += `*MEDIOS:*\n`;
            if (totalCamionetas > 0) {
                texto += `${pluralize(totalCamionetas, 'Camioneta', 'Camionetas')}: ${String(totalCamionetas).padStart(2, '0')}\n`;
            }
            if (totalCamiones > 0) {
                texto += `${pluralize(totalCamiones, 'Camión', 'Camiones')}: ${String(totalCamiones).padStart(2, '0')}\n`;
            }
            if (totalBuses > 0) {
                texto += `${pluralize(totalBuses, 'Bus', 'Buses')}: ${String(totalBuses).padStart(2, '0')}\n`;
            }
        }


        return texto;
    };

    // Construir reporte
    let reportText = encabezado + "\n";
    cantones.forEach(canton => {
        reportText += `-${canton}\n`;
    });

    reportText += "\n*PROTECCIÓN DE LAS ZONAS DE SEGURIDAD DEL ESTADO*\n";

    // A) ARS
    const totalARS = tipoARS.reduce((sum, tipo) => sum + (operacionesPorTipo[tipo]?.length || 0), 0);
    if (totalARS > 0) {
        reportText += `\n*A) PROTECCIÓN DE LAS ÁREAS RESERVADAS DE SEGURIDAD TERRESTRE ARS (${String(totalARS).padStart(2, '0')})*\n`;
        reportText += generarSeccionOperacion(tipoARS, "ARS");
    }

    // B) COMPETENCIA LEGAL
    const totalCompetencia =
        tipoRastrillaje.reduce((sum, tipo) => sum + (operacionesPorTipo[tipo]?.length || 0), 0) +
        tipoCombate.reduce((sum, tipo) => sum + (operacionesPorTipo[tipo]?.length || 0), 0) +
        tipoEjes.reduce((sum, tipo) => sum + (operacionesPorTipo[tipo]?.length || 0), 0) +
        tipoCAMEX.reduce((sum, tipo) => sum + (operacionesPorTipo[tipo]?.length || 0), 0);

    if (totalCompetencia > 0) {
        reportText += `\n*B) COMPETENCIA LEGAL DE FUERZAS ARMADAS (${String(totalCompetencia).padStart(2, '0')})*\n`;

        if (tipoRastrillaje.length > 0) {
            reportText += generarSeccionOperacion(tipoRastrillaje, "RASTRILLAJE");
        }
        if (tipoCombate.length > 0) {
            reportText += generarSeccionOperacion(tipoCombate, "COMBATE URBANO");
        }
        if (tipoEjes.length > 0) {
            reportText += generarSeccionOperacion(tipoEjes, "EJES VIALES");
        }
        if (tipoCAMEX.length > 0) {
            reportText += generarSeccionOperacion(tipoCAMEX, "CAMEX");
        }
    }

    document.getElementById('generatedMessage').value = reportText.toUpperCase();
}

function copyMessage() {
    const textarea = document.getElementById('generatedMessage');
    textarea.select();
    document.execCommand('copy');
    alert("Copiado al portapapeles");
}

// ========== CRUD Y EDICIÓN ==========
function openEditModal(id) {
    const baseItem = rawData.find(x => x.id === id);
    if (!baseItem) return;

    // Obtener todos los items que pertenecen al mismo grupo
    const dateKey = baseItem.fechaPlanificacion.toISOString().split('T')[0];
    const groupKey = `${baseItem.tipoOp}|${baseItem.provincia}|${baseItem.canton}|${baseItem.parroquia}|${dateKey}`.toUpperCase();

    const itemsInGroup = rawData.filter(it => {
        const itDateKey = it.fechaPlanificacion.toISOString().split('T')[0];
        const itKey = `${it.tipoOp}|${it.provincia}|${it.canton}|${it.parroquia}|${itDateKey}`.toUpperCase();
        return itKey === groupKey;
    });

    document.getElementById('editId').value = id;
    document.getElementById('editTipo').value = baseItem.tipoOp;

    const container = document.getElementById('opsEditContainer');
    container.innerHTML = `<h6 class="mb-3 text-primary">Grupo: ${baseItem.tipoOp} (${itemsInGroup.length} registros)</h6>`;

    itemsInGroup.forEach((item, idx) => {
        container.innerHTML += `
            <div class="card mb-3 shadow-sm border-start border-primary border-4 multi-edit-row" data-id="${item.id}">
                <div class="card-body p-3">
                    <div class="d-flex justify-content-between mb-2">
                        <span class="badge bg-secondary">Registro #${idx + 1}</span>
                        <span class="fw-bold">${item.horaMilitar}</span>
                    </div>
                    <div class="row g-2">
                        <div class="col-md-6">
                            <label class="form-label small mb-1">Cantón</label>
                            <input type="text" class="form-control form-control-sm" name="canton" value="${item.canton}">
                        </div>
                        <div class="col-md-6">
                            <label class="form-label small mb-1">Parroquia</label>
                            <input type="text" class="form-control form-control-sm" name="parroquia" value="${item.parroquia}">
                        </div>
                        <div class="col-12">
                            <label class="form-label small mb-1">Resultados</label>
                            <textarea class="form-control form-control-sm" name="resultados" rows="2">${item.resultados}</textarea>
                        </div>
                        <div class="col-md-4">
                            <label class="form-label small mb-1">OF</label>
                            <input type="number" class="form-control form-control-sm" name="ofi" value="${item.personal?.oficiales || 0}">
                        </div>
                        <div class="col-md-4">
                            <label class="form-label small mb-1">AE</label>
                            <input type="number" class="form-control form-control-sm" name="aerot" value="${item.personal?.aerotecnicos || 0}">
                        </div>
                        <div class="col-md-4">
                            <label class="form-label small mb-1">RE</label>
                            <input type="number" class="form-control form-control-sm" name="res" value="${item.personal?.reservistas || 0}">
                        </div>
                        <div class="col-12 mt-2">
                            <label class="form-label small mb-1 fw-bold">Medios Empleados</label>
                        </div>
                        <div class="col-md-4">
                            <label class="form-label small mb-1">Camionetas</label>
                            <input type="number" class="form-control form-control-sm" name="camionetas" value="${item.medios?.camionetas || 0}" min="0">
                        </div>
                        <div class="col-md-4">
                            <label class="form-label small mb-1">Buses</label>
                            <input type="number" class="form-control form-control-sm" name="buses" value="${item.medios?.buses || 0}" min="0">
                        </div>
                        <div class="col-md-4">
                            <label class="form-label small mb-1">Camiones</label>
                            <input type="number" class="form-control form-control-sm" name="camiones" value="${item.medios?.camiones || 0}" min="0">
                        </div>
                    </div>
                </div>
            </div>
        `;
    });

    const modalElement = document.getElementById('editModal');
    const modal = new bootstrap.Modal(modalElement);

    // Asegurar que el modal tenga z-index alto
    modalElement.style.zIndex = '1060';

    // Hacer el modal draggable (arrastrable)
    makeDraggable(modalElement);

    modal.show();
}

// Función para hacer un modal draggable
function makeDraggable(modalElement) {
    const dialog = modalElement.querySelector('.modal-dialog');
    const header = modalElement.querySelector('.modal-header');

    if (!header || !dialog) return;

    let isDragging = false;
    let currentX;
    let currentY;
    let initialX;
    let initialY;

    header.addEventListener('mousedown', (e) => {
        isDragging = true;
        initialX = e.clientX - (dialog.offsetLeft || 0);
        initialY = e.clientY - (dialog.offsetTop || 0);

        // Cambiar posición a absolute para poder moverlo
        dialog.style.position = 'absolute';
        dialog.style.margin = '0';
    });

    document.addEventListener('mousemove', (e) => {
        if (isDragging) {
            e.preventDefault();
            currentX = e.clientX - initialX;
            currentY = e.clientY - initialY;

            dialog.style.left = currentX + 'px';
            dialog.style.top = currentY + 'px';
        }
    });

    document.addEventListener('mouseup', () => {
        isDragging = false;
    });
}

function saveChanges() {
    const rows = document.querySelectorAll('.multi-edit-row');
    const newTipo = document.getElementById('editTipo').value.toUpperCase();

    rows.forEach(row => {
        const id = parseInt(row.dataset.id);
        const item = rawData.find(x => x.id === id);
        if (!item) return;

        item.tipoOp = newTipo;
        item.canton = row.querySelector('[name="canton"]').value.toUpperCase();
        item.parroquia = row.querySelector('[name="parroquia"]').value.toUpperCase();
        item.resultados = row.querySelector('[name="resultados"]').value;

        const ofi = parseInt(row.querySelector('[name="ofi"]').value) || 0;
        const aerot = parseInt(row.querySelector('[name="aerot"]').value) || 0;
        const res = parseInt(row.querySelector('[name="res"]').value) || 0;

        const camionetas = parseInt(row.querySelector('[name="camionetas"]').value) || 0;
        const buses = parseInt(row.querySelector('[name="buses"]').value) || 0;
        const camiones = parseInt(row.querySelector('[name="camiones"]').value) || 0;

        item.personal = { oficiales: ofi, aerotecnicos: aerot, reservistas: res };
        item.medios = { camionetas: camionetas, buses: buses, camiones: camiones };
        item.pmp = ofi + aerot + res;
    });

    const modalElement = document.getElementById('editModal');
    const modalInstance = bootstrap.Modal.getInstance(modalElement);
    if (modalInstance) modalInstance.hide();

    applyFilters();

    // Regenerar el reporte si el modal de reportes está abierto
    const messageModal = document.getElementById('messageModal');
    if (messageModal && messageModal.classList.contains('show')) {
        generateCustomReport();
    }
}

// Funciones para Report & Gestión Modal
function renderReportCrud() {
    const tbody = document.getElementById('crudTableBody');
    if (!tbody) return;

    // Obtener el rango de fecha/hora del reporte
    const dStart = document.getElementById('reportStartDate').value;
    const tStart = document.getElementById('reportStartTime').value || "00:00";
    const dEnd = document.getElementById('reportEndDate').value;
    const tEnd = document.getElementById('reportEndTime').value || "23:59";

    const filterStart = parseLocalDate(dStart, tStart);
    const filterEnd = parseLocalDate(dEnd, tEnd);

    // Filtrar datos ejecutados en el rango
    let crudData = rawData.filter(item => {
        if (!filterStart || !filterEnd) return true; // Si no hay filtro, mostrar todo

        const isExecuted = (item.startDate.getTime() <= filterEnd.getTime() && item.endDate.getTime() >= filterStart.getTime());
        return isExecuted;
    });

    // Agrupar por tipo, provincia, cantón, parroquia y fecha (similar a la tabla principal)
    const groups = {};
    crudData.forEach(item => {
        const dateKey = item.fechaPlanificacion.toISOString().split('T')[0];
        const groupKey = `${item.tipoOp}|${item.provincia}|${item.canton}|${item.parroquia}|${dateKey}`.toUpperCase();

        if (!groups[groupKey]) {
            groups[groupKey] = {
                tipoOp: item.tipoOp,
                provincia: item.provincia,
                canton: item.canton,
                parroquia: item.parroquia,
                sector: item.sector,
                fechaPlanificacion: item.fechaPlanificacion,
                horaMilitar: item.horaMilitar,
                sumPlanif: 0,
                sumEjecut: 0,
                sumPmp: 0,
                oficiales: 0,
                aerotecnicos: 0,
                reservistas: 0,
                camionetas: 0,
                buses: 0,
                camiones: 0,
                items: []
            };
        }

        groups[groupKey].sumPlanif++;
        if (item.startDate.getTime() <= filterEnd.getTime() && item.endDate.getTime() >= filterStart.getTime()) {
            groups[groupKey].sumEjecut++;
        }
        groups[groupKey].sumPmp += item.pmp;
        groups[groupKey].oficiales += (item.personal?.oficiales || 0);
        groups[groupKey].aerotecnicos += (item.personal?.aerotecnicos || 0);
        groups[groupKey].reservistas += (item.personal?.reservistas || 0);
        groups[groupKey].camionetas += (item.medios?.camionetas || 0);
        groups[groupKey].buses += (item.medios?.buses || 0);
        groups[groupKey].camiones += (item.medios?.camiones || 0);
        groups[groupKey].items.push(item);
    });

    // Convertir a array y ordenar
    const groupedData = Object.values(groups).sort((a, b) => {
        if (a.tipoOp !== b.tipoOp) return a.tipoOp.localeCompare(b.tipoOp);
        if (a.provincia !== b.provincia) return a.provincia.localeCompare(b.provincia);
        if (a.canton !== b.canton) return a.canton.localeCompare(b.canton);
        return a.fechaPlanificacion - b.fechaPlanificacion;
    });

    // Poblar el filtro de tipos de operación
    const tiposUnicos = [...new Set(groupedData.map(g => g.tipoOp))].sort();
    const filterSelect = document.getElementById('crudFilterTipo');
    const currentFilter = filterSelect ? filterSelect.value : 'TODOS';

    if (filterSelect) {
        filterSelect.innerHTML = '<option value="TODOS">-- Todos los Tipos --</option>';
        tiposUnicos.forEach(tipo => {
            const option = document.createElement('option');
            option.value = tipo;
            option.textContent = tipo;
            if (tipo === currentFilter) option.selected = true;
            filterSelect.appendChild(option);
        });
    }

    // Aplicar filtro de tipo de operación
    const filteredGroupedData = currentFilter === 'TODOS'
        ? groupedData
        : groupedData.filter(g => g.tipoOp === currentFilter);

    tbody.innerHTML = '';
    filteredGroupedData.forEach(group => {
        const tr = document.createElement('tr');
        const firstItemId = group.items[0].id;
        tr.innerHTML = `
            <td>${group.fechaPlanificacion.toLocaleDateString()}</td>
            <td>${group.canton} / ${group.parroquia}</td>
            <td>${group.tipoOp}</td>
            <td>${group.sumPmp}</td>
            <td>O:${group.oficiales} A:${group.aerotecnicos} R:${group.reservistas}</td>
            <td>C:${group.camionetas} B:${group.buses} CM:${group.camiones}</td>
            <td class="text-center">
                <button class="btn btn-sm btn-info" onclick="openEditModal(${firstItemId})">Edit</button>
            </td>
        `;
        tbody.appendChild(tr);
    });
}

// ========== INICIALIZACIÓN DE EVENTOS ==========
dom.btnFilter.addEventListener('click', applyFilters);

// Función para resetear filtros
function resetFilters() {
    if (rawData.length === 0) {
        alert("No hay datos cargados. Por favor, cargue un archivo Excel primero.");
        return;
    }

    // Resetear filtro de tipo activo en gráficos
    activeTypeFilter = null;

    // Resetear campos de fecha a vacío (dd/mm/aaaa)
    dom.filterStart.value = "";
    dom.filterEnd.value = "";
    dom.filterTimeStart.value = "00:00";
    dom.filterTimeEnd.value = "23:59";

    // Resetear otros filtros
    dom.filterProvincia.value = "TODOS";
    dom.filterCanton.value = "TODOS";
    dom.filterTipo.value = "TODOS";
    document.getElementById('searchInput').value = "";

    // Resetear selector de periodo a "Ver Todas"
    dom.sheetSelector.value = "ALL";

    // Aplicar filtros
    applyFilters();

    console.log(`🔄 Filtros reseteados a valores por defecto`);
}

// Event listener para el botón de resetear
const btnReset = document.getElementById('btnReset');
if (btnReset) {
    btnReset.addEventListener('click', resetFilters);
}

// Vincular filtros con actualización automática de gráficos
dom.filterProvincia.addEventListener('change', applyFilters);
dom.filterCanton.addEventListener('change', applyFilters);
dom.filterTipo.addEventListener('change', applyFilters);

const searchInput = document.getElementById('searchInput');
if (searchInput) {
    searchInput.addEventListener('input', () => {
        // Usar un pequeño delay (debounce) para no saturar al escribir
        clearTimeout(window.searchTimeout);
        window.searchTimeout = setTimeout(applyFilters, 300);
    });
}

dom.btnMessage.addEventListener('click', () => {
    const modal = new bootstrap.Modal(document.getElementById('messageModal'));
    // Inicializar fechas del reporte con el filtro actual
    document.getElementById('reportStartDate').value = dom.filterStart.value;
    document.getElementById('reportEndDate').value = dom.filterEnd.value;
    generateCustomReport();
    modal.show();
});

document.addEventListener('DOMContentLoaded', () => {
    // Configurar fechas por defecto si es necesario
    const hoy = new Date().toISOString().split('T')[0];
    dom.filterStart.value = hoy;
    dom.filterEnd.value = hoy;

    // Inicializar fechas de la matriz de evaluación
    const matrizDateStart = document.getElementById('matrizDateStart');
    const matrizDateEnd = document.getElementById('matrizDateEnd');
    if (matrizDateStart) matrizDateStart.value = hoy;
    if (matrizDateEnd) matrizDateEnd.value = hoy;
});

// ========== MÓDULO: PANEL PARA MATRIZ DE EVALUACIÓN ==========

/**
 * Aplica filtros independientes para la Matriz de Evaluación
 * Filtra por rango de tiempo y agrupa por Cantón
 */
function applyMatrizFilters() {
    console.log('🔷 MATRIZ DE EVALUACIÓN: Aplicando filtros...');

    // Obtener valores de los filtros independientes (fecha completa + hora)
    const dateStart = document.getElementById('matrizDateStart').value;
    const timeStart = document.getElementById('matrizTimeStart').value;
    const dateEnd = document.getElementById('matrizDateEnd').value;
    const timeEnd = document.getElementById('matrizTimeEnd').value;

    if (!dateStart || !timeStart || !dateEnd || !timeEnd) {
        alert('Por favor, configure fecha y hora de inicio y fin');
        return;
    }

    const filterStart = parseLocalDate(dateStart, timeStart);
    const filterEnd = parseLocalDate(dateEnd, timeEnd);

    if (!filterStart || !filterEnd) {
        alert('Error al parsear las fechas. Verifique los valores.');
        return;
    }

    console.log(`   Rango: ${filterStart.toLocaleString()} - ${filterEnd.toLocaleString()}`);

    // Filtrar operaciones ejecutadas en el rango de tiempo
    const filtered = rawData.filter(item => {
        // Verificar solapamiento de horarios
        const isExecuted = (item.startDate.getTime() <= filterEnd.getTime() &&
            item.endDate.getTime() >= filterStart.getTime());
        return isExecuted;
    });

    console.log(`   ✅ Operaciones filtradas: ${filtered.length}`);

    // Agrupar SOLO por Cantón
    const groups = {};

    filtered.forEach(item => {
        const key = item.canton; // Solo cantón

        if (!groups[key]) {
            groups[key] = {
                canton: item.canton,
                count: 0,
                pmp: 0,
                oficiales: 0,
                aerotecnicos: 0,
                reservistas: 0,
                tipos: {} // Desglose por tipo dentro del cantón
            };
        }

        groups[key].count++;
        groups[key].pmp += item.pmp;
        groups[key].oficiales += (item.personal?.oficiales || 0);
        groups[key].aerotecnicos += (item.personal?.aerotecnicos || 0);
        groups[key].reservistas += (item.personal?.reservistas || 0);

        // Contar por tipo dentro del cantón
        const tipo = item.tipoOp;
        if (!groups[key].tipos[tipo]) {
            groups[key].tipos[tipo] = 0;
        }
        groups[key].tipos[tipo]++;
    });

    const groupedData = Object.values(groups);
    console.log(`   📊 Grupos creados: ${groupedData.length}`);

    // Actualizar contador total
    document.getElementById('matrizTotalOps').textContent = filtered.length;

    // Renderizar tarjetas
    renderMatrizCards(groupedData);
}

/**
 * Renderiza las tarjetas dinámicamente en el grid
 * @param {Array} groups - Array de objetos con datos agrupados
 */
function renderMatrizCards(groups) {
    const container = document.getElementById('matrizCardsContainer');

    if (!container) return;

    // Limpiar contenedor
    container.innerHTML = '';

    if (groups.length === 0) {
        container.innerHTML = `
            <div class="col-12 text-center text-muted py-5">
                <span class="material-icons" style="font-size: 48px; opacity: 0.5;">search_off</span>
                <p class="mt-2 opacity-75">No se encontraron operaciones en el rango especificado</p>
            </div>
        `;
        return;
    }

    // Generar tarjetas por Cantón
    groups.forEach(group => {
        const card = document.createElement('div');
        card.className = 'matriz-card';

        // Generar lista de tipos de operación
        let tiposHtml = '';
        for (const [tipo, count] of Object.entries(group.tipos)) {
            const tipoColor = getTipoColor(tipo);
            tiposHtml += `
                <div class="tipo-item" style="border-left: 3px solid ${tipoColor};">
                    <span class="tipo-name">${tipo}</span>
                    <span class="tipo-count">${count}</span>
                </div>
            `;
        }

        card.innerHTML = `
            <div class="card-header" style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);">
                <h6 class="mb-0 text-white">
                    <span class="material-icons" style="font-size: 16px;">location_on</span>
                    ${group.canton}
                </h6>
            </div>
            <div class="card-body">
                <div class="stats-row mb-3">
                    <div class="stat-item">
                        <span class="stat-label">Operaciones</span>
                        <span class="stat-value">${group.count}</span>
                    </div>
                    <div class="stat-item">
                        <span class="stat-label">PMP</span>
                        <span class="stat-value">${group.pmp}</span>
                    </div>
                </div>
                <div class="personal-row mb-3">
                    <small class="text-muted">
                        OF: ${group.oficiales} | AE: ${group.aerotecnicos} | RE: ${group.reservistas}
                    </small>
                </div>
                <div class="tipos-desglose">
                    <strong class="d-block mb-2" style="font-size: 0.75rem; color: #6c757d;">Tipos de Operación:</strong>
                    ${tiposHtml}
                </div>
            </div>
        `;

        container.appendChild(card);
    });

    console.log(`   ✅ ${groups.length} tarjetas renderizadas`);
}

/**
 * Obtiene un color según el tipo de operación
 * @param {string} tipo - Tipo de operación
 * @returns {string} Color en formato hexadecimal
 */
function getTipoColor(tipo) {
    const tipoUpper = tipo.toUpperCase();

    if (tipoUpper.includes('RASTRILLAJE')) return '#667eea';
    if (tipoUpper.includes('COMBATE')) return '#764ba2';
    if (tipoUpper.includes('ARS') || tipoUpper.includes('PROTECCIÓN')) return '#f093fb';
    if (tipoUpper.includes('CAMEX') || tipoUpper.includes('ARMAS')) return '#4facfe';
    if (tipoUpper.includes('EJES') || tipoUpper.includes('VIAL')) return '#43e97b';

    return '#6c757d'; // Color por defecto
}
