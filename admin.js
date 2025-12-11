// Verificar sesión y permisos de administrador
document.addEventListener('DOMContentLoaded', () => {
    const userData = sessionStorage.getItem('vacaciones_user');
    if (!userData) {
        window.location.href = 'index.html';
        return;
    }

    const user = JSON.parse(userData);
    if (!(user.admin === 1 || user.admin === '1')) {
        window.location.href = 'calendar.html';
        return;
    }

    // Inicializar la aplicación de administración
    initAdminApp();
});

// Variables globales
let employees = [];
let changeMap = {};
let dataTable = null;
let filteredEmployees = [];

// Elementos DOM
const DOM = {
    tbody: document.getElementById('employees-tbody'),
    reloadBtn: document.getElementById('reload-btn'),
    logoutBtn: document.getElementById('logout-btn'),
    empresaFilter: document.getElementById('empresa-filter'),
    exportExcelBtn: document.getElementById('export-excel-btn')
};

// Función para cerrar sesión
function handleLogout() {
    sessionStorage.removeItem('vacaciones_user');
    window.location.href = 'index.html';
}

// Inicializar la aplicación de administración
function initAdminApp() {
    // Inicializar Firebase
    firebase.initializeApp(firebaseConfig);
    const db = firebase.firestore();

    // Configurar eventos
    DOM.logoutBtn.onclick = handleLogout;
    DOM.reloadBtn.onclick = () => loadEmployees(db);
    DOM.exportExcelBtn.onclick = handleExportExcel;

    // Cargar empleados
    loadEmployees(db);
}

// Manejar exportación a Excel
function handleExportExcel() {
    const dataToExport = filteredEmployees.length > 0 ? filteredEmployees : employees;
    exportToExcel(dataToExport, DOM.exportExcelBtn);
}

// Cargar lista de empleados
async function loadEmployees(db) {
    DOM.tbody.innerHTML = `<tr><td colspan="23" style="text-align:center;padding:20px">Cargando...</td></tr>`;
    try {
        const snap = await db.collection('empleados').orderBy('nombre').get();
        employees = snap.docs
            .map(d => ({ login: d.id, ...d.data() }))
            .filter(e => !(e.admin && Number(e.admin) === 1));
        
        renderEmpresaFilter();
        renderTable(employees);
    } catch (err) {
        DOM.tbody.innerHTML = `<tr><td colspan="23" style="color:#f44336;padding:20px;text-align:center">Error: ${escapeHtml(err.message)}</td></tr>`;
    }
}

// Renderizar filtro de empresas
function renderEmpresaFilter() {
    const empresas = [...new Set(employees.map(e => e.empresa?.trim()).filter(Boolean))].sort();
    DOM.empresaFilter.innerHTML = 
        `<option value="">Todas las empresas</option>` +
        empresas.map(e => `<option value="${escapeHtml(e)}">${escapeHtml(e)}</option>`).join('');
    
    DOM.empresaFilter.onchange = filtraPorEmpresa;
}

// Filtrar por empresa
function filtraPorEmpresa() {
    const val = DOM.empresaFilter.value;
    filteredEmployees = val ? employees.filter(e => e.empresa?.trim() === val) : [...employees];
    
    if (dataTable) {
        dataTable.clear();
        dataTable.rows.add(filteredEmployees.map(renderRowData));
        dataTable.draw();
        attachRowEvents();
    } else {
        renderTable(filteredEmployees);
    }
}

// Preparar datos para la fila
function renderRowData(emp) {
    return emp;
}

// Renderizar tabla con DataTables
function renderTable(data) {
    if (!data.length) {
        DOM.tbody.innerHTML = `<tr><td colspan="23" style="padding:20px;text-align:center;color:#9e9e9e">No hay empleados</td></tr>`;
        if (dataTable) dataTable.destroy();
        return;
    }

    if (!dataTable) {
        dataTable = $('#employees-table').DataTable({
            paging: true,
            pageLength: 10,
            lengthMenu: [10, 20, 40],
            searching: true,
            ordering: true,
            info: true,
            autoWidth: false,
            data: data,
            columns: [
                { 
                    data: 'nombre', 
                    render: (d, _, r) => `<span class="td-nombre" title="${escapeHtml(d || '')}">${escapeHtml(d || '')}</span>` 
                },
                { 
                    data: 'dni', 
                    render: d => escapeHtml(d || '') 
                },
                { 
                    data: 'empresa', 
                    render: d => escapeHtml(d || '') 
                },
                // Columna para tipo (select editable)
                { 
                    data: 'tipo', 
                    render: (d, _, r) => {
                        const tipo = d || '1';
                        return `
                            <select class="tipo-select" data-login="${r.login}">
                                <option value="1" ${tipo === '1' || tipo === 1 ? 'selected' : ''}>15 días</option>
                                <option value="2" ${tipo === '2' || tipo === 2 ? 'selected' : ''}>7 días</option>
                            </select>
                        `;
                    }
                },
                // Período 1 - EMPIEZA AQUÍ DESPUÉS DEL TIPO
                { 
                    data: null, 
                    render: (_, __, r) => `<input class="input-range calendar-input" data-type="start" data-period="1" value="${r.per1start || ''}">` 
                },
                { 
                    data: null, 
                    render: (_, __, r) => `<input class="input-range calendar-input" data-type="end" data-period="1" value="${r.per1end || ''}">` 
                },
                { 
                    data: null, 
                    className: 'cell-clear col-narrow', 
                    render: (_, __, r) => `<button class="clear-btn" data-period="1" title="Limpiar"><span class="material-icons" style="font-size:16px">clear</span></button>` 
                },
                { 
                    data: null, 
                    className: 'cell-exc', 
                    render: (_, __, r) => {
                        const exc = calcOver(r, 1);
                        return exc ? `<span class="${exc.startsWith('+') ? 'exc-pos' : 'exc-neg'}">${exc}</span>` : '';
                    }
                },
                // Período 2
                { 
                    data: null, 
                    render: (_, __, r) => `<input class="input-range calendar-input" data-type="start" data-period="2" value="${r.per2start || ''}">` 
                },
                { 
                    data: null, 
                    render: (_, __, r) => `<input class="input-range calendar-input" data-type="end" data-period="2" value="${r.per2end || ''}">` 
                },
                { 
                    data: null, 
                    className: 'cell-clear col-narrow', 
                    render: (_, __, r) => `<button class="clear-btn" data-period="2" title="Limpiar"><span class="material-icons" style="font-size:16px">clear</span></button>` 
                },
                { 
                    data: null, 
                    className: 'cell-exc', 
                    render: (_, __, r) => {
                        const exc = calcOver(r, 2);
                        return exc ? `<span class="${exc.startsWith('+') ? 'exc-pos' : 'exc-neg'}">${exc}</span>` : '';
                    }
                },
                // Período 3
                { 
                    data: null, 
                    render: (_, __, r) => `<input class="input-range calendar-input" data-type="start" data-period="3" value="${r.per3start || ''}">` 
                },
                { 
                    data: null, 
                    render: (_, __, r) => `<input class="input-range calendar-input" data-type="end" data-period="3" value="${r.per3end || ''}">` 
                },
                { 
                    data: null, 
                    className: 'cell-clear col-narrow', 
                    render: (_, __, r) => `<button class="clear-btn" data-period="3" title="Limpiar"><span class="material-icons" style="font-size:16px">clear</span></button>` 
                },
                { 
                    data: null, 
                    className: 'cell-exc', 
                    render: (_, __, r) => {
                        const exc = calcOver(r, 3);
                        return exc ? `<span class="${exc.startsWith('+') ? 'exc-pos' : 'exc-neg'}">${exc}</span>` : '';
                    }
                },
                // Período 4
                { 
                    data: null, 
                    render: (_, __, r) => `<input class="input-range calendar-input" data-type="start" data-period="4" value="${r.per4start || ''}">` 
                },
                { 
                    data: null, 
                    render: (_, __, r) => `<input class="input-range calendar-input" data-type="end" data-period="4" value="${r.per4end || ''}">` 
                },
                { 
                    data: null, 
                    className: 'cell-clear col-narrow', 
                    render: (_, __, r) => `<button class="clear-btn" data-period="4" title="Limpiar"><span class="material-icons" style="font-size:16px">clear</span></button>` 
                },
                { 
                    data: null, 
                    className: 'cell-exc', 
                    render: (_, __, r) => {
                        const exc = calcOver(r, 4);
                        return exc ? `<span class="${exc.startsWith('+') ? 'exc-pos' : 'exc-neg'}">${exc}</span>` : '';
                    }
                },
                { 
                    data: 'login', 
                    render: (d) => 
                        `<button class="mdl-button mdl-js-button mdl-button--raised mdl-button--colored save-btn" data-login="${d}">Guardar</button>` 
                }
            ],
            language: {
                url: "https://cdn.datatables.net/plug-ins/2.0.8/i18n/es-ES.json"
            },
            columnDefs: [
                { targets: [0, 1, 2, 3, 4, 5, 6], orderable: true },
                { targets: '_all', className: 'mdl-data-table__cell--non-numeric' }
            ],
            dom: 
                "<'dt-top clearfix'<'dt-search'f>>" +
                "<'dt-table'tr>" +
                "<'dt-bottom clearfix'<'dt-info'i><'dt-pag'p>>"
        });

        // Vincular eventos después de cada renderizado
        $('#employees-table').on('draw.dt', attachRowEvents);
        attachRowEvents();
    } else {
        dataTable.clear();
        dataTable.rows.add(data);
        dataTable.draw();
        attachRowEvents();
    }
}

// Vincular eventos a las filas
function attachRowEvents() {
    const rows = document.querySelectorAll('#employees-table tbody tr');
    rows.forEach(row => {
        const login = row.querySelector('.save-btn')?.dataset.login || row.dataset.login;
        if (!login) return;

        // Configurar eventos para el select de tipo
        const tipoSelect = row.querySelector('.tipo-select');
        if (tipoSelect) {
            tipoSelect.addEventListener('change', () => {
                markChanged(row, login);
                // Recalcular todos los períodos cuando cambia el tipo
                for (let i = 1; i <= 4; i++) {
                    updateExc(row, login, i);
                }
            });
        }

        // Configurar eventos para cada período
        for (let i = 1; i <= 4; i++) {
            const start = row.querySelector(`[data-type="start"][data-period="${i}"]`);
            const end = row.querySelector(`[data-type="end"][data-period="${i}"]`);
            const clearBtn = row.querySelector(`.clear-btn[data-period="${i}"]`);
            
            // Inicializar Flatpickr si no está ya inicializado
            if (start && !start._flatpickr) {
                initFlatpickrStart(start, end, login, i, row);
            }
            if (end && !end._flatpickr) {
                initFlatpickrEnd(end, start, login, i, row);
            }
            
            // Evento para limpiar período
            if (clearBtn) {
                clearBtn.addEventListener('click', e => {
                    e.stopPropagation();
                    clearPeriod(row, login, i);
                });
            }
            
            // Eventos de cambio para inputs
            if (start) {
                start.addEventListener('input', () => {
                    markChanged(row, login);
                    updateExc(row, login, i);
                    // Actualizar minDate del input de fin si existe
                    updateEndMinDate(row, login, i);
                });
            }
            if (end) {
                end.addEventListener('input', () => {
                    markChanged(row, login);
                    updateExc(row, login, i);
                });
            }
        }
        
        // Evento para guardar
        const saveBtn = row.querySelector('button[data-login]');
        if (saveBtn) {
            saveBtn.addEventListener('click', e => {
                e.stopPropagation();
                saveRow(row, login);
            });
        }
    });
}

// Actualizar minDate del input de fin
function updateEndMinDate(row, login, periodIndex) {
    const start = row.querySelector(`[data-type="start"][data-period="${periodIndex}"]`);
    const end = row.querySelector(`[data-type="end"][data-period="${periodIndex}"]`);
    
    if (start && end && end._flatpickr) {
        const startValue = start.value;
        if (startValue) {
            // Establecer minDate como la fecha de inicio
            end._flatpickr.set('minDate', startValue);
        } else {
            // Si no hay fecha de inicio, establecer minDate a 2026-01-01
            end._flatpickr.set('minDate', "2026-01-01");
        }
    }
}

// Inicializar Flatpickr para input de inicio
function initFlatpickrStart(input, endInput, login, periodIndex, row) {
    if (!input || input._flatpickr) return;
    
    // Configuración básica de Flatpickr para inicio
    const config = {
        dateFormat: "Y/m/d",
        allowInput: true,
        locale: "es",
        minDate: "2026-01-01",
        maxDate: "2026-12-31",
        onChange: function(selectedDates, dateStr) {
            input.value = dateStr;
            markChanged(row, login);
            updateExc(row, login, periodIndex);
            
            // Actualizar minDate del input de fin
            updateEndMinDate(row, login, periodIndex);
        }
    };
    
    // Solo establecer fecha por defecto si el input ya tiene un valor
    if (input.value && input.value.trim() !== '') {
        config.defaultDate = input.value;
    }
    
    flatpickr(input, config);
}

// Inicializar Flatpickr para input de fin
function initFlatpickrEnd(input, startInput, login, periodIndex, row) {
    if (!input || input._flatpickr) return;
    
    // Configuración básica de Flatpickr para fin
    const config = {
        dateFormat: "Y/m/d",
        allowInput: true,
        locale: "es",
        minDate: "2026-01-01", // Por defecto, se actualizará si hay fecha de inicio
        maxDate: "2026-12-31",
        onChange: function(selectedDates, dateStr) {
            input.value = dateStr;
            markChanged(row, login);
            updateExc(row, login, periodIndex);
        }
    };
    
    // Si hay fecha de inicio, establecerla como minDate
    if (startInput && startInput.value && startInput.value.trim() !== '') {
        config.minDate = startInput.value;
    }
    
    // Solo establecer fecha por defecto si el input ya tiene un valor
    if (input.value && input.value.trim() !== '') {
        config.defaultDate = input.value;
    }
    
    flatpickr(input, config);
}

// Marcar fila como modificada
function markChanged(row, login) {
    changeMap[login] = true;
    const btn = row.querySelector('.save-btn');
    btn?.classList.add('save-changed');
}

// Actualizar columna de exceso/deficiencia
function updateExc(row, login, periodIndex) {
    const start = row.querySelector(`[data-type="start"][data-period="${periodIndex}"]`)?.value || '';
    const end = row.querySelector(`[data-type="end"][data-period="${periodIndex}"]`)?.value || '';
    
    let exc = '';
    if (start && end) {
        const tipo = detectTipoFromRow(row);
        const diff = calcOverCustom(start, end, tipo);
        if (diff !== 0) exc = (diff > 0 ? '+' : '') + diff;
    }
    
    const cell = row.querySelector(`.cell-exc:nth-of-type(${periodIndex})`);
    if (cell) {
        cell.innerHTML = exc ? `<span class="${exc.startsWith('+') ? 'exc-pos' : 'exc-neg'}">${exc}</span>` : '';
    }
}

// Detectar tipo de empleado desde la fila (select)
function detectTipoFromRow(row) {
    const tipoSelect = row.querySelector('.tipo-select');
    return tipoSelect ? Number(tipoSelect.value) : 1;
}

// Detectar tipo de empleado (para uso con objeto empleado)
function detectTipo(login) {
    const emp = employees.find(e => e.login === login);
    return emp?.tipo === 2 || emp?.tipo === '2' ? 2 : 1;
}

// Limpiar período
function clearPeriod(row, login, periodIndex) {
    const start = row.querySelector(`[data-type="start"][data-period="${periodIndex}"]`);
    const end = row.querySelector(`[data-type="end"][data-period="${periodIndex}"]`);
    
    if (start) {
        start.value = '';
        if (start._flatpickr) {
            start._flatpickr.clear();
        }
    }
    if (end) {
        end.value = '';
        if (end._flatpickr) {
            end._flatpickr.clear();
            // Resetear minDate a 2026-01-01
            end._flatpickr.set('minDate', "2026-01-01");
        }
    }
    
    markChanged(row, login);
    updateExc(row, login, periodIndex);
}

// Guardar cambios de una fila - CON VALIDACIÓN DE FECHAS
async function saveRow(row, login) {
    const db = firebase.firestore();
    row.classList.add('row-saving');
    
    try {
        const updates = {};
        
        // Recopilar datos del tipo
        const tipoSelect = row.querySelector('.tipo-select');
        if (tipoSelect) {
            const tipoValue = tipoSelect.value;
            updates['tipo'] = tipoValue === '1' ? 1 : 2;
        }
        
        // Recopilar datos de cada período y validar fechas
        for (let i = 1; i <= 4; i++) {
            const start = row.querySelector(`[data-type="start"][data-period="${i}"]`)?.value || '';
            const end = row.querySelector(`[data-type="end"][data-period="${i}"]`)?.value || '';
            
            // Validar que si hay fechas, ambas estén presentes
            if ((start && !end) || (!start && end)) {
                alert(`Error en período ${i}: Debe completar ambas fechas o dejar ambas vacías.`);
                row.classList.remove('row-saving');
                return;
            }
            
            if (start && end) {
                // Validar que la fecha de fin sea mayor que la de inicio
                const startDate = parseDate(start);
                const endDate = parseDate(end);
                
                if (!startDate || !endDate) {
                    alert(`Error en período ${i}: Formato de fecha inválido.`);
                    row.classList.remove('row-saving');
                    return;
                }
                
                if (endDate <= startDate) {
                    alert(`Error en período ${i}: La fecha de fin debe ser posterior a la fecha de inicio.`);
                    row.classList.remove('row-saving');
                    return;
                }
                
                updates[`per${i}start`] = start.replace(/\//g, '-');
                updates[`per${i}end`] = end.replace(/\//g, '-');
            } else {
                updates[`per${i}start`] = firebase.firestore.FieldValue.delete();
                updates[`per${i}end`] = firebase.firestore.FieldValue.delete();
            }
        }
        
        // Guardar en Firestore
        await db.collection('empleados').doc(login).update(updates);
        
        // Actualizar datos locales
        const doc = await db.collection('empleados').doc(login).get();
        const fresh = { login: doc.id, ...doc.data() };
        const idx = employees.findIndex(e => e.login === login);
        
        if (idx >= 0) {
            employees[idx] = fresh;
            
            // Actualizar en la tabla filtrada si existe
            const filteredIdx = filteredEmployees.findIndex(e => e.login === login);
            if (filteredIdx >= 0) {
                filteredEmployees[filteredIdx] = fresh;
            }
        }
        
        // Actualizar DataTable
        if (dataTable) {
            const rowIndex = dataTable.row(row).index();
            dataTable.row(rowIndex).data(fresh).draw(false);
            
            // Re-vincular eventos después de un breve delay
            setTimeout(() => attachRowEvents(), 50);
        }
        
        // Quitar marca de cambio
        delete changeMap[login];
        
        // Quitar clase de cambio del botón
        const saveBtn = row.querySelector('.save-btn');
        if (saveBtn) saveBtn.classList.remove('save-changed');
        
    } catch (err) {
        alert(`Error al guardar: ${err.message}`);
    } finally {
        row.classList.remove('row-saving');
    }
}

// ===== FUNCIONES HELPER =====

// Calcular exceso/deficiencia para un período
function calcOver(emp, i) {
    const s = emp[`per${i}start`];
    const e = emp[`per${i}end`];
    if (!s || !e) return '';
    
    const days = daysInclusive(s, e);
    const tipo = emp.tipo === 2 || emp.tipo === '2' ? 2 : 1;
    const base = tipo === 1 ? 15 : 7;
    const over = days - base;
    return over === 0 ? '' : over > 0 ? `+${over}` : String(over);
}

// Calcular exceso/deficiencia personalizado
function calcOverCustom(s, e, tipo) {
    const d1 = parseDate(s);
    const d2 = parseDate(e);
    if (!d1 || !d2) return 0;
    
    const base = tipo === 1 ? 15 : 7;
    return Math.floor((d2 - d1) / 86400000) + 1 - base;
}

// Calcular días entre dos fechas (inclusive)
function daysInclusive(s, e) {
    const d1 = parseDate(s);
    const d2 = parseDate(e);
    return d1 && d2 ? Math.floor((d2 - d1) / 86400000) + 1 : 0;
}

// Parsear fecha
function parseDate(str) {
    if (!str) return null;
    
    // Convertir formato Y/m/d o Y-m-d a Date
    const cleanStr = str.replace(/\//g, '-');
    const [year, month, day] = cleanStr.split('-').map(Number);
    
    if (!year || !month || !day || isNaN(year) || isNaN(month) || isNaN(day)) {
        return null;
    }
    
    const date = new Date(year, month - 1, day);
    return isNaN(date.getTime()) ? null : date;
}

// Escapar HTML para prevenir XSS
function escapeHtml(s) {
    if (s == null) return '';
    
    return String(s)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#039;');
}