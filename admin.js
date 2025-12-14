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
    setupGroupDialog();
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
    exportExcelBtn: document.getElementById('export-excel-btn'),
    groupDialog: document.getElementById('group-info-dialog'),
    incompatibleTbody: document.getElementById('incompatible-tbody')
};

function setupGroupDialog() {
    const closeBtn = DOM.groupDialog.querySelector('.close-dialog');
    const closeIconBtn = DOM.groupDialog.querySelector('.close-dialog-btn');
    if (closeBtn) {
        closeBtn.addEventListener('click', () => { DOM.groupDialog.close(); });
    }
    if (closeIconBtn) {
        closeIconBtn.addEventListener('click', () => { DOM.groupDialog.close(); });
    }
    DOM.groupDialog.addEventListener('click', (e) => {
        if (e.target === DOM.groupDialog) DOM.groupDialog.close();
    });
    document.addEventListener('keydown', (e) => {
        if (e.key === 'Escape' && DOM.groupDialog.open) DOM.groupDialog.close();
    });
}

function handleLogout() {
    sessionStorage.removeItem('vacaciones_user');
    window.location.href = 'index.html';
}

function initAdminApp() {
    firebase.initializeApp(firebaseConfig);
    const db = firebase.firestore();
    DOM.logoutBtn.onclick = handleLogout;
    DOM.reloadBtn.onclick = () => loadEmployees(db);
    DOM.exportExcelBtn.onclick = handleExportExcel;
    loadEmployees(db);
}

function handleExportExcel() {
    const dataToExport = filteredEmployees.length > 0 ? filteredEmployees : employees;
    exportToExcel(dataToExport, DOM.exportExcelBtn);
}

async function loadEmployees(db) {
    DOM.tbody.innerHTML = `<tr><td colspan="25" style="text-align:center;padding:20px">Cargando...</td></tr>`;
    try {
        const snap = await db.collection('empleados').orderBy('nombre').get();
        employees = snap.docs
            .map(d => ({ login: d.id, ...d.data() }))
            .filter(e => !(e.admin && Number(e.admin) === 1));
        renderEmpresaFilter();
        renderTable(employees);
    } catch (err) {
        DOM.tbody.innerHTML = `<tr><td colspan="25" style="color:#f44336;padding:20px;text-align:center">Error: ${escapeHtml(err.message)}</td></tr>`;
    }
}

function renderEmpresaFilter() {
    const empresas = [...new Set(employees.map(e => e.empresa?.trim()).filter(Boolean))].sort();
    DOM.empresaFilter.innerHTML = 
        `<option value="">Todas las empresas</option>` +
        empresas.map(e => `<option value="${escapeHtml(e)}">${escapeHtml(e)}</option>`).join('');
    DOM.empresaFilter.onchange = filtraPorEmpresa;
}

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

function showGroupInfo(login) {
    const emp = employees.find(e => e.login === login);
    if (!emp) return;
    
    document.getElementById('dialog-employee-name').textContent = emp.nombre || 'N/A';
    document.getElementById('dialog-employee-empresa').textContent = emp.empresa ? `Empresa: ${emp.empresa}` : 'Empresa: N/A';
    
    // Construir la información de grupo
    let groupInfo = '';
    if (emp.grupo) {
        groupInfo += `Grupo: ${emp.grupo}`;
    }
    if (emp.subgrupo) {
        groupInfo += groupInfo ? ` | Subgrupo: ${emp.subgrupo}` : `Subgrupo: ${emp.subgrupo}`;
    }
    if (!emp.grupo && !emp.subgrupo) {
        groupInfo = 'Sin grupo/subgrupo';
    }
    document.getElementById('dialog-employee-grupo').textContent = groupInfo;
    
    // Añadir información de exclusiones - MODIFICADO
    const tieneExclusiones = emp.exclusiones && emp.exclusiones.trim() !== '';
    const exclusionElement = document.getElementById('dialog-employee-subgrupo');
    
    if (tieneExclusiones) {
        exclusionElement.innerHTML = `<i class="material-icons" style="font-size:14px">warning</i> Con exclusiones: ${escapeHtml(emp.exclusiones)}`;
        exclusionElement.className = 'employee-group-item exclusion-warning';
        exclusionElement.style.display = 'block';
    } else {
        // Si no tiene exclusiones, ocultar completamente el elemento
        exclusionElement.innerHTML = '';
        exclusionElement.className = '';
        exclusionElement.style.display = 'none';
    }
    
    const incompatibleEmployees = findIncompatibleEmployees(emp);
    updateIncompatibleTable(incompatibleEmployees);
    DOM.groupDialog.showModal();
}

function updateIncompatibleTable(incompatibleEmployees) {
    if (!DOM.incompatibleTbody) return;
    if (incompatibleEmployees.length === 0) {
        DOM.incompatibleTbody.innerHTML = `
            <tr>
                <td colspan="3" style="text-align:center; padding:15px; color:#666;">
                    <i class="material-icons" style="color:#4CAF50; font-size:24px; display:block; margin-bottom:8px;">check_circle</i>
                    Este empleado no tiene restricciones de grupo
                </td>
            </tr>
        `;
    } else {
        let html = '';
        incompatibleEmployees.forEach(emp => {
            const grupoInfo = emp.grupo || 'N/A';
            const subgrupoInfo = emp.subgrupo ? `/ ${emp.subgrupo}` : '';
            // Verificar si este empleado incompatible tiene exclusiones - MODIFICADO
            const tieneExclusiones = emp.exclusiones && emp.exclusiones.trim() !== '';
            const exclusionIcon = tieneExclusiones ? ' <i class="material-icons" style="font-size:14px; color:#f44336; vertical-align:middle" title="Con exclusiones">warning</i>' : '';
            
            html += `
                <tr>
                    <td class="employee-name">${escapeHtml(emp.nombre || emp.login)}${exclusionIcon}</td>
                    <td class="employee-empresa">${escapeHtml(emp.empresa || 'N/A')}</td>
                    <td class="employee-grupo">${escapeHtml(grupoInfo)} ${escapeHtml(subgrupoInfo)}</td>
                </tr>
            `;
        });
        DOM.incompatibleTbody.innerHTML = html;
    }
}

function findIncompatibleEmployees(selectedEmp) {
    if (!selectedEmp.grupo && !selectedEmp.subgrupo) return [];
    return employees.filter(emp => {
        if (emp.login === selectedEmp.login) return false;
        const sameGroup = selectedEmp.grupo && emp.grupo && selectedEmp.grupo === emp.grupo;
        let sameSubgroup = false;
        if (selectedEmp.subgrupo) {
            sameSubgroup = emp.subgrupo && selectedEmp.subgrupo === emp.subgrupo;
        }
        const groupSubgroupMatch = (selectedEmp.grupo && emp.subgrupo && selectedEmp.grupo === emp.subgrupo) ||
                                  (selectedEmp.subgrupo && emp.grupo && selectedEmp.subgrupo === emp.grupo);
        return sameGroup || sameSubgroup || groupSubgroupMatch;
    });
}

function renderRowData(emp) { return emp; }

function calculateTotalDays(emp) {
    let total = 0;
    for (let i = 1; i <= 4; i++) {
        const start = emp[`per${i}start`];
        const end = emp[`per${i}end`];
        if (start && end) {
            const days = daysInclusive(start, end);
            total += days;
        }
    }
    return total;
}

function updateTotalColumn(row) {
    const login = row.querySelector('.save-btn')?.dataset.login;
    if (!login) return;
    const emp = employees.find(e => e.login === login);
    if (!emp) return;
    let total = 0;
    for (let i = 1; i <= 4; i++) {
        const startInput = row.querySelector(`[data-type="start"][data-period="${i}"]`);
        const endInput = row.querySelector(`[data-type="end"][data-period="${i}"]`);
        if (startInput && endInput && startInput.value && endInput.value) {
            const start = startInput.value.replace(/\//g, '-');
            const end = endInput.value.replace(/\//g, '-');
            const days = daysInclusive(start, end);
            total += days;
        } else if (emp[`per${i}start`] && emp[`per${i}end`]) {
            const days = daysInclusive(emp[`per${i}start`], emp[`per${i}end`]);
            total += days;
        }
    }
    const totalCell = row.querySelector('.cell-total');
    if (totalCell) {
        const colorClass = total > 30 ? 'total-over' : 'total-normal';
        totalCell.innerHTML = `<span class="${colorClass}" title="Total de días de vacaciones">${total}</span>`;
    }
}

function determineRowStatus(row, login) {
    const tipoSelect = row.querySelector('.tipo-select');
    const tipo = tipoSelect ? Number(tipoSelect.value) : 1;
    let completedPeriods = 0;
    for (let i = 1; i <= 4; i++) {
        const start = row.querySelector(`[data-type="start"][data-period="${i}"]`)?.value || '';
        const end = row.querySelector(`[data-type="end"][data-period="${i}"]`)?.value || '';
        if (start && end) completedPeriods++;
    }
    row.classList.remove('row-complete', 'row-incomplete');
    if (tipo === 1) {
        disablePeriodInputs(row, 3);
        disablePeriodInputs(row, 4);
        const hasPeriod1 = row.querySelector(`[data-type="start"][data-period="1"]`)?.value && 
                          row.querySelector(`[data-type="end"][data-period="1"]`)?.value;
        const hasPeriod2 = row.querySelector(`[data-type="start"][data-period="2"]`)?.value && 
                          row.querySelector(`[data-type="end"][data-period="2"]`)?.value;
        if (hasPeriod1 && hasPeriod2) {
            row.classList.add('row-complete');
        } else if (completedPeriods > 0 && completedPeriods < 2) {
            row.classList.add('row-incomplete');
        }
    } else if (tipo === 2) {
        enablePeriodInputs(row, 3);
        enablePeriodInputs(row, 4);
        let allPeriodsComplete = true;
        for (let i = 1; i <= 4; i++) {
            const start = row.querySelector(`[data-type="start"][data-period="${i}"]`)?.value || '';
            const end = row.querySelector(`[data-type="end"][data-period="${i}"]`)?.value || '';
            if (!start || !end) {
                allPeriodsComplete = false;
                break;
            }
        }
        if (allPeriodsComplete) {
            row.classList.add('row-complete');
        } else if (completedPeriods > 0 && completedPeriods < 4) {
            row.classList.add('row-incomplete');
        }
    }
}

function disablePeriodInputs(row, periodIndex) {
    const startInput = row.querySelector(`[data-type="start"][data-period="${periodIndex}"]`);
    const endInput = row.querySelector(`[data-type="end"][data-period="${periodIndex}"]`);
    const clearBtn = row.querySelector(`.clear-btn[data-period="${periodIndex}"]`);
    if (startInput) {
        startInput.disabled = true;
        // Fix: check for _flatpickr, not just null, but also not already destroyed
        if (startInput._flatpickr && !startInput._flatpickr.destroyed && typeof startInput._flatpickr.destroy === 'function') {
            startInput._flatpickr.destroy();
            startInput._flatpickr = undefined;
        }
    }
    if (endInput) {
        endInput.disabled = true;
        if (endInput._flatpickr && !endInput._flatpickr.destroyed && typeof endInput._flatpickr.destroy === 'function') {
            endInput._flatpickr.destroy();
            endInput._flatpickr = undefined;
        }
    }
    if (clearBtn) { clearBtn.disabled = true; }
}

function enablePeriodInputs(row, periodIndex) {
    const startInput = row.querySelector(`[data-type="start"][data-period="${periodIndex}"]`);
    const endInput = row.querySelector(`[data-type="end"][data-period="${periodIndex}"]`);
    const clearBtn = row.querySelector(`.clear-btn[data-period="${periodIndex}"]`);
    if (startInput) {
        startInput.disabled = false;
        if (!startInput._flatpickr) {
            const endInput = row.querySelector(`[data-type="end"][data-period="${periodIndex}"]`);
            initFlatpickrStart(startInput, endInput, '', periodIndex, row);
        }
    }
    if (endInput) {
        endInput.disabled = false;
        if (!endInput._flatpickr) {
            const startInput = row.querySelector(`[data-type="start"][data-period="${periodIndex}"]`);
            initFlatpickrEnd(endInput, startInput, '', periodIndex, row);
        }
    }
    if (clearBtn) { clearBtn.disabled = false; }
}

function checkInitialStatus(emp) {
    const tipo = emp.tipo === 2 || emp.tipo === '2' ? 2 : 1;
    if (tipo === 1) {
        const hasPeriod1 = emp.per1start && emp.per1end;
        const hasPeriod2 = emp.per2start && emp.per2end;
        return {
            isComplete: hasPeriod1 && hasPeriod2,
            isIncomplete: (hasPeriod1 || hasPeriod2) && !(hasPeriod1 && hasPeriod2)
        };
    } else {
        const hasPeriod1 = emp.per1start && emp.per1end;
        const hasPeriod2 = emp.per2start && emp.per2end;
        const hasPeriod3 = emp.per3start && emp.per3end;
        const hasPeriod4 = emp.per4start && emp.per4end;
        const allPeriods = hasPeriod1 && hasPeriod2 && hasPeriod3 && hasPeriod4;
        const somePeriods = hasPeriod1 || hasPeriod2 || hasPeriod3 || hasPeriod4;
        return {
            isComplete: allPeriods,
            isIncomplete: somePeriods && !allPeriods
        };
    }
}

function renderTable(data) {
    if (!data.length) {
        DOM.tbody.innerHTML = `<tr><td colspan="25" style="padding:20px;text-align:center;color:#9e9e9e">No hay empleados</td></tr>`;
        if (dataTable) dataTable.destroy();
        return;
    }
    if (!dataTable) {
        dataTable = $('#employees-table').DataTable({
            paging: true,
            pageLength: 15,
            lengthMenu: [15, 30, 60, 100],
            searching: true,
            ordering: true,
            info: true,
            autoWidth: false,
            data: data,
            columns: [
                { 
                    data: 'nombre', 
                    className: 'td-nombre',
                    render: (d, _, r) => `<span class="td-nombre" title="${escapeHtml(d || '')}">${escapeHtml(d || '')}</span>` 
                },
                { 
                    data: 'dni', 
                    render: d => escapeHtml(d || '') 
                },
                { 
                    data: 'empresa', 
                    className: 'td-empresa',
                    render: function(d, type, row) {
                        if (!d) return '';
                        const empresa = escapeHtml(d);
                        // Recortar a 14 caracteres y añadir ...
                        const truncated = empresa.length > 14 ? empresa.substring(0, 14) + '...' : empresa;
                        return `<span title="${empresa}">${truncated}</span>`;
                    }
                },
                { 
                    data: null, 
                    className: 'cell-group', 
                    render: (_, __, r) => { 
                        const hasGroup = r.grupo || r.subgrupo; 
                        return hasGroup ? 
                            `<button class="group-info-btn" data-login="${r.login}" title="Ver trabajadores que no pueden coincidir"><i class="material-icons">group</i></button>` : 
                            `<span class="no-group" title="Sin restricciones de grupo">-</span>`; 
                    } 
                },
                { 
                    data: 'tipo', 
                    render: (d, _, r) => { 
                        const tipo = d || '1'; 
                        const initialStatus = checkInitialStatus(r); 
                        let initialClass = ''; 
                        if (initialStatus.isComplete) initialClass = 'row-complete'; 
                        else if (initialStatus.isIncomplete) initialClass = 'row-incomplete'; 
                        return `<select class="tipo-select" data-login="${r.login}">
                            <option value="1" ${tipo === '1' || tipo === 1 ? 'selected' : ''}>15 días</option>
                            <option value="2" ${tipo === '2' || tipo === 2 ? 'selected' : ''}>7 días</option>
                        </select>`; 
                    } 
                },
                // Períodos
                { data: null, render: (_, __, r) => `<input class="input-range calendar-input" data-type="start" data-period="1" value="${r.per1start || ''}">` },
                { data: null, render: (_, __, r) => `<input class="input-range calendar-input" data-type="end" data-period="1" value="${r.per1end || ''}">` },
                { data: null, className: 'cell-clear col-narrow', render: () => `<button class="clear-btn" data-period="1" title="Limpiar"><span class="material-icons" style="font-size:16px">clear</span></button>` },
                { data: null, className: 'cell-exc', render: (_, __, r) => { const exc = calcOver(r, 1); return exc ? `<span class="${exc.startsWith('+') ? 'exc-pos' : 'exc-neg'}">${exc}</span>` : ''; } },
                { data: null, render: (_, __, r) => `<input class="input-range calendar-input" data-type="start" data-period="2" value="${r.per2start || ''}">` },
                { data: null, render: (_, __, r) => `<input class="input-range calendar-input" data-type="end" data-period="2" value="${r.per2end || ''}">` },
                { data: null, className: 'cell-clear col-narrow', render: () => `<button class="clear-btn" data-period="2" title="Limpiar"><span class="material-icons" style="font-size:16px">clear</span></button>` },
                { data: null, className: 'cell-exc', render: (_, __, r) => { const exc = calcOver(r, 2); return exc ? `<span class="${exc.startsWith('+') ? 'exc-pos' : 'exc-neg'}">${exc}</span>` : ''; } },
                { data: null, render: (_, __, r) => `<input class="input-range calendar-input" data-type="start" data-period="3" value="${r.per3start || ''}">` },
                { data: null, render: (_, __, r) => `<input class="input-range calendar-input" data-type="end" data-period="3" value="${r.per3end || ''}">` },
                { data: null, className: 'cell-clear col-narrow', render: () => `<button class="clear-btn" data-period="3" title="Limpiar"><span class="material-icons" style="font-size:16px">clear</span></button>` },
                { data: null, className: 'cell-exc', render: (_, __, r) => { const exc = calcOver(r, 3); return exc ? `<span class="${exc.startsWith('+') ? 'exc-pos' : 'exc-neg'}">${exc}</span>` : ''; } },
                { data: null, render: (_, __, r) => `<input class="input-range calendar-input" data-type="start" data-period="4" value="${r.per4start || ''}">` },
                { data: null, render: (_, __, r) => `<input class="input-range calendar-input" data-type="end" data-period="4" value="${r.per4end || ''}">` },
                { data: null, className: 'cell-clear col-narrow', render: () => `<button class="clear-btn" data-period="4" title="Limpiar"><span class="material-icons" style="font-size:16px">clear</span></button>` },
                { data: null, className: 'cell-exc', render: (_, __, r) => { const exc = calcOver(r, 4); return exc ? `<span class="${exc.startsWith('+') ? 'exc-pos' : 'exc-neg'}">${exc}</span>` : ''; } },
                { data: null, className: 'cell-total', render: (_, __, r) => { const total = calculateTotalDays(r); const colorClass = total > 30 ? 'total-over' : 'total-normal'; return `<span class="${colorClass}" title="Total de días de vacaciones">${total}</span>`; } },
                { data: 'login', render: (d) => `<button class="mdl-button mdl-js-button mdl-button--raised mdl-button--colored save-btn" data-login="${d}">Guardar</button>` }
            ],
            language: { url: "https://cdn.datatables.net/plug-ins/2.0.8/i18n/es-ES.json" },
            columnDefs: [
                { targets: [0, 1, 2, 3, 4, 5, 6], orderable: true },
                { targets: '_all', className: 'mdl-data-table__cell--non-numeric'}
            ],
            dom: 
                "<'dt-top clearfix'<'dt-length'l><'dt-search'f>>" +
                "<'dt-table'tr>" +
                "<'dt-bottom clearfix'<'dt-info'i><'dt-pag'p>>"
        });
        $('#employees-table').on('draw.dt', function() {
            attachRowEvents();
            const rows = document.querySelectorAll('#employees-table tbody tr');
            rows.forEach(row => {
                const login = row.querySelector('.save-btn')?.dataset.login;
                if (login) determineRowStatus(row, login);
            });
        });
        $('#employees-table').on('change', '.tipo-select, .input-range', function() {
            const row = $(this).closest('tr')[0];
            if (row) updateTotalColumn(row);
        });
        attachRowEvents();
    } else {
        dataTable.clear();
        dataTable.rows.add(data);
        dataTable.draw();
        attachRowEvents();
    }
}

function attachRowEvents() {
    const rows = document.querySelectorAll('#employees-table tbody tr');
    rows.forEach(row => {
        const login = row.querySelector('.save-btn')?.dataset.login || row.dataset.login;
        if (!login) return;
        determineRowStatus(row, login);
        const groupBtn = row.querySelector('.group-info-btn');
        if (groupBtn) {
            const newGroupBtn = groupBtn.cloneNode(true);
            groupBtn.parentNode.replaceChild(newGroupBtn, groupBtn);
            newGroupBtn.addEventListener('click', (e) => {
                e.stopPropagation();
                e.preventDefault();
                showGroupInfo(login);
            });
        }
        const tipoSelect = row.querySelector('.tipo-select');
        if (tipoSelect) {
            const newTipoSelect = tipoSelect.cloneNode(true);
            tipoSelect.parentNode.replaceChild(newTipoSelect, tipoSelect);
            newTipoSelect.addEventListener('change', () => {
                changeMap[login] = true;
                const btn = row.querySelector('.save-btn');
                if (btn) { btn.classList.add('save-changed'); }
                for (let i = 1; i <= 4; i++) { updateExc(row, login, i); }
                updateTotalColumn(row);
                determineRowStatus(row, login);
            });
        }
        for (let i = 1; i <= 4; i++) {
            const start = row.querySelector(`[data-type="start"][data-period="${i}"]`);
            const end = row.querySelector(`[data-type="end"][data-period="${i}"]`);
            const clearBtn = row.querySelector(`.clear-btn[data-period="${i}"]`);
            if (start && !start._flatpickr) initFlatpickrStart(start, end, login, i, row);
            if (end && !end._flatpickr) initFlatpickrEnd(end, start, login, i, row);
            if (clearBtn) {
                const newClearBtn = clearBtn.cloneNode(true);
                clearBtn.parentNode.replaceChild(newClearBtn, clearBtn);
                newClearBtn.addEventListener('click', e => {
                    e.stopPropagation();
                    e.preventDefault();
                    clearPeriod(row, login, i);
                });
            }
            if (start) {
                start.addEventListener('change', () => {
                    changeMap[login] = true;
                    const btn = row.querySelector('.save-btn');
                    if (btn) btn.classList.add('save-changed');
                    updateExc(row, login, i);
                    updateTotalColumn(row);
                    updateEndMinDate(row, login, i);
                    determineRowStatus(row, login);
                });
            }
            if (end) {
                end.addEventListener('change', () => {
                    changeMap[login] = true;
                    const btn = row.querySelector('.save-btn');
                    if (btn) btn.classList.add('save-changed');
                    updateExc(row, login, i);
                    updateTotalColumn(row);
                    determineRowStatus(row, login);
                });
            }
        }
    });
}

document.addEventListener('click', function(e) {
    if (e.target && e.target.closest('.save-btn')) {
        const btn = e.target.closest('.save-btn');
        const row = btn.closest('tr');
        const login = btn.dataset.login;
        if (row && login) {
            e.preventDefault();
            e.stopPropagation();
            saveRow(row, login);
        }
    }
});

function updateEndMinDate(row, login, periodIndex) {
    const start = row.querySelector(`[data-type="start"][data-period="${periodIndex}"]`);
    const end = row.querySelector(`[data-type="end"][data-period="${periodIndex}"]`);
    if (start && end && end._flatpickr) {
        const startValue = start.value;
        end._flatpickr.set('minDate', startValue || "2026-01-01");
    }
}
function initFlatpickrStart(input, endInput, login, periodIndex, row) {
    if (!input || input._flatpickr) return;
    const config = {
        dateFormat: "Y/m/d",
        allowInput: true,
        locale: "es",
        minDate: "2026-01-01",
        maxDate: "2026-12-31",
        onChange: function(selectedDates, dateStr) {
            input.value = dateStr;
            changeMap[login] = true;
            const btn = row.querySelector('.save-btn');
            if (btn) btn.classList.add('save-changed');
            updateExc(row, login, periodIndex);
            updateTotalColumn(row);
            updateEndMinDate(row, login, periodIndex);
            determineRowStatus(row, login);
        }
    };
    if (input.value && input.value.trim() !== '') config.defaultDate = input.value;
    flatpickr(input, config);
}
function initFlatpickrEnd(input, startInput, login, periodIndex, row) {
    if (!input || input._flatpickr) return;
    const config = {
        dateFormat: "Y/m/d",
        allowInput: true,
        locale: "es",
        minDate: "2026-01-01",
        maxDate: "2026-12-31",
        onChange: function(selectedDates, dateStr) {
            input.value = dateStr;
            changeMap[login] = true;
            const btn = row.querySelector('.save-btn');
            if (btn) btn.classList.add('save-changed');
            updateExc(row, login, periodIndex);
            updateTotalColumn(row);
            determineRowStatus(row, login);
        }
    };
    if (startInput && startInput.value && startInput.value.trim() !== '') config.minDate = startInput.value;
    if (input.value && input.value.trim() !== '') config.defaultDate = input.value;
    flatpickr(input, config);
}

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
function detectTipoFromRow(row) {
    const tipoSelect = row.querySelector('.tipo-select');
    return tipoSelect ? Number(tipoSelect.value) : 1;
}
function detectTipo(login) {
    const emp = employees.find(e => e.login === login);
    return emp?.tipo === 2 || emp?.tipo === '2' ? 2 : 1;
}

function clearPeriod(row, login, periodIndex) {
    const start = row.querySelector(`[data-type="start"][data-period="${periodIndex}"]`);
    const end = row.querySelector(`[data-type="end"][data-period="${periodIndex}"]`);
    if (start) {
        start.value = '';
        if (start._flatpickr && typeof start._flatpickr.clear === 'function') start._flatpickr.clear();
        start.dispatchEvent(new Event('change'));
    }
    if (end) {
        end.value = '';
        if (end._flatpickr && typeof end._flatpickr.clear === 'function') {
            end._flatpickr.clear();
            if (typeof end._flatpickr.set === 'function') end._flatpickr.set('minDate', "2026-01-01");
        }
        end.dispatchEvent(new Event('change'));
    }
    changeMap[login] = true;
    const btn = row.querySelector('.save-btn');
    if (btn) btn.classList.add('save-changed');
    updateExc(row, login, periodIndex);
    updateTotalColumn(row);
    determineRowStatus(row, login);
}

async function saveRow(row, login) {
    const db = firebase.firestore();
    row.classList.add('row-saving');
    try {
        const updates = {};
        const tipoSelect = row.querySelector('.tipo-select');
        if (tipoSelect) {
            const tipoValue = tipoSelect.value;
            updates['tipo'] = tipoValue === '1' ? 1 : 2;
        }
        for (let i = 1; i <= 4; i++) {
            const startInput = row.querySelector(`[data-type="start"][data-period="${i}"]`);
            const endInput = row.querySelector(`[data-type="end"][data-period="${i}"]`);
            const start = startInput ? startInput.value.trim() : '';
            const end = endInput ? endInput.value.trim() : '';
            if ((start && !end) || (!start && end)) {
                alert(`Error en período ${i}: Debe completar ambas fechas o dejar ambas vacías.`);
                row.classList.remove('row-saving');
                return;
            }
            if (start && end) {
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
        await db.collection('empleados').doc(login).update(updates);
        const doc = await db.collection('empleados').doc(login).get();
        const fresh = { login: doc.id, ...doc.data() };
        const idx = employees.findIndex(e => e.login === login);
        if (idx >= 0) {
            employees[idx] = fresh;
            const filteredIdx = filteredEmployees.findIndex(e => e.login === login);
            if (filteredIdx >= 0) filteredEmployees[filteredIdx] = fresh;
        }
        if (dataTable) {
            const rowIndex = dataTable.row(row).index();
            dataTable.row(rowIndex).data(fresh).draw(false);
            setTimeout(() => attachRowEvents(), 50);
        }
        delete changeMap[login];
        const saveBtn = row.querySelector('.save-btn');
        if (saveBtn) saveBtn.classList.remove('save-changed');
        determineRowStatus(row, login);
    } catch (err) {
        alert(`Error al guardar: ${err.message}`);
    } finally { row.classList.remove('row-saving'); }
}

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
function calcOverCustom(s, e, tipo) {
    const d1 = parseDate(s);
    const d2 = parseDate(e);
    if (!d1 || !d2) return 0;
    const base = tipo === 1 ? 15 : 7;
    return Math.floor((d2 - d1) / 86400000) + 1 - base;
}
function daysInclusive(s, e) {
    const d1 = parseDate(s);
    const d2 = parseDate(e);
    return d1 && d2 ? Math.floor((d2 - d1) / 86400000) + 1 : 0;
}
function parseDate(str) {
    if (!str) return null;
    const cleanStr = str.replace(/\//g, '-');
    const [year, month, day] = cleanStr.split('-').map(Number);
    if (!year || !month || !day || isNaN(year) || isNaN(month) || isNaN(day)) return null;
    const date = new Date(year, month - 1, day);
    return isNaN(date.getTime()) ? null : date;
}
function escapeHtml(s) {
    if (s == null) return '';
    return String(s)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#039;');
}