// Festivos 2026 (Vigo)
const HOLIDAYS_2026 = [
    '2026-01-01',
    '2026-01-06',
    '2026-03-19',
    '2026-03-28',
    '2026-04-02',
    '2026-04-03',
    '2026-05-01',
    '2026-06-24',
    '2026-07-25',
    '2026-08-15',
    '2026-08-17',
    '2026-10-12',
    '2026-12-08',
    '2026-12-25'
];

const EXCLUSION_RANGES = [
    { id: 1, start: '2026-07-27', end: '2026-08-16' },
    { id: 1, start: '2026-12-21', end: '2026-12-31' },
    { id: 1, start: '2026-03-30', end: '2026-04-05' }
];

const DEBUG_MODE = false;

let currentUser = null;
let selectedRanges = [];
let allUsers = [];
let isSaving = false;

const DOM = {
    header: document.getElementById('header'),
    userNameSpan: document.getElementById('user-name'),
    instructionsText: document.getElementById('instructions-text'),
    calendarElement: document.getElementById('calendar'),
    saveBtn: document.getElementById('save-btn'),
    logoutBtn: document.getElementById('logout-btn'),
    floatingActions: document.getElementById('floating-actions'),
    successDialog: document.getElementById('success-dialog'),
    errorDialog: document.getElementById('error-dialog'),
    confirmDialog: document.getElementById('confirm-dialog'),
    confirmText: document.getElementById('confirm-text'),
    errorMessage: document.getElementById('error-message'),
    loadingOverlay: document.getElementById('loading-overlay'),
    colorLegend: document.getElementById('color-legend'),
    debugTable: null
};

const fmt = d => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
const rangeDays = (start, end) => {
    const s = new Date(start), e = new Date(end), days = [];
    while (s <= e) {
        days.push(fmt(s));
        s.setDate(s.getDate() + 1);
    }
    return days;
};

document.addEventListener('DOMContentLoaded', async () => {
    const userData = sessionStorage.getItem('vacaciones_user');
    if (!userData) {
        window.location.href = 'index.html';
        return;
    }

    currentUser = JSON.parse(userData);
    if (currentUser.admin === 1 || currentUser.admin === '1') {
        window.location.href = 'admin.html';
        return;
    }

    try {
        if (firebase.firestore && firebase.firestore()._persistenceEnabled) {
            await firebase.firestore().clearPersistence();
        }
    } catch (e) {
        console.log('No se pudo limpiar persistencia:', e);
    }

    firebase.initializeApp(firebaseConfig);
    const db = firebase.firestore();

    if (window.componentHandler) {
        window.componentHandler.upgradeAllRegistered();
        const spinner = DOM.loadingOverlay.querySelector('.mdl-spinner');
        if (spinner) window.componentHandler.upgradeElement(spinner);
    }

    DOM.saveBtn.addEventListener('click', showConfirmDialog);
    DOM.logoutBtn.addEventListener('click', handleLogout);

    document.querySelectorAll('.close, .confirm-cancel').forEach(btn => {
        btn.addEventListener('click', () => btn.closest('dialog').close());
    });
    document.querySelector('.confirm-save').addEventListener('click', () => saveVacations(db));
    
    const refreshBtn = document.getElementById('refresh-btn');
    if (refreshBtn) {
        refreshBtn.addEventListener('click', async () => {
            DOM.errorDialog.close();
            toggleLock(true);
            try {
                await loadUserData(db);
                generateCalendar();
            } catch (err) {
                console.error('Error al refrescar:', err);
            } finally {
                toggleLock(false);
            }
        });
    }
    
    DOM.errorDialog.addEventListener('close', () => {
        if (refreshBtn) refreshBtn.style.display = 'none';
    });

    document.addEventListener('keydown', e => {
        if (e.key === 'Escape') {
            DOM.successDialog.close();
            DOM.errorDialog.close();
            DOM.confirmDialog.close();
        }
    });

    const legend = document.getElementById('color-legend');
    if (legend && DEBUG_MODE) {
        const div = document.createElement('div');
        div.id = 'debug-table-container';
        legend.parentNode.insertBefore(div, legend.nextSibling);
        DOM.debugTable = div;
    } else if (legend && !DEBUG_MODE) {
        const existingDebug = document.getElementById('debug-table-container');
        if (existingDebug) {
            existingDebug.remove();
        }
        DOM.debugTable = null;
    }

    await loadUserData(db);
    renderLegend();
    renderDebugTable();
    generateCalendar();
    showMainScreen();
});

function showCollisionError(message, showRefresh = true) {
    DOM.errorMessage.textContent = message;
    const refreshBtn = document.getElementById('refresh-btn');
    if (refreshBtn) refreshBtn.style.display = showRefresh ? 'inline-block' : 'none';
    DOM.errorDialog.showModal();
}

async function loadUserData(db) {
    try {
        const doc = await db.collection('empleados').doc(currentUser.login).get({ source: 'server' });
        if (doc.exists) {
            currentUser = { ...currentUser, ...doc.data() };
            if (currentUser.tipo === undefined) currentUser.tipo = 1;
            const groupPromises = [ 
                db.collection('empleados').where('grupo', '==', currentUser.grupo).get({ source: 'server' })
            ];
            if (currentUser.subgrupo && currentUser.subgrupo.length > 0) {
                groupPromises.push(
                    db.collection('empleados').where('grupo', '==', currentUser.subgrupo).get({ source: 'server' })
                );
            }
            const groupSnaps = await Promise.all(groupPromises);
            const users = new Map();
            groupSnaps.forEach(snap => {
                snap.docs.forEach(doc => {
                    if (doc.id !== currentUser.login) {
                        users.set(doc.id, { login: doc.id, ...doc.data() });
                    }
                });
            });
            allUsers = Array.from(users.values());
            validateCurrentSelectionsAgainstDatabase();
        }
    } catch (err) {
        console.error('Error al cargar datos del usuario:', err);
        showCollisionError('Error al cargar datos del usuario. Por favor, refresca la página.', true);
    }
}

function showMainScreen() {
    DOM.userNameSpan.textContent = currentUser.nombre;
    DOM.instructionsText.innerHTML = currentUser.tipo === 1
        ? `Selecciona <strong>lunes o sábado</strong> para iniciar una quincena (15 días + festivos contiguos). Puedes seleccionar hasta 2 quincenas.`
        : `Selecciona <strong>lunes o sábado no festivo</strong> o alternativo festivo para marcar una semana (7 días desde el día pulsado y festivos/domingos pegados). Máximo 4.`;
}

function renderLegend() {
    const $legend = DOM.colorLegend || document.getElementById('color-legend');
    if (!$legend) return;
    const hasExclusions = currentUser.exclusiones === '1';
    $legend.innerHTML = `
        <div class="color-item">
            <span class="color-box" style="background: var(--selectable);"></span>
            <span>Seleccionable</span>
        </div>
        <div class="color-item">
            <span class="color-box" style="background: var(--selected);"></span>
            <span>Seleccionado no guardado</span>
        </div>
        <div class="color-item">
            <span class="color-box" style="background: var(--saved);"></span>
            <span>Vacaciones guardadas</span>
        </div>
        <div class="color-item">
            <span class="color-box" style="background: var(--holiday);"></span>
            <span>Festivo</span>
        </div>
        <div class="color-item">
            <span class="color-box" style="background: var(--occupied);"></span>
            <span>Vacaciones de otro</span>
        </div>
        ${hasExclusions ? `
        <div class="color-item">
            <span class="color-box" style="background: var(--excluded);"></span>
            <span>Excluido por empresa</span>
        </div>
        ` : ''}
        <div class="color-item">
            <span class="color-box" style="background: var(--blocked);"></span>
            <span>Solapamiento</span>
        </div>
    `;
}

function renderDebugTable() {
    if (!DOM.debugTable || !DEBUG_MODE) return;
    const hasExclusions = currentUser.exclusiones === '1';
    let html = `<h4 style="font-size:1.07em;color:#888;margin-top:18px;margin-bottom:7px;">Debug No puede coincidir con :</h4>
    <table style="width:100%;border-collapse:collapse;margin-top:0;font-size:0.95em">
        <thead style="background:#eee">
            <tr>
                <th style="border-bottom:1px solid #bbb;padding:4px 8px">Nombre</th>
                <th style="border-bottom:1px solid #bbb;padding:4px 8px">Empresa</th>
                <th style="border-bottom:1px solid #bbb;padding:4px 8px">Grupo</th>
            </tr>
        </thead>
        <tbody>
    `;
    allUsers.forEach(u => {
        html += `<tr>
            <td style="border-bottom:1px solid #ddd;padding:2px 8px">${u.nombre || u.login}</td>
            <td style="border-bottom:1px solid #ddd;padding:2px 8px">${u.empresa || ''}</td>
            <td style="border-bottom:1px solid #ddd;padding:2px 8px">${u.grupo || ''}</td>
        </tr>`;
    });
    html += "</tbody></table>";
    if (hasExclusions) {
        html += `<div style="margin-top:15px;padding:10px;background:#fff8e1;border-left:4px solid #ffb300;border-radius:4px;">
            <h5 style="margin:0 0 5px 0;color:#e65100;">Exclusiones activas:</h5>
            <ul style="margin:0;padding-left:20px;">
                ${EXCLUSION_RANGES.filter(r => r.id === 1).map(r => 
                    `<li><strong>${r.start} → ${r.end}</strong></li>`
                ).join('')}
            </ul>
        </div>`;
    }
    DOM.debugTable.innerHTML = html;
}

function handleLogout() {
    sessionStorage.removeItem('vacaciones_user');
    window.location.href = 'index.html';
}

function toggleLock(lock) {
    isSaving = lock;
    DOM.saveBtn.disabled = lock || !selectedRanges.length;
    DOM.logoutBtn.disabled = lock;
    document.querySelectorAll('input, button').forEach(el => el.disabled = lock);
    document.querySelectorAll('.day.selectable').forEach(el => {
        el.style.pointerEvents = lock ? 'none' : 'auto';
        el.style.opacity = lock ? '0.6' : '1';
    });
    DOM.loadingOverlay.style.display = lock ? 'flex' : 'none';
}

function generateCalendar() {
    DOM.calendarElement.innerHTML = '';
    for (let m = 0; m < 12; m++) {
        DOM.calendarElement.appendChild(createMonth(m));
    }
    DOM.saveBtn.disabled = selectedRanges.length === 0;
}

function createMonth(month) {
    const el = document.createElement('div');
    el.className = 'month';
    const monthName = new Date(2026, month).toLocaleDateString('es-ES', { month: 'long' });
    el.innerHTML = `<div class="month-header">${monthName.charAt(0).toUpperCase() + monthName.slice(1)}</div>
                    <div class="weekdays">${['Lu','Ma','Mi','Ju','Vi','Sá','Do'].map(d => `<div class="weekday">${d}</div>`).join('')}</div>
                    <div class="days"></div>`;
    const daysEl = el.querySelector('.days');
    const first = new Date(2026, month, 1);
    const offset = (first.getDay() || 7) - 1;
    for (let i = 0; i < offset; i++) {
        const d = new Date(first);
        d.setDate(d.getDate() - (offset - i));
        daysEl.appendChild(createDay(d, true));
    }
    const lastDay = new Date(2026, month + 1, 0).getDate();
    for (let d = 1; d <= lastDay; d++) {
        daysEl.appendChild(createDay(new Date(2026, month, d)));
    }
    return el;
}

function createDay(date, isOtherMonth = false) {
    const ds = fmt(date);
    const el = document.createElement('div');
    el.className = `day${isOtherMonth ? ' other-month' : ''}`;
    el.textContent = date.getDate();
    el.dataset.date = ds;

    const isHol = HOLIDAYS_2026.includes(ds);
    const isSun = date.getDay() === 0;
    const isSaved = isInAnyRange(ds, currentUser);
    const isExcluded = currentUser.exclusiones === '1' && isInExclusionRange(ds);
    const isOccupied = !isSaved && !isExcluded && isOccupiedForCurrent(ds);
    const isSelected = selectedRanges.some(([start, end]) => {
        const days = rangeDays(start, end);
        return days.includes(ds);
    });

    let selectableInfo = isSelectableDay(date);
    let isSelectable = selectableInfo.selectable;
    let isSpecialBlocked = selectableInfo.blocked;
    let isDesignatedDay = (!selectableInfo.realDate || fmt(selectableInfo.realDate) === ds);

    // CAMBIO CLAVE AQUÍ:
    let isBlocked = false;
    if (isSelectable) {
        if (typeof isSpecialBlocked !== "undefined") {
            isBlocked = isSpecialBlocked;
        } else {
            isBlocked = wouldOverlap(date);
        }
    }

    if (isOtherMonth) {
        el.style.visibility = 'hidden';
        el.style.pointerEvents = 'none';
        return el;
    }

    if (isSaved) {
        el.classList.add('saved');
    } else if (isSelected) {
        el.classList.add('selected');
        el.addEventListener('click', () => removeSelection(ds));
    } else if (isExcluded) {
        el.classList.add('excluded');
    } else if (isOccupied) {
        el.classList.add('occupied');
    }

    if (!isSaved && !isSelected && !isExcluded && !isOccupied) {
        if (isHol) el.classList.add('holiday');
        else if (isSun) el.classList.add('sunday');
    }
    if (!isSaved && !isSelected && !isExcluded && !isOccupied) {
        if (isBlocked) {
            el.classList.add('blocked');
        }
        else if (isSelectable && isDesignatedDay) {
            el.classList.add('selectable');
            el.addEventListener('click', () => toggleSelection(date));
        }
    }

    return el;
}

// --- NUEVA lógica para quincena: día alternativo para lunes/sábado festivo o su viernes o martes contiguos
function isQuincenaSelectableDay(date) {
    const w = date.getDay();
    const ds = fmt(date);

    // LUNES o SÁBADO normal no festivo
    if ((w === 1 || w === 6) && !HOLIDAYS_2026.includes(ds)) 
        return { selectable: true, realDate: date };

    // ¿Este día es alternativo anterior a un sábado festivo?
    if (w >= 1 && w < 6) {
        const next = new Date(date); next.setDate(next.getDate() + 1);
        const nDay = next.getDay();
        const nDs = fmt(next);
        if (nDay === 6 && HOLIDAYS_2026.includes(nDs)) {
            const blockResult = checkSelectableOverride(date);
            return { selectable: blockResult === true, realDate: date, blocked: blockResult === 'blocked' };
        }
    }

    // ¿Este día es alternativo posterior a un lunes festivo?
    if (w > 1 && w <= 6) {
        const prev = new Date(date); prev.setDate(prev.getDate() - 1);
        const pDay = prev.getDay();
        const pDs = fmt(prev);
        if (pDay === 1 && HOLIDAYS_2026.includes(pDs)) {
            const blockResult = checkSelectableOverride(date);
            return { selectable: blockResult === true, realDate: date, blocked: blockResult === 'blocked' };
        }
    }

    // La lógica original para el propio lunes o sábado festivo pulsado
    return isValidQuincenaStart(date);
}

function isValidQuincenaStart(date) {
    const w = date.getDay();
    const ds = fmt(date);

    if (w === 1 && !HOLIDAYS_2026.includes(ds)) return { selectable: true, realDate: date };
    if (w === 6 && !HOLIDAYS_2026.includes(ds)) return { selectable: true, realDate: date };

    if (w === 1 && HOLIDAYS_2026.includes(ds)) {
        let next = new Date(date);
        for (let i = 1; i <= 6; i++) {
            next.setDate(next.getDate() + 1);
            const nDay = next.getDay();
            const nDs = fmt(next);
            if (!(HOLIDAYS_2026.includes(nDs) || nDay === 0)) {
                const blockResult = checkSelectableOverride(next);
                return { selectable: blockResult === true, realDate: next, blocked: blockResult === 'blocked' };
            }
        }
        return { selectable: false };
    }
    if (w === 6 && HOLIDAYS_2026.includes(ds)) {
        let prev = new Date(date);
        for (let i = 1; i <= 6; i++) {
            prev.setDate(prev.getDate() - 1);
            const pDay = prev.getDay();
            const pDs = fmt(prev);
            if (!(HOLIDAYS_2026.includes(pDs) || pDay === 0)) {
                const blockResult = checkSelectableOverride(prev);
                return { selectable: blockResult === true, realDate: prev, blocked: blockResult === 'blocked' };
            }
        }
        return { selectable: false };
    }
    if (isAltForHoliday(date)) {
        return { selectable: true, realDate: date };
    }
    return { selectable: false };
}

function isSelectableDay(date) {
    if (!currentUser) return { selectable: false };
    const guardadas = getCurrentSavedRanges();
    if (currentUser.tipo === 1) {
        if (guardadas.length >= 2) return { selectable: false };
        return isQuincenaSelectableDay(date);
    } else {
        if (guardadas.length >= 4) return { selectable: false };
        return { selectable: isValidWeekStart(date), realDate: date };
    }
}

function checkSelectableOverride(date) {
    const ds = fmt(date);
    const r = [ds, fmt(calcRangeEnd(date, 14))];
    if (
        isInAnyRange(ds, currentUser) ||
        wouldOverlapRange(r)
    ) {
        return "blocked";
    }
    return true;
}

function isInExclusionRange(ds) {
    if (currentUser.exclusiones !== '1') return false;
    const date = new Date(ds);
    for (const range of EXCLUSION_RANGES) {
        const startDate = new Date(range.start);
        const endDate = new Date(range.end);
        if (date >= startDate && date <= endDate) {
            return true;
        }
    }
    return false;
}

function isOccupiedForCurrent(ds) {
    const occupiedByOthers = allUsers.some(u => isInAnyRange(ds, u));
    const isExcluded = currentUser.exclusiones === '1' && isInExclusionRange(ds);
    return occupiedByOthers || isExcluded;
}

function toggleSelection(date) {
    const prevRanges = getCurrentSavedRanges();
    if (currentUser.tipo === 1) {
        if (prevRanges.length >= 2) return;
        const r = createQuincenaRange(date);
        if (!r) return;
        const idx = selectedRanges.findIndex(x => x[0] === r[0] && x[1] === r[1]);
        if (idx >= 0) {
            selectedRanges.splice(idx, 1);
        } else {
            const availableSlots = 2 - (prevRanges.length + selectedRanges.length);
            if (availableSlots <= 0) return;
            if (!wouldOverlapRange(r)) {
                selectedRanges.push(r);
            }
        }
    } else if (currentUser.tipo === 2) {
        if (prevRanges.length >= 4) return;
        const r = createWeekRange(date);
        if (!r) return;
        const idx = selectedRanges.findIndex(x => x[0] === r[0] && x[1] === r[1]);
        if (idx >= 0) {
            selectedRanges.splice(idx, 1);
        } else {
            const availableSlots = 4 - (prevRanges.length + selectedRanges.length);
            if (availableSlots <= 0) return;
            if (!wouldOverlapRange(r)) {
                selectedRanges.push(r);
            }
        }
    }
    generateCalendar();
}

function removeSelection(ds) {
    const idx = selectedRanges.findIndex(([start, end]) => {
        const days = rangeDays(start, end);
        return days.includes(ds);
    });
    if (idx !== -1) {
        selectedRanges.splice(idx, 1);
        generateCalendar();
    }
}

function getCurrentSavedRanges() {
    const ranges = [];
    for (let i = 1; i <= 4; i++) {
        const start = currentUser[`per${i}start`];
        const end = currentUser[`per${i}end`];
        if (start && end) ranges.push([start, end]);
    }
    return ranges;
}

const createQuincenaRange = (start) => {
    const valid = isValidQuincenaStart(start);
    if (!valid.selectable || valid.blocked) return null;
    const end = calcRangeEnd(valid.realDate, 14);
    return [fmt(valid.realDate), fmt(end)];
};

const createWeekRange = (date) => {
    if (!isValidWeekStart(date)) return null;
    let start = new Date(date);
    let end = new Date(start);
    end.setDate(end.getDate() + 6);
    let next = new Date(end);
    next.setDate(next.getDate() + 1);
    while (HOLIDAYS_2026.includes(fmt(next)) || next.getDay() === 0) {
        end = next;
        next = new Date(end);
        next.setDate(next.getDate() + 1);
    }
    return [fmt(start), fmt(end)];
};

function calcRangeEnd(start, baseDays) {
    let end = new Date(start);
    end.setDate(end.getDate() + baseDays);
    let next = new Date(end);
    next.setDate(next.getDate() + 1);
    while (HOLIDAYS_2026.includes(fmt(next)) || next.getDay() === 0) {
        end = next;
        next = new Date(end);
        next.setDate(next.getDate() + 1);
    }
    return end;
}

function isValidWeekStart(date) {
    const w = date.getDay();
    if ((w === 1 || w === 6) && !HOLIDAYS_2026.includes(fmt(date))) return true;
    return isAltForHoliday(date);
}

function isAltForHoliday(date) {
    const w = date.getDay();
    if (w < 2 || w > 5) return false;
    const mon = new Date(date);
    mon.setDate(date.getDate() - (w - 1));
    if (!HOLIDAYS_2026.includes(fmt(mon))) return false;
    for (let d = 1; d < w; d++) {
        const dd = new Date(date);
        dd.setDate(dd.getDate() - (w - d));
        if (!HOLIDAYS_2026.includes(fmt(dd))) return false;
    }
    return true;
}

function isInAnyRange(ds, user) {
    for (let i = 1; i <= 4; i++) {
        const s = user[`per${i}start`];
        const e = user[`per${i}end`];
        if (s && e && ds >= s && ds <= e) return true;
    }
    return false;
}

function wouldOverlap(date) {
    if (currentUser.tipo === 1) {
        const r = createQuincenaRange(date);
        if (!r) return true;
        return wouldOverlapRange(r);
    } else {
        const r = createWeekRange(date);
        if (!r) return true;
        return wouldOverlapRange(r);
    }
}

function wouldOverlapRange([s1, e1], forceRealTimeCheck = false) {
    const start1 = new Date(s1);
    const end1 = new Date(e1);
    const users = forceRealTimeCheck ? allUsers : [currentUser, ...allUsers];
    for (const u of users) {
        for (let i = 1; i <= 4; i++) {
            const s2 = u[`per${i}start`];
            const e2 = u[`per${i}end`];
            if (!s2 || !e2) continue;
            const start2 = new Date(s2);
            const end2 = new Date(e2);
            if (!(end1 < start2 || start1 > end2)) return true;
        }
    }
    for (const [s2, e2] of selectedRanges) {
        if (s1 === s2 && e1 === e2) continue;
        const start2 = new Date(s2);
        const end2 = new Date(e2);
        if (!(end1 < start2 || start1 > end2)) return true;
    }
    if (currentUser.exclusiones === '1') {
        const rangeDaysList = rangeDays(s1, e1);
        for (const day of rangeDaysList) {
            if (isInExclusionRange(day)) {
                return true;
            }
        }
    }
    return false;
}

function validateCurrentSelectionsAgainstDatabase() {
    if (selectedRanges.length === 0) return;
    const invalidRanges = [];
    selectedRanges.forEach((range, index) => {
        if (wouldOverlapRange(range, true)) {
            invalidRanges.push(index);
        }
    });
    if (invalidRanges.length > 0) {
        invalidRanges.sort((a, b) => b - a);
        invalidRanges.forEach(index => {
            selectedRanges.splice(index, 1);
        });
        if (selectedRanges.length < invalidRanges.length) {
            setTimeout(() => {
                showCollisionError('¡Atención! Otro compañero acaba de solicitar días de vacaciones que se solapaban con tus selecciones.\n\nTus selecciones han sido actualizadas automáticamente.', true);
            }, 500);
        }
        generateCalendar();
    }
}

function showConfirmDialog() {
    const prevRanges = getCurrentSavedRanges();
    const allRanges = [...prevRanges, ...selectedRanges].slice(0, currentUser.tipo === 1 ? 2 : 4);
    const periods = allRanges.map(([s, e]) => `${s} → ${e}`).join('<br>');
    DOM.confirmText.innerHTML = `¿Confirmar?<br><small>${periods}</small>`;
    DOM.confirmDialog.showModal();
}

async function saveVacations(db) {
    DOM.confirmDialog.close();
    if (!selectedRanges.length || isSaving) return;
    toggleLock(true);
    try {
        const latestUserDoc = await db.collection('empleados').doc(currentUser.login).get({ source: 'server' });
        const latestUserData = latestUserDoc.exists ? latestUserDoc.data() : {};
        const latestGroupPromises = [
            db.collection('empleados').where('grupo', '==', currentUser.grupo).get({ source: 'server' })
        ];
        if (currentUser.subgrupo && currentUser.subgrupo.length > 0) {
            latestGroupPromises.push(
                db.collection('empleados').where('grupo', '==', currentUser.subgrupo).get({ source: 'server' })
            );
        }
        const latestGroupSnaps = await Promise.all(latestGroupPromises);
        const latestUsers = new Map();
        latestGroupSnaps.forEach(snap => {
            snap.docs.forEach(doc => {
                if (doc.id !== currentUser.login) {
                    latestUsers.set(doc.id, { login: doc.id, ...doc.data() });
                }
            });
        });
        for (const [start, end] of selectedRanges) {
            for (let i = 1; i <= 4; i++) {
                const s2 = latestUserData[`per${i}start`];
                const e2 = latestUserData[`per${i}end`];
                if (!s2 || !e2) continue;
                const start2 = new Date(s2);
                const end2 = new Date(e2);
                if (!(new Date(end) < start2 || new Date(start) > end2)) {
                    throw new Error('Colisión detectada con tus propias vacaciones recién actualizadas. Por favor, recarga la página.');
                }
            }
            for (const user of latestUsers.values()) {
                for (let i = 1; i <= 4; i++) {
                    const s2 = user[`per${i}start`];
                    const e2 = user[`per${i}end`];
                    if (!s2 || !e2) continue;
                    const start2 = new Date(s2);
                    const end2 = new Date(e2);
                    if (!(new Date(end) < start2 || new Date(start) > end2)) {
                        throw new Error('¡Atención! Parece que otro compañero acaba de solicitar días de vacaciones que se solapan con los que intentas reservar.\n\nPor favor, refresca el calendario para ver los cambios más recientes.');
                    }
                }
            }
        }
        const prevRanges = getCurrentSavedRanges();
        let allRanges = [...prevRanges];
        selectedRanges.forEach(sr => {
            const existe = allRanges.some(([s, e]) => sr[0] === s && sr[1] === e);
            if (!existe) allRanges.push(sr);
        });
        allRanges = allRanges.slice(0, currentUser.tipo === 1 ? 2 : 4);
        const updates = {};
        allRanges.forEach(([start, end], i) => {
            updates[`per${i + 1}start`] = start;
            updates[`per${i + 1}end`] = end;
        });
        for (let i = allRanges.length + 1; i <= 4; i++) {
            updates[`per${i}start`] = firebase.firestore.FieldValue.delete();
            updates[`per${i}end`] = firebase.firestore.FieldValue.delete();
        }
        await db.collection('empleados').doc(currentUser.login).update(updates);
        const updatedDoc = await db.collection('empleados').doc(currentUser.login).get({ source: 'server' });
        currentUser = { login: currentUser.login, ...updatedDoc.data() };
        if (currentUser.tipo === undefined) currentUser.tipo = 1;
        selectedRanges = [];
        DOM.successDialog.showModal();
        generateCalendar();
        renderDebugTable();
    } catch (err) {
        const showRefresh = err.message.includes('otro compañero') || err.message.includes('refresca') || err.message.includes('recarga');
        showCollisionError(err.message, showRefresh);
        try {
            await loadUserData(db);
            generateCalendar();
        } catch (reloadErr) {
            console.error('Error al recargar datos:', reloadErr);
        }
    } finally {
        toggleLock(false);
    }
}

function loadPendingSelections() {
    selectedRanges = [];
}