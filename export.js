async function exportToExcel(employeesData, exportButton) {
    try {
        // Mostrar indicador de carga
        exportButton.disabled = true;
        exportButton.innerHTML = '<i class="material-icons" style="margin-right: 8px;">hourglass_empty</i> Generando Excel...';
        
        if (employeesData.length === 0) {
            alert('No hay datos para exportar');
            exportButton.disabled = false;
            exportButton.innerHTML = '<i class="material-icons" style="margin-right: 8px;">download</i> Exportar Excel';
            return;
        }
        
        // Crear un nuevo libro de Excel
        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('Vacaciones 2026');
        
        // Configurar el ancho de las columnas (incluyendo exclusiones)
        const columnWidths = [
            { header: 'Nombre', key: 'nombre', width: 75 }, // 3 veces más ancha
            { header: 'DNI', key: 'dni', width: 15 },
            { header: 'Empresa', key: 'empresa', width: 20 },
            { header: 'Grupo', key: 'grupo', width: 15 },
            { header: 'Subgrupo', key: 'subgrupo', width: 15 },
            { header: 'Exclusiones', key: 'exclusiones', width: 25 }, // COLUMNA AÑADIDA
            { header: 'Tipo Vacaciones', key: 'tipo', width: 15 },
            { header: 'Rango1Ini', key: 'r1ini', width: 12 },
            { header: 'Rango1Fin', key: 'r1fin', width: 12 },
            { header: 'Días1', key: 'dias1', width: 8 },
            { header: 'Rango2Ini', key: 'r2ini', width: 12 },
            { header: 'Rango2Fin', key: 'r2fin', width: 12 },
            { header: 'Días2', key: 'dias2', width: 8 },
            { header: 'Rango3Ini', key: 'r3ini', width: 12 },
            { header: 'Rango3Fin', key: 'r3fin', width: 12 },
            { header: 'Días3', key: 'dias3', width: 8 },
            { header: 'Rango4Ini', key: 'r4ini', width: 12 },
            { header: 'Rango4Fin', key: 'r4fin', width: 12 },
            { header: 'Días4', key: 'dias4', width: 8 },
            { header: 'Día Suelto 1', key: 'dia1', width: 12 },
            { header: 'Día Suelto 2', key: 'dia2', width: 12 }
        ];
        
        // Agregar las columnas para cada día del año (2026 no es bisiesto)
        const meses = ['ENERO', 'FEBRERO', 'MARZO', 'ABRIL', 'MAYO', 'JUNIO', 
                      'JULIO', 'AGOSTO', 'SEPTIEMBRE', 'OCTUBRE', 'NOVIEMBRE', 'DICIEMBRE'];
        const diasPorMes = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
        
        // Crear dos filas de cabecera
        const headerRow1 = [
            'Nombre', 'DNI', 'Empresa', 'Grupo', 'Subgrupo', 'Exclusiones', 'Tipo Vacaciones', // EXCLUSIONES AÑADIDA
            'Rango1Ini', 'Rango1Fin', 'Días1',
            'Rango2Ini', 'Rango2Fin', 'Días2',
            'Rango3Ini', 'Rango3Fin', 'Días3',
            'Rango4Ini', 'Rango4Fin', 'Días4',
            'Día Suelto 1', 'Día Suelto 2'
        ];
        
        const headerRow2 = Array(21).fill(''); // Cambiado de 20 a 21 por la nueva columna
        
        // Agregar los meses a la primera fila y los días a la segunda
        let colIndex = 21; // Cambiado de 20 a 21 por la nueva columna
        const monthStartColumns = {};
        const allDates = []; // Almacenar todas las fechas del año
        
        // Generar fechas para cada día del año (2026-01-01 a 2026-12-31)
        let currentDate = new Date(2026, 0, 1); // 1 de enero de 2026
        
        meses.forEach((mes, mesIndex) => {
            monthStartColumns[mes] = colIndex + 1;
            headerRow1.push(mes);
            
            for (let d = 1; d <= diasPorMes[mesIndex]; d++) {
                // Crear fecha sin horas (00:00:00)
                const date = new Date(Date.UTC(2026, mesIndex, d));
                allDates.push(date);
                headerRow2.push(date.getDate()); // Solo mostrar el día
                colIndex++;
            }
        });
        
        // Agregar las dos filas de cabecera
        const row1 = worksheet.addRow(headerRow1);
        const row2 = worksheet.addRow(headerRow2);
        
        // Aplicar estilos a las cabeceras
        row1.font = { bold: true, size: 12 };
        row2.font = { bold: true, size: 10 };
        row1.alignment = { horizontal: 'center', vertical: 'center' };
        row2.alignment = { horizontal: 'center', vertical: 'center' };
        
        // Estilo especial para los nombres de los meses
        meses.forEach((mes, mesIndex) => {
            const startCol = monthStartColumns[mes];
            if (startCol) {
                for (let d = 0; d < diasPorMes[mesIndex]; d++) {
                    const cell = row1.getCell(startCol + d);
                    cell.fill = {
                        type: 'pattern',
                        pattern: 'solid',
                        fgColor: { argb: 'FFE0E0E0' }
                    };
                }
            }
        });
        
        // Fusionar celdas para los meses (colspan)
        let currentCol = 22; // Cambiado de 21 a 22 por la nueva columna
        meses.forEach((mes, mesIndex) => {
            const daysInMonth = diasPorMes[mesIndex];
            if (daysInMonth > 1) {
                worksheet.mergeCells(row1.number, currentCol, row1.number, currentCol + daysInMonth - 1);
            }
            currentCol += daysInMonth;
        });
        
        // Fusionar celdas verticalmente para las primeras 21 columnas (rowspan)
        for (let i = 1; i <= 21; i++) { // Cambiado de 20 a 21
            worksheet.mergeCells(row1.number, i, row2.number, i);
        }
        
        // Procesar cada empleado
        employeesData.forEach((emp, empIndex) => {
            const rowNumber = empIndex + 3; // +3 porque las filas 1 y 2 son cabeceras
            
            // Formatear exclusiones (convertir array a string si es necesario)
            let exclusionesFormatted = '';
            if (emp.exclusiones) {
                if (Array.isArray(emp.exclusiones)) {
                    exclusionesFormatted = emp.exclusiones.join(', ');
                } else {
                    exclusionesFormatted = String(emp.exclusiones);
                }
            }
            
            // Preparar datos de la fila
            const rowData = {
                nombre: emp.nombre || '',
                dni: emp.dni || '',
                empresa: emp.empresa || '',
                grupo: emp.grupo || '',
                subgrupo: emp.subgrupo || '', // Asegurar que se incluye
                exclusiones: exclusionesFormatted, // EXCLUSIONES AÑADIDAS
                tipo: detectTipoExport(emp) === 1 ? 'Quincena' : 'Semana', // CORREGIDO: Solo "Quincena" o "Semana"
                // Fechas como objetos Date sin horas
                r1ini: emp.per1start ? createDateWithoutTime(emp.per1start) : null,
                r1fin: emp.per1end ? createDateWithoutTime(emp.per1end) : null,
                dias1: emp.per1start && emp.per1end ? calculateDaysDifference(emp.per1start, emp.per1end) : '',
                r2ini: emp.per2start ? createDateWithoutTime(emp.per2start) : null,
                r2fin: emp.per2end ? createDateWithoutTime(emp.per2end) : null,
                dias2: emp.per2start && emp.per2end ? calculateDaysDifference(emp.per2start, emp.per2end) : '',
                r3ini: emp.per3start ? createDateWithoutTime(emp.per3start) : null,
                r3fin: emp.per3end ? createDateWithoutTime(emp.per3end) : null,
                dias3: emp.per3start && emp.per3end ? calculateDaysDifference(emp.per3start, emp.per3end) : '',
                r4ini: emp.per4start ? createDateWithoutTime(emp.per4start) : null,
                r4fin: emp.per4end ? createDateWithoutTime(emp.per4end) : null,
                dias4: emp.per4start && emp.per4end ? calculateDaysDifference(emp.per4start, emp.per4end) : '',
                dia1: null,
                dia2: null
            };
            
            // Crear la fila con los primeros 21 valores
            const rowArray = Object.values(rowData);
            const newRow = worksheet.addRow(rowArray);
            
            // Aplicar formato de fecha a las columnas de fechas
            const dateColumns = [8, 9, 11, 12, 14, 15, 17, 18, 20, 21]; // Ajustados por la nueva columna
            dateColumns.forEach(colIndex => {
                const cell = newRow.getCell(colIndex);
                if (cell.value instanceof Date) {
                    cell.numFmt = 'dd/mm/yyyy';
                    cell.alignment = { horizontal: 'center', vertical: 'center' };
                } else if (cell.value === null || cell.value === '') {
                    cell.value = '';
                }
            });
            
            // Aplicar alineación centrada a las columnas de días
            const daysColumns = [10, 13, 16, 19]; // Ajustados por la nueva columna
            daysColumns.forEach(colIndex => {
                const cell = newRow.getCell(colIndex);
                cell.alignment = { horizontal: 'center', vertical: 'center' };
            });
            
            // Aplicar alineación a la columna de exclusiones
            const exclusionesCell = newRow.getCell(6);
            exclusionesCell.alignment = { horizontal: 'left', vertical: 'center', wrapText: true };
            
            // Agregar fórmulas condicionales para cada día del año
            // Las fórmulas verificarán si la fecha está dentro de algún rango
            allDates.forEach((date, dateIndex) => {
                const colNumber = 22 + dateIndex; // Columna actual (1-based) - Cambiado de 21 a 22
                const colLetter = getExcelColumnLetter(colNumber);
                const rowNum = newRow.number;
                
                // Crear fórmula que verifica si la fecha está en algún rango
                // Usar referencias absolutas para las columnas de fechas
                let formula = '';
                
                // Array para almacenar las condiciones de cada rango
                const conditions = [];
                
                // Para cada rango (1-4), crear una condición
                for (let i = 1; i <= 4; i++) {
                    const startCol = 7 + (i-1)*3 + 1; // Columna de inicio del rango (ajustado por nueva columna)
                    const endCol = startCol + 1; // Columna de fin del rango
                    const startLetter = getExcelColumnLetter(startCol);
                    const endLetter = getExcelColumnLetter(endCol);
                    
                    // Condición: fecha >= inicio AND fecha <= fin
                    // Usar referencias absolutas para las columnas de rango
                    conditions.push(
                        `AND(${colLetter}${rowNum}>=${startLetter}${rowNum},${colLetter}${rowNum}<=${endLetter}${rowNum})`
                    );
                }
                
                // Combinar todas las condiciones con OR
                if (conditions.length > 0) {
                    formula = `=IF(OR(${conditions.join(',')}),1,0)`;
                }
                
                // Asignar la fórmula a la celda
                const cell = newRow.getCell(colNumber);
                cell.value = { formula: formula };
                cell.alignment = { horizontal: 'center', vertical: 'center' };
            });
        });
        
        // Crear reglas de formato condicional para las celdas de días
        // Pintar de rojo las celdas donde la fórmula devuelva 1
        const startCol = 22; // Primera columna de días (cambiado de 21 a 22)
        const endCol = startCol + allDates.length - 1; // Última columna de días
        const startRow = 3; // Primera fila de datos
        const endRow = employeesData.length + 2; // Última fila de datos
        
        const startColLetter = getExcelColumnLetter(startCol);
        const endColLetter = getExcelColumnLetter(endCol);
        
        // Agregar formato condicional CORREGIDO
        worksheet.addConditionalFormatting({
            ref: `${startColLetter}${startRow}:${endColLetter}${endRow}`,
            rules: [
                {
                    type: 'expression',
                    formulae: [`${startColLetter}${startRow}=1`],
                    style: {
                        fill: {
                            type: 'pattern',
                            pattern: 'solid',
                            fgColor: { argb: 'FFFF0000' } // Rojo
                        },
                        font: {
                            color: { argb: 'FFFFFFFF' }, // Blanco
                            bold: true
                        }
                    }
                }
            ]
        });
        
        // Aplicar bordes a todas las celdas
        worksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
            row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
                cell.border = {
                    top: { style: 'thin' },
                    left: { style: 'thin' },
                    bottom: { style: 'thin' },
                    right: { style: 'thin' }
                };
            });
        });
        
        // Congelar paneles (fijar las primeras 2 filas y las primeras 21 columnas)
        worksheet.views = [
            { state: 'frozen', xSplit: 21, ySplit: 2, activeCell: 'A1' } // Cambiado de 20 a 21
        ];
        
        // Generar el archivo Excel
        const buffer = await workbook.xlsx.writeBuffer();
        const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        
        // Descargar el archivo
        const fecha = new Date().toISOString().slice(0, 10).replace(/-/g, '');
        saveAs(blob, `vacaciones_2026_${fecha}.xlsx`);
        
    } catch (err) {
        console.error('Error al exportar a Excel:', err);
        alert(`Error al exportar a Excel: ${err.message}`);
    } finally {
        // Restaurar el botón
        exportButton.disabled = false;
        exportButton.innerHTML = '<i class="material-icons" style="margin-right: 8px;">download</i> Exportar Excel';
    }
}

// ===== FUNCIONES HELPER PARA EXPORTACIÓN =====

// Crear fecha sin horas (00:00:00)
function createDateWithoutTime(dateString) {
    if (!dateString) return null;
    
    // Convertir formato Y/m/d o Y-m-d a Date
    const cleanStr = dateString.replace(/\//g, '-');
    const [year, month, day] = cleanStr.split('-').map(Number);
    
    if (!year || !month || !day || isNaN(year) || isNaN(month) || isNaN(day)) {
        return null;
    }
    
    // Crear fecha UTC sin horas
    return new Date(Date.UTC(year, month - 1, day));
}

// Calcular diferencia de días entre dos fechas (inclusive)
function calculateDaysDifference(startStr, endStr) {
    const startDate = createDateWithoutTime(startStr);
    const endDate = createDateWithoutTime(endStr);
    
    if (!startDate || !endDate) return 0;
    
    // Calcular diferencia en días
    const diffTime = Math.abs(endDate - startDate);
    const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
    
    return diffDays + 1; // +1 porque es inclusive
}

// Obtener letra de columna Excel (1 -> A, 27 -> AA, etc.)
function getExcelColumnLetter(columnNumber) {
    let columnLetter = '';
    while (columnNumber > 0) {
        const remainder = (columnNumber - 1) % 26;
        columnLetter = String.fromCharCode(65 + remainder) + columnLetter;
        columnNumber = Math.floor((columnNumber - 1) / 26);
    }
    return columnLetter;
}

// Detectar tipo de empleado (1: quincena, 2: semana) - SOLO "QUINCENA" O "SEMANA"
function detectTipoExport(emp) {
    return emp?.tipo === 2 || emp?.tipo === '2' ? 2 : 1;
}