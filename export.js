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
        
        // ===== CONFIGURACIÓN DE COLUMNAS =====
        // Definir anchos de columnas según la estructura solicitada
        worksheet.columns = [
            { key: 'nombre', width: 25 },    // A: Nombre
            { key: 'dni', width: 12 },       // B: DNI
            { key: 'empresa', width: 18 },   // C: Empresa
            { key: 'exclusiones', width: 6 }, // D: Exclusiones (numérico pequeño)
            { key: 'tipo', width: 10 },      // E: Tipo
            { key: 'r1ini', width: 12 },     // F: Rango1Ini
            { key: 'r1fin', width: 12 },     // G: Rango1Fin
            { key: 'dias1', width: 6 },      // H: Días1 (fórmula)
            { key: 'r2ini', width: 12 },     // I: Rango2Ini
            { key: 'r2fin', width: 12 },     // J: Rango2Fin
            { key: 'dias2', width: 6 },      // K: Días2 (fórmula)
            { key: 'r3ini', width: 12 },     // L: Rango3Ini
            { key: 'r3fin', width: 12 },     // M: Rango3Fin
            { key: 'dias3', width: 6 },      // N: Días3 (fórmula)
            { key: 'r4ini', width: 12 },     // O: Rango4Ini
            { key: 'r4fin', width: 12 },     // P: Rango4Fin
            { key: 'dias4', width: 6 },      // Q: Días4 (fórmula)
            { key: 'dia1', width: 12 },      // R: DiaSuelto1
            { key: 'dia2', width: 12 },      // S: DiaSuelto2
            { key: 'total', width: 10 }      // T: Total Días (fórmula)
        ];
        
        // ===== PRIMERA FILA (TÍTULOS DE SECCIONES) =====
        const row1 = worksheet.addRow([]);
        
        // Fusionar celdas y agregar títulos para cada sección
        worksheet.mergeCells('F1:H1');
        worksheet.getCell('F1').value = 'Vacaciones 1';
        
        worksheet.mergeCells('I1:K1');
        worksheet.getCell('I1').value = 'Vacaciones 2';
        
        worksheet.mergeCells('L1:N1');
        worksheet.getCell('L1').value = 'Vacaciones 3';
        
        worksheet.mergeCells('O1:Q1');
        worksheet.getCell('O1').value = 'Vacaciones 4';
        
        worksheet.mergeCells('R1:S1');
        worksheet.getCell('R1').value = 'Días sueltos';
        
        worksheet.getCell('T1').value = 'Total Días';
        
        // Estilo para la primera fila (títulos de secciones)
        ['F1', 'I1', 'L1', 'O1', 'R1', 'T1'].forEach(cell => {
            const cellObj = worksheet.getCell(cell);
            cellObj.font = { bold: true, size: 12 };
            cellObj.alignment = { horizontal: 'center', vertical: 'center' };
            cellObj.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFE0E0E0' }
            };
            cellObj.border = {
                top: { style: 'thin' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' }
            };
        });
        
        // ===== SEGUNDA FILA (ENCABEZADOS DE COLUMNAS) =====
        const headers = [
            'Nombre', 'DNI', 'Empresa', 'Exc.', 'Tipo',
            'Rango1Ini', 'Rango1Fin', 'Días1',
            'Rango2Ini', 'Rango2Fin', 'Días2',
            'Rango3Ini', 'Rango3Fin', 'Días3',
            'Rango4Ini', 'Rango4Fin', 'Días4',
            'DiaSuelto1', 'DiaSuelto2',
            'Total Días'
        ];
        
        const row2 = worksheet.addRow(headers);
        
        // Estilo para la segunda fila
        row2.eachCell((cell, colNumber) => {
            cell.font = { bold: true, size: 11 };
            cell.alignment = { horizontal: 'center', vertical: 'center' };
            cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFF0F0F0' }
            };
            cell.border = {
                top: { style: 'thin' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' }
            };
        });
        
        // ===== DATOS DE EMPLEADOS (A PARTIR DE FILA 3) =====
        employeesData.forEach((emp, index) => {
            const rowNumber = index + 3; // Comenzar en fila 3
            
            // Formatear exclusiones como número (contar la cantidad)
            let exclusionesNumerico = 0;
            if (emp.exclusiones) {
                if (Array.isArray(emp.exclusiones)) {
                    // Si es array, contar elementos no vacíos
                    exclusionesNumerico = emp.exclusiones.filter(item => 
                        item !== null && item !== undefined && String(item).trim() !== ''
                    ).length;
                } else if (typeof emp.exclusiones === 'string') {
                    // Si es string, contar elementos separados por comas
                    const items = emp.exclusiones.split(',').map(item => item.trim());
                    exclusionesNumerico = items.filter(item => item !== '').length;
                } else if (typeof emp.exclusiones === 'number') {
                    // Si ya es número, usarlo directamente
                    exclusionesNumerico = emp.exclusiones;
                }
            }
            
            // Determinar tipo
            const tipo = emp?.tipo === 2 || emp?.tipo === '2' ? 'Semana' : 'Quincena';
            
            // Crear la fila con los datos básicos
            const rowData = [
                emp.nombre || '',                    // A: Nombre
                emp.dni || '',                      // B: DNI
                emp.empresa || '',                  // C: Empresa
                exclusionesNumerico,                // D: Exclusiones (numérico)
                tipo,                               // E: Tipo
                createDateWithoutTime(emp.per1start), // F: Rango1Ini
                createDateWithoutTime(emp.per1end),   // G: Rango1Fin
                '',                                  // H: Días1 (fórmula)
                createDateWithoutTime(emp.per2start), // I: Rango2Ini
                createDateWithoutTime(emp.per2end),   // J: Rango2Fin
                '',                                  // K: Días2 (fórmula)
                createDateWithoutTime(emp.per3start), // L: Rango3Ini
                createDateWithoutTime(emp.per3end),   // M: Rango3Fin
                '',                                  // N: Días3 (fórmula)
                createDateWithoutTime(emp.per4start), // O: Rango4Ini
                createDateWithoutTime(emp.per4end),   // P: Rango4Fin
                '',                                  // Q: Días4 (fórmula)
                null,                               // R: DiaSuelto1 (vacío por ahora)
                null,                               // S: DiaSuelto2 (vacío por ahora)
                ''                                  // T: Total Días (fórmula)
            ];
            
            const row = worksheet.addRow(rowData);
            
            // Aplicar formato numérico a la columna de exclusiones (columna D)
            const exclusionCell = row.getCell(4); // Columna D (4 = 1-indexed)
            exclusionCell.numFmt = '0';
            exclusionCell.alignment = { horizontal: 'center', vertical: 'center' };
            
            // Aplicar bordes a toda la fila
            row.eachCell((cell) => {
                cell.border = {
                    top: { style: 'thin' },
                    left: { style: 'thin' },
                    bottom: { style: 'thin' },
                    right: { style: 'thin' }
                };
            });
            
            // ===== FÓRMULAS PARA DÍAS DE CADA RANGO =====
            const currentRow = rowNumber;
            
            // Fórmula para Días1 (columna H)
            const dias1Formula = `=IF(AND(F${currentRow}<>"",G${currentRow}<>""),G${currentRow}-F${currentRow}+1,0)`;
            row.getCell(8).value = { formula: dias1Formula }; // H (8 = 1-indexed)
            
            // Fórmula para Días2 (columna K)
            const dias2Formula = `=IF(AND(I${currentRow}<>"",J${currentRow}<>""),J${currentRow}-I${currentRow}+1,0)`;
            row.getCell(11).value = { formula: dias2Formula }; // K (11 = 1-indexed)
            
            // Fórmula para Días3 (columna N)
            const dias3Formula = `=IF(AND(L${currentRow}<>"",M${currentRow}<>""),M${currentRow}-L${currentRow}+1,0)`;
            row.getCell(14).value = { formula: dias3Formula }; // N (14 = 1-indexed)
            
            // Fórmula para Días4 (columna Q)
            const dias4Formula = `=IF(AND(O${currentRow}<>"",P${currentRow}<>""),P${currentRow}-O${currentRow}+1,0)`;
            row.getCell(17).value = { formula: dias4Formula }; // Q (17 = 1-indexed)
            
            // ===== FÓRMULA PARA TOTAL DÍAS =====
            // Sumar todos los días de rangos + días sueltos (si hay)
            const totalFormula = `=H${currentRow}+K${currentRow}+N${currentRow}+Q${currentRow}+IF(R${currentRow}<>"",1,0)+IF(S${currentRow}<>"",1,0)`;
            const totalCell = row.getCell(20); // T (20 = 1-indexed)
            totalCell.value = { formula: totalFormula };
            totalCell.numFmt = '0';
            totalCell.alignment = { horizontal: 'center', vertical: 'center' };
            totalCell.font = { bold: true };
            
            // Aplicar color según el tipo
            if (tipo === 'Quincena') {
                totalCell.fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: 'FFE8F5E8' } // Verde claro
                };
            } else {
                totalCell.fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: 'FFF5F5F5' } // Gris claro
                };
            }
            
            // Aplicar formato numérico a las columnas de días (H, K, N, Q)
            [8, 11, 14, 17].forEach(colIndex => {
                const cell = row.getCell(colIndex);
                cell.numFmt = '0';
                cell.alignment = { horizontal: 'center', vertical: 'center' };
                cell.font = { bold: true };
                
                // Color de fondo para días calculados
                if (tipo === 'Quincena') {
                    cell.fill = {
                        type: 'pattern',
                        pattern: 'solid',
                        fgColor: { argb: 'FFF0F8E8' } // Verde muy claro
                    };
                } else {
                    cell.fill = {
                        type: 'pattern',
                        pattern: 'solid',
                        fgColor: { argb: 'FFF8F8F8' } // Gris muy claro
                    };
                }
            });
        });
        
        // ===== APLICAR FORMATO DE FECHA DD/MM/AAAA =====
        // Columnas de fechas: F, G, I, J, L, M, O, P, R, S
        const fechaColumns = [6, 7, 9, 10, 12, 13, 15, 16, 18, 19];
        
        fechaColumns.forEach(colIndex => {
            // Aplicar formato de fecha "dd/mm/aaaa" a toda la columna
            const column = worksheet.getColumn(colIndex);
            
            // Configurar formato de fecha español (dd/mm/aaaa)
            column.numFmt = 'dd/mm/yyyy';
            
            // Aplicar alineación centrada a toda la columna
            column.alignment = { horizontal: 'center', vertical: 'center' };
        });
        
        // ===== CALENDARIO VISUAL (A PARTIR DE LA COLUMNA U) =====
        const startCalendarCol = 21; // Columna U (1-indexed)
        const meses = [
            { nombre: 'ENERO', dias: 31 },
            { nombre: 'FEBRERO', dias: 28 }, // 2026 no es bisiesto
            { nombre: 'MARZO', dias: 31 },
            { nombre: 'ABRIL', dias: 30 },
            { nombre: 'MAYO', dias: 31 },
            { nombre: 'JUNIO', dias: 30 },
            { nombre: 'JULIO', dias: 31 },
            { nombre: 'AGOSTO', dias: 31 },
            { nombre: 'SEPTIEMBRE', dias: 30 },
            { nombre: 'OCTUBRE', dias: 31 },
            { nombre: 'NOVIEMBRE', dias: 30 },
            { nombre: 'DICIEMBRE', dias: 31 }
        ];
        
        let currentCol = startCalendarCol;
        
        // Primera fila del calendario (meses)
        meses.forEach((mes, mesIndex) => {
            if (mes.dias > 0) {
                const startCol = currentCol;
                const endCol = currentCol + mes.dias - 1;
                
                // Fusionar celdas del mes
                if (mes.dias > 1) {
                    worksheet.mergeCells(1, startCol, 1, endCol);
                }
                
                // Poner el nombre del mes
                const monthCell = worksheet.getCell(1, startCol);
                monthCell.value = mes.nombre;
                monthCell.font = { bold: true, size: 10 };
                monthCell.alignment = { horizontal: 'center', vertical: 'center' };
                monthCell.fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: 'FFF0F0F0' }
                };
                monthCell.border = {
                    top: { style: 'thin' },
                    left: { style: 'thin' },
                    bottom: { style: 'thin' },
                    right: { style: 'thin' }
                };
                
                // Segunda fila del calendario (días del mes)
                for (let dia = 1; dia <= mes.dias; dia++) {
                    const dayCell = worksheet.getCell(2, currentCol);
                    dayCell.value = dia;
                    dayCell.font = { bold: true, size: 9 };
                    dayCell.alignment = { horizontal: 'center', vertical: 'center' };
                    dayCell.border = {
                        top: { style: 'thin' },
                        left: { style: 'thin' },
                        bottom: { style: 'thin' },
                        right: { style: 'thin' }
                    };
                    
                    // Ajustar ancho de la columna
                    worksheet.getColumn(currentCol).width = 4;
                    
                    currentCol++;
                }
            }
        });
        
// ===== FÓRMULAS PARA EL CALENDARIO VISUAL =====
        employeesData.forEach((emp, empIndex) => {
            const rowNumber = empIndex + 3; // Fila del empleado
            let currentDayCol = startCalendarCol;
            
            // Para cada mes
            meses.forEach((mes, mesIndex) => {
                // Para cada día del mes
                for (let dia = 1; dia <= mes.dias; dia++) {
                    // CORRECCIÓN PREVIA: Sin el "="
                    const fechaExcel = `DATE(2026,${mesIndex + 1},${dia})`; 
                    
                    const colLetter = getExcelColumnLetter(currentDayCol);
                    
                    const formula = `=IF(OR(` +
                        `AND(F${rowNumber}<>"", G${rowNumber}<>"", ${fechaExcel}>=F${rowNumber},${fechaExcel}<=G${rowNumber}),` +
                        `AND(I${rowNumber}<>"", J${rowNumber}<>"", ${fechaExcel}>=I${rowNumber},${fechaExcel}<=J${rowNumber}),` +
                        `AND(L${rowNumber}<>"", M${rowNumber}<>"", ${fechaExcel}>=L${rowNumber},${fechaExcel}<=M${rowNumber}),` +
                        `AND(O${rowNumber}<>"", P${rowNumber}<>"", ${fechaExcel}>=O${rowNumber},${fechaExcel}<=P${rowNumber}),` +
                        `${fechaExcel}=R${rowNumber},${fechaExcel}=S${rowNumber}),1,0)`;
                    
                    const cell = worksheet.getCell(rowNumber, currentDayCol);
                    cell.value = { formula: formula };

                    // === NUEVA LÍNEA MÁGICA ===
                    // Esto oculta el texto (1 y 0) pero mantiene el valor para el formato condicional
                    cell.numFmt = ';;;'; 
                    
                    cell.alignment = { horizontal: 'center', vertical: 'center' };
                    cell.border = {
                        top: { style: 'hair' },
                        left: { style: 'hair' },
                        bottom: { style: 'hair' },
                        right: { style: 'hair' }
                    };
                    cell.font = { size: 8 };
                    
                    currentDayCol++;
                }
            });
        });
        
        // ===== APLICAR FORMATO CONDICIONAL AL CALENDARIO =====
        const lastRow = employeesData.length + 2;
        const lastCalendarCol = startCalendarCol + 365 - 1; // 365 días del año
        
        // Calcular el rango del calendario (desde U3 hasta la última celda)
        const startColLetter = getExcelColumnLetter(startCalendarCol);
        const endColLetter = getExcelColumnLetter(lastCalendarCol);
        const calendarRange = `${startColLetter}3:${endColLetter}${lastRow}`;
        
        // Agregar formato condicional usando el tipo correcto
        // En ExcelJS, el formato condicional debe usar referencias relativas a la celda superior izquierda
        // Para un rango U3:ZZ100, la fórmula debe referirse a U3 (celda superior izquierda)
        
        worksheet.addConditionalFormatting({
            ref: calendarRange,
            rules: [
                {
                    type: 'cellIs',
                    operator: 'equal',
                    formulae: ['1'],
                    style: {
                        fill: {
                            type: 'pattern',
                            pattern: 'solid',
                            fgColor: { argb: 'FFFF0000' } // Rojo
                        },
                        font: {
                            color: { argb: 'FFFFFFFF' }, // Blanco
                            bold: true,
                            size: 8
                        }
                    }
                }
            ]
        });
        
        // También podemos agregar un formato para celdas vacías (opcional)
        worksheet.addConditionalFormatting({
            ref: calendarRange,
            rules: [
                {
                    type: 'cellIs',
                    operator: 'equal',
                    formulae: ['0'],
                    style: {
                        fill: {
                            type: 'pattern',
                            pattern: 'solid',
                            fgColor: { argb: 'FFFFFFFF' } // Blanco
                        },
                        font: {
                            color: { argb: 'FFCCCCCC' }, // Gris claro
                            size: 8
                        }
                    }
                }
            ]
        });
        
        // ===== NOTA EXPLICATIVA =====
        // Agregar una fila con instrucciones
        const noteRow = worksheet.getRow(lastRow + 2);
        noteRow.getCell(1).value = 'INSTRUCCIONES:';
        noteRow.getCell(1).font = { bold: true, color: { argb: 'FF0000FF' } };
        
        const noteRow2 = worksheet.getRow(lastRow + 3);
        noteRow2.getCell(1).value = '1. El calendario muestra los días de vacaciones en rojo (valor 1 = vacaciones, 0 = trabajo)';
        noteRow2.getCell(1).font = { size: 10 };
        
        const noteRow3 = worksheet.getRow(lastRow + 4);
        noteRow3.getCell(1).value = '2. Los rangos se pueden modificar y las fórmulas recalcularán automáticamente';
        noteRow3.getCell(1).font = { size: 10 };
        
        const noteRow4 = worksheet.getRow(lastRow + 5);
        noteRow4.getCell(1).value = '3. Para días sueltos, escribir la fecha en las columnas R y S (DiaSuelto1 y DiaSuelto2)';
        noteRow4.getCell(1).font = { size: 10 };
        
        // ===== CONFIGURACIÓN FINAL =====
        
        // Ajustar altura de las dos primeras filas
        worksheet.getRow(1).height = 25;
        worksheet.getRow(2).height = 20;
        
        // Ajustar altura de las filas de datos
        for (let i = 3; i <= employeesData.length + 2; i++) {
            worksheet.getRow(i).height = 18;
        }
        
        // Congelar las dos primeras filas y las primeras 20 columnas
        worksheet.views = [
            { 
                state: 'frozen', 
                xSplit: 20, // Congelar columnas A-T
                ySplit: 2,  // Congelar filas 1-2
                activeCell: 'A3',
                showGridLines: true
            }
        ];
        
        // Aplicar filtros a la segunda fila (solo columnas A-T)
        worksheet.autoFilter = {
            from: 'A2',
            to: 'T2'
        };
        
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

// ===== FUNCIONES HELPER =====

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