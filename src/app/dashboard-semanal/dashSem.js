const XLSX = require('xlsx');

/**
 * Procesa el archivo Excel semanal y extrae las métricas del Dashboard.
 * @param {Buffer|ArrayBuffer} fileBuffer - El archivo Excel subido por el usuario.
 * @param {string} tipoArchivo - 'TIENDAS', 'AGRICULTORES' o 'VIAJEROS'
 */
function procesarExcelSemanal(fileBuffer, tipoArchivo) {
    // 1. Leer el libro de trabajo de Excel
    const workbook = XLSX.read(fileBuffer, { type: 'buffer' });
    
    // Convertir la hoja principal a un arreglo de objetos JSON
    // Se recomienda que la hoja de datos crudos sea la primera o se busque por nombre (ej: 'Detalle regiones')
    const sheetName = workbook.SheetNames[0]; 
    const sheetData = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);

    // Variables de control para el Dashboard
    let totalViajes = 0;
    let viajesCumple = 0;
    let numeroSemana = null;
    const detallesDashboard = [];

    // 2. Mapeo dinámico según la estructura real de cada archivo
    sheetData.forEach(fila => {
        let semana, entidad, estado, vehiculo, tiempo;

        if (tipoArchivo === 'TIENDAS') {
            semana = fila['SEMANA'];
            entidad = fila['Región'];
            estado = fila['CUMPLIMIENTO']; // 'CUMPLE' o 'NO CUMPLE'
            vehiculo = fila['Vehículo'];
            tiempo = fila['TIEMPO EN REGION'];
        } 
        else if (tipoArchivo === 'AGRICULTORES') {
            semana = fila['semana'];
            entidad = fila['AGRICULTOR'];
            estado = fila['CUMPLIMIENTO LLEGADA AGRICULTOR'];
            vehiculo = fila['PLACA'] || fila['PRODUCTO'];
            tiempo = fila['TIEMPO EN LLEGAR ACRICULTOR'];
        } 
        else if (tipoArchivo === 'VIAJEROS') {
            semana = fila['SEMANAS'];
            entidad = fila['SUCURSAL'];
            estado = fila['CUMPLIMIENTO LLEGADA VEHICULO'] || fila['CUMPLIMIENTO CITA'];
            vehiculo = fila['PLACA'];
            tiempo = fila['TIEMPO LLEGADA VEHICULO'];
        }

        // Validar que la fila contenga información útil
        if (semana && estado) {
            if (!numeroSemana) numeroSemana = semana; // Captura la semana del reporte
            
            totalViajes++;
            if (estado.trim().toUpperCase() === 'CUMPLE') {
                viajesCumple++;
            }

            // Guardar el registro limpio para la tabla detallada del Dashboard
            detallesDashboard.push({
                modulo: tipoArchivo,
                entidad: entidad || 'No especificado',
                vehiculo: vehiculo || 'N/A',
                tiempo_registro: tiempo || '00:00:00',
                estado: estado.trim().toUpperCase()
            });
        }
    });

    // 3. Calcular porcentaje matemático de efectividad semanal
    // Fórmula estándar: (Cumple / Total) * 100
    const porcentajeEfectividad = totalViajes > 0 
        ? parseFloat(((viajesCumple / totalViajes) * 100).toFixed(1)) 
        : 0;

    // 4. Retornar la estructura final procesada
    return {
        semana: numeroSemana,
        metricas: {
            total_viajes: totalViajes,
            viajes_cumplen: viajesCumple,
            viajes_no_cumplen: totalViajes - viajesCumple,
            efectividad_global: porcentajeEfectividad
        },
        data_graficos: detallesDashboard
    };
}