const mongoose = require('mongoose');
const ExcelJS = require('exceljs'); // <-- Asegúrate de tener exceljs instalado

const CombustibleSchema = new mongoose.Schema({
  placa: { type: String, required: true },
  conductor: { type: String, required: true },
  fecha: { type: Date, required: true },
  kmActual: { type: Number, required: true },
  galones: { type: Number, required: true },
  precioTotal: { type: Number, required: true },
  estacionServicio: { type: String },
  metodoPago: { type: String },
  kmRecorridos: { type: Number, default: 0 },
  rendimientoKmpg: { type: Number, default: 0 },
  costoPorKm: { type: Number, default: 0 }
}, { 
  versionKey: false,
  collection: 'combustibles' // Fuerza el nombre exacto de la tabla en MongoDB Compass
});

// 🔥 METODO ESTÁTICO: Para generar y descargar el Excel directamente desde las rutas
CombustibleSchema.statics.exportarAExcel = async function (filtros, res) {
  try {
    const { placa, fechaInicio, fechaFin, fechaExacta } = filtros;
    let query = {};

    // Filtro opcional por placa
    if (placa) {
      query.placa = { $regex: new RegExp(placa, 'i') };
    }

    // Filtro por fecha exacta o rango de fechas
    if (fechaExacta) {
      const inicioDia = new Date(fechaExacta);
      inicioDia.setHours(0, 0, 0, 0);
      const finDia = new Date(fechaExacta);
      finDia.setHours(23, 59, 59, 999);
      query.fecha = { $gte: inicioDia, $lte: finDia };
    } 
    else if (fechaInicio || fechaFin) {
      query.fecha = {};
      if (fechaInicio) query.fecha.$gte = new Date(`${fechaInicio}T00:00:00.000Z`);
      if (fechaFin) query.fecha.$lte = new Date(`${fechaFin}T23:59:59.999Z`);
    }

    // Consultar la base de datos con los filtros
    const registros = await this.find(query).sort({ fecha: -1 });

    // Crear el libro de ExcelJS
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Reporte de Combustible');

    worksheet.columns = [
      { header: 'Fecha', key: 'fecha', width: 15 },
      { header: 'Placa', key: 'placa', width: 12 },
      { header: 'Conductor', key: 'conductor', width: 20 },
      { header: 'KM Registro', key: 'kmActual', width: 15 },
      { header: 'KM Recorridos', key: 'kmRecorridos', width: 15 },
      { header: 'Galones', key: 'galones', width: 12 },
      { header: 'Valor Total', key: 'precioTotal', width: 15 },
      { header: 'Rendimiento Real', key: 'rendimientoKmpg', width: 18 },
      { header: 'Costo / KM', key: 'costoPorKm', width: 15 },
      { header: 'Estación de Servicio', key: 'estacionServicio', width: 22 },
      { header: 'Método de Pago', key: 'metodoPago', width: 18 }
    ];

    // Estilos visuales para la cabecera
    worksheet.getRow(1).font = { name: 'Arial', bold: true, color: { argb: 'FFFFFF' } };
    worksheet.getRow(1).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: '4B1A7A' }
    };
    worksheet.getRow(1).alignment = { vertical: 'middle', horizontal: 'center' };

    // Inyectar filas
    registros.forEach(reg => {
      worksheet.addRow({
        fecha: reg.fecha ? reg.fecha.toISOString().split('T')[0] : '---',
        placa: reg.placa,
        conductor: reg.conductor,
        kmActual: reg.kmActual,
        kmRecorridos: reg.kmRecorridos || 0,
        galones: reg.galones,
        precioTotal: reg.precioTotal,
        rendimientoKmpg: reg.rendimientoKmpg || 0,
        costoPorKm: reg.costoPorKm || 0,
        estacionServicio: reg.estacionServicio || '---',
        metodoPago: reg.metodoPago || '---'
      });
    });

    // Configurar cabeceras de respuesta HTTP para forzar la descarga en el navegador
    res.setHeader(
      'Content-Type',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    );
    res.setHeader(
      'Content-Disposition',
      `attachment; filename=reporte-combustible-${Date.now()}.xlsx`
    );

    await workbook.xlsx.write(res);
    res.end();

  } catch (error) {
    console.error('Error al generar Excel desde el modelo:', error);
    if (!res.headersSent) {
      res.status(500).json({ error: 'No se pudo generar el reporte en Excel' });
    }
  }
};

// 🔥 HOOK AUTOMÁTICO: Con Verificaciones y Seguridad Integradas
CombustibleSchema.pre('save', async function (next) {
  const documentoActual = this;

  try {
    // 🔍 VERIFICACIÓN 1: Validar que los valores numéricos del formulario sean positivos y válidos
    if (documentoActual.kmActual <= 0 || documentoActual.galones <= 0 || documentoActual.precioTotal <= 0) {
      return next(new Error('Validación fallida: Kilometraje, galones y precio total deben ser números mayores a cero.'));
    }

    // Busca dinámicamente el último registro en la colección usando Mongoose
    const registroAnterior = await mongoose.model('Combustible').findOne({
      placa: documentoActual.placa, 
      fecha: { $lt: documentoActual.fecha } 
    }).sort({ fecha: -1 });

    // Si existe historial previo de este vehículo, ejecutamos los cálculos automáticos
    if (registroAnterior && registroAnterior.kmActual) {
      const kmAnterior = registroAnterior.kmActual;

      // 🔍 VERIFICACIÓN 2: Validar inconsistencia de kilometraje (Evita kilómetros en reversa)
      if (documentoActual.kmActual < kmAnterior) {
        return next(new Error(`Validación fallida: El kilometraje actual (${documentoActual.kmActual} KM) no puede ser menor al del tanqueo anterior (${kmAnterior} KM).`));
      }

      // 🔍 VERIFICACIÓN 3: Validar kilometraje idéntico (Evita duplicados o registros sospechosos)
      if (documentoActual.kmActual === kmAnterior) {
        return next(new Error(`Validación fallida: El kilometraje ingresado es idéntico al registro anterior (${kmAnterior} KM). El vehículo debe haber recorrido alguna distancia.`));
      }

      // 1. Kilómetros Recorridos
      documentoActual.kmRecorridos = documentoActual.kmActual - kmAnterior;

      // 2. Rendimiento Real (KM / Galones) - Verificando que galones sea mayor a 0
      if (documentoActual.galones > 0) {
        const rendimiento = documentoActual.kmRecorridos / documentoActual.galones;
        documentoActual.rendimientoKmpg = Number(rendimiento.toFixed(2));
      }

      // 3. Costo por Kilómetro ($ / KM) - Verificando que kilómetros recorridos sea mayor a 0
      if (documentoActual.kmRecorridos > 0) {
        const costo = documentoActual.precioTotal / documentoActual.kmRecorridos;
        documentoActual.costoPorKm = Number(costo.toFixed(2));
      }

    } else {
      // Si es el primer registro de esta placa, los valores por defecto quedan en 0
      documentoActual.kmRecorridos = 0;
      documentoActual.rendimientoKmpg = 0;
      documentoActual.costoPorKm = 0;
    }

    next(); // Permite que Mongoose termine de guardar el registro en MongoDB de forma segura
  } catch (error) {
    next(error); // Si algo falla, cancela el guardado para proteger la consistencia de la base de datos
  }
});

// Exportamos como objeto para que el require del server.js funcione idéntico
module.exports = { 
  Combustible: mongoose.model('Combustible', CombustibleSchema) 
};