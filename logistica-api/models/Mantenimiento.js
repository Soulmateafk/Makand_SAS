const mongoose = require('mongoose');

const MantenimientoSchema = new mongoose.Schema({
  // Información General
  placa: { 
    type: String, 
    required: true, 
    uppercase: true, 
    trim: true 
  },
  marcaModelo: { type: String },
  numMotor: { type: String },
  conductor: { type: String, required: true },
  
  // Datos de Servicio
  kilometraje: { type: Number, default: 0 },
  kilometrajeProximo: { type: Number, default: 0 }, 
  fechaEntrada: { type: String, required: true }, 
  fechaProxima: { type: String, required: true },
  tipoIntervencion: { 
    type: String, 
    enum: ['Preventivo', 'Correctivo', 'Predictivo'], 
    default: 'Preventivo' 
  },
  taller: { type: String },
  
  // CORREGIDO: Se añade 'Operativo' al enum para que coincida con tu frontend
  estadoVehiculo: { 
    type: String, 
    enum: ['En Taller', 'Operativo', 'Finalizado'], 
    default: 'En Taller' 
  },
  
  // Detalles Técnicos y Costos
  descripcion: { type: String },
  repuestos: { type: String },
  costoManoObra: { type: Number, default: 0 },
  costoTotal: { type: Number, default: 0 },
  
  // Documentación
  vencimientoSoat: { type: String },
  vencimientoTecno: { type: String },
  diasAlertaPrevia: { type: Number, default: 15 }, 
  
  // Checklist (Objeto anidado)
  checklist: {
    frenos: { type: Boolean, default: false },
    aceite: { type: Boolean, default: false },
    neumaticos: { type: Boolean, default: false },
    bateria: { type: Boolean, default: false },
    luces: { type: Boolean, default: false },
    refrigerante: { type: Boolean, default: false }
  },
  
  // Evidencias (Array de Strings para Base64 o URLs)
  fotos: [{ type: String }],
  
  // Metadatos automáticos
  createdAt: { type: Date, default: Date.now }
});

module.exports = mongoose.model('Mantenimiento', MantenimientoSchema);