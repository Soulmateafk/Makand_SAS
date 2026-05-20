const mongoose = require('mongoose');

const EntregaSchema = new mongoose.Schema({
  placa: { type: String, required: true },
  nombre: String, 
  origen: String, 
  entregas: { type: Number, default: 0 }, 
  devoluciones: { type: Number, default: 0 },
  conductor: String, 
  estadoProducto: String, 
  explicacion: String, 
  foto: String, // Base64 de la imagen
  tasa: String,
  // Aseguramos que la fecha sea siempre un objeto Date para los filtros
  fecha: { 
    type: Date, 
    default: Date.now,
    index: true // Añadimos índice para que buscar por años/semanas sea veloz
  }
});

// Tip técnico: Este índice compuesto ayuda si luego quieres buscar 
// "todas las entregas de X placa en tal año"
EntregaSchema.index({ placa: 1, fecha: -1 });

module.exports = mongoose.model('Entrega', EntregaSchema);