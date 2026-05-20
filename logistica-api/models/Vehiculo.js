const mongoose = require('mongoose');

const VehiculoSchema = new mongoose.Schema({
  // Identificación y Marca
  placa: { 
    type: String, 
    required: true, 
    unique: true, 
    uppercase: true 
  },
  marcaModelo: { type: String, default: "" }, // Coincide con [(ngModel)]="vehiculo.marcaModelo"
  anio: { type: Number, default: null },       // Coincide con [(ngModel)]="vehiculo.anio"
  numFlota: { type: String, default: "" },    // Coincide con [(ngModel)]="vehiculo.numFlota"

  // Documentación Legal
  soat: { type: Date, default: null },
  tecno: { type: Date, default: null },
  seguro: { type: Date, default: null },

  // Estado y Operación
  kmActual: { type: Number, default: 0 },     // Coincide con [(ngModel)]="vehiculo.kmActual"
  kmProximo: { type: Number, default: 0 },    // Coincide con [(ngModel)]="vehiculo.kmProximo"
  estado: { 
    type: String, 
    enum: ['Bueno', 'Regular', 'Malo', 'En Taller'], // Ajustado a tu HTML
    default: 'Bueno' 
  },

  // Responsables
  conductor: { type: String, default: "" },
  area: { type: String, default: "" }
});

module.exports = mongoose.model('Vehiculo', VehiculoSchema);