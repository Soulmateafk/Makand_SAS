const mongoose = require('mongoose');

const ConductorSchema = new mongoose.Schema({
  nombre: { type: String, required: true },
  cedula: { type: String, required: true, unique: true },
  fechaNacimiento: { type: String }, 
  telefono: { type: String },
  fechaUnion: { type: String },      
  email: { type: String }
});

module.exports = mongoose.model('Conductor', ConductorSchema);