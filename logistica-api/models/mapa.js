// models/Mapa.js
const mapaSchema = new mongoose.Schema({
  // ... tus otros campos existentes ...
  nombre: String,
  conductor: String,
  // ... AGREGA ESTO:
  lat: { type: Number, required: true }, 
  lon: { type: Number, required: true }
});