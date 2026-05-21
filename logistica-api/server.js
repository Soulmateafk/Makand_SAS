require('dotenv').config(); // Cargar variables de entorno
const express = require('express');
const cors = require('cors');
const helmet = require('helmet'); // Seguridad adicional
const mongoose = require('mongoose'); 
const jwt = require('jsonwebtoken');
const app = express();

// Definimos el puerto aquí para que esté disponible en todo el archivo
const PORT = process.env.PORT || 3000;

// 🛡️ IMPORTAR EL VIGILANTE DE SEGURIDAD (MIDDLEWARE)
const verificarAuth = require('./middlewares/auth');

// 🛠️ CONFIGURACIÓN DE SEGURIDAD Y CORS
app.use(helmet()); 
app.use(cors({
  origin: '*', // Permitido acceso general para evitar bloqueos
  methods: ['GET', 'POST', 'PUT', 'DELETE'],
  allowedHeaders: ['Content-Type', 'Authorization']
}));

app.use(express.json({ limit: '50mb' })); 
app.use(express.urlencoded({ limit: '50mb', extended: true }));

// RUTA DE SALUD DEL SERVIDOR
app.get('/', (req, res) => res.send('API de Logística funcionando correctamente 🚀'));

const mongoURI = process.env.MONGO_URI || 'mongodb://127.0.0.1:27017/logistica'; 

// --- MODIFICACIÓN: LA CONEXIÓN Y EL LISTEN AHORA VAN JUNTOS ---
// Hemos eliminado el app.listen inicial para evitar conflictos y asegurar que la BD conecte primero
mongoose.connect(mongoURI)
  .then(() => {
    console.log("📍 Conectado a:", process.env.MONGO_URI ? "ATLAS (Nube)" : "LOCAL (Tu PC)");
    
    // El servidor solo arranca una vez confirmada la conexión y usando la variable correcta PORT
    app.listen(PORT, () => { 
        console.log(`🚀 Servidor listo en http://localhost:${PORT}`); 
    });
  })
  .catch(err => {
    console.error("❌ ERROR CRÍTICO AL CONECTAR A LA BD:", err);
  });
// -----------------------------------------------------------

const Formulario = require('./models/entrega'); 
const Vehiculo = require('./models/Vehiculo');
const Conductor = require('./models/Conductor');
const Mantenimiento = require('./models/Mantenimiento');
const Usuario = require('./models/Usuario');
const { Combustible } = require('./models/combustible'); 

// ==========================================
// RUTA: LOGIN (CORREGIDA Y SEGURA)
// ==========================================
app.post('/api/login', async (req, res) => {
  const { email, password } = req.body;
  try {
    const user = await Usuario.findOne({ email, password });
    if (user) {
      const claveSecreta = process.env.JWT_SECRET;
      
      if (!claveSecreta) {
        throw new Error('FATAL: JWT_SECRET no está definido en el archivo .env');
      }

      const token = jwt.sign(
        { id: user._id, nombre: user.nombre, email: user.email, rol: user.rol }, 
        claveSecreta, 
        { expiresIn: '10h' }
      );
      res.json({ message: "Login exitoso", token, user: { nombre: user.nombre, email: user.email, rol: user.rol } });
    } else {
      res.status(401).json({ message: "Credenciales inválidas" });
    }
  } catch (error) {
    console.error("--- ERROR EN EL SERVIDOR ---");
    console.error(error); 
    res.status(500).json({ error: "Error en el servidor", detalles: error.message });
  }
}); 

// ==========================================
// RUTAS: ENTREGAS
// ==========================================
app.get('/api/entregas', verificarAuth, async (req, res) => {
  try {
    const lista = await Formulario.find().sort({ fecha: -1 });
    res.json(lista);
  } catch (error) { res.status(500).json({ error: 'Error al obtener entregas', detalles: error.message }); }
});

app.get('/api/entregas/historial', verificarAuth, async (req, res) => {
  try {
    const { anio, semana, fechaExacta } = req.query;
    let filtro = {};
    if (fechaExacta) {
      const inicioDia = new Date(fechaExacta); inicioDia.setHours(0, 0, 0, 0);
      const finDia = new Date(fechaExacta); finDia.setHours(23, 59, 59, 999);
      filtro.fecha = { $gte: inicioDia, $lte: finDia };
    } else if (anio && semana) {
      const primeroEnero = new Date(anio, 0, 1);
      const diasAMarzo = (semana - 1) * 7;
      const inicioSemana = new Date(anio, 0, 1 + diasAMarzo);
      const finSemana = new Date(anio, 0, 1 + diasAMarzo + 6);
      finSemana.setHours(23, 59, 59, 999);
      filtro.fecha = { $gte: inicioSemana, $lte: finSemana };
    } else if (anio) {
      const inicioAnio = new Date(anio, 0, 1);
      const finAnio = new Date(anio, 11, 31, 23, 59, 59);
      filtro.fecha = { $gte: inicioAnio, $lte: finAnio };
    }
    const resultados = await Formulario.find(filtro).sort({ fecha: -1 });
    res.json(resultados);
  } catch (error) { res.status(500).json({ error: 'Error en el filtro de historial', detalles: error.message }); }
});

app.get('/api/mapa', verificarAuth, async (req, res) => {
  try { 
    const lista = await Formulario.find().sort({ fecha: -1 }); 
    const listaConCoordenadas = lista.map(item => {
      const doc = item.toObject ? item.toObject() : item;
      return {
        ...doc,
        lat: doc.lat || 4.7110,  
        lon: doc.lon || -74.0721 
      };
    });
    res.json(listaConCoordenadas); 
  } 
  catch (error) { 
    res.status(500).json({ error: 'Error al obtener datos' }); 
  }
});

app.post('/api/entregas', verificarAuth, async (req, res) => {
  try {
    const datos = req.body;
    if (!datos.fecha) datos.fecha = new Date();
    const nuevoRegistro = new Formulario(datos);
    await nuevoRegistro.save();
    res.status(201).json(nuevoRegistro);
  } catch (error) { res.status(400).json({ error: 'Error al guardar registro' }); }
});

app.put('/api/entregas/:id', verificarAuth, async (req, res) => {
  try { const actualizado = await Formulario.findByIdAndUpdate(req.params.id, req.body, { new: true }); res.json(actualizado); } 
  catch (error) { res.status(400).json({ error: 'Error al actualizar' }); }
});

app.delete('/api/entregas/:id', verificarAuth, async (req, res) => {
  try { await Formulario.findByIdAndDelete(req.params.id); res.json({ message: "Eliminado con éxito" }); } 
  catch (error) { res.status(500).json({ error: 'Error al eliminar' }); }
});

// ==========================================
// RUTAS: CONDUCTORES
// ==========================================
app.get('/api/conductores', verificarAuth, async (req, res) => {
  try { const lista = await Conductor.find().sort({ _id: -1 }); res.json(lista); } 
  catch (error) { res.status(500).json({ error: 'Error al obtener conductores' }); }
});

app.post('/api/conductores', verificarAuth, async (req, res) => {
  try { const nuevo = new Conductor({ ...req.body, estado: "Activo" }); const guardado = await nuevo.save(); res.status(201).json(guardado); } 
  catch (error) { if (error.code === 11000) return res.status(400).json({ message: "La cédula ya existe" }); res.status(400).json({ message: "Error al guardar conductor", detalles: error.message }); }
});

app.put('/api/conductores/:id', verificarAuth, async (req, res) => {
  try { const actualizado = await Conductor.findByIdAndUpdate(req.params.id, { $set: req.body }, { new: true, runValidators: true }); res.json(actualizado); } 
  catch (error) { res.status(400).json({ message: "No se pudo actualizar", detalles: error.message }); }
});

app.delete('/api/conductores/:id', verificarAuth, async (req, res) => {
  try { await Conductor.findByIdAndDelete(req.params.id); res.json({ message: "Conductor eliminado con éxito" }); } 
  catch (error) { res.status(500).json({ error: 'Error al eliminar conductor' }); }
});

// ==========================================
// RUTAS: VEHÍCULOS
// ==========================================
app.get('/api/vehiculos', verificarAuth, async (req, res) => {
  try { const lista = await Vehiculo.find().sort({ _id: -1 }); res.json(lista); } 
  catch (error) { res.status(500).json({ message: "Error al obtener vehículos", error }); }
});

app.post('/api/vehiculos', verificarAuth, async (req, res) => {
  try { const nuevo = new Vehiculo(req.body); await nuevo.save(); res.status(201).json(nuevo); } 
  catch (error) { if (error.code === 11000) return res.status(400).json({ message: "La placa ya existe" }); res.status(400).json({ message: "Error de validación", detalles: error.message }); }
});

app.put('/api/vehiculos/:id', verificarAuth, async (req, res) => {
  try { const actualizado = await Vehiculo.findByIdAndUpdate(req.params.id, req.body, { new: true }); res.json(actualizado); }
  catch (error) { res.status(400).json({ message: "No se pudo actualizar", error }); }
});

app.delete('/api/vehiculos/:id', verificarAuth, async (req, res) => {
  try { await Vehiculo.findByIdAndDelete(req.params.id); res.json({ message: "Vehículo eliminado" }); }
  catch (error) { res.status(500).json({ message: "Error al eliminar", error }); }
});

// ==========================================
// RUTAS: AUTENTICACIÓN
// ==========================================
app.post('/api/register', async (req, res) => {
  try { const nuevoUsuario = new Usuario(req.body); await nuevoUsuario.save(); res.status(201).json({ message: "Usuario creado correctamente" }); } 
  catch (error) { res.status(400).json({ message: "Error al crear usuario", error }); }
});

// ==========================================
// RUTAS: MANTENIMIENTOS
// ==========================================
const fnGetMant = async (req, res) => {
  try { const lista = await Mantenimiento.find().sort({ fechaEntrada: -1 }); res.json(lista); } 
  catch (error) { res.status(500).json({ message: "Error", detalles: error.message }); }
};
app.get('/api/mantenimiento', verificarAuth, fnGetMant);
app.get('/api/mantenimientos', verificarAuth, fnGetMant);

const fnPostMant = async (req, res) => {
  try { const nuevo = new Mantenimiento(req.body); await nuevo.save(); res.status(201).json(nuevo); } 
  catch (error) { res.status(400).json({ message: "Error", detalles: error.message }); }
};
app.post('/api/mantenimiento', verificarAuth, fnPostMant);
app.post('/api/mantenimientos', verificarAuth, fnPostMant);

const fnPutMant = async (req, res) => {
  try {
    const { id } = req.params;
    const { _id, id: bodyId, __v, ...datosLimpios } = req.body;
    if (datosLimpios.vehiculo && typeof datosLimpios.vehiculo === 'object') { datosLimpios.vehiculo = datosLimpios.vehiculo._id || datosLimpios.vehiculo; }
    const actualizado = await Mantenimiento.findByIdAndUpdate(id, { $set: datosLimpios }, { new: true, runValidators: true });
    if (!actualizado) return res.status(404).json({ message: "No encontrado" });
    res.json(actualizado);
  } catch (error) { res.status(400).json({ message: "Error al actualizar", detalles: error.message }); }
};
app.put('/api/mantenimiento/:id', verificarAuth, fnPutMant);
app.put('/api/mantenimientos/:id', verificarAuth, fnPutMant);

const fnDeleteMant = async (req, res) => {
  try { const eliminado = await Mantenimiento.findByIdAndDelete(req.params.id); if (!eliminado) return res.status(404).json({ message: "No encontrado" }); res.json({ message: "Eliminado con éxito" }); } 
  catch (error) { res.status(500).json({ message: "Error al eliminar" }); }
};
app.delete('/api/mantenimiento/:id', verificarAuth, fnDeleteMant);
app.delete('/api/mantenimientos/:id', verificarAuth, fnDeleteMant);

app.get('/api/mantenimiento/stats', verificarAuth, async (req, res) => {
  try {
    const todos = await Mantenimiento.find({}, 'estadoVehiculo fechaProxima kilometraje kilometrajeProximo');
    const stats = { total: todos.length, enTaller: 0, operativos: 0, vencidosFecha: 0, alertaKilometraje: 0 };
    const hoy = new Date(); hoy.setHours(0, 0, 0, 0);
    todos.forEach(m => {
      if (m.estadoVehiculo === 'En Taller') stats.enTaller++;
      if (m.estadoVehiculo === 'Operativo') stats.operativos++;
      if (m.fechaProxima) { const fProxima = new Date(m.fechaProxima); if (fProxima <= hoy) stats.vencidosFecha++; }
      if (m.kilometrajeProximo && m.kilometraje) { const margen = m.kilometrajeProximo * 0.90; if (m.kilometraje >= margen) stats.alertaKilometraje++; }
    });
    res.json(stats);
  } catch (error) { res.status(500).json({ error: 'No se pudieron calcular las estadísticas' }); }
});

// ==========================================
// RUTAS: COMBUSTIBLE
// ==========================================
app.get('/api/combustibles/exportar-excel', verificarAuth, async (req, res) => {
  try { await Combustible.exportarAExcel(req.query, res); } catch (error) { res.status(500).json({ error: 'Error al procesar la exportación' }); }
});

app.post('/api/combustible', verificarAuth, async (req, res) => {
  try {
    const { placa, conductor, fecha, kmActual, galones, precioTotal, estacionServicio, metodoPago } = req.body;
    const ultimoTanqueo = await Combustible.findOne({ placa: placa.toUpperCase() }).sort({ kmActual: -1 });
    let kmRecorridos = 0, rendimientoKmpg = 0, costoPorKm = 0;
    if (ultimoTanqueo && kmActual > ultimoTanqueo.kmActual) {
      kmRecorridos = kmActual - ultimoTanqueo.kmActual;
      if (galones > 0) rendimientoKmpg = kmRecorridos / galones;
      if (kmRecorridos > 0) costoPorKm = precioTotal / kmRecorridos;
    }
    const nuevoRegistro = new Combustible({ placa, conductor, fecha, kmActual, galones, precioTotal, estacionServicio, metodoPago, kmRecorridos, rendimientoKmpg, costoPorKm });
    await nuevoRegistro.save();
    res.status(201).json(nuevoRegistro);
  } catch (error) { res.status(500).json({ error: error.message }); }
});

app.get('/api/combustible', verificarAuth, async (req, res) => {
  try { const historial = await Combustible.find().sort({ fecha: -1 }); res.status(200).json(historial); } 
  catch (error) { res.status(500).json({ error: error.message }); }
});

app.delete('/api/combustible/:id', verificarAuth, async (req, res) => {
  try { await Combustible.findByIdAndDelete(req.params.id); res.status(200).json({ message: 'Eliminado correctamente' }); } 
  catch (error) { res.status(500).json({ error: error.message }); }
});

// NUEVO: MANEJADOR DE ERRORES GLOBAL
app.use((err, req, res, next) => {
  console.error(err.stack);
  res.status(500).json({ message: 'Algo salió mal en el servidor', error: err.message });
});