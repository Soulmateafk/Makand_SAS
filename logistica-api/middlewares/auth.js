const jwt = require('jsonwebtoken');

const verificarAuth = (req, res, next) => {
  // 1. Capturamos la cabecera Authorization
  const authHeader = req.headers['authorization'];
  
  // 🔍 LOG DE CONTROL: Ver si Angular está enviando algo
  console.log('--- NUEVA PETICIÓN PROTEGIDA ---');
  console.log('Cabecera Authorization recibida:', authHeader);

  // El formato suele ser "Bearer TOKEN", así que separamos el texto del token real
  const token = authHeader && authHeader.split(' ')[1]; 

  if (!token) {
    console.log('❌ RECHAZADO: No se encontró ningún token en la petición.');
    return res.status(401).json({ message: 'Acceso denegado. No hay token.' });
  }

  try {
    // ⚠️ SEGURIDAD ESTRICTA: Usamos solo process.env.JWT_SECRET
    const claveSecreta = process.env.JWT_SECRET;

    // Si por algún error de configuración la variable no llega, detenemos la petición
    if (!claveSecreta) {
        console.error('❌ ERROR CRÍTICO: JWT_SECRET no está definido en el archivo .env');
        return res.status(500).json({ message: 'Error de configuración del servidor' });
    }
    
    // Verificamos y desciframos el token
    const verificado = jwt.verify(token, claveSecreta);
    req.user = verificado; 
    
    console.log('✅ PERMITIDO: Token válido para el usuario:', verificado.email);
    next(); // Continuar a la ruta (ej: /api/vehiculos)
  } catch (error) {
    console.log('❌ RECHAZADO: El token llegó pero es inválido. Motivo:', error.message);
    return res.status(401).json({ message: 'Token inválido o expirado' });
  }
};

module.exports = verificarAuth;