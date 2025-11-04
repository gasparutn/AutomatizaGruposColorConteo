function actualizarContadorGruposCompleto() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaRegistro = ss.getSheetByName(NOMBRE_HOJA_REGISTRO);
  const hojaContador = ss.getSheetByName(NOMBRE_HOJA_CONTADOR);

  if (!hojaRegistro) {
    Logger.log(`Error: No se encontró la hoja "${NOMBRE_HOJA_REGISTRO}"`);
    return;
  }
  if (!hojaContador) {
    Logger.log(`Error: No se encontró la hoja "${NOMBRE_HOJA_CONTADOR}"`);
    return;
  }

  const ultimaFilaRegistros = hojaRegistro.getLastRow();
  if (ultimaFilaRegistros < 2) {
    Logger.log("No hay datos en Registros.");
    return; // No hay datos para contar
  }

  // 1. Leer solo las columnas necesarias (de A hasta I)
  // Leemos hasta la COL_GRUPOS (Col I, número 9) 
  // Esto incluye automáticamente la COL_MARCA_N_E_A (Col C, número 3) 
  const rangoDatosRegistros = hojaRegistro.getRange(2, 1, ultimaFilaRegistros - 1, COL_GRUPOS); 
  const valoresRegistros = rangoDatosRegistros.getValues();

  // 2. Procesar los datos y contarlos en un mapa
  const mapaContador = {};

  for (let i = 0; i < valoresRegistros.length; i++) {
    const fila = valoresRegistros[i];
    
    // --- (INICIO DE LA CORRECCIÓN) ---
    // Usar las constantes correctas según tu aclaración
    
    // Columna I (Grupos) 
    const grupo = fila[COL_GRUPOS - 1]; 
    
    // Columna C (N/E/A) 
    const jornada = fila[COL_MARCA_N_E_A - 1]; 
    
    // --- (FIN DE LA CORRECCIÓN) ---

    if (!grupo || grupo === "") {
      continue; // Ignorar filas sin grupo
    }

    // Inicializar el grupo en el mapa si no existe
    if (!mapaContador[grupo]) {
      mapaContador[grupo] = {
        total: 0,
        normal: 0,
        extendida: 0
      };
    }

    // Contar el total
    mapaContador[grupo].total++;

    // Contar por jornada (según tu lógica)
    const jornadaStr = String(jornada).trim();
    
    if (jornadaStr === "Normal" || jornadaStr === "Normal (Pre-Venta)") {
      mapaContador[grupo].normal++;
    } else if (jornadaStr === "Extendida") {
      mapaContador[grupo].extendida++;
    }
  }

  // 3. Leer los grupos de la hoja Contador
  const ultimaFilaContador = hojaContador.getLastRow();
  if (ultimaFilaContador < 2) {
    Logger.log("No hay grupos para contar en la hoja Contador.");
    return;
  }

  // Col A de "Contador" [cite: 366]
  const rangoGruposContador = hojaContador.getRange(2, COL_GRUPOS_CONTADOR_NOMBRE, ultimaFilaContador - 1, 1); 
  const nombresGruposContador = rangoGruposContador.getValues();

  // 4. Preparar los datos de salida
  const datosSalida = [];

  for (let i = 0; i < nombresGruposContador.length; i++) {
    const nombreGrupo = nombresGruposContador[i][0];
    const counts = mapaContador[nombreGrupo];

    if (counts) {
      // Si encontramos el grupo en nuestro mapa, ponemos los valores
      datosSalida.push([
        counts.total,       // Col B
        counts.normal,      // Col C
        counts.extendida    // Col D
      ]);
    } else {
      // Si el grupo no está en el mapa (0 registros), ponemos 0
      datosSalida.push([0, 0, 0]);
    }
  }

  // 5. Escribir los resultados de una sola vez
  hojaContador.getRange(
    2, // Fila inicial
    COL_GRUPOS_CONTADOR_TOTAL, // Col B (Columna 2) [cite: 371]
    datosSalida.length, // Número de filas
    3 // Número de columnas (Total, Normal, Extendida)
  ).setValues(datosSalida);
}