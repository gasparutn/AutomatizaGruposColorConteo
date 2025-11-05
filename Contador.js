/**
 * Función de ayuda para ordenar "Grupo 2 años" antes que "Grupo 10 años"
 */
function sortGrupos(a, b) {
  const numA = parseInt(a.match(/\d+/));
  const numB = parseInt(b.match(/\d+/));
  
  if (numA && numB) {
    return numA - numB;
  }
  // Si no puede extraer un número (ej. "Fuera de rango"), usa orden alfabético
  return a.localeCompare(b);
}

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
  
  // Limpiar el contador si no hay registros
  if (ultimaFilaRegistros < 2) {
    Logger.log("No hay datos en Registros.");
    hojaContador.getRange("A2:D" + hojaContador.getMaxRows()).clearContent();
    return; // No hay datos, limpia el contador
  }

  // 1. Leer datos Y colores de la hoja Registros
  // Leemos desde la Col 1 (A) hasta la Col 9 (I - GRUPOS)
  const rangoDatosRegistros = hojaRegistro.getRange(2, 1, ultimaFilaRegistros - 1, COL_GRUPOS); 
  const valoresRegistros = rangoDatosRegistros.getValues();
  // Leemos los colores de fondo del mismo rango
  const coloresRegistros = rangoDatosRegistros.getBackgrounds();

  // 2. Procesar los datos y contarlos en un mapa
  const mapaContador = {};

  for (let i = 0; i < valoresRegistros.length; i++) {
    const fila = valoresRegistros[i];
    
    // Usamos las constantes
    const grupo = fila[COL_GRUPOS - 1]; // Columna I
    const jornada = fila[COL_MARCA_N_E_A - 1]; // Columna C
    
    // --- ESTA ES TU PETICIÓN ---
    // Guardamos el color de fondo de la celda del grupo
    const color = coloresRegistros[i][COL_GRUPOS - 1]; 
    // ---------------------------

    if (!grupo || grupo === "") {
      continue; // Ignorar filas sin grupo
    }

    // Inicializar el grupo en el mapa si no existe
    if (!mapaContador[grupo]) {
      mapaContador[grupo] = {
        total: 0,
        normal: 0,
        extendida: 0,
        color: color // Guardamos el color de la primera aparición
      };
    }

    // Contar el total
    mapaContador[grupo].total++;

    // Contar jornadas (lógica corregida)
    const jornadaStr = String(jornada).trim().toLowerCase();
    
    if (jornadaStr.includes("extendida")) {
      // Contará "Extendida" y "Extendida (Pre-Venta)"
      mapaContador[grupo].extendida++;
    } else if (jornadaStr.includes("normal")) {
      // Contará "Normal" y "Normal (Pre-Venta)"
      mapaContador[grupo].normal++;
    }
  }

  // 3. Preparar los datos de salida (ordenados)
  const gruposEncontrados = Object.keys(mapaContador).sort(sortGrupos);
  
  const datosSalida = [];
  const coloresSalida = [];

  for (const nombreGrupo of gruposEncontrados) {
    const counts = mapaContador[nombreGrupo];
    
    // Añadimos la fila de datos (A, B, C, D)
    datosSalida.push([
      nombreGrupo,      // Col A (El nombre del grupo)
      counts.total,       // Col B
      counts.normal,      // Col C
      counts.extendida    // Col D
    ]);
    
    // --- ESTA ES TU PETICIÓN ---
    // Añadimos la fila de colores para A, B, C, D
    coloresSalida.push([
      counts.color, // Col A
      counts.color, // Col B
      counts.color, // Col C
      counts.color  // Col D
    ]);
    // ---------------------------
  }

  // 4. Limpiar el contenido anterior (A2:D...)
  // Esto NO tocará tus notas en la columna E o más allá.
  const rangoLleno = hojaContador.getLastRow();
  if (rangoLleno > 1) {
    // Limpia A, B, C, D desde la fila 2 hasta el final
    hojaContador.getRange(2, 1, rangoLleno - 1, 4).clearContent(); 
  }

  // 5. Escribir los nuevos valores Y colores de una sola vez
  if (datosSalida.length > 0) {
    // Escribir valores en A2:D...
    hojaContador.getRange(2, 1, datosSalida.length, 4)
                .setValues(datosSalida);
                
    // Escribir colores en A2:D...
    hojaContador.getRange(2, 1, coloresSalida.length, 4)
                .setBackgrounds(coloresSalida);
  }
}










/*
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

*/

