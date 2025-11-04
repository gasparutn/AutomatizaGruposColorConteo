function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Escuela Hípico')
    .addSubMenu(ui.createMenu('🧰Utililidad')
      .addItem('Eliminar Espaciados en Datos', 'limpiarColumnasPorHoja')
      .addItem('Actualizar Grupos y Colores', 'actualizarGruposManual')
      .addItem('Actualizar Contador (Normal/Extendida)', 'actualizarContadorGruposCompleto')) // <-- ESTA ES LA LÍNEA NUEVA
    .addToUi();
}
/*
function onEdit(e) {
  // 'e' es el objeto de evento
  if (!e || !e.range) {
    return; // No es un evento de edición válido o no hay rango
  }

  const hojaEditada = e.source.getActiveSheet();
  const nombreHojaEditada = hojaEditada.getName();

  // 1. Comprobar si la hoja editada es "Registros"
  // (Usamos la constante que ya tienes)
  if (nombreHojaEditada === NOMBRE_HOJA_REGISTRO) {
    
    const columnaEditada = e.range.getColumn();
    
    // 2. Comprobar si la columna editada es "GRUPOS"
    // (Usamos la constante que ya tienes)
    if (columnaEditada === COL_GRUPOS) {
      
      // 3. Si todo coincide, llamamos a la función de conteo
      actualizarContadorDeGruposConColor();
    }
  }
}
*/
function onEdit(e) {
  try {
    if (!e) return; // Salir si no hay evento (ej. al ejecutar desde el editor)

    const ss = e.source;
    const hojaActiva = ss.getActiveSheet();
    const celdaActiva = e.range;

    // Si la edición fue en la hoja "Registros"
    // Y fue en la columna de Grupos (I) o Jornada (C)
    if (hojaActiva.getName() === NOMBRE_HOJA_REGISTRO &&
        (celdaActiva.getColumn() === COL_GRUPOS || celdaActiva.getColumn() === COL_MARCA_N_E_A)) { // [cite: 304, 310]

      // Llamar a la función principal de conteo
      actualizarContadorGruposCompleto();
    }
  } catch (err) {
    Logger.log(`Error en onEdit: ${err.message}`);
  }
}


function actualizarGruposManual() {
  const ui = SpreadsheetApp.getUi();
  try {
    // Estas constantes deben estar definidas en tu proyecto (ej. en Constantes.js)
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hojaConfig = ss.getSheetByName(NOMBRE_HOJA_CONFIG);
    const hojaRegistro = ss.getSheetByName(NOMBRE_HOJA_REGISTRO);

    if (!hojaConfig) {
      ui.alert('Error: No se encuentra la hoja "Config".');
      return;
    }
    if (!hojaRegistro) {
      ui.alert('Error: No se encuentra la hoja "Registros".');
      return;
    }

    // 1. Crear un mapa de colores desde la hoja "Config"
    // (MODIFICADO) Rango extendido a 14 filas (A30:B43) para cubrir 2 a 15 años
    const rangoGruposConfig = hojaConfig.getRange("A30:B43");
    const valoresGruposConfig = rangoGruposConfig.getValues();
    const coloresGruposConfig = rangoGruposConfig.getBackgrounds();
    
    const mapaColores = {};
    for (let i = 0; i < valoresGruposConfig.length; i++) {
      const nombreGrupo = valoresGruposConfig[i][0]; // Col A: "GRUPO 2 AÑOS"
      const color = coloresGruposConfig[i][1];       // Col B: Color de fondo
      if (nombreGrupo && nombreGrupo.trim() !== "") {
        mapaColores[nombreGrupo.trim().toUpperCase()] = color;
      }
    }

    // 2. Leer todos los datos de la hoja "Registros"
    const ultimaFila = hojaRegistro.getLastRow();
    if (ultimaFila <= 1) {
       ui.alert('No hay datos para actualizar en "Registros".');
       return;
    }
    
    // Leer la columna H (Fecha de Nacimiento)
    const rangoFechas = hojaRegistro.getRange(2, COL_FECHA_NACIMIENTO_REGISTRO, ultimaFila - 1, 1); 
    const valoresFechas = rangoFechas.getValues();
    
    const nuevosValoresGrupo = [];
    const nuevosColoresGrupo = [];

    // 3. Procesar cada fila en memoria (muy rápido)
    for (let i = 0; i < valoresFechas.length; i++) {
      const fechaNacObj = valoresFechas[i][0]; // Es un objeto Date de la planilla
      let textoGrupo = "Sin Fecha";
      let colorGrupo = "#ffffff"; // Blanco por defecto

      if (fechaNacObj && fechaNacObj instanceof Date) {
        try {
          // Convertir el objeto Date de la planilla a un string YYYY-MM-DD
          // Se usa "GMT" para evitar corrimientos de zona horaria al convertir
          const fechaNacStr = Utilities.formatDate(fechaNacObj, "GMT", "yyyy-MM-dd");
          
          // Usar la nueva función de lógica de fechas
          textoGrupo = obtenerGrupoPorFechaNacimiento(fechaNacStr);

          // Buscar el color en el mapa
          const claveMapa = textoGrupo.trim().toUpperCase();
          if (mapaColores[claveMapa]) {
            colorGrupo = mapaColores[claveMapa];
          } else {
            // Si el grupo no está en el mapa (ej. "Fuera de rango"), dejarlo blanco
            colorGrupo = "#ffffff";
          }
        } catch (e) {
           Logger.log("Error procesando fecha en fila " + (i+2) + ": " + e.message);
           textoGrupo = "Error Fecha";
           colorGrupo = "#ffffff";
        }
      }
      
      nuevosValoresGrupo.push([textoGrupo]);
      nuevosColoresGrupo.push([colorGrupo]);
    }

    // 4. Escribir los datos de vuelta en la hoja (Optimizado)
    
    // Escribir los textos en la Columna I (GRUPOS)
    hojaRegistro.getRange(2, COL_GRUPOS, nuevosValoresGrupo.length, 1).setValues(nuevosValoresGrupo);
    
    // Escribir los colores en la Columna I (GRUPOS)
    hojaRegistro.getRange(2, COL_GRUPOS, nuevosColoresGrupo.length, 1).setBackgrounds(nuevosColoresGrupo);

    // --- ¡¡AQUÍ ESTÁ LA INTEGRACIÓN!! ---
    // 5. Llamar a la función para actualizar el contador después de actualizar los grupos.
    actualizarContadorDeGruposConColor();
    // ------------------------------------

   // ui.alert('¡Proceso completado! Se actualizaron ' + nuevosValoresGrupo.length + ' filas con la nueva lógica de grupos.');
  
  } catch (e) {
    Logger.log("Error en actualizarGruposManual: " + e.message);
    ui.alert("Ocurrió un error: " + e.message);
  }
}


function obtenerGrupoPorFechaNacimiento(fechaNacStr) {
  if (!fechaNacStr) return "Sin Fecha";

  try {
    // Normalizar la fecha. "2020-10-15" se interpreta como "2020-10-15T00:00:00Z" (UTC)
    const fechaNac = new Date(fechaNacStr + "T00:00:00Z"); 

    // Definir las fechas de corte (Formato: YYYY, Mes (0-11), Dia)
    // Mes 6 = Julio
    // La lógica se basa en la fórmula del usuario: (>= 1-Jul-AAAA) y (< 1-Jul-AAAA+1)
    
    if (fechaNac >= new Date(Date.UTC(2022, 6, 1)) && fechaNac < new Date(Date.UTC(2023, 6, 1))) return "Grupo 2 años";
    if (fechaNac >= new Date(Date.UTC(2021, 6, 1)) && fechaNac < new Date(Date.UTC(2022, 6, 1))) return "Grupo 3 años";
    if (fechaNac >= new Date(Date.UTC(2020, 6, 1)) && fechaNac < new Date(Date.UTC(2021, 6, 1))) return "Grupo 4 años";
    if (fechaNac >= new Date(Date.UTC(2019, 6, 1)) && fechaNac < new Date(Date.UTC(2020, 6, 1))) return "Grupo 5 años";
    if (fechaNac >= new Date(Date.UTC(2018, 6, 1)) && fechaNac < new Date(Date.UTC(2019, 6, 1))) return "Grupo 6 años";
    if (fechaNac >= new Date(Date.UTC(2017, 6, 1)) && fechaNac < new Date(Date.UTC(2018, 6, 1))) return "Grupo 7 años";
    if (fechaNac >= new Date(Date.UTC(2016, 6, 1)) && fechaNac < new Date(Date.UTC(2017, 6, 1))) return "Grupo 8 años";
    if (fechaNac >= new Date(Date.UTC(2015, 6, 1)) && fechaNac < new Date(Date.UTC(2016, 6, 1))) return "Grupo 9 años";
    if (fechaNac >= new Date(Date.UTC(2014, 6, 1)) && fechaNac < new Date(Date.UTC(2015, 6, 1))) return "Grupo 10 años";
    if (fechaNac >= new Date(Date.UTC(2013, 6, 1)) && fechaNac < new Date(Date.UTC(2014, 6, 1))) return "Grupo 11 años";
    if (fechaNac >= new Date(Date.UTC(2012, 6, 1)) && fechaNac < new Date(Date.UTC(2013, 6, 1))) return "Grupo 12 años";
    
    // (Extendido a 15)
    if (fechaNac >= new Date(Date.UTC(2011, 6, 1)) && fechaNac < new Date(Date.UTC(2012, 6, 1))) return "Grupo 13 años";
    if (fechaNac >= new Date(Date.UTC(2010, 6, 1)) && fechaNac < new Date(Date.UTC(2011, 6, 1))) return "Grupo 14 años";
    if (fechaNac >= new Date(Date.UTC(2009, 6, 1)) && fechaNac < new Date(Date.UTC(2010, 6, 1))) return "Grupo 15 años";

    return "Fuera de rango"; // Default

  } catch (e) {
    Logger.log("Error al parsear fecha en obtenerGrupoPorFechaNacimiento: " + fechaNacStr + " | Error: " + e.message);
    return "Error Fecha";
  }
}


function actualizarContadorDeGruposConColor() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const hojaRegistros = ss.getSheetByName(NOMBRE_HOJA_REGISTRO);
  const hojaContador = ss.getSheetByName(NOMBRE_HOJA_CONTADOR);

  if (!hojaRegistros) {
    Logger.log(`Error: No se encontró la hoja '${NOMBRE_HOJA_REGISTRO}'.`);
    return;
  }
  if (!hojaContador) {
    Logger.log(`Error: No se encontró la hoja '${NOMBRE_HOJA_CONTADOR}'.`);
    return;
  }

  // 1. OBTENER DATOS Y COLORES
  const ultimaFila = hojaRegistros.getLastRow();
  if (ultimaFila < 2) {
    Logger.log("No hay datos para contar en 'Registros'.");
    return;
  }
  
  const rangoDatos = hojaRegistros.getRange(2, COL_GRUPOS, ultimaFila - 1, 1);
  const valores = rangoDatos.getValues();
  const colores = rangoDatos.getBackgrounds();

  // 2. CONTAR VALORES Y GUARDAR EL PRIMER COLOR
  const grupoInfo = {}; 
  for (let i = 0; i < valores.length; i++) {
    const grupo = valores[i][0];
    const color = colores[i][0];
    
    if (grupo && grupo.toString().trim() !== "") { 
      if (!grupoInfo[grupo]) {
        grupoInfo[grupo] = {
          count: 1,
          color: color 
        };
      } else {
        grupoInfo[grupo].count++;
      }
    }
  }

  // 3. PREPARAR LOS DATOS DE SALIDA
  const arrayResultados = [["Grupo", "Total"]]; 
  for (const grupoNombre in grupoInfo) {
    arrayResultados.push([grupoNombre, grupoInfo[grupoNombre].count]);
  }

  // --- ¡¡CAMBIO IMPORTANTE!! ---
  // 4. LIMPIAR CONTENIDO (PRESERVANDO FORMATO) Y APLICAR DATOS
  
  // Borra solo el contenido de las columnas A y B, dejando el formato
  hojaContador.getRange(1, 1, hojaContador.getMaxRows(), 2).clearContent();
  
  if (arrayResultados.length > 1) { 
    // Escribir todos los valores de una vez
    hojaContador.getRange(1, COL_GRUPOS_CONTADOR_NOMBRE, arrayResultados.length, 2)
                .setValues(arrayResultados);

    // Aplicar los colores de fondo (esto sí es necesario)
    for (let i = 1; i < arrayResultados.length; i++) { 
      const grupoNombre = arrayResultados[i][0];
      const colorParaAplicar = grupoInfo[grupoNombre].color;
      
      hojaContador.getRange(i + 1, COL_GRUPOS_CONTADOR_NOMBRE, 1, 2)
                  .setBackground(colorParaAplicar);
    }
    
    // --- ¡LÍNEAS ELIMINADAS! ---
    // Ya no se auto-ajustan las columnas ni se pone la negrita.
    // El formato que pongas manualmente se respetará.
  }

  Logger.log("Contador de grupos con color actualizado (formato manual preservado).");
}