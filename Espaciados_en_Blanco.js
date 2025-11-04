function limpiarColumnasPorHoja() {
  // Define las hojas y sus columnas objetivo (A=1, B=2, etc.)
  const configuracion = {
    'PRE-VENTA': [4, 5],        // Por ejemplo: columnas B y D
    'Base de Datos': [3, 4], // Por ejemplo: columnas A, C y E
    'Registros': [5, 6,7] // Por ejemplo: columnas A, C y E
  };

  const libro = SpreadsheetApp.getActiveSpreadsheet();

  for (const nombreHoja in configuracion) {
    const hoja = libro.getSheetByName(nombreHoja);
    if (!hoja) {
      Logger.log(`La hoja "${nombreHoja}" no existe.`);
      continue;
    }

    const columnasObjetivo = configuracion[nombreHoja];
    const rango = hoja.getDataRange();
    const valores = rango.getValues();

    for (let i = 0; i < valores.length; i++) {
      for (let c = 0; c < columnasObjetivo.length; c++) {
        const colIndex = columnasObjetivo[c] - 1;
        let celda = valores[i][colIndex];

        if (typeof celda === 'string') {
          // Elimina espacios al inicio/final y comas/puntos finales
          celda = celda.trim().replace(/[\s]*[.,]$/, '');
          valores[i][colIndex] = celda;
        }
      }
    }

    rango.setValues(valores);
    Logger.log(`Hoja "${nombreHoja}" limpiada correctamente.`);
  }
  // Mostrar ventana emergente al finalizar
  SpreadsheetApp.getUi().alert(
    '✅ Limpieza completada',
    'Limpieza de espaciados de inicio, fin y caracteres (,) limpiados satisfactoriamente.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}



