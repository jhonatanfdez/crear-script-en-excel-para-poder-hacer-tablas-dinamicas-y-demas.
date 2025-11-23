/**
 * SCRIPT: Generación de Gran Total
 * OBJETIVO: Calcular y agregar una fila final de "Total general" que sume los totales de las 3 tablas anteriores.
 * UBICACIÓN: Se coloca 2 filas debajo de la tabla de "Horas No Laborables".
 */
function main(workbook: ExcelScript.Workbook) {
    let sheet = workbook.getWorksheet("Para compartir");
    if (!sheet) return;

    // ==========================================
    // 1. OBTENER RANGOS DE LAS 3 TABLAS
    // ==========================================
    // Usamos getSurroundingRegion para detectar dinámicamente el tamaño de cada tabla
    let rangoProy = sheet.getRange("A6").getSurroundingRegion();
    let rangoAdmin = sheet.getRange("A26").getSurroundingRegion();
    let rangoNL = sheet.getRange("A36").getSurroundingRegion();

    // ==========================================
    // 2. IDENTIFICAR FILAS DE TOTALES INDIVIDUALES
    // ==========================================
    // La fila de totales es la última fila de cada región detectada
    // getLastRow() devuelve un objeto Range que representa esa fila completa
    let filaTotalProy = rangoProy.getLastRow();
    let filaTotalAdmin = rangoAdmin.getLastRow();
    let filaTotalNL = rangoNL.getLastRow();

    // ==========================================
    // 3. DETERMINAR POSICIÓN DEL GRAN TOTAL
    // ==========================================
    // Calculamos dónde debe ir la nueva fila:
    // Índice de inicio de la última tabla + número de filas = índice de la primera fila vacía
    let indexUltimaFilaTablaNL = rangoNL.getRowIndex() + rangoNL.getRowCount();
    
    // Dejamos 1 fila vacía de separación, así que sumamos 1 al índice
    let filaGranTotalIndex = indexUltimaFilaTablaNL + 1;

    // ==========================================
    // 4. ESCRIBIR ETIQUETA
    // ==========================================
    // Escribimos "Total general" en la columna A de la nueva fila
    sheet.getRangeByIndexes(filaGranTotalIndex, 0, 1, 1).setValue("Total general");

    // ==========================================
    // 5. CALCULAR SUMAS POR COLUMNA
    // ==========================================
    // Asumimos que las 3 tablas tienen la misma estructura de columnas (mismos recursos en mismo orden)
    let columnCount = rangoProy.getColumnCount();

    // Iteramos desde la columna 1 (B) hasta la última columna de datos
    // La columna 0 es la de etiquetas ("Total general"), por eso empezamos en 1
    for (let col = 1; col < columnCount; col++) {
        // Obtenemos los valores de las celdas de totales de cada tabla
        // getCell(0, col) obtiene la celda en la fila relativa 0 (la única fila del rango) y columna relativa 'col'
        let valProy = filaTotalProy.getCell(0, col).getValue() as number;
        let valAdmin = filaTotalAdmin.getCell(0, col).getValue() as number;
        let valNL = filaTotalNL.getCell(0, col).getValue() as number;

        // Sumamos los valores, tratando nulos/undefined como 0
        let suma = (valProy || 0) + (valAdmin || 0) + (valNL || 0);

        // Escribimos el resultado en la fila del Gran Total
        sheet.getRangeByIndexes(filaGranTotalIndex, col, 1, 1).setValue(suma);
    }
}
