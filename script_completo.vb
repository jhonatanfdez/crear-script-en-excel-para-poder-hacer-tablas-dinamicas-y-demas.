/**
 * SCRIPT MAESTRO: Ejecución Completa de Reportes
 * OBJETIVO: Ejecutar secuencialmente todos los pasos para generar el reporte de horas.
 * NOTA: Este script consolida la lógica de los 10 scripts individuales en orden cronológico.
 */

function main(workbook: ExcelScript.Workbook) {
    // 0. Limpieza Inicial (Eliminar hojas previas)
    step0_LimpiezaInicial(workbook);

    // 1. Generar reporte de Horas Proyecto
    step1_HorasProyecto(workbook);

    // 2. Generar reporte de Horas Admin
    step2_HorasAdmin(workbook);

    // 3. Generar reporte de Horas No Laborables
    step3_HorasNoLaborables(workbook);

    // 4. Aplicar filtros inteligentes a las tablas dinámicas
    step4_HacerFiltros(workbook);

    // 5. Generar hoja de presentación y títulos
    step5_GenerarTitulo(workbook);

    // 6. Copiar datos de Proyectos a la presentación
    step6_GenerarProyectos(workbook);

    // 7. Copiar datos de Admin a la presentación
    step7_GenerarHorasAdmin(workbook);

    // 8. Copiar datos de No Laborables a la presentación
    step8_GenerarHorasNoLaborables(workbook);

    // 9. Aplicar formatos de bordes y alineación (Parte 1)
    step9_Ajustes1(workbook);

    // 10. Ajustes finales de columnas y filas (Parte 2)
    step10_Ajustes2(workbook);

    // 11. Generar Gran Total (Sumatoria de las 3 tablas)
    step11_GenerarGranTotal(workbook);

    // 12. Aplicar estilos finales (Negrita y Azul)
    step12_EstilosFinales(workbook);
}

// ==========================================================================
// 0. LIMPIEZA INICIAL
// ==========================================================================
/**
 * SCRIPT: Limpieza Inicial
 * OBJETIVO: Eliminar las hojas generadas en ejecuciones anteriores para evitar errores de duplicidad.
 * HOJAS A ELIMINAR: "Para compartir", "Horas Proyectos", "Horas Admin", "Horas No Laborables".
 */
function step0_LimpiezaInicial(workbook: ExcelScript.Workbook) {
    // Lista de nombres de hojas que queremos limpiar
    const hojasAEliminar = [
        "Para compartir", 
        "Horas Proyectos", 
        "Horas Admin", 
        "Horas No Laborables"
    ];

    // Recorremos la lista y eliminamos si existen
    hojasAEliminar.forEach(nombreHoja => {
        let hoja = workbook.getWorksheet(nombreHoja);
        if (hoja) {
            hoja.delete();
            console.log(`Hoja eliminada: ${nombreHoja}`);
        } else {
            console.log(`Hoja no encontrada (ya estaba limpia): ${nombreHoja}`);
        }
    });
}

// ==========================================================================
// 1. HORAS PROYECTO
// ==========================================================================
/**
 * SCRIPT: Creación de Reporte "Horas Proyectos"
 * OBJETIVO: Generar una nueva hoja con una tabla dinámica basada en los datos maestros.
 * FUENTE DE DATOS: Hoja "Datos TM+", Rango A1:AA884.
 * SALIDA: Hoja nueva llamada "Horas Proyectos".
 */
function step1_HorasProyecto(workbook: ExcelScript.Workbook) {

	// ==========================================
	// 1. PREPARACIÓN DE LAS HOJAS
	// ==========================================

	// Crea una nueva hoja en blanco donde se colocará el reporte
	let hojaReporte = workbook.addWorksheet();

	// Obtiene la hoja que contiene la data cruda (Fuente de la verdad)
	let hojaDatosOrigen = workbook.getWorksheet("Datos TM+");

	// ==========================================
	// 2. CREACIÓN DE LA TABLA DINÁMICA
	// ==========================================

	// Se inserta la tabla dinámica llamada "TablaDinámica11".
	// - Origen: Rango dinámico desde A1 de 'Datos TM+'.
	// - Destino: La nueva hoja creada, comenzando en la celda A3.
	let newPivotTable = workbook.addPivotTable(
		"TablaDinámica11",
		hojaDatosOrigen.getRange("A1").getSurroundingRegion(),
		hojaReporte.getRange("A3")
	);

	// ==========================================
	// 3. ORGANIZACIÓN DEL LIBRO
	// ==========================================

	// Se renombra la hoja nueva para estandarizar el nombre del reporte
	hojaReporte.setName("Horas Proyectos");

	// Se actualiza la referencia a la hoja con el nuevo nombre
	let horas_Proyectos = workbook.getWorksheet("Horas Proyectos");

	// Se mueve la hoja a la posición índice 2 (es decir, la 3ra pestaña del libro)
	// Esto ayuda a mantener el orden visual para el usuario final.
	horas_Proyectos.setPosition(2);

	// ==========================================
	// 4. CONFIGURACIÓN DE CAMPOS (FILAS Y COLUMNAS)
	// ==========================================

	// -- Configuración de FILAS --
	// Agrega el campo "Recurso" a las filas para agrupar por persona
	newPivotTable.addRowHierarchy(newPivotTable.getHierarchy("Recurso"));

	// -- Configuración de COLUMNAS --
	// Agrega también "Recurso" a las columnas (Matriz cruzada)
	newPivotTable.addColumnHierarchy(newPivotTable.getHierarchy("Recurso"));

	// Asegura que la jerarquía de columna "Recurso" esté en la posición inicial (índice 0)
	newPivotTable.getColumnHierarchy("Recurso").setPosition(0);

	// -- Nivel Adicional en FILAS --
	// Agrega "Nombre Proyecto" debajo de "Recurso" en las filas.
	// Esto crea un desglose: Recurso -> Proyectos en los que trabajó.
	newPivotTable.addRowHierarchy(newPivotTable.getHierarchy("Nombre Proyecto"));

	// ==========================================
	// 5. CONFIGURACIÓN DE VALORES (DATOS)
	// ==========================================

	// Agrega el campo numérico "Horas del proyecto" al área de valores.
	// Por defecto, esto sumará las horas registradas para cada cruce de Recurso/Proyecto.
	newPivotTable.addDataHierarchy(newPivotTable.getHierarchy("Horas del proyecto"));
}

// ==========================================================================
// 2. HORAS ADMIN
// ==========================================================================
/**
 * SCRIPT: Creación de Reporte "Horas Admin"
 * OBJETIVO: Generar una tabla dinámica para analizar las horas administrativas.
 * FUENTE DE DATOS: Hoja "Datos TM+", Rango A1:AA884.
 * SALIDA: Hoja nueva llamada "Horas Admin".
 */
function step2_HorasAdmin(workbook: ExcelScript.Workbook) {

	// ==========================================
	// 1. PREPARACIÓN DE LA HOJA
	// ==========================================

	// Agrega una nueva hoja de cálculo al libro
	let hoja14 = workbook.addWorksheet();
	
	// Obtiene la hoja de datos origen
	let datos_TM_ = workbook.getWorksheet("Datos TM+");

	// ==========================================
	// 2. CREACIÓN DE LA TABLA DINÁMICA
	// ==========================================

	// Crea la tabla dinámica "TablaDinámica12" en la nueva hoja, celda A3
	let newPivotTable = workbook.addPivotTable(
		"TablaDinámica12", 
		datos_TM_.getRange("A1").getSurroundingRegion(), 
		hoja14.getRange("A3")
	);

	// ==========================================
	// 3. CONFIGURACIÓN DE CAMPOS
	// ==========================================

	// -- Filas --
	// Agrega "Categoría de tiempo" a las filas
	newPivotTable.addRowHierarchy(newPivotTable.getHierarchy("Categoría de tiempo"));

	// -- Valores --
	// Agrega "Horas de admin" a los valores para sumarizar
	newPivotTable.addDataHierarchy(newPivotTable.getHierarchy("Horas de admin"));

	// -- Filas Adicionales --
	// Agrega "Recurso" a las filas
	newPivotTable.addRowHierarchy(newPivotTable.getHierarchy("Recurso"));

	// -- Columnas --
	// Agrega "Recurso" a las columnas para matriz cruzada
	newPivotTable.addColumnHierarchy(newPivotTable.getHierarchy("Recurso"));

	// Asegura que "Recurso" sea la primera jerarquía en columnas
	newPivotTable.getColumnHierarchy("Recurso").setPosition(0);

	// ==========================================
	// 4. FINALIZACIÓN Y ORGANIZACIÓN
	// ==========================================

	// Mueve la hoja a la posición 3 (índice visual)
	hoja14.setPosition(3);

	// Renombra la hoja a "Horas Admin"
	hoja14.setName("Horas Admin");
}

// ==========================================================================
// 3. HORAS NO LABORABLES
// ==========================================================================
/**
 * SCRIPT: Creación de Reporte "Horas No Laborables"
 * OBJETIVO: Generar una tabla dinámica para analizar tiempos no laborables (Vacaciones, Licencias, etc.).
 * FUENTE DE DATOS: Hoja "Datos TM+", Rango A1:AA884.
 * SALIDA: Hoja nueva llamada "Horas No Laborables".
 */
function step3_HorasNoLaborables(workbook: ExcelScript.Workbook) {

	// ==========================================
	// 1. PREPARACIÓN DE LA HOJA
	// ==========================================

	// Agrega una nueva hoja de cálculo
	let hoja19 = workbook.addWorksheet();
	
	// Referencia a la hoja de datos
	let datos_TM_ = workbook.getWorksheet("Datos TM+");

	// ==========================================
	// 2. CREACIÓN DE LA TABLA DINÁMICA
	// ==========================================

	// Crea la tabla dinámica "TablaDinámica14" en la nueva hoja
	let newPivotTable = workbook.addPivotTable(
		"TablaDinámica14", 
		datos_TM_.getRange("A1").getSurroundingRegion(), 
		hoja19.getRange("A3")
	);

	// ==========================================
	// 3. CONFIGURACIÓN DE CAMPOS
	// ==========================================

	// -- Filas --
	// Agrupa por "Categoría de tiempo"
	newPivotTable.addRowHierarchy(newPivotTable.getHierarchy("Categoría de tiempo"));

	// -- Valores --
	// Suma las "Horas no laborables"
	newPivotTable.addDataHierarchy(newPivotTable.getHierarchy("Horas no laborables"));

	// -- Filas Adicionales --
	// Desglose por "Recurso"
	newPivotTable.addRowHierarchy(newPivotTable.getHierarchy("Recurso"));

	// -- Columnas --
	// Matriz por "Recurso"
	newPivotTable.addColumnHierarchy(newPivotTable.getHierarchy("Recurso"));

	// Ajusta la posición de la columna "Recurso"
	newPivotTable.getColumnHierarchy("Recurso").setPosition(0);

	// ==========================================
	// 4. FINALIZACIÓN
	// ==========================================

	// Mueve la hoja a la posición 4
	hoja19.setPosition(4);

	// Renombra la hoja
	hoja19.setName("Horas No Laborables");
}

// ==========================================================================
// 4. HACER FILTROS
// ==========================================================================
/**
 * SCRIPT: Aplicación de Filtros Inteligentes
 * OBJETIVO: Filtrar las tablas dinámicas de "Horas Admin" y "Horas No Laborables" según reglas de negocio.
 * LÓGICA: 
 *   - Admin: Oculta Feriados/Licencias/Blancos.
 *   - No Laborables: Muestra SOLO Feriados/Licencias.
 */
function step4_HacerFiltros(workbook: ExcelScript.Workbook) {
  // =====================================================
  // 1. DEFINICIÓN DE PREFERENCIAS Y REGLAS
  // =====================================================
  const nombreCampo = "Categoría de tiempo";
  // Lista de categorías consideradas "No Laborables"
  const listaClave = ["Feriados", "Licencias", "Permisos", "Vacaciones"];
  const nombreEnBlanco = "(en blanco)";

  // =====================================================
  // 2. PROCESAR HOJA "HORAS ADMIN"
  // =====================================================
  procesarTablaInteligente(workbook, "Horas Admin", nombreCampo, (nombreItem) => {
    // REGLA DE NEGOCIO:
    // - Si es blanco -> Ocultar
    // - Si está en la lista de No Laborables -> Ocultar
    // - Todo lo demás (ej. Reuniones, Entrenamientos) -> Mostrar
    if (nombreItem === nombreEnBlanco) return false;
    if (listaClave.includes(nombreItem)) return false;
    return true; 
  });

  // =====================================================
  // 3. PROCESAR HOJA "HORAS NO LABORABLES"
  // =====================================================
  procesarTablaInteligente(workbook, "Horas No Laborables", nombreCampo, (nombreItem) => {
    // REGLA DE NEGOCIO:
    // - Si es blanco -> Ocultar
    // - Si está en la lista de No Laborables -> Mostrar
    // - Todo lo demás -> Ocultar
    if (nombreItem === nombreEnBlanco) return false;
    if (listaClave.includes(nombreItem)) return true;
    return false; 
  });
}

// =====================================================
// FUNCIÓN AUXILIAR: PROCESAMIENTO ROBUSTO DE TABLAS
// =====================================================
/**
 * Recorre los ítems de una tabla dinámica y aplica visibilidad según una función de regla.
 * Maneja casos especiales como "(en blanco)" que a veces es string vacío.
 */
function procesarTablaInteligente(
  workbook: ExcelScript.Workbook,
  nombreHoja: string,
  nombreCampo: string,
  reglaVisibilidad: (nombre: string) => boolean
) {
  const hoja = workbook.getWorksheet(nombreHoja);
  if (!hoja) return;

  const tabla = hoja.getPivotTables()[0];
  if (!tabla) return;

  console.log(`--- Procesando: ${nombreHoja} ---`);

  // Obtenemos los textos visibles de la tabla para iterar sobre lo que existe
  const rangoTabla = tabla.getLayout().getRange();
  const textos = rangoTabla.getTexts();

  const jerarquia = tabla.getRowHierarchy(nombreCampo);
  if (!jerarquia) {
    console.log("Campo no encontrado");
    return;
  }
  const campo = jerarquia.getFields()[0];

  // Recorremos las filas leídas (empezando en 1 para saltar encabezados)
  for (let i = 1; i < textos.length; i++) {
    let nombreItemDetectado = textos[i][0]; // Texto en la celda

    if (nombreItemDetectado && nombreItemDetectado !== "Total general") {

      // Evaluamos la regla
      const debeVerse = reglaVisibilidad(nombreItemDetectado);

      // === INTENTO 1: Buscar por el nombre exacto ===
      let item = campo.getPivotItem(nombreItemDetectado);

      // === INTENTO 2 (Plan B): Manejo de blancos ===
      if (!item && (nombreItemDetectado === "(en blanco)" || nombreItemDetectado === "(blank)")) {
        item = campo.getPivotItem("");
      }

      // Aplicar visibilidad si se encontró el ítem
      if (item) {
        // Optimización: Solo cambiar si es necesario
        if (item.getVisible() !== debeVerse) {
          item.setVisible(debeVerse);
        }
      }
    }
  }
  // console.log("✅ Proceso completado."); // Comentado para rendimiento
}

// ==========================================================================
// 5. GENERAR TITULO
// ==========================================================================
/**
 * SCRIPT: Generación de Títulos y Hoja "Para compartir"
 * OBJETIVO: Crear la hoja final de presentación y configurar los encabezados institucionales.
 * SALIDA: Hoja nueva llamada "Para compartir" con formato inicial.
 */
function step5_GenerarTitulo(workbook: ExcelScript.Workbook) {

	// ==========================================
	// 1. CREACIÓN DE LA HOJA DE PRESENTACIÓN
	// ==========================================

	// Agrega una nueva hoja llamada "Para compartir"
	let para_compartir = workbook.addWorksheet("Para compartir");

	// ==========================================
	// 2. CONFIGURACIÓN DE ENCABEZADOS
	// ==========================================

	// Establece los textos del encabezado en el rango A1:A4
	// Incluye Nombre de la Empresa, Vicepresidencia, Gerencia y Título del Reporte
	para_compartir.getRange("A1:A4").setValues([
		["NOMBRE DE LA EMPRESA S.A."],
		["Vicepresidencia de Auditoria"],
		["Gerencia de …"],
		["Horas del trimestre mes - mes 2025"]
	]);

	// ==========================================
	// 3. FORMATO VISUAL
	// ==========================================

	// Oculta las líneas de cuadrícula para una apariencia más limpia
	para_compartir.setShowGridlines(false);

	// Aplica negrita a los encabezados (A1:A4)
	para_compartir.getRange("A1:A4").getFormat().getFont().setBold(true);
}

// ==========================================================================
// 6. GENERAR PROYECTOS
// ==========================================================================
/**
 * SCRIPT: Generación de Sección "Proyectos"
 * OBJETIVO: Copiar los datos procesados de la tabla dinámica de Proyectos a la hoja de presentación.
 * FUENTE: Hoja "Horas Proyectos", Rango dinámico desde A3.
 * DESTINO: Hoja "Para compartir", celda A6.
 */
function step6_GenerarProyectos(workbook: ExcelScript.Workbook) {

	// ==========================================
	// 1. REFERENCIAS A HOJAS
	// ==========================================
	let para_compartir = workbook.getWorksheet("Para compartir");
	let horas_Proyectos = workbook.getWorksheet("Horas Proyectos");

	// ==========================================
	// 2. ORGANIZACIÓN
	// ==========================================

	// Mueve la hoja "Para compartir" a la posición 2 (para que sea la principal visible)
	para_compartir.setPosition(2);

	// ==========================================
	// 3. COPIADO DE DATOS
	// ==========================================

	// Identifica el rango completo de la tabla dinámica de forma robusta
	// Usamos el objeto PivotTable para obtener el rango exacto, evitando problemas con celdas vacías intermedias
	let pivotTable = horas_Proyectos.getPivotTables()[0];
	let rangoOrigen = pivotTable.getLayout().getRange();

	// Copia los valores de la tabla dinámica de proyectos
	// Origen: Rango dinámico de "Horas Proyectos"
	// Destino: A6 de "Para compartir"
	// Solo copia valores (sin formato de tabla dinámica)
	para_compartir.getRange("A6").copyFrom(
		rangoOrigen, 
		ExcelScript.RangeCopyType.values, 
		false, 
		false
	);
}

// ==========================================================================
// 7. GENERAR HORAS ADMIN
// ==========================================================================
/**
 * SCRIPT: Generación de Sección "Horas Admin"
 * OBJETIVO: Copiar datos administrativos y aplicar formato de bordes y alineación.
 * FUENTE: Hoja "Horas Admin", Rango dinámico desde A3.
 * DESTINO: Hoja "Para compartir", posición dinámica después de la tabla anterior.
 */
function step7_GenerarHorasAdmin(workbook: ExcelScript.Workbook) {

	// ==========================================
	// 1. REFERENCIAS Y COPIADO
	// ==========================================
	let para_compartir = workbook.getWorksheet("Para compartir");
	let horas_Admin = workbook.getWorksheet("Horas Admin");

	// Identifica el rango completo de la tabla dinámica de forma robusta
	let pivotTable = horas_Admin.getPivotTables()[0];
	let rangoOrigen = pivotTable.getLayout().getRange();

	// --- LÓGICA DE POSICIONAMIENTO DINÁMICO ROBUSTA ---
	// Buscamos la última fila real con datos en la columna A para evitar sobrescribir.
	// getUsedRange(true) ignora celdas que solo tienen formato pero no valores.
	let usedRangeA = para_compartir.getRange("A:A").getUsedRange(true);
	let lastRow = usedRangeA ? usedRangeA.getLastRow().getRowIndex() : 5; // Default seguro si está vacía
	
	// Dejamos 1 fila vacía de separación (lastRow + 2)
	let targetRowIndex = lastRow + 2;
	
	let celdaDestino = para_compartir.getRangeByIndexes(targetRowIndex, 0, 1, 1); // Columna A

	// Copia los datos de horas administrativas a la hoja de presentación
	celdaDestino.copyFrom(
		rangoOrigen, 
		ExcelScript.RangeCopyType.values, 
		false, 
		false
	);

	// ==========================================
	// 2. FORMATO DE BORDES (TABLA COMPLETA)
	// ==========================================
	
	// Usamos las dimensiones del rango origen para asegurar que cubrimos toda la tabla,
	// incluso si hay celdas vacías que romperían un getExtendedRange.
	let rowCount = rangoOrigen.getRowCount();
	let colCount = rangoOrigen.getColumnCount();
	
	let rangoCompleto = celdaDestino.getResizedRange(rowCount - 1, colCount - 1);
	let formatoCompleto = rangoCompleto.getFormat();

	// Limpia diagonales
	formatoCompleto.getRangeBorder(ExcelScript.BorderIndex.diagonalDown).setStyle(ExcelScript.BorderLineStyle.none);
	formatoCompleto.getRangeBorder(ExcelScript.BorderIndex.diagonalUp).setStyle(ExcelScript.BorderLineStyle.none);

	// Aplica bordes externos e internos finos a TODO el rango
	let borderStyle = ExcelScript.BorderLineStyle.continuous;
	let borderWeight = ExcelScript.BorderWeight.thin;

	formatoCompleto.getRangeBorder(ExcelScript.BorderIndex.edgeLeft).setStyle(borderStyle);
	formatoCompleto.getRangeBorder(ExcelScript.BorderIndex.edgeLeft).setWeight(borderWeight);
	formatoCompleto.getRangeBorder(ExcelScript.BorderIndex.edgeTop).setStyle(borderStyle);
	formatoCompleto.getRangeBorder(ExcelScript.BorderIndex.edgeTop).setWeight(borderWeight);
	formatoCompleto.getRangeBorder(ExcelScript.BorderIndex.edgeBottom).setStyle(borderStyle);
	formatoCompleto.getRangeBorder(ExcelScript.BorderIndex.edgeBottom).setWeight(borderWeight);
	formatoCompleto.getRangeBorder(ExcelScript.BorderIndex.edgeRight).setStyle(borderStyle);
	formatoCompleto.getRangeBorder(ExcelScript.BorderIndex.edgeRight).setWeight(borderWeight);
	
	formatoCompleto.getRangeBorder(ExcelScript.BorderIndex.insideVertical).setStyle(borderStyle);
	formatoCompleto.getRangeBorder(ExcelScript.BorderIndex.insideVertical).setWeight(borderWeight);
	formatoCompleto.getRangeBorder(ExcelScript.BorderIndex.insideHorizontal).setStyle(borderStyle);
	formatoCompleto.getRangeBorder(ExcelScript.BorderIndex.insideHorizontal).setWeight(borderWeight);

	// ==========================================
	// 4. ALINEACIÓN Y AJUSTE DE TEXTO
	// ==========================================

	// -- Encabezados de Sección --
	// El encabezado ocupa todo el ancho de la tabla pegada.
	// (Reutilizamos colCount calculado en la sección 2)
	let rangoEncabezadoCompleto = celdaDestino.getResizedRange(0, colCount - 1);
	
	rangoEncabezadoCompleto.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
	rangoEncabezadoCompleto.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.bottom);
	rangoEncabezadoCompleto.getFormat().setWrapText(false);
	rangoEncabezadoCompleto.getFormat().setTextOrientation(0);
	rangoEncabezadoCompleto.merge(false); // Combina celdas para título centrado

	// Reajuste posterior a la izquierda
	rangoEncabezadoCompleto.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);
	rangoEncabezadoCompleto.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.bottom);
	rangoEncabezadoCompleto.merge(false);

	// -- Encabezados Superiores (Proyectos) --
	// NOTA: Este bloque modificaba A6:G6. Como estamos en step7 (Admin), 
	// esto parece código residual o que afecta a la tabla anterior. 
	// Lo dejaremos comentado o adaptado si es necesario, pero A6 es fijo para Proyectos.
	// Si se requiere, se debe manejar en step6 o step9.
}

// ==========================================================================
// 8. GENERAR HORAS NO LABORABLES
// ==========================================================================
/**
 * SCRIPT: Generación de Sección "Horas No Laborables"
 * OBJETIVO: Copiar los datos de tiempos no laborables a la hoja de presentación.
 * FUENTE: Hoja "Horas No Laborables", Rango dinámico desde A3 (incluye totales).
 * DESTINO: Hoja "Para compartir", posición dinámica después de la tabla anterior.
 */
function step8_GenerarHorasNoLaborables(workbook: ExcelScript.Workbook) {

	// ==========================================
	// 1. REFERENCIAS Y COPIADO
	// ==========================================
	let para_compartir = workbook.getWorksheet("Para compartir");
	let horas_No_Laborables = workbook.getWorksheet("Horas No Laborables");

	// Identifica el rango completo de la tabla dinámica de forma robusta
	let pivotTable = horas_No_Laborables.getPivotTables()[0];
	let rangoOrigen = pivotTable.getLayout().getRange();

	// --- LÓGICA DE POSICIONAMIENTO DINÁMICO ROBUSTA ---
	// Buscamos la última fila real con datos en la columna A
	let usedRangeA = para_compartir.getRange("A:A").getUsedRange(true);
	let lastRow = usedRangeA ? usedRangeA.getLastRow().getRowIndex() : 5;

	// Dejamos 1 fila vacía de separación (lastRow + 2)
	let targetRowIndex = lastRow + 2;
	
	let celdaDestino = para_compartir.getRangeByIndexes(targetRowIndex, 0, 1, 1); // Columna A

	// Copia los datos a la hoja de presentación
	celdaDestino.copyFrom(
		rangoOrigen, 
		ExcelScript.RangeCopyType.values, 
		false, 
		false
	);
}

// ==========================================================================
// 9. AJUSTES 1
// ==========================================================================
/**
 * SCRIPT: Ajustes de Formato - Parte 1
 * OBJETIVO: Aplicar bordes, alineación y estilos a las tablas de "Proyectos" y "Horas No Laborables".
 * ALCANCE: Rangos dinámicos detectados por encabezados.
 */
function step9_Ajustes1(workbook: ExcelScript.Workbook) {
	let para_compartir = workbook.getWorksheet("Para compartir");

	// ==========================================
	// 1. FORMATO TABLA PROYECTOS (A6)
	// ==========================================
	// Asumimos que Proyectos siempre empieza en A6
	let rangoProyectos = para_compartir.getRange("A6").getSurroundingRegion();
	let formatoProyectos = rangoProyectos.getFormat();

	// Limpia diagonales
	formatoProyectos.getRangeBorder(ExcelScript.BorderIndex.diagonalDown).setStyle(ExcelScript.BorderLineStyle.none);
	formatoProyectos.getRangeBorder(ExcelScript.BorderIndex.diagonalUp).setStyle(ExcelScript.BorderLineStyle.none);

	// Aplica bordes finos continuos en todos los lados e interiores
	formatoProyectos.getRangeBorder(ExcelScript.BorderIndex.edgeLeft).setStyle(ExcelScript.BorderLineStyle.continuous);
	formatoProyectos.getRangeBorder(ExcelScript.BorderIndex.edgeLeft).setWeight(ExcelScript.BorderWeight.thin);
	formatoProyectos.getRangeBorder(ExcelScript.BorderIndex.edgeTop).setStyle(ExcelScript.BorderLineStyle.continuous);
	formatoProyectos.getRangeBorder(ExcelScript.BorderIndex.edgeTop).setWeight(ExcelScript.BorderWeight.thin);
	formatoProyectos.getRangeBorder(ExcelScript.BorderIndex.edgeBottom).setStyle(ExcelScript.BorderLineStyle.continuous);
	formatoProyectos.getRangeBorder(ExcelScript.BorderIndex.edgeBottom).setWeight(ExcelScript.BorderWeight.thin);
	formatoProyectos.getRangeBorder(ExcelScript.BorderIndex.edgeRight).setStyle(ExcelScript.BorderLineStyle.continuous);
	formatoProyectos.getRangeBorder(ExcelScript.BorderIndex.edgeRight).setWeight(ExcelScript.BorderWeight.thin);
	formatoProyectos.getRangeBorder(ExcelScript.BorderIndex.insideVertical).setStyle(ExcelScript.BorderLineStyle.continuous);
	formatoProyectos.getRangeBorder(ExcelScript.BorderIndex.insideVertical).setWeight(ExcelScript.BorderWeight.thin);
	formatoProyectos.getRangeBorder(ExcelScript.BorderIndex.insideHorizontal).setStyle(ExcelScript.BorderLineStyle.continuous);
	formatoProyectos.getRangeBorder(ExcelScript.BorderIndex.insideHorizontal).setWeight(ExcelScript.BorderWeight.thin);

	// -- Alineación General --
	formatoProyectos.setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
	formatoProyectos.setVerticalAlignment(ExcelScript.VerticalAlignment.bottom);
	formatoProyectos.setWrapText(false);
	formatoProyectos.setTextOrientation(0);

	// -- Encabezado Específico (A6) --
	let rangoA6 = para_compartir.getRange("A6");
	rangoA6.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);
	rangoA6.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.bottom);
	rangoA6.merge(false);

	// -- Columna A (Nombres) --
	// Extiende desde A7 hacia abajo para alinear nombres a la izquierda
	// Usamos el rango detectado para saber hasta dónde llegar
	let rowCount = rangoProyectos.getRowCount();
	// A7 es la segunda fila del rango (index 1)
	if (rowCount > 1) {
		let rangoNombres = rangoProyectos.getRow(1).getResizedRange(rowCount - 2, 0).getColumn(0);
		rangoNombres.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);
		rangoNombres.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.bottom);
	}

	// ==========================================
	// 2. FORMATO TABLA HORAS ADMIN
	// ==========================================
	// Buscamos la tabla de Admin dinámicamente para aplicar alineación
	let foundAdmin = para_compartir.getRange("A:A").find("Suma de Horas de admin", {
		completeMatch: false,
		matchCase: false,
		searchDirection: ExcelScript.SearchDirection.forward
	});

	if (foundAdmin) {
		let rangoAdmin = foundAdmin.getSurroundingRegion();
		
		// 1. Alineación General: Centrada (para que los números queden centrados)
		rangoAdmin.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
		rangoAdmin.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
		rangoAdmin.getFormat().setWrapText(false);

		// 2. Alineación Columna A: Izquierda (para los textos de las filas y el título)
		// Seleccionamos la primera columna del rango detectado
		let colA = rangoAdmin.getColumn(0);
		colA.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);
	}

	// ==========================================
	// 3. FORMATO TABLA NO LABORABLES
	// ==========================================
	// Buscamos la tabla de No Laborables.
	
	// Estrategia: Buscar "Suma de Horas no laborables" en columna A
	let foundRange = para_compartir.getRange("A:A").find("Suma de Horas no laborables", {
		completeMatch: false,
		matchCase: false,
		searchDirection: ExcelScript.SearchDirection.forward
	});

    // Intento alternativo si no se encuentra la cabecera exacta (por si cambia el nombre del campo)
    if (!foundRange) {
        foundRange = para_compartir.getRange("A:A").find("Horas no laborables", {
            completeMatch: false,
            matchCase: false,
            searchDirection: ExcelScript.SearchDirection.forward
        });
    }

	if (foundRange) {
		let rangoNL = foundRange.getSurroundingRegion();
		let formatoCuerpoNL = rangoNL.getFormat();

		// Bordes
		formatoCuerpoNL.getRangeBorder(ExcelScript.BorderIndex.diagonalDown).setStyle(ExcelScript.BorderLineStyle.none);
		formatoCuerpoNL.getRangeBorder(ExcelScript.BorderIndex.diagonalUp).setStyle(ExcelScript.BorderLineStyle.none);
		formatoCuerpoNL.getRangeBorder(ExcelScript.BorderIndex.edgeLeft).setStyle(ExcelScript.BorderLineStyle.continuous);
		formatoCuerpoNL.getRangeBorder(ExcelScript.BorderIndex.edgeLeft).setWeight(ExcelScript.BorderWeight.thin);
		formatoCuerpoNL.getRangeBorder(ExcelScript.BorderIndex.edgeTop).setStyle(ExcelScript.BorderLineStyle.continuous);
		formatoCuerpoNL.getRangeBorder(ExcelScript.BorderIndex.edgeTop).setWeight(ExcelScript.BorderWeight.thin);
		formatoCuerpoNL.getRangeBorder(ExcelScript.BorderIndex.edgeBottom).setStyle(ExcelScript.BorderLineStyle.continuous);
		formatoCuerpoNL.getRangeBorder(ExcelScript.BorderIndex.edgeBottom).setWeight(ExcelScript.BorderWeight.thin);
		formatoCuerpoNL.getRangeBorder(ExcelScript.BorderIndex.edgeRight).setStyle(ExcelScript.BorderLineStyle.continuous);
		formatoCuerpoNL.getRangeBorder(ExcelScript.BorderIndex.edgeRight).setWeight(ExcelScript.BorderWeight.thin);
		formatoCuerpoNL.getRangeBorder(ExcelScript.BorderIndex.insideVertical).setStyle(ExcelScript.BorderLineStyle.continuous);
		formatoCuerpoNL.getRangeBorder(ExcelScript.BorderIndex.insideVertical).setWeight(ExcelScript.BorderWeight.thin);
		formatoCuerpoNL.getRangeBorder(ExcelScript.BorderIndex.insideHorizontal).setStyle(ExcelScript.BorderLineStyle.continuous);
		formatoCuerpoNL.getRangeBorder(ExcelScript.BorderIndex.insideHorizontal).setWeight(ExcelScript.BorderWeight.thin);

		// Alineación
		formatoCuerpoNL.setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
		formatoCuerpoNL.setVerticalAlignment(ExcelScript.VerticalAlignment.center);
		formatoCuerpoNL.setWrapText(false);

		// Alineación Columna A: Izquierda (para etiquetas)
		rangoNL.getColumn(0).getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);

		// Encabezado alineado a la izquierda
		foundRange.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);
	}
}

// ==========================================================================
// 10. AJUSTES 2
// ==========================================================================
/**
 * SCRIPT: Ajustes Finales de Formato (Parte 2)
 * OBJETIVO: Ajustar automáticamente el ancho de columnas y alinear celdas específicas al final de la hoja.
 * ALCANCE: Hoja "Para compartir", ajuste global y rangos dinámicos al final.
 */
function step10_Ajustes2(workbook: ExcelScript.Workbook) {
	let para_compartir = workbook.getWorksheet("Para compartir");

	// ==========================================
	// 1. AJUSTE AUTOMÁTICO DE COLUMNAS
	// ==========================================
	// Ajusta el ancho de todas las columnas según el contenido
	para_compartir.getRange().getFormat().autofitColumns();

	// ==========================================
	// 2. ALINEACIÓN DE CELDAS ESPECÍFICAS (Navegación Dinámica)
	// ==========================================
	// Navega desde A1 hacia abajo múltiples veces para encontrar la sección final (probablemente totales o notas al pie)
	let rangoFinal = para_compartir.getRange("A1")
		.getRangeEdge(ExcelScript.KeyboardDirection.down)
		.getRangeEdge(ExcelScript.KeyboardDirection.down)
		.getRangeEdge(ExcelScript.KeyboardDirection.down)
		.getRangeEdge(ExcelScript.KeyboardDirection.down)
		.getRangeEdge(ExcelScript.KeyboardDirection.down)
		.getRangeEdge(ExcelScript.KeyboardDirection.down);

	// Formato para este rango encontrado dinámicamente
	let formatoFinal = rangoFinal.getFormat();
	formatoFinal.setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);
	formatoFinal.setVerticalAlignment(ExcelScript.VerticalAlignment.center);
	formatoFinal.setIndentLevel(0);
	formatoFinal.setWrapText(false);
	formatoFinal.setTextOrientation(0);
	
	// Fusionar celdas en el rango encontrado
	rangoFinal.merge(false);

	// ==========================================
	// 3. FORMATO RANGO FINAL (A37:A41)
	// ==========================================
	let rangoCierre = para_compartir.getRange("A37:A41");
	let formatoCierre = rangoCierre.getFormat();
	
	formatoCierre.setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);
	formatoCierre.setVerticalAlignment(ExcelScript.VerticalAlignment.center);
	formatoCierre.setIndentLevel(0);
    formatoCierre.setWrapText(false);
    formatoCierre.setTextOrientation(0);
}

// ==========================================================================
// 11. GENERAR GRAN TOTAL
// ==========================================================================
/**
 * SCRIPT: Generación de Gran Total
 * OBJETIVO: Calcular y agregar una fila final de "Total general" que sume los totales de las 3 tablas anteriores.
 * UBICACIÓN: Se coloca 2 filas debajo de la tabla de "Horas No Laborables".
 */
function step11_GenerarGranTotal(workbook: ExcelScript.Workbook) {
    let sheet = workbook.getWorksheet("Para compartir");
    if (!sheet) return;

    // ==========================================
    // 1. OBTENER RANGOS DE LAS 3 TABLAS (DINÁMICO)
    // ==========================================
    // Buscamos las tablas por sus encabezados
    let rangoProy = sheet.getRange("A6").getSurroundingRegion();
    
    // Buscar Admin
    let foundAdmin = sheet.getRange("A:A").find("Suma de Horas de admin", {
        completeMatch: false, matchCase: false, searchDirection: ExcelScript.SearchDirection.forward
    });
    let rangoAdmin = foundAdmin ? foundAdmin.getSurroundingRegion() : null;

    // Buscar No Laborables
    let foundNL = sheet.getRange("A:A").find("Suma de Horas no laborables", {
        completeMatch: false, matchCase: false, searchDirection: ExcelScript.SearchDirection.forward
    });

    if (!foundNL) {
        foundNL = sheet.getRange("A:A").find("Horas no laborables", {
            completeMatch: false, matchCase: false, searchDirection: ExcelScript.SearchDirection.forward
        });
    }

    let rangoNL = foundNL ? foundNL.getSurroundingRegion() : null;

    if (!rangoAdmin || !rangoNL) {
        console.log("No se encontraron todas las tablas para el Gran Total");
        return;
    }

    // ==========================================
    // 2. IDENTIFICAR FILAS DE TOTALES INDIVIDUALES
    // ==========================================
    // La fila de totales es la última fila de cada región detectada
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
    // 5. CALCULAR SUMAS POR COLUMNA (OPTIMIZADO)
    // ==========================================
    // Asumimos que las 3 tablas tienen la misma estructura de columnas (mismos recursos en mismo orden)
    let columnCount = rangoProy.getColumnCount();

    // Leemos los valores de una sola vez para evitar llamadas en bucle (Optimización de rendimiento)
    // getValues() devuelve una matriz 2D [fila][columna]. Como es una sola fila, usamos [0].
    let valoresProy = filaTotalProy.getValues()[0];
    let valoresAdmin = filaTotalAdmin.getValues()[0];
    let valoresNL = filaTotalNL.getValues()[0];
    
    // Preparamos una matriz para escribir los resultados de una sola vez
    // Inicializamos con la etiqueta "Total general" para la columna 0
    let valoresGranTotal: (string | number | boolean)[] = ["Total general"];

    // Iteramos desde la columna 1 (B) hasta la última columna de datos
    for (let col = 1; col < columnCount; col++) {
        // Obtenemos los valores de los arrays leídos previamente
        let valProy = valoresProy[col] as number;
        let valAdmin = valoresAdmin[col] as number;
        let valNL = valoresNL[col] as number;

        // Sumamos los valores, tratando nulos/undefined como 0
        let suma = (valProy || 0) + (valAdmin || 0) + (valNL || 0);
        
        valoresGranTotal.push(suma);
    }
    
    // Escribimos toda la fila de una vez
    // El rango destino empieza en filaGranTotalIndex, columna 0, 1 fila, columnCount columnas
    sheet.getRangeByIndexes(filaGranTotalIndex, 0, 1, columnCount).setValues([valoresGranTotal]);
}

// ==========================================================================
// 12. ESTILOS FINALES (AZUL Y NEGRITA)
// ==========================================================================
/**
 * SCRIPT: Aplicar Estilos de Encabezados y Totales
 * OBJETIVO: Aplicar negrita y color de fondo azul (#C0E6F5) a los encabezados y filas de totales.
 * ALCANCE: Hoja "Para compartir", rangos de encabezados y filas finales de tablas.
 */
function step12_EstilosFinales(workbook: ExcelScript.Workbook) {
	// ==========================================
	// 1. REFERENCIA A LA HOJA
	// ==========================================
	// Intenta obtener la hoja específica "Para compartir", si no, usa la activa
	let selectedSheet = workbook.getWorksheet("Para compartir");
	if (!selectedSheet) {
		selectedSheet = workbook.getActiveWorksheet();
	}

	// Color azul claro corporativo/reporte
	const colorAzul = "#C0E6F5";

	// ==========================================
	// 2. FORMATO DE ENCABEZADOS Y TOTALES (DINÁMICO)
	// ==========================================
	
	// --- Función auxiliar para aplicar estilos a una tabla dada su cabecera ---
	function aplicarEstilosTabla(textoCabecera: string) {
		let found = selectedSheet.getRange("A:A").find(textoCabecera, {
			completeMatch: false, matchCase: false, searchDirection: ExcelScript.SearchDirection.forward
		});

		// Fallback: Si no encuentra "Suma de X", busca solo "X"
		if (!found && textoCabecera.indexOf("Suma de ") === 0) {
			let textoCorto = textoCabecera.replace("Suma de ", "");
			found = selectedSheet.getRange("A:A").find(textoCorto, {
				completeMatch: false, matchCase: false, searchDirection: ExcelScript.SearchDirection.forward
			});
		}
		
		if (found) {
			// ESTRATEGIA ROBUSTA: Usar la fila de encabezados de columna (fila siguiente)
			// para determinar el tamaño real de la tabla, ya que la fila de título puede tener huecos.
			let filaEncabezados = found.getOffsetRange(1, 0); // Baja 1 fila
			let regionDatos = filaEncabezados.getSurroundingRegion();
			let numColumnas = regionDatos.getColumnCount();
			
			// 1. UNIR Y ETIQUETAR ENCABEZADO SUPERIOR (Columnas B hasta el final)
			// Rango objetivo: Desde B(fila título) hasta la última columna
			// getResizedRange(0, numColumnas - 2) expande desde la columna B
			let rangoEncabezadoSuperior = found.getOffsetRange(0, 1).getResizedRange(0, numColumnas - 2);
			
			// Limpiamos merges previos para evitar conflictos y unimos
			rangoEncabezadoSuperior.unmerge();
			rangoEncabezadoSuperior.getCell(0, 0).setValue("Etiquetas de columna");
			rangoEncabezadoSuperior.merge(false);
			
			// Alineación centrada
			rangoEncabezadoSuperior.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
			rangoEncabezadoSuperior.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);

			// 2. ESTILOS DE COLOR Y FUENTE
			// Aplicamos a las dos filas superiores (Título y Encabezados)
			// Usamos resize desde found para cubrir todo el ancho
			let rangoTitulo = found.getResizedRange(0, numColumnas - 1); // Fila 1 (A..Final)
			let rangoSubtitulos = filaEncabezados.getResizedRange(0, numColumnas - 1); // Fila 2 (A..Final)
			
			rangoTitulo.getFormat().getFont().setBold(true);
			rangoTitulo.getFormat().getFill().setColor(colorAzul);
			
			rangoSubtitulos.getFormat().getFont().setBold(true);
			rangoSubtitulos.getFormat().getFill().setColor(colorAzul);

			// 3. ESTILOS DE TOTALES (Última fila de la región detectada)
			let rangoTotal = regionDatos.getLastRow();
			rangoTotal.getFormat().getFill().setColor(colorAzul);
			rangoTotal.getFormat().getFont().setBold(true);
		}
	}

	// Aplicar a las 3 tablas
	aplicarEstilosTabla("Suma de Horas del proyecto");
	aplicarEstilosTabla("Suma de Horas de admin");
	aplicarEstilosTabla("Suma de Horas no laborables");

    // ==========================================
	// 4. FORMATO GRAN TOTAL (NUEVO)
	// ==========================================
	// Buscamos la fila del "Total general" final.
	// Sabemos que está después de la tabla de No Laborables.
	// Una forma segura es buscar la última celda usada en la columna A.
	
	let ultimaCeldaA = selectedSheet.getRange("A:A").getUsedRange().getLastRow().getCell(0, 0);
	
	// Verificamos si es "Total general" para estar seguros
	if (ultimaCeldaA.getValue() === "Total general") {
		let rangoGranTotal = ultimaCeldaA.getExtendedRange(ExcelScript.KeyboardDirection.right);
		
		// 1. Color y Negrita
		rangoGranTotal.getFormat().getFill().setColor(colorAzul);
		rangoGranTotal.getFormat().getFont().setBold(true);
		
		// 2. Alineación Centrada
		rangoGranTotal.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
		rangoGranTotal.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);

		// Alineación Izquierda para la etiqueta "Total general"
		rangoGranTotal.getCell(0, 0).getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);

		// 3. Bordes Completos (Caja y divisiones internas)
		let formatoGT = rangoGranTotal.getFormat();
		
		// Bordes externos
		formatoGT.getRangeBorder(ExcelScript.BorderIndex.edgeTop).setStyle(ExcelScript.BorderLineStyle.continuous);
		formatoGT.getRangeBorder(ExcelScript.BorderIndex.edgeBottom).setStyle(ExcelScript.BorderLineStyle.continuous);
		formatoGT.getRangeBorder(ExcelScript.BorderIndex.edgeLeft).setStyle(ExcelScript.BorderLineStyle.continuous);
		formatoGT.getRangeBorder(ExcelScript.BorderIndex.edgeRight).setStyle(ExcelScript.BorderLineStyle.continuous);
		
		// Bordes internos verticales (para separar columnas)
		formatoGT.getRangeBorder(ExcelScript.BorderIndex.insideVertical).setStyle(ExcelScript.BorderLineStyle.continuous);
	}
}