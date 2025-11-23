/**
 * SCRIPT MAESTRO: Ejecución Completa de Reportes
 * OBJETIVO: Ejecutar secuencialmente todos los pasos para generar el reporte de horas.
 * NOTA: Este script consolida la lógica de los 10 scripts individuales en orden cronológico.
 */

function main(workbook: ExcelScript.Workbook) {
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

    // 11. Aplicar estilos finales (Negrita y Azul)
    step11_EstilosFinales(workbook);
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
	// - Origen: Hoja 'Datos TM+', rango fijo de A1 a AA884.
	// - Destino: La nueva hoja creada, comenzando en la celda A3.
	let newPivotTable = workbook.addPivotTable(
		"TablaDinámica11",
		hojaDatosOrigen.getRange("A1:AA884"),
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
		datos_TM_.getRange("A1:AA884"), 
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
		datos_TM_.getRange("A1:AA884"), 
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
        try {
          // Optimización: Solo cambiar si es necesario
          if (item.getVisible() !== debeVerse) {
            item.setVisible(debeVerse);
            console.log(`   Ajustado: "${nombreItemDetectado}" -> ${debeVerse}`);
          }
        } catch (e) {
          // Fallback en caso de error de lectura de estado
          item.setVisible(debeVerse);
        }
      }
    }
  }
  console.log("✅ Proceso completado.");
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
	// Incluye Nombre del Banco, Vicepresidencia, Gerencia y Título del Reporte
	para_compartir.getRange("A1:A4").setValues([
		["BANCO MULTIPLE ADEMI S.A."],
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

	// Identifica el rango completo de la tabla dinámica automáticamente
	let rangoOrigen = horas_Proyectos.getRange("A3").getSurroundingRegion();

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
 * DESTINO: Hoja "Para compartir", celda A26.
 */
function step7_GenerarHorasAdmin(workbook: ExcelScript.Workbook) {

	// ==========================================
	// 1. REFERENCIAS Y COPIADO
	// ==========================================
	let para_compartir = workbook.getWorksheet("Para compartir");
	let horas_Admin = workbook.getWorksheet("Horas Admin");

	// Identifica el rango completo de la tabla dinámica automáticamente
	let rangoOrigen = horas_Admin.getRange("A3").getSurroundingRegion();

	// Copia los datos de horas administrativas a la hoja de presentación
	para_compartir.getRange("A26").copyFrom(
		rangoOrigen, 
		ExcelScript.RangeCopyType.values, 
		false, 
		false
	);

	// ==========================================
	// 2. FORMATO DE BORDES (TABLA SUPERIOR)
	// ==========================================
	
	// Configura los bordes para la celda de encabezado A26
	let rangoA26 = para_compartir.getRange("A26");
	let formatoA26 = rangoA26.getFormat();

	// Limpia diagonales
	formatoA26.getRangeBorder(ExcelScript.BorderIndex.diagonalDown).setStyle(ExcelScript.BorderLineStyle.none);
	formatoA26.getRangeBorder(ExcelScript.BorderIndex.diagonalUp).setStyle(ExcelScript.BorderLineStyle.none);

	// Aplica bordes externos e internos finos
	formatoA26.getRangeBorder(ExcelScript.BorderIndex.edgeLeft).setStyle(ExcelScript.BorderLineStyle.continuous);
	formatoA26.getRangeBorder(ExcelScript.BorderIndex.edgeLeft).setWeight(ExcelScript.BorderWeight.thin);
	formatoA26.getRangeBorder(ExcelScript.BorderIndex.edgeTop).setStyle(ExcelScript.BorderLineStyle.continuous);
	formatoA26.getRangeBorder(ExcelScript.BorderIndex.edgeTop).setWeight(ExcelScript.BorderWeight.thin);
	formatoA26.getRangeBorder(ExcelScript.BorderIndex.edgeBottom).setStyle(ExcelScript.BorderLineStyle.continuous);
	formatoA26.getRangeBorder(ExcelScript.BorderIndex.edgeBottom).setWeight(ExcelScript.BorderWeight.thin);
	formatoA26.getRangeBorder(ExcelScript.BorderIndex.edgeRight).setStyle(ExcelScript.BorderLineStyle.continuous);
	formatoA26.getRangeBorder(ExcelScript.BorderIndex.edgeRight).setWeight(ExcelScript.BorderWeight.thin);
	formatoA26.getRangeBorder(ExcelScript.BorderIndex.insideVertical).setStyle(ExcelScript.BorderLineStyle.continuous);
	formatoA26.getRangeBorder(ExcelScript.BorderIndex.insideVertical).setWeight(ExcelScript.BorderWeight.thin);
	formatoA26.getRangeBorder(ExcelScript.BorderIndex.insideHorizontal).setStyle(ExcelScript.BorderLineStyle.continuous);
	formatoA26.getRangeBorder(ExcelScript.BorderIndex.insideHorizontal).setWeight(ExcelScript.BorderWeight.thin);

	// ==========================================
	// 3. FORMATO DE BORDES (CUERPO DE LA TABLA)
	// ==========================================

	// Extiende la selección desde A7 hacia abajo y derecha para aplicar bordes al bloque principal
	let rangoCuerpo = para_compartir.getRange("A7").getExtendedRange(ExcelScript.KeyboardDirection.down).getExtendedRange(ExcelScript.KeyboardDirection.right);
	let formatoCuerpo = rangoCuerpo.getFormat();

	// Limpia diagonales del cuerpo
	formatoCuerpo.getRangeBorder(ExcelScript.BorderIndex.diagonalDown).setStyle(ExcelScript.BorderLineStyle.none);
	formatoCuerpo.getRangeBorder(ExcelScript.BorderIndex.diagonalUp).setStyle(ExcelScript.BorderLineStyle.none);

	// Aplica bordes completos al cuerpo
	formatoCuerpo.getRangeBorder(ExcelScript.BorderIndex.edgeLeft).setStyle(ExcelScript.BorderLineStyle.continuous);
	formatoCuerpo.getRangeBorder(ExcelScript.BorderIndex.edgeLeft).setWeight(ExcelScript.BorderWeight.thin);
	formatoCuerpo.getRangeBorder(ExcelScript.BorderIndex.edgeTop).setStyle(ExcelScript.BorderLineStyle.continuous);
	formatoCuerpo.getRangeBorder(ExcelScript.BorderIndex.edgeTop).setWeight(ExcelScript.BorderWeight.thin);
	formatoCuerpo.getRangeBorder(ExcelScript.BorderIndex.edgeBottom).setStyle(ExcelScript.BorderLineStyle.continuous);
	formatoCuerpo.getRangeBorder(ExcelScript.BorderIndex.edgeBottom).setWeight(ExcelScript.BorderWeight.thin);
	formatoCuerpo.getRangeBorder(ExcelScript.BorderIndex.edgeRight).setStyle(ExcelScript.BorderLineStyle.continuous);
	formatoCuerpo.getRangeBorder(ExcelScript.BorderIndex.edgeRight).setWeight(ExcelScript.BorderWeight.thin);
	formatoCuerpo.getRangeBorder(ExcelScript.BorderIndex.insideVertical).setStyle(ExcelScript.BorderLineStyle.continuous);
	formatoCuerpo.getRangeBorder(ExcelScript.BorderIndex.insideVertical).setWeight(ExcelScript.BorderWeight.thin);
	formatoCuerpo.getRangeBorder(ExcelScript.BorderIndex.insideHorizontal).setStyle(ExcelScript.BorderLineStyle.continuous);
	formatoCuerpo.getRangeBorder(ExcelScript.BorderIndex.insideHorizontal).setWeight(ExcelScript.BorderWeight.thin);

	// ==========================================
	// 4. ALINEACIÓN Y AJUSTE DE TEXTO
	// ==========================================

	// -- Encabezados de Sección (A26:G26) --
	let rangoEncabezado = para_compartir.getRange("A26:G26");
	rangoEncabezado.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
	rangoEncabezado.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.bottom);
	rangoEncabezado.getFormat().setWrapText(false);
	rangoEncabezado.getFormat().setTextOrientation(0);
	rangoEncabezado.merge(false); // Combina celdas para título centrado

	// Reajuste posterior a la izquierda (parece redundante en el original, pero se mantiene la lógica)
	rangoEncabezado.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);
	rangoEncabezado.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.bottom);
	rangoEncabezado.merge(false);

	// -- Encabezados Superiores (A6:G6) --
	let rangoTituloSup = para_compartir.getRange("A6:G6");
	rangoTituloSup.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
	rangoTituloSup.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.bottom);
	rangoTituloSup.getFormat().setWrapText(false);
	rangoTituloSup.merge(false);
	
	// Reajuste a la izquierda
	rangoTituloSup.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);
}

// ==========================================================================
// 8. GENERAR HORAS NO LABORABLES
// ==========================================================================
/**
 * SCRIPT: Generación de Sección "Horas No Laborables"
 * OBJETIVO: Copiar los datos de tiempos no laborables a la hoja de presentación.
 * FUENTE: Hoja "Horas No Laborables", Rango dinámico desde A3 (incluye totales).
 * DESTINO: Hoja "Para compartir", celda A36.
 */
function step8_GenerarHorasNoLaborables(workbook: ExcelScript.Workbook) {

	// ==========================================
	// 1. REFERENCIAS Y COPIADO
	// ==========================================
	let para_compartir = workbook.getWorksheet("Para compartir");
	let horas_No_Laborables = workbook.getWorksheet("Horas No Laborables");

	// Identifica el rango completo de la tabla dinámica automáticamente
	// getSurroundingRegion() selecciona todo el bloque de datos contiguos desde A3
	let rangoOrigen = horas_No_Laborables.getRange("A3").getSurroundingRegion();

	// Copia los datos a la hoja de presentación
	// Origen: Rango dinámico de "Horas No Laborables"
	// Destino: A36 de "Para compartir"
	para_compartir.getRange("A36").copyFrom(
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
 * ALCANCE: Rangos A6:G24 (Proyectos) y A26:G34/A36:G36 (Admin/No Laborables).
 */
function step9_Ajustes1(workbook: ExcelScript.Workbook) {
	let para_compartir = workbook.getWorksheet("Para compartir");

	// ==========================================
	// 1. FORMATO TABLA PROYECTOS (A6:G24)
	// ==========================================
	
	// -- Bordes --
	let rangoProyectos = para_compartir.getRange("A6:G24");
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
	let rangoNombres = para_compartir.getRange("A7").getExtendedRange(ExcelScript.KeyboardDirection.down);
	rangoNombres.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);
	rangoNombres.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.bottom);

	// ==========================================
	// 2. FORMATO TABLA HORAS ADMIN (A26:G34)
	// ==========================================
	
	// -- Alineación --
	let rangoAdmin = para_compartir.getRange("A26:G34");
	rangoAdmin.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
	rangoAdmin.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center); // Centrado vertical
	rangoAdmin.getFormat().setWrapText(false);

	// -- Bordes --
	let formatoAdmin = rangoAdmin.getFormat();
	formatoAdmin.getRangeBorder(ExcelScript.BorderIndex.diagonalDown).setStyle(ExcelScript.BorderLineStyle.none);
	formatoAdmin.getRangeBorder(ExcelScript.BorderIndex.diagonalUp).setStyle(ExcelScript.BorderLineStyle.none);
	
	// Bordes externos e internos
	formatoAdmin.getRangeBorder(ExcelScript.BorderIndex.edgeLeft).setStyle(ExcelScript.BorderLineStyle.continuous);
	formatoAdmin.getRangeBorder(ExcelScript.BorderIndex.edgeLeft).setWeight(ExcelScript.BorderWeight.thin);
	formatoAdmin.getRangeBorder(ExcelScript.BorderIndex.edgeTop).setStyle(ExcelScript.BorderLineStyle.continuous);
	formatoAdmin.getRangeBorder(ExcelScript.BorderIndex.edgeTop).setWeight(ExcelScript.BorderWeight.thin);
	formatoAdmin.getRangeBorder(ExcelScript.BorderIndex.edgeBottom).setStyle(ExcelScript.BorderLineStyle.continuous);
	formatoAdmin.getRangeBorder(ExcelScript.BorderIndex.edgeBottom).setWeight(ExcelScript.BorderWeight.thin);
	formatoAdmin.getRangeBorder(ExcelScript.BorderIndex.edgeRight).setStyle(ExcelScript.BorderLineStyle.continuous);
	formatoAdmin.getRangeBorder(ExcelScript.BorderIndex.edgeRight).setWeight(ExcelScript.BorderWeight.thin);
	formatoAdmin.getRangeBorder(ExcelScript.BorderIndex.insideVertical).setStyle(ExcelScript.BorderLineStyle.continuous);
	formatoAdmin.getRangeBorder(ExcelScript.BorderIndex.insideVertical).setWeight(ExcelScript.BorderWeight.thin);
	formatoAdmin.getRangeBorder(ExcelScript.BorderIndex.insideHorizontal).setStyle(ExcelScript.BorderLineStyle.continuous);
	formatoAdmin.getRangeBorder(ExcelScript.BorderIndex.insideHorizontal).setWeight(ExcelScript.BorderWeight.thin);

	// -- Encabezado (A26:G26) --
	let rangoEncabezadoAdmin = para_compartir.getRange("A26:G26");
	rangoEncabezadoAdmin.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);
	rangoEncabezadoAdmin.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
	rangoEncabezadoAdmin.merge(false);

	// -- Nombres (A27:A33) --
	let rangoNombresAdmin = para_compartir.getRange("A27:A33");
	rangoNombresAdmin.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);
	rangoNombresAdmin.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);

	// -- Ajuste especial (A27 extendido hacia arriba) --
	// Esto parece corregir un encabezado o celda específica
	let rangoEspecial = para_compartir.getRange("A27").getRangeEdge(ExcelScript.KeyboardDirection.up).getRangeEdge(ExcelScript.KeyboardDirection.up);
	rangoEspecial.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
	rangoEspecial.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.bottom);

	// ==========================================
	// 3. FORMATO TABLA NO LABORABLES (A36:G36 y extensión)
	// ==========================================

	// -- Encabezado (A36:G36) --
	let rangoEncabezadoNL = para_compartir.getRange("A36:G36");
	rangoEncabezadoNL.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);
	rangoEncabezadoNL.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.bottom);
	rangoEncabezadoNL.merge(false);

	// -- Cuerpo (Extendido desde A36:G36 hacia abajo) --
	let rangoCuerpoNL = para_compartir.getRange("A36:G36").getExtendedRange(ExcelScript.KeyboardDirection.down);
	let formatoCuerpoNL = rangoCuerpoNL.getFormat();

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
// 11. ESTILOS FINALES (AZUL Y NEGRITA)
// ==========================================================================
/**
 * SCRIPT: Aplicar Estilos de Encabezados y Totales
 * OBJETIVO: Aplicar negrita y color de fondo azul (#C0E6F5) a los encabezados y filas de totales.
 * ALCANCE: Hoja "Para compartir", rangos de encabezados y filas finales de tablas.
 */
function step11_EstilosFinales(workbook: ExcelScript.Workbook) {
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
	// 2. FORMATO DE ENCABEZADOS (RANGOS FIJOS)
	// ==========================================
	
	// -- Tabla Proyectos (A6:G7) --
	let rangoEncabezadoProy = selectedSheet.getRange("A6:G7");
	rangoEncabezadoProy.getFormat().getFont().setBold(true);
	rangoEncabezadoProy.getFormat().getFill().setColor(colorAzul);

	// -- Tabla Admin (A26:G27) --
	let rangoEncabezadoAdmin = selectedSheet.getRange("A26:G27");
	rangoEncabezadoAdmin.getFormat().getFont().setBold(true);
	rangoEncabezadoAdmin.getFormat().getFill().setColor(colorAzul);

	// -- Tabla No Laborables (A36:G37) --
	let rangoEncabezadoNL = selectedSheet.getRange("A36:G37");
	rangoEncabezadoNL.getFormat().getFont().setBold(true);
	rangoEncabezadoNL.getFormat().getFill().setColor(colorAzul);

	// ==========================================
	// 3. FORMATO DE TOTALES (RANGOS DINÁMICOS)
	// ==========================================
	// Busca la última fila de cada bloque y aplica formato

	// -- Total Proyectos --
	// Desde A6, baja hasta el final del bloque y selecciona esa fila completa hacia la derecha
	let rangoTotalProy = selectedSheet.getRange("A6")
		.getRangeEdge(ExcelScript.KeyboardDirection.down)
		.getExtendedRange(ExcelScript.KeyboardDirection.right);
	
	rangoTotalProy.getFormat().getFill().setColor(colorAzul);
	rangoTotalProy.getFormat().getFont().setBold(true);

	// -- Total Admin --
	// Desde A26, baja hasta el final del bloque
	let rangoTotalAdmin = selectedSheet.getRange("A26")
		.getRangeEdge(ExcelScript.KeyboardDirection.down)
		.getExtendedRange(ExcelScript.KeyboardDirection.right);

	rangoTotalAdmin.getFormat().getFill().setColor(colorAzul);
	rangoTotalAdmin.getFormat().getFont().setBold(true);

	// -- Total No Laborables --
	// Desde A36, baja hasta el final del bloque
	let rangoTotalNL = selectedSheet.getRange("A36")
		.getRangeEdge(ExcelScript.KeyboardDirection.down)
		.getExtendedRange(ExcelScript.KeyboardDirection.right);

	rangoTotalNL.getFormat().getFill().setColor(colorAzul);
	rangoTotalNL.getFormat().getFont().setBold(true);
}