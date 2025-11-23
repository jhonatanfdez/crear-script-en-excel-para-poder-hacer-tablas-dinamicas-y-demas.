/**
 * SCRIPT: Aplicar Estilos de Encabezados y Totales
 * OBJETIVO: Aplicar negrita y color de fondo azul (#C0E6F5) a los encabezados y filas de totales.
 * ALCANCE: Hoja "Para compartir", rangos de encabezados y filas finales de tablas.
 */
function main(workbook: ExcelScript.Workbook) {
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