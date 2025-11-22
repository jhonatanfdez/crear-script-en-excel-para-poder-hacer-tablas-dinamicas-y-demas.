/**
 * SCRIPT: Ajustes Finales de Formato (Parte 2)
 * OBJETIVO: Ajustar automáticamente el ancho de columnas y alinear celdas específicas al final de la hoja.
 * ALCANCE: Hoja "Para compartir", ajuste global y rangos dinámicos al final.
 */

function main(workbook: ExcelScript.Workbook) {
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