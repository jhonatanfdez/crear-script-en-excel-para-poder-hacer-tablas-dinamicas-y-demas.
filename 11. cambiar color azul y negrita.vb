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
}