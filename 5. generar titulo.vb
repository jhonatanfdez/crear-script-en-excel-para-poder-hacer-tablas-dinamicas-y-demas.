/**
 * SCRIPT: Generación de Títulos y Hoja "Para compartir"
 * OBJETIVO: Crear la hoja final de presentación y configurar los encabezados institucionales.
 * SALIDA: Hoja nueva llamada "Para compartir" con formato inicial.
 */

function main(workbook: ExcelScript.Workbook) {

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
