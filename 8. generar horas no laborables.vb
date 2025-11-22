/**
 * SCRIPT: Generación de Sección "Horas No Laborables"
 * OBJETIVO: Copiar los datos de tiempos no laborables a la hoja de presentación.
 * FUENTE: Hoja "Horas No Laborables", Rango dinámico desde A3 (incluye totales).
 * DESTINO: Hoja "Para compartir", celda A36.
 */

function main(workbook: ExcelScript.Workbook) {

	// ==========================================
	// 1. REFERENCIAS Y COPIADO
	// ==========================================
	let para_compartir = workbook.getWorksheet("Para compartir");
	let horas_No_Laborables = workbook.getWorksheet("Horas No Laborables");

	// Copia los datos a la hoja de presentación
	// Origen: A3:G9 de "Horas No Laborables"
	// Destino: A36 de "Para compartir"
	para_compartir.getRange("A36").copyFrom(
		horas_No_Laborables.getRange("A3:G9"), 
		ExcelScript.RangeCopyType.values, 
		false, 
		false
	);
}