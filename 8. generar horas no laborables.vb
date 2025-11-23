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