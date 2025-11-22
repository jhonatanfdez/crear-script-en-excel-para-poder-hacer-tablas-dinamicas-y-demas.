/**
 * SCRIPT: Generación de Sección "Proyectos"
 * OBJETIVO: Copiar los datos procesados de la tabla dinámica de Proyectos a la hoja de presentación.
 * FUENTE: Hoja "Horas Proyectos", Rango dinámico desde A3.
 * DESTINO: Hoja "Para compartir", celda A6.
 */

function main(workbook: ExcelScript.Workbook) {

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

	// Copia los valores de la tabla dinámica de proyectos
	// Origen: A3:G21 de "Horas Proyectos"
	// Destino: A6 de "Para compartir"
	// Solo copia valores (sin formato de tabla dinámica)
	para_compartir.getRange("A6").copyFrom(
		horas_Proyectos.getRange("A3:G21"), 
		ExcelScript.RangeCopyType.values, 
		false, 
		false
	);
}