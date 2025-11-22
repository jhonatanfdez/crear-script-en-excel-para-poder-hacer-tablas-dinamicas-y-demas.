/**
 * SCRIPT: Creación de Reporte "Horas No Laborables"
 * OBJETIVO: Generar una tabla dinámica para analizar tiempos no laborables (Vacaciones, Licencias, etc.).
 * FUENTE DE DATOS: Hoja "Datos TM+", Rango A1:AA884.
 * SALIDA: Hoja nueva llamada "Horas No Laborables".
 */

function main(workbook: ExcelScript.Workbook) {

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