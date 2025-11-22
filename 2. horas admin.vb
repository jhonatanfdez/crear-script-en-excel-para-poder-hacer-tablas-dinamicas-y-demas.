/**
 * SCRIPT: Creación de Reporte "Horas Admin"
 * OBJETIVO: Generar una tabla dinámica para analizar las horas administrativas.
 * FUENTE DE DATOS: Hoja "Datos TM+", Rango A1:AA884.
 * SALIDA: Hoja nueva llamada "Horas Admin".
 */

function main(workbook: ExcelScript.Workbook) {

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