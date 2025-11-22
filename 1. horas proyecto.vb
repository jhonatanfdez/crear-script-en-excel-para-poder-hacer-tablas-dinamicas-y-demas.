/**
 * SCRIPT: Creación de Reporte "Horas Proyectos"
 * OBJETIVO: Generar una nueva hoja con una tabla dinámica basada en los datos maestros.
 * FUENTE DE DATOS: Hoja "Datos TM+", Rango A1:AA884.
 * SALIDA: Hoja nueva llamada "Horas Proyectos".
 */

function main(workbook: ExcelScript.Workbook) {

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