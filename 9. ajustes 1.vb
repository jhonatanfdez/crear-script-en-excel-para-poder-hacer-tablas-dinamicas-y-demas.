/**
 * SCRIPT: Ajustes de Formato - Parte 1
 * OBJETIVO: Aplicar bordes, alineación y estilos a las tablas de "Proyectos" y "Horas No Laborables".
 * ALCANCE: Rangos A6:G24 (Proyectos) y A26:G34/A36:G36 (Admin/No Laborables).
 */

function main(workbook: ExcelScript.Workbook) {
	let para_compartir = workbook.getWorksheet("Para compartir");

	// ==========================================
	// 1. FORMATO TABLA PROYECTOS (A6:G24)
	// ==========================================
	
	// -- Bordes --
	let rangoProyectos = para_compartir.getRange("A6:G24");
	let formatoProyectos = rangoProyectos.getFormat();

	// Limpia diagonales
	formatoProyectos.getRangeBorder(ExcelScript.BorderIndex.diagonalDown).setStyle(ExcelScript.BorderLineStyle.none);
	formatoProyectos.getRangeBorder(ExcelScript.BorderIndex.diagonalUp).setStyle(ExcelScript.BorderLineStyle.none);

	// Aplica bordes finos continuos en todos los lados e interiores
	formatoProyectos.getRangeBorder(ExcelScript.BorderIndex.edgeLeft).setStyle(ExcelScript.BorderLineStyle.continuous);
	formatoProyectos.getRangeBorder(ExcelScript.BorderIndex.edgeLeft).setWeight(ExcelScript.BorderWeight.thin);
	formatoProyectos.getRangeBorder(ExcelScript.BorderIndex.edgeTop).setStyle(ExcelScript.BorderLineStyle.continuous);
	formatoProyectos.getRangeBorder(ExcelScript.BorderIndex.edgeTop).setWeight(ExcelScript.BorderWeight.thin);
	formatoProyectos.getRangeBorder(ExcelScript.BorderIndex.edgeBottom).setStyle(ExcelScript.BorderLineStyle.continuous);
	formatoProyectos.getRangeBorder(ExcelScript.BorderIndex.edgeBottom).setWeight(ExcelScript.BorderWeight.thin);
	formatoProyectos.getRangeBorder(ExcelScript.BorderIndex.edgeRight).setStyle(ExcelScript.BorderLineStyle.continuous);
	formatoProyectos.getRangeBorder(ExcelScript.BorderIndex.edgeRight).setWeight(ExcelScript.BorderWeight.thin);
	formatoProyectos.getRangeBorder(ExcelScript.BorderIndex.insideVertical).setStyle(ExcelScript.BorderLineStyle.continuous);
	formatoProyectos.getRangeBorder(ExcelScript.BorderIndex.insideVertical).setWeight(ExcelScript.BorderWeight.thin);
	formatoProyectos.getRangeBorder(ExcelScript.BorderIndex.insideHorizontal).setStyle(ExcelScript.BorderLineStyle.continuous);
	formatoProyectos.getRangeBorder(ExcelScript.BorderIndex.insideHorizontal).setWeight(ExcelScript.BorderWeight.thin);

	// -- Alineación General --
	formatoProyectos.setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
	formatoProyectos.setVerticalAlignment(ExcelScript.VerticalAlignment.bottom);
	formatoProyectos.setWrapText(false);
	formatoProyectos.setTextOrientation(0);

	// -- Encabezado Específico (A6) --
	let rangoA6 = para_compartir.getRange("A6");
	rangoA6.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);
	rangoA6.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.bottom);
	rangoA6.merge(false);

	// -- Columna A (Nombres) --
	// Extiende desde A7 hacia abajo para alinear nombres a la izquierda
	let rangoNombres = para_compartir.getRange("A7").getExtendedRange(ExcelScript.KeyboardDirection.down);
	rangoNombres.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);
	rangoNombres.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.bottom);

	// ==========================================
	// 2. FORMATO TABLA HORAS ADMIN (A26:G34)
	// ==========================================
	
	// -- Alineación --
	let rangoAdmin = para_compartir.getRange("A26:G34");
	rangoAdmin.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
	rangoAdmin.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center); // Centrado vertical
	rangoAdmin.getFormat().setWrapText(false);

	// -- Bordes --
	let formatoAdmin = rangoAdmin.getFormat();
	formatoAdmin.getRangeBorder(ExcelScript.BorderIndex.diagonalDown).setStyle(ExcelScript.BorderLineStyle.none);
	formatoAdmin.getRangeBorder(ExcelScript.BorderIndex.diagonalUp).setStyle(ExcelScript.BorderLineStyle.none);
	
	// Bordes externos e internos
	formatoAdmin.getRangeBorder(ExcelScript.BorderIndex.edgeLeft).setStyle(ExcelScript.BorderLineStyle.continuous);
	formatoAdmin.getRangeBorder(ExcelScript.BorderIndex.edgeLeft).setWeight(ExcelScript.BorderWeight.thin);
	formatoAdmin.getRangeBorder(ExcelScript.BorderIndex.edgeTop).setStyle(ExcelScript.BorderLineStyle.continuous);
	formatoAdmin.getRangeBorder(ExcelScript.BorderIndex.edgeTop).setWeight(ExcelScript.BorderWeight.thin);
	formatoAdmin.getRangeBorder(ExcelScript.BorderIndex.edgeBottom).setStyle(ExcelScript.BorderLineStyle.continuous);
	formatoAdmin.getRangeBorder(ExcelScript.BorderIndex.edgeBottom).setWeight(ExcelScript.BorderWeight.thin);
	formatoAdmin.getRangeBorder(ExcelScript.BorderIndex.edgeRight).setStyle(ExcelScript.BorderLineStyle.continuous);
	formatoAdmin.getRangeBorder(ExcelScript.BorderIndex.edgeRight).setWeight(ExcelScript.BorderWeight.thin);
	formatoAdmin.getRangeBorder(ExcelScript.BorderIndex.insideVertical).setStyle(ExcelScript.BorderLineStyle.continuous);
	formatoAdmin.getRangeBorder(ExcelScript.BorderIndex.insideVertical).setWeight(ExcelScript.BorderWeight.thin);
	formatoAdmin.getRangeBorder(ExcelScript.BorderIndex.insideHorizontal).setStyle(ExcelScript.BorderLineStyle.continuous);
	formatoAdmin.getRangeBorder(ExcelScript.BorderIndex.insideHorizontal).setWeight(ExcelScript.BorderWeight.thin);

	// -- Encabezado (A26:G26) --
	let rangoEncabezadoAdmin = para_compartir.getRange("A26:G26");
	rangoEncabezadoAdmin.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);
	rangoEncabezadoAdmin.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
	rangoEncabezadoAdmin.merge(false);

	// -- Nombres (A27:A33) --
	let rangoNombresAdmin = para_compartir.getRange("A27:A33");
	rangoNombresAdmin.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);
	rangoNombresAdmin.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);

	// -- Ajuste especial (A27 extendido hacia arriba) --
	// Esto parece corregir un encabezado o celda específica
	let rangoEspecial = para_compartir.getRange("A27").getRangeEdge(ExcelScript.KeyboardDirection.up).getRangeEdge(ExcelScript.KeyboardDirection.up);
	rangoEspecial.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
	rangoEspecial.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.bottom);

	// ==========================================
	// 3. FORMATO TABLA NO LABORABLES (A36:G36 y extensión)
	// ==========================================

	// -- Encabezado (A36:G36) --
	let rangoEncabezadoNL = para_compartir.getRange("A36:G36");
	rangoEncabezadoNL.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);
	rangoEncabezadoNL.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.bottom);
	rangoEncabezadoNL.merge(false);

	// -- Cuerpo (Extendido desde A36:G36 hacia abajo) --
	let rangoCuerpoNL = para_compartir.getRange("A36:G36").getExtendedRange(ExcelScript.KeyboardDirection.down);
	let formatoCuerpoNL = rangoCuerpoNL.getFormat();

	// Bordes
	formatoCuerpoNL.getRangeBorder(ExcelScript.BorderIndex.diagonalDown).setStyle(ExcelScript.BorderLineStyle.none);
	formatoCuerpoNL.getRangeBorder(ExcelScript.BorderIndex.diagonalUp).setStyle(ExcelScript.BorderLineStyle.none);
	formatoCuerpoNL.getRangeBorder(ExcelScript.BorderIndex.edgeLeft).setStyle(ExcelScript.BorderLineStyle.continuous);
	formatoCuerpoNL.getRangeBorder(ExcelScript.BorderIndex.edgeLeft).setWeight(ExcelScript.BorderWeight.thin);
	formatoCuerpoNL.getRangeBorder(ExcelScript.BorderIndex.edgeTop).setStyle(ExcelScript.BorderLineStyle.continuous);
	formatoCuerpoNL.getRangeBorder(ExcelScript.BorderIndex.edgeTop).setWeight(ExcelScript.BorderWeight.thin);
	formatoCuerpoNL.getRangeBorder(ExcelScript.BorderIndex.edgeBottom).setStyle(ExcelScript.BorderLineStyle.continuous);
	formatoCuerpoNL.getRangeBorder(ExcelScript.BorderIndex.edgeBottom).setWeight(ExcelScript.BorderWeight.thin);
	formatoCuerpoNL.getRangeBorder(ExcelScript.BorderIndex.edgeRight).setStyle(ExcelScript.BorderLineStyle.continuous);
	formatoCuerpoNL.getRangeBorder(ExcelScript.BorderIndex.edgeRight).setWeight(ExcelScript.BorderWeight.thin);
	formatoCuerpoNL.getRangeBorder(ExcelScript.BorderIndex.insideVertical).setStyle(ExcelScript.BorderLineStyle.continuous);
	formatoCuerpoNL.getRangeBorder(ExcelScript.BorderIndex.insideVertical).setWeight(ExcelScript.BorderWeight.thin);
	formatoCuerpoNL.getRangeBorder(ExcelScript.BorderIndex.insideHorizontal).setStyle(ExcelScript.BorderLineStyle.continuous);
	formatoCuerpoNL.getRangeBorder(ExcelScript.BorderIndex.insideHorizontal).setWeight(ExcelScript.BorderWeight.thin);

	// Alineación
	formatoCuerpoNL.setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
	formatoCuerpoNL.setVerticalAlignment(ExcelScript.VerticalAlignment.center);
	formatoCuerpoNL.setWrapText(false);
}