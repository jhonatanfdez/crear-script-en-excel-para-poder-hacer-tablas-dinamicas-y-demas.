/**
 * SCRIPT: Generación de Sección "Horas Admin"
 * OBJETIVO: Copiar datos administrativos y aplicar formato de bordes y alineación.
 * FUENTE: Hoja "Horas Admin", Rango dinámico desde A3.
 * DESTINO: Hoja "Para compartir", celda A26.
 */

function main(workbook: ExcelScript.Workbook) {

	// ==========================================
	// 1. REFERENCIAS Y COPIADO
	// ==========================================
	let para_compartir = workbook.getWorksheet("Para compartir");
	let horas_Admin = workbook.getWorksheet("Horas Admin");

	// Copia los datos de horas administrativas a la hoja de presentación
	para_compartir.getRange("A26").copyFrom(
		horas_Admin.getRange("A3:G11"), 
		ExcelScript.RangeCopyType.values, 
		false, 
		false
	);

	// ==========================================
	// 2. FORMATO DE BORDES (TABLA SUPERIOR)
	// ==========================================
	
	// Configura los bordes para la celda de encabezado A26
	let rangoA26 = para_compartir.getRange("A26");
	let formatoA26 = rangoA26.getFormat();

	// Limpia diagonales
	formatoA26.getRangeBorder(ExcelScript.BorderIndex.diagonalDown).setStyle(ExcelScript.BorderLineStyle.none);
	formatoA26.getRangeBorder(ExcelScript.BorderIndex.diagonalUp).setStyle(ExcelScript.BorderLineStyle.none);

	// Aplica bordes externos e internos finos
	formatoA26.getRangeBorder(ExcelScript.BorderIndex.edgeLeft).setStyle(ExcelScript.BorderLineStyle.continuous);
	formatoA26.getRangeBorder(ExcelScript.BorderIndex.edgeLeft).setWeight(ExcelScript.BorderWeight.thin);
	formatoA26.getRangeBorder(ExcelScript.BorderIndex.edgeTop).setStyle(ExcelScript.BorderLineStyle.continuous);
	formatoA26.getRangeBorder(ExcelScript.BorderIndex.edgeTop).setWeight(ExcelScript.BorderWeight.thin);
	formatoA26.getRangeBorder(ExcelScript.BorderIndex.edgeBottom).setStyle(ExcelScript.BorderLineStyle.continuous);
	formatoA26.getRangeBorder(ExcelScript.BorderIndex.edgeBottom).setWeight(ExcelScript.BorderWeight.thin);
	formatoA26.getRangeBorder(ExcelScript.BorderIndex.edgeRight).setStyle(ExcelScript.BorderLineStyle.continuous);
	formatoA26.getRangeBorder(ExcelScript.BorderIndex.edgeRight).setWeight(ExcelScript.BorderWeight.thin);
	formatoA26.getRangeBorder(ExcelScript.BorderIndex.insideVertical).setStyle(ExcelScript.BorderLineStyle.continuous);
	formatoA26.getRangeBorder(ExcelScript.BorderIndex.insideVertical).setWeight(ExcelScript.BorderWeight.thin);
	formatoA26.getRangeBorder(ExcelScript.BorderIndex.insideHorizontal).setStyle(ExcelScript.BorderLineStyle.continuous);
	formatoA26.getRangeBorder(ExcelScript.BorderIndex.insideHorizontal).setWeight(ExcelScript.BorderWeight.thin);

	// ==========================================
	// 3. FORMATO DE BORDES (CUERPO DE LA TABLA)
	// ==========================================

	// Extiende la selección desde A7 hacia abajo y derecha para aplicar bordes al bloque principal
	let rangoCuerpo = para_compartir.getRange("A7").getExtendedRange(ExcelScript.KeyboardDirection.down).getExtendedRange(ExcelScript.KeyboardDirection.right);
	let formatoCuerpo = rangoCuerpo.getFormat();

	// Limpia diagonales del cuerpo
	formatoCuerpo.getRangeBorder(ExcelScript.BorderIndex.diagonalDown).setStyle(ExcelScript.BorderLineStyle.none);
	formatoCuerpo.getRangeBorder(ExcelScript.BorderIndex.diagonalUp).setStyle(ExcelScript.BorderLineStyle.none);

	// Aplica bordes completos al cuerpo
	formatoCuerpo.getRangeBorder(ExcelScript.BorderIndex.edgeLeft).setStyle(ExcelScript.BorderLineStyle.continuous);
	formatoCuerpo.getRangeBorder(ExcelScript.BorderIndex.edgeLeft).setWeight(ExcelScript.BorderWeight.thin);
	formatoCuerpo.getRangeBorder(ExcelScript.BorderIndex.edgeTop).setStyle(ExcelScript.BorderLineStyle.continuous);
	formatoCuerpo.getRangeBorder(ExcelScript.BorderIndex.edgeTop).setWeight(ExcelScript.BorderWeight.thin);
	formatoCuerpo.getRangeBorder(ExcelScript.BorderIndex.edgeBottom).setStyle(ExcelScript.BorderLineStyle.continuous);
	formatoCuerpo.getRangeBorder(ExcelScript.BorderIndex.edgeBottom).setWeight(ExcelScript.BorderWeight.thin);
	formatoCuerpo.getRangeBorder(ExcelScript.BorderIndex.edgeRight).setStyle(ExcelScript.BorderLineStyle.continuous);
	formatoCuerpo.getRangeBorder(ExcelScript.BorderIndex.edgeRight).setWeight(ExcelScript.BorderWeight.thin);
	formatoCuerpo.getRangeBorder(ExcelScript.BorderIndex.insideVertical).setStyle(ExcelScript.BorderLineStyle.continuous);
	formatoCuerpo.getRangeBorder(ExcelScript.BorderIndex.insideVertical).setWeight(ExcelScript.BorderWeight.thin);
	formatoCuerpo.getRangeBorder(ExcelScript.BorderIndex.insideHorizontal).setStyle(ExcelScript.BorderLineStyle.continuous);
	formatoCuerpo.getRangeBorder(ExcelScript.BorderIndex.insideHorizontal).setWeight(ExcelScript.BorderWeight.thin);

	// ==========================================
	// 4. ALINEACIÓN Y AJUSTE DE TEXTO
	// ==========================================

	// -- Encabezados de Sección (A26:G26) --
	let rangoEncabezado = para_compartir.getRange("A26:G26");
	rangoEncabezado.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
	rangoEncabezado.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.bottom);
	rangoEncabezado.getFormat().setWrapText(false);
	rangoEncabezado.getFormat().setTextOrientation(0);
	rangoEncabezado.merge(false); // Combina celdas para título centrado

	// Reajuste posterior a la izquierda (parece redundante en el original, pero se mantiene la lógica)
	rangoEncabezado.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);
	rangoEncabezado.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.bottom);
	rangoEncabezado.merge(false);

	// -- Encabezados Superiores (A6:G6) --
	let rangoTituloSup = para_compartir.getRange("A6:G6");
	rangoTituloSup.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
	rangoTituloSup.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.bottom);
	rangoTituloSup.getFormat().setWrapText(false);
	rangoTituloSup.merge(false);
	
	// Reajuste a la izquierda
	rangoTituloSup.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);
}