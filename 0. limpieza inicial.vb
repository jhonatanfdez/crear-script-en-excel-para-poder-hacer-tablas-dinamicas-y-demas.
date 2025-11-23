/**
 * SCRIPT: Limpieza Inicial
 * OBJETIVO: Eliminar las hojas generadas en ejecuciones anteriores para evitar errores de duplicidad.
 * HOJAS A ELIMINAR: "Para compartir", "Horas Proyectos", "Horas Admin", "Horas No Laborables".
 */
function main(workbook: ExcelScript.Workbook) {
    // Lista de nombres de hojas que queremos limpiar
    const hojasAEliminar = [
        "Para compartir", 
        "Horas Proyectos", 
        "Horas Admin", 
        "Horas No Laborables"
    ];

    // Recorremos la lista y eliminamos si existen
    hojasAEliminar.forEach(nombreHoja => {
        let hoja = workbook.getWorksheet(nombreHoja);
        if (hoja) {
            hoja.delete();
            console.log(`Hoja eliminada: ${nombreHoja}`);
        } else {
            console.log(`Hoja no encontrada (ya estaba limpia): ${nombreHoja}`);
        }
    });
}
