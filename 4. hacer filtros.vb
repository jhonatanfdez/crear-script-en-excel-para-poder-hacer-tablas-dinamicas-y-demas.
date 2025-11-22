/**
 * SCRIPT: Aplicación de Filtros Inteligentes
 * OBJETIVO: Filtrar las tablas dinámicas de "Horas Admin" y "Horas No Laborables" según reglas de negocio.
 * LÓGICA: 
 *   - Admin: Oculta Feriados/Licencias/Blancos.
 *   - No Laborables: Muestra SOLO Feriados/Licencias.
 */

function main(workbook: ExcelScript.Workbook) {
  // =====================================================
  // 1. DEFINICIÓN DE PREFERENCIAS Y REGLAS
  // =====================================================
  const nombreCampo = "Categoría de tiempo";
  // Lista de categorías consideradas "No Laborables"
  const listaClave = ["Feriados", "Licencias", "Permisos", "Vacaciones"];
  const nombreEnBlanco = "(en blanco)";

  // =====================================================
  // 2. PROCESAR HOJA "HORAS ADMIN"
  // =====================================================
  procesarTablaInteligente(workbook, "Horas Admin", nombreCampo, (nombreItem) => {
    // REGLA DE NEGOCIO:
    // - Si es blanco -> Ocultar
    // - Si está en la lista de No Laborables -> Ocultar
    // - Todo lo demás (ej. Reuniones, Entrenamientos) -> Mostrar
    if (nombreItem === nombreEnBlanco) return false;
    if (listaClave.includes(nombreItem)) return false;
    return true; 
  });

  // =====================================================
  // 3. PROCESAR HOJA "HORAS NO LABORABLES"
  // =====================================================
  procesarTablaInteligente(workbook, "Horas No Laborables", nombreCampo, (nombreItem) => {
    // REGLA DE NEGOCIO:
    // - Si es blanco -> Ocultar
    // - Si está en la lista de No Laborables -> Mostrar
    // - Todo lo demás -> Ocultar
    if (nombreItem === nombreEnBlanco) return false;
    if (listaClave.includes(nombreItem)) return true;
    return false; 
  });
}

// =====================================================
// FUNCIÓN AUXILIAR: PROCESAMIENTO ROBUSTO DE TABLAS
// =====================================================
/**
 * Recorre los ítems de una tabla dinámica y aplica visibilidad según una función de regla.
 * Maneja casos especiales como "(en blanco)" que a veces es string vacío.
 */
function procesarTablaInteligente(
  workbook: ExcelScript.Workbook,
  nombreHoja: string,
  nombreCampo: string,
  reglaVisibilidad: (nombre: string) => boolean
) {
  const hoja = workbook.getWorksheet(nombreHoja);
  if (!hoja) return;

  const tabla = hoja.getPivotTables()[0];
  if (!tabla) return;

  console.log(`--- Procesando: ${nombreHoja} ---`);

  // Obtenemos los textos visibles de la tabla para iterar sobre lo que existe
  const rangoTabla = tabla.getLayout().getRange();
  const textos = rangoTabla.getTexts();

  const jerarquia = tabla.getRowHierarchy(nombreCampo);
  if (!jerarquia) {
    console.log("Campo no encontrado");
    return;
  }
  const campo = jerarquia.getFields()[0];

  // Recorremos las filas leídas (empezando en 1 para saltar encabezados)
  for (let i = 1; i < textos.length; i++) {
    let nombreItemDetectado = textos[i][0]; // Texto en la celda

    if (nombreItemDetectado && nombreItemDetectado !== "Total general") {

      // Evaluamos la regla
      const debeVerse = reglaVisibilidad(nombreItemDetectado);

      // === INTENTO 1: Buscar por el nombre exacto ===
      let item = campo.getPivotItem(nombreItemDetectado);

      // === INTENTO 2 (Plan B): Manejo de blancos ===
      if (!item && (nombreItemDetectado === "(en blanco)" || nombreItemDetectado === "(blank)")) {
        item = campo.getPivotItem("");
      }

      // Aplicar visibilidad si se encontró el ítem
      if (item) {
        try {
          // Optimización: Solo cambiar si es necesario
          if (item.getVisible() !== debeVerse) {
            item.setVisible(debeVerse);
            console.log(`   Ajustado: "${nombreItemDetectado}" -> ${debeVerse}`);
          }
        } catch (e) {
          // Fallback en caso de error de lectura de estado
          item.setVisible(debeVerse);
        }
      }
    }
  }
  console.log("✅ Proceso completado.");
}