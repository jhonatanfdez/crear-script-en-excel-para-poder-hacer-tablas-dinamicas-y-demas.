# Automatización de Reporte de Horas - Excel Office Scripts

Este proyecto contiene un conjunto de scripts de Office (TypeScript) diseñados para automatizar la generación del reporte trimestral de horas.

## Descripción

El objetivo principal es procesar los datos crudos de la hoja "Datos TM+" y generar un reporte ejecutivo en una hoja limpia llamada "Para compartir". El proceso incluye la creación de tablas dinámicas, filtrado inteligente, copiado de valores, formateo estético y cálculo de totales generales.

## Estructura del Proyecto

El proyecto consta de scripts individuales para desarrollo/pruebas y un script maestro para la ejecución productiva.

### Script Maestro
*   `script_completo.vb`: Contiene toda la lógica unificada. Es el único script que se debe ejecutar para generar el reporte completo de principio a fin.

### Scripts Individuales (Módulos)
*   `0. limpieza inicial.vb`: Elimina las hojas generadas previamente para reiniciar el proceso.
*   `1. horas proyecto.vb`: Genera la tabla dinámica de horas de proyectos.
*   `2. horas admin.vb`: Genera la tabla dinámica de horas administrativas.
*   `3. horas no laborables.vb`: Genera la tabla dinámica de horas no laborables.
*   `4. hacer filtros.vb`: Aplica reglas de negocio para filtrar categorías (ej. ocultar feriados en admin).
*   `5. generar titulo.vb`: Crea la hoja "Para compartir" con encabezados institucionales.
*   `6. generar proyectos.vb`: Copia los datos de proyectos a la hoja final.
*   `7. generar horas admin.vb`: Copia los datos administrativos a la hoja final.
*   `8. generar horas no laborables.vb`: Copia los datos no laborables a la hoja final.
*   `9. ajustes 1.vb`: Aplica formatos básicos de bordes y alineación.
*   `10 ajustes 2.vb`: Ajusta anchos de columna y formatos finales.
*   `11. generar gran total.vb`: Calcula la sumatoria total de las tres secciones.
*   `12. estilos finales.vb`: Aplica estilos corporativos (Azul #C0E6F5, Negrita) y bordes finales.

## Flujo de Ejecución

Al ejecutar `script_completo.vb`, el sistema realiza los siguientes pasos secuencialmente:

1.  **Limpieza**: Elimina hojas antiguas ("Para compartir", "Horas Proyectos", etc.).
2.  **Generación de Tablas**: Crea 3 tablas dinámicas en hojas temporales basadas en "Datos TM+".
3.  **Filtrado**: Aplica filtros para excluir o incluir categorías específicas (Vacaciones, Feriados, etc.).
4.  **Maquetación**: Prepara la hoja "Para compartir" con los títulos institucionales.
5.  **Consolidación**: Copia los valores de las tablas dinámicas a la hoja de presentación.
6.  **Formato**: Aplica bordes, alineaciones y ajusta el ancho de las columnas.
7.  **Cálculo**: Genera una fila de "Total general" sumando los totales de Proyectos + Admin + No Laborables.
8.  **Estilizado**: Aplica el color azul corporativo y negritas a encabezados y totales.

## Requisitos

*   Microsoft Excel (Web o Desktop) con soporte para Office Scripts.
*   Hoja de datos fuente llamada `Datos TM+` con la estructura de columnas esperada.

## Autor
Jhonatan David Fernandez Rosa
