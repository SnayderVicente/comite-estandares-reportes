# COMITÉ DE ESTÁNDARES: GENERACIÓN DE REPORTES Y EXPORTACIÓN DE DATOS (PDF / Excel / CSV)
**Proyecto:** PRS04 - SIGRC (Sistema Integrado de Gestión y Registro Clínico)  
**Institución:** Instituto de Educación Superior Valle Grande  
**Auditoría de Componentes:** `vg-ms-therapies` (Backend Spring Boot WebFlux) y `vg-web-sigrc` (Frontend Angular)

---

## 🔗 Enlace Oficial a la Hoja de Cálculo
- **Google Sheets (En vivo):** [Abrir metricas_reportes en Google Sheets](https://docs.google.com/spreadsheets/d/12E1OdZqn3ZDSasf_8IGkvfK_FIxXD2BQ/edit?usp=sharing)
  - Incluye las 5 pestañas reglamentarias: `Resumen` (Dashboard 100% Cumple), `Checklist PR`, `Datos`, `Métricas` y `Reglas`.

---

## Estructura de Entregables del Comité

Esta carpeta contiene los entregables oficiales e independientes del Comité de Estándares:

1. **`CHECKLIST_ESTANDAR_REPORTES.md`**: *(Entregable 1 - Smartdown/Markdown)*
   - Matriz formal de 12 ítems del estándar con referencias (§2 a §18), evidencias de código, responsables y fechas.

2. **`checklist_estandar_reportes.html`**: *(Entregable 1 - Instrumento Interactivo Web)*
   - Aplicación web autónoma con diseño UI moderno, dashboard en tiempo real, KPIs dinámicos, filtros por estado, búsqueda y exportación a JSON, CSV e impresión.

3. **`AGENTE_IA_ENTRENAMIENTO.md`**: *(Entregable 2 - Documento de Entrenamiento y Arquitectura)*
   - Especificación completa del Agente de IA (`ReportStandardsAuditAgent`), System Prompt, cadena de razonamiento (Chain-of-Thought) y ejemplos Few-Shot de inferencia.

4. **`audit_agent.py`**: *(Entregable 2 - Motor Ejecutable del Agente de IA)*
   - Script en Python que audita automáticamente el código de los microservicios y compila el reporte.

5. **`metricas_reportes.xlsx`**: *(Entregable 2 - Hoja de Cálculo Oficial Generada)*
   - Libro Excel con las 5 pestañas oficiales: `Resumen`, `Checklist PR`, `Datos`, `Métricas` y `Reglas`, con fórmulas de Excel nativas y estilos corporativos.

---

## Cómo Ejecutar el Agente de IA

```bash
# 1. Instalar dependencia requerida (si no está instalada)
pip install openpyxl

# 2. Ejecutar la auditoría automatizada
python audit_agent.py

# El agente auditará los componentes y actualizará automáticamente 'metricas_reportes.xlsx'
```