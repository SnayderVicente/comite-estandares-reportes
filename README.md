# Comité de Estándares: Reportes y Exportación de Datos

**Proyecto:** PRS04 - SIGRC  
**Instituto:** Valle Grande  
**Componentes revisados:** `vg-ms-therapies` (Backend) y `vg-web-sigrc` (Frontend)  

---

Repositorio de trabajo del **Comité de Estándares de Generación de Reportes y Exportación de Datos (PDF / Excel / CSV)**. Aquí se encuentra el checklist de cumplimiento técnico del estándar y la automatización desarrollada para auditar el código y generar el reporte de métricas.

---

## 📊 Hoja de Cálculo en Google Sheets

El reporte con el dashboard y el checklist completo se puede consultar en el siguiente enlace:

👉 **[Abrir métricas y checklist en Google Sheets](https://docs.google.com/spreadsheets/d/12E1OdZqn3ZDSasf_8IGkvfK_FIxXD2BQ/edit?usp=sharing)**  
*(Contiene las 5 pestañas: Resumen, Checklist PR, Datos, Métricas y Reglas).*

---

## Entregables del Comité

### 1. Checklist de Cumplimiento del Estándar (Entregable 1 - 5 pts)
- **`CHECKLIST_ESTANDAR_REPORTES.md`**: Matriz en Markdown con los 12 criterios evaluados del estándar (§2 al §18), detallando el estado, responsable y la ubicación exacta de las evidencias en el código.
- **`checklist_estandar_reportes.html`**: Instrumento interactivo en HTML. Permite abrir la matriz en el navegador, ver el porcentaje de avance, filtrar por estado y consultar las evidencias de forma visual.

### 2. Agente de IA y Automatización (Entregable 2 - 5 pts)
- **`AGENTE_IA_ENTRENAMIENTO.md`**: Documento de configuración y entrenamiento del agente (`ReportStandardsAuditAgent`), donde se definen su prompt del sistema, las reglas de validación y la lógica para auditar el repositorio.
- **`audit_agent.py`**: Script en Python que recorre los archivos del proyecto, valida las reglas del estándar (dependencias no bloqueantes, seguridad JWT, precisión en montos, nombres de archivos) y genera el archivo Excel con las fórmulas calculadas.
- **`metricas_reportes.xlsx`**: Hoja de cálculo generada por el script con las 5 pestañas y fórmulas nativas de cálculo.

---

## Cómo ejecutar la auditoría

Para correr el script y generar o actualizar el reporte en Excel:

```bash
# Instalar librería requerida
pip install openpyxl

# Ejecutar el script de auditoría
python audit_agent.py
```

Al terminar, el script revisará el código y actualizará el archivo `metricas_reportes.xlsx` con los resultados de la revisión.