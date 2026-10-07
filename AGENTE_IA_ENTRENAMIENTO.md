# Entrenamiento y Configuración del Agente de IA

**Proyecto:** PRS04 - SIGRC  
**Microservicios evaluados:** `vg-ms-therapies` (Backend) y `vg-web-sigrc` (Frontend)  
**Comité:** Generación de Reportes y Exportación de Datos  
**Entregable:** 2 (Automatización con IA y reporte en hoja de cálculo)  

---

## 1. Objetivo y funcionamiento del agente

El objetivo de este agente es automatizar la revisión de código para verificar que el equipo cumpla con el estándar de reportes y exportación de datos. En lugar de revisar manualmente archivo por archivo, el agente analiza el código fuente, busca las evidencias técnicas correspondientes a cada regla del estándar y genera automáticamente el reporte en Excel (`metricas_reportes.xlsx`).

### Flujo de revisión:
1. **Lectura de archivos:** Revisa el `pom.xml`, configuraciones de seguridad (`SecurityConfig.java`), controladores, servicios y componentes de Angular.
2. **Evaluación de reglas:** Compara lo encontrado contra los 12 criterios del estándar (formatos, lazy loading, patrones de nombres, JWT, etc.).
3. **Extracción de evidencias:** Identifica las clases, métodos y líneas exactas donde se cumple o falla cada criterio.
4. **Generación del reporte:** Construye el archivo de métricas con fórmulas automáticas para el cálculo de porcentajes y KPIs.

---

## 2. Prompt del Sistema (System Prompt)

Para que el modelo de IA entienda el contexto del proyecto y evalúe correctamente el código, se le proporciona la siguiente instrucción base:

```text
ROL:
Eres un auditor técnico de código para el proyecto PRS04 (SIGRC). Tu tarea es revisar el código fuente de los microservicios en Spring Boot y Angular para verificar el cumplimiento del estándar de reportes y exportación de datos.

INSTRUCCIONES:
1. Revisa los archivos de código fuente y busca evidencias reales (nombre de archivo, clase y línea).
2. Evalúa cada uno de los 12 criterios del estándar:
   - Justificación de formatos (PDF para comprobantes, Excel para tablas, CSV para interoperabilidad).
   - Lazy loading de librerías como jsPDF o XLSX en el frontend.
   - Patrón de nombres de archivo: [Tipo]_[PREFIJO]-[YYYY-MM-DD]-[CORRELATIVO].[ext].
   - Documentación OpenAPI/Swagger en los endpoints de reportes.
   - Aplicación de los mismos filtros de pantalla en las consultas del backend.
   - Botón deshabilitado y spinner mientras se genera la exportación.
   - Uso obligatorio de BigDecimal para montos monetarios (prohibido float/double) y fechas en ISO-8601 UTC.
   - Encabezado institucional Valle Grande y paginación en PDFs.
   - Protección de endpoints con token JWT y roles autorizados.
   - Backend reactivo con Spring WebFlux (sin motores bloqueantes como JasperReports o POI en el hilo principal).
   - Definición del patrón de generación (A, B o C).
   - Pruebas unitarias de controladores, DTOs y correlativos.
3. Clasifica cada criterio en: Cumple, No cumple, Pendiente o NA.
4. Exporta el resultado a la hoja de cálculo respetando la estructura de pestañas: Resumen, Checklist PR, Datos, Métricas y Reglas.
```

---

## 3. Ejemplos de validación (Few-Shot Prompting)

Para asegurar que el agente razone de forma consistente y no invente datos, se le entregaron ejemplos prácticos de cómo debe analizar el código:

### Ejemplo 1: Verificación de tipos de datos para montos (Criterio 7)
* **Código analizado:**
  ```java
  // TherapyPrice.java
  public class TherapyPrice {
      private BigDecimal price;
      private LocalDateTime createdAt;
  }
  ```
* **Lógica del agente:**
  - Se revisa el atributo `price`. El tipo de dato es `BigDecimal`, evitando errores de precisión de punto flotante (`float`/`double`).
  - El campo `createdAt` maneja fecha en formato ISO/UTC.
* **Resultado:**
  - **Estado:** Cumple  
  - **Evidencia:** `TherapyPrice.java#L20: private BigDecimal price; uso de BigDecimal para importes monetarios`.

### Ejemplo 2: Verificación de nomenclatura de archivos (Criterio 3)
* **Código analizado:**
  ```java
  // TicketApplicationService.java
  private static final String PREFIX = "TRAT";
  return String.format("%s-%s-%06d", PREFIX, today, seq.getSequence());
  ```
* **Lógica del agente:**
  - Se valida el prefijo del módulo (`TRAT`).
  - La fecha sigue el formato `YYYY-MM-DD` y el correlativo tiene 6 dígitos con ceros a la izquierda.
* **Resultado:**
  - **Estado:** Cumple  
  - **Evidencia:** `TicketApplicationService.java#L43: formato TRAT-YYYY-MM-DD-000000 generado con correlativo de BD`.

### Ejemplo 3: Ausencia de librerías bloqueantes en WebFlux (Criterio 10)
* **Archivo analizado:** `pom.xml`
* **Lógica del agente:**
  - Se confirma la presencia de dependencias reactivas (`spring-boot-starter-webflux`, `spring-boot-starter-data-mongodb-reactive`).
  - Se verifica que no existan dependencias bloqueantes como `jasperreports` o `org.apache.poi` que congelen el bucle de eventos reactivo de Netty.
* **Resultado:**
  - **Estado:** Cumple  
  - **Evidencia:** `pom.xml: stack 100% reactivo con Spring WebFlux; no incluye motores bloqueantes`.

---

## 4. Script de automatización (`audit_agent.py`)

Para llevar a la práctica la evaluación, se implementó el script `audit_agent.py` en Python. Este script:
1. Escanea los archivos del proyecto `vg-ms-therapies` y verifica las evidencias técnicas.
2. Evalúa los 12 puntos del estándar.
3. Genera el archivo `metricas_reportes.xlsx` con las 5 pestañas solicitadas:
   - **Resumen:** Conteo de ítems, porcentaje de cumplimiento y estado general.
   - **Checklist PR:** Detalle de cada criterio con su evidencia y responsable.
   - **Datos:** Muestra de 20 reportes procesados con formato, peso en KB y tiempo en segundos.
   - **Métricas:** Comparativo de tiempos, tasa de éxito y metas operativas.
   - **Reglas:** Guía de uso y reglas de cálculo.

---

## 5. Ejecución del script

Para correr el agente y actualizar la hoja de cálculo:

```bash
pip install openpyxl
python audit_agent.py
```
