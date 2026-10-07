# Checklist de Cumplimiento: Estándar de Reportes y Exportación

**Comité:** Generación de Reportes y Exportación de Datos (PDF / Excel / CSV)  
**Proyecto:** PRS04 - SIGRC  
**Componentes revisados:** `vg-ms-therapies` (Backend Spring Boot WebFlux) y `vg-web-sigrc` (Frontend Angular)  
**Versión del estándar:** 2.4.0  

---

## 1. Propósito del Checklist

Este documento sirve como guía y lista de verificación técnica para validar que las funcionalidades de generación y descarga de datos (PDF, Excel, CSV) en el proyecto PRS04 cumplan con las reglas de arquitectura reactiva, seguridad JWT, precisión en cálculos y buenas prácticas de interfaz de usuario.

### Criterios de calificación:
* **Cumple (C):** El criterio está implementado correctamente y cuenta con evidencia verificable en el código. Suma al porcentaje de cumplimiento.
* **No Cumple (NC):** El criterio aplica al componente pero no cumple las reglas del estándar.
* **Pendiente (P):** Criterio en proceso de implementación o revisión pendiente. No penaliza directamente el porcentaje pero reduce el avance total.
* **No Aplica (NA):** Criterio que no corresponde a la tecnología del componente analizado.

### Fórmulas de cálculo:
* **% Cumplimiento** = `Cumple / (Cumple + No Cumple) * 100`
* **% Avance de Revisión** = `(Cumple + No Cumple) / Total Ítems Aplicables * 100`

---

## 2. Matriz de Verificación del Estándar

| # | Criterio | Ref. | Verificación Técnica | Estado | Evidencia en Código | Responsable | Fecha | Observaciones |
|---|----------|------|----------------------|:------:|---------------------|-------------|:-----:|---------------|
| **1** | **Formato justificado por tipo de usuario** | §3 | Los formatos responden a una necesidad clara: PDF para comprobantes inmutables de pacientes; Excel para análisis masivo de datos; CSV para integración y analítica. | **Cumple** | `TherapyController.java`: Comprobantes clínicos emitidos en PDF y consultas masivas preparadas para tabular. | Comité Estándar PRS04 | 2026-10-06 | PDF para documentos finales y Excel para exportaciones tabulares. |
| **2** | **Lazy import de jsPDF / XLSX** | §5, §7 | En el frontend, librerías pesadas (`jspdf`, `xlsx`) no deben ir en el bundle principal. Se cargan bajo demanda mediante importación dinámica (`import()`). | **Cumple** | `vg-web-sigrc/angular.json` y dynamic imports: `import('jspdf').then(...)`. No infla el bundle inicial de Angular. | Frontend Lead | 2026-10-06 | Las dependencias pesadas se aíslan en chunks asíncronos. |
| **3** | **Patrón de nombre de archivo** | §6.4 | El nombre del archivo sigue el patrón: `[Tipo]_[PREFIJO]-[YYYY-MM-DD]-[CORRELATIVO].[ext]`. | **Cumple** | `TicketApplicationService.java#L43`: `String.format("%s-%s-%06d", PREFIX, today, seq.getSequence())`. Salida: `Comprobante_TRAT-2026-09-27-000005.pdf`. | Backend Lead | 2026-10-06 | Generación con correlativo incremental desde MongoDB. |
| **4** | **Endpoint fuente documentado** | §8 | Los endpoints que entregan datos para reportes están documentados en OpenAPI / Swagger, con parámetros y esquemas. | **Cumple** | `SecurityConfig.java#L32-34` permite `/v3/api-docs/**` y `/swagger-ui/**`. Controladores documentados en `/api/v1/therapies/**`. | Backend Dev | 2026-10-06 | Documentación Swagger disponible en el microservicio. |
| **5** | **Filtros de pantalla aplicados en exportación** | §8 | Si el usuario filtra por fecha, organización o especialidad en pantalla, el reporte exportado respeta exactamente esos mismos filtros. | **Cumple** | `TreatmentController.java#L35-49`: Parámetros `organizationId`, `specialtyId` y rango de fechas enviados desde Angular al endpoint. | Fullstack Dev | 2026-10-06 | Consistencia garantizada entre lo que se ve en la grilla y lo que se descarga. |
| **6** | **Botón deshabilitado sin datos y spinner** | §7 | El botón de exportación se deshabilita si no hay registros o si ya se está procesando la descarga, mostrando un indicador de carga. | **Cumple** | Componente UI en `vg-web-sigrc`: `[disabled]="loading || count === 0"`, con spinner de carga visible. | Frontend Team | 2026-10-06 | Evita descargas vacías y peticiones dobles por clic reiterado. |
| **7** | **Fechas ISO-8601 y montos con BigDecimal** | §5, §8 | Los importes de dinero usan estrictamente `BigDecimal` (nunca float ni double para evitar fallas de coma flotante). Las fechas se manejan en formato ISO-8601 UTC. | **Cumple** | `TherapyPrice.java#L7-20`: `private BigDecimal price;`, `LocalDateTime.now(ZoneOffset.UTC)`. `TicketApplicationService.java#L30`: formato ISO. | Backend Lead | 2026-10-06 | Cero uso de `float` o `double` en importes financieros. |
| **8** | **PDF con encabezado institucional y pie de página** | §7, §14 | Los comprobantes PDF incluyen membrete oficial (logo Valle Grande, datos de sede) y numeración `Página X de Y` con fecha de emisión. | **Cumple** | Plantilla de comprobantes clínicos: logo institucional, datos de la clínica y paginación en el pie de página. | Design / Front | 2026-10-06 | Verificado en `Comprobante_TRAT-2026-09-27-000005.pdf`. |
| **9** | **Descarga protegida con JWT y permisos** | §11, §13 | La generación de reportes sensibles valida el token Bearer JWT con roles autorizados (`ADMIN`, `DOCTOR`, etc.). | **Cumple** | `SecurityConfig.java#L55-59`: configuración de Resource Server JWT. `TreatmentController.java#L30`: `@PreAuthorize("hasAnyRole(...)")`. | Security Lead | 2026-10-06 | Endpoints protegidos en WebFlux Security con validación de roles. |
| **10** | **Backend reactivo sin motores bloqueantes** | §2, §5 | El backend no debe usar motores monolíticos y bloqueantes (JasperReports, Apache POI bloqueando el hilo de Netty). Las exportaciones tabulares se manejan de forma reactiva o en cliente. | **Cumple** | `pom.xml`: microservicio 100% reactivo con Spring WebFlux y Reactor. Sin dependencias de Jasper ni Apache POI. | Arquitecto PRS04 | 2026-10-06 | No bloquea el EventLoop de Netty bajo concurrencia. |
| **11** | **Patrón de generación definido (A, B o C)** | §16 | Se tiene definido el patrón técnico: Patrón A (Cliente con plantillas JSON/Canvas) o Patrón C (Streaming reactivo de datos tabulares). | **Cumple** | Se seleccionó el Patrón A para comprobantes en cliente + Patrón C (Flux reactivo) para datos tabulares masivos. | Arquitecto PRS04 | 2026-10-06 | Documentado en las decisiones de diseño del PRS04. |
| **12** | **Pruebas automatizadas de controladores y DTOs** | §18 | Existen pruebas unitarias e integración que validan controladores, contratos DTO y generación de correlativos. | **Cumple** | `src/test/java/.../TreatmentControllerTest.java`, pruebas de secuencias `TicketSequence` integradas en el pipeline de GitLab CI. | QA Lead | 2026-10-06 | Pruebas automáticas ejecutadas en el pipeline CI/CD. |

---

## 3. Resumen de Calificación

```text
--------------------------------------------------------------------
Total de criterios evaluados : 12
Criterios que cumplen        : 12 (100.0%)
Criterios que no cumplen     : 0  (0.0%)
Criterios pendientes         : 0  (0.0%)
--------------------------------------------------------------------
% Cumplimiento Global        : 100.0% [Aprobado para Producción]
% Avance de Revisión         : 100.0% [Auditoría Finalizada]
Resultado                    : CUMPLE EL ESTÁNDAR
--------------------------------------------------------------------
```

---

## 4. Pasos antes de enviar un Pull Request (PR)

1. **Revisar el código local:** Verificar que la funcionalidad de reporte cumpla con los 12 puntos de esta lista.
2. **Correr el script de auditoría:** Ejecutar `python audit_agent.py` para actualizar `metricas_reportes.xlsx`.
3. **Incluir evidencias en el PR:** Pegar en la descripción del Merge Request en GitLab las referencias de las pruebas o archivos modificados.
