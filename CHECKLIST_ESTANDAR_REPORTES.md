# INSTRUMENTO DE VERIFICACIÓN: CHECKLIST DE CUMPLIMIENTO DEL ESTÁNDAR
## Comité de Estándar: Generación de Reportes y Exportación de Datos (PDF / Excel / CSV)
**Proyecto PRS:** PRS04 - SIGRC (Sistema Integrado de Gestión y Registro Clínico)  
**Componentes Evaluados:** `vg-ms-therapies` (Backend Spring Boot WebFlux) & `vg-web-sigrc` (Frontend Angular)  
**Versión del Estándar:** 2.4.0 | **Código de Documento:** CK-EST-REP-PRS04  

---

## 1. Presentación y Propósito del Instrumento

El presente **CheckList de Verificación de Cumplimiento del Estándar de Reportes y Exportación** tiene como finalidad servir como instrumento formal de auditoría y aseguramiento de la calidad de software para el PRS04. Permite validar que todo componente que genere, exporte o procese información en formatos descargables (PDF, Excel, CSV) cumpla estrictamente con las directrices de arquitectura, rendimiento, seguridad, estandarización de datos y experiencia de usuario establecidas por el Comité de Estándares.

### Reglas de Calificación del Checklist:
- **Cumple (C):** El ítem ha sido implementado satisfactoriamente y cuenta con evidencia técnica verificable (código, pruebas, configuración). Suma al porcentaje de cumplimiento.
- **No Cumple (NC):** El ítem aplica al componente pero no respeta las pautas del estándar. Se contabiliza y reduce el porcentaje.
- **Pendiente (P):** El ítem está en proceso de revisión o implementación pendiente. No penaliza directamente el cumplimiento pero reduce el avance total de revisión.
- **No Aplica (NA):** El criterio no corresponde a la naturaleza técnica del componente auditado. Se excluye de los cálculos.

**Fórmulas Oficiales:**
$$\% \text{ Cumplimiento} = \frac{\text{Cumple}}{\text{Cumple} + \text{No Cumple}} \times 100$$
$$\% \text{ Avance de Revisión} = \frac{\text{Cumple} + \text{No Cumple}}{\text{Total Ítems Aplicables}} \times 100$$

---

## 2. Matriz de Verificación del Estándar (Checklist de PR y Release)

| # | Ítem de Verificación | Ref. Estándar | Criterio de Aceptación / Verificación Técnica | Estado | Evidencia / Ubicación de Código | Responsable | Fecha | Observaciones |
|---|----------------------|---------------|-----------------------------------------------|:------:|----------------------------------|-------------|:-----:|---------------|
| **1** | **Formato elegido según la audiencia y justificado** | §3 / 02-formatos | La elección de formato responde a la necesidad funcional: PDF para comprobantes y visualización oficial; Excel para auditorías tabulares masivas; CSV para interoperabilidad analítica. | **Cumple** | `TherapyController.java`, Comprobantes de tratamiento clínicos emitidos en PDF para pacientes y reportes tabulares de ventas. | Comité Estándar PRS04 | 2026-10-06 | Justificado: Comprobantes clínicos requieren inmutabilidad (PDF); auditorías requieren Excel. |
| **2** | **Lazy import de jsPDF / xlsx (cero bloat en bundle inicial)** | §5 · §7 | En el frontend, las librerías pesadas (`jspdf`, `xlsx`, `exceljs`) no deben estar en el bundle principal de la aplicación. Deben cargarse mediante `import()` dinámico solo al accionar el botón de exportación. | **Cumple** | `vg-web-sigrc/angular.json` y dynamic imports `import('jspdf').then(...)`. Bundle vendor inicial optimizado sin dependencias de reporte precargadas. | Front Lead PRS04 | 2026-10-06 | Se valida mediante análisis de bundle chunk; las bibliotecas pesadas se aíslan en chunks asíncronos bajo demanda. |
| **3** | **Nombre de archivo con el patrón definido** | §6.4 | La nomenclatura de los archivos generados debe seguir estrictamente la convención de prefijo de negocio, fecha ISO y correlativo: `[Tipo]_[PREFIJO]-[YYYY-MM-DD]-[CORRELATIVO].[ext]`. | **Cumple** | `TicketApplicationService.java#L43`: `String.format("%s-%s-%06d", PREFIX, today, seq.getSequence())`. Archivos emitidos: `Comprobante_TRAT-2026-09-27-000005.pdf`. | Backend Lead PRS04 | 2026-10-06 | Cumple al 100%. Generación determinística con correlativo incremental por organización en MongoDB. |
| **4** | **Endpoint fuente documentado** | §8 / 05-operaciones | Todo endpoint de datos que alimente la exportación debe estar documentado en OpenAPI 3.0 / Swagger, indicando query params, filtros y esquema de respuesta. | **Cumple** | `SecurityConfig.java#L32-34` permite `/v3/api-docs/**` y `/swagger-ui/**`. Controladores documentados con tags de API en endpoints `/api/v1/therapies/**`. | Backend Dev | 2026-10-06 | Endpoints reactivos WebFlux expuestos en la especificación OpenAPI institucional. |
| **5** | **Filtros de pantalla reflejados en la exportación** | §8 | Si el usuario aplica filtros en la interfaz (fechas, sucursal, estado, especialidad), el reporte generado debe respetar exactamente los mismos filtros aplicados en la vista actual. | **Cumple** | `TreatmentController.java#L35-49` y frontend query params sincronizados: filtros de `organizationId`, `specialtyId`, rango de fechas propagados a la consulta. | Fullstack Dev | 2026-10-06 | Garantiza consistencia entre lo visualizado en la grilla y el reporte descargado. |
| **6** | **Botón deshabilitado sin datos; spinner durante la generación** | §7 | Prevención de errores de usuario: el botón de exportación debe inhabilitarse si el dataset está vacío, y mostrar un indicador visual (spinner/skeleton) con estado disabled mientras se compila el reporte. | **Cumple** | Componente UI `vg-web-sigrc/src/app/shared/components/ui/button`: estado `[disabled]="loading \|\| count === 0"`, con spinner reactivo integrado. | Frontend Team | 2026-10-06 | Evita descargas vacías y peticiones dobles concurrentes por clic múltiple. |
| **7** | **Fechas ISO-8601 y montos sin Double/Float** | §5 · §8 | Estricta precisión numérica: montos de dinero deben manejar exclusivamente `BigDecimal` (Java) / enteros en céntimos (evitar fallas de coma flotante de `double`/`float`). Fechas formateadas en estándar `ISO-8601` (`YYYY-MM-DD` / UTC). | **Cumple** | `TherapyPrice.java#L7-20`: `private BigDecimal price;`, `LocalDateTime.now(ZoneOffset.UTC)`. `TicketApplicationService.java#L30`: `DateTimeFormatter.ISO_LOCAL_DATE`. | Backend Lead PRS04 | 2026-10-06 | Verificado en código de dominio y DTOs de request/response. Cero uso de `float`/`double` para finanzas. |
| **8** | **PDF externo con encabezado institucional y pie de paginación** | §7 · §14 | Los reportes PDF deben incluir cabecera corporativa estandarizada (logo Valle Grande, RUC, datos del paciente/sede) y pie de página con paginación formal `Página X de Y` y fecha/hora de emisión. | **Cumple** | Plantilla de comprobantes y reportes clínicos: logo institucional, membrete oficial, paginación dinámica y fecha de generación auditada. | Front/Design Team | 2026-10-06 | Verificado en comprobantes emitidos: `Comprobante_TRAT-2026-09-27-000005.pdf`. |
| **9** | **Descarga protegida con JWT y permiso reports:export** | §11 · §13 | La generación y descarga de información analítica sensible debe validar token Bearer JWT con el rol o authority correspondiente (`reports:export` o roles autorizados RBAC). | **Cumple** | `SecurityConfig.java#L55-59`: validación de OAuth2 Resource Server JWT. `TreatmentController.java#L30`: `@PreAuthorize("hasAnyRole('ADMIN', 'DOCTOR', 'USER', ...)")`. | Security Lead | 2026-10-06 | Endpoints protegidos en WebFlux Security con claims de autenticación y aislamiento multi-tenant. |
| **10** | **Server-side solo con justificación y sin motores prohibidos** | §2 · §5 | El servidor backend no debe utilizar motores de renderizado pesado y bloqueante (JasperReports, iText no reactivo, Apache POI en el thread reactivo de Netty). Generación masiva delegada a colas o procesada en cliente. | **Cumple** | `pom.xml`: microservicio 100% no-bloqueante con Spring WebFlux y Project Reactor. Sin dependencias de motores monolíticos prohibidos en el pipeline reactivo. | Arquitecto PRS04 | 2026-10-06 | Cumplimiento total. El pipeline reactivo de EventLoop no sufre degradación de I/O bloqueante. |
| **11** | **Patrón de plantilla/recursos elegido (A, B o C)** | §16 | Debe definirse formalmente el patrón de generación: Patrón A (Cliente con plantillas JSON/Canvas), Patrón B (Servicio de reportes asíncrono con Worker) o Patrón C (Streaming reactivo de datos tabulares). | **Cumple** | Patrón A seleccionado para comprobantes interactivos del cliente + Patrón C (Reactive Flux streaming) para exportaciones masivas tabulares. | Arquitecto PRS04 | 2026-10-06 | Patrón formalizado en el documento de arquitectura de software de PRS04. |
| **12** | **Servicio de reporte con pruebas de estructura y naming** | §18 | La suite de pruebas automatizadas debe validar que el correlativo, los nombres de archivos, los headers HTTP (`Content-Disposition`) y los contratos de datos se cumplan en CI/CD. | **Cumple** | `src/test/java/.../TreatmentControllerTest.java`, pruebas de secuencias `TicketSequence` y tests de DTOs con validaciones de campos no nulos. | QA Automation Lead | 2026-10-06 | Suite de pruebas automatizada ejecutada en el pipeline de GitLab CI (`.gitlab-ci.yml`). |

---

## 3. Resumen Ejecutivo del Cumplimiento

```
====================================================================
           PANEL DE CONTROL DE AUDITORÍA DE ESTÁNDARES PRS04
====================================================================
 Total Ítems Evaluados       : 12
 Ítems que Cumplen           : 12 (100.0%)
 Ítems que No Cumplen        : 0  (0.0%)
 Ítems Pendientes            : 0  (0.0%)
 Ítems No Aplica             : 0
--------------------------------------------------------------------
 % CUMPLIMIENTO GLOBAL       : 100.0% [APROBADO PARA PRODUCCIÓN]
 % AVANCE DE REVISIÓN        : 100.0% [AUDITORÍA FINALIZADA]
 ESTADO GLOBAL DE CALIDAD    : CUMPLE EL ESTÁNDAR
====================================================================
```

---

## 4. Instrucciones para los Desarrolladores antes de enviar un Pull Request (PR)

1. **Revisión Local:** Todo desarrollador debe contrastar su funcionalidad de reporte con los 12 puntos de esta lista.
2. **Ejecución del Agente Auditor:** Ejecutar el comando local de verificación con el Agente de IA (`python scripts/audit_agent.py`) para regenerar el reporte `metricas_reportes.xlsx`.
3. **Adjuntar Evidencia en el PR:** Pegar la tabla resumen del checklist y el enlace al commit de pruebas en la descripción del Merge Request en GitLab.
