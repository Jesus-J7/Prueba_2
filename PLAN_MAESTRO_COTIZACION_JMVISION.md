# PLAN MAESTRO DE COTIZACIÓN Y VENTAS — JMVISION

## 1. Objetivo general
Diseñar un archivo `JMVISION_Cotizaciones.xlsx` que permita:
- Cotizar productos y servicios de forma rápida y consistente.
- Mantener trazabilidad completa desde la cotización hasta la venta.
- Medir desempeño comercial (monto cotizado, tasa de aprobación y ventas cerradas).

---

## 2. Alcance funcional
El sistema en Excel cubrirá 4 procesos:
1. **Parametrización:** configuración de empresa, impuestos, vigencias y listas.
2. **Maestros:** clientes y productos.
3. **Operación:** creación de cotizaciones con cálculo automático.
4. **Seguimiento:** histórico, conversión a venta e indicadores.

---

## 3. Estructura de hojas (arquitectura del archivo)

## 3.1 `CONFIG`
Parámetros globales de la solución.

### Campos sugeridos
| Celda | Campo | Tipo | Ejemplo |
|---|---|---|---|
| B2 | Empresa | Texto | JMVISION |
| B3 | NIT/ID fiscal | Texto | 900123456 |
| B4 | Teléfono | Texto | +57 300 000 0000 |
| B5 | Correo ventas | Texto | ventas@jmvision.com |
| B6 | Moneda | Lista | COP |
| B7 | IVA estándar | Porcentaje | 19% |
| B8 | Días vigencia cotización | Número | 15 |
| B9 | Prefijo cotización | Texto | COT-JMV |
| B10 | Prefijo venta | Texto | VTA-JMV |

---

## 3.2 `CLIENTES`
Base maestra de clientes.

| Columna | Campo |
|---|---|
| A | ID_Cliente |
| B | Tipo_Cliente (Empresa/Persona) |
| C | Nombre_Cliente |
| D | Documento/NIT |
| E | Contacto |
| F | Teléfono |
| G | Correo |
| H | Ciudad |
| I | Dirección |
| J | Canal_Origen |
| K | Estado (Activo/Inactivo) |

**Regla:** `ID_Cliente` único (ejemplo `CLI-0001`).

---

## 3.3 `PRODUCTOS`
Catálogo comercial de JMVISION.

| Columna | Campo |
|---|---|
| A | SKU |
| B | Producto_Servicio |
| C | Categoría |
| D | Unidad |
| E | Precio_Lista |
| F | Costo_Referencial |
| G | Margen_%_Objetivo |
| H | IVA_Aplica (SI/NO) |
| I | Estado (Activo/Inactivo) |

**Regla:** no usar productos inactivos en cotización.

---

## 3.4 `COTIZACION`
Plantilla operativa para emitir cotizaciones.

### Encabezado
| Campo | Celda | Fórmula/Origen |
|---|---|---|
| No_Cotización | B2 | Secuencial con prefijo |
| Fecha | B3 | Manual / `=HOY()` |
| ID_Cliente | B4 | Lista desde `CLIENTES` |
| Cliente | B5 | `BUSCARX` |
| Correo | B6 | `BUSCARX` |
| Vigencia hasta | B7 | `=B3+CONFIG!B8` |
| Asesor | B8 | Manual |
| Moneda | B9 | `=CONFIG!B6` |

### Tabla de ítems (desde fila 12)
| Columna | Campo | Fórmula recomendada |
|---|---|---|
| A | SKU | Lista desde `PRODUCTOS` |
| B | Descripción | `=SI.ERROR(BUSCARX(A12,PRODUCTOS!A:A,PRODUCTOS!B:B),"")` |
| C | Cantidad | Manual |
| D | Precio_Unitario | `=SI.ERROR(BUSCARX(A12,PRODUCTOS!A:A,PRODUCTOS!E:E),0)` |
| E | Desc_% | Manual |
| F | Subtotal | `=C12*D12*(1-E12)` |
| G | IVA | `=SI(SI.ERROR(BUSCARX(A12,PRODUCTOS!A:A,PRODUCTOS!H:H),"NO")="SI",F12*CONFIG!B7,0)` |
| H | Total_Línea | `=F12+G12` |

### Totales
- Subtotal: `=SUMA(F12:F200)`
- IVA: `=SUMA(G12:G200)`
- Total: `=SUMA(H12:H200)`

---

## 3.5 `HISTORICO_COTIZACIONES`
Bitácora central.

| Columna | Campo |
|---|---|
| A | No_Cotización |
| B | Fecha |
| C | ID_Cliente |
| D | Cliente |
| E | Subtotal |
| F | IVA |
| G | Total |
| H | Estado |
| I | Fecha_Respuesta |
| J | Motivo_Pérdida |
| K | Observaciones |

Estados sugeridos: `Borrador`, `Enviada`, `Aprobada`, `Rechazada`, `Vencida`.

---

## 3.6 `VENTAS`
Conversión de cotizaciones aprobadas.

| Columna | Campo |
|---|---|
| A | No_Venta |
| B | Fecha_Venta |
| C | No_Cotización |
| D | Cliente |
| E | Total_Venta |
| F | Medio_Pago |
| G | Estado_Pago |
| H | Fecha_Cobro |

---

## 3.7 `DASHBOARD`
Indicadores de gestión comercial.

KPIs mínimos:
- Cotizado mes actual.
- Número de cotizaciones emitidas.
- Tasa de aprobación (% aprobadas / enviadas).
- Ventas cerradas del mes.
- Ticket promedio.

Gráficos sugeridos:
- Embudo: Enviada → Aprobada → Venta.
- Barras por asesor.
- Tendencia mensual de cotización/venta.

---

## 4. Reglas de negocio
1. Una cotización debe tener mínimo un ítem con cantidad > 0.
2. La vigencia nunca puede ser menor a la fecha de emisión.
3. Solo cotizaciones `Aprobada` pasan a `VENTAS`.
4. Si pasa la vigencia sin respuesta, estado automático recomendado: `Vencida`.
5. Descuentos mayores al 20% requieren aprobación interna (control manual o color de alerta).

---

## 5. Estandarización y control de calidad
- Convertir todos los rangos en **Tablas de Excel** (`Ctrl + T`).
- Activar validación de datos para IDs, SKU y estados.
- Bloquear celdas con fórmulas y proteger hoja con contraseña.
- Usar formato condicional para:
  - vigencias vencidas,
  - descuentos altos,
  - estados críticos.

---

## 6. Numeración recomendada
- Cotización: `COT-JMV-AAAA-0001`
- Venta: `VTA-JMV-AAAA-0001`

Ejemplos:
- `COT-JMV-2026-0008`
- `VTA-JMV-2026-0003`

---

## 7. Flujo operativo (SOP)
1. Registrar/validar cliente en `CLIENTES`.
2. Crear cotización en `COTIZACION`.
3. Guardar registro en `HISTORICO_COTIZACIONES` con estado `Enviada`.
4. Hacer seguimiento comercial y actualizar estado.
5. Si es `Aprobada`, generar registro en `VENTAS`.
6. Revisar `DASHBOARD` semanalmente para decisiones.

---

## 8. Plan de implementación por fases

### Fase 1 (día 1)
- Crear hojas base.
- Cargar parámetros y catálogos iniciales.
- Construir formato `COTIZACION` con fórmulas.

### Fase 2 (día 2)
- Activar validaciones, formatos y protección.
- Crear histórico y módulo de ventas.

### Fase 3 (día 3)
- Construir dashboard y validar KPIs.
- Pruebas con 10 cotizaciones reales/simuladas.

### Fase 4 (día 4)
- Ajustes finales.
- Capacitación interna (30-45 min).

---

## 9. Checklist de salida a producción
- [ ] Catálogo de productos actualizado.
- [ ] Lista de clientes inicial cargada.
- [ ] Fórmulas probadas sin errores `#N/A`.
- [ ] Estados y listas desplegables funcionando.
- [ ] Protección de celdas activada.
- [ ] Dashboard mostrando datos correctos.

---

## 10. Resultado esperado
Con este plan maestro, JMVISION tendrá un archivo de cotizaciones robusto, con control operativo y visión comercial para convertir más cotizaciones en ventas.
