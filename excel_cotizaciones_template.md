# Estructura recomendada para tu Excel de cotizaciones y ventas

Esta guía está pensada para que tengas un archivo ordenado, escalable y fácil de usar.

---

## 1) Hoja: `CONFIG`
Usa esta hoja para parámetros globales.

### Campos sugeridos
| Celda | Descripción | Ejemplo |
|---|---|---|
| B2 | Nombre de la empresa | Mi Empresa S.A.S |
| B3 | NIT / ID fiscal | 900123456 |
| B4 | Teléfono | +57 300 000 0000 |
| B5 | Correo | ventas@miempresa.com |
| B6 | Moneda | COP |
| B7 | IVA (%) | 19% |
| B8 | Validez cotización (días) | 15 |

---

## 2) Hoja: `CLIENTES`
Registro maestro de clientes.

### Columnas sugeridas
| Columna | Nombre |
|---|---|
| A | ID_Cliente |
| B | Nombre_Cliente |
| C | Documento |
| D | Teléfono |
| E | Correo |
| F | Dirección |
| G | Ciudad |
| H | Observaciones |

> Tip: usa un ID único tipo `CLI-0001`.

---

## 3) Hoja: `PRODUCTOS`
Catálogo de productos/servicios.

### Columnas sugeridas
| Columna | Nombre |
|---|---|
| A | SKU |
| B | Producto |
| C | Categoría |
| D | Precio_Unitario |
| E | IVA_Aplica (SI/NO) |
| F | Stock |
| G | Activo (SI/NO) |

> Tip: usa formato moneda en `Precio_Unitario`.

---

## 4) Hoja: `COTIZACION`
Formato para generar cada cotización.

## Encabezado sugerido
| Campo | Celda |
|---|---|
| N° Cotización | B2 |
| Fecha | B3 |
| ID Cliente | B4 |
| Nombre Cliente (autocompletar) | B5 |
| Validez hasta | B6 |
| Vendedor | B7 |

### Fórmulas útiles en encabezado
- **Nombre cliente (B5)**
  ```excel
  =SI.ERROR(BUSCARX(B4,CLIENTES!A:A,CLIENTES!B:B),"Cliente no encontrado")
  ```
- **Validez hasta (B6)**
  ```excel
  =B3+CONFIG!B8
  ```

## Tabla de ítems (desde fila 12)
| Columna | Nombre | Ejemplo de fórmula |
|---|---|---|
| A | SKU | Selección manual/lista |
| B | Descripción | `=SI.ERROR(BUSCARX(A12,PRODUCTOS!A:A,PRODUCTOS!B:B),"")` |
| C | Cantidad | valor manual |
| D | Precio Unitario | `=SI.ERROR(BUSCARX(A12,PRODUCTOS!A:A,PRODUCTOS!D:D),0)` |
| E | Descuento % | valor manual |
| F | Subtotal | `=C12*D12*(1-E12)` |
| G | IVA | `=SI(SI.ERROR(BUSCARX(A12,PRODUCTOS!A:A,PRODUCTOS!E:E),"NO")="SI",F12*CONFIG!B7,0)` |
| H | Total línea | `=F12+G12` |

Copia las fórmulas hacia abajo para todas las filas de ítems.

## Resumen de totales
| Campo | Fórmula sugerida |
|---|---|
| Subtotal General | `=SUMA(F12:F100)` |
| IVA Total | `=SUMA(G12:G100)` |
| Total Cotización | `=SUMA(H12:H100)` |

---

## 5) Hoja: `HISTORICO_COTIZACIONES`
Control de todas las cotizaciones emitidas.

### Columnas sugeridas
| Columna | Nombre |
|---|---|
| A | N_Cotizacion |
| B | Fecha |
| C | ID_Cliente |
| D | Cliente |
| E | Subtotal |
| F | IVA |
| G | Total |
| H | Estado (Enviada/Aprobada/Rechazada/Vencida) |
| I | Fecha_Respuesta |
| J | Observaciones |

---

## 6) Hoja: `VENTAS`
Para convertir cotizaciones aprobadas en ventas.

### Columnas sugeridas
| Columna | Nombre |
|---|---|
| A | N_Venta |
| B | Fecha_Venta |
| C | N_Cotizacion |
| D | Cliente |
| E | Total_Venta |
| F | Medio_Pago |
| G | Estado_Pago |

---

## 7) Automatizaciones recomendadas en Excel
1. **Validación de datos**
   - Lista desplegable para `ID_Cliente`, `SKU` y `Estado`.
2. **Formato condicional**
   - Resaltar cotizaciones vencidas en rojo.
   - Estado `Aprobada` en verde.
3. **Bloqueo de fórmulas**
   - Protege celdas de cálculo y deja editables solo las de captura.
4. **Panel de indicadores (opcional)**
   - Cotizaciones del mes.
   - Tasa de aprobación.
   - Ventas cerradas.

---

## 8) Convención de numeración recomendada
- Cotización: `COT-AAAA-0001`
- Venta: `VTA-AAAA-0001`

Ejemplo:
- `COT-2026-0001`
- `VTA-2026-0001`

---

## 9) Flujo sugerido
1. Registrar cliente (si no existe).
2. Crear cotización en hoja `COTIZACION`.
3. Guardar resumen en `HISTORICO_COTIZACIONES`.
4. Cambiar estado cuando el cliente responda.
5. Si aprueba, registrar en `VENTAS`.

---

Con esta estructura tendrás un archivo profesional para cotizar y controlar ventas sin perder trazabilidad.
