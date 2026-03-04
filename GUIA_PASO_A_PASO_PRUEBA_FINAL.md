# Guía paso a paso: crear el archivo final Excel y hacer la prueba

Esta guía te lleva desde cero hasta una prueba real de cotización en menos de 1 hora.

## 1) Crear el archivo base
1. Abre Excel y crea un libro nuevo.
2. Guárdalo como: `JMVISION_Cotizaciones.xlsx`.
3. Crea estas hojas en este orden:
   - `CONFIG`
   - `CLIENTES`
   - `PRODUCTOS`
   - `COTIZACION`
   - `HISTORICO_COTIZACIONES`
   - `VENTAS`
   - `DASHBOARD` (opcional en la primera prueba)

## 2) Configurar parámetros globales (`CONFIG`)
Carga valores iniciales:
- Empresa
- IVA
- Moneda
- Días de vigencia
- Prefijos (`COT-JMV`, `VTA-JMV`)

> Recomendación: usa formato porcentaje para IVA y formato fecha para campos de fecha.

## 3) Cargar datos maestros
### `CLIENTES`
Ingresa al menos 3 clientes de prueba con `ID_Cliente` único (`CLI-0001`, `CLI-0002`, ...).

### `PRODUCTOS`
Ingresa al menos 10 productos reales o de prueba, incluyendo:
- SKU
- Nombre
- Categoría
- Unidad
- Precio
- IVA aplica
- Stock
- Estado

> Para tu caso, incluye varios productos que contengan la palabra **cámaras** en la descripción para validar búsqueda.

## 4) Armar la hoja de cotización (`COTIZACION`)
1. Crea el encabezado (No cotización, fecha, cliente, vigencia, asesor).
2. Crea la tabla de ítems desde la fila 12.
3. Aplica las fórmulas de autocompletado (`BUSCARX`) desde la guía principal.
4. Agrega la búsqueda por palabra clave:
   - Celda `J4`: término de búsqueda (ej. `camaras`).
   - Celda `J6`: fórmula para devolver SKU sugerido.
   - Columna SKU (`A12`): toma el SKU sugerido para autocompletar el resto.

## 5) Configurar validaciones y formato
1. Crea listas desplegables para:
   - ID de cliente
   - SKU
   - Estado de cotización
2. Aplica formato condicional:
   - Vencida en rojo
   - Aprobada en verde
3. Bloquea celdas con fórmulas y deja editables solo campos de captura.

## 6) Prueba funcional mínima (obligatoria)
Haz estas 5 pruebas rápidas:

1. **Búsqueda por palabra**
   - Escribe `camaras` en `J4`.
   - Verifica que aparezca SKU en `J6`.

2. **Autocompletado de producto**
   - Verifica que se complete descripción, precio, IVA, categoría, unidad, stock y estado.

3. **Cálculo de totales**
   - Cambia cantidad y descuento.
   - Verifica subtotal, IVA y total general.

4. **Cliente y vigencia**
   - Cambia ID cliente.
   - Verifica nombre/correo autocompletados y vigencia correcta.

5. **Trazabilidad**
   - Registra la cotización en `HISTORICO_COTIZACIONES`.
   - Marca estado `Enviada`.

## 7) Prueba piloto real (día 1)
1. Crea 3 cotizaciones reales o simuladas.
2. Registra todas en histórico.
3. Marca una como `Aprobada` y pásala a `VENTAS`.
4. Revisa si el flujo fue claro y toma notas de mejora.

## 8) Criterio de éxito para darlo “listo”
Tu Excel queda listo cuando:
- No aparecen errores `#N/A` en flujo normal.
- La búsqueda por palabra funciona (ej. `camaras`).
- Los totales coinciden manualmente.
- Puedes pasar una cotización aprobada a venta sin reprocesar datos.

---

Si quieres, el siguiente ajuste recomendado es crear una versión **v1.1** con botones (macros simples) para:
- “Guardar cotización en histórico”
- “Convertir a venta”
- “Limpiar formulario de cotización”
