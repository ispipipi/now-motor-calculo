# NOW Motor de Calculo

Motor local y modular para reproducir el calculo de variables remuneracionales NOW Fase 1.

## Alcance actual

- Lee las fuentes Excel reales desde rutas locales.
- Calcula RGU por tecnico desde hoja `RGU`.
- Aplica trabajos especiales GSA/Avanzadas como complemento diario hasta 4,5 RGU cuando no viene `Pts`.
- Calcula productividad, calidad, llave de tabla doble entrada, valor 100% movil y valor pago por asistencia.
- Mantiene Calidad como modulo separado que alimenta el calculo de bono.
- Genera conciliacion de Calidad Reactiva a 30 dias desde `BBDD` contra la hoja `CALIDAD` usada en julio 2026.
- Calcula y valida variables secundarias separadas: combustible y concurso/HHEE.
- Extrae programaticamente la matriz RM/ZN desde `Ratios_2`.
- Compara contra `Resumen Red Neutra RM` y `Resumen Red Neutra ZN`.
- Exporta resultados y diferencias a `outputs/<periodo>/`.

## Fuentes sensibles

No copies los Excel reales dentro del proyecto. Pasalos por CLI:

```bash
npm run validate -- \
  --bbdd "/Users/julioespinoza/Downloads/BBDD Julio 2026_NOW (1).xlsx" \
  --pago "/Users/julioespinoza/Downloads/Pago TAC Jul_2026_Rev Nuevo modelo (2).xlsx" \
  --novedades "/Users/julioespinoza/Downloads/Novedades Julio 2026 (1).xlsx" \
  --calidad "/Users/julioespinoza/Downloads/Calidad Junio-Julio 2026.xlsx" \
  --validacion-now-calidad "/ruta/validacion_now_calidad_2026-07.csv" \
  --period 2026-07 \
  --quality-period 2026-06
```

Pedro confirmo que los 3 ajustes RGU pendientes son ajustes manuales que no debiesen existir. El motor no debe aplicarlos para cuadrar el Ranking.

## Salidas

- `resultados_mensuales.json`: resultado calculado por tecnico.
- `validacion_resumen.json`: resumen de calce contra Excel real.
- `validacion_detalle.csv`: diferencia por tecnico y campo.
- `variables_secundarias_resumen.json`: totales de combustible y concurso/HHEE.
- `validacion_variables_secundarias.csv`: diferencia por tecnico y campo secundario.
- `referencias_llaves_observadas.csv`: control de llaves observadas RM/ZN contra `Ratios_2`.
- `ajustes_especiales_detalle.csv`: trazabilidad de OT Avanzada y GSA aplicados, duplicados o no mapeados.
- `calidad_reactiva_resumen.json`: resumen general y por tecnico de calidad reactiva a 30 dias.
- `calidad_reactiva_detalle.csv`: actividades origen evaluables y estado de reiteracion/maduracion.
- `calidad_reactiva_reiteraciones.csv`: reparaciones asociadas a una actividad origen.
- `calidad_reactiva_pareto_cierre.csv`: Pareto de codigos de cierre que generan reiteracion.
- `calidad_reactiva_vs_bono.csv`: conciliacion por tecnico entre reglas calculadas desde `BBDD` y resultado `CALIDAD` usado por bono.
- `calidad_fuente_vs_bono.csv`: conciliacion entre la planilla fuente de calidad junio-julio y la hoja `CALIDAD` usada por bono.
- `docs/formato-validacion-now-calidad.md`: propuesta de archivo para registrar ajustes/apelaciones de ejecutivos NOW contra la calidad que tiene Claro.
- `calidad_validacion_now_resumen.json`: estado de iteraciones y aplicacion de ajustes NOW.
- `calidad_validacion_now_detalle.csv`: trazabilidad por OT de cada ajuste, iteracion y resultado.
- `docs/gantt-proyecto-now.md`: Gantt viva del motor y de la futura plataforma modular.
- `data/assumptions/julio-2026-ajustes-pedro-pendientes.json`: evidencia de ajustes manuales rechazados por Pedro. El archivo esta deshabilitado con `enabled: false`.

## Notas tecnicas

- La Fase 1 no incluye interfaz.
- Combustible replica las columnas finales de `Variables`: `Asignacion de combustible y TAG + Adicional = Total`, `Total - Consumo = Saldo`, `Movilizacion especial` si el saldo es positivo y `Eficiencia de consumo` si el saldo es negativo. Las columnas `Suma de...` quedan como auditoria cruda.
- Concurso/HHEE se normaliza desde las columnas finales de `Variables`.
- Capacidad ociosa queda fuera de alcance por el momento. La hoja `Cap_Ociosa` no se carga ni se usa en calculo, resultado mensual o validacion secundaria.
- Calidad es un modulo separado. El pago de un mes usa la calidad del mes anterior: pago julio usa calidad junio; pago agosto usa calidad julio.
- La calidad vence 30 dias despues del cierre del mes evaluado. Calidad junio vence el 30 de julio y alimenta pago julio. Calidad julio vence el 30 de agosto y alimenta pago agosto.
- El CLI asume por defecto que `--quality-period` es el mes anterior a `--period`. Para una corrida de pago julio se debe usar `--period 2026-07 --quality-period 2026-06`; si se omite `--quality-period`, el motor lo infiere como `2026-06`.
- En pago julio 2026 el bono sigue alimentandose desde la hoja `CALIDAD`, porque es la fuente que cuadra contra el pago observado. Las reglas de Calidad Reactiva se calculan desde `BBDD` y se concilian contra esa fuente hasta lograr el mismo resultado.
- La version calculada desde `BBDD` aplica: actividades completadas, universo valido, misma `Direccion` exacta como llave de cruce, regla secuencial de ultima actividad origen, deduplicacion por `Fecha + OT`, excepcion `REITERADO_0D_OBSERVAR` para modificaciones y maduracion de ventana. Ya no exige `Access ID`, porque esos datos no estan disponibles en la BBDD por proteccion de informacion.
- La planilla `Calidad Junio-Julio 2026.xlsx` se puede pasar con `--calidad`. El motor no reemplaza automaticamente la hoja `CALIDAD` final del bono; genera `calidad_fuente_vs_bono.csv` para revisar denominador y numerador antes de activarla como fuente primaria.
- Calidad tiene dos capas: calculo por reglas y ajuste/apelacion de ejecutivos NOW contra la calidad que tiene Claro. Si la hoja `CALIDAD` trae `Reiterados - apelacion`, el bono usa ese numerador final; si no viene, el motor calcula `Reiterados - Apelacion`.
- Pedro confirmo que la validacion de Calidad debe permitir al menos dos iteraciones. La fuente operativa propuesta para el bono es motor + validacion NOW; la hoja `CALIDAD` final se mantiene como control de conciliacion.
- Los ajustes de Calidad se registraran preferentemente por `OT`, que sera la llave principal. El `TAC` quedara en una columna separada para analizar variaciones. El detalle de base debe conservarse dia a dia con campos como `Tipo de Actividad`, `Tipo de red` y `Codigo de Cierre`.
- En preliquidacion pueden cambiar tanto el numerador como el denominador, pero son excepciones que deben originarse en una gestion de jefatura y quedar trazables.
- Solo se consideran tareas `Completada`. En procesos manuales como GSA, si la orden corresponde pero aparece como no realizada, primero debe corregirse manualmente a `Completada` y luego incorporarse.
- GSA/Avanzadas quedan automatizadas con la regla confirmada por Pedro: la fuente ideal debe traer `RUT`, `Fecha`, `OT` y `Pts`; si `Pts` no viene, el motor calcula `max(0, 4,5 - RGU del tecnico en esa fecha)`.
- Si `Pts` viene, el motor usa ese valor de fuente. En julio la hoja GSA trae `Pts` pero no trae fecha util, por lo que esos puntos se aplican como fuente y el detalle queda en `ajustes_especiales_detalle.csv`.
- Si una fila GSA/Avanzadas llega sin `Pts` ni `Fecha`, debe completarse antes de la corrida. El motor no inventa ese valor; queda pendiente definir si la correccion sera manual o vendra desde otra fuente autorizada.
- Fallback manual de `Tabla` se aplica cuando una orden completada trae `Total RGU = 0`. Pedro confirmo que es una contingencia momentanea: primero se pide actualizacion del bot y, si no vuelve valorizada, se deja el valor base de la tabla.

## Estado validacion julio 2026

Ultima corrida local:

```text
totalResultados: 69
totalTecnicosValidados: 69
referencias llaves observadas: 46/46
totalRgu: 62/69
productividadNominal: 62/69
factorCalidad: 69/69
llave: 65/69
valor100Movil: 65/69
valorPagoAsistencia: 65/69
variables secundarias: 11/11 campos al 100%
calidad reactiva: conciliacion generada contra CALIDAD usada por bono
```

Las diferencias quedan detalladas en `outputs/2026-07/validacion_detalle.csv`.

Residuos RGU pendientes de auditoria con regla diaria GSA/Avanzadas:

```text
20337158-6 RM: motor 71.25 vs Excel 80.25 (-9), cambia llave/valor
21243461-2 RM: motor 93.75 vs Excel 102.75 (-9), cambia llave/valor
21440718-3 RM: motor 68.5 vs Excel 77 (-8.5), cambia llave/valor
17159419-7 RM: motor 85 vs Excel 94 (-9), cambia llave/valor
12717769-4 ZN: motor 82.75 vs Excel 83.75 (-1), no cambia llave/valor
15009906-4 ZN: motor 111.5 vs Excel 112.5 (-1), no cambia llave/valor
18898707-9 ZN: motor 75.75 vs Excel 77.75 (-2), no cambia llave/valor
```

Pedro confirmo que los residuos previos de 21440718-3, 12717769-4 y 15009906-4 eran ajustes manuales que no debiesen existir. Las nuevas diferencias RM aparecen al aplicar la regla diaria correcta sobre una fuente historica de julio que no trae todos los campos diarios necesarios para reproducir el Excel observado.

Las respuestas quedaron consolidadas en `docs/respuestas-pedro.md`.

## Ajustes manuales rechazados

El archivo de ajustes se conserva solo como evidencia de auditoria:

Archivo actual:

```text
data/assumptions/julio-2026-ajustes-pedro-pendientes.json
```

Esta marcado con `enabled: false`, por lo que aunque se entregue por CLI con `--ajustes`, el motor no aplica esos puntos.

Pendientes de aclaracion con Pedro:

- Capacidad ociosa queda excluida hasta que se decida retomarla.
- Para cuadrar `CALIDAD` contra bono se requiere una fuente que incluya el mes de calidad y sus reparaciones hasta el vencimiento. Para pago julio se necesita calidad junio observada hasta el 30 de julio; para pago agosto se necesita calidad julio observada hasta el 30 de agosto.
- En la planilla `Calidad Junio-Julio 2026.xlsx`, el denominador de junio cuadra contra el pago de julio, pero el numerador no coincide completamente con la hoja `CALIDAD` final. Eso corresponde a la capa posterior de validacion/apelacion de ejecutivos NOW contra la calidad que tiene Claro; queda conciliado en `calidad_fuente_vs_bono.csv` mediante columnas de ajuste neto.
- El formato de validacion NOW fue revisado con Pedro. Falta implementarlo como entrada formal, incluyendo iteraciones, ajustes por OT, TAC asociado, detalle diario y trazabilidad de jefatura.
- Falta definir el procedimiento exacto para completar filas GSA/Avanzadas sin `Pts` ni `Fecha`.

## Riesgo dependencia

`xlsx@0.18.5` es la dependencia minima usada para leer Excel en Node. `npm install` reporta una vulnerabilidad alta conocida en el arbol de esa libreria. Para esta fase local, con archivos controlados, se mantiene por simplicidad; antes de exponer carga de archivos en una plataforma, conviene reemplazarla o aislarla.
