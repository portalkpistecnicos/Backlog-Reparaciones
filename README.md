# Backlog Repara &middot; Zona Norte (Elecnor)

Dashboard diario del KPI **Backlog Repara**, publicado con GitHub Pages.

## Formula

```
Backlog Repara = En_Proceso2 / ( (Cerrado + Cancelado)_6d_habiles / 6 )
```

- **En Proceso2**: incidencias con estado `pendiente` en el corte del dia.
- **Divisor**: promedio diario de cierres (Cerrado + Cancelado, por fecha real de
  cierre) de los ultimos 6 dias habiles (lunes a sabado; domingo no cuenta).

## Factor FBR

Convierte el ratio Backlog Repara en un factor de desempeno entre 0,93 y 1,08:

```
FBR = 1,08                                    si BR <= 0,5
      0,93                                    si BR >= 1,7
      interpolacion lineal (1,08 -> 1,00)     si 0,5 < BR <= 1
      interpolacion lineal (1,00 -> 0,93)     si 1 <= BR < 1,7
```

## Actualizacion diaria

`update_dashboard.bat` (llama a `update_dashboard.ps1`):

1. Busca el CSV mas reciente `p67_base_backlog_reparaciones_mod-detalle_*.csv`
   en la carpeta de origen (KPI's/Backlog Repara).
2. Recalcula la serie diaria y la composicion del backlog actual.
3. Reemplaza el bloque `AUTO-DATA` dentro de `index.html`.
4. Hace commit y push al repositorio.

Para dejarlo corriendo solo cada dia, agrega `update_dashboard.bat` al
Programador de tareas de Windows (Task Scheduler) con la frecuencia deseada.

**Importante**: el CSV de origen contiene datos de clientes y nunca se sube a
este repositorio (ver `.gitignore`). Solo se publican los totales agregados
del KPI.
