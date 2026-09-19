<#
  update_dashboard.ps1
  Recalcula el KPI "Backlog Repara" desde el CSV de origen y actualiza index.html
  (bloque AUTO-DATA: resumen de zona, tarjetas por agencia y explorador por bucket),
  luego hace commit + push al repositorio de GitHub.

  Ejecutar via update_dashboard.bat (doble clic o tarea programada diaria).
#>

$ErrorActionPreference = 'Stop'
$inv = [System.Globalization.CultureInfo]::InvariantCulture

# ---------------------------------------------------------------------------
# 0. Rutas
# ---------------------------------------------------------------------------
$RepoDir   = $PSScriptRoot
$IndexPath = Join-Path $RepoDir 'index.html'
$SourceDir = "G:\Mi unidad\Respaldo\Documents\KPI's\Backlog Repara"
$LogPath   = Join-Path $RepoDir 'update.log'

function Write-Log($msg) {
    $line = "[{0}] {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $msg
    Write-Host $line
    Add-Content -Path $LogPath -Value $line -Encoding UTF8
}

Write-Log "== Iniciando actualizacion =="

# ---------------------------------------------------------------------------
# 1. Ubicar el CSV mas reciente (por nombre de mes, no por fecha de modificacion:
#    un archivo de un mes anterior puede haberse tocado despues por sync de Drive)
# ---------------------------------------------------------------------------
$csv = Get-ChildItem -Path $SourceDir -Filter 'p67_base_backlog_reparaciones_mod-detalle_*.csv' |
       Sort-Object Name -Descending | Select-Object -First 1
if (-not $csv) { Write-Log "ERROR: no se encontro ningun CSV en $SourceDir"; exit 1 }
Write-Log "CSV fuente: $($csv.FullName)"

$data = Import-Csv -Path $csv.FullName -Delimiter ';' -Encoding UTF8
Write-Log "Filas leidas: $($data.Count)"

# ---------------------------------------------------------------------------
# 2. Entidades HTML / normalizacion de texto (los codigos del CSV vienen en
#    MAYUSCULAS sin tildes; se aplica una capa de nombres "bonitos" conocidos
#    y se cae a Title Case si aparece un codigo nuevo)
# ---------------------------------------------------------------------------
$prettyMap = @{
    'VALPARAISO'   = 'Valpara&iacute;so'
    'VINA DEL MAR' = 'Vi&ntilde;a del Mar'
    'SAN ANTONIO'  = 'San Antonio'
}

function ToTitleCase($s) {
    $ti = (Get-Culture).TextInfo
    return $ti.ToTitleCase($s.ToLower())
}

function HtmlEscape($s) {
    if ($null -eq $s) { return '' }
    $s = $s -replace '&','&amp;' -replace '<','&lt;' -replace '>','&gt;' -replace "'","&#39;"
    $s = $s -replace '\xE1','&aacute;' -replace '\xE9','&eacute;' -replace '\xED','&iacute;' `
             -replace '\xF3','&oacute;' -replace '\xFA','&uacute;' -replace '\xF1','&ntilde;' `
             -replace '\xC1','&Aacute;' -replace '\xC9','&Eacute;' -replace '\xCD','&Iacute;' `
             -replace '\xD3','&Oacute;' -replace '\xDA','&Uacute;' -replace '\xD1','&Ntilde;'
    return $s
}

function PrettyName($raw) {
    if ($prettyMap.ContainsKey($raw)) { return $prettyMap[$raw] }
    if ([string]::IsNullOrWhiteSpace($raw)) { return 'Sin dato' }
    return HtmlEscape (ToTitleCase $raw)
}

function Round2($v) { [math]::Round($v, 2) }

# ---------------------------------------------------------------------------
# 3. Cierres reales (por fecha_cierre), desglosados por territorio y estado
# ---------------------------------------------------------------------------
$closuresRaw = $data | Where-Object { $_.fecha_cierre -ne '' } |
    Select-Object rdy_id_incidencia, fecha_cierre, rdy_estado, rdy_cod_territorio -Unique

$closuresByDateZone            = @{}   # date -> count
$closuresByDateZoneEstado      = @{}   # date -> @{estado=count}
$closuresByDateTerritorio      = @{}   # date -> @{territorio=count}
$closuresByDateTerritorioEstado= @{}   # date -> @{territorio=@{estado=count}}

foreach ($row in $closuresRaw) {
    $d = $row.fecha_cierre; $t = $row.rdy_cod_territorio; $e = $row.rdy_estado
    if (-not $closuresByDateZone.ContainsKey($d)) { $closuresByDateZone[$d] = 0 }
    $closuresByDateZone[$d]++

    if (-not $closuresByDateZoneEstado.ContainsKey($d)) { $closuresByDateZoneEstado[$d] = @{} }
    if (-not $closuresByDateZoneEstado[$d].ContainsKey($e)) { $closuresByDateZoneEstado[$d][$e] = 0 }
    $closuresByDateZoneEstado[$d][$e]++

    if (-not $closuresByDateTerritorio.ContainsKey($d)) { $closuresByDateTerritorio[$d] = @{} }
    if (-not $closuresByDateTerritorio[$d].ContainsKey($t)) { $closuresByDateTerritorio[$d][$t] = 0 }
    $closuresByDateTerritorio[$d][$t]++

    if (-not $closuresByDateTerritorioEstado.ContainsKey($d)) { $closuresByDateTerritorioEstado[$d] = @{} }
    if (-not $closuresByDateTerritorioEstado[$d].ContainsKey($t)) { $closuresByDateTerritorioEstado[$d][$t] = @{} }
    if (-not $closuresByDateTerritorioEstado[$d][$t].ContainsKey($e)) { $closuresByDateTerritorioEstado[$d][$t][$e] = 0 }
    $closuresByDateTerritorioEstado[$d][$t][$e]++
}

$enProcesoByDateZone = @{}         # date -> count
$enProcesoByDateTerritorio = @{}   # date -> @{territorio=count}
foreach ($row in ($data | Where-Object { $_.rdy_estado -eq 'pendiente' })) {
    $d = $row.partition_date; $t = $row.rdy_cod_territorio
    if (-not $enProcesoByDateZone.ContainsKey($d)) { $enProcesoByDateZone[$d] = 0 }
    $enProcesoByDateZone[$d]++
    if (-not $enProcesoByDateTerritorio.ContainsKey($d)) { $enProcesoByDateTerritorio[$d] = @{} }
    if (-not $enProcesoByDateTerritorio[$d].ContainsKey($t)) { $enProcesoByDateTerritorio[$d][$t] = 0 }
    $enProcesoByDateTerritorio[$d][$t]++
}

$allDates = $data.partition_date | Where-Object { $_ -ne '' } | Select-Object -Unique |
    Sort-Object { [datetime]::ParseExact($_, 'dd-MM-yyyy', $null) }
$lastDate = $allDates[$allDates.Count - 1]
$lastDt   = [datetime]::ParseExact($lastDate, 'dd-MM-yyyy', $null)
Write-Log "Ultima fecha: $lastDate"

# Suma movil de cierres en los ultimos 6 dias habiles (lun-sab) terminando en $dt,
# para un hashtable por-fecha ($map, date -> count) opcionalmente filtrado por $key
# dentro de un hashtable anidado (date -> @{key=count}).
function Sum6Habiles($dt, $mapByDate, $key) {
    $sum = 0; $found = 0; $back = 0
    while ($found -lt 6 -and $back -lt 21) {
        $dd = $dt.AddDays(-$back)
        if ($dd.DayOfWeek -ne [System.DayOfWeek]::Sunday) {
            $k = $dd.ToString('dd-MM-yyyy')
            if ($mapByDate.ContainsKey($k)) {
                if ($null -eq $key) { $sum += $mapByDate[$k] }
                elseif ($mapByDate[$k].ContainsKey($key)) { $sum += $mapByDate[$k][$key] }
            }
            $found++
        }
        $back++
    }
    return $sum
}

function BacklogRatio($enProceso, $sum6) {
    $denom = $sum6 / 6.0
    if ($denom -le 0) { return $null }
    return [math]::Round($enProceso / $denom, 3)
}

function Fbr($br) {
    if ($null -eq $br) { return $null }
    if ($br -le 0.5) { return 1.08 }
    if ($br -ge 1.7) { return 0.93 }
    if ($br -gt 0.5 -and $br -le 1) { return [math]::Round(((1.08-1)*($br-1))/(0.5-1) + 1, 4) }
    return [math]::Round(((1-0.93)*($br-1.7)/(1-1.7)) + 0.93, 4)
}

# ---------------------------------------------------------------------------
# 4. Resumen de zona (corte = ultima fecha)
# ---------------------------------------------------------------------------
$zoneEnProceso = if ($enProcesoByDateZone.ContainsKey($lastDate)) { $enProcesoByDateZone[$lastDate] } else { 0 }
$zoneSum6      = Sum6Habiles $lastDt $closuresByDateZone $null
$zoneBr        = BacklogRatio $zoneEnProceso $zoneSum6
$zoneFbr       = Fbr $zoneBr
$zoneTerminadasHoy  = if ($closuresByDateZoneEstado.ContainsKey($lastDate) -and $closuresByDateZoneEstado[$lastDate].ContainsKey('Cerrado'))   { $closuresByDateZoneEstado[$lastDate]['Cerrado'] }   else { 0 }
$zoneCanceladasHoy  = if ($closuresByDateZoneEstado.ContainsKey($lastDate) -and $closuresByDateZoneEstado[$lastDate].ContainsKey('Cancelado')) { $closuresByDateZoneEstado[$lastDate]['Cancelado'] } else { 0 }

Write-Log "Zona: en_proceso=$zoneEnProceso backlog=$zoneBr fbr=$zoneFbr"

# ---------------------------------------------------------------------------
# 5. Tarjetas por agencia (territorio)
# ---------------------------------------------------------------------------
$territorios = $data.rdy_cod_territorio | Where-Object { $_ -ne '' } | Select-Object -Unique
$lastPendAll = $data | Where-Object { $_.partition_date -eq $lastDate -and $_.rdy_estado -eq 'pendiente' }

$agencies = @()
foreach ($t in $territorios) {
    $ep = if ($enProcesoByDateTerritorio.ContainsKey($lastDate) -and $enProcesoByDateTerritorio[$lastDate].ContainsKey($t)) { $enProcesoByDateTerritorio[$lastDate][$t] } else { 0 }
    $sum6 = Sum6Habiles $lastDt $closuresByDateTerritorio $t
    $br = BacklogRatio $ep $sum6
    $termHoy = if ($closuresByDateTerritorioEstado.ContainsKey($lastDate) -and $closuresByDateTerritorioEstado[$lastDate].ContainsKey($t) -and $closuresByDateTerritorioEstado[$lastDate][$t].ContainsKey('Cerrado')) { $closuresByDateTerritorioEstado[$lastDate][$t]['Cerrado'] } else { 0 }
    $cancHoy = if ($closuresByDateTerritorioEstado.ContainsKey($lastDate) -and $closuresByDateTerritorioEstado[$lastDate].ContainsKey($t) -and $closuresByDateTerritorioEstado[$lastDate][$t].ContainsKey('Cancelado')) { $closuresByDateTerritorioEstado[$lastDate][$t]['Cancelado'] } else { 0 }

    $pendTerr = @($lastPendAll | Where-Object { $_.rdy_cod_territorio -eq $t })
    $reiterCount = @($pendTerr | Where-Object { $_.rdy_es_reiterada_pro -eq '1' }).Count

    $agencies += [PSCustomObject]@{
        name = (PrettyName $t); ep = $ep; termHoy = $termHoy; cancHoy = $cancHoy
        br = $br; reiterCount = $reiterCount; reiterTotal = $ep
    }
}
$agencies = @($agencies | Sort-Object { if ($null -eq $_.br) { [double]::MaxValue } else { $_.br } })
Write-Log "Agencias: $($agencies.Count)"

# ---------------------------------------------------------------------------
# 6. Backlog por bucket (acumulado de todo el periodo, todas las filas/snapshots)
# ---------------------------------------------------------------------------
$bucketGroups = $data | Where-Object { $_.cod_bucket -ne '' } | Group-Object cod_bucket
$buckets = @()
foreach ($g in $bucketGroups) {
    $p = @($g.Group | Where-Object { $_.rdy_estado -eq 'pendiente' }).Count
    $t = @($g.Group | Where-Object { $_.rdy_estado -eq 'Cerrado' }).Count
    $c = @($g.Group | Where-Object { $_.rdy_estado -eq 'Cancelado' }).Count
    $buckets += [PSCustomObject]@{ name = $g.Name; c = $c; p = $p; t = $t }
}
Write-Log "Buckets: $($buckets.Count)"

# ---------------------------------------------------------------------------
# 7. Generar el bloque JS
# ---------------------------------------------------------------------------
function Num($v) { if ($null -eq $v) { return 'null' }; return $v.ToString($inv) }
function JsStr($s) { return "'" + ($s -replace "'", "\'") + "'" }

$sb = New-Object System.Text.StringBuilder
[void]$sb.AppendLine('  // ===AUTO-DATA-START=== (generado por update_dashboard.ps1 - no editar a mano)')
[void]$sb.AppendLine("  var GENERATED_AT = `"$((Get-Date).ToString('yyyy-MM-ddTHH:mm:ss'))`";")
[void]$sb.AppendLine("  var lastDate = `"$($lastDt.ToString('yyyy-MM-dd'))`";")
[void]$sb.AppendLine("  var zone = { enProceso:$zoneEnProceso, termHoy:$zoneTerminadasHoy, cancHoy:$zoneCanceladasHoy, br:$(Num $zoneBr), fbr:$(Num $zoneFbr) };")
[void]$sb.AppendLine('  var agencies = [')
for ($i = 0; $i -lt $agencies.Count; $i++) {
    $a = $agencies[$i]
    $comma = if ($i -lt $agencies.Count - 1) { ',' } else { '' }
    [void]$sb.AppendLine("    {name:$(JsStr $a.name), ep:$($a.ep), termHoy:$($a.termHoy), cancHoy:$($a.cancHoy), br:$(Num $a.br), reiterCount:$($a.reiterCount), reiterTotal:$($a.reiterTotal)}$comma")
}
[void]$sb.AppendLine('  ];')
[void]$sb.AppendLine('  var bucketData = {')
for ($i = 0; $i -lt $buckets.Count; $i++) {
    $b = $buckets[$i]
    $comma = if ($i -lt $buckets.Count - 1) { ',' } else { '' }
    [void]$sb.AppendLine("    $(JsStr $b.name): { c:$($b.c), p:$($b.p), t:$($b.t) }$comma")
}
[void]$sb.AppendLine('  };')
[void]$sb.AppendLine('  // ===AUTO-DATA-END===')

$newBlock = $sb.ToString().TrimEnd("`r","`n")

# ---------------------------------------------------------------------------
# 8. Inyectar en index.html (sin BOM, UTF-8 puro)
# ---------------------------------------------------------------------------
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
$html = [System.IO.File]::ReadAllText($IndexPath, [System.Text.Encoding]::UTF8)

$pattern = '(?s)  // ===AUTO-DATA-START===.*?// ===AUTO-DATA-END==='
if ($html -notmatch $pattern) {
    Write-Log "ERROR: no se encontraron los marcadores AUTO-DATA en index.html"
    exit 1
}
$html = [regex]::Replace($html, $pattern, { param($m) $newBlock }, 1)

[System.IO.File]::WriteAllText($IndexPath, $html, $utf8NoBom)
Write-Log "index.html actualizado (corte $lastDate, zona en_proceso=$zoneEnProceso, backlog=$zoneBr)"

# ---------------------------------------------------------------------------
# 9. Commit + push (EAP en Continue: git escribe avisos normales a stderr,
#    que PowerShell 5.1 convierte en error terminante si se captura con 2>&1
#    bajo Stop)
#
#    IMPORTANTE: se agrega TODO lo que cambio en el repo (git add -A), no solo
#    index.html/update.log. Si este script (o el .bat/README) se edita a mano,
#    ese cambio debe quedar commiteado tambien -- de lo contrario una carpeta
#    local recreada mas adelante (reclone, restauracion) vuelve a traer la
#    version vieja del script y revive bugs ya corregidos.
# ---------------------------------------------------------------------------
$ErrorActionPreference = 'Continue'
Set-Location $RepoDir

git add -A 2>$null 1>$null
$statusPorcelain = git status --porcelain
if ([string]::IsNullOrWhiteSpace($statusPorcelain)) {
    Write-Log "Sin cambios, no se genera commit."
} else {
    $commitMsg = "Actualizacion diaria backlog - $lastDate (zona en_proceso=$zoneEnProceso)"
    git commit -m $commitMsg 2>$null 1>$null
    git push origin HEAD 2>$null 1>$null
    if ($LASTEXITCODE -eq 0) { Write-Log "Push completado." }
    else { Write-Log "ERROR: git push devolvio codigo $LASTEXITCODE" }
}
$ErrorActionPreference = 'Stop'

Write-Log "== Fin =="
