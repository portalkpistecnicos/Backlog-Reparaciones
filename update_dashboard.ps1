<#
  update_dashboard.ps1
  Recalcula el KPI "Backlog Repara" desde el CSV de origen y actualiza index.html
  (bloque AUTO-DATA), luego hace commit + push al repositorio de GitHub.

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
# 1. Ubicar el CSV mas reciente
# ---------------------------------------------------------------------------
$csv = Get-ChildItem -Path $SourceDir -Filter 'p67_base_backlog_reparaciones_mod-detalle_*.csv' |
       Sort-Object LastWriteTime -Descending | Select-Object -First 1
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
    'VALPARAISO'                          = 'Valpara&iacute;so'
    'VINA DEL MAR'                        = 'Vi&ntilde;a del Mar'
    'SAN ANTONIO'                         = 'San Antonio'
    'ATENCION CLIENTES'                   = 'Atenci&oacute;n Clientes'
    'CASA CERRADA'                        = 'Casa Cerrada'
    'CASOS REMOTOS'                       = 'Casos Remotos'
    'PLAT NIVEL 2'                        = 'Plataforma Nivel 2'
    'NIVEL 1'                             = 'Nivel 1'
    'LOGISTICA'                           = 'Log&iacute;stica'
    'AGENDA Y VALIDACION'                 = 'Agenda y Validaci&oacute;n'
    'FALLA MASIVA PLANTA INTERNA'         = 'Falla Masiva Pta. Interna'
    'FALLA MASIVA PLANTA EXTERNA'         = 'Falla Masiva Pta. Externa'
    'PLANTA EXTERNA'                      = 'Planta Externa'
    'PEXT ATC'                            = 'PEXT ATC'
    ''                                    = 'Sin dato'
    'PLATAFORMA TERRENO'                  = 'Plataforma Terreno'
    'PLATAFORMA N2 BA'                    = 'Plataforma N2 BA'
    'PLATAFORMA CERRADOS POR UNIFICA'     = 'Plataforma Cerrados por Unifica'
    'PLATAFORMA N2 IPTV'                  = 'Plataforma N2 IPTV'
    'PLATAFORMA CASA CERRADA'             = 'Plataforma Casa Cerrada'
    'GESTION APP STREAMING'               = 'Gesti&oacute;n App Streaming'
    'PLATAFORMA REPARADO CASA CERRADA'    = 'Plataforma Reparado Casa Cerrada'
    'FALLA MASIVA MIGRACION'              = 'Falla Masiva Migraci&oacute;n'
    'PLATAFORMA FALLA MASIVA BANDA ANCHA' = 'Plataforma Falla Masiva Banda Ancha'
    'DERIVADO A ONNET'                    = 'Derivado a Onnet'
    'FALLA MASIVA FO'                     = 'Falla Masiva FO'
    'FALLA MASIVA PTA EXT CU'             = 'Falla Masiva Pta Ext CU'
    'DEVUELTO DE PLANTA EXTERNA'          = 'Devuelto de Planta Externa'
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

# ---------------------------------------------------------------------------
# 3. Serie diaria: En Proceso2 (pendientes del dia) y cierres reales
# ---------------------------------------------------------------------------
$enProceso = @{}
$data | Where-Object { $_.rdy_estado -eq 'pendiente' } | Group-Object partition_date |
    ForEach-Object { $enProceso[$_.Name] = $_.Count }

$closuresRaw = $data | Where-Object { $_.fecha_cierre -ne '' } |
    Select-Object rdy_id_incidencia, fecha_cierre, rdy_estado -Unique
$closuresByDate = @{}
$closuresRaw | Group-Object fecha_cierre | ForEach-Object { $closuresByDate[$_.Name] = $_.Count }

$allDates = $data.partition_date | Where-Object { $_ -ne '' } | Select-Object -Unique |
    Sort-Object { [datetime]::ParseExact($_, 'dd-MM-yyyy', $null) }

$dowMap = @{
    'Monday'='Lun'; 'Tuesday'='Mar'; 'Wednesday'='Mi&eacute;'; 'Thursday'='Jue'
    'Friday'='Vie'; 'Saturday'='S&aacute;b'; 'Sunday'='Dom'
}

$rows = @()
foreach ($d in $allDates) {
    $dt = [datetime]::ParseExact($d, 'dd-MM-yyyy', $null)
    $isSunday = ($dt.DayOfWeek -eq [System.DayOfWeek]::Sunday)
    $ep = if ($enProceso.ContainsKey($d)) { $enProceso[$d] } else { 0 }

    $sum6 = 0; $found = 0; $back = 0
    while ($found -lt 6 -and $back -lt 21) {
        $dd = $dt.AddDays(-$back)
        if ($dd.DayOfWeek -ne [System.DayOfWeek]::Sunday) {
            $key = $dd.ToString('dd-MM-yyyy')
            if ($closuresByDate.ContainsKey($key)) { $sum6 += $closuresByDate[$key] }
            $found++
        }
        $back++
    }
    $denom = $sum6 / 6.0
    $br = if ($denom -gt 0 -and -not $isSunday) { [math]::Round($ep / $denom, 3) } else { $null }

    $fbr = $null
    if ($null -ne $br) {
        if ($br -le 0.5) { $fbr = 1.08 }
        elseif ($br -ge 1.7) { $fbr = 0.93 }
        elseif ($br -gt 0.5 -and $br -le 1) { $fbr = ((1.08-1)*($br-1))/(0.5-1) + 1 }
        else { $fbr = ((1-0.93)*($br-1.7)/(1-1.7)) + 0.93 }
        $fbr = [math]::Round($fbr, 4)
    }

    $iso = $dt.ToString('yyyy-MM-dd')
    $cd  = if ($closuresByDate.ContainsKey($d)) { $closuresByDate[$d] } else { 0 }

    $rows += [PSCustomObject]@{
        iso = $iso; dow = $dowMap[$dt.DayOfWeek.ToString()]; ep = $ep; cd = $cd
        c6 = $sum6; denom = [math]::Round($denom,2); br = $br; fbr = $fbr; sun = $isSunday
    }
}
Write-Log "Dias procesados: $($rows.Count)"

# ---------------------------------------------------------------------------
# 4. Composicion del backlog actual (ultima fecha con datos)
# ---------------------------------------------------------------------------
$lastDate = $allDates[$allDates.Count - 1]
$lastPend = $data | Where-Object { $_.partition_date -eq $lastDate -and $_.rdy_estado -eq 'pendiente' }
$totalPend = $lastPend.Count
Write-Log "Ultima fecha: $lastDate  -  pendientes: $totalPend"

function BuildBreakdown($groupField, $maxItems) {
    $groups = $lastPend | Group-Object $groupField | Sort-Object Count -Descending
    $top = $groups | Select-Object -First $maxItems
    $rest = $groups | Select-Object -Skip $maxItems
    $items = @()
    foreach ($g in $top) {
        $items += [PSCustomObject]@{ name = (PrettyName $g.Name); value = $g.Count }
    }
    if ($rest.Count -gt 0) {
        $restSum = ($rest | Measure-Object Count -Sum).Sum
        $items += [PSCustomObject]@{ name = "Otros ($($rest.Count) grupos)"; value = $restSum }
    }
    return $items
}

$bdTerritorio = BuildBreakdown 'rdy_cod_territorio' 10
$bdAmbito     = BuildBreakdown 'mad_ambito' 8
$bdGrupo      = BuildBreakdown 'rdy_grupo_asignado' 6

# ---------------------------------------------------------------------------
# 5. Generar el bloque JS
# ---------------------------------------------------------------------------
function Num($v) {
    if ($null -eq $v) { return 'null' }
    return $v.ToString($inv)
}
function JsStr($s) { return "'" + ($s -replace "'", "\'") + "'" }

$sb = New-Object System.Text.StringBuilder
[void]$sb.AppendLine('  // ===AUTO-DATA-START=== (generado por update_dashboard.ps1 - no editar a mano)')
[void]$sb.AppendLine("  var GENERATED_AT = `"$((Get-Date).ToString('yyyy-MM-ddTHH:mm:ss'))`";")
[void]$sb.AppendLine('  var daily = [')
foreach ($r in $rows) {
    $line = "    {{d:`"{0}`", dow:`"{1}`", ep:{2}, cd:{3}, c6:{4}, denom:{5}, br:{6}, fbr:{7}, sun:{8}}}," -f `
        $r.iso, $r.dow, $r.ep, $r.cd, $r.c6, (Num $r.denom), (Num $r.br), (Num $r.fbr), $r.sun.ToString().ToLower()
    [void]$sb.AppendLine($line)
}
[void]$sb.AppendLine('  ];')
[void]$sb.AppendLine('')
[void]$sb.AppendLine('  var breakdowns = {')

function AppendGroup($name, $items, $total, $isLast) {
    [void]$sb.AppendLine("    ${name}: { total: $total, items: [")
    for ($i = 0; $i -lt $items.Count; $i++) {
        $comma = if ($i -lt $items.Count - 1) { ',' } else { '' }
        [void]$sb.AppendLine("      {name:$(JsStr $items[$i].name), value:$($items[$i].value)}$comma")
    }
    $closeComma = if ($isLast) { '' } else { ',' }
    [void]$sb.AppendLine("    ]}$closeComma")
}
AppendGroup 'territorio' $bdTerritorio $totalPend $false
AppendGroup 'ambito'     $bdAmbito     $totalPend $false
AppendGroup 'grupo'      $bdGrupo      $totalPend $true
[void]$sb.AppendLine('  };')
[void]$sb.AppendLine('  // ===AUTO-DATA-END===')

$newBlock = $sb.ToString().TrimEnd("`r","`n")

# ---------------------------------------------------------------------------
# 6. Inyectar en index.html (sin BOM, UTF-8 puro)
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
Write-Log "index.html actualizado ($($rows.Count) dias, corte $lastDate, $totalPend pendientes)"

# ---------------------------------------------------------------------------
# 7. Commit + push (EAP en Continue: git escribe avisos normales a stderr,
#    que PowerShell 5.1 convierte en error terminante si se captura con 2>&1
#    bajo Stop)
# ---------------------------------------------------------------------------
$ErrorActionPreference = 'Continue'
Set-Location $RepoDir

git add index.html update.log 2>$null 1>$null
$statusPorcelain = git status --porcelain -- index.html
if ([string]::IsNullOrWhiteSpace($statusPorcelain)) {
    Write-Log "Sin cambios en index.html, no se genera commit."
} else {
    $commitMsg = "Actualizacion diaria backlog - $lastDate ($totalPend pendientes)"
    git commit -m $commitMsg 2>$null 1>$null
    git push origin HEAD 2>$null 1>$null
    if ($LASTEXITCODE -eq 0) { Write-Log "Push completado." }
    else { Write-Log "ERROR: git push devolvio codigo $LASTEXITCODE" }
}
$ErrorActionPreference = 'Stop'

Write-Log "== Fin =="
