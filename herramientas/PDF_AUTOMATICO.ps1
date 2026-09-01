# =====================================================================
#  HNAL · PDF automatico
#  Vigila las carpetas donde la app deja los documentos y convierte a PDF
#  cada Word o Excel nuevo, usando el propio Word/Excel (respeta membrete,
#  logo y area de impresion).
#
#  Uso:  clic derecho en INICIAR_PDF_AUTOMATICO.bat  ->  Ejecutar
#  Para que arranque solo con Windows, ver LEEME.txt
# =====================================================================

# ---------- 1) CARPETAS A VIGILAR (edita esta lista) ----------
$Carpetas = @(
    "$env:USERPROFILE\Documents\HNAL\Validaciones",
    "$env:USERPROFILE\Documents\HNAL\Memos",
    "$env:USERPROFILE\Documents\HNAL\Comparativos",
    "$env:USERPROFILE\Documents\HNAL\Notas",
    "$env:USERPROFILE\Documents\HNAL\Cartas"
)

# Cada cuantos segundos revisa
$CadaSegundos = 10
# Donde deja el PDF: "" = misma carpeta; o pon una ruta fija
$CarpetaPDF = ""

# ---------------------------------------------------------------------
$Log = Join-Path $PSScriptRoot "pdf_automatico.log"
function Escribir($txt) {
    $linea = "{0}  {1}" -f (Get-Date -Format "dd/MM/yyyy HH:mm:ss"), $txt
    Write-Host $linea
    Add-Content -Path $Log -Value $linea -Encoding UTF8
}

# Espera a que el archivo termine de escribirse (si no, Word lo abre a medias)
function ArchivoListo($ruta) {
    for ($i = 0; $i -lt 12; $i++) {
        try {
            $f = [System.IO.File]::Open($ruta, 'Open', 'Read', 'None')
            $f.Close(); $f.Dispose()
            return $true
        } catch { Start-Sleep -Milliseconds 500 }
    }
    return $false
}

function RutaPDF($ruta) {
    $nombre = [System.IO.Path]::GetFileNameWithoutExtension($ruta) + ".pdf"
    if ($CarpetaPDF -and (Test-Path $CarpetaPDF)) { return Join-Path $CarpetaPDF $nombre }
    return Join-Path ([System.IO.Path]::GetDirectoryName($ruta)) $nombre
}

# Ya convertido y sin cambios posteriores -> no se rehace
function YaConvertido($ruta) {
    $pdf = RutaPDF $ruta
    return (Test-Path $pdf) -and ((Get-Item $pdf).LastWriteTime -ge (Get-Item $ruta).LastWriteTime)
}

$word = $null; $excel = $null

function ConvertirWord($ruta) {
    if (-not $script:word) {
        $script:word = New-Object -ComObject Word.Application
        $script:word.Visible = $false
        $script:word.DisplayAlerts = 0
    }
    $doc = $script:word.Documents.Open($ruta, $false, $true)   # solo lectura
    try   { $doc.ExportAsFixedFormat((RutaPDF $ruta), 17) }    # 17 = wdExportFormatPDF
    finally { $doc.Close($false) }
}

function ConvertirExcel($ruta) {
    if (-not $script:excel) {
        $script:excel = New-Object -ComObject Excel.Application
        $script:excel.Visible = $false
        $script:excel.DisplayAlerts = $false
    }
    $wb = $script:excel.Workbooks.Open($ruta, $false, $true)   # solo lectura
    try   { $wb.ExportAsFixedFormat(0, (RutaPDF $ruta)) }      # 0 = xlTypePDF
    finally { $wb.Close($false) }
}

function CerrarOffice {
    if ($script:word)  { try { $script:word.Quit() }  catch {}; $script:word = $null }
    if ($script:excel) { try { $script:excel.Quit() } catch {}; $script:excel = $null }
    [System.GC]::Collect()
}

Escribir "=== PDF automatico iniciado. Ctrl+C para detener. ==="
foreach ($c in $Carpetas) {
    if (Test-Path $c) { Escribir "Vigilando: $c" }
    else              { Escribir "AVISO: no existe la carpeta $c (se ignora)" }
}

$sinTrabajo = 0
while ($true) {
    $hubo = $false
    foreach ($carpeta in $Carpetas) {
        if (-not (Test-Path $carpeta)) { continue }
        $archivos = Get-ChildItem -Path $carpeta -File -Include *.docx, *.xlsx -Recurse -ErrorAction SilentlyContinue |
                    Where-Object { $_.Name -notlike '~$*' }
        foreach ($a in $archivos) {
            if (YaConvertido $a.FullName) { continue }
            if (-not (ArchivoListo $a.FullName)) {
                Escribir "En uso, se reintenta luego: $($a.Name)"
                continue
            }
            try {
                if ($a.Extension -eq ".docx") { ConvertirWord  $a.FullName }
                else                          { ConvertirExcel $a.FullName }
                Escribir "PDF listo: $($a.Name)"
                $hubo = $true
            } catch {
                Escribir "ERROR con $($a.Name): $($_.Exception.Message)"
            }
        }
    }
    # si lleva un rato sin nada que hacer, se cierra Word/Excel para no ocupar memoria
    if ($hubo) { $sinTrabajo = 0 } else { $sinTrabajo++ }
    if ($sinTrabajo -ge 30) { CerrarOffice; $sinTrabajo = 0 }
    Start-Sleep -Seconds $CadaSegundos
}
