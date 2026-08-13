<#
.SYNOPSIS
    Script para Triagem Forense Computacional, Resposta a Incidentes (IR) e Auditoria de Conformidade em Segurança da Informação.
.DESCRIPTION
    Busca profunda em PDF, Word e Excel usando motores nativos.
    Inclui cronometro, quarentena e compactacao 7-Zip.
.NOTES
    Autor: Marcelo Soares
    Data: Agosto/2026
    Versão: 1.1 (DevSecOps - Encoding Fixed)
#>
param(
    [Parameter(Mandatory = $true)][string[]]$Termos,
    [string]$CaminhoAlvo = ".",
    [string]$RelatorioNome = "Relatorio_Forense_Final.html",
    [switch]$AtivarQuarentena,
    [string]$SenhaZip = "infected"
)

# --- 1. INÍCIO DO CRONÔMETRO ---
$TempoInicio = [System.Diagnostics.Stopwatch]::StartNew()

$Resultados = New-Object System.Collections.Generic.List[PSObject]
$Pattern = $Termos -join "|"
$DataID = Get-Date -Format "ddMMyyyy_HHmm"
$QuarentenaPath = Join-Path $env:USERPROFILE "Desktop\Quarentena_Forense_$DataID"
$Path7Zip = "C:\Program Files\7-Zip\7z.exe"

if ($AtivarQuarentena -and !(Test-Path $QuarentenaPath)) {
    New-Item -ItemType Directory -Path $QuarentenaPath -Force | Out-Null
}

$Arquivos = Get-ChildItem -Path $CaminhoAlvo -Include *.doc*, *.xls*, *.ppt*, *.pdf, *.txt, *.csv, *.log, *.crypt* -Recurse -ErrorAction SilentlyContinue

Write-Host "--- Iniciando Auditoria Forense em $($Arquivos.Count) arquivos ---" -ForegroundColor Cyan

# Inicialização dos Motores Office
try {
    $word = New-Object -ComObject Word.Application; $word.Visible = $false
    $excel = New-Object -ComObject Excel.Application; $excel.Visible = $false
}
catch { Write-Warning "Componentes Office não disponíveis. A busca será limitada a texto simples." }

# --- 2. VARREDURA ---
foreach ($File in $Arquivos) {
    $Encontrou = $false
    $TipoIncidente = ""
    $Ext = $File.Extension.ToLower()

    try {
        # A. RANSOMWARE
        if ($Ext -eq ".crypt" -or $Ext -eq ".locked") {
            $Encontrou = $true; $TipoIncidente = "Extensao Suspeita (Ransomware)"
        }
        # B. WORD & PDF (Via Motor Word)
        elseif ($Ext -match "doc" -or $Ext -eq ".pdf") {
            if ($word) {
                $doc = $word.Documents.Open($File.FullName, $false, $true)
                $find = $doc.Content.Find
                foreach ($termo in $Termos) {
                    if ($find.Execute($termo)) { $Encontrou = $true; $TipoIncidente = "Detectado via Motor Word ($($Ext.ToUpper()))"; break }
                }
                $doc.Close($false)
            }
        }
        # C. EXCEL (Via Motor Excel)
        elseif ($Ext -match "xls") {
            if ($excel) {
                $wb = $excel.Workbooks.Open($File.FullName)
                foreach ($sheet in $wb.Worksheets) {
                    if ($sheet.UsedRange.Find($Termos[0])) { $Encontrou = $true; $TipoIncidente = "Termo em Planilha Excel"; break }
                }
                $wb.Close($false)
            }
        }
        # D. TEXTO SIMPLES / LOGS (Fallback)
        else {
            $Content = Get-Content -Path $File.FullName -Raw -ErrorAction SilentlyContinue
            if ($Content -match $Pattern) { $Encontrou = $true; $TipoIncidente = "Termo em Arquivo de Texto/Log" }
        }

        # --- 3. REGISTRO E QUARENTENA ---
        if ($Encontrou) {
            $Hash = (Get-FileHash $File.FullName -Algorithm SHA256).Hash
            $StatusAcao = "Detectado"
            if ($AtivarQuarentena) {
                $StatusAcao = "Isolado em Quarentena"
                $Destino = Join-Path $QuarentenaPath "$($File.Name)_$($Hash.Substring(0,8))"
                Move-Item -Path $File.FullName -Destination $Destino -Force
            }
            $Resultados.Add([PSCustomObject]@{
                    DataHora = Get-Date -Format "dd/MM/yyyy HH:mm:ss"
                    Arquivo  = $File.Name
                    Tipo     = $TipoIncidente
                    Acao     = $StatusAcao
                    Caminho  = $File.FullName
                    SHA256   = $Hash
                })
            Write-Host "[!] $StatusAcao : $($File.Name)" -ForegroundColor Yellow
        }
    }
    catch { 
        # Log silencioso de erros de abertura
    }
}

# --- 4. FINALIZAÇÃO ---
if ($word) { $word.Quit() }; if ($excel) { $excel.Quit() }
$TempoInicio.Stop()
$TotalSegundos = [math]::Round($TempoInicio.Elapsed.TotalSeconds, 2)

# Compactação 7-Zip
if ($AtivarQuarentena -and $Resultados.Count -gt 0 -and (Test-Path $Path7Zip)) {
    $ZipDestino = Join-Path $env:USERPROFILE "Desktop\Evidencias_Protegidas_$DataID.zip"
    $Args7z = "a `"$ZipDestino`" `"$QuarentenaPath\*`" -p$SenhaZip -mhe=on"
    Start-Process -FilePath $Path7Zip -ArgumentList $Args7z -Wait -NoNewWindow
    Remove-Item $QuarentenaPath -Recurse -Force
}

# Geração de Relatório com Métricas
$Style = '<style>body{font-family:sans-serif;padding:30px;background:#f0f2f5;} .card{background:white;padding:20px;border-radius:8px;box-shadow:0 2px 4px rgba(0,0,0,0.1);} table{width:100%;border-collapse:collapse;margin-top:20px;} th{background:#1a3a5f;color:white;padding:12px;text-align:left;} td{border:1px solid #dee2e6;padding:10px;font-size:11px;} .metricas{color:#555;font-size:14px;margin-bottom:20px;}</style>'
$Html = "<html><head>$Style</head><body><div class='card'><h2>Relatorio de Investigacao Digital</h2>"
$Html += "<div class='metricas'><b>Perito:</b> Marcelo Soares | <b>Duracao:</b> $TotalSegundos seg | <b>Termos:</b> $($Termos -join ', ')</div>"
$Html += ($Resultados | ConvertTo-Html -Fragment)
$Html += "</div></body></html>"
$Html | Out-File $RelatorioNome -Encoding UTF8

Write-Host "`n--- Auditoria Concluida em $TotalSegundos segundos ---" -ForegroundColor Green
Write-Host "Relatorio disponivel em: $RelatorioNome" -ForegroundColor Green