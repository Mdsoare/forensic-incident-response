<#
.SYNOPSIS
    Finaliza processos do Office deixados em segundo plano.
.DESCRIPTION
    Este script encerra instancias de Excel, Word e PowerPoint que nao possuem janelas visiveis,
    comumente deixadas por scripts de automacao ou auditoria.
#>

$Processos = @("EXCEL", "WINWORD", "POWERPNT")

Write-Host "--- Iniciando limpeza de processos Office ---" -ForegroundColor Cyan

foreach ($Nome in $Processos) {
    # Busca processos que nao tem uma janela principal (indicativo de execucao via script)
    $Instancias = Get-Process -Name $Nome -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowHandle -eq 0 }

    if ($Instancias) {
        foreach ($i in $Instancias) {
            try {
                Stop-Process -Id $i.Id -Force
                Write-Host "[OK] Processo $Nome (PID: $($i.Id)) encerrado." -ForegroundColor Green
            } catch {
                Write-Warning "Nao foi possivel encerrar $Nome (PID: $($i.Id))."
            }
        }
    } else {
        Write-Host "Nenhuma instancia 'zumbi' de $Nome encontrada." -ForegroundColor Gray
    }
}

Write-Host "--- Limpeza concluida ---" -ForegroundColor Cyan