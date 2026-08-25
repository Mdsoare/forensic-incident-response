<#
.SYNOPSIS
    Finaliza processos do Office deixados em segundo plano.
.DESCRIPTION
    Este script encerra instancias de Excel, Word e PowerPoint que nao possuem janelas visiveis,
    comumente deixadas por scripts de automacao ou auditoria.
#>

$Processos = @("EXCEL", "WINWORD", "POWERPNT")

Write-Output "--- Iniciando limpeza de processos Office ---"

foreach ($Nome in $Processos) {
    # Busca processos que nao tem uma janela principal (indicativo de execucao via script)
    $Instancias = Get-Process -Name $Nome -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowHandle -eq 0 }

    if ($Instancias) {
        foreach ($i in $Instancias) {
            try {
                Stop-Process -Id $i.Id -Force
                Write-Output "[OK] Processo $Nome (PID: $($i.Id)) encerrado."
            } catch {
                Write-Warning "Nao foi possivel encerrar $Nome (PID: $($i.Id))."
            }
        }
    } else {
        Write-Output "Nenhuma instancia 'zumbi' de $Nome encontrada."
    }
}

Write-Output "--- Limpeza concluida ---"