# ============================================================================================= #
# Limpeza real de versões no SharePoint - REMOVE versões excedentes                             #
# Autor: Jean Carlo Freitas                                                                     #
# Criação: 26/09/2025                                                                           #
# ============================================================================================= #

# Configurações
$SiteUrl   = "https://sascartecnologia.sharepoint.com/sites/BORDEROS"
$Tenant    = "sascartecnologia.onmicrosoft.com"
$ClientId  = "4b6028cc-ad3b-4884-8d1c-2b6210dadddf"
$KeepCount = 20   # Quantidade de versões a manter
$LogFile   = "C:\Temp\SharePoint_Limpeza_Real.log"

# Conectar
Connect-PnPOnline -Url $SiteUrl -Tenant $Tenant -ClientId $ClientId -Interactive
$ctx = Get-PnPContext
$ctx.RequestTimeout = 1000 * 60 * 10

# Função de log
function Write-Log {
    param([string]$msg)
    $line = "[{0}] {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $msg
    Add-Content -Path $LogFile -Value $line
    Write-Host $msg
}

# Bibliotecas
$lists = Get-PnPList | Where-Object { $_.BaseTemplate -eq 101 -and -not $_.Hidden }

foreach ($list in $lists) {
    Write-Log ">> Processando biblioteca: $($list.Title)"
    $items = Get-PnPListItem -List $list.Title -PageSize 500 -Fields "FileLeafRef","FileDirRef","FSObjType","FileRef" |
             Where-Object { $_["FSObjType"] -eq 0 }

    foreach ($item in $items) {
        try {
            $file = Get-PnPFile -Url $item["FileRef"] -AsListItem -ErrorAction Stop
            $versions = Get-PnPProperty -ClientObject $file.File -Property Versions

            if ($versions.Count -gt $KeepCount) {
                $excedentes = $versions | Sort-Object Created -Descending | Select-Object -Skip $KeepCount

                foreach ($v in $excedentes) {
                    Write-Log ("Removendo versão {0} de {1} ({2} KB)" -f $v.VersionLabel, $file["FileLeafRef"], [Math]::Round($v.Size/1KB,2))
                    $v.DeleteObject()
                }

                Invoke-PnPQuery
            }
        }
        catch {
            Write-Log ("[ERRO] {0} | {1}" -f $item["FileRef"], $_.Exception.Message)
        }
    }
}

Write-Log ">> Processo de limpeza concluído."