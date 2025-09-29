# ============================================================================================= #
# Coleta Informações completas de armazenamento das versões de arquivos em sites do SharePoint. #
# Autor: Jean Carlo Freitas                                                                     #
# Criação: 12/09/2025 | Revisão: 29/09/2025                                                     #
# Requisitos:                                                                                  #
#   - Registrar aplicativo para leitura do conteúdo dos sites                                   #
#   - Usuário precisa ser administrador do site                                                 #
#   - Executar no PowerShell 7                                                                 #
# ============================================================================================= #

# ============================
# Configuração e conexão
# ============================
$SiteUrl  = "https://tenanturl.sharepoint.com/sites/site" # Raiz do site
$Tenant   = "tenant.onmicrosoft.com"
$ClientId = "YOUR APP REGISTRATION ID"
$Output   = "C:\SharePoint_SiteName_Versions.xlsx"
$ErrorLog = "C:\SharePoint_SiteName_Errors.log"

# Conecta ao site
Connect-PnPOnline -Url $SiteUrl -Tenant $Tenant -ClientId $ClientId -Interactive

# Aumenta o timeout do contexto (10 minutos)
$ctx = Get-PnPContext
$ctx.RequestTimeout = 1000 * 60 * 10

# ============================
# Buscar bibliotecas do site
# ============================
Write-Host "Buscando as bibliotecas do site..." -ForegroundColor Green
$lists = Get-PnPList | Where-Object { $_.BaseTemplate -eq 101 -and -not $_.Hidden }

# Lista genérica (melhor performance que +=)
$report = [System.Collections.Generic.List[object]]::new()

foreach ($list in $lists) {
    Write-Host ">> Processando biblioteca: $($list.Title)" -ForegroundColor Yellow
    
    # Busca arquivos
    $items = Get-PnPListItem -List $list.Title -PageSize 500 -Fields "FileLeafRef","FileDirRef","FSObjType","FileRef" |
             Where-Object { $_["FSObjType"] -eq 0 }

    foreach ($item in $items) {
        try {
            $file = Get-PnPFile -Url $item["FileRef"] -AsListItem -ErrorAction Stop
            $versions = Get-PnPProperty -ClientObject $file.File -Property Versions

            if ($versions) {
                $versionCount = $versions.Count
                $versionSize  = ($versions | Measure-Object -Property Size -Sum).Sum / 1MB
            }
            else {
                $versionCount = 1
                $versionSize  = $file.File.Length / 1MB
            }

            # Cria objeto e adiciona na lista
            $obj = [PSCustomObject]@{
                Library        = $list.Title
                FileName       = $file["FileLeafRef"]
                FilePath       = $file["FileDirRef"]
                VersionCount   = $versionCount
                TotalVersionMB = [Math]::Round($versionSize,2)
                LastModified   = $file["Modified"]
                ModifiedBy     = $file["Editor"].Email
            }
            $report.Add($obj)
        }
        catch {
            $msg = ("[ERRO] {0} | {1} | {2}" -f (Get-Date), $item["FileRef"], $_.Exception.Message)
            Write-Host $msg -ForegroundColor Red
            Add-Content -Path $ErrorLog -Value $msg
        }
    }
}

# ============================
# Exportação
# ============================
$report | Sort-Object TotalVersionMB -Descending |
    Export-Excel -Path $Output -AutoSize -BoldTopRow -FreezeTopRow -WorksheetName "Ranking"

Write-Host ">> Relatório final gerado em $Output" -ForegroundColor Green
Write-Host ">> Log de erros salvo em $ErrorLog" -ForegroundColor Yellow