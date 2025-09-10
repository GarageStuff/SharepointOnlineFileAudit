$SharePointAdminSiteURL = "https://tenant-admin.sharepoint.com"

#Azure App registration ID
$AppID = ""

$conn = Connect-PnPOnline -Url $SharePointAdminSiteURL -clientid $AppID

# Set Variables
$dateTime = (Get-Date).toString("dd-MM-yyyy")
$invocation = (Get-Variable MyInvocation).Value
$directorypath = Split-Path $invocation.MyCommand.Path
$fileName = "\SiteStoragReport-" + $dateTime + ".csv"
$OutputSite = $directorypath + $fileName
$fileName = "\FileStorageReport-" + $dateTime + ".csv"
$OutPutFile = $directorypath + $fileName

$arraySite = New-Object System.Collections.ArrayList
$arrayFile = New-Object System.Collections.ArrayList

#Exclude certain libraries
#$ExcludedLibraries = @("Form Templates", "Preservation Hold Library", "Site Assets", "Site Pages", "Images", "Pages", "Settings", "Videos", "Site Collection Documents", "Site Collection Images", "Style Library", "AppPages", "Apps for SharePoint", "Apps for Office")

function ReportStorageVersions($site) {
    try {
        $fileSizes = @(); 
        $fileSize = 0 
        $TotalVersionSize = 0
        $DocLibraries = Get-PnPList -Includes BaseType, Hidden, Title -Connection $siteconn | Where-Object { $_.BaseType -eq "DocumentLibrary" -and $_.Hidden -eq $False -and $_.Title -notin $ExcludedLibraries }
        $DocLibraries | ForEach-Object {
            Write-host "Processing Document Library:" $_.Title -f Yellow
            $library = $_
            $listItems = Get-PnPListItem -List $library.Title -Fields "ID" -PageSize 1000 -Connection $siteconn

            #Get file zize
            $listItems | ForEach-Object {
                $listitem = $_
                $fileVersionSize = 0
                $file = Get-PnPFile -Url $listitem["FileRef"] -AsFileObject -ErrorAction SilentlyContinue -Connection $siteconn 

                if ($file) {
                    $fileSize += $file.Length          
                    $elementFile = "" | Select-Object SiteUrl, siteName, siteStorage, FileRef,FileSize,TotalVersionSize,VersionCount,StartTime, EndTime
                    $elementFile.SiteUrl = $site.Url
                    $elementFile.siteName = $site.Title
                    $elementFile.siteStorage = "$siteStorage MB"
                    $elementFile.StartTime = (Get-Date).toString("dd-MM-yyyy HH:mm:ss")
                    $elementFile.FileRef  =   $listitem["FileRef"]
                    $fileversions = Get-PnPFileVersion -Url $listitem["FileRef"] -Connection $siteconn
                    if ($fileversions) {
                        # Calculate the total version size
                        $fileVersionSize = $fileversions | Measure-Object -Property Size -Sum | Select-Object -ExpandProperty Sum                                                   
                    }

                    $elementFile.FileSize = "$([Math]::Round(($file.Length/1MB),3)) MB" 
                    $elementFile.TotalVersionSize = "$([Math]::Round(($fileVersionSize/1MB),3)) MB"
                    $elementFile.VersionCount = $fileversions.Count
                    $totalVersionSize += $fileVersionSize
                    $elementFile.EndTime = (Get-Date).toString("dd-MM-yyyy HH:mm:ss")
                    $arrayFile.Add($elementFile) | Out-Null 
                }        
            }
        }
        $fileSizes += $fileSize
        $fileSizes += $totalVersionSize   

        return $fileSizes
    }
    catch {
        Write-Output "An exception was thrown: $($_.Exception.Message)" -ForegroundColor Red
    } 
}

# Get total storage use for the site collection, amend query to run reports against site collection(s), e.g. filter by $_.StorageUsageCurrent -gt 10000
Get-PnPTenantSite -Connection $conn | Where-Object { ($_.Template -eq "GROUP#0" -or $_.Template -eq "SITEPAGEPUBLISHING#0") } | ForEach-Object {
    $site = $_
    $siteStorage = $site.StorageUsageCurrent
    #$siteStorage = $siteStorage/1024l
    #$siteStorage = [Math]::Round($siteStorage, 2)
    Write-Host "Site storage: $siteStorage MB"
    $siteconn = Connect-PnPOnline -Url $site.Url -clientid 96e6e507-ed1b-48ee-b55f-514cf15fd3d0 -ReturnConnection

    $element = "" | Select-Object SiteUrl, siteName, siteStorage, FileSize, StartTime,TotalVersionSize, RecycleBinSize,EndTime
    $element.SiteUrl = $site.Url
    $element.siteName = $site.Title
    $element.siteStorage = "$siteStorage MB"
    $element.StartTime = (Get-Date).toString("dd-MM-yyyy HH:mm:ss")
    $FileSizeVersions = ReportStorageVersions -site $site
    $element.FileSize = "$([Math]::Round(($FileSizeVersions[0]/1MB),3)) MB" 
    $element.TotalVersionSize = "$([Math]::Round(($FileSizeVersions[1]/1MB),3)) MB"
    $RecycleBinItemsSize = Get-PnPRecycleBinItem -Connection $siteconn | Measure-Object -Property Size -Sum | Select-Object -ExpandProperty Sum
    $element.RecycleBinSize = "$([Math]::Round(($RecycleBinItemsSize/1MB),3)) MB"
    $element.EndTime = (Get-Date).toString("dd-MM-yyyy HH:mm:ss")

    $arraySite.Add($element) | Out-Null 
}  

$arraySite | Export-Csv -Path $OutputSite -NoTypeInformation -Force 
$arrayFile | Export-Csv -Path $OutputFile -NoTypeInformation -Force


