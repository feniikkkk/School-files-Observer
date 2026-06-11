# Note: Comments are written by me and function as a user/personal guide :)

# Detects School OneDrive automatically, alternately, a seperate path can replace this one 
if ($env:OneDriveCommercial) { $pathToSchoolFolder = $env:OneDriveCommercial } 
elseif ($env:OneDrive) { $pathToSchoolFolder = $env:OneDrive }

# Message popup for inputing a personal School files path
else {  
    Add-Type -AssemblyName PresentationFramework, System.Windows.Forms
    $personalPath = 'UNSET'
    if (Test-Path $personalPath) { $pathToSchoolFolder = $personalPath}

    while ($true) {
        $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
        $dialog.Description = "Choose a path to store your school files in"

        if ($dialog.ShowDialog() -eq 'OK') {
            if (Test-Path $dialog.SelectedPath) {
                (Get-Content $PSCommandPath) -replace '\$personalPath = ''UNSET''', "`$personalPath = `"$($dialog.SelectedPath)`"" | Set-Content $PSCommandPath # **
                $pathToSchoolFolder = $dialog.SelectedPath
                break;
            }
            [System.Windows.MessageBox]::Show("The selected path is invalid or does not exist.", "Path Error", "OK", "Error") 
        }
        else {
            [System.Windows.MessageBox]::Show("A valid destination directory is required to proceed.", "Operation Canceled", "OK", "Warning")
            exit
        }
    }
 }

# Configue correct path inside OneDrive, f.E if a folder named as the same Module exists 
function pathFinder($keyword) {
    $items = Get-Childitem -Path $pathToSchoolFolder
    foreach($item in $items) {
        if ($item.Name -match $keyword) {
            $newPath = $item.FullName 
            return $newPath
        }
    }

    (New-Item -Path $pathToSchoolFolder -Name $keyword -ItemType Directory).FullName

}

# Finds Downloads folder automatically, alternately, a seperate path can replace this one
$watchFolderPath = (New-Object -ComObject Shell.Application).NameSpace('shell:Downloads').Self.Path

$observer = New-Object System.IO.FileSystemWatcher 
$observer.path = $watchFolderPath
$observer.EnableRaisingEvents = $true

# Speed of the watcher: 
$observer.InternalBufferSize = 64KB  

$action = {
    Start-Sleep -Seconds 2
    $path = $Event.SourceEventArgs.FullPath
    $currentFile = Get-Item -Path $path
    if ($currentFile.Extension -match "crdownload|tmp") { return } 
    $origin = getOrigin($path)
    if ($origin.length -gt 0) {

        $target = pathFinder($origin)
        $expectedPath = Join-Path $target $currentFile.Name
        $baseTargetPath = $expectedPath

        $counter = 1 
        while (Test-Path $expectedPath) {                       # ** --> 
            $folder = Split-Path $baseTargetPath -Parent
            $name = (Get-Item $baseTargetPath).BaseName
            $ext = (Get-Item $baseTargetPath).Extension
            $expectedPath = "$folder\$name ($counter)$ext"
            $counter++
        }                                                       # ** <--
    
        Move-Item $currentFile $expectedPath -ErrorAction Stop 
        if ($origin -eq 'M231' -and $expectedPath -match '\.zip$') {
            $targetFolder = Split-Path $expectedPath -Parent
            Expand-Archive -Path $expectedPath -DestinationPath $targetFolder
            Remove-Item -Path $expectedPath -Force
        } 
    }
} 


function getOrigin($path) {
    $zoneData = Get-Content -Path $path -Stream Zone.Identifier -ErrorAction SilentlyContinue
    $hostUrl = ($zoneData | Where-Object { $_ -match "^HostUrl="}) -replace "HostUrl=", ""
    $urlData = ($hostUrl -split "/")
    $moduleName = ""
    foreach ($data in $urlData) {
        if ( $data -match 'M\d{3}') {
            $moduleName = [regex]::Match($data, "M\d{3}").Value
            break
        }
    }
    if ($moduleName -eq "") { return $moduleName} 
    return $moduleName
    
}



Register-ObjectEvent -InputObject $observer -EventName 'Created' -Action $action

# Any lines that are fully written by an Artificial Intelligence are marked with "**" 

# https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_regular_expressions?view=powershell-7.6
# https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_comparison_operators?view=powershell-7.6
# https://stackoverflow.com/questions/57947150/where-is-the-downloads-folder-located
# https://learn.microsoft.com/en-us/dotnet/api/system.management.automation.host.pshostuserinterface.promptforchoice?view=powershellsdk-7.6.0

# Links for ADS stuff, basically finding the origin of where a file was downloaded from
# https://hshrzd.wordpress.com/2016/03/19/introduction-to-ads-alternate-data-streams/?source=post_page-----c0e4a2402563---------------------------------------

# Message Box popup 
# https://learn.microsoft.com/en-us/dotnet/api/system.windows.messagebox.show?view=windowsdesktop-10.0
# https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_automatic_variables?view=powershell-7.6#pscommandpath