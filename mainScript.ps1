# Note: Comments are written by me and function as a user/personal guide :)

$global:scriptDir = $PSScriptRoot

# Detects School OneDrive automatically.
if ($env:OneDriveCommercial) { $global:pathToSchoolFolder = $env:OneDriveCommercial } 
elseif ($env:OneDrive) { $global:pathToSchoolFolder = $env:OneDrive }

# Message popup for inputing a personal School files path
else {  
    Add-Type -AssemblyName PresentationFramework, System.Windows.Forms
    $personalPath = "UNSET"
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
$global:watchFolderPath = (New-Object -ComObject Shell.Application).NameSpace('shell:Downloads').Self.Path

# Observer setup
$observer = New-Object System.IO.FileSystemWatcher 
$observer.path = $global:watchFolderPath
$observer.EnableRaisingEvents = $true

# Speed of the watcher: 
$observer.InternalBufferSize = 64KB  

# Main script
$action = {
    $path = $Event.SourceEventArgs.FullPath
    $currentFile = Get-Item -Path $path
    if ($currentFile.Extension -match "crdownload|tmp") { return } 
    # ** -->
    $fileIsLocked = $true
    while ($fileIsLocked) {
    try {
        # Tries to open the file exclusively. If Chrome is downloading, this fails.
        $stream = [System.IO.File]::Open($path, 'Open', 'Read', 'None')
        $stream.Close()
        $fileIsLocked = $false 
    } catch {
        if (-not (Test-Path $path)) { return }
        Start-Sleep -Milliseconds 500
    }
    # ** <--
}



    $origin = getOrigin($path)

    # Manual Module name checks ** -->
    $isZip = ($currentFile.Extension -eq ".zip")
    $tempDir = ""

    if ($origin -eq "" -or $origin -eq "SF") {
        if ($isZip) {
            $tempDir = Join-Path $global:watchFolderPath "Temp_$($currentFile.BaseName)" 
            Expand-Archive -Path $path -DestinationPath $tempDir -Force
            $scanList = Get-ChildItem -Path $tempDir -Recurse
        } else {
            $scanList = @($currentFile) 
        }
    # ** <--
        foreach ($file in $scanList) {
            if ($file.Name -match 'M\d{3}') {
                $origin = [regex]::Match($file.Name, "M\d{3}").Value
                break
            }
            else {
            switch ($file.Extension) {

                { $_ -match "\.txt|\.md"} {
                    $text = Get-Content $file.FullName -Raw -ErrorAction SilentlyContinue
                    if ($text -match 'M\d{3}') {
                            $origin = [regex]::Match($text, "M\d{3}").Value 
                            break
                    }
                }
                ".pptx" {
                    $tempDirPPT = Join-Path $global:watchFolderPath "Tmp_$($file.BaseName)"
                    Expand-Archive -Path $file.FullName -DestinationPath $tempDirPPT -Force
                    $slides = Get-ChildItem -Path "$tempDirPPT\ppt\slides\*.xml"
                    foreach ($slide in $slides) {
                        $text = Get-Content $slide.FullName -Raw -ErrorAction SilentlyContinue
                        if ($text -match 'M\d{3}') {
                            $origin = [regex]::Match($text, "M\d{3}").Value 
                            break
                        }
                    }
                    Remove-Item -Path $tempDirPPT -Recurse -Force -ErrorAction SilentlyContinue

                    if ($origin -match 'M\d{3}') { break }
                }
                ".pdf" {
                    $toolPath = Join-Path $global:scriptDir "PDFReader\pdftotext.exe"
                    $text = & $toolPath $file.FullName -
                    if ($text -match 'M\d{3}') {
                            $origin = [regex]::Match($text, "M\d{3}").Value 
                            break
                    }
                }
            }
        }
        }
        
        if ($origin -notmatch 'M\d{3}' -and $origin -eq "SF") { 
            $origin = "Unsorted" 
        }
    }

    #  Move Items 
    if ($isZip -and $tempDir -ne "") {
        if ($origin -ne "Unsorted") {
            $targetFolder = pathFinder($origin)
            foreach ($item in Get-ChildItem -Path $tempDir) {
                Move-Item -Path $item.FullName -Destination (Join-Path $targetFolder $item.Name) -Force
            }
            Remove-Item -Path $path -Force
        }
        Remove-Item -Path $tempDir -Recurse -Force
        if ($origin -ne "Unsorted") { return }
    }
    # File Collision Handling and M231 Unzipping 
    if ($origin.length -gt 0) {

        $target = pathFinder($origin)
        $expectedPath = Join-Path $target $currentFile.Name
        $baseTargetPath = $expectedPath

        $counter = 1 
        while (Test-Path $expectedPath) {                       # * --> 
            $folder = Split-Path $baseTargetPath -Parent
            $name = (Get-Item $baseTargetPath).BaseName
            $ext = (Get-Item $baseTargetPath).Extension
            $expectedPath = "$folder\$name ($counter)$ext"
            $counter++
        }                                                       # * <--
    
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
    $isFromTeams = $false
    foreach ($data in $urlData) {
        if ( $data -match 'M\d{3}') {
            $moduleName = [regex]::Match($data, "M\d{3}").Value
            break
        }
        if ( $data -match 'sluz.sharepoint.com') {
            $isFromTeams = $true
        }
    }

    if ($moduleName -eq "") {
        if ($isFromTeams) {
            return "SF"     #SF = School File
        }
        return ""
    } 
    return $moduleName
    
}



Register-ObjectEvent -InputObject $observer -EventName 'Created' -Action $action
Register-ObjectEvent -InputObject $observer -EventName 'Renamed' -Action $action


# Any lines that are fully written by an Artificial Intelligence are marked with "**" 
# Any lines that are partially written by an Artificial Intelligence are marked with "*" 

# https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_automatic_variables?view=powershell-7.6#:~:text=PSScriptRoot%20%2D%20Contains%20the%20full%20path,that%20invoked%20the%20current%20command.
# https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_regular_expressions?view=powershell-7.6
# https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_comparison_operators?view=powershell-7.6
# https://stackoverflow.com/questions/57947150/where-is-the-downloads-folder-located
# https://learn.microsoft.com/en-us/dotnet/api/system.management.automation.host.pshostuserinterface.promptforchoice?view=powershellsdk-7.6.0
# https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_operators?view=powershell-7.6

# ADS stuff, basically finding the origin of where a file was downloaded from
# https://hshrzd.wordpress.com/2016/03/19/introduction-to-ads-alternate-data-streams/?source=post_page-----c0e4a2402563---------------------------------------

# Message Box popup 
# https://learn.microsoft.com/en-us/dotnet/api/system.windows.messagebox.show?view=windowsdesktop-10.0
# https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_automatic_variables?view=powershell-7.6#pscommandpath


# PDF Reader 
# https://www.xpdfreader.com/download.html