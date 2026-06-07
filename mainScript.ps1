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
                (Get-Content $PSCommandPath) -replace '\$SavedPath = "UNSET"', "`$SavedPath = `"$($dialog.SelectedPath)`"" | Set-Content $PSCommandPath
                $pathToSchoolFolder = $dialog.SelectedPath
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

    New-Item -Path $pathToSchoolFolder -ChildPath $keyword

}

# Finds Downloads folder automatically, alternately, a seperate path can replace this one
$watchFolderPath = (New-Object -ComObject Shell.Application).NameSpace('shell:Downloads').Self.Path

$observer = New-Object System.IO.FileSystemWatcher 
$observer.path = $watchFolderPath
$observer.EnableRaisingEvents = $true

# Speed of the watcher: 
$observer.InternalBufferSize = 64KB  

$action = {
    $path = $Event.SourceEventArgs.FullPath
    $currentFile = Get-Item -Path $path
    if ( $path -match 'M231' ) {
        # TODO: Get Obsidian stuff into the correct Obsidian directory 
    }
    elseif ( $path -match 'M\d{3}' ) { # Searches for classic Module IT Style "M**" f.E M122 
        #TODO: File collision handling 
        Move-Item $currentFile $pathToSchoolFolder -ErrorAction Stop 
        # TODO: Organization 
    }
    
    # TODO: More school-like checks.
}

# Register-ObjectEvent -InputObject $observer -EventName 'Created' -Action $action

# TODO: If enough time, archive-system 


# https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_regular_expressions?view=powershell-7.6
# https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_comparison_operators?view=powershell-7.6
# https://stackoverflow.com/questions/57947150/where-is-the-downloads-folder-located
# https://learn.microsoft.com/en-us/dotnet/api/system.management.automation.host.pshostuserinterface.promptforchoice?view=powershellsdk-7.6.0

# Links for ADS stuff, basically finding the origin of where a file was downloaded from
# https://hshrzd.wordpress.com/2016/03/19/introduction-to-ads-alternate-data-streams/?source=post_page-----c0e4a2402563---------------------------------------

# Message Box popup 
# https://learn.microsoft.com/en-us/dotnet/api/system.windows.messagebox.show?view=windowsdesktop-10.0
# https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_automatic_variables?view=powershell-7.6#pscommandpath