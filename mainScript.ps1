# Note: Comments are written by me and function as a user and personal guide :)

# Detects School OneDrive automatically, alternately, a seperate path can replace this one 
if ( $env:OneDriveCommercial ) { $pathToSchoolFolder = $env:OneDriveCommercial } 
else { $pathToSchoolFolder = $env:OneDrive }
#Todo: Configue correct path inside OneDrive, f.E if a folder named as the same Module exists 


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

# TODO: If enough time, archive system 


# https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_regular_expressions?view=powershell-7.6
# https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_comparison_operators?view=powershell-7.6
# https://stackoverflow.com/questions/57947150/where-is-the-downloads-folder-located