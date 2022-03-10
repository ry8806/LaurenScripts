Import-Module BitsTransfer

# Tests for the Unzip path and if not, creates it
$unzipPath = "C:\Temp\unzips"
If(!(test-path $unzipPath))
{
    New-Item -ItemType Directory -Force -Path $unzipPath | Out-Null
}

Function Select_File($InitialDirectory)
{
    Add-Type -AssemblyName System.Windows.Forms
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.Title = "Please Select File"
    #$OpenFileDialog.InitialDirectory = $InitialDirectory
    #$OpenFileDialog.filter = “Zip files (*.7z;*.zip)| *.7z;*.zip”
    $OpenFileDialog.filter = “Zip files (*.7z;*.zip;*.mkv)| *.7z;*.zip;*.mkv”
    $OpenFileDialog.RestoreDirectory = $True

    If ($OpenFileDialog.ShowDialog() -eq "Cancel")
    {
        Write-Host "No file chosen, exiting"
        Exit
    }
    $Global:SelectedFile = $OpenFileDialog.FileName
}

# Gets the folder where 7-Zip lives
$zipLocation = (Get-ChildItem HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall |
% { Get-ItemProperty $_.PsPath } | Where-Object { $_.DisplayName -like '7-Zip*' } | Select InstallLocation).InstallLocation

if ([string]::IsNullOrEmpty($zipLocation)){
    Write-Host "Can't find 7-Zip using Registry"

    $7zipStandardLocation = "C:\Program Files\7-Zip\"

    if (!(Test-Path -Path $7zipStandardLocation)) {
        Write-Host "7-Zip not found, falling back to Windows Extractor" 

        # TODO: Fall back here
    }
    else {
        $zipLocation = $7zipStandardLocation
    }
}

# Open the Select File dialog
Select_File([Environment]::GetFolderPath('Desktop'))
$fileToCopy = $Global:SelectedFile

Write-Host "Got File:" $Global:SelectedFile
Write-Host "Moving Selected file to Temp Folder: " $unzipPath
Start-BitsTransfer -Source "$fileToCopy" -Destination "$unzipPath" -Description "Destination: $unzipPath" -DisplayName "Copying: $fileToCopy"

$newItemPath = Join-Path -Path "$unzipPath" -ChildPath (Split-Path $fileToCopy -leaf)
$newItemFolder = Join-Path -Path $unzipPath -ChildPath ([System.IO.Path]::GetFileNameWithoutExtension($fileToCopy))

# Clean up the folder (if we've already unzipped before)
if (Test-Path -Path $newItemFolder) {
    Remove-Item -LiteralPath $newItemFolder -Force -Recurse
}

# Prepare the 7-Zip Path and Arguments
$7zip = Join-Path -Path $zipLocation -ChildPath "7z.exe"
$7zip_args = "x `"$newItemPath`" -o`"$newItemFolder`""

Write-Host "Unzipping file..."

$pinfo = New-Object System.Diagnostics.ProcessStartInfo
$pinfo.FileName = "$7zip"
$pinfo.RedirectStandardError = $true
$pinfo.RedirectStandardOutput = $true
$pinfo.UseShellExecute = $false
$pinfo.Arguments = $7zip_args
$p = New-Object System.Diagnostics.Process
$p.StartInfo = $pinfo
$p.Start() | Out-Null
$p.WaitForExit()
$stdout = $p.StandardOutput.ReadToEnd()
$stderr = $p.StandardError.ReadToEnd()

If(![string]::IsNullOrEmpty($stderr))
{
    Write-Host "Error: $stderr"
    Write-Host "Unzipping encountered an error. Please contact Ryan"
}

If(![string]::IsNullOrEmpty($stdout))
{
    Write-Host "7-Zip Information: $stdout"
    Write-Host "Completed!"
    # Clean up the zip file we moved and no longer need
    Remove-Item $newItemPath -Force
    ii $newItemFolder
}
