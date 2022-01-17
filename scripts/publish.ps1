# Build script for MXSPyCOM.
# MXSPyCOM.exe is built in release, with the version number incremented by 0.01, if the user requests.
# MXSPyCOM.exe , the demonstration HelloWorld Python and MaxScript scripts, the initialize_COM_server MaxScript, and 
# the readme.md file are archived in an MXSPyCOM.zip file and placed on the desktop, ready for including in a release
# in the GitHub repository.

[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

$inc_ver_num = [System.Windows.Forms.MessageBox]::Show(
    "Increment the version number?", "Publish MXSPyCOM", [System.Windows.Forms.MessageBoxButtons]::YesNo)

if ($inc_ver_num -like "yes") {
    [xml]$axml= Get-Content "MXSPyCOM.csproj"
    $ver_num_parts = $axml.Project.PropertyGroup.Version.Split(".")
    $major = [int]$ver_num_parts[0]
    $minor = [int]$ver_num_parts[-1] + 1
    if ($minor -gt 99) {
        $major += 1
        $minor = 00
    }
    $axml.Project.PropertyGroup.Version = $major.ToString() + "." + $minor.ToString()
    $axml.Save("MXSPyCOM.csproj")
}

dotnet publish "MXSPyCOM.csproj" -c Release

$temp_path = Join-Path $env:TEMP "MXSPyCOM"
mkdir $temp_path -Force
Copy-Item "bin\release\net6.0-windows\win-x64\publish\MXSPyCOM.exe" $temp_path
Copy-Item "README.md" $temp_path
Copy-Item "hello_world.ms" $temp_path
Copy-Item "hello_world.py" $temp_path
Copy-Item "initialize_COM_server.ms" $temp_path

$temp_filepath = Join-Path $temp_path "*"
try {
    # The user's desktop is in the default location on their hard drive.
    $archive_filepath = Join-Path $env:USERPROFILE "DESKTOP" "MXSPyCOM.zip"
    Compress-Archive -Path $temp_filepath -DestinationPath $archive_filepath -Force
}
catch { # It would be better if a specific Compress-Archive exception is caught, but none can be found in docs.
    # The user's desktop is being managed by OneDrive.
    $archive_filepath = Join-Path $env:ONEDRIVE "Desktop" "MXSPyCOM.zip"
    Compress-Archive -Path $temp_filepath -DestinationPath $archive_filepath -Force
}
finally {
   Remove-Item $temp_path -Recurse -Force
   [System.Windows.Forms.MessageBox]::Show(
       "Finished.
       
Include the MXSPyCOM.zip file that was written to the desktop 
in the new binary release on GitHub.", 
       "Publish MXSPyCOM")
}