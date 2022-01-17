$inc_ver_num = Read-Host -Prompt "Increment the version number? (Y/N)?"
if ($inc_ver_num -like "y") {
    [xml]$axml= Get-Content "MXSPyCOM.csproj"
    $ver_num_parts = $axml.Project.PropertyGroup.Version.Split(".")
    $minor = [int]$ver_num_parts[-1] + 1
    $axml.Project.PropertyGroup.Version = $ver_num_parts[0] + "." + $minor.ToString()
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
   Write-Output "Include the MXSPyCOM.zip file that was written to your desktop in the new binary release on GitHub."
}