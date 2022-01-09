$inc_ver_num = Read-Host -Prompt "Increment the version number? (Y/N)?"
if ($inc_ver_num -like "y") {
    [xml]$axml= Get-Content "MXSPyCOM.csproj"
    $ver_num_parts = $axml.Project.PropertyGroup.Version.Split(".")
    $minor = [int]$ver_num_parts[-1] + 1
    $axml.Project.PropertyGroup.Version = $ver_num_parts[0] + "." + $minor.ToString()
    $axml.Save("MXSPyCOM.csproj")
}

dotnet publish "MXSPyCOM.csproj" -c Release

mkdir "$env:TEMP\MXSPyCOM" -Force
Copy-Item ".\bin\release\net6.0-windows\win-x64\publish\MXSPyCOM.exe" "$env:TEMP\MXSPyCOM"
Copy-Item ".\README.md" "$env:TEMP\MXSPyCOM"
Copy-Item ".\hello_world.ms" "$env:TEMP\MXSPyCOM"
Copy-Item ".\hello_world.py" "$env:TEMP\MXSPyCOM"
Copy-Item ".\initialize_COM_server.ms" "$env:TEMP\MXSPyCOM"

Compress-Archive -Path "$env:TEMP\MXSPyCOM\*" -DestinationPath "$env:USERPROFILE\DESKTOP\MXSPyCOM.zip" -Force

Remove-Item "$env:TEMP\MXSPyCOM" -Recurse -Force

Write-Output "Include the MXSPyCOM.zip file that was written to your desktop in the new binary release on GitHub."