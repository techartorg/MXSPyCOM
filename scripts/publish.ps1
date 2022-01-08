dotnet publish "MXSPyCOM.csproj" -c Release

mkdir "$env:TEMP\MXSPyCOM" -Force
Copy-Item ".\bin\release\net6.0-windows\win-x64\publish\MXSPyCOM.exe" "$env:TEMP\MXSPyCOM"
Copy-Item ".\README.md" "$env:TEMP\MXSPyCOM"
Copy-Item ".\hello_world.ms" "$env:TEMP\MXSPyCOM"
Copy-Item ".\hello_world.py" "$env:TEMP\MXSPyCOM"
Copy-Item ".\initialize_COM_server.ms" "$env:TEMP\MXSPyCOM"

Compress-Archive -Path "$env:TEMP\MXSPyCOM\*" -DestinationPath "$env:USERPROFILE\DESKTOP\MXSPyCOM.zip" -Force

Remove-Item "$env:TEMP\MXSPyCOM" -Recurse -Force