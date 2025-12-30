# Plant 3D JSON Exporter

Exports Plant 3D project drawings and entity data to a JSON file via an AutoCAD Plant 3D command.

## Requirements
- Windows
- AutoCAD Plant 3D 2026 installed
- Plant 3D SDK headers/assemblies available locally
- .NET SDK installed (net8.0 target)

## Build
```powershell
cd addin
dotnet build -c Release
```

## .NET version used
This project targets `net8.0-windows` (see `addin/PlantJsonExporter.csproj`).  
If you have multiple SDKs installed, the build uses the latest SDK available unless a `global.json` pins a version. With your setup, the SDK 10.0.101 is used to build a net8.0 target.

## SDK reference path
Update the SDK path if needed in:
- `addin/PlantJsonExporter.csproj` (`PlantSdkDir`)
