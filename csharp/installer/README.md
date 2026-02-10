# Installer - Odoo AutoCAD Integration (C#)

## Quick Start

```cmd
build.bat
```

## Build Options

| Flag | Description |
|------|-------------|
| `--installer-only` | Skip dotnet publish, repackage existing `publish/` dir |
| `--skip-tests` | Skip running unit tests |
| `--clean` | Clean `bin/obj/publish` before build |
| `--keep-symbols` | Keep debug symbols (PDB) in output |

```cmd
build.bat --clean --skip-tests
build.bat --installer-only
```

## Prerequisites

| Tool | Version | Download |
|------|---------|----------|
| .NET SDK | 8.0+ | https://dotnet.microsoft.com/download/dotnet/8.0 |
| Inno Setup | 6.x | https://jrsoftware.org/isdl.php |

## Build Flow

```
build.bat
  [1] Check prerequisites (.NET SDK, Inno Setup)
  [2] Clean bin/obj/publish (optional --clean)
  [3] Restore NuGet packages
  [3.5] Run tests (optional, skip with --skip-tests)
  [4] dotnet publish (Release, win-x64, self-contained)
  [5] Inno Setup ISCC -> setup.exe
```

## Output

```
csharp/
├── publish/                                  # 203 MB (self-contained .NET 8)
│   ├── OdooAutoCAD.exe                       # Main executable
│   ├── appsettings.json                      # Configuration
│   └── ...                                   # Runtime + dependencies
└── installer/
    └── odoo-autocad-integration-X.Y.Z-setup.exe   # ~63 MB installer
```

## Files

| File | Purpose |
|------|---------|
| `build.bat` | Build + package script |
| `odoo-autocad-csharp-setup.iss` | Inno Setup installer definition |
| `README.md` | This file |

## Runtime Requirements

- Windows 10/11 (64-bit)
- No .NET runtime needed (self-contained)
- AutoCAD 2020+ (optional, for COM integration)
- Odoo 14+ server (optional, for API integration)

## Versioning

Version is defined in `odoo-autocad-csharp-setup.iss`:
```ini
#define MyAppVersion "6.0.0"
```

Update this value before building a new release.
