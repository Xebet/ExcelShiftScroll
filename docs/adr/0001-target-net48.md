# ADR 0001: Target .NET Framework 4.8

- Status: Accepted
- Date: 2026-09-03

## Context

Excel-DNA 1.9.0 is the current stable release. Its official runtime guide supports .NET Framework 4.x and .NET 6 or newer. The guide recommends net472/net48 for broad distribution and notes that modern .NET requires a matching x86/x64 Desktop Runtime and permits only one modern .NET runtime major in an Excel process.

Sources checked before implementation:

- <https://excel-dna.net/docs/guides-basic/dotnet-runtime-support/>
- <https://github.com/Excel-DNA/ExcelDna/releases>
- <https://www.nuget.org/packages/ExcelDna.AddIn/1.9.0>

## Decision

Use an SDK-style C# project targeting `net48`, with `Microsoft.NETFramework.ReferenceAssemblies` as a private build dependency. Build with a current .NET SDK. Package x86 and x64 single-file XLLs through Excel-DNA 1.9.0.

## Reasons

- .NET Framework 4.8 is serviced as a Windows component, so users do not install another runtime.
- Separate AppDomains give mature isolation from other Excel-DNA/.NET Framework add-ins.
- It avoids modern .NET runtime-major conflicts with unrelated Excel add-ins.
- Win32, COM `dynamic`, Ribbon, settings, and tests need no modern runtime-only feature.
- The project can still use modern C# compiler features through the SDK build.

## Consequences

- The code cannot depend on APIs available only in modern .NET.
- Build agents require a .NET SDK, but the reference-assemblies package avoids machine-global targeting-pack assumptions.
- If future Windows or Excel support changes, a modern .NET migration will require a new ADR and coexistence testing.
