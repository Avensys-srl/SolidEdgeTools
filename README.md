# SolidEdgeTools

[![Version](https://img.shields.io/badge/version-1.2.1.0-blue)](./SolidEdgeTools/My%20Project/AssemblyInfo.vb)
[![VB.NET](https://img.shields.io/badge/language-VB.NET-5C2D91)](./SolidEdgeTools)
[![.NET Framework](https://img.shields.io/badge/.NET%20Framework-4.6-512BD4)](./SolidEdgeTools/SolidEdgeTools.vbproj)
[![Platform](https://img.shields.io/badge/platform-Windows%20x86-0078D6)](./SolidEdgeTools/SolidEdgeTools.vbproj)
[![CAD](https://img.shields.io/badge/CAD-Solid%20Edge%20V20-0A7B83)](./handoff.md)
[![Interop](https://img.shields.io/badge/API-COM%20Interop-important)](./handoff.md)
[![Status](https://img.shields.io/badge/status-legacy%20refactoring-orange)](./handoff.md)

Legacy VB.NET WinForms utility for **Solid Edge V20** automation, evolved incrementally toward a more structured CAD automation engine while preserving existing production workflows.

## What It Does

The current tool already supports:

- BOM generation
- property table export to Excel
- DXF export for sheet metal parts
- STL / STP export
- DFT generation for sheet metal and 3D views
- optional automatic DFT layout for sheet metal views
- DFT publishing to PDF / DWG
- image export
- project property coding
- chained `Produzione Lamiera` workflow

This is still primarily a **production automation / output tool**, but the repository now includes the first architectural layers for stronger configuration-driven CAD automation.

## Technology

- Solid Edge V20 (UGS era)
- COM-based automation API
- VB.NET
- .NET Framework 4.6
- WinForms
- x86 target

## Repository Structure

```text
SolidEdgeTools.sln
handoff.md
README.md
SolidEdgeTools/
  Interop/
    SolidEdgeSessionHelpers.vb
  Models/
    BOMItem.vb
    ProductConfiguration.vb
    GeometryModels.vb
    WorkflowOptions.vb
  Services/
    BomService.vb
    ConfigurationEngine.vb
    ConfigurationValidator.vb
    DraftGenerationService.vb
    DraftPublishService.vb
    FilePropertyService.vb
    FlatDxfExportService.vb
    GeometryPlanService.vb
    ImageExportService.vb
    MaterialFilter.vb
    NeutralExportService.vb
    OccurrenceWalker.vb
    SolidEdgeWorkflowService.vb
    TemplateGeometryService.vb
  My Project/
  SET_MainForm.vb
  SET_MainForm.Designer.vb
  SolidEdgeTools.vbproj
```

## Current Architecture

Current practical layering:

```text
SET_MainForm
  -> ProductConfiguration / WorkflowOptions
  -> ConfigurationEngine / ConfigurationValidator
  -> SolidEdgeWorkflowService
  -> Output services
  -> Geometry planning scaffolding
```

Main reusable services already extracted from the original form:

- [`BomService`](./SolidEdgeTools/Services/BomService.vb)
- [`FlatDxfExportService`](./SolidEdgeTools/Services/FlatDxfExportService.vb)
- [`NeutralExportService`](./SolidEdgeTools/Services/NeutralExportService.vb)
- [`DraftGenerationService`](./SolidEdgeTools/Services/DraftGenerationService.vb)
- [`DraftPublishService`](./SolidEdgeTools/Services/DraftPublishService.vb)
- [`ImageExportService`](./SolidEdgeTools/Services/ImageExportService.vb)
- [`OccurrenceWalker`](./SolidEdgeTools/Services/OccurrenceWalker.vb)
- [`SolidEdgeWorkflowService`](./SolidEdgeTools/Services/SolidEdgeWorkflowService.vb)

First configuration / geometry foundation:

- [`ProductConfiguration`](./SolidEdgeTools/Models/ProductConfiguration.vb)
- [`ConfigurationEngine`](./SolidEdgeTools/Services/ConfigurationEngine.vb)
- [`ConfigurationValidator`](./SolidEdgeTools/Services/ConfigurationValidator.vb)
- [`GeometryPlanService`](./SolidEdgeTools/Services/GeometryPlanService.vb)
- [`TemplateGeometryService`](./SolidEdgeTools/Services/TemplateGeometryService.vb)

## Current Status

The repository is in an **incremental refactoring state**:

- working production workflows are preserved
- COM/session handling has been hardened pragmatically
- UI coupling has been reduced with typed option/config models
- orchestration has started moving out of the form
- geometry automation is only at the template-planning stage
- assembly composition and intent-driven engineering rules are not implemented yet

## Build Notes

This project is **not** a modern SDK-style .NET project.

Build requirements:

- Visual Studio with .NET Framework targeting pack
- Full MSBuild / Visual Studio build tools
- Windows environment
- Solid Edge COM dependencies available

Important:

- `dotnet build` / `dotnet msbuild` are not sufficient for this project because it uses legacy COM reference resolution
- target platform is `x86`

## Runtime Notes

This tool is designed around **Solid Edge V20-compatible COM automation patterns**.

Key constraints:

- no dependency on Synchronous Technology
- feature-based legacy API assumptions
- deterministic COM cleanup is critical
- document/session ownership must be handled carefully to avoid closing a user-owned Solid Edge session

## Refactoring Direction

The target direction is documented in [`handoff.md`](./handoff.md).

In short, the intended path is:

```text
UI -> Configuration -> Geometry -> Assembly -> Output
```

Current progress:

- service extraction from the WinForms monolith: largely completed for output workflows
- request / options objects: implemented
- COM/session hardening: implemented pragmatically
- workflow orchestration layer: implemented
- configuration layer: started and already used
- geometry layer: started as template-driven scaffolding
- `.psm` DFT auto-layout: first usable implementation
- assembly engine: not started
- engineering rule engine: not started

## Known Limitations

- no automated test suite yet
- no CI pipeline configured
- some legacy UI code and Excel late-binding remain
- PDF generation still depends on local environment behavior
- geometry generation is not yet driving real `.par` / `.psm` feature creation
- `.psm` DFT auto-layout is usable but still visually iterative on edge-on cases

## Near-Term Ideas

- reuse the most recent `Disegni di Piega_old` file for the same part when present, updating model views and preserving existing dimensions/annotations when possible
- add a dedicated `Produzione Plastiche` workflow using material `FabLab`, supplier BOM export with `Plastiche_` prefix, and STL export chain

## Debugging Notes

If Visual Studio reports missing symbols while debugging:

- ensure `Debug|x86` is used
- consider disabling `Just My Code`
- verify optimization settings for the Debug configuration if source-level debugging is required

## Compatibility Policy

When in doubt, prioritize:

- Solid Edge V20 compatibility
- COM-safe automation patterns
- deterministic cleanup
- stable reference strategies
- incremental refactoring over rewrite

## Ownership / License

This repository currently does **not** include a published open-source license file.

Until a `LICENSE` file is added, treat the codebase as proprietary/internal-use by default.
