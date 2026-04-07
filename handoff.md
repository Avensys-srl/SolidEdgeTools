# Solid Edge V20 Automation Tool - Handoff Document

## 1. Objective

Design and evolve an existing **VB.NET Solid Edge automation tool** into a **modular CAD automation engine** capable of:

- Generating and modifying 3D geometry (Part / Sheet Metal)
- Creating assemblies programmatically
- Applying engineering rules (DFP, HVAC logic)
- Automating BOM, DXF and publishing workflows

---

## 2. Technology Context

### CAD System
- Solid Edge V20 (UGS era)
- COM-based API (Automation)

### Programming Stack
- VB.NET (.NET Framework 2.0-4.x compatible)
- COM Interop

### API Access Pattern

```vbnet
Dim app As SolidEdgeFramework.Application
app = Marshal.GetActiveObject("SolidEdge.Application")
```

---

## 3. Core API Concepts

### 3.1 Object Model

```text
Application
  `-- Documents
       |-- PartDocument
       |-- SheetMetalDocument
       |-- AssemblyDocument
       `-- DraftDocument
```

---

### 3.2 Geometry Creation Workflow

1. Create document
2. Create reference planes
3. Create profile (2D sketch)
4. Add geometry (lines, arcs, etc.)
5. Apply relationships
6. Create feature (extrude, cut, etc.)

---

### 3.3 Assembly Workflow

- Insert components (Occurrences)
- Position components
- Apply constraints

---

### 3.4 COM Constraint (CRITICAL)

```vbnet
Marshal.ReleaseComObject(obj)
obj = Nothing
```

Failure causes:
- memory leaks
- Solid Edge crashes
- locked documents

---

## 4. Existing Tool (Assumed)

Current VB.NET tool already handles:

- BOM extraction
- DXF generation
- File handling / naming

Current nature: **post-processing tool**

---

## 4.1 Current Repository Status (Updated)

The repository is no longer only a single-form utility.

The original WinForms application still exists, but the ongoing refactoring has already extracted a first service layer around the legacy codebase.

Current extracted modules in the repository:

- `Interop/SolidEdgeSessionHelpers.vb`
- `Models/BOMItem.vb`
- `Services/FilePropertyService.vb`
- `Services/MaterialFilter.vb`
- `Services/BomService.vb`
- `Services/NeutralExportService.vb`
- `Services/FlatDxfExportService.vb`
- `Services/DraftGenerationService.vb`
- `Services/DraftPublishService.vb`
- `Services/ImageExportService.vb`
- `Services/OccurrenceWalker.vb`

Current practical state:

- UI still starts from `SET_MainForm`
- major export and BOM workflows are already delegated to services
- recursive assembly traversal has been centralized for most workflows
- geometry-generation capability is still not implemented
- direct control reads from the form are still present and should be replaced later with request/config objects

This means the codebase is now in an **incremental transition state**:

- not yet a CAD automation engine
- no longer just a monolithic form-based utility

---

## 5. Target Architecture

Transform into a:

### CAD Automation Engine

---

## 6. Proposed Architecture

### 6.1 High-Level Flow

```text
[UI / Input]
     |
     v
[Configuration Engine]
     |
     v
[Geometry Engine]
     |
     v
[Assembly Engine]
     |
     v
[Output Engine]
```

---

### 6.2 Modules

#### Product Configuration Engine
- Interprets input
- Applies engineering rules
- Builds logical model

Example:

```vbnet
If Unit.Type = "HRU" And Unit.Connection = "OSC" Then
    UseCounterFlowLayout()
End If
```

---

#### Geometry Engine
- Creates or modifies parts
- Handles sketches, features, holes, etc.

---

#### Assembly Engine
- Builds full assembly
- Inserts and constrains components

---

#### Output Engine
- BOM
- DXF
- Naming
- Revision

---

## 7. Design Strategy

### Avoid
- referencing random faces
- unstable geometry references

### Use
- named reference planes
- coordinate systems
- variables
- stable sketches

---

## 8. Automation Levels

| Level | Description |
|------:|-------------|
| 1 | Parametric control |
| 2 | Configurable geometry |
| 3 | Generated geometry |
| 4 | Full automation |

---

## 9. Coding Guidelines

### Always use Try/Finally

```vbnet
Try
   ' logic
Finally
   Marshal.ReleaseComObject(obj)
End Try
```

---

### Use Wrapper Classes

```vbnet
Class SEApplication
    Public App As SolidEdgeFramework.Application
End Class
```

---

### Avoid COM spread

Centralize:
- object creation
- document access

### Incremental Refactoring Rule

Do not rewrite working production workflows unnecessarily.

Preferred approach:

1. Extract logic from `SET_MainForm`
2. Preserve public behavior
3. Introduce service classes
4. Introduce typed request/config objects
5. Only then introduce stronger geometry / assembly automation

### Current Refactoring Boundary

The following workflows should now be treated as service-owned logic and extended there rather than re-expanded inside the form:

- BOM generation
- STL/STP export
- DXF export
- DFT generation
- DFT publish to PDF/DWG
- image export
- occurrence traversal

The form should progressively become:

- input collection
- command orchestration
- user feedback

Not the main location for CAD/business logic.

---

## 10. Suggested Class Structure

```vbnet
Class UnitModel
    Public Width As Double
    Public Height As Double
    Public Depth As Double
    Public Configuration As String
End Class

Class GeometryBuilder
    Sub CreatePanel(model As UnitModel)
    End Sub
End Class

Class AssemblyBuilder
    Sub BuildUnit(model As UnitModel)
    End Sub
End Class

Class ExportService
    Sub ExportDXF()
    Sub ExportBOM()
End Class
```

---

### 10.1 Current Implemented Structure

The repository already contains a first practical structure aligned with the above direction:

```text
SET_MainForm
  -> SolidEdgeSessionHelpers
  -> FilePropertyService
  -> MaterialFilter
  -> BomService
  -> NeutralExportService
  -> FlatDxfExportService
  -> DraftGenerationService
  -> DraftPublishService
  -> ImageExportService
  -> OccurrenceWalker
```

This should be considered the current baseline for future refactoring.

Future additions should build on this structure rather than reintroducing logic into the form.

---

## 11. Execution Flow

```text
INPUT -> CONFIG -> GEOMETRY -> ASSEMBLY -> EXPORT
```

Example:

```vbnet
Dim unit As New UnitModel(...)
Dim config = ConfigEngine.Process(unit)

GeometryEngine.Build(config)
AssemblyEngine.Build(config)

Exporter.Run()
```

---

## 12. Capabilities

Solid Edge V20 API supports:

- geometry creation
- assembly building
- document control
- property management

---

## 13. Constraints (V20)

- No Synchronous Technology
- Feature-based only
- Older API

Still fully usable for advanced automation.

---

## 14. Expected Outcome

- Reduced manual CAD work
- Automated unit generation
- Consistent engineering rules
- Fewer production errors

---

## 15. Open Questions

Codex must analyze:

1. Current project structure
2. Entry point
3. BOM/DXF implementation
4. Template usage
5. Naming conventions
6. File storage

---

## 16. Next Steps

Codex should:

1. Analyze VB.NET project
2. Extract reusable logic
3. Refactor architecture
4. Implement Geometry Engine
5. Integrate existing export logic

---

## 16.1 Updated Next Steps After Current Refactoring

The first extraction step is already underway and partially completed.

Recommended next sequence from the current repository state:

1. Introduce typed request/options classes to remove direct UI-control coupling from workflows
2. Harden COM lifecycle management for transient COM objects inside technical export routines
3. Consolidate Solid Edge session ownership rules so the tool does not accidentally close a user-owned session
4. Introduce a configuration layer for product/input intent
5. Add template-driven geometry services for `.par` / `.psm`
6. Add assembly-composition services only after stable template and configuration patterns exist

Near-term priority:

- stabilize architecture
- reduce COM risk
- preserve current production outputs

Before any major geometry automation:

- named references
- template discipline
- deterministic document lifecycle

must be in place.

---

## Final Note

This document is designed to support:

- industrial CAD automation
- HVAC product logic
- scalable architecture

Not a tutorial, but a **foundation for a production-ready CAD engine**.

---

## Reference Material

Use the following references when reasoning about Solid Edge API and COM behavior:

- https://support.industrysoftware.automation.siemens.com/trainings/se/107/api/ProgrammersGuide.html
- https://support.industrysoftware.automation.siemens.com/trainings/se/106/api/SolidEdgeFramework~Application.html
- https://support.industrysoftware.automation.siemens.com/trainings/se/106/api/SolidEdgePart_P.html
- https://support.industrysoftware.automation.siemens.com/trainings/se/106/api/SolidEdgeAssembly_P.html
- https://support.industrysoftware.automation.siemens.com/trainings/se/106/api/SolidEdgeDraft_P.html
- https://support.industrysoftware.automation.siemens.com/trainings/se/107/api/V20SP11-WhatsNew.html
- https://learn.microsoft.com/en-us/dotnet/api/system.runtime.interopservices.marshal.releasecomobject
- https://learn.microsoft.com/en-us/dotnet/standard/native-interop/cominterop

When there is ambiguity, prioritize compatibility with Solid Edge V20 (COM-based API, no synchronous technology).
