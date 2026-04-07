# Solid Edge V20 Automation Tool – Handoff Document

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
- VB.NET (.NET Framework 2.0–4.x compatible)
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
 └── Documents
      ├── PartDocument
      ├── SheetMetalDocument
      ├── AssemblyDocument
      └── DraftDocument
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

👉 Current nature: **post-processing tool**

---

## 5. Target Architecture

Transform into a:

### CAD Automation Engine

---

## 6. Proposed Architecture

### 6.1 High-Level Flow

```text
[UI / Input]
     ↓
[Configuration Engine]
     ↓
[Geometry Engine]
     ↓
[Assembly Engine]
     ↓
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

## 11. Execution Flow

```text
INPUT → CONFIG → GEOMETRY → ASSEMBLY → EXPORT
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

✔ Still fully usable for advanced automation

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

## Final Note

This document is designed to support:

- industrial CAD automation
- HVAC product logic
- scalable architecture

Not a tutorial, but a **foundation for a production-ready CAD engine**.
