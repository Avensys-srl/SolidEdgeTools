Imports System.IO

Public Class TemplateGeometryService

    Public Function CreateDocumentFromTemplate(seApplication As SolidEdgeFramework.Application,
                                               plan As GeometryBuildPlan) As Boolean

        If seApplication Is Nothing Then
            Throw New ArgumentNullException("seApplication")
        End If

        If plan Is Nothing Then
            Throw New ArgumentNullException("plan")
        End If

        If Not File.Exists(plan.TemplatePath) Then
            Throw New FileNotFoundException("Template file not found.", plan.TemplatePath)
        End If

        If Not Directory.Exists(Path.GetDirectoryName(plan.OutputPath)) Then
            Directory.CreateDirectory(Path.GetDirectoryName(plan.OutputPath))
        End If

        File.Copy(plan.TemplatePath, plan.OutputPath, True)

        ' V20-safe baseline: start with deterministic template cloning.
        ' Variable-driven edits can be layered here once stable named variables
        ' and production templates are defined.
        Return True
    End Function
End Class
