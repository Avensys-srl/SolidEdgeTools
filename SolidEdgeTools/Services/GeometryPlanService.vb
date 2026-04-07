Imports System.IO

Public Class GeometryPlanService

    Public Function CreatePartPlan(configuration As ProductConfiguration,
                                   templatePath As String,
                                   outputPath As String) As GeometryBuildPlan

        Return BuildPlan(configuration, templatePath, outputPath, ".par")
    End Function

    Public Function CreateSheetMetalPlan(configuration As ProductConfiguration,
                                         templatePath As String,
                                         outputPath As String) As GeometryBuildPlan

        Return BuildPlan(configuration, templatePath, outputPath, ".psm")
    End Function

    Private Function BuildPlan(configuration As ProductConfiguration,
                               templatePath As String,
                               outputPath As String,
                               expectedExtension As String) As GeometryBuildPlan

        Dim plan As New GeometryBuildPlan() With {
            .TemplatePath = templatePath,
            .OutputPath = outputPath,
            .DocumentType = expectedExtension
        }

        If configuration IsNot Nothing AndAlso configuration.Unit IsNot Nothing Then
            plan.Variables.Add(New GeometryVariable() With {
                .Name = "Scale",
                .Value = configuration.Unit.Scale
            })
        End If

        Return plan
    End Function

    Public Function IsCompatibleTemplate(plan As GeometryBuildPlan) As Boolean
        If plan Is Nothing Then
            Return False
        End If

        If String.IsNullOrWhiteSpace(plan.TemplatePath) OrElse String.IsNullOrWhiteSpace(plan.OutputPath) Then
            Return False
        End If

        Return String.Equals(Path.GetExtension(plan.TemplatePath), plan.DocumentType, StringComparison.OrdinalIgnoreCase)
    End Function
End Class
