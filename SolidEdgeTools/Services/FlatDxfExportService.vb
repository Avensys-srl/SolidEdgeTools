Imports System.IO
Imports System.Windows.Forms

Public Class FlatDxfExportService

    Private ReadOnly _dxfExporter As Action(Of SolidEdgeFramework.Application, String, String)
    Private ReadOnly _errorHandler As Func(Of Exception, String, MessageBoxButtons, MessageBoxIcon, DialogResult)
    Private ReadOnly _occurrenceWalker As New OccurrenceWalker()

    Public Sub New(dxfExporter As Action(Of SolidEdgeFramework.Application, String, String),
                   errorHandler As Func(Of Exception, String, MessageBoxButtons, MessageBoxIcon, DialogResult))

        _dxfExporter = dxfExporter
        _errorHandler = errorHandler
    End Sub

    Public Function ExportAssembly(seApplication As SolidEdgeFramework.Application,
                                   assembly As SolidEdgeAssembly.AssemblyDocument,
                                   options As FlatDxfExportOptions,
                                   Optional progress As Action(Of Integer, Integer, String) = Nothing,
                                   Optional shouldCancel As Func(Of Boolean) = Nothing) As Boolean

        Dim targetFiles = GetTargetFiles(assembly, options)
        Return ExportFiles(seApplication, assembly.Path, options, targetFiles, progress, shouldCancel)
    End Function

    Public Function ExportFiles(seApplication As SolidEdgeFramework.Application,
                                rootAssemblyPath As String,
                                options As FlatDxfExportOptions,
                                targetFiles As IEnumerable(Of String),
                                Optional progress As Action(Of Integer, Integer, String) = Nothing,
                                Optional shouldCancel As Func(Of Boolean) = Nothing) As Boolean

        Dim resolvedTargets = New List(Of String)(targetFiles)
        Dim processed As Integer = 0

        If progress IsNot Nothing Then
            progress(0, resolvedTargets.Count, "")
        End If

        For Each occurrenceFileName In resolvedTargets
            If shouldCancel IsNot Nothing AndAlso shouldCancel() Then
                Return False
            End If

            If Not ExportFile(seApplication, rootAssemblyPath, options, occurrenceFileName) Then
                Return False
            End If

            processed += 1

            If progress IsNot Nothing Then
                progress(processed, resolvedTargets.Count, occurrenceFileName)
            End If
        Next

        Return True
    End Function

    Private Function GetTargetFiles(assembly As SolidEdgeAssembly.AssemblyDocument,
                                    options As FlatDxfExportOptions) As List(Of String)

        Dim targetFiles As New List(Of String)
        Dim uniqueFiles As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

        _occurrenceWalker.Walk(
            assembly.Occurrences,
            options.IncludeSubAssemblies,
            Function(item)
                If item.Type <> SolidEdgeFramework.ObjectType.igPart Then
                    Return True
                End If

                If Path.GetExtension(item.OccurrenceFileName).ToLowerInvariant() <> ".psm" Then
                    Return True
                End If

                If Not MaterialFilter.MatchesSelectedMaterial(FilePropertyService.GetPropertyValue(item.OccurrenceFileName, "MechanicalModeling", "Material"), options.MaterialSelection.SelectedMaterials) Then
                    Return True
                End If

                If uniqueFiles.Add(item.OccurrenceFileName) Then
                    targetFiles.Add(item.OccurrenceFileName)
                End If

                Return True
            End Function)

        Return targetFiles
    End Function

    Private Function ExportFile(seApplication As SolidEdgeFramework.Application,
                                rootAssemblyPath As String,
                                options As FlatDxfExportOptions,
                                occurrenceFileName As String) As Boolean

        Do While True
            Try
                _dxfExporter(seApplication,
                             occurrenceFileName,
                             Path.Combine(rootAssemblyPath,
                                          "dxf",
                                          options.Prefix & Path.ChangeExtension(Path.GetFileName(occurrenceFileName), "dxf")))
                Return True
            Catch ex As Exception
                Select Case _errorHandler(ex,
                                          String.Format("Errore durante l'esportazione {0}.", occurrenceFileName),
                                          MessageBoxButtons.AbortRetryIgnore,
                                          MessageBoxIcon.Error)
                    Case DialogResult.Ignore
                        Return True
                    Case DialogResult.Abort
                        Return False
                End Select
            End Try
        Loop

        Return False
    End Function
End Class
