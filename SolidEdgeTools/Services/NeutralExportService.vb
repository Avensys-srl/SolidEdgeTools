Imports System.IO
Imports System.Windows.Forms

Public Class NeutralExportService

    Private ReadOnly _partExporter As Action(Of SolidEdgeFramework.Application, String, String)
    Private ReadOnly _sheetMetalExporter As Action(Of SolidEdgeFramework.Application, String, String)
    Private ReadOnly _errorHandler As Func(Of Exception, String, MessageBoxButtons, MessageBoxIcon, DialogResult)
    Private ReadOnly _occurrenceWalker As New OccurrenceWalker()

    Public Sub New(partExporter As Action(Of SolidEdgeFramework.Application, String, String),
                   sheetMetalExporter As Action(Of SolidEdgeFramework.Application, String, String),
                   errorHandler As Func(Of Exception, String, MessageBoxButtons, MessageBoxIcon, DialogResult))

        _partExporter = partExporter
        _sheetMetalExporter = sheetMetalExporter
        _errorHandler = errorHandler
    End Sub

    Public Function ExportAssembly(seApplication As SolidEdgeFramework.Application,
                                   assembly As SolidEdgeAssembly.AssemblyDocument,
                                   options As NeutralExportOptions,
                                   Optional progress As Action(Of Integer, Integer, String) = Nothing,
                                   Optional shouldCancel As Func(Of Boolean) = Nothing) As Boolean

        Dim targetFiles = GetTargetFiles(assembly, options)
        Dim processed As Integer = 0

        If progress IsNot Nothing Then
            progress(0, targetFiles.Count, "")
        End If

        For Each occurrenceFileName In targetFiles
            If shouldCancel IsNot Nothing AndAlso shouldCancel() Then
                Return False
            End If

            Dim exporter As Action(Of SolidEdgeFramework.Application, String, String) = Nothing
            Dim extension = Path.GetExtension(occurrenceFileName).ToLowerInvariant()

            If extension = ".par" Then
                exporter = _partExporter
            ElseIf extension = ".psm" Then
                exporter = _sheetMetalExporter
            End If

            If exporter Is Nothing Then
                Continue For
            End If

            If Not ExportFile(seApplication, assembly.Path, options, occurrenceFileName, exporter) Then
                Return False
            End If

            processed += 1

            If progress IsNot Nothing Then
                progress(processed, targetFiles.Count, occurrenceFileName)
            End If
        Next

        Return True
    End Function

    Private Function GetTargetFiles(assembly As SolidEdgeAssembly.AssemblyDocument,
                                    options As NeutralExportOptions) As List(Of String)

        Dim targetFiles As New List(Of String)
        Dim uniqueFiles As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

        _occurrenceWalker.Walk(
            assembly.Occurrences,
            True,
            Function(item)
                If item.Type <> SolidEdgeFramework.ObjectType.igPart Then
                    Return True
                End If

                Dim extension = Path.GetExtension(item.OccurrenceFileName).ToLowerInvariant()
                If extension <> ".par" AndAlso extension <> ".psm" Then
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
                                options As NeutralExportOptions,
                                occurrenceFileName As String,
                                exporter As Action(Of SolidEdgeFramework.Application, String, String)) As Boolean

        Do While True
            Try
                exporter(seApplication,
                         occurrenceFileName,
                         Path.Combine(rootAssemblyPath,
                                      options.ExportType,
                                      options.Prefix & Path.ChangeExtension(Path.GetFileName(occurrenceFileName), options.ExportType)))
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
