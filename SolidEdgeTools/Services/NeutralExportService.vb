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
                                   options As NeutralExportOptions) As Boolean

        Dim occurrenceFileNames As New Dictionary(Of String, Integer)

        Return _occurrenceWalker.Walk(
            assembly.Occurrences,
            True,
            Function(item)
                If item.Type <> SolidEdgeFramework.ObjectType.igPart Then
                    Return True
                End If

                If Not MaterialFilter.MatchesSelectedMaterial(FilePropertyService.GetPropertyValue(item.OccurrenceFileName, "MechanicalModeling", "Material"), options.MaterialSelection.SelectedMaterials) Then
                    Return True
                End If

                If Path.GetExtension(item.OccurrenceFileName) = ".par" Then
                    Return ExportFile(seApplication, occurrenceFileNames, assembly.Path, options, item.OccurrenceFileName, _partExporter)
                End If

                If Path.GetExtension(item.OccurrenceFileName) = ".psm" Then
                    Return ExportFile(seApplication, occurrenceFileNames, assembly.Path, options, item.OccurrenceFileName, _sheetMetalExporter)
                End If

                Return True
            End Function)
    End Function

    Private Function ExportFile(seApplication As SolidEdgeFramework.Application,
                                occurrenceFileNames As Dictionary(Of String, Integer),
                                rootAssemblyPath As String,
                                options As NeutralExportOptions,
                                occurrenceFileName As String,
                                exporter As Action(Of SolidEdgeFramework.Application, String, String)) As Boolean

        If occurrenceFileNames.ContainsKey(occurrenceFileName) Then
            Return True
        End If

        Do While True
            Try
                exporter(seApplication,
                         occurrenceFileName,
                         Path.Combine(rootAssemblyPath,
                                      options.ExportType,
                                      options.Prefix & Path.ChangeExtension(Path.GetFileName(occurrenceFileName), options.ExportType)))

                occurrenceFileNames.Add(occurrenceFileName, 0)
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
    End Function
End Class
