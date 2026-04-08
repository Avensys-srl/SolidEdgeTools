Imports System.IO
Imports System.Windows.Forms

Public Class ImageExportService

    Private ReadOnly _imageExporter As Action(Of SolidEdgeFramework.Application, String, String)
    Private ReadOnly _errorHandler As Func(Of Exception, String, MessageBoxButtons, MessageBoxIcon, DialogResult)
    Private ReadOnly _occurrenceWalker As New OccurrenceWalker()

    Public Sub New(imageExporter As Action(Of SolidEdgeFramework.Application, String, String),
                   errorHandler As Func(Of Exception, String, MessageBoxButtons, MessageBoxIcon, DialogResult))

        _imageExporter = imageExporter
        _errorHandler = errorHandler
    End Sub

    Public Function ExportAssembly(seApplication As SolidEdgeFramework.Application,
                                   assembly As SolidEdgeAssembly.AssemblyDocument,
                                   options As ImageExportOptions) As Boolean

        Dim occurrenceFileNames As New Dictionary(Of String, Integer)

        Return _occurrenceWalker.Walk(
            assembly.Occurrences,
            options.IncludeSubAssemblies,
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

                Return ExportFile(seApplication, occurrenceFileNames, assembly.Path, options, item.OccurrenceFileName)
            End Function)
    End Function

    Private Function ExportFile(seApplication As SolidEdgeFramework.Application,
                                occurrenceFileNames As Dictionary(Of String, Integer),
                                rootAssemblyPath As String,
                                options As ImageExportOptions,
                                occurrenceFileName As String) As Boolean

        If occurrenceFileNames.ContainsKey(occurrenceFileName) Then
            Return True
        End If

        Do While True
            Try
                _imageExporter(seApplication,
                               occurrenceFileName,
                               Path.Combine(rootAssemblyPath,
                                            "image",
                                            options.Prefix & Path.ChangeExtension(Path.GetFileName(occurrenceFileName), "jpg")))

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

        Return False
    End Function
End Class
