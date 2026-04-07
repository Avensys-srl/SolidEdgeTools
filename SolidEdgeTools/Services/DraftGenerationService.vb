Imports System.IO
Imports System.Windows.Forms

Public Class DraftGenerationService

    Private ReadOnly _draftExporter As Action(Of SolidEdgeFramework.Application, String, String)
    Private ReadOnly _errorHandler As Func(Of Exception, String, MessageBoxButtons, MessageBoxIcon, DialogResult)
    Private ReadOnly _occurrenceWalker As New OccurrenceWalker()

    Public Sub New(draftExporter As Action(Of SolidEdgeFramework.Application, String, String),
                   errorHandler As Func(Of Exception, String, MessageBoxButtons, MessageBoxIcon, DialogResult))

        _draftExporter = draftExporter
        _errorHandler = errorHandler
    End Sub

    Public Function GenerateForAssembly(seApplication As SolidEdgeFramework.Application,
                                        assembly As SolidEdgeAssembly.AssemblyDocument,
                                        options As DraftGenerationOptions) As Boolean

        Dim processedFiles As New Dictionary(Of String, Integer)

        Return _occurrenceWalker.Walk(
            assembly.Occurrences,
            True,
            Function(item)
                If item.Type <> SolidEdgeFramework.ObjectType.igPart Then
                    Return True
                End If

                Dim extension = Path.GetExtension(item.OccurrenceFileName)
                If extension <> ".psm" AndAlso extension <> ".par" Then
                    Return True
                End If

                If Not MaterialFilter.MatchesSelectedMaterial(FilePropertyService.GetPropertyValue(item.OccurrenceFileName, "MechanicalModeling", "Material"), options.MaterialSelection.SelectedMaterials) Then
                    Return True
                End If

                Return ExportFile(seApplication, processedFiles, assembly.Path, options, item.OccurrenceFileName)
            End Function)
    End Function

    Private Function ExportFile(seApplication As SolidEdgeFramework.Application,
                                processedFiles As Dictionary(Of String, Integer),
                                rootAssemblyPath As String,
                                options As DraftGenerationOptions,
                                occurrenceFileName As String) As Boolean

        If processedFiles.ContainsKey(occurrenceFileName) Then
            Return True
        End If

        Do While True
            Try
                _draftExporter(seApplication,
                               BuildOutputPath(rootAssemblyPath, options.Prefix, occurrenceFileName),
                               occurrenceFileName)

                processedFiles.Add(occurrenceFileName, 0)
                Return True
            Catch ex As Exception
                Select Case _errorHandler(ex,
                                          String.Format("Errore durante la generazione {0}.", occurrenceFileName),
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

    Private Function BuildOutputPath(rootAssemblyPath As String,
                                     prefix As String,
                                     occurrenceFileName As String) As String

        Dim targetFolder As String

        If Path.GetExtension(occurrenceFileName) = ".psm" Then
            targetFolder = "Disegni di Piega"
        Else
            targetFolder = "Viste 3D"
        End If

        Return Path.Combine(rootAssemblyPath,
                            targetFolder,
                            prefix & Path.ChangeExtension(Path.GetFileName(occurrenceFileName), "dft"))
    End Function
End Class
