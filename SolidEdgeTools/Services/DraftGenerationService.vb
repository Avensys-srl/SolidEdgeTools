Imports System.IO
Imports System.Windows.Forms

Public Class DraftGenerationService

    Private ReadOnly _materialMatcher As Func(Of String, Boolean)
    Private ReadOnly _draftExporter As Action(Of SolidEdgeFramework.Application, String, String)
    Private ReadOnly _errorHandler As Func(Of Exception, String, MessageBoxButtons, MessageBoxIcon, DialogResult)

    Public Sub New(materialMatcher As Func(Of String, Boolean),
                   draftExporter As Action(Of SolidEdgeFramework.Application, String, String),
                   errorHandler As Func(Of Exception, String, MessageBoxButtons, MessageBoxIcon, DialogResult))

        _materialMatcher = materialMatcher
        _draftExporter = draftExporter
        _errorHandler = errorHandler
    End Sub

    Public Function GenerateForAssembly(seApplication As SolidEdgeFramework.Application,
                                        assembly As SolidEdgeAssembly.AssemblyDocument,
                                        prefix As String) As Boolean

        Dim processedFiles As New Dictionary(Of String, Integer)

        Return ScanNode(seApplication, processedFiles, assembly, assembly.Occurrences, prefix)
    End Function

    Private Function ScanNode(seApplication As SolidEdgeFramework.Application,
                              processedFiles As Dictionary(Of String, Integer),
                              rootAssembly As SolidEdgeAssembly.AssemblyDocument,
                              occurrences As SolidEdgeAssembly.Occurrences,
                              prefix As String) As Boolean

        For Each item As SolidEdgeAssembly.Occurrence In occurrences
            Select Case item.Type
                Case SolidEdgeFramework.ObjectType.igSubAssembly
                    If Not ScanNode(seApplication, processedFiles, rootAssembly, item.OccurrenceDocument.Occurrences, prefix) Then
                        Return False
                    End If

                Case SolidEdgeFramework.ObjectType.igPart
                    Dim extension = Path.GetExtension(item.OccurrenceFileName)

                    If extension = ".psm" OrElse extension = ".par" Then
                        If _materialMatcher(FilePropertyService.GetPropertyValue(item.OccurrenceFileName, "MechanicalModeling", "Material")) Then
                            If Not processedFiles.ContainsKey(item.OccurrenceFileName) Then
                                Do While True
                                    Try
                                        _draftExporter(seApplication,
                                                       BuildOutputPath(rootAssembly.Path, prefix, item.OccurrenceFileName),
                                                       item.OccurrenceFileName)

                                        processedFiles.Add(item.OccurrenceFileName, 0)
                                        Exit Do
                                    Catch ex As Exception
                                        Select Case _errorHandler(ex,
                                                                  String.Format("Errore durante la generazione {0}.", item.OccurrenceFileName),
                                                                  MessageBoxButtons.AbortRetryIgnore,
                                                                  MessageBoxIcon.Error)
                                            Case DialogResult.Ignore
                                                Exit Do
                                            Case DialogResult.Abort
                                                Return False
                                        End Select
                                    End Try
                                Loop
                            End If
                        End If
                    End If
            End Select
        Next

        Return True
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
