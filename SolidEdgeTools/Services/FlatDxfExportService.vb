Imports System.IO
Imports System.Windows.Forms

Public Class FlatDxfExportService

    Private ReadOnly _materialMatcher As Func(Of String, Boolean)
    Private ReadOnly _dxfExporter As Action(Of SolidEdgeFramework.Application, String, String)
    Private ReadOnly _errorHandler As Func(Of Exception, String, MessageBoxButtons, MessageBoxIcon, DialogResult)

    Public Sub New(materialMatcher As Func(Of String, Boolean),
                   dxfExporter As Action(Of SolidEdgeFramework.Application, String, String),
                   errorHandler As Func(Of Exception, String, MessageBoxButtons, MessageBoxIcon, DialogResult))

        _materialMatcher = materialMatcher
        _dxfExporter = dxfExporter
        _errorHandler = errorHandler
    End Sub

    Public Function ExportAssembly(seApplication As SolidEdgeFramework.Application,
                                   assembly As SolidEdgeAssembly.AssemblyDocument,
                                   prefix As String,
                                   includeSubAssemblies As Boolean) As Boolean

        Dim occurrenceFileNames As New Dictionary(Of String, Integer)

        Return ScanNode(seApplication, occurrenceFileNames, assembly, assembly.Occurrences, prefix, includeSubAssemblies)
    End Function

    Private Function ScanNode(seApplication As SolidEdgeFramework.Application,
                              occurrenceFileNames As Dictionary(Of String, Integer),
                              rootAssembly As SolidEdgeAssembly.AssemblyDocument,
                              occurrences As SolidEdgeAssembly.Occurrences,
                              prefix As String,
                              includeSubAssemblies As Boolean) As Boolean

        For Each item As SolidEdgeAssembly.Occurrence In occurrences
            Select Case item.Type
                Case SolidEdgeFramework.ObjectType.igSubAssembly
                    If includeSubAssemblies Then
                        If Not ScanNode(seApplication, occurrenceFileNames, rootAssembly, item.OccurrenceDocument.Occurrences, prefix, includeSubAssemblies) Then
                            Return False
                        End If
                    End If

                Case SolidEdgeFramework.ObjectType.igPart
                    If Path.GetExtension(item.OccurrenceFileName) = ".psm" Then
                        If _materialMatcher(FilePropertyService.GetPropertyValue(item.OccurrenceFileName, "MechanicalModeling", "Material")) Then
                            If Not occurrenceFileNames.ContainsKey(item.OccurrenceFileName) Then
                                Do While True
                                    Try
                                        _dxfExporter(seApplication,
                                                     item.OccurrenceFileName,
                                                     Path.Combine(rootAssembly.Path,
                                                                  "dxf",
                                                                  prefix & Path.ChangeExtension(Path.GetFileName(item.OccurrenceFileName), "dxf")))

                                        occurrenceFileNames.Add(item.OccurrenceFileName, 0)
                                        Exit Do
                                    Catch ex As Exception
                                        Select Case _errorHandler(ex,
                                                                  String.Format("Errore durante l'esportazione {0}.", item.OccurrenceFileName),
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
End Class
