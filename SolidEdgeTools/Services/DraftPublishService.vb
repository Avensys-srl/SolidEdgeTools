Imports System.IO

Public Class DraftPublishService

    Private ReadOnly _pdfPrinterName As String

    Public Sub New(Optional pdfPrinterName As String = "Adobe PDF")
        _pdfPrinterName = pdfPrinterName
    End Sub

    Public Function PublishPdf(seApplication As SolidEdgeFramework.Application,
                               options As DraftPublishOptions) As Boolean

        Dim seDocuments As SolidEdgeFramework.Documents = Nothing
        Dim objDraft As SolidEdgeDraft.DraftDocument = Nothing

        Try
            seDocuments = seApplication.Documents

            For Each dftPath As String In Directory.GetFiles(options.InputDirectory, "*.dft")
                objDraft = seDocuments.Open(dftPath)

                Dim outPDFFilePath = Path.Combine(Path.GetDirectoryName(dftPath),
                                                  "Pdf",
                                                  Path.GetFileNameWithoutExtension(dftPath) + ".pdf")

                EnsureOutputPath(outPDFFilePath)

                objDraft.PrintOut(_pdfPrinterName,
                    Orientation:=Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6.PrinterObjectConstants.vbPRORLandscape)

                objDraft.Close()
                SolidEdgeSessionHelpers.ReleaseCOMReference(objDraft)
            Next
        Finally
            SolidEdgeSessionHelpers.ReleaseCOMReference(seDocuments)
        End Try

        Return True
    End Function

    Public Function PublishDwg(seApplication As SolidEdgeFramework.Application,
                               options As DraftPublishOptions) As Boolean

        Dim seDocuments As SolidEdgeFramework.Documents = Nothing
        Dim objDraft As SolidEdgeDraft.DraftDocument = Nothing

        Try
            seDocuments = seApplication.Documents

            For Each dftPath As String In Directory.GetFiles(options.InputDirectory, "*.dft")
                objDraft = seDocuments.Open(dftPath)

                Dim outDWGFilePath = Path.Combine(Path.GetDirectoryName(dftPath),
                                                  "DWG",
                                                  Path.GetFileNameWithoutExtension(dftPath) + ".dwg")

                EnsureOutputPath(outDWGFilePath)

                objDraft.SaveAs(outDWGFilePath)

                objDraft.Close()
                SolidEdgeSessionHelpers.ReleaseCOMReference(objDraft)
            Next
        Finally
            SolidEdgeSessionHelpers.ReleaseCOMReference(seDocuments)
        End Try

        Return True
    End Function

    Private Sub EnsureOutputPath(outputFilePath As String)
        If Not Directory.Exists(Path.GetDirectoryName(outputFilePath)) Then
            Directory.CreateDirectory(Path.GetDirectoryName(outputFilePath))
        End If

        If File.Exists(outputFilePath) Then
            File.Delete(outputFilePath)
        End If
    End Sub
End Class
