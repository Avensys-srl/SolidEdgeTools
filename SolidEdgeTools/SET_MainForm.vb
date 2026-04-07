Imports SolidEdgeCommunity.Extensions
Imports System.Runtime.InteropServices
Imports System.IO

Public Class SET_MainForm


#Region "====[ Generate BOM ]===="

    Private Sub btnPropTable_Click(sender As System.Object, e As System.EventArgs) Handles btnPropTable.Click

        Dim objApp As SolidEdgeFramework.Application = Nothing
        Dim objDocuments As SolidEdgeFramework.Documents = Nothing
        Dim objDocument As SolidEdgeFramework.SolidEdgeDocument = Nothing
        Dim xlsArray(100, 3) As String
        Dim index As Integer = 0
        Dim objPropSets As SolidEdgeFileProperties.PropertySets = New SolidEdgeFileProperties.PropertySets
        Dim objProp As SolidEdgeFileProperties.Property = Nothing
        Dim objProps As SolidEdgeFileProperties.Properties = Nothing

        Try
            If ofdSelectPSMFile.ShowDialog() = Windows.Forms.DialogResult.OK Then
                sfdSelectXLSFile.FileName = Prefisso.Text + "Proprietà_" + Path.GetFileNameWithoutExtension(ofdSelectPSMFile.FileName)
                If sfdSelectXLSFile.ShowDialog() = Windows.Forms.DialogResult.OK Then


                    ' Register with OLE to handle concurrency issues on the current thread.
                    SolidEdgeCommunity.OleMessageFilter.Register()
                    ' Connect to or start Solid Edge.
                    objApp = SolidEdgeCommunity.SolidEdgeUtils.Connect(True, True)
                    ' Make Solid Edge visible
                    objApp.Visible = True 'se_off.Checked --> sembra non salvi in background
                    ' Turn off alerts. Weldment environment will display a warning
                    objApp.DisplayAlerts = True
                    ' Get a reference to the Documents collection
                    objDocuments = objApp.Documents
                    ' Create an instance of each document environment
                    Dim sDocument As String = ofdSelectPSMFile.FileName

                    Prefisso.Text = ofdSelectPSMFile.FileName



                    objPropSets.Open(sDocument, True)

                    xlsArray.SetValue("Classe", index, 0)
                    xlsArray.SetValue("Proprietà", index, 1)
                    xlsArray.SetValue("Valore", index, 2)



                    For Each objProps In objPropSets
                        For Each objProp In objProps
                            index = index + 1
                            xlsArray.SetValue(IIf(objProps.Name Is Nothing, "", objProps.Name.ToString), index, 0)
                            xlsArray.SetValue(objProp.Name, index, 1)
                            xlsArray.SetValue(Convert.ToString(objProp.Value), index, 2)
                        Next
                    Next




                    WriteSpreadsheetFromArray(xlsArray, sfdSelectXLSFile.FileName)

                    objDocuments.Close()
                    objApp.Quit()

                End If


            End If
        Catch exception As Exception
            DisplayException(exception)
        Finally
            If Not objProp Is Nothing Then
                Marshal.ReleaseComObject(objProp)
                objProp = Nothing
            End If
            If Not objProps Is Nothing Then
                Marshal.ReleaseComObject(objProps)
                objProps = Nothing
            End If
            If Not objPropSets Is Nothing Then
                objPropSets.Close()
                Marshal.ReleaseComObject(objPropSets)
                objPropSets = Nothing
            End If
        End Try
    End Sub

    Private Sub btnGenerateBOMSupplier_Click(sender As System.Object, e As System.EventArgs) Handles btnGenerateBOMSupplier.Click
        BOM_Generate(False)
    End Sub

    Private Sub BOM_Generate(PropBom As Boolean)
        Dim objApp As SolidEdgeFramework.Application = Nothing
        Dim objDocuments As SolidEdgeFramework.Documents = Nothing
        Dim objAssembly As SolidEdgeAssembly.AssemblyDocument = Nothing
        Dim bomService As New BomService(AddressOf PsmGetProperty)
        Dim bomAssembly As BOMAssembly
        Dim xlsArray As Array

        Try
            If ofdSelectASMFile.ShowDialog() = Windows.Forms.DialogResult.OK Then
                sfdSelectXLSFile.FileName = "Lista_" + Path.GetFileNameWithoutExtension(ofdSelectASMFile.FileName)
                If sfdSelectXLSFile.ShowDialog() = Windows.Forms.DialogResult.OK Then

                    ' Register with OLE to handle concurrency issues on the current thread.
                    SolidEdgeCommunity.OleMessageFilter.Register()
                    ' Connect to or start Solid Edge.
                    objApp = SolidEdgeCommunity.SolidEdgeUtils.Connect(True, True)
                    ' Make Solid Edge visible
                    objApp.Visible = True 'se_off.Checked --> sembra non salvi in background
                    ' Turn off alerts. Weldment environment will display a warning
                    objApp.DisplayAlerts = False
                    ' Get a reference to the Documents collection
                    objDocuments = objApp.Documents
                    ' Create an instance of each document environment
                    objAssembly = objDocuments.Open(ofdSelectASMFile.FileName)

                    bomAssembly = bomService.Build(objAssembly.FullName, objAssembly.Occurrences)

                    If PropBom Then
                        xlsArray = bomService.ToPropertyArray(bomAssembly, AddressOf CheckMaterial)
                    Else
                        xlsArray = bomService.ToSupplierArray(bomAssembly, Prefisso.Text, AddressOf CheckMaterial)
                    End If

                    WriteSpreadsheetFromArray(xlsArray, sfdSelectXLSFile.FileName)

                    objDocuments.Close()
                    objApp.Quit()

                End If
            End If

        Catch exception As Exception
            DisplayException(exception)
        End Try

    End Sub
    Private Function PsmGetProperty(path As String,
                                    propertySetName As String,
                                    propertyName As String)
        Return FilePropertyService.GetPropertyValue(path, propertySetName, propertyName)
    End Function

    'Private Sub BOM_ToI24Array_AddArray(values As Array, ByRef index As Integer, ByVal level As String, item As BOMItem)
    '    values.SetValue(IIf(String.IsNullOrEmpty(level), "0", level), index, 0)
    '    values.SetValue(item.Description, index, 1)
    '    values.SetValue(item.Material, index, 2)
    '    values.SetValue(Path.GetFileName(item.Name), index, 3)
    '    values.SetValue(item.Count.ToString(), index, 4)
    '    index += 1
    '    If TypeOf item Is BOMAssembly Then
    '        Dim elementIndex As Integer = 1

    '        For Each subItem As BOMItem In item.Items
    '            BOM_ToI24Array_AddArray(values,
    '                index,
    '                IIf(String.IsNullOrEmpty(level),
    '                    String.Format("{0}", elementIndex),
    '                    String.Format("{0}.{1}", level, elementIndex)), subItem)
    '            elementIndex += 1
    '        Next
    '    End If
    'End Sub

    'Private Function BOM_ToI24Array(bomAssembly As BOMAssembly)

    '    Dim count As Integer = 0
    '    Dim values(0, 3) As String
    '    Dim index As Integer = 0

    '    ' Calcola il numero totale di elementi
    '    For Each assemblyKeyValuePair As KeyValuePair(Of String, BOMAssembly) In m_BOMAssemblies
    '        count += 1 + assemblyKeyValuePair.Value.Items.Count
    '    Next

    '    ReDim values(count, 3)
    '    BOM_ToI24Array_AddArray(values, index, "", bomAssembly)

    '    Return values

    'End Function

    Private Sub BOMPrint(item As BOMItem, level As Integer)

        If TypeOf item Is BOMAssembly Then

            Debug.WriteLine(String.Format("{0} [ASM] {1} (={2})", New String(" ", level), item.Name, item.Count))
            For Each subItem As BOMItem In item.Items
                BOMPrint(subItem, level + 1)
            Next
        ElseIf TypeOf item Is BOMItem Then
            Debug.WriteLine(String.Format("{0} [ITM] {1} (={2})", New String(" ", level), item.Name, item.Count))
        End If

    End Sub

    Public Function CheckMaterial(item_material As String) As Boolean
        Dim selectedMaterials As New List(Of String)

        For Each item In Material.CheckedItems
            selectedMaterials.Add(item.ToString())
        Next

        Return MaterialFilter.MatchesSelectedMaterial(item_material, selectedMaterials)

    End Function
    Sub GetFileProps(filename As String, i As Integer)
        'Dim objPropSets As SolidEdgeFramework.PropertySets
        'Dim objProps As SolidEdgeFramework.Properties
        'Dim objProp As SolidEdgeFramework.Property


        'objPropSets = CreateObject("SolidEdge.FileProperties")
        'Call objPropSets.Open(filename)

        'objProps = objPropSets.Item("ProjectInformation")
        'objProp = objProps.Item("Document Number")
        'Data(i, 5) = objProp.Value
        'objProp = objProps.Item("Revision")
        'Data(i, 4) = objProp.Value

        ''For Each objProps In objPropSets
        ''    For Each objProp In objProps
        ''     Debug.Print objProps.Name, ": ", objProp.Name, " = ", objProp.Value
        ''    Next
        ''Next

        ''objProps = objPropSets.Item("ProjectInformation")
        ''For Each objProp In objProps
        ''    Debug.Print(objProp.Name, " = ", objProp.Value)
        ''Next

        ''objProps = objPropSets.Item("SummaryInformation")
        ''For Each objProp In objProps
        ''    Debug.Print(objProp.Name, " = ", objProp.Value)
        ''Next


        ''objProps = objPropSets.Item("MechanicalModeling")
        ''For Each objProp In objProps
        ''    Debug.Print(objProp.Name, " = ", objProp.Value)
        ''Next

        ''objProps = objPropSets.Item("Custom")
        ''For Each objProp In objProps
        ''    Debug.Print(objProp.Name, " = ", objProp.Value)
        ''Next

        'End
    End Sub

    Public Sub WriteSpreadsheetFromArray(strOutputArray As Array, Optional ByVal strExcelFileOutPath As String = "")
        'To avoid conflicts with different versions of Excel...We are using late binding.
        Dim objxlOutApp As Object = Nothing 'Excel.Application
        Dim objxlOutWBook As Object = Nothing 'Excel.Workbook
        Dim objxlOutSheet As Object = Nothing 'Excel.Worksheet
        Dim objxlRange As Object = Nothing 'Excel.Range
        Try
            'Try to Open Excel, Add a workbook and worksheet
            objxlOutApp = CreateObject("Excel.Application") 'New Excel.Application
            objxlOutWBook = objxlOutApp.Workbooks.Add '.Add.Sheets
            objxlOutSheet = objxlOutWBook.Sheets.Item(1)
        Catch ex As Exception
            MessageBox.Show("While trying to Open Excel recieved error:" & ex.Message, "Export to Excel Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Try
                If Not IsNothing(objxlOutWBook) Then
                    objxlOutWBook.Close()  'If an error occured we want to close the workbook
                End If
                If Not IsNothing(objxlOutApp) Then
                    objxlOutApp.Quit() 'If an error occured we want to close Excel
                End If
            Catch
            End Try
            objxlOutSheet = Nothing
            objxlOutWBook = Nothing
            If Not IsNothing(objxlOutApp) Then
                System.Runtime.InteropServices.Marshal.ReleaseComObject(objxlOutApp)  'This will release the object reference
            End If
            objxlOutApp = Nothing
            Exit Sub 'An error occured so we don't want to continue
        End Try
        Try
            objxlOutApp.DisplayAlerts = False    'This will prevent any message prompts from Excel (IE.."Do you want to save before closing?")
            objxlOutApp.Visible = False    'We don't want the app visible while we are populating it.
            'This is the easiest way I have found to populate a spreadsheet
            'First we get the range based on the size of our array

            objxlRange = objxlOutSheet.Range(Chr(strOutputArray.GetLowerBound(1) + 1 + 64) & (strOutputArray.GetLowerBound(0) + 1) & ":" & Chr(strOutputArray.GetUpperBound(1) + 1 + 64) & (strOutputArray.GetUpperBound(0) + 1))
            'Next we set the value of that range to our array
            objxlRange.Value = strOutputArray
            'This final part is optional, but we Auto Fit the columns of the spreadsheet.
            objxlRange.Columns.AutoFit()
            If strExcelFileOutPath.Length > 0 Then 'If a file name is passed
                Dim objFileInfo As New IO.FileInfo(strExcelFileOutPath)
                If Not objFileInfo.Directory.Exists Then 'Check if folder exists
                    objFileInfo.Directory.Create() 'If not we create it
                End If
                objFileInfo = Nothing
                objxlOutWBook.SaveAs(strExcelFileOutPath)  'Then we save our file.
            End If
            objxlOutApp.Visible = True 'Make excel visible
        Catch ex As Exception
            MessageBox.Show("While trying to Export to Excel recieved error:" & ex.Message, "Export to Excel Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Try
                objxlOutWBook.Close()  'If an error occured we want to close the workbook
                objxlOutApp.Quit() 'If an error occured we want to close Excel
            Catch
            End Try
        Finally
            objxlOutSheet = Nothing
            objxlOutWBook = Nothing
            If Not IsNothing(objxlOutApp) Then
                System.Runtime.InteropServices.Marshal.ReleaseComObject(objxlOutApp) 'This will release the object reference
            End If
            objxlOutApp = Nothing
        End Try
    End Sub

#End Region

#Region "====[ Solid Edge Functions ]===="

    Private Function SE_OpenApplication(makeVisible As Boolean) As SolidEdgeFramework.Application
        Return SolidEdgeSessionHelpers.OpenApplication(makeVisible)

    End Function

    Private Sub SE_CloseApplication(ByRef seApplication As SolidEdgeFramework.Application, quit As Boolean)
        SolidEdgeSessionHelpers.CloseApplication(seApplication, quit)
    End Sub

    Private Sub ReleaseCOMReference(ByRef comObject As Object)
        SolidEdgeSessionHelpers.ReleaseCOMReference(comObject)
    End Sub

    Private Sub ExportPartDocument(ByVal seApplication As SolidEdgeFramework.Application,
                        inPARFilePath As String,
                        outFilePath As String)

        Dim seDocuments As SolidEdgeFramework.Documents = Nothing
        Dim sePARDocument As SolidEdgePart.PartDocument = Nothing

        Try
            seDocuments = seApplication.Documents

            ' Apre il par file
            sePARDocument = seDocuments.Open(inPARFilePath)

            If Not Directory.Exists(Path.GetDirectoryName(outFilePath)) Then
                Directory.CreateDirectory(Path.GetDirectoryName(outFilePath))
            End If


            If File.Exists(outFilePath) Then
                File.Delete(outFilePath)
            End If

            ' Export
            ' MessageBox.Show(Me, "Sto salvando: " + outFilePath, "Informazione", MessageBoxButtons.OK, MessageBoxIcon.Information)

            sePARDocument.SaveCopyAs(outFilePath)
            sePARDocument.Close()

        Finally
            ReleaseCOMReference(seDocuments)
            ReleaseCOMReference(sePARDocument)
        End Try

    End Sub

    Private Sub ExportSheetMetalDocumentDocument(ByVal seApplication As SolidEdgeFramework.Application,
                        inPSMFilePath As String,
                        outFilePath As String)

        Dim seDocuments As SolidEdgeFramework.Documents = Nothing
        Dim sePSMDocument As SolidEdgePart.SheetMetalDocument = Nothing

        Try
            seDocuments = seApplication.Documents

            ' Apre il par file
            sePSMDocument = seDocuments.Open(inPSMFilePath)

            If Not Directory.Exists(Path.GetDirectoryName(outFilePath)) Then
                Directory.CreateDirectory(Path.GetDirectoryName(outFilePath))
            End If

            If File.Exists(outFilePath) Then
                File.Delete(outFilePath)
            End If

            ' Export

            sePSMDocument.SaveCopyAs(outFilePath)
            sePSMDocument.Close()

        Finally
            ReleaseCOMReference(seDocuments)
            ReleaseCOMReference(sePSMDocument)
        End Try

    End Sub

    Private Sub ExportPartDocumentImage(ByVal seApplication As SolidEdgeFramework.Application,
                        inPARFilePath As String,
                        outFilePath As String)

        Dim seDocuments As SolidEdgeFramework.Documents = Nothing
        Dim sePARDocument As SolidEdgePart.PartDocument = Nothing
        Dim seRefPlanes As SolidEdgePart.RefPlanes = Nothing
        Dim seRefSketchs As SolidEdgePart.Sketchs = Nothing
        Dim seView As SolidEdgeFramework.View = Nothing
        Dim seViewStyle As SolidEdgeFramework.ViewStyle = Nothing
        Dim seWindow As SolidEdgeFramework.Window = Nothing
        Dim seSketch As SolidEdgePart.Sketch = Nothing
        Dim seNamedViews As SolidEdgeFramework.NamedViews = Nothing
        Dim index As Integer = 0
        Dim view As Object = Nothing


        Try
            seDocuments = seApplication.Documents

            ' Apre il par file
            sePARDocument = seDocuments.Open(inPARFilePath)

            If Not Directory.Exists(Path.GetDirectoryName(outFilePath)) Then
                Directory.CreateDirectory(Path.GetDirectoryName(outFilePath))
            End If


            If File.Exists(outFilePath) Then
                File.Delete(outFilePath)
            End If



            seWindow = TryCast(seApplication.ActiveWindow, SolidEdgeFramework.Window)

            seRefPlanes = sePARDocument.RefPlanes
            seRefSketchs = sePARDocument.Sketches



            For Each plane In seRefPlanes
                plane.Visible = False
            Next

            For Each sketch In seRefSketchs
                sketch.ShowSketchColors = False
            Next

            seView = seWindow.View

            seView.SaveAsImage(outFilePath)

            sePARDocument.Close()

        Finally
            ReleaseCOMReference(seDocuments)
            ReleaseCOMReference(sePARDocument)
        End Try

    End Sub


    Private Sub ExportSheetMetalDocumentToDxf(ByVal seApplication As SolidEdgeFramework.Application,
                        inPSMFilePath As String,
                        outFilePath As String)

        Dim seDocuments As SolidEdgeFramework.Documents = Nothing
        Dim sePSMDocument As SolidEdgePart.SheetMetalDocument = Nothing
        Dim seBody As SolidEdgeGeometry.Body = Nothing
        Dim seFace As SolidEdgeGeometry.Face = Nothing
        Dim seBiggestFace As SolidEdgeGeometry.Face = Nothing
        Dim seFirstEdge As SolidEdgeGeometry.Edge = Nothing
        Dim seStartVertex As SolidEdgeGeometry.Vertex = Nothing
        Dim maxOpp As Double = 0

        Dim t As Integer

        Try
            seDocuments = seApplication.Documents

            ' Apre il par file
            sePSMDocument = seDocuments.Open(inPSMFilePath)



            seBody = sePSMDocument.Models.Item(1).Body

            For t = 1 To seBody.Faces(FaceType:=SolidEdgeConstants.FeatureTopologyQueryTypeConstants.igQueryAll).Count

                Dim dblParam(3) As Double
                Dim dblMaxTang(0 To 3) As Double
                Dim dblMaxCurv(0 To 3) As Double
                Dim dblMinCurv(0 To 3) As Double

                dblParam(0) = 3.141592
                dblParam(1) = 0.05
                dblParam(2) = 3 / 2 * 3.141592
                dblParam(3) = 0.1


                seFace = seBody.Faces(SolidEdgeConstants.FeatureTopologyQueryTypeConstants.igQueryAll).Item(t)

                seFace.GetCurvatures(2, dblParam, dblMaxTang, dblMaxCurv, dblMinCurv)

                If seFace.Area > maxOpp AndAlso dblMaxCurv(0) = 0 Then

                    maxOpp = seFace.Area
                    seBiggestFace = seFace
                End If
            Next

            For Each edge In seBiggestFace.Edges

                Dim dblParams(0 To 0) As Double
                Dim dblDirections(0 To 0) As Double
                Dim dblCurvatures(0 To 0) As Double

                dblParams(0) = 0
                edge.GetCurvature(1, dblParams,
                    dblDirections,
                    dblCurvatures)

                If cir_on.Checked = False Then

                    If dblCurvatures(0) = 0 AndAlso dblDirections(0) + dblDirections(1) + dblDirections(2) = 0 Then
                        seFirstEdge = edge
                        Exit For
                    End If
                Else
                    seFirstEdge = edge
                End If


            Next

            If Not seFirstEdge Is Nothing Then
                seStartVertex = seFirstEdge.StartVertex
            End If

            If Not seStartVertex Is Nothing Then

                If Not Directory.Exists(Path.GetDirectoryName(outFilePath)) Then
                    Directory.CreateDirectory(Path.GetDirectoryName(outFilePath))
                End If

                sePSMDocument.Models.SaveAsFlatDXF(outFilePath, seBiggestFace, seFirstEdge, seStartVertex)

            ElseIf seStartVertex Is Nothing AndAlso cir_on.Checked = True Then

                If Not Directory.Exists(Path.GetDirectoryName(outFilePath)) Then
                    Directory.CreateDirectory(Path.GetDirectoryName(outFilePath))
                End If

                sePSMDocument.Models.SaveAsFlatDXF(outFilePath, seBiggestFace, seFirstEdge, seFirstEdge)

            End If

            sePSMDocument.Close()

        Finally
            ReleaseCOMReference(seDocuments)
            ReleaseCOMReference(sePSMDocument)
        End Try

    End Sub




#End Region

#Region "====[ Generate 'Disegni di Piega' ]===="

    Public Function GenerateDisegniDiPiega_Execute(asmFilePath As String) As Boolean

        Dim seApplication As SolidEdgeFramework.Application = Nothing
        Dim seDocuments As SolidEdgeFramework.Documents = Nothing
        Dim seAssembly As SolidEdgeAssembly.AssemblyDocument = Nothing
        Dim draftService As New DraftGenerationService(AddressOf CheckMaterial,
                                                       AddressOf DisegniDiPiega_ExportDFT,
                                                       AddressOf DisplayException)

        Try
            seApplication = SE_OpenApplication(se_off.CheckState)

            seApplication.DisplayAlerts = False
            seDocuments = seApplication.Documents

            ' Load file asm
            seAssembly = seDocuments.Open(asmFilePath)

            If draftService.GenerateForAssembly(seApplication, seAssembly, Prefisso.Text) = False Then
                Return False
            End If

        Finally
            ReleaseCOMReference(seDocuments)
            ReleaseCOMReference(seAssembly)
            SE_CloseApplication(seApplication, True)
        End Try

        Return True

    End Function

    Public Sub DisegniDiPiega_ExportDFT(ByVal seApplication As SolidEdgeFramework.Application,
                        outputDFTFilePath As String,
                        modelLinkPath As String)

        Dim objDocuments As SolidEdgeFramework.Documents = Nothing
        Dim objDraft As SolidEdgeDraft.DraftDocument = Nothing
        Dim objSheet As SolidEdgeDraft.Sheet = Nothing
        Dim objModelLinks As SolidEdgeDraft.ModelLinks = Nothing
        Dim objModelLink As SolidEdgeDraft.ModelLink = Nothing
        Dim objDrawingViews As SolidEdgeDraft.DrawingViews = Nothing
        Dim objDrawingView As SolidEdgeDraft.DrawingView = Nothing

        Try
            objDocuments = seApplication.Documents

            ' Add a Draft document
            objDraft = objDocuments.Add("SolidEdge.DraftDocument")

            ' Get a reference to the active sheet
            objSheet = objDraft.ActiveSheet

            ' Get a reference to the model links collection
            objModelLinks = objDraft.ModelLinks

            ' Add a new model link
            objModelLink = objModelLinks.Add(modelLinkPath)

            ' Get a reference to the drawing views collection
            objDrawingViews = objSheet.DrawingViews

            If Path.GetExtension(modelLinkPath) = ".psm" Then

                ' Add a FRONT view
                objDrawingView = objDrawingViews.AddSheetMetalView(
                objModelLink,
                SolidEdgeDraft.ViewOrientationConstants.igFrontView,
                CDbl(txtScale.Text),
                0.1,
                0.3,
                SolidEdgeDraft.SheetMetalDrawingViewTypeConstants.seSheetMetalDesignedView)

                objDrawingViews.AddByFold(objDrawingView,
                    SolidEdgeDraft.FoldTypeConstants.igFoldRight,
                    0.3, 0.3)
                objDrawingViews.AddByFold(objDrawingView,
                    SolidEdgeDraft.FoldTypeConstants.igFoldDown,
                    0.1, 0.1)
                objDrawingViews.AddByFold(objDrawingView,
                SolidEdgeDraft.FoldTypeConstants.igFoldDownRight,
                0.3, 0.1)
            End If


            If Path.GetExtension(modelLinkPath) = ".par" Then

                ' Add a FRONT view
                objDrawingView = objDrawingViews.AddPartView(
                objModelLink,
                SolidEdgeDraft.ViewOrientationConstants.igBottomFrontRightView,
                CDbl(txtScale.Text),
                0.12,
                0.3,
                SolidEdgeDraft.PartDrawingViewTypeConstants.sePartDesignedView)

            End If

            ' Assign a caption
            'objDrawingView.Caption = "Da decidere"
            ' Ensure caption is displayed
            'objDrawingView.DisplayCaption = True

            If Not Directory.Exists(Path.GetDirectoryName(outputDFTFilePath)) Then
                Directory.CreateDirectory(Path.GetDirectoryName(outputDFTFilePath))
            End If

            If File.Exists(outputDFTFilePath) Then
                File.Delete(outputDFTFilePath)
            End If

            objDraft.SaveAs(outputDFTFilePath)
            objDraft.Close()

        Finally
            ReleaseCOMReference(objDocuments)
            ReleaseCOMReference(objDraft)
            ReleaseCOMReference(objSheet)
            ReleaseCOMReference(objModelLinks)
            ReleaseCOMReference(objModelLink)
            ReleaseCOMReference(objDrawingViews)
            ReleaseCOMReference(objDrawingView)
        End Try
    End Sub

    'Public Sub DisegniDiPiega_RelinkDFT(inputDFTDirectory As String)

    '    Dim RMApp As RevisionManager.Application = Nothing
    '    Dim objDraft As RevisionManager.Document = Nothing



    '    Try

    '        For Each dftPath As String In Directory.GetFiles(inputDFTDirectory, "*.dft")

    '            ' Load file dft
    '            objDraft = RMApp.OpenFileInRevisionManager(dftPath)

    '            obj

    '            outPDFFilePath = Path.Combine(Path.GetDirectoryName(dftPath), "DWG",
    '                Path.GetFileNameWithoutExtension(dftPath) + ".dwg")

    '            If Not Directory.Exists(Path.GetDirectoryName(outPDFFilePath)) Then
    '                Directory.CreateDirectory(Path.GetDirectoryName(outPDFFilePath))
    '            End If

    '            If File.Exists(outPDFFilePath) Then
    '                File.Delete(outPDFFilePath)
    '            End If

    '            objDraft.SaveAs(outPDFFilePath)

    '            objDraft.Close()

    '            ReleaseCOMReference(objDraft)

    '        Next


    '    Finally

    '        ReleaseCOMReference(objDraft)
    '        ReleaseCOMReference(RMApp)

    '    End Try
    'End Sub









    Public Sub DisegniDiPiega_ExportJPG(ByVal seApplication As SolidEdgeFramework.Application,
                        outputDFTFilePath As String,
                        modelLinkPath As String)

        Dim objDocuments As SolidEdgeFramework.Documents = Nothing
        Dim objDraft As SolidEdgeDraft.DraftDocument = Nothing
        Dim objSheet As SolidEdgeDraft.Sheet = Nothing
        Dim objModelLinks As SolidEdgeDraft.ModelLinks = Nothing
        Dim objModelLink As SolidEdgeDraft.ModelLink = Nothing
        Dim objDrawingViews As SolidEdgeDraft.DrawingViews = Nothing
        Dim objDrawingView As SolidEdgeDraft.DrawingView = Nothing
        Dim seView As SolidEdgeFramework.View = Nothing
        Dim seViewStyle As SolidEdgeFramework.ViewStyle = Nothing
        Dim seWindow As SolidEdgeFramework.Window = Nothing

        Try
            objDocuments = seApplication.Documents

            ' Add a Draft document
            objDraft = objDocuments.Add("SolidEdge.DraftDocument")

            ' Get a reference to the active sheet
            objSheet = objDraft.ActiveSheet

            ' Get a reference to the model links collection
            objModelLinks = objDraft.ModelLinks

            ' Add a new model link
            objModelLink = objModelLinks.Add(modelLinkPath)

            ' Get a reference to the drawing views collection
            objDrawingViews = objSheet.DrawingViews

            ' Add a FRONT view
            objDrawingView = objDrawingViews.AddSheetMetalView(
                objModelLink,
                SolidEdgeDraft.ViewOrientationConstants.igFrontView,
                CDbl(txtScale.Text),
                0.12,
                0.3,
                SolidEdgeDraft.SheetMetalDrawingViewTypeConstants.seSheetMetalDesignedView)

            objDrawingViews.AddByFold(objDrawingView,
                SolidEdgeDraft.FoldTypeConstants.igFoldRight,
                0.3, 0.3)
            objDrawingViews.AddByFold(objDrawingView,
                SolidEdgeDraft.FoldTypeConstants.igFoldDown,
                0.1, 0.1)
            objDrawingViews.AddByFold(objDrawingView,
                SolidEdgeDraft.FoldTypeConstants.igFoldDownRight,
                0.3, 0.1)

            ' Assign a caption
            'objDrawingView.Caption = "Da decidere"
            ' Ensure caption is displayed
            'objDrawingView.DisplayCaption = True

            If Not Directory.Exists(Path.GetDirectoryName(outputDFTFilePath)) Then
                Directory.CreateDirectory(Path.GetDirectoryName(outputDFTFilePath))
            End If

            If File.Exists(outputDFTFilePath) Then
                File.Delete(outputDFTFilePath)
            End If

            Dim image As Imaging.Metafile


            image = objSheet.GetEnhancedMetafile()




            Dim Width As Object = 1920
            Dim Height As Object = 1080
            Dim AltViewStyle As Object = "Default"
            Dim Resolution As Object = 1
            Dim ColorDepth As Object = 24
            Dim ImageQuality = SolidEdgeFramework.SeImageQualityType.seImageQualityHigh
            Dim Invert As Boolean = False


            seWindow = TryCast(seApplication.ActiveWindow, SolidEdgeFramework.Window)

            If seWindow Is Nothing Then
                MessageBox.Show(Me, outputDFTFilePath, "Informazione", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Else
                MessageBox.Show(Me, "TryCast OK", "Informazione", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If

            seView = seWindow.View
            seView.SaveAsImage(outputDFTFilePath, Width, Height, AltViewStyle, Resolution, ColorDepth, ImageQuality, Invert)


            objDraft.Close()

        Finally
            ReleaseCOMReference(objDocuments)
            ReleaseCOMReference(objDraft)
            ReleaseCOMReference(objSheet)
            ReleaseCOMReference(objModelLinks)
            ReleaseCOMReference(objModelLink)
            ReleaseCOMReference(objDrawingViews)
            ReleaseCOMReference(objDrawingView)
        End Try
    End Sub

    Private Sub btnGenerateDisegniDiPiega_Click(sender As System.Object, e As System.EventArgs) Handles btnGenerateDisegniDiPiega.Click
        Try
            If ofdSelectASMFile.ShowDialog() = Windows.Forms.DialogResult.OK Then
                GenerateDisegniDiPiega_Execute(ofdSelectASMFile.FileName)
                MessageBox.Show(Me, "Generazione Disegni di Piega completata.", "Informazione", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If

        Catch exception As Exception
            DisplayException(exception)
        End Try

    End Sub

#End Region

#Region "====[ Export to STL/STP (PAR/PSM) ]===="

    Public Function Export_Execute(asmFilePath As String, type As String) As Boolean

        Dim seApplication As SolidEdgeFramework.Application = Nothing
        Dim seDocuments As SolidEdgeFramework.Documents = Nothing
        Dim seAssembly As SolidEdgeAssembly.AssemblyDocument = Nothing
        Dim exportService As New NeutralExportService(AddressOf CheckMaterial,
                                                      AddressOf ExportPartDocument,
                                                      AddressOf ExportSheetMetalDocumentDocument,
                                                      AddressOf DisplayException)

        Try
            seApplication = SE_OpenApplication(se_off.CheckState)

            seApplication.DisplayAlerts = False
            seDocuments = seApplication.Documents

            ' Load asm file
            seAssembly = seDocuments.Open(asmFilePath)

            If exportService.ExportAssembly(seApplication, seAssembly, Prefisso.Text, type) = False Then
                Return False
            End If

        Finally
            SE_CloseApplication(seApplication, True)
            ReleaseCOMReference(seDocuments)
            ReleaseCOMReference(seAssembly)
        End Try

        Return True
    End Function

    Private Sub btnExportSTL_Click(sender As System.Object, e As System.EventArgs) Handles btnExportSTL.Click
        Try
            If ofdSelectASMFile.ShowDialog() = Windows.Forms.DialogResult.OK Then
                Export_Execute(ofdSelectASMFile.FileName, "stl")
                MessageBox.Show(Me, "Esportazione in STL (PAR/PSM) completata.", "Informazione", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If

        Catch exception As Exception
            DisplayException(exception)
        End Try

    End Sub

    Private Sub btnExportSTP_Click(sender As Object, e As EventArgs) Handles btnExportSTP.Click
        Try
            If ofdSelectASMFile.ShowDialog() = Windows.Forms.DialogResult.OK Then
                Export_Execute(ofdSelectASMFile.FileName, "stp")
                MessageBox.Show(Me, "Esportazione in STP (PAR/PSM) completata.", "Informazione", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If

        Catch exception As Exception
            DisplayException(exception)
        End Try
    End Sub


#End Region

#Region "====[ Export to DXF (PSM) ]===="

    Public Function ExportDXF_Execute(asmFilePath As String) As Boolean

        Dim seApplication As SolidEdgeFramework.Application = Nothing
        Dim seDocuments As SolidEdgeFramework.Documents = Nothing
        Dim seAssembly As SolidEdgeAssembly.AssemblyDocument = Nothing
        Dim exportService As New FlatDxfExportService(AddressOf CheckMaterial,
                                                      AddressOf ExportSheetMetalDocumentToDxf,
                                                      AddressOf DisplayException)

        Try
            seApplication = SE_OpenApplication(se_off.CheckState)

            seApplication.DisplayAlerts = False
            seDocuments = seApplication.Documents

            ' Load asm file
            seAssembly = seDocuments.Open(asmFilePath)

            If Not exportService.ExportAssembly(seApplication, seAssembly, Prefisso.Text, all_subasm.Checked) Then
                Return False
            End If

        Finally
            SE_CloseApplication(seApplication, True)
            ReleaseCOMReference(seDocuments)
            ReleaseCOMReference(seAssembly)
        End Try

        Return True

    End Function

    Private Sub btnExportDXF_Click(sender As System.Object, e As System.EventArgs) Handles btnExportDXF.Click

        Try
            If ofdSelectASMFile.ShowDialog() = Windows.Forms.DialogResult.OK Then
                ExportDXF_Execute(ofdSelectASMFile.FileName)
                MessageBox.Show(Me, "Esportazione in DXF (PSM) completata.", "Informazione", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If

        Catch exception As Exception
            DisplayException(exception)
        End Try

    End Sub

#End Region

#Region "====[ Convert 'Disegni di Piega' to PDF ]===="

    Private Sub btnConvertDisegniDiPiegaToPdf_Click(sender As System.Object, e As System.EventArgs) Handles btnConvertDisegniDiPiegaToPdf.Click
        Try
            If fbdSelectDisegniDiPiegaFolder.ShowDialog() = Windows.Forms.DialogResult.OK Then
                ConvertDisegniDiPiegaToPdf_Execute(fbdSelectDisegniDiPiegaFolder.SelectedPath)
                MessageBox.Show(Me, "Conversione PDF completata.", "Informazione", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If

        Catch exception As Exception
            DisplayException(exception)
        End Try
    End Sub

    Public Function ConvertDisegniDiPiegaToPdf_Execute(inputDFTDirectory As String) As Boolean

        Dim seApplication As SolidEdgeFramework.Application = Nothing
        Dim publishService As New DraftPublishService()

        Try
            seApplication = SE_OpenApplication(se_off.CheckState)

            seApplication.DisplayAlerts = False
            publishService.PublishPdf(seApplication, inputDFTDirectory)

        Finally
            SE_CloseApplication(seApplication, True)
        End Try

        Return True

    End Function

    Public Function ConvertDisegniDiPiegaToDWG_Execute(inputDFTDirectory As String) As Boolean

        Dim seApplication As SolidEdgeFramework.Application = Nothing
        Dim publishService As New DraftPublishService()

        Try
            seApplication = SE_OpenApplication(se_off.CheckState)

            seApplication.DisplayAlerts = False
            publishService.PublishDwg(seApplication, inputDFTDirectory)

        Finally
            SE_CloseApplication(seApplication, True)
        End Try

        Return True

    End Function


#End Region

    Private Function DisplayException(exception As Exception,
        Optional ByVal text As String = "",
        Optional buttons As MessageBoxButtons = MessageBoxButtons.OK,
        Optional icon As MessageBoxIcon = MessageBoxIcon.Error) As DialogResult

        Return MessageBox.Show(Me, text & vbNewLine & exception.Message, "Errore", buttons, icon)

    End Function

    Private Function inPARFilePath() As String
        Throw New NotImplementedException
    End Function

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles SoloMateriale.CheckedChanged
        If SoloMateriale.Checked Then
            SubFolders.Enabled = True
            Material.Enabled = True
        Else
            SubFolders.Enabled = False
            Material.Enabled = False
        End If
    End Sub

    Private Sub SET_MainForm_Load(sender As Object, e As EventArgs) Handles Me.Load
        ' Attiva tutti i materiali in partenza
        For i As Int16 = 0 To Material.Items.Count - 1
            Material.SetItemChecked(i, True)
        Next
        If SoloMateriale.Checked Then
            Material.Enabled = True
        Else
            Material.Enabled = False
        End If
    End Sub

    Private Sub bntPropBOM_Click(sender As Object, e As EventArgs) Handles bntPropBOM.Click
        BOM_Generate(True)
    End Sub


#Region "====[ Crea file JPG  ]===="

    Private Sub btnExportJPG_Click(sender As Object, e As EventArgs) Handles btnExportJPG.Click


        Try
            If ofdSelectASMFile.ShowDialog() = Windows.Forms.DialogResult.OK Then
                ExportJPG_Execute(ofdSelectASMFile.FileName)
                MessageBox.Show(Me, "Esportazione in JPG (PAR) completata.", "Informazione", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If

        Catch exception As Exception
            DisplayException(exception)
        End Try

    End Sub

    Public Function ExportJPG_Execute(asmFilePath As String) As Boolean

        Dim seApplication As SolidEdgeFramework.Application = Nothing
        Dim seDocuments As SolidEdgeFramework.Documents = Nothing
        Dim seAssembly As SolidEdgeAssembly.AssemblyDocument = Nothing
        Dim exportService As New ImageExportService(AddressOf CheckMaterial,
                                                    AddressOf ExportPartDocumentImage,
                                                    AddressOf DisplayException)

        Try
            seApplication = SE_OpenApplication(se_off.CheckState)

            seApplication.DisplayAlerts = False
            seDocuments = seApplication.Documents

            ' Load asm file
            seAssembly = seDocuments.Open(asmFilePath)

            If Not exportService.ExportAssembly(seApplication, seAssembly, Prefisso.Text, all_subasm.Checked) Then
                Return False
            End If

        Finally
            SE_CloseApplication(seApplication, True)
            ReleaseCOMReference(seDocuments)
            ReleaseCOMReference(seAssembly)
        End Try

        Return True
    End Function

    Private Sub btnCodificaProgetto_Click(sender As Object, e As EventArgs) Handles btnCodificaProgetto.Click
        Dim objApp As SolidEdgeFramework.Application = Nothing
        Dim objDocuments As SolidEdgeFramework.Documents = Nothing
        Dim xlsArray(100, 3) As String
        Dim index As Integer = 0
        Dim objPropSets As SolidEdgeFileProperties.PropertySets = New SolidEdgeFileProperties.PropertySets
        Dim objProp As SolidEdgeFileProperties.Property = Nothing
        Dim objProps As SolidEdgeFileProperties.Properties = Nothing


        Try
            If ofdSelectPSMFile.ShowDialog() = Windows.Forms.DialogResult.OK Then


                ' Register with OLE to handle concurrency issues on the current thread.
                SolidEdgeCommunity.OleMessageFilter.Register()
                ' Connect to or start Solid Edge.
                objApp = SolidEdgeCommunity.SolidEdgeUtils.Connect(True, True)
                ' Make Solid Edge visible
                objApp.Visible = True 'se_off.Checked --> sembra non salvi in background
                ' Turn off alerts. Weldment environment will display a warning
                objApp.DisplayAlerts = True
                ' Get a reference to the Documents collection
                objDocuments = objApp.Documents
                ' Create an instance of each document environment
                Dim sDocument As String = ofdSelectPSMFile.FileName


                objPropSets.Open(sDocument, False)


                objProps = objPropSets.Item("ProjectInformation")

                objProps.Item("Project Name").Value = txtProgetto.Text
                objProps.Item("Revision").Value = txtVersione.Text
                objProps.Item("Document Number").Value = txtProgressivo.Text

                objProps.Save()
                objPropSets.Save()
                objPropSets.Close()
                objApp.Quit()

            End If

        Catch exception As Exception
            DisplayException(exception)
        Finally
            If Not objProp Is Nothing Then
                Marshal.ReleaseComObject(objProp)
                objProp = Nothing
            End If
            If Not objProps Is Nothing Then
                Marshal.ReleaseComObject(objProps)
                objProps = Nothing
            End If
            If Not objPropSets Is Nothing Then
                objPropSets.Close()
                Marshal.ReleaseComObject(objPropSets)
                objPropSets = Nothing
            End If
        End Try
    End Sub

#End Region

#Region "====[ Genera Viste 3D, propedeutico per STL/STP list]===="

    Private Sub btnGenerateDisegni3D_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub btnConvertDisegniDiPiegaToDWG_Click(sender As Object, e As EventArgs) Handles btnConvertDisegniDiPiegaToDWG.Click
        Try
            If fbdSelectDisegniDiPiegaFolder.ShowDialog() = Windows.Forms.DialogResult.OK Then
                ConvertDisegniDiPiegaToDWG_Execute(fbdSelectDisegniDiPiegaFolder.SelectedPath)
                MessageBox.Show(Me, "Conversione DWG completata.", "Informazione", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If

        Catch exception As Exception
            DisplayException(exception)
        End Try
    End Sub



#End Region


End Class
