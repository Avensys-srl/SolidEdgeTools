Imports SolidEdgeCommunity.Extensions
Imports System.Runtime.InteropServices
Imports System.IO
Imports System.Linq
Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Drawing.Imaging

Public Class SET_MainForm

    Private ReadOnly _configurationEngine As New ConfigurationEngine()
    Private ReadOnly _workflowService As New SolidEdgeWorkflowService()
    Private _currentOperationName As String = ""
    Private _lastAssemblyPath As String = ""
    Private _lastDraftFolderPath As String = ""
    Private _cancelRequested As Boolean = False


#Region "====[ Generate BOM ]===="

    Private Sub btnPropTable_Click(sender As System.Object, e As System.EventArgs) Handles btnPropTable.Click

        Dim xlsArray(100, 3) As String
        Dim index As Integer = 0
        Dim objPropSets As SolidEdgeFileProperties.PropertySets = New SolidEdgeFileProperties.PropertySets
        Dim objProp As SolidEdgeFileProperties.Property = Nothing
        Dim objProps As SolidEdgeFileProperties.Properties = Nothing

        Try
            If ofdSelectPSMFile.ShowDialog() = Windows.Forms.DialogResult.OK Then
                sfdSelectXLSFile.FileName = Prefisso.Text + "Proprietà_" + Path.GetFileNameWithoutExtension(ofdSelectPSMFile.FileName)
                If sfdSelectXLSFile.ShowDialog() = Windows.Forms.DialogResult.OK Then

                    Dim sDocument As String = ofdSelectPSMFile.FileName

                    objPropSets.Open(sDocument, False)

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

                End If


            End If
        Catch exception As Exception
            DisplayException(exception)
        Finally
            ReleaseCOMReference(objProp)
            ReleaseCOMReference(objProps)
            If Not objPropSets Is Nothing Then
                objPropSets.Close()
                ReleaseCOMReference(objPropSets)
            End If
        End Try
    End Sub

    Private Sub btnGenerateBOMSupplier_Click(sender As System.Object, e As System.EventArgs) Handles btnGenerateBOMSupplier.Click
        BOM_Generate(False)
    End Sub

    Private Sub BOM_Generate(PropBom As Boolean)
        Dim bomService As New BomService(AddressOf PsmGetProperty)
        Dim bomOptions = GetBomExportOptions()

        Try
            If ofdSelectASMFile.ShowDialog() = Windows.Forms.DialogResult.OK Then
                RememberAssemblyPath(ofdSelectASMFile.FileName)
                If PropBom Then
                    sfdSelectXLSFile.FileName = "Lista_Proprietà_" + Path.GetFileNameWithoutExtension(ofdSelectASMFile.FileName)
                Else
                    sfdSelectXLSFile.FileName = "Lista_" + Path.GetFileNameWithoutExtension(ofdSelectASMFile.FileName)
                End If
                If sfdSelectXLSFile.ShowDialog() = Windows.Forms.DialogResult.OK Then
                    _workflowService.ExecuteWithAssembly(
                        ofdSelectASMFile.FileName,
                        GetApplicationOptions(),
                        False,
                        Function(app, assembly)
                            Dim bomAssembly = bomService.Build(assembly.FullName, assembly.Occurrences)
                            Dim xlsArray As Array

                            If PropBom Then
                                xlsArray = bomService.ToPropertyArray(bomAssembly, bomOptions)
                            Else
                                xlsArray = bomService.ToSupplierArray(bomAssembly, bomOptions)
                            End If

                            WriteSpreadsheetFromArray(xlsArray, sfdSelectXLSFile.FileName)
                            Return True
                        End Function)
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
        Return MaterialFilter.MatchesSelectedMaterial(item_material, GetMaterialSelectionOptions().SelectedMaterials)

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

    Public Sub WriteSpreadsheetFromArray(strOutputArray As Array,
                                         Optional ByVal strExcelFileOutPath As String = "",
                                         Optional ByVal showExcel As Boolean = True)
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
            objxlOutApp.Visible = showExcel 'Make excel visible only when requested
        Catch ex As Exception
            MessageBox.Show("While trying to Export to Excel recieved error:" & ex.Message, "Export to Excel Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Try
                objxlOutWBook.Close()  'If an error occured we want to close the workbook
                objxlOutApp.Quit() 'If an error occured we want to close Excel
            Catch
            End Try
        Finally
            ReleaseCOMReference(objxlRange)
            ReleaseCOMReference(objxlOutSheet)
            ReleaseCOMReference(objxlOutWBook)
            If Not showExcel AndAlso Not IsNothing(objxlOutApp) Then
                Try
                    objxlOutApp.Quit()
                Catch
                End Try
            End If
            If Not IsNothing(objxlOutApp) Then
                System.Runtime.InteropServices.Marshal.ReleaseComObject(objxlOutApp) 'This will release the object reference
            End If
            objxlOutApp = Nothing
        End Try
    End Sub

#End Region

#Region "====[ Solid Edge Functions ]===="

    Private Function SE_OpenApplication(options As SolidEdgeApplicationOptions) As SolidEdgeSessionContext
        Return SolidEdgeSessionHelpers.OpenApplication(options.MakeVisible)

    End Function

    Private Sub SE_CloseApplication(ByRef session As SolidEdgeSessionContext, quit As Boolean)
        SolidEdgeSessionHelpers.CloseApplication(session, quit)
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

    Private Sub ExportModelDocumentImage(ByVal seApplication As SolidEdgeFramework.Application,
                        inputFilePath As String,
                        outFilePath As String)

        Dim seDocuments As SolidEdgeFramework.Documents = Nothing
        Dim seDocument As SolidEdgeFramework.SolidEdgeDocument = Nothing
        Dim seRefPlanes As Object = Nothing
        Dim seRefSketchs As Object = Nothing
        Dim seView As SolidEdgeFramework.View = Nothing
        Dim seWindow As SolidEdgeFramework.Window = Nothing
        Dim viewStyle As Object = Nothing
        Dim previousRenderMode As SolidEdgeFramework.SeRenderModeType
        Dim previousSilhouettesEnabled As Boolean = False
        Dim previousBackgroundType As Object = Nothing
        Dim previousBackgroundImageDisplayed As Object = Nothing
        Dim previousReflections As Object = Nothing
        Dim previousFloorReflection As Object = Nothing
        Dim previousDropShadow As Object = Nothing
        Dim previousCastShadows As Object = Nothing
        Dim previousTextures As Object = Nothing
        Dim previousStyleSilhouettesEnabled As Object = Nothing


        Try
            seDocuments = seApplication.Documents

            seDocument = DirectCast(seDocuments.Open(inputFilePath), SolidEdgeFramework.SolidEdgeDocument)
            seDocument.Activate()

            If Not Directory.Exists(Path.GetDirectoryName(outFilePath)) Then
                Directory.CreateDirectory(Path.GetDirectoryName(outFilePath))
            End If


            If File.Exists(outFilePath) Then
                File.Delete(outFilePath)
            End If

            seWindow = TryCast(seApplication.ActiveWindow, SolidEdgeFramework.Window)
            If seWindow Is Nothing Then
                Throw New InvalidOperationException("Finestra Solid Edge non disponibile per export JPG.")
            End If

            seView = seWindow.View
            viewStyle = seView.ViewStyle
            previousRenderMode = seView.RenderModeType
            previousSilhouettesEnabled = seView.SilhouettesEnabled

            If viewStyle IsNot Nothing Then
                Try
                    previousBackgroundType = viewStyle.BackgroundType
                Catch
                End Try

                Try
                    previousBackgroundImageDisplayed = viewStyle.IsBackgroundImageDisplayed
                Catch
                End Try

                Try
                    previousReflections = viewStyle.Reflections
                Catch
                End Try

                Try
                    previousFloorReflection = viewStyle.FloorReflection
                Catch
                End Try

                Try
                    previousDropShadow = viewStyle.DropShadow
                Catch
                End Try

                Try
                    previousCastShadows = viewStyle.CastShadows
                Catch
                End Try

                Try
                    previousTextures = viewStyle.Textures
                Catch
                End Try

                Try
                    previousStyleSilhouettesEnabled = viewStyle.SilhouettesEnabled
                Catch
                End Try
            End If

            Try
                seRefPlanes = CallByName(seDocument, "RefPlanes", CallType.Get)
            Catch
            End Try

            Try
                seRefSketchs = CallByName(seDocument, "Sketches", CallType.Get)
            Catch
            End Try

            If seRefPlanes IsNot Nothing Then
                For Each plane In seRefPlanes
                    plane.Visible = False
                Next
            End If

            If seRefSketchs IsNot Nothing Then
                For Each sketch In seRefSketchs
                    sketch.ShowSketchColors = False
                Next
            End If

            ConfigureImageView(seView, viewStyle)
            FitModelInView(seView)
            SaveViewImageWithFallback(seView, outFilePath)

            seDocument.Close()

        Finally
            RestoreImageView(seView,
                             viewStyle,
                             previousRenderMode,
                             previousSilhouettesEnabled,
                             previousBackgroundType,
                             previousBackgroundImageDisplayed,
                             previousReflections,
                             previousFloorReflection,
                             previousDropShadow,
                             previousCastShadows,
                             previousTextures,
                             previousStyleSilhouettesEnabled)
            ReleaseCOMReference(seView)
            ReleaseCOMReference(seWindow)
            ReleaseCOMReference(viewStyle)
            ReleaseCOMReference(seRefPlanes)
            ReleaseCOMReference(seRefSketchs)
            ReleaseCOMReference(seDocuments)
            ReleaseCOMReference(seDocument)
        End Try

    End Sub

    Private Sub ConfigureImageView(seView As SolidEdgeFramework.View,
                                   viewStyle As Object)

        If seView Is Nothing Then
            Return
        End If

        seView.SetRenderMode(SolidEdgeFramework.SeRenderModeType.seRenderModeOutline)
        seView.RenderModeType = SolidEdgeFramework.SeRenderModeType.seRenderModeOutline
        seView.SilhouettesEnabled = True

        If viewStyle IsNot Nothing Then
            Try
                viewStyle.BeginPropertyBuffer()
            Catch
            End Try

            Try
                viewStyle.RenderModeType = SolidEdgeFramework.SeRenderModeType.seRenderModeOutline
            Catch
            End Try

            Try
                viewStyle.BackgroundType = SolidEdgeFramework.SeBackgroundType.seBackgroundTypeGradient
            Catch
            End Try

            Try
                viewStyle.IsBackgroundImageDisplayed = 0
            Catch
            End Try

            Try
                viewStyle.Reflections = 0
            Catch
            End Try

            Try
                viewStyle.FloorReflection = 0
            Catch
            End Try

            Try
                viewStyle.DropShadow = 0
            Catch
            End Try

            Try
                viewStyle.CastShadows = 0
            Catch
            End Try

            Try
                viewStyle.Textures = 0
            Catch
            End Try

            Try
                viewStyle.SilhouettesEnabled = True
            Catch
            End Try

            Try
                viewStyle.Perspective = 0
            Catch
            End Try

            Try
                viewStyle.SetGradientBackground(
                    SolidEdgeFramework.SeGradientType.seGradientTypeVertical,
                    System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White),
                    System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White),
                    System.Reflection.Missing.Value,
                    System.Reflection.Missing.Value)
            Catch
            End Try

            Try
                viewStyle.FlushPropertyBuffer()
            Catch
            End Try
        End If
    End Sub

    Private Sub RestoreImageView(seView As SolidEdgeFramework.View,
                                 viewStyle As Object,
                                 previousRenderMode As SolidEdgeFramework.SeRenderModeType,
                                 previousSilhouettesEnabled As Boolean,
                                 previousBackgroundType As Object,
                                 previousBackgroundImageDisplayed As Object,
                                 previousReflections As Object,
                                 previousFloorReflection As Object,
                                 previousDropShadow As Object,
                                 previousCastShadows As Object,
                                 previousTextures As Object,
                                 previousStyleSilhouettesEnabled As Object)

        If seView Is Nothing Then
            Return
        End If

        Try
            seView.RenderModeType = previousRenderMode
        Catch
        End Try

        Try
            seView.SilhouettesEnabled = previousSilhouettesEnabled
        Catch
        End Try

        If viewStyle IsNot Nothing Then
            Try
                If previousBackgroundType IsNot Nothing Then viewStyle.BackgroundType = previousBackgroundType
            Catch
            End Try

            Try
                If previousBackgroundImageDisplayed IsNot Nothing Then viewStyle.IsBackgroundImageDisplayed = previousBackgroundImageDisplayed
            Catch
            End Try

            Try
                If previousReflections IsNot Nothing Then viewStyle.Reflections = previousReflections
            Catch
            End Try

            Try
                If previousFloorReflection IsNot Nothing Then viewStyle.FloorReflection = previousFloorReflection
            Catch
            End Try

            Try
                If previousDropShadow IsNot Nothing Then viewStyle.DropShadow = previousDropShadow
            Catch
            End Try

            Try
                If previousCastShadows IsNot Nothing Then viewStyle.CastShadows = previousCastShadows
            Catch
            End Try

            Try
                If previousTextures IsNot Nothing Then viewStyle.Textures = previousTextures
            Catch
            End Try

            Try
                If previousStyleSilhouettesEnabled IsNot Nothing Then viewStyle.SilhouettesEnabled = previousStyleSilhouettesEnabled
            Catch
            End Try
        End If
    End Sub

    Private Sub FitModelInView(seView As SolidEdgeFramework.View)
        If seView Is Nothing Then
            Return
        End If

        Try
            seView.Fit()
        Catch
        End Try

        Try
            seView.Update()
        Catch
        End Try

        Try
            seView.ZoomCamera(0.9)
        Catch
        End Try
    End Sub

    Private Sub SaveViewImageWithFallback(seView As SolidEdgeFramework.View,
                                          outputJpgFilePath As String)

        Dim widths() As Integer = {2048, 1600, 1280}
        Dim heights() As Integer = {2048, 1600, 1280}
        Dim altViewStyle As Object = System.Reflection.Missing.Value
        Dim resolutions() As Integer = {1, 1, 1, 1}
        Dim colorDepth As Object = 24
        Dim imageQuality = SolidEdgeFramework.SeImageQualityType.seImageQualityHigh
        Dim invert As Boolean = False
        Dim lastException As Exception = Nothing
        Dim sourceImagePath = Path.Combine(Path.GetDirectoryName(outputJpgFilePath),
                                           Path.GetFileNameWithoutExtension(outputJpgFilePath) & "_source_tmp.jpg")

        For i As Integer = 0 To widths.Length - 1
            Try
                seView.SaveAsImage(sourceImagePath,
                                   widths(i),
                                   heights(i),
                                   altViewStyle,
                                   resolutions(i),
                                   colorDepth,
                                   imageQuality,
                                   invert)
                NormalizePreviewImage(sourceImagePath, outputJpgFilePath)
                Return
            Catch ex As COMException
                lastException = ex

                If ex.HResult <> &H8007000E Then
                    Throw
                End If
            Finally
                If File.Exists(sourceImagePath) Then
                    File.Delete(sourceImagePath)
                End If
            End Try
        Next

        If lastException IsNot Nothing Then
            Throw lastException
        End If
    End Sub

    Private Sub NormalizePreviewImage(sourceImagePath As String, outputImagePath As String)
        Const edgeSampleWidth As Integer = 6
        Const contentThreshold As Integer = 24
        Const outputCanvasSize As Integer = 1200
        Const contentFillRatio As Double = 0.92
        Const analysisSize As Integer = 512
        Dim tempImagePath = Path.Combine(Path.GetDirectoryName(outputImagePath),
                                         Path.GetFileNameWithoutExtension(outputImagePath) & "_preview_tmp.jpg")

        Using sourceBitmap As New Bitmap(sourceImagePath)
            Dim analysisWidth = Math.Min(analysisSize, sourceBitmap.Width)
            Dim analysisHeight = Math.Min(analysisSize, sourceBitmap.Height)
            Dim cropRectangle As Rectangle

            Using analysisBitmap As New Bitmap(analysisWidth, analysisHeight, PixelFormat.Format24bppRgb)
                Using graphicsContext As System.Drawing.Graphics = System.Drawing.Graphics.FromImage(analysisBitmap)
                    graphicsContext.Clear(Color.White)
                    graphicsContext.InterpolationMode = InterpolationMode.HighQualityBicubic
                    graphicsContext.DrawImage(sourceBitmap,
                                              New Rectangle(0, 0, analysisWidth, analysisHeight),
                                              New Rectangle(0, 0, sourceBitmap.Width, sourceBitmap.Height),
                                              GraphicsUnit.Pixel)
                End Using

                Dim analysisBounds = FindContentBounds(analysisBitmap, contentThreshold, edgeSampleWidth)

                If analysisBounds.Width <= 0 OrElse analysisBounds.Height <= 0 Then
                    sourceBitmap.Save(tempImagePath, ImageFormat.Jpeg)
                    ReplaceNormalizedPreviewImage(outputImagePath, tempImagePath)
                    Return
                End If

                Dim scaleX = sourceBitmap.Width / CDbl(analysisWidth)
                Dim scaleY = sourceBitmap.Height / CDbl(analysisHeight)
                cropRectangle = New Rectangle(
                    Math.Max(0, CInt(Math.Floor(analysisBounds.Left * scaleX))),
                    Math.Max(0, CInt(Math.Floor(analysisBounds.Top * scaleY))),
                    Math.Min(sourceBitmap.Width, CInt(Math.Ceiling(analysisBounds.Width * scaleX))),
                    Math.Min(sourceBitmap.Height, CInt(Math.Ceiling(analysisBounds.Height * scaleY))))
            End Using

            Using croppedBitmap As New Bitmap(cropRectangle.Width, cropRectangle.Height, PixelFormat.Format24bppRgb)
                Using cropGraphics As System.Drawing.Graphics = System.Drawing.Graphics.FromImage(croppedBitmap)
                    cropGraphics.Clear(Color.White)
                    cropGraphics.InterpolationMode = InterpolationMode.HighQualityBicubic
                    cropGraphics.DrawImage(sourceBitmap,
                                           New Rectangle(0, 0, cropRectangle.Width, cropRectangle.Height),
                                           cropRectangle,
                                           GraphicsUnit.Pixel)
                End Using

                NormalizePreviewBackground(croppedBitmap, contentThreshold, edgeSampleWidth)

                Dim normalizedBounds = FindContentBounds(croppedBitmap, contentThreshold, edgeSampleWidth)
                If normalizedBounds.Width <= 0 OrElse normalizedBounds.Height <= 0 Then
                    normalizedBounds = New Rectangle(0, 0, croppedBitmap.Width, croppedBitmap.Height)
                End If

                Using previewBitmap As New Bitmap(outputCanvasSize, outputCanvasSize, PixelFormat.Format24bppRgb)
                    Using graphicsContext As System.Drawing.Graphics = System.Drawing.Graphics.FromImage(previewBitmap)
                        graphicsContext.Clear(Color.White)
                        graphicsContext.InterpolationMode = InterpolationMode.HighQualityBicubic
                        graphicsContext.SmoothingMode = SmoothingMode.AntiAlias
                        graphicsContext.PixelOffsetMode = PixelOffsetMode.HighQuality
                        graphicsContext.CompositingQuality = CompositingQuality.HighQuality

                        Dim scale = Math.Min((outputCanvasSize * contentFillRatio) / normalizedBounds.Width,
                                             (outputCanvasSize * contentFillRatio) / normalizedBounds.Height)
                        Dim targetWidth = Math.Max(1, CInt(normalizedBounds.Width * scale))
                        Dim targetHeight = Math.Max(1, CInt(normalizedBounds.Height * scale))
                        Dim targetX = (outputCanvasSize - targetWidth) \ 2
                        Dim targetY = (outputCanvasSize - targetHeight) \ 2

                        graphicsContext.DrawImage(croppedBitmap,
                                                  New Rectangle(targetX, targetY, targetWidth, targetHeight),
                                                  normalizedBounds,
                                                  GraphicsUnit.Pixel)
                    End Using

                    previewBitmap.Save(tempImagePath, ImageFormat.Jpeg)
                End Using
            End Using
        End Using

        ReplaceNormalizedPreviewImage(outputImagePath, tempImagePath)
    End Sub

    Private Sub ReplaceNormalizedPreviewImage(imagePath As String, tempImagePath As String)
        If File.Exists(imagePath) Then
            File.Delete(imagePath)
        End If

        File.Move(tempImagePath, imagePath)
    End Sub

    Private Function EstimateRowBackground(sourceBitmap As Bitmap,
                                           row As Integer,
                                           sampleWidth As Integer) As Color

        Dim effectiveSampleWidth = Math.Min(sampleWidth, Math.Max(1, sourceBitmap.Width \ 10))
        Dim totalR As Integer = 0
        Dim totalG As Integer = 0
        Dim totalB As Integer = 0
        Dim sampleCount As Integer = 0

        For x As Integer = 0 To effectiveSampleWidth - 1
            Dim leftPixel = sourceBitmap.GetPixel(x, row)
            totalR += leftPixel.R
            totalG += leftPixel.G
            totalB += leftPixel.B
            sampleCount += 1

            Dim rightPixel = sourceBitmap.GetPixel(sourceBitmap.Width - 1 - x, row)
            totalR += rightPixel.R
            totalG += rightPixel.G
            totalB += rightPixel.B
            sampleCount += 1
        Next

        Return Color.FromArgb(totalR \ sampleCount, totalG \ sampleCount, totalB \ sampleCount)
    End Function

    Private Function FindContentBounds(sourceBitmap As Bitmap,
                                       contentThreshold As Integer,
                                       sampleWidth As Integer) As Rectangle

        Dim minX As Integer = sourceBitmap.Width - 1
        Dim minY As Integer = sourceBitmap.Height - 1
        Dim maxX As Integer = 0
        Dim maxY As Integer = 0
        Dim hasContent As Boolean = False

        For y As Integer = 0 To sourceBitmap.Height - 1
            Dim backgroundColor = EstimateRowBackground(sourceBitmap, y, sampleWidth)

            For x As Integer = 0 To sourceBitmap.Width - 1
                Dim pixel = sourceBitmap.GetPixel(x, y)
                Dim distance = Math.Abs(CInt(pixel.R) - CInt(backgroundColor.R)) +
                               Math.Abs(CInt(pixel.G) - CInt(backgroundColor.G)) +
                               Math.Abs(CInt(pixel.B) - CInt(backgroundColor.B))

                If distance > contentThreshold Then
                    hasContent = True
                    If x < minX Then minX = x
                    If y < minY Then minY = y
                    If x > maxX Then maxX = x
                    If y > maxY Then maxY = y
                End If
            Next
        Next

        If Not hasContent Then
            Return Rectangle.Empty
        End If

        Dim marginX = Math.Max(4, CInt((maxX - minX + 1) * 0.12))
        Dim marginY = Math.Max(4, CInt((maxY - minY + 1) * 0.12))
        Dim cropLeft = Math.Max(0, minX - marginX)
        Dim cropTop = Math.Max(0, minY - marginY)
        Dim cropRight = Math.Min(sourceBitmap.Width - 1, maxX + marginX)
        Dim cropBottom = Math.Min(sourceBitmap.Height - 1, maxY + marginY)

        Return Rectangle.FromLTRB(cropLeft,
                                  cropTop,
                                  cropRight + 1,
                                  cropBottom + 1)
    End Function

    Private Sub NormalizePreviewBackground(previewBitmap As Bitmap,
                                           contentThreshold As Integer,
                                           sampleWidth As Integer)

        For y As Integer = 0 To previewBitmap.Height - 1
            Dim backgroundColor = EstimateRowBackground(previewBitmap, y, sampleWidth)

            For x As Integer = 0 To previewBitmap.Width - 1
                Dim pixel = previewBitmap.GetPixel(x, y)
                Dim distance = Math.Abs(CInt(pixel.R) - CInt(backgroundColor.R)) +
                               Math.Abs(CInt(pixel.G) - CInt(backgroundColor.G)) +
                               Math.Abs(CInt(pixel.B) - CInt(backgroundColor.B))

                If distance <= contentThreshold Then
                    previewBitmap.SetPixel(x, y, Color.White)
                End If
            Next
        Next
    End Sub

    Private Sub SaveSheetWindowImageWithFallback(sheetWindow As Object,
                                                 outputJpgFilePath As String)

        Dim widths() As Integer = {3840, 2560, 1920, 1600, 1280}
        Dim heights() As Integer = {2160, 1440, 1080, 900, 720}
        Dim resolutions() As Integer = {1, 1, 1, 1, 1}
        Dim colorDepth As Object = 24
        Dim imageQuality = SolidEdgeFramework.SeImageQualityType.seImageQualityHigh
        Dim invert As Boolean = False
        Dim lastException As Exception = Nothing

        For i As Integer = 0 To widths.Length - 1
            Try
                sheetWindow.SaveAsImage(outputJpgFilePath,
                                        widths(i),
                                        heights(i),
                                        resolutions(i),
                                        colorDepth,
                                        imageQuality,
                                        invert)
                Return
            Catch ex As COMException
                lastException = ex

                If ex.HResult <> &H8007000E Then
                    Throw
                End If
            End Try
        Next

        If lastException IsNot Nothing Then
            Throw lastException
        End If
    End Sub


    Private Sub ExportSheetMetalDocumentToDxf(ByVal seApplication As SolidEdgeFramework.Application,
                        inPSMFilePath As String,
                        outFilePath As String)

        Dim seDocuments As SolidEdgeFramework.Documents = Nothing
        Dim sePSMDocument As SolidEdgePart.SheetMetalDocument = Nothing
        Dim seModels As Object = Nothing
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



            seModels = sePSMDocument.Models
            seBody = seModels.Item(1).Body

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

                seModels.SaveAsFlatDXF(outFilePath, seBiggestFace, seFirstEdge, seStartVertex)

            ElseIf seStartVertex Is Nothing AndAlso cir_on.Checked = True Then

                If Not Directory.Exists(Path.GetDirectoryName(outFilePath)) Then
                    Directory.CreateDirectory(Path.GetDirectoryName(outFilePath))
                End If

                seModels.SaveAsFlatDXF(outFilePath, seBiggestFace, seFirstEdge, seFirstEdge)

            End If

            sePSMDocument.Close()

        Finally
            ReleaseCOMReference(seStartVertex)
            ReleaseCOMReference(seFirstEdge)
            ReleaseCOMReference(seBiggestFace)
            ReleaseCOMReference(seFace)
            ReleaseCOMReference(seBody)
            ReleaseCOMReference(seModels)
            ReleaseCOMReference(seDocuments)
            ReleaseCOMReference(sePSMDocument)
        End Try

    End Sub




#End Region

#Region "====[ Generate 'Disegni di Piega' ]===="

    Public Function GenerateDisegniDiPiega_Execute(asmFilePath As String) As Boolean
        Dim draftOptions = GetDraftGenerationOptions()
        Dim draftService As New DraftGenerationService(Sub(app, outputPath, modelLinkPath) DisegniDiPiega_ExportDFT(app, outputPath, modelLinkPath, draftOptions.Scale, draftOptions.AutoLayoutSheetMetalViews),
                                                       AddressOf DisplayException)

        Return ExecuteWithProgress(
            "Generazione DFT",
            Function(progress)
                Return _workflowService.ExecuteWithAssembly(
                    asmFilePath,
                    GetApplicationOptions(),
                    False,
                    Function(app, assembly) draftService.GenerateForAssembly(app, assembly, draftOptions, progress, AddressOf IsCancellationRequested))
            End Function)

    End Function

    Public Sub DisegniDiPiega_ExportDFT(ByVal seApplication As SolidEdgeFramework.Application,
                        outputDFTFilePath As String,
                        modelLinkPath As String,
                        scale As Double,
                        Optional autoLayoutSheetMetalViews As Boolean = False)

        Dim objDocuments As SolidEdgeFramework.Documents = Nothing
        Dim objDraft As SolidEdgeDraft.DraftDocument = Nothing
        Dim objSheet As SolidEdgeDraft.Sheet = Nothing
        Dim objSheetSetup As SolidEdgeDraft.SheetSetup = Nothing
        Dim objModelLinks As SolidEdgeDraft.ModelLinks = Nothing
        Dim objModelLink As SolidEdgeDraft.ModelLink = Nothing
        Dim objDrawingViews As SolidEdgeDraft.DrawingViews = Nothing
        Dim objDrawingView As SolidEdgeDraft.DrawingView = Nothing
        Dim objFoldedView As SolidEdgeDraft.DrawingView = Nothing
        Dim objTopView As SolidEdgeDraft.DrawingView = Nothing
        Dim objRightView As SolidEdgeDraft.DrawingView = Nothing
        Dim objIsoView As SolidEdgeDraft.DrawingView = Nothing

        Try
            If Path.GetExtension(modelLinkPath).Equals(".psm", StringComparison.OrdinalIgnoreCase) Then
                If TryReuseArchivedSheetMetalDraft(seApplication, outputDFTFilePath, modelLinkPath) Then
                    Return
                End If
            End If

            objDocuments = seApplication.Documents
            CreateDraftDocumentContext(objDocuments,
                                       modelLinkPath,
                                       objDraft,
                                       objSheet,
                                       objSheetSetup,
                                       objModelLinks,
                                       objModelLink,
                                       objDrawingViews)

            If Path.GetExtension(modelLinkPath) = ".psm" Then
                If autoLayoutSheetMetalViews Then
                    autoLayoutSheetMetalViews = TryCreateAutoLayoutSheetMetalViews(objDocuments,
                                                                                   modelLinkPath,
                                                                                   objSheetSetup,
                                                                                   objModelLink,
                                                                                   objDrawingViews,
                                                                                   scale,
                                                                                   objDrawingView,
                                                                                   objTopView,
                                                                                   objRightView,
                                                                                   objIsoView)

                    If Not autoLayoutSheetMetalViews Then
                        ResetDraftDocumentContext(objDocuments,
                                                  modelLinkPath,
                                                  objDraft,
                                                  objSheet,
                                                  objSheetSetup,
                                                  objModelLinks,
                                                  objModelLink,
                                                  objDrawingViews,
                                                  objDrawingView,
                                                  objTopView,
                                                  objRightView,
                                                  objIsoView,
                                                  objFoldedView)
                    End If
                End If

                If Not autoLayoutSheetMetalViews Then
                    objDrawingView = objDrawingViews.AddSheetMetalView(
                    objModelLink,
                    SolidEdgeDraft.ViewOrientationConstants.igFrontView,
                    scale,
                    0.1,
                    0.3,
                    SolidEdgeDraft.SheetMetalDrawingViewTypeConstants.seSheetMetalDesignedView)

                    objFoldedView = objDrawingViews.AddByFold(objDrawingView,
                        SolidEdgeDraft.FoldTypeConstants.igFoldRight,
                        0.3, 0.3)
                    ReleaseCOMReference(objFoldedView)
                    objFoldedView = Nothing
                    objFoldedView = objDrawingViews.AddByFold(objDrawingView,
                        SolidEdgeDraft.FoldTypeConstants.igFoldDown,
                        0.1, 0.1)
                    ReleaseCOMReference(objFoldedView)
                    objFoldedView = Nothing
                    objFoldedView = objDrawingViews.AddByFold(objDrawingView,
                    SolidEdgeDraft.FoldTypeConstants.igFoldDownRight,
                    0.3, 0.1)
                End If
            End If


            If Path.GetExtension(modelLinkPath) = ".par" Then

                ' Add a FRONT view
                objDrawingView = objDrawingViews.AddPartView(
                objModelLink,
                SolidEdgeDraft.ViewOrientationConstants.igBottomFrontRightView,
                scale,
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
            ReleaseCOMReference(objIsoView)
            ReleaseCOMReference(objRightView)
            ReleaseCOMReference(objTopView)
            ReleaseCOMReference(objFoldedView)
            ReleaseCOMReference(objDocuments)
            ReleaseCOMReference(objDraft)
            ReleaseCOMReference(objSheet)
            ReleaseCOMReference(objSheetSetup)
            ReleaseCOMReference(objModelLinks)
            ReleaseCOMReference(objModelLink)
            ReleaseCOMReference(objDrawingViews)
            ReleaseCOMReference(objDrawingView)
        End Try
    End Sub

    Private Function TryReuseArchivedSheetMetalDraft(seApplication As SolidEdgeFramework.Application,
                                                     outputDFTFilePath As String,
                                                     modelLinkPath As String) As Boolean

        Dim archivedDraftPath As String = Nothing
        Dim objDocuments As SolidEdgeFramework.Documents = Nothing
        Dim objDraft As SolidEdgeDraft.DraftDocument = Nothing
        Dim objModelLinks As SolidEdgeDraft.ModelLinks = Nothing
        Dim objModelLink As SolidEdgeDraft.ModelLink = Nothing
        Dim reused As Boolean = False

        Try
            archivedDraftPath = FindMostRecentArchivedDraft(outputDFTFilePath)

            If String.IsNullOrWhiteSpace(archivedDraftPath) OrElse Not File.Exists(archivedDraftPath) Then
                Return False
            End If

            If Not Directory.Exists(Path.GetDirectoryName(outputDFTFilePath)) Then
                Directory.CreateDirectory(Path.GetDirectoryName(outputDFTFilePath))
            End If

            If File.Exists(outputDFTFilePath) Then
                File.Delete(outputDFTFilePath)
            End If

            File.Copy(archivedDraftPath, outputDFTFilePath, True)

            objDocuments = seApplication.Documents
            objDraft = CType(objDocuments.Open(outputDFTFilePath), SolidEdgeDraft.DraftDocument)
            objModelLinks = objDraft.ModelLinks

            If objModelLinks Is Nothing OrElse objModelLinks.Count = 0 Then
                Return False
            End If

            For index As Integer = 1 To objModelLinks.Count
                objModelLink = objModelLinks.Item(index)

                If ShouldReuseArchivedModelLink(objModelLink, modelLinkPath, objModelLinks.Count) Then
                    objModelLink.ChangeSource(modelLinkPath)

                    Try
                        objModelLink.ForceUpdateViews()
                    Catch
                        objModelLink.UpdateViews()
                    End Try

                    reused = True
                End If

                ReleaseCOMReference(objModelLink)
                objModelLink = Nothing
            Next

            If Not reused Then
                Return False
            End If

            objDraft.Save()
            objDraft.Close()
            Return True
        Catch
            Return False
        Finally
            Try
                If objDraft IsNot Nothing Then
                    objDraft.Close()
                End If
            Catch
            End Try

            If Not reused AndAlso File.Exists(outputDFTFilePath) Then
                Try
                    File.Delete(outputDFTFilePath)
                Catch
                End Try
            End If

            ReleaseCOMReference(objModelLink)
            ReleaseCOMReference(objModelLinks)
            ReleaseCOMReference(objDraft)
            ReleaseCOMReference(objDocuments)
        End Try
    End Function

    Private Function FindMostRecentArchivedDraft(outputDFTFilePath As String) As String
        Dim outputFolder = Path.GetDirectoryName(outputDFTFilePath)
        Dim folderParent = Path.GetDirectoryName(outputFolder)
        Dim folderName = Path.GetFileName(outputFolder)
        Dim draftFileName = Path.GetFileName(outputDFTFilePath)
        Dim archivedFolders As String()

        If String.IsNullOrWhiteSpace(outputFolder) OrElse String.IsNullOrWhiteSpace(folderParent) OrElse String.IsNullOrWhiteSpace(folderName) Then
            Return Nothing
        End If

        If Not Directory.Exists(folderParent) Then
            Return Nothing
        End If

        archivedFolders = Directory.GetDirectories(folderParent, folderName & "_old*")

        If archivedFolders Is Nothing OrElse archivedFolders.Length = 0 Then
            Return Nothing
        End If

        For Each archivedFolder In archivedFolders.OrderByDescending(Function(path) Directory.GetLastWriteTimeUtc(path))
            Dim archivedDraftPath = Path.Combine(archivedFolder, draftFileName)

            If File.Exists(archivedDraftPath) Then
                Return archivedDraftPath
            End If
        Next

        Return Nothing
    End Function

    Private Function ShouldReuseArchivedModelLink(modelLink As SolidEdgeDraft.ModelLink,
                                                  modelLinkPath As String,
                                                  modelLinkCount As Integer) As Boolean

        If modelLink Is Nothing Then
            Return False
        End If

        If modelLinkCount = 1 Then
            Return True
        End If

        Dim existingFileName = Path.GetFileNameWithoutExtension(modelLink.FileName)
        Dim targetFileName = Path.GetFileNameWithoutExtension(modelLinkPath)
        Dim existingExtension = Path.GetExtension(modelLink.FileName)
        Dim targetExtension = Path.GetExtension(modelLinkPath)

        Return existingFileName.Equals(targetFileName, StringComparison.OrdinalIgnoreCase) AndAlso
               existingExtension.Equals(targetExtension, StringComparison.OrdinalIgnoreCase)
    End Function

    Private Sub CreateDraftDocumentContext(objDocuments As SolidEdgeFramework.Documents,
                                           modelLinkPath As String,
                                           ByRef objDraft As SolidEdgeDraft.DraftDocument,
                                           ByRef objSheet As SolidEdgeDraft.Sheet,
                                           ByRef objSheetSetup As SolidEdgeDraft.SheetSetup,
                                           ByRef objModelLinks As SolidEdgeDraft.ModelLinks,
                                           ByRef objModelLink As SolidEdgeDraft.ModelLink,
                                           ByRef objDrawingViews As SolidEdgeDraft.DrawingViews)

        objDraft = objDocuments.Add("SolidEdge.DraftDocument")
        objSheet = objDraft.ActiveSheet
        objSheet.BackgroundVisible = True
        objSheetSetup = objSheet.SheetSetup
        Try
            objSheetSetup.ClearDrawingViewForSheetScale()
        Catch
        End Try
        objModelLinks = objDraft.ModelLinks
        objModelLink = objModelLinks.Add(modelLinkPath)
        objDrawingViews = objSheet.DrawingViews
    End Sub

    Private Sub ResetDraftDocumentContext(objDocuments As SolidEdgeFramework.Documents,
                                          modelLinkPath As String,
                                          ByRef objDraft As SolidEdgeDraft.DraftDocument,
                                          ByRef objSheet As SolidEdgeDraft.Sheet,
                                          ByRef objSheetSetup As SolidEdgeDraft.SheetSetup,
                                          ByRef objModelLinks As SolidEdgeDraft.ModelLinks,
                                          ByRef objModelLink As SolidEdgeDraft.ModelLink,
                                          ByRef objDrawingViews As SolidEdgeDraft.DrawingViews,
                                          ByRef objDrawingView As SolidEdgeDraft.DrawingView,
                                          ByRef objTopView As SolidEdgeDraft.DrawingView,
                                          ByRef objRightView As SolidEdgeDraft.DrawingView,
                                          ByRef objIsoView As SolidEdgeDraft.DrawingView,
                                          ByRef objFoldedView As SolidEdgeDraft.DrawingView)

        Try
            If objDraft IsNot Nothing Then
                objDraft.Close()
            End If
        Catch
        End Try

        ReleaseCOMReference(objIsoView)
        ReleaseCOMReference(objRightView)
        ReleaseCOMReference(objTopView)
        ReleaseCOMReference(objFoldedView)
        ReleaseCOMReference(objDrawingView)
        ReleaseCOMReference(objDrawingViews)
        ReleaseCOMReference(objModelLink)
        ReleaseCOMReference(objModelLinks)
        ReleaseCOMReference(objSheetSetup)
        ReleaseCOMReference(objSheet)
        ReleaseCOMReference(objDraft)

        objIsoView = Nothing
        objRightView = Nothing
        objTopView = Nothing
        objFoldedView = Nothing
        objDrawingView = Nothing
        objDrawingViews = Nothing
        objModelLink = Nothing
        objModelLinks = Nothing
        objSheetSetup = Nothing
        objSheet = Nothing
        objDraft = Nothing

        CreateDraftDocumentContext(objDocuments,
                                   modelLinkPath,
                                   objDraft,
                                   objSheet,
                                   objSheetSetup,
                                   objModelLinks,
                                   objModelLink,
                                   objDrawingViews)
    End Sub

    Private Function TryCreateAutoLayoutSheetMetalViews(objDocuments As SolidEdgeFramework.Documents,
                                                        modelLinkPath As String,
                                                        sheetSetup As SolidEdgeDraft.SheetSetup,
                                                        modelLink As SolidEdgeDraft.ModelLink,
                                                        drawingViews As SolidEdgeDraft.DrawingViews,
                                                        requestedScale As Double,
                                                        ByRef frontView As SolidEdgeDraft.DrawingView,
                                                        ByRef topView As SolidEdgeDraft.DrawingView,
                                                        ByRef rightView As SolidEdgeDraft.DrawingView,
                                                        ByRef isoView As SolidEdgeDraft.DrawingView) As Boolean

        Dim usableWidth As Double
        Dim usableHeight As Double
        Dim leftMargin As Double
        Dim bottomMargin As Double
        Dim sheetCenterX As Double
        Dim sheetCenterY As Double
        Dim viewPlan As SheetMetalAutoLayoutPlan

        Try
            If objDocuments Is Nothing OrElse sheetSetup Is Nothing OrElse modelLink Is Nothing OrElse drawingViews Is Nothing Then
                Return False
            End If

            viewPlan = MeasureSheetMetalAutoLayoutPlan(objDocuments, modelLinkPath, requestedScale)
            If viewPlan Is Nothing Then
                Return False
            End If

            usableWidth = CDbl(sheetSetup.SheetWidth) - CDbl(sheetSetup.LeftMargin) - CDbl(sheetSetup.RightMargin)
            usableHeight = CDbl(sheetSetup.SheetHeight) - CDbl(sheetSetup.TopMargin) - CDbl(sheetSetup.BottomMargin)
            leftMargin = CDbl(sheetSetup.LeftMargin)
            bottomMargin = CDbl(sheetSetup.BottomMargin)
            sheetCenterX = leftMargin + (usableWidth / 2)
            sheetCenterY = bottomMargin + (usableHeight / 2)

            frontView = drawingViews.AddSheetMetalView(
                modelLink,
                viewPlan.MainOrientation,
                viewPlan.FinalScale,
                sheetCenterX,
                sheetCenterY,
                SolidEdgeDraft.SheetMetalDrawingViewTypeConstants.seSheetMetalDesignedView)

            rightView = drawingViews.AddByFold(frontView,
                                               SolidEdgeDraft.FoldTypeConstants.igFoldRight,
                                               sheetCenterX,
                                               sheetCenterY)

            topView = drawingViews.AddByFold(frontView,
                                             SolidEdgeDraft.FoldTypeConstants.igFoldDown,
                                             sheetCenterX,
                                             sheetCenterY)

            isoView = drawingViews.AddByFold(frontView,
                                             SolidEdgeDraft.FoldTypeConstants.igFoldDownRight,
                                             sheetCenterX,
                                             sheetCenterY)

            ApplyAutoLayoutToSheetMetalViews(sheetSetup,
                                             sheetCenterX,
                                             sheetCenterY,
                                             frontView,
                                             topView,
                                             rightView,
                                             isoView)

            Return True
        Catch
            Return False
        End Try
    End Function

    Private Function MeasureSheetMetalAutoLayoutPlan(objDocuments As SolidEdgeFramework.Documents,
                                                     modelLinkPath As String,
                                                     requestedScale As Double) As SheetMetalAutoLayoutPlan

        Dim tempDraft As SolidEdgeDraft.DraftDocument = Nothing
        Dim tempSheet As SolidEdgeDraft.Sheet = Nothing
        Dim tempSheetSetup As SolidEdgeDraft.SheetSetup = Nothing
        Dim tempModelLinks As SolidEdgeDraft.ModelLinks = Nothing
        Dim tempModelLink As SolidEdgeDraft.ModelLink = Nothing
        Dim tempDrawingViews As SolidEdgeDraft.DrawingViews = Nothing
        Dim tempFrontView As SolidEdgeDraft.DrawingView = Nothing
        Dim tempTopView As SolidEdgeDraft.DrawingView = Nothing
        Dim tempRightView As SolidEdgeDraft.DrawingView = Nothing
        Dim tempIsoView As SolidEdgeDraft.DrawingView = Nothing
        Dim infoMain As ViewLayoutInfo = Nothing
        Dim infoTop As ViewLayoutInfo = Nothing
        Dim infoSide As ViewLayoutInfo = Nothing
        Dim infoIso As ViewLayoutInfo = Nothing
        Dim viewPlan As SheetMetalAutoLayoutPlan = Nothing
        Dim usableWidth As Double
        Dim usableHeight As Double
        Dim leftMargin As Double
        Dim bottomMargin As Double
        Dim centerX As Double
        Dim centerY As Double
        Dim seedScale As Double

        Try
            tempDraft = objDocuments.Add("SolidEdge.DraftDocument")
            tempSheet = tempDraft.ActiveSheet
            tempSheet.BackgroundVisible = True
            tempSheetSetup = tempSheet.SheetSetup
            tempModelLinks = tempDraft.ModelLinks
            tempModelLink = tempModelLinks.Add(modelLinkPath)
            tempDrawingViews = tempSheet.DrawingViews

            usableWidth = CDbl(tempSheetSetup.SheetWidth) - CDbl(tempSheetSetup.LeftMargin) - CDbl(tempSheetSetup.RightMargin)
            usableHeight = CDbl(tempSheetSetup.SheetHeight) - CDbl(tempSheetSetup.TopMargin) - CDbl(tempSheetSetup.BottomMargin)
            leftMargin = CDbl(tempSheetSetup.LeftMargin)
            bottomMargin = CDbl(tempSheetSetup.BottomMargin)
            centerX = leftMargin + (usableWidth / 2)
            centerY = bottomMargin + (usableHeight / 2)
            seedScale = Math.Max(0.05, Math.Min(0.25, requestedScale))

            viewPlan = BuildSheetMetalAutoLayoutPlan(tempModelLink)

            tempFrontView = tempDrawingViews.AddSheetMetalView(
                tempModelLink,
                viewPlan.MainOrientation,
                seedScale,
                centerX,
                centerY,
                SolidEdgeDraft.SheetMetalDrawingViewTypeConstants.seSheetMetalDesignedView)

            tempRightView = tempDrawingViews.AddByFold(tempFrontView,
                                                       SolidEdgeDraft.FoldTypeConstants.igFoldRight,
                                                       centerX,
                                                       centerY)

            tempTopView = tempDrawingViews.AddByFold(tempFrontView,
                                                     SolidEdgeDraft.FoldTypeConstants.igFoldDown,
                                                     centerX,
                                                     centerY)

            tempIsoView = tempDrawingViews.AddByFold(tempFrontView,
                                                     SolidEdgeDraft.FoldTypeConstants.igFoldDownRight,
                                                     centerX,
                                                     centerY)

            infoMain = CreateViewLayoutInfo(tempFrontView, centerX, centerY)
            infoTop = CreateViewLayoutInfo(tempTopView, centerX, centerY)
            infoSide = CreateViewLayoutInfo(tempRightView, centerX, centerY)
            infoIso = CreateViewLayoutInfo(tempIsoView, centerX, centerY)

            PrepareAutoLayoutView(infoMain.View)
            PrepareAutoLayoutView(infoTop.View)
            PrepareAutoLayoutView(infoSide.View)
            PrepareAutoLayoutView(infoIso.View)

            RefreshViewLayoutInfo(infoMain)
            RefreshViewLayoutInfo(infoTop)
            RefreshViewLayoutInfo(infoSide)
            RefreshViewLayoutInfo(infoIso)

            viewPlan.FinalScale = CalculateAutoLayoutScaleFactor(tempSheetSetup,
                                                                 infoMain,
                                                                 infoTop,
                                                                 infoSide,
                                                                 infoIso,
                                                                 seedScale)

            Return viewPlan
        Catch
            Return Nothing
        Finally
            Try
                If tempDraft IsNot Nothing Then
                    tempDraft.Close(False)
                End If
            Catch
            End Try

            ReleaseCOMReference(tempIsoView)
            ReleaseCOMReference(tempRightView)
            ReleaseCOMReference(tempTopView)
            ReleaseCOMReference(tempFrontView)
            ReleaseCOMReference(tempDrawingViews)
            ReleaseCOMReference(tempModelLink)
            ReleaseCOMReference(tempModelLinks)
            ReleaseCOMReference(tempSheetSetup)
            ReleaseCOMReference(tempSheet)
            ReleaseCOMReference(tempDraft)
        End Try
    End Function

    Private Sub ApplyAutoLayoutToSheetMetalViews(sheetSetup As SolidEdgeDraft.SheetSetup,
                                                 initialOriginX As Double,
                                                 initialOriginY As Double,
                                                 frontView As SolidEdgeDraft.DrawingView,
                                                 topView As SolidEdgeDraft.DrawingView,
                                                 rightView As SolidEdgeDraft.DrawingView,
                                                 isoView As SolidEdgeDraft.DrawingView)

        If sheetSetup Is Nothing Then
            Return
        End If

        Dim views As New List(Of ViewLayoutInfo) From {
            CreateViewLayoutInfo(frontView, initialOriginX, initialOriginY),
            CreateViewLayoutInfo(topView, initialOriginX, initialOriginY),
            CreateViewLayoutInfo(rightView, initialOriginX, initialOriginY),
            CreateViewLayoutInfo(isoView, initialOriginX, initialOriginY)
        }

        For Each info In views
            PrepareAutoLayoutView(info.View)
            RefreshViewLayoutInfo(info)
        Next

        Dim usableWidth = CDbl(sheetSetup.SheetWidth) - CDbl(sheetSetup.LeftMargin) - CDbl(sheetSetup.RightMargin)
        Dim usableHeight = CDbl(sheetSetup.SheetHeight) - CDbl(sheetSetup.TopMargin) - CDbl(sheetSetup.BottomMargin)
        Dim gap As Double = Math.Min(usableWidth, usableHeight) * 0.004
        Dim outerMargin As Double = Math.Min(usableWidth, usableHeight) * 0.006
        Dim layoutLeft As Double = CDbl(sheetSetup.LeftMargin) + outerMargin
        Dim titleBlockReservedHeight As Double = usableHeight * 0.07
        Dim layoutBottom As Double = CDbl(sheetSetup.BottomMargin) + titleBlockReservedHeight + outerMargin
        Dim layoutWidth As Double = usableWidth - (outerMargin * 2)
        Dim layoutHeight As Double = usableHeight - titleBlockReservedHeight - (outerMargin * 2)
        Dim leftColumnWidthTarget As Double = layoutWidth * 0.67
        Dim rightColumnWidthTarget As Double = layoutWidth - leftColumnWidthTarget - gap
        Dim topRowHeightTarget As Double = layoutHeight * 0.58
        Dim bottomRowHeightTarget As Double = layoutHeight - topRowHeightTarget - gap
        Dim isoCellHeightTarget As Double = bottomRowHeightTarget
        Dim topTargetMinY As Double
        Dim sideTargetMinX As Double
        Dim isoTargetMinX As Double
        Dim isoTargetMinY As Double
        Dim targetGroupCenterX As Double
        Dim targetGroupCenterY As Double

        CenterDrawingViewInCell(views(0), layoutLeft, layoutBottom + bottomRowHeightTarget + gap, leftColumnWidthTarget, topRowHeightTarget)
        RefreshViewLayoutInfo(views(0))
        RefreshViewLayoutInfo(views(1))
        RefreshViewLayoutInfo(views(2))
        RefreshViewLayoutInfo(views(3))

        topTargetMinY = layoutBottom + Math.Max(0, (bottomRowHeightTarget - views(1).Height) / 2)
        MoveDrawingViewToLowerLeft(views(1),
                                   views(0).MinX + Math.Max(0, (views(0).Width - views(1).Width) / 2),
                                   topTargetMinY)

        sideTargetMinX = layoutLeft + leftColumnWidthTarget + gap + Math.Max(0, (rightColumnWidthTarget - views(2).Width) / 2)
        MoveDrawingViewToLowerLeft(views(2),
                                   sideTargetMinX,
                                   views(0).MinY + Math.Max(0, (views(0).Height - views(2).Height) / 2))

        isoTargetMinX = layoutLeft + leftColumnWidthTarget + gap + Math.Max(0, (rightColumnWidthTarget - views(3).Width) / 2)
        isoTargetMinY = layoutBottom + Math.Max(0, (isoCellHeightTarget - views(3).Height) / 2)
        MoveDrawingViewToLowerLeft(views(3), isoTargetMinX, isoTargetMinY)

        targetGroupCenterX = CDbl(sheetSetup.LeftMargin) + (usableWidth / 2)
        targetGroupCenterY = layoutBottom + (layoutHeight / 2)
        CenterViewGroup(views, targetGroupCenterX, targetGroupCenterY)
    End Sub

    Private Function CalculateAutoLayoutScaleFactor(sheetSetup As SolidEdgeDraft.SheetSetup,
                                                    mainView As ViewLayoutInfo,
                                                    topView As ViewLayoutInfo,
                                                    sideView As ViewLayoutInfo,
                                                    isoView As ViewLayoutInfo,
                                                    baseScale As Double) As Double

        Dim usableWidth = CDbl(sheetSetup.SheetWidth) - CDbl(sheetSetup.LeftMargin) - CDbl(sheetSetup.RightMargin)
        Dim usableHeight = CDbl(sheetSetup.SheetHeight) - CDbl(sheetSetup.TopMargin) - CDbl(sheetSetup.BottomMargin)
        Dim gap As Double = Math.Min(usableWidth, usableHeight) * 0.004
        Dim outerMargin As Double = Math.Min(usableWidth, usableHeight) * 0.006
        Dim layoutWidth As Double = usableWidth - (outerMargin * 2)
        Dim titleBlockReservedHeight As Double = usableHeight * 0.07
        Dim layoutHeight As Double = usableHeight - titleBlockReservedHeight - (outerMargin * 2)
        Dim leftColumnWidthTarget As Double = layoutWidth * 0.67
        Dim rightColumnWidthTarget As Double = layoutWidth - leftColumnWidthTarget - gap
        Dim topRowHeightTarget As Double = layoutHeight * 0.58
        Dim bottomRowHeightTarget As Double = layoutHeight - topRowHeightTarget - gap
        Dim isoCellHeightTarget As Double = bottomRowHeightTarget
        Dim scaleFactor As Double = Double.MaxValue

        scaleFactor = Math.Min(scaleFactor, GetCellScaleFactor(mainView, leftColumnWidthTarget, topRowHeightTarget, 0.999))
        scaleFactor = Math.Min(scaleFactor, GetCellScaleFactor(topView, leftColumnWidthTarget, bottomRowHeightTarget, 0.997))
        scaleFactor = Math.Min(scaleFactor, GetCellScaleFactor(sideView, rightColumnWidthTarget, topRowHeightTarget, 0.997))
        scaleFactor = Math.Min(scaleFactor, GetCellScaleFactor(isoView, rightColumnWidthTarget, isoCellHeightTarget, 0.995))

        If Double.IsInfinity(scaleFactor) OrElse Double.IsNaN(scaleFactor) OrElse scaleFactor <= 0 Then
            Return 1.0
        End If

        Return Math.Max(0.01, baseScale * scaleFactor)
    End Function

    Private Sub PrepareAutoLayoutView(view As SolidEdgeDraft.DrawingView)
        If view Is Nothing Then
            Return
        End If

        Try
            view.DisplayCaption = False
        Catch
        End Try

        Try
            view.DisplayScale = False
        Catch
        End Try

        Try
            view.DisplaySuffix = False
        Catch
        End Try

        Try
            view.SetPerspectiveOff()
        Catch
        End Try

        Try
            view.Update()
        Catch
        End Try
    End Sub

    Private Function CreateViewLayoutInfo(view As SolidEdgeDraft.DrawingView,
                                          originX As Double,
                                          originY As Double) As ViewLayoutInfo

        Return New ViewLayoutInfo With {
            .View = view,
            .OriginX = originX,
            .OriginY = originY
        }
    End Function

    Private Sub RefreshViewLayoutInfo(info As ViewLayoutInfo)
        If info Is Nothing OrElse info.View Is Nothing Then
            Return
        End If

        Dim minX As Double = 0
        Dim minY As Double = 0
        Dim maxX As Double = 0
        Dim maxY As Double = 0

        info.View.Range(minX, minY, maxX, maxY)
        info.MinX = minX
        info.MinY = minY
        info.MaxX = maxX
        info.MaxY = maxY
        info.Width = maxX - minX
        info.Height = maxY - minY
    End Sub

    Private Sub MoveDrawingViewToLowerLeft(info As ViewLayoutInfo,
                                           targetMinX As Double,
                                           targetMinY As Double)

        If info Is Nothing OrElse info.View Is Nothing Then
            Return
        End If

        RefreshViewLayoutInfo(info)

        Dim deltaX = targetMinX - info.MinX
        Dim deltaY = targetMinY - info.MinY

        info.OriginX += deltaX
        info.OriginY += deltaY
        info.View.SetOrigin(info.OriginX, info.OriginY)
        info.View.Update()
        RefreshViewLayoutInfo(info)
    End Sub

    Private Sub CenterDrawingViewInCell(info As ViewLayoutInfo,
                                        cellLeft As Double,
                                        cellBottom As Double,
                                        cellWidth As Double,
                                        cellHeight As Double)

        If info Is Nothing Then
            Return
        End If

        Dim targetMinX As Double = cellLeft + Math.Max(0, (cellWidth - info.Width) / 2)
        Dim targetMinY As Double = cellBottom + Math.Max(0, (cellHeight - info.Height) / 2)

        MoveDrawingViewToLowerLeft(info, targetMinX, targetMinY)
    End Sub

    Private Sub CenterViewGroup(views As IEnumerable(Of ViewLayoutInfo),
                                targetCenterX As Double,
                                targetCenterY As Double)

        Dim groupMinX As Double = Double.MaxValue
        Dim groupMinY As Double = Double.MaxValue
        Dim groupMaxX As Double = Double.MinValue
        Dim groupMaxY As Double = Double.MinValue
        Dim groupCenterX As Double
        Dim groupCenterY As Double
        Dim deltaX As Double
        Dim deltaY As Double

        For Each info In views
            If info Is Nothing OrElse info.View Is Nothing Then
                Continue For
            End If

            RefreshViewLayoutInfo(info)
            groupMinX = Math.Min(groupMinX, info.MinX)
            groupMinY = Math.Min(groupMinY, info.MinY)
            groupMaxX = Math.Max(groupMaxX, info.MaxX)
            groupMaxY = Math.Max(groupMaxY, info.MaxY)
        Next

        If groupMinX = Double.MaxValue OrElse groupMinY = Double.MaxValue Then
            Return
        End If

        groupCenterX = groupMinX + ((groupMaxX - groupMinX) / 2)
        groupCenterY = groupMinY + ((groupMaxY - groupMinY) / 2)
        deltaX = targetCenterX - groupCenterX
        deltaY = targetCenterY - groupCenterY

        For Each info In views
            If info Is Nothing OrElse info.View Is Nothing Then
                Continue For
            End If

            info.OriginX += deltaX
            info.OriginY += deltaY
            info.View.SetOrigin(info.OriginX, info.OriginY)
            info.View.Update()
            RefreshViewLayoutInfo(info)
        Next
    End Sub

    Private Function GetCellScaleFactor(info As ViewLayoutInfo,
                                        cellWidth As Double,
                                        cellHeight As Double,
                                        Optional fillRatio As Double = 0.98) As Double

        If info Is Nothing OrElse info.Width <= 0 OrElse info.Height <= 0 Then
            Return 1.0
        End If

        Return Math.Min((cellWidth * fillRatio) / info.Width, (cellHeight * fillRatio) / info.Height)
    End Function

    Private Function BuildSheetMetalAutoLayoutPlan(modelLink As SolidEdgeDraft.ModelLink) As SheetMetalAutoLayoutPlan
        Dim frontSize = GetProjectedRangeSize(modelLink, SolidEdgeDraft.ViewOrientationConstants.igFrontView)
        Dim topSize = GetProjectedRangeSize(modelLink, SolidEdgeDraft.ViewOrientationConstants.igTopView)
        Dim rightSize = GetProjectedRangeSize(modelLink, SolidEdgeDraft.ViewOrientationConstants.igRightView)

        Dim frontArea = frontSize.Width * frontSize.Height
        Dim topArea = topSize.Width * topSize.Height
        Dim rightArea = rightSize.Width * rightSize.Height

        If topArea >= frontArea AndAlso topArea >= rightArea Then
            Return New SheetMetalAutoLayoutPlan With {
                .MainOrientation = SolidEdgeDraft.ViewOrientationConstants.igTopView
            }
        End If

        If rightArea >= frontArea AndAlso rightArea >= topArea Then
            Return New SheetMetalAutoLayoutPlan With {
                .MainOrientation = SolidEdgeDraft.ViewOrientationConstants.igRightView
            }
        End If

        Return New SheetMetalAutoLayoutPlan With {
            .MainOrientation = SolidEdgeDraft.ViewOrientationConstants.igFrontView
        }
    End Function

    Private Function GetProjectedRangeSize(modelLink As SolidEdgeDraft.ModelLink,
                                           orientation As SolidEdgeDraft.ViewOrientationConstants) As ProjectedRangeSize
        Dim minX As Double = 0
        Dim minY As Double = 0
        Dim maxX As Double = 0
        Dim maxY As Double = 0
        Dim missingValue As Object = System.Reflection.Missing.Value

        modelLink.Range2d(orientation, minX, minY, maxX, maxY, missingValue, missingValue)

        Return New ProjectedRangeSize With {
            .Width = Math.Abs(maxX - minX),
            .Height = Math.Abs(maxY - minY)
        }
    End Function

    Private Class ViewLayoutInfo
        Public Property View As SolidEdgeDraft.DrawingView
        Public Property OriginX As Double
        Public Property OriginY As Double
        Public Property MinX As Double
        Public Property MinY As Double
        Public Property MaxX As Double
        Public Property MaxY As Double
        Public Property Width As Double
        Public Property Height As Double
    End Class

    Private Class ProjectedRangeSize
        Public Property Width As Double
        Public Property Height As Double
    End Class

    Private Class SheetMetalAutoLayoutPlan
        Public Property MainOrientation As SolidEdgeDraft.ViewOrientationConstants
        Public Property FinalScale As Double
    End Class

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
                        modelLinkPath As String,
                        outputJpgFilePath As String)

        Dim objDocuments As SolidEdgeFramework.Documents = Nothing
        Dim objDraft As SolidEdgeDraft.DraftDocument = Nothing
        Dim objSheet As SolidEdgeDraft.Sheet = Nothing
        Dim objModelLinks As SolidEdgeDraft.ModelLinks = Nothing
        Dim objModelLink As SolidEdgeDraft.ModelLink = Nothing
        Dim objDrawingViews As SolidEdgeDraft.DrawingViews = Nothing
        Dim objDrawingView As SolidEdgeDraft.DrawingView = Nothing
        Dim objSheetSetup As SolidEdgeDraft.SheetSetup = Nothing
        Dim sheetWindow As Object = Nothing
        Dim missingValue As Object = System.Reflection.Missing.Value

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
            objSheetSetup = objSheet.SheetSetup

            sheetWindow = seApplication.ActiveWindow
            sheetWindow.Activate()

            Dim usableWidth As Double = CDbl(objSheetSetup.SheetWidth) - CDbl(objSheetSetup.LeftMargin) - CDbl(objSheetSetup.RightMargin)
            Dim usableHeight As Double = CDbl(objSheetSetup.SheetHeight) - CDbl(objSheetSetup.TopMargin) - CDbl(objSheetSetup.BottomMargin)
            Dim orientation = SolidEdgeDraft.ViewOrientationConstants.igTrimetricTopFrontRightView
            Dim minX As Double = 0
            Dim minY As Double = 0
            Dim maxX As Double = 0
            Dim maxY As Double = 0
            Dim projectedWidth As Double
            Dim projectedHeight As Double
            Dim viewScale As Double
            Dim centerX As Double
            Dim centerY As Double
            Dim fileExtension = Path.GetExtension(modelLinkPath).ToLowerInvariant()

            objModelLink.Range2d(orientation, minX, minY, maxX, maxY, missingValue, missingValue)

            projectedWidth = Math.Abs(maxX - minX)
            projectedHeight = Math.Abs(maxY - minY)

            If projectedWidth <= 0 Then projectedWidth = 1
            If projectedHeight <= 0 Then projectedHeight = 1

            viewScale = Math.Min((usableWidth * 0.92) / projectedWidth,
                                 (usableHeight * 0.92) / projectedHeight)

            centerX = CDbl(objSheetSetup.LeftMargin) + (usableWidth / 2)
            centerY = CDbl(objSheetSetup.BottomMargin) + (usableHeight / 2)

            If fileExtension = ".par" Then
                objDrawingView = objDrawingViews.AddPartView(
                    objModelLink,
                    orientation,
                    viewScale,
                    centerX,
                    centerY,
                    SolidEdgeDraft.PartDrawingViewTypeConstants.sePartDesignedView)
            ElseIf fileExtension = ".psm" Then
                objDrawingView = objDrawingViews.AddSheetMetalView(
                    objModelLink,
                    orientation,
                    viewScale,
                    centerX,
                    centerY,
                    SolidEdgeDraft.SheetMetalDrawingViewTypeConstants.seSheetMetalDesignedView)
            Else
                Throw New NotSupportedException(String.Format("Estensione non supportata per export JPG: {0}", fileExtension))
            End If

            objDrawingView.DisplayCaption = False
            sheetWindow.Fit()
            sheetWindow.Update()

            If Not Directory.Exists(Path.GetDirectoryName(outputJpgFilePath)) Then
                Directory.CreateDirectory(Path.GetDirectoryName(outputJpgFilePath))
            End If

            If File.Exists(outputJpgFilePath) Then
                File.Delete(outputJpgFilePath)
            End If

            SaveSheetWindowImageWithFallback(sheetWindow, outputJpgFilePath)

            objDraft.Close()

        Finally
            ReleaseCOMReference(sheetWindow)
            ReleaseCOMReference(objDocuments)
            ReleaseCOMReference(objDraft)
            ReleaseCOMReference(objSheet)
            ReleaseCOMReference(objSheetSetup)
            ReleaseCOMReference(objModelLinks)
            ReleaseCOMReference(objModelLink)
            ReleaseCOMReference(objDrawingViews)
            ReleaseCOMReference(objDrawingView)
        End Try
    End Sub

    Private Sub btnGenerateDisegniDiPiega_Click(sender As System.Object, e As System.EventArgs) Handles btnGenerateDisegniDiPiega.Click
        Try
            If ofdSelectASMFile.ShowDialog() = Windows.Forms.DialogResult.OK Then
                RememberAssemblyPath(ofdSelectASMFile.FileName)
                If GenerateDisegniDiPiega_Execute(ofdSelectASMFile.FileName) Then
                    MessageBox.Show(Me, "Generazione Disegni di Piega completata.", "Informazione", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Else
                    MessageBox.Show(Me, "Generazione Disegni di Piega interrotta.", "Informazione", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If
            End If

        Catch exception As Exception
            DisplayException(exception)
        End Try

    End Sub

#End Region

#Region "====[ Export to STL/STP (PAR/PSM) ]===="

    Public Function Export_Execute(asmFilePath As String, type As String) As Boolean
        Dim exportOptions = GetNeutralExportOptions(type)
        Dim exportService As New NeutralExportService(AddressOf ExportPartDocument,
                                                      AddressOf ExportSheetMetalDocumentDocument,
                                                      AddressOf DisplayException)

        Return ExecuteWithProgress(
            String.Format("Esportazione {0}", type.ToUpperInvariant()),
            Function(progress)
                Return _workflowService.ExecuteWithAssembly(
                    asmFilePath,
                    GetApplicationOptions(),
                    False,
                    Function(app, assembly) exportService.ExportAssembly(app, assembly, exportOptions, progress, AddressOf IsCancellationRequested))
            End Function)
    End Function

    Private Sub btnExportSTL_Click(sender As System.Object, e As System.EventArgs) Handles btnExportSTL.Click
        Try
            If ofdSelectASMFile.ShowDialog() = Windows.Forms.DialogResult.OK Then
                RememberAssemblyPath(ofdSelectASMFile.FileName)
                If Export_Execute(ofdSelectASMFile.FileName, "stl") Then
                    MessageBox.Show(Me, "Esportazione in STL (PAR/PSM) completata.", "Informazione", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Else
                    MessageBox.Show(Me, "Esportazione in STL (PAR/PSM) interrotta.", "Informazione", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If
            End If

        Catch exception As Exception
            DisplayException(exception)
        End Try

    End Sub

    Private Sub btnExportSTP_Click(sender As Object, e As EventArgs) Handles btnExportSTP.Click
        Try
            If ofdSelectASMFile.ShowDialog() = Windows.Forms.DialogResult.OK Then
                RememberAssemblyPath(ofdSelectASMFile.FileName)
                If Export_Execute(ofdSelectASMFile.FileName, "stp") Then
                    MessageBox.Show(Me, "Esportazione in STP (PAR/PSM) completata.", "Informazione", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Else
                    MessageBox.Show(Me, "Esportazione in STP (PAR/PSM) interrotta.", "Informazione", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If
            End If

        Catch exception As Exception
            DisplayException(exception)
        End Try
    End Sub


#End Region

#Region "====[ Export to DXF (PSM) ]===="

    Public Function ExportDXF_Execute(asmFilePath As String) As Boolean
        Dim exportOptions = GetFlatDxfExportOptions()
        Dim exportService As New FlatDxfExportService(AddressOf ExportSheetMetalDocumentToDxf,
                                                      AddressOf DisplayException)

        Return ExecuteWithProgress(
            "Esportazione DXF",
            Function(progress)
                Return _workflowService.ExecuteWithAssembly(
                    asmFilePath,
                    GetApplicationOptions(),
                    False,
                    Function(app, assembly) exportService.ExportAssembly(app, assembly, exportOptions, progress, AddressOf IsCancellationRequested))
            End Function)

    End Function

    Private Sub btnExportDXF_Click(sender As System.Object, e As System.EventArgs) Handles btnExportDXF.Click

        Try
            If ofdSelectASMFile.ShowDialog() = Windows.Forms.DialogResult.OK Then
                RememberAssemblyPath(ofdSelectASMFile.FileName)
                If ExportDXF_Execute(ofdSelectASMFile.FileName) Then
                    MessageBox.Show(Me, "Esportazione in DXF (PSM) completata.", "Informazione", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Else
                    MessageBox.Show(Me, "Esportazione in DXF (PSM) interrotta.", "Informazione", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If
            End If

        Catch exception As Exception
            DisplayException(exception)
        End Try

    End Sub

#End Region

#Region "====[ Convert 'Disegni di Piega' to PDF ]===="

    Private Sub btnConvertDisegniDiPiegaToPdf_Click(sender As System.Object, e As System.EventArgs) Handles btnConvertDisegniDiPiegaToPdf.Click
        Try
            PrepareDraftFolderDialog()
            If fbdSelectDisegniDiPiegaFolder.ShowDialog() = Windows.Forms.DialogResult.OK Then
                RememberDraftFolderPath(fbdSelectDisegniDiPiegaFolder.SelectedPath)
                If ConvertDisegniDiPiegaToPdf_Execute(fbdSelectDisegniDiPiegaFolder.SelectedPath) Then
                    MessageBox.Show(Me, "Conversione PDF completata.", "Informazione", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Else
                    MessageBox.Show(Me, "Conversione PDF interrotta.", "Informazione", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If
            End If

        Catch exception As Exception
            DisplayException(exception)
        End Try
    End Sub

    Public Function ConvertDisegniDiPiegaToPdf_Execute(inputDFTDirectory As String) As Boolean

        Dim publishService As New DraftPublishService()
        Dim publishOptions = GetDraftPublishOptions(inputDFTDirectory)

        Return ExecuteWithProgress(
            "Conversione PDF",
            Function(progress)
                Return _workflowService.ExecuteWithApplication(
                    GetApplicationOptions(),
                    False,
                    Function(app)
                        Return publishService.PublishPdf(app, publishOptions, progress, AddressOf IsCancellationRequested)
                    End Function)
            End Function)

    End Function

    Public Function ConvertDisegniDiPiegaToDWG_Execute(inputDFTDirectory As String) As Boolean

        Dim publishService As New DraftPublishService()
        Dim publishOptions = GetDraftPublishOptions(inputDFTDirectory)

        Return ExecuteWithProgress(
            "Conversione DWG",
            Function(progress)
                Return _workflowService.ExecuteWithApplication(
                    GetApplicationOptions(),
                    False,
                    Function(app)
                        Return publishService.PublishDwg(app, publishOptions, progress, AddressOf IsCancellationRequested)
                    End Function)
            End Function)

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

        lblVersion.Text = String.Format("Versione {0}", System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString())
        lblProgress.Text = "Pronto."
        progressOperations.Minimum = 0
        progressOperations.Value = 0
        btnCancelOperation.Enabled = False
    End Sub

    Private Function ExecuteWithProgress(operationName As String,
                                         work As Func(Of Action(Of Integer, Integer, String), Boolean)) As Boolean

        BeginProgress(operationName)

        Try
            Return work(AddressOf ReportProgress)
        Finally
            EndProgress()
        End Try
    End Function

    Private Sub BeginProgress(operationName As String)
        _currentOperationName = operationName
        _cancelRequested = False
        progressOperations.Style = ProgressBarStyle.Marquee
        progressOperations.MarqueeAnimationSpeed = 30
        progressOperations.Value = 0
        lblProgress.Text = String.Format("{0}: preparazione...", operationName)
        UseWaitCursor = True
        btnCancelOperation.Enabled = True
        RefreshProgressUI()
    End Sub

    Private Sub ReportProgress(processed As Integer,
                               total As Integer,
                               currentFilePath As String)

        progressOperations.Style = ProgressBarStyle.Continuous
        progressOperations.MarqueeAnimationSpeed = 0
        progressOperations.Maximum = Math.Max(1, total)
        progressOperations.Value = Math.Min(processed, progressOperations.Maximum)

        If total <= 0 Then
            lblProgress.Text = String.Format("{0}: nessun file da processare.", _currentOperationName)
        ElseIf processed <= 0 Then
            lblProgress.Text = String.Format("{0}: 0/{1}", _currentOperationName, total)
        Else
            lblProgress.Text = String.Format("{0}: {1}/{2} - {3}",
                                             _currentOperationName,
                                             processed,
                                             total,
                                             Path.GetFileName(currentFilePath))
        End If

        RefreshProgressUI()
    End Sub

    Private Sub EndProgress()
        UseWaitCursor = False
        progressOperations.Style = ProgressBarStyle.Continuous
        progressOperations.MarqueeAnimationSpeed = 0
        progressOperations.Value = 0
        lblProgress.Text = "Pronto."
        _currentOperationName = ""
        btnCancelOperation.Enabled = False
        RefreshProgressUI()
    End Sub

    Private Sub RefreshProgressUI()
        lblProgress.Refresh()
        progressOperations.Refresh()
        btnCancelOperation.Refresh()
        Me.Refresh()
        Application.DoEvents()
    End Sub

    Private Function IsCancellationRequested() As Boolean
        Return _cancelRequested
    End Function

    Private Sub btnCancelOperation_Click(sender As Object, e As EventArgs) Handles btnCancelOperation.Click
        _cancelRequested = True
        lblProgress.Text = String.Format("{0}: interruzione richiesta...", _currentOperationName)
        btnCancelOperation.Enabled = False
        RefreshProgressUI()
    End Sub

    Private Sub RememberAssemblyPath(asmFilePath As String)
        If Not String.IsNullOrWhiteSpace(asmFilePath) AndAlso File.Exists(asmFilePath) Then
            _lastAssemblyPath = asmFilePath
            ofdSelectASMFile.InitialDirectory = Path.GetDirectoryName(asmFilePath)
        End If
    End Sub

    Private Sub RememberDraftFolderPath(folderPath As String)
        If Not String.IsNullOrWhiteSpace(folderPath) AndAlso Directory.Exists(folderPath) Then
            _lastDraftFolderPath = folderPath
        End If
    End Sub

    Private Sub PrepareDraftFolderDialog()
        Dim defaultFolderPath As String = Nothing

        If Not String.IsNullOrWhiteSpace(_lastDraftFolderPath) AndAlso Directory.Exists(_lastDraftFolderPath) Then
            defaultFolderPath = _lastDraftFolderPath
        ElseIf Not String.IsNullOrWhiteSpace(_lastAssemblyPath) AndAlso File.Exists(_lastAssemblyPath) Then
            Dim assemblyDirectory = Path.GetDirectoryName(_lastAssemblyPath)
            Dim draftDirectory = Path.Combine(assemblyDirectory, "Disegni di Piega")

            If Directory.Exists(draftDirectory) Then
                defaultFolderPath = draftDirectory
            ElseIf Directory.Exists(assemblyDirectory) Then
                defaultFolderPath = assemblyDirectory
            End If
        End If

        If Not String.IsNullOrWhiteSpace(defaultFolderPath) Then
            fbdSelectDisegniDiPiegaFolder.SelectedPath = defaultFolderPath
        End If
    End Sub

    Private Sub bntPropBOM_Click(sender As Object, e As EventArgs) Handles bntPropBOM.Click
        BOM_Generate(True)
    End Sub


#Region "====[ Crea file JPG  ]===="

    Private Sub btnExportJPG_Click(sender As Object, e As EventArgs) Handles btnExportJPG.Click


        Try
            If ofdSelectASMFile.ShowDialog() = Windows.Forms.DialogResult.OK Then
                RememberAssemblyPath(ofdSelectASMFile.FileName)
                If ExportJPG_Execute(ofdSelectASMFile.FileName) Then
                    MessageBox.Show(Me, "Esportazione in JPG (PAR/PSM) completata.", "Informazione", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Else
                    MessageBox.Show(Me, "Esportazione in JPG (PAR/PSM) interrotta.", "Informazione", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If
            End If

        Catch exception As Exception
            DisplayException(exception)
        End Try

    End Sub

    Private Sub btnProduzioneLamiera_Click(sender As Object, e As EventArgs) Handles btnProduzioneLamiera.Click
        Try
            If ofdSelectASMFile.ShowDialog() = Windows.Forms.DialogResult.OK Then
                RememberAssemblyPath(ofdSelectASMFile.FileName)

                If ProduzioneLamiera_Execute(ofdSelectASMFile.FileName) Then
                    MessageBox.Show(Me, "Produzione Lamiera completata.", "Informazione", MessageBoxButtons.OK, MessageBoxIcon.Information)
                ElseIf _cancelRequested Then
                    MessageBox.Show(Me, "Produzione Lamiera interrotta.", "Informazione", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If
            End If

        Catch exception As Exception
            DisplayException(exception)
        End Try
    End Sub

    Public Function ExportJPG_Execute(asmFilePath As String) As Boolean
        Dim exportOptions = GetImageExportOptions()
        Dim exportService As New ImageExportService(AddressOf ExportModelDocumentImage,
                                                    AddressOf DisplayException)

        Return ExecuteWithProgress(
            "Esportazione JPG",
            Function(progress)
                Return _workflowService.ExecuteWithAssembly(
                    asmFilePath,
                    GetApplicationOptions(),
                    False,
                    Function(app, assembly) exportService.ExportAssembly(app, assembly, exportOptions, progress, AddressOf IsCancellationRequested))
            End Function)
    End Function

    Private Function ProduzioneLamiera_Execute(asmFilePath As String) As Boolean
        Dim sheetMetalMaterials = GetSheetMetalProductionMaterialSelection()
        Dim bomOptions = BuildSheetMetalProductionBomOptions(sheetMetalMaterials)
        Dim dxfOptions = BuildSheetMetalProductionDxfOptions(sheetMetalMaterials)
        Dim draftOptions = BuildSheetMetalProductionDraftOptions(sheetMetalMaterials)
        Dim bomService As New BomService(AddressOf PsmGetProperty)
        Dim dxfService As New FlatDxfExportService(AddressOf ExportSheetMetalDocumentToDxf,
                                                   AddressOf DisplayException)
        Dim draftService As New DraftGenerationService(Sub(app, outputPath, modelLinkPath) DisegniDiPiega_ExportDFT(app, outputPath, modelLinkPath, draftOptions.Scale, draftOptions.AutoLayoutSheetMetalViews),
                                                       AddressOf DisplayException)
        Dim assemblyDirectory = Path.GetDirectoryName(asmFilePath)
        Dim supplierBomPath = Path.Combine(assemblyDirectory, "Lista_" & Path.GetFileNameWithoutExtension(asmFilePath) & ".xlsx")
        Dim expectedTargetFiles As List(Of String) = Nothing
        Dim expectedLabels As HashSet(Of String) = Nothing
        Dim expectedCount As Integer = 0

        BeginProgress("Produzione Lamiera - BOM")

        Try
            Dim bomBuilt = _workflowService.ExecuteWithAssembly(
                asmFilePath,
                GetApplicationOptions(),
                False,
                Function(app, assembly)
                    expectedTargetFiles = GetSheetMetalProductionTargets(assembly, dxfOptions)
                    expectedLabels = New HashSet(Of String)(
                        expectedTargetFiles.Select(Function(targetPath) bomOptions.Prefix & System.IO.Path.GetFileNameWithoutExtension(targetPath)),
                        StringComparer.OrdinalIgnoreCase)

                    Dim bomAssembly = bomService.Build(assembly.FullName, assembly.Occurrences)
                    Dim supplierArray = bomService.ToSupplierArray(bomAssembly, bomOptions)
                    Dim filteredArray = FilterSpreadsheetArrayByFirstColumn(supplierArray, expectedLabels)

                    expectedCount = CountSpreadsheetDataRows(filteredArray)
                    WriteSpreadsheetFromArray(filteredArray, supplierBomPath, False)
                    Return True
                End Function)

            If Not bomBuilt Then
                Return False
            End If
        Finally
            EndProgress()
        End Try

        If expectedCount = 0 Then
            MessageBox.Show(Me, "Nessun particolare in lamiera trovato nell'assieme selezionato.", "Validazione Produzione Lamiera", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return False
        End If

        Dim dxfExported = ExecuteWithProgress(
            "Produzione Lamiera - DXF",
            Function(progress)
                Return _workflowService.ExecuteWithAssembly(
                    asmFilePath,
                    GetApplicationOptions(),
                    False,
                    Function(app, assembly) dxfService.ExportAssembly(app, assembly, dxfOptions, progress, AddressOf IsCancellationRequested))
            End Function)

        If Not dxfExported Then
            Return False
        End If

        Dim dxfIssues = ValidateExportedFiles(Path.Combine(assemblyDirectory, "dxf"),
                                              expectedTargetFiles,
                                              bomOptions.Prefix,
                                              ".dxf",
                                              15 * 1024)

        If dxfIssues.Count > 0 Then
            MessageBox.Show(Me,
                            "Verifica DXF non soddisfatta:" & Environment.NewLine & String.Join(Environment.NewLine, dxfIssues),
                            "Validazione Produzione Lamiera",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Warning)
            Return False
        End If

        ArchiveExistingProductionFolder(Path.Combine(assemblyDirectory, "Disegni di Piega"))

        Dim dftGenerated = ExecuteWithProgress(
            "Produzione Lamiera - DFT",
            Function(progress)
                Return _workflowService.ExecuteWithAssembly(
                    asmFilePath,
                    GetApplicationOptions(),
                    False,
                    Function(app, assembly) draftService.GenerateForAssembly(app, assembly, draftOptions, progress, AddressOf IsCancellationRequested))
            End Function)

        If Not dftGenerated Then
            Return False
        End If

        Dim dftIssues = ValidateExportedFiles(Path.Combine(assemblyDirectory, "Disegni di Piega"),
                                              expectedTargetFiles,
                                              bomOptions.Prefix,
                                              ".dft",
                                              0)

        If dftIssues.Count > 0 Then
            MessageBox.Show(Me,
                            "Verifica DFT non soddisfatta:" & Environment.NewLine & String.Join(Environment.NewLine, dftIssues),
                            "Validazione Produzione Lamiera",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Warning)
            Return False
        End If

        Return True
    End Function

    Private Sub btnCodificaProgetto_Click(sender As Object, e As EventArgs) Handles btnCodificaProgetto.Click
        Dim session As SolidEdgeSessionContext = Nothing
        Dim objApp As SolidEdgeFramework.Application = Nothing
        Dim objDocuments As SolidEdgeFramework.Documents = Nothing
        Dim xlsArray(100, 3) As String
        Dim index As Integer = 0
        Dim projectCodingOptions = GetProjectCodingOptions()
        Dim objPropSets As SolidEdgeFileProperties.PropertySets = New SolidEdgeFileProperties.PropertySets
        Dim objProp As SolidEdgeFileProperties.Property = Nothing
        Dim objProps As SolidEdgeFileProperties.Properties = Nothing


        Try
            If ofdSelectPSMFile.ShowDialog() = Windows.Forms.DialogResult.OK Then


                session = SolidEdgeSessionHelpers.OpenApplication(True)
                objApp = session.Application
                ' Turn off alerts. Weldment environment will display a warning
                objApp.DisplayAlerts = True
                ' Get a reference to the Documents collection
                objDocuments = objApp.Documents
                ' Create an instance of each document environment
                Dim sDocument As String = ofdSelectPSMFile.FileName


                objPropSets.Open(sDocument, False)


                objProps = objPropSets.Item("ProjectInformation")

                objProps.Item("Project Name").Value = projectCodingOptions.ProjectName
                objProps.Item("Revision").Value = projectCodingOptions.Revision
                objProps.Item("Document Number").Value = projectCodingOptions.DocumentNumber

                objProps.Save()
                objPropSets.Save()
                objPropSets.Close()

            End If

        Catch exception As Exception
            DisplayException(exception)
        Finally
            ReleaseCOMReference(objProp)
            ReleaseCOMReference(objProps)
            If Not objPropSets Is Nothing Then
                objPropSets.Close()
                ReleaseCOMReference(objPropSets)
            End If
            ReleaseCOMReference(objDocuments)
            SE_CloseApplication(session, True)
        End Try
    End Sub

#End Region

#Region "====[ Genera Viste 3D, propedeutico per STL/STP list]===="

    Private Sub btnGenerateDisegni3D_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub btnConvertDisegniDiPiegaToDWG_Click(sender As Object, e As EventArgs) Handles btnConvertDisegniDiPiegaToDWG.Click
        Try
            PrepareDraftFolderDialog()
            If fbdSelectDisegniDiPiegaFolder.ShowDialog() = Windows.Forms.DialogResult.OK Then
                RememberDraftFolderPath(fbdSelectDisegniDiPiegaFolder.SelectedPath)
                If ConvertDisegniDiPiegaToDWG_Execute(fbdSelectDisegniDiPiegaFolder.SelectedPath) Then
                    MessageBox.Show(Me, "Conversione DWG completata.", "Informazione", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Else
                    MessageBox.Show(Me, "Conversione DWG interrotta.", "Informazione", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If
            End If

        Catch exception As Exception
            DisplayException(exception)
        End Try
    End Sub

    Private Function GetApplicationOptions() As SolidEdgeApplicationOptions
        Return _configurationEngine.CreateApplicationOptions(GetProductConfiguration())
    End Function

    Private Function GetMaterialSelectionOptions() As MaterialSelectionOptions
        Return _configurationEngine.CreateMaterialSelectionOptions(GetProductConfiguration())
    End Function

    Private Function GetBomExportOptions() As BomExportOptions
        Return _configurationEngine.CreateBomExportOptions(GetProductConfiguration())
    End Function

    Private Function GetNeutralExportOptions(exportType As String) As NeutralExportOptions
        Return _configurationEngine.CreateNeutralExportOptions(GetProductConfiguration(), exportType)
    End Function

    Private Function GetFlatDxfExportOptions() As FlatDxfExportOptions
        Return _configurationEngine.CreateFlatDxfExportOptions(GetProductConfiguration())
    End Function

    Private Function GetImageExportOptions() As ImageExportOptions
        Return _configurationEngine.CreateImageExportOptions(GetProductConfiguration())
    End Function

    Private Function GetDraftGenerationOptions() As DraftGenerationOptions
        Return _configurationEngine.CreateDraftGenerationOptions(GetProductConfiguration())
    End Function

    Private Function GetDraftPublishOptions(inputDirectory As String) As DraftPublishOptions
        Return New DraftPublishOptions() With {
            .InputDirectory = inputDirectory
        }
    End Function

    Private Function GetProjectCodingOptions() As ProjectCodingOptions
        Return _configurationEngine.CreateProjectCodingOptions(GetProductConfiguration())
    End Function

    Private Function GetSheetMetalProductionMaterialSelection() As MaterialSelectionOptions
        Dim options As New MaterialSelectionOptions()

        For Each item As Object In Material.Items
            Dim materialName = item.ToString()

            If materialName.IndexOf("LAMIER", StringComparison.OrdinalIgnoreCase) >= 0 Then
                options.SelectedMaterials.Add(materialName)
            End If
        Next

        Return options
    End Function

    Private Function BuildSheetMetalProductionBomOptions(materialSelection As MaterialSelectionOptions) As BomExportOptions
        Return New BomExportOptions() With {
            .Prefix = Prefisso.Text,
            .MaterialSelection = materialSelection
        }
    End Function

    Private Function BuildSheetMetalProductionDxfOptions(materialSelection As MaterialSelectionOptions) As FlatDxfExportOptions
        Return New FlatDxfExportOptions() With {
            .Prefix = Prefisso.Text,
            .IncludeSubAssemblies = all_subasm.Checked,
            .MaterialSelection = materialSelection
        }
    End Function

    Private Function BuildSheetMetalProductionDraftOptions(materialSelection As MaterialSelectionOptions) As DraftGenerationOptions
        Return New DraftGenerationOptions() With {
            .Prefix = Prefisso.Text,
            .Scale = CDbl(txtScale.Text),
            .AutoLayoutSheetMetalViews = chkAutoLayoutDft.Checked,
            .MaterialSelection = materialSelection
        }
    End Function

    Private Function GetSheetMetalProductionTargets(assembly As SolidEdgeAssembly.AssemblyDocument,
                                                    options As FlatDxfExportOptions) As List(Of String)

        Dim walker As New OccurrenceWalker()
        Dim targets As New List(Of String)
        Dim uniqueFiles As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

        walker.Walk(
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
                    targets.Add(item.OccurrenceFileName)
                End If

                Return True
            End Function)

        Return targets
    End Function

    Private Function FilterSpreadsheetArrayByFirstColumn(sourceArray As Array,
                                                         allowedValues As HashSet(Of String)) As Array

        Dim rowIndexes As New List(Of Integer) From {0}
        Dim columnUpperBound = sourceArray.GetUpperBound(1)

        For rowIndex As Integer = 1 To sourceArray.GetUpperBound(0)
            Dim firstColumnValue = Convert.ToString(sourceArray.GetValue(rowIndex, 0))

            If allowedValues.Contains(firstColumnValue) Then
                rowIndexes.Add(rowIndex)
            End If
        Next

        Dim filteredArray(rowIndexes.Count - 1, columnUpperBound) As String

        For targetRow As Integer = 0 To rowIndexes.Count - 1
            Dim sourceRow = rowIndexes(targetRow)

            For columnIndex As Integer = 0 To columnUpperBound
                filteredArray(targetRow, columnIndex) = Convert.ToString(sourceArray.GetValue(sourceRow, columnIndex))
            Next
        Next

        Return filteredArray
    End Function

    Private Function CountSpreadsheetDataRows(sourceArray As Array) As Integer
        Dim count As Integer = 0

        For rowIndex As Integer = 1 To sourceArray.GetUpperBound(0)
            If Not String.IsNullOrWhiteSpace(Convert.ToString(sourceArray.GetValue(rowIndex, 0))) Then
                count += 1
            End If
        Next

        Return count
    End Function

    Private Function ValidateExportedFiles(outputDirectory As String,
                                           expectedTargetFiles As IEnumerable(Of String),
                                           prefix As String,
                                           extension As String,
                                           minFileSizeBytes As Long) As List(Of String)

        Dim issues As New List(Of String)
        Dim actualCount As Integer = 0
        Dim targetFiles As New List(Of String)(expectedTargetFiles)
        Dim expectedCount = targetFiles.Count

        If Not Directory.Exists(outputDirectory) Then
            issues.Add(String.Format("Cartella output mancante: {0}", outputDirectory))
            Return issues
        End If

        For Each targetFile In targetFiles
            Dim expectedOutputFile = Path.Combine(outputDirectory,
                                                  prefix & Path.GetFileNameWithoutExtension(targetFile) & extension)

            If Not File.Exists(expectedOutputFile) Then
                issues.Add(String.Format("File mancante: {0}", Path.GetFileName(expectedOutputFile)))
                Continue For
            End If

            actualCount += 1

            If minFileSizeBytes > 0 Then
                Dim fileLength = New FileInfo(expectedOutputFile).Length

                If fileLength <= minFileSizeBytes Then
                    issues.Add(String.Format("File troppo piccolo ({0} KB): {1}",
                                             fileLength \ 1024,
                                             Path.GetFileName(expectedOutputFile)))
                End If
            End If
        Next

        If actualCount <> expectedCount Then
            issues.Insert(0, String.Format("Quantità attesa {0}, trovata {1}.", expectedCount, actualCount))
        End If

        Return issues
    End Function

    Private Sub ArchiveExistingProductionFolder(folderPath As String)
        If Not Directory.Exists(folderPath) Then
            Return
        End If

        Dim archivedFolderPath = folderPath & "_old"
        Dim suffixIndex As Integer = 1

        While Directory.Exists(archivedFolderPath)
            archivedFolderPath = String.Format("{0}_old_{1:00}", folderPath, suffixIndex)
            suffixIndex += 1
        End While

        Directory.Move(folderPath, archivedFolderPath)
    End Sub

    Private Function GetProductConfiguration() As ProductConfiguration
        Dim input As New ConfigurationInputModel() With {
            .Prefix = Prefisso.Text,
            .Scale = CDbl(txtScale.Text),
            .AutoLayoutSheetMetalViews = chkAutoLayoutDft.Checked,
            .IncludeSubAssemblies = all_subasm.Checked,
            .MakeApplicationVisible = se_off.CheckState,
            .ProjectName = txtProgetto.Text,
            .Revision = txtVersione.Text,
            .DocumentNumber = txtProgressivo.Text
        }

        For Each item In Material.CheckedItems
            input.SelectedMaterials.Add(item.ToString())
        Next

        Return _configurationEngine.Build(input)
    End Function



#End Region


End Class
