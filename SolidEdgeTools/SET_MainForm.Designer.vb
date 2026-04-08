<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class SET_MainForm
    Inherits System.Windows.Forms.Form

    'Form esegue l'override del metodo Dispose per pulire l'elenco dei componenti.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Richiesto da Progettazione Windows Form
    Private components As System.ComponentModel.IContainer

    'NOTA: la procedura che segue è richiesta da Progettazione Windows Form
    'Può essere modificata in Progettazione Windows Form.  
    'Non modificarla nell'editor del codice.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(SET_MainForm))
        Me.btnPropTable = New System.Windows.Forms.Button()
        Me.btnGenerateDisegniDiPiega = New System.Windows.Forms.Button()
        Me.btnExportDXF = New System.Windows.Forms.Button()
        Me.btnExportSTL = New System.Windows.Forms.Button()
        Me.ofdSelectASMFile = New System.Windows.Forms.OpenFileDialog()
        Me.sfdSelectXLSFile = New System.Windows.Forms.SaveFileDialog()
        Me.Prefisso = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.se_off = New System.Windows.Forms.CheckBox()
        Me.btnGenerateBOMSupplier = New System.Windows.Forms.Button()
        Me.btnConvertDisegniDiPiegaToPdf = New System.Windows.Forms.Button()
        Me.fbdSelectDisegniDiPiegaFolder = New System.Windows.Forms.FolderBrowserDialog()
        Me.all_subasm = New System.Windows.Forms.CheckBox()
        Me.cir_on = New System.Windows.Forms.CheckBox()
        Me.SoloMateriale = New System.Windows.Forms.CheckBox()
        Me.SubFolders = New System.Windows.Forms.CheckBox()
        Me.Material = New System.Windows.Forms.CheckedListBox()
        Me.ofdSelectPSMFile = New System.Windows.Forms.OpenFileDialog()
        Me.bntPropBOM = New System.Windows.Forms.Button()
        Me.btnExportJPG = New System.Windows.Forms.Button()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtScale = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.txtVersione = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.txtProgetto = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.txtProgressivo = New System.Windows.Forms.TextBox()
        Me.btnCodificaProgetto = New System.Windows.Forms.Button()
        Me.btnExportSTP = New System.Windows.Forms.Button()
        Me.btnConvertDisegniDiPiegaToDWG = New System.Windows.Forms.Button()
        Me.lblVersion = New System.Windows.Forms.Label()
        Me.lblProgress = New System.Windows.Forms.Label()
        Me.progressOperations = New System.Windows.Forms.ProgressBar()
        Me.btnCancelOperation = New System.Windows.Forms.Button()
        Me.btnProduzioneLamiera = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'btnPropTable
        '
        Me.btnPropTable.Location = New System.Drawing.Point(105, 36)
        Me.btnPropTable.Name = "btnPropTable"
        Me.btnPropTable.Size = New System.Drawing.Size(117, 25)
        Me.btnPropTable.TabIndex = 0
        Me.btnPropTable.Text = "Tabella Proprietà"
        Me.btnPropTable.UseVisualStyleBackColor = True
        '
        'btnGenerateDisegniDiPiega
        '
        Me.btnGenerateDisegniDiPiega.Location = New System.Drawing.Point(12, 98)
        Me.btnGenerateDisegniDiPiega.Name = "btnGenerateDisegniDiPiega"
        Me.btnGenerateDisegniDiPiega.Size = New System.Drawing.Size(185, 25)
        Me.btnGenerateDisegniDiPiega.TabIndex = 0
        Me.btnGenerateDisegniDiPiega.Text = "Genera Disegni di Piega o Viste 3D"
        Me.btnGenerateDisegniDiPiega.UseVisualStyleBackColor = True
        '
        'btnExportDXF
        '
        Me.btnExportDXF.Location = New System.Drawing.Point(143, 67)
        Me.btnExportDXF.Name = "btnExportDXF"
        Me.btnExportDXF.Size = New System.Drawing.Size(79, 25)
        Me.btnExportDXF.TabIndex = 0
        Me.btnExportDXF.Text = "Esporta DXF"
        Me.btnExportDXF.UseVisualStyleBackColor = True
        '
        'btnExportSTL
        '
        Me.btnExportSTL.Location = New System.Drawing.Point(228, 67)
        Me.btnExportSTL.Name = "btnExportSTL"
        Me.btnExportSTL.Size = New System.Drawing.Size(74, 25)
        Me.btnExportSTL.TabIndex = 0
        Me.btnExportSTL.Text = "Esporta STL"
        Me.btnExportSTL.UseVisualStyleBackColor = True
        '
        'ofdSelectASMFile
        '
        Me.ofdSelectASMFile.Filter = "File asm|*.asm"
        Me.ofdSelectASMFile.Title = "Seleziona il file SolidEdge"
        '
        'sfdSelectXLSFile
        '
        Me.sfdSelectXLSFile.Filter = "File xls|*.xlsx"
        Me.sfdSelectXLSFile.Title = "Seleziona il file Excel per l'output"
        '
        'Prefisso
        '
        Me.Prefisso.Location = New System.Drawing.Point(95, 223)
        Me.Prefisso.Multiline = True
        Me.Prefisso.Name = "Prefisso"
        Me.Prefisso.Size = New System.Drawing.Size(202, 20)
        Me.Prefisso.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 226)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(77, 13)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Prefisso dei file"
        '
        'se_off
        '
        Me.se_off.AutoSize = True
        Me.se_off.Location = New System.Drawing.Point(139, 139)
        Me.se_off.Name = "se_off"
        Me.se_off.Size = New System.Drawing.Size(118, 17)
        Me.se_off.TabIndex = 3
        Me.se_off.Text = "Solide Edge Visibile"
        Me.se_off.UseVisualStyleBackColor = True
        '
        'btnGenerateBOMSupplier
        '
        Me.btnGenerateBOMSupplier.Location = New System.Drawing.Point(12, 67)
        Me.btnGenerateBOMSupplier.Name = "btnGenerateBOMSupplier"
        Me.btnGenerateBOMSupplier.Size = New System.Drawing.Size(125, 25)
        Me.btnGenerateBOMSupplier.TabIndex = 0
        Me.btnGenerateBOMSupplier.Text = "BOM Fornitore"
        Me.btnGenerateBOMSupplier.UseVisualStyleBackColor = True
        '
        'btnConvertDisegniDiPiegaToPdf
        '
        Me.btnConvertDisegniDiPiegaToPdf.Location = New System.Drawing.Point(202, 98)
        Me.btnConvertDisegniDiPiegaToPdf.Name = "btnConvertDisegniDiPiegaToPdf"
        Me.btnConvertDisegniDiPiegaToPdf.Size = New System.Drawing.Size(81, 25)
        Me.btnConvertDisegniDiPiegaToPdf.TabIndex = 0
        Me.btnConvertDisegniDiPiegaToPdf.Text = "DFT --> PDF"
        Me.btnConvertDisegniDiPiegaToPdf.UseVisualStyleBackColor = True
        '
        'all_subasm
        '
        Me.all_subasm.AutoSize = True
        Me.all_subasm.Checked = True
        Me.all_subasm.CheckState = System.Windows.Forms.CheckState.Checked
        Me.all_subasm.Location = New System.Drawing.Point(270, 139)
        Me.all_subasm.Name = "all_subasm"
        Me.all_subasm.Size = New System.Drawing.Size(150, 17)
        Me.all_subasm.TabIndex = 4
        Me.all_subasm.Text = "Considera Sottoassemblati"
        Me.all_subasm.UseVisualStyleBackColor = True
        '
        'cir_on
        '
        Me.cir_on.AutoSize = True
        Me.cir_on.Location = New System.Drawing.Point(12, 139)
        Me.cir_on.Name = "cir_on"
        Me.cir_on.Size = New System.Drawing.Size(121, 17)
        Me.cir_on.TabIndex = 5
        Me.cir_on.Text = "Profili Circolari (DXF)"
        Me.cir_on.UseVisualStyleBackColor = True
        '
        'SoloMateriale
        '
        Me.SoloMateriale.AutoSize = True
        Me.SoloMateriale.Location = New System.Drawing.Point(426, 139)
        Me.SoloMateriale.Name = "SoloMateriale"
        Me.SoloMateriale.Size = New System.Drawing.Size(129, 17)
        Me.SoloMateriale.TabIndex = 6
        Me.SoloMateriale.Text = "Esporta solo materiale"
        Me.SoloMateriale.UseVisualStyleBackColor = True
        '
        'SubFolders
        '
        Me.SubFolders.AutoSize = True
        Me.SubFolders.Enabled = False
        Me.SubFolders.Location = New System.Drawing.Point(562, 139)
        Me.SubFolders.Name = "SubFolders"
        Me.SubFolders.Size = New System.Drawing.Size(110, 17)
        Me.SubFolders.TabIndex = 8
        Me.SubFolders.Text = "Crea Sottocartelle"
        Me.SubFolders.UseVisualStyleBackColor = True
        '
        'Material
        '
        Me.Material.FormattingEnabled = True
        Me.Material.Items.AddRange(New Object() {"LAMIERA ZN", "LAMIERA ZNVN", "LAMIERA ZNVV", "LAMIERA FESD", "FabLab", "Polistirolo", "PPE"})
        Me.Material.Location = New System.Drawing.Point(391, 36)
        Me.Material.MultiColumn = True
        Me.Material.Name = "Material"
        Me.Material.Size = New System.Drawing.Size(281, 94)
        Me.Material.TabIndex = 9
        '
        'ofdSelectPSMFile
        '
        Me.ofdSelectPSMFile.Filter = "File psm|*.psm"
        '
        'bntPropBOM
        '
        Me.bntPropBOM.Location = New System.Drawing.Point(12, 36)
        Me.bntPropBOM.Name = "bntPropBOM"
        Me.bntPropBOM.Size = New System.Drawing.Size(87, 25)
        Me.bntPropBOM.TabIndex = 10
        Me.bntPropBOM.Text = "BOM Proprietà"
        Me.bntPropBOM.UseVisualStyleBackColor = True
        '
        'btnExportJPG
        '
        Me.btnExportJPG.Enabled = True
        Me.btnExportJPG.Location = New System.Drawing.Point(228, 36)
        Me.btnExportJPG.Name = "btnExportJPG"
        Me.btnExportJPG.Size = New System.Drawing.Size(74, 25)
        Me.btnExportJPG.TabIndex = 11
        Me.btnExportJPG.Text = "Esporta JPG"
        Me.btnExportJPG.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(558, 229)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(58, 13)
        Me.Label2.TabIndex = 12
        Me.Label2.Text = "Scala DFT"
        '
        'txtScale
        '
        Me.txtScale.Location = New System.Drawing.Point(622, 226)
        Me.txtScale.Multiline = True
        Me.txtScale.Name = "txtScale"
        Me.txtScale.Size = New System.Drawing.Size(50, 20)
        Me.txtScale.TabIndex = 13
        Me.txtScale.Text = "0.5"
        Me.txtScale.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(200, 170)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(83, 13)
        Me.Label3.TabIndex = 2
        Me.Label3.Text = "Codice Progetto"
        '
        'txtVersione
        '
        Me.txtVersione.Location = New System.Drawing.Point(466, 167)
        Me.txtVersione.Multiline = True
        Me.txtVersione.Name = "txtVersione"
        Me.txtVersione.Size = New System.Drawing.Size(66, 20)
        Me.txtVersione.TabIndex = 1
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(412, 170)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(48, 13)
        Me.Label4.TabIndex = 2
        Me.Label4.Text = "Versione"
        '
        'txtProgetto
        '
        Me.txtProgetto.Location = New System.Drawing.Point(289, 167)
        Me.txtProgetto.Multiline = True
        Me.txtProgetto.Name = "txtProgetto"
        Me.txtProgetto.Size = New System.Drawing.Size(117, 20)
        Me.txtProgetto.TabIndex = 1
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(538, 170)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(62, 13)
        Me.Label5.TabIndex = 2
        Me.Label5.Text = "Progressivo"
        '
        'txtProgressivo
        '
        Me.txtProgressivo.Location = New System.Drawing.Point(606, 167)
        Me.txtProgressivo.Multiline = True
        Me.txtProgressivo.Name = "txtProgressivo"
        Me.txtProgressivo.Size = New System.Drawing.Size(66, 20)
        Me.txtProgressivo.TabIndex = 1
        '
        'btnCodificaProgetto
        '
        Me.btnCodificaProgetto.Location = New System.Drawing.Point(12, 164)
        Me.btnCodificaProgetto.Name = "btnCodificaProgetto"
        Me.btnCodificaProgetto.Size = New System.Drawing.Size(185, 25)
        Me.btnCodificaProgetto.TabIndex = 0
        Me.btnCodificaProgetto.Text = "Codifica Progetto"
        Me.btnCodificaProgetto.UseVisualStyleBackColor = True
        '
        'btnExportSTP
        '
        Me.btnExportSTP.Location = New System.Drawing.Point(308, 67)
        Me.btnExportSTP.Name = "btnExportSTP"
        Me.btnExportSTP.Size = New System.Drawing.Size(77, 25)
        Me.btnExportSTP.TabIndex = 15
        Me.btnExportSTP.Text = "Esporta STP"
        Me.btnExportSTP.UseVisualStyleBackColor = True
        '
        'btnConvertDisegniDiPiegaToDWG
        '
        Me.btnConvertDisegniDiPiegaToDWG.Location = New System.Drawing.Point(289, 98)
        Me.btnConvertDisegniDiPiegaToDWG.Name = "btnConvertDisegniDiPiegaToDWG"
        Me.btnConvertDisegniDiPiegaToDWG.Size = New System.Drawing.Size(96, 25)
        Me.btnConvertDisegniDiPiegaToDWG.TabIndex = 0
        Me.btnConvertDisegniDiPiegaToDWG.Text = "DFT --> DWG"
        Me.btnConvertDisegniDiPiegaToDWG.UseVisualStyleBackColor = True
        '
        'lblVersion
        '
        Me.lblVersion.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblVersion.AutoSize = False
        Me.lblVersion.Location = New System.Drawing.Point(559, 12)
        Me.lblVersion.Name = "lblVersion"
        Me.lblVersion.Size = New System.Drawing.Size(113, 13)
        Me.lblVersion.TabIndex = 16
        Me.lblVersion.Text = "Versione 1.0.0.0"
        Me.lblVersion.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblProgress
        '
        Me.lblProgress.AutoSize = False
        Me.lblProgress.Location = New System.Drawing.Point(12, 249)
        Me.lblProgress.Name = "lblProgress"
        Me.lblProgress.Size = New System.Drawing.Size(660, 13)
        Me.lblProgress.TabIndex = 17
        Me.lblProgress.Text = "Pronto."
        '
        'progressOperations
        '
        Me.progressOperations.Location = New System.Drawing.Point(12, 265)
        Me.progressOperations.Name = "progressOperations"
        Me.progressOperations.Size = New System.Drawing.Size(561, 12)
        Me.progressOperations.TabIndex = 18
        '
        'btnCancelOperation
        '
        Me.btnCancelOperation.Enabled = False
        Me.btnCancelOperation.Location = New System.Drawing.Point(579, 259)
        Me.btnCancelOperation.Name = "btnCancelOperation"
        Me.btnCancelOperation.Size = New System.Drawing.Size(93, 23)
        Me.btnCancelOperation.TabIndex = 19
        Me.btnCancelOperation.Text = "Interrompi"
        Me.btnCancelOperation.UseVisualStyleBackColor = True
        '
        'btnProduzioneLamiera
        '
        Me.btnProduzioneLamiera.Location = New System.Drawing.Point(303, 220)
        Me.btnProduzioneLamiera.Name = "btnProduzioneLamiera"
        Me.btnProduzioneLamiera.Size = New System.Drawing.Size(249, 25)
        Me.btnProduzioneLamiera.TabIndex = 20
        Me.btnProduzioneLamiera.Text = "Produzione Lamiera"
        Me.btnProduzioneLamiera.UseVisualStyleBackColor = True
        '
        'SET_MainForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(684, 289)
        Me.Controls.Add(Me.btnProduzioneLamiera)
        Me.Controls.Add(Me.btnCancelOperation)
        Me.Controls.Add(Me.progressOperations)
        Me.Controls.Add(Me.lblProgress)
        Me.Controls.Add(Me.btnExportSTP)
        Me.Controls.Add(Me.lblVersion)
        Me.Controls.Add(Me.txtScale)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.btnExportJPG)
        Me.Controls.Add(Me.bntPropBOM)
        Me.Controls.Add(Me.Material)
        Me.Controls.Add(Me.SubFolders)
        Me.Controls.Add(Me.SoloMateriale)
        Me.Controls.Add(Me.cir_on)
        Me.Controls.Add(Me.all_subasm)
        Me.Controls.Add(Me.se_off)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.txtProgetto)
        Me.Controls.Add(Me.txtProgressivo)
        Me.Controls.Add(Me.txtVersione)
        Me.Controls.Add(Me.Prefisso)
        Me.Controls.Add(Me.btnExportSTL)
        Me.Controls.Add(Me.btnExportDXF)
        Me.Controls.Add(Me.btnCodificaProgetto)
        Me.Controls.Add(Me.btnConvertDisegniDiPiegaToDWG)
        Me.Controls.Add(Me.btnConvertDisegniDiPiegaToPdf)
        Me.Controls.Add(Me.btnGenerateDisegniDiPiega)
        Me.Controls.Add(Me.btnGenerateBOMSupplier)
        Me.Controls.Add(Me.btnPropTable)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximumSize = New System.Drawing.Size(700, 328)
        Me.MinimumSize = New System.Drawing.Size(700, 328)
        Me.Name = "SET_MainForm"
        Me.Text = "SolidEdgeTools"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnPropTable As System.Windows.Forms.Button
    Friend WithEvents btnGenerateDisegniDiPiega As System.Windows.Forms.Button
    Friend WithEvents btnExportDXF As System.Windows.Forms.Button
    Friend WithEvents btnExportSTL As System.Windows.Forms.Button
    Friend WithEvents ofdSelectASMFile As System.Windows.Forms.OpenFileDialog
    Friend WithEvents sfdSelectXLSFile As System.Windows.Forms.SaveFileDialog
    Friend WithEvents Prefisso As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents se_off As System.Windows.Forms.CheckBox
    Friend WithEvents btnGenerateBOMSupplier As System.Windows.Forms.Button
    Friend WithEvents btnConvertDisegniDiPiegaToPdf As System.Windows.Forms.Button
    Friend WithEvents fbdSelectDisegniDiPiegaFolder As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents all_subasm As System.Windows.Forms.CheckBox
    Friend WithEvents cir_on As System.Windows.Forms.CheckBox
    Friend WithEvents SoloMateriale As CheckBox
    Friend WithEvents SubFolders As CheckBox
    Friend WithEvents Material As CheckedListBox
    Friend WithEvents ofdSelectPSMFile As OpenFileDialog
    Friend WithEvents bntPropBOM As Button
    Friend WithEvents btnExportJPG As Button
    Friend WithEvents Label2 As Label
    Friend WithEvents txtScale As TextBox
    Friend WithEvents Label3 As Label
    Friend WithEvents txtVersione As TextBox
    Friend WithEvents Label4 As Label
    Friend WithEvents txtProgetto As TextBox
    Friend WithEvents Label5 As Label
    Friend WithEvents txtProgressivo As TextBox
    Friend WithEvents btnCodificaProgetto As Button
    Friend WithEvents btnExportSTP As Button
    Friend WithEvents btnConvertDisegniDiPiegaToDWG As Button
    Friend WithEvents lblVersion As Label
    Friend WithEvents lblProgress As Label
    Friend WithEvents progressOperations As ProgressBar
    Friend WithEvents btnCancelOperation As Button
    Friend WithEvents btnProduzioneLamiera As Button
End Class
