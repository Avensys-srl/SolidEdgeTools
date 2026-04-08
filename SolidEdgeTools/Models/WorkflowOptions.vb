Public Class MaterialSelectionOptions
    Public Property SelectedMaterials As New List(Of String)
End Class

Public Class SolidEdgeApplicationOptions
    Public Property MakeVisible As Boolean
End Class

Public Class BomExportOptions
    Public Property Prefix As String
    Public Property MaterialSelection As MaterialSelectionOptions
End Class

Public Class NeutralExportOptions
    Public Property Prefix As String
    Public Property ExportType As String
    Public Property MaterialSelection As MaterialSelectionOptions
End Class

Public Class FlatDxfExportOptions
    Public Property Prefix As String
    Public Property IncludeSubAssemblies As Boolean
    Public Property MaterialSelection As MaterialSelectionOptions
End Class

Public Class ImageExportOptions
    Public Property Prefix As String
    Public Property IncludeSubAssemblies As Boolean
    Public Property MaterialSelection As MaterialSelectionOptions
End Class

Public Class DraftGenerationOptions
    Public Property Prefix As String
    Public Property Scale As Double
    Public Property AutoLayoutSheetMetalViews As Boolean
    Public Property MaterialSelection As MaterialSelectionOptions
End Class

Public Class DraftPublishOptions
    Public Property InputDirectory As String
End Class

Public Class ProjectCodingOptions
    Public Property ProjectName As String
    Public Property Revision As String
    Public Property DocumentNumber As String
End Class
