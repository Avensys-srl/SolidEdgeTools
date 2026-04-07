Public Class ConfigurationInputModel
    Public Property Prefix As String
    Public Property Scale As Double
    Public Property IncludeSubAssemblies As Boolean
    Public Property MakeApplicationVisible As Boolean
    Public Property SelectedMaterials As New List(Of String)
    Public Property ProjectName As String
    Public Property Revision As String
    Public Property DocumentNumber As String
End Class

Public Class ProjectIdentity
    Public Property ProjectName As String
    Public Property Revision As String
    Public Property DocumentNumber As String
End Class

Public Class UnitModel
    Public Property Prefix As String
    Public Property Configuration As String
    Public Property Scale As Double
    Public Property SelectedMaterials As New List(Of String)
End Class

Public Class ProductConfiguration
    Public Property ApplicationOptions As SolidEdgeApplicationOptions
    Public Property MaterialSelection As MaterialSelectionOptions
    Public Property IncludeSubAssemblies As Boolean
    Public Property ProjectIdentity As ProjectIdentity
    Public Property Unit As UnitModel
End Class
