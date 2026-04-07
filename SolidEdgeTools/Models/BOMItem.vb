Public Class BOMItem
    Public Name As String
    Public LastSaveDate As String
    Public RevisionNumber As String
    Public Thickness As String
    Public BendRadius As String
    Public NeutralFactor As String
    Public Material As String
    Public Count As Integer
    Public Items As New List(Of BOMItem)
End Class

Public Class BOMAssembly
    Inherits BOMItem

    Public Sub IncCount()
        Count += 1
    End Sub
End Class
