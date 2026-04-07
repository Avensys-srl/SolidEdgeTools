Public Class GeometryVariable
    Public Property Name As String
    Public Property Value As Double
End Class

Public Class GeometryBuildPlan
    Public Property TemplatePath As String
    Public Property OutputPath As String
    Public Property DocumentType As String
    Public Property Variables As New List(Of GeometryVariable)
End Class

Public Class ConfigurationValidationIssue
    Public Property FieldName As String
    Public Property Message As String
End Class

Public Class ConfigurationValidationResult
    Public Property Issues As New List(Of ConfigurationValidationIssue)

    Public ReadOnly Property IsValid As Boolean
        Get
            Return Issues.Count = 0
        End Get
    End Property
End Class
