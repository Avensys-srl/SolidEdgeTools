Public Class ConfigurationValidator

    Public Function Validate(configuration As ProductConfiguration) As ConfigurationValidationResult
        Dim result As New ConfigurationValidationResult()

        If configuration Is Nothing Then
            result.Issues.Add(New ConfigurationValidationIssue() With {
                .FieldName = "Configuration",
                .Message = "Product configuration is required."
            })
            Return result
        End If

        If configuration.Unit Is Nothing Then
            result.Issues.Add(New ConfigurationValidationIssue() With {
                .FieldName = "Unit",
                .Message = "Unit model is required."
            })
            Return result
        End If

        If String.IsNullOrWhiteSpace(configuration.Unit.Prefix) Then
            result.Issues.Add(New ConfigurationValidationIssue() With {
                .FieldName = "Prefix",
                .Message = "Prefix is required."
            })
        End If

        If configuration.Unit.Scale <= 0 Then
            result.Issues.Add(New ConfigurationValidationIssue() With {
                .FieldName = "Scale",
                .Message = "Scale must be greater than zero."
            })
        End If

        If configuration.MaterialSelection Is Nothing Then
            result.Issues.Add(New ConfigurationValidationIssue() With {
                .FieldName = "MaterialSelection",
                .Message = "Material selection is required."
            })
        End If

        If configuration.ApplicationOptions Is Nothing Then
            result.Issues.Add(New ConfigurationValidationIssue() With {
                .FieldName = "ApplicationOptions",
                .Message = "Application options are required."
            })
        End If

        Return result
    End Function
End Class
