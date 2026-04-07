Imports System.Runtime.InteropServices

Public Module FilePropertyService

    Public Function GetPropertyValue(path As String,
                                     propertySetName As String,
                                     propertyName As String) As String

        Dim objPropertySets As SolidEdgeFileProperties.PropertySets = Nothing
        Dim objProperties As SolidEdgeFileProperties.Properties = Nothing
        Dim objProperty As SolidEdgeFileProperties.Property = Nothing

        Try
            objPropertySets = New SolidEdgeFileProperties.PropertySets
            objPropertySets.Open(path, True)

            For Each objProperties In objPropertySets
                If objProperties.Name = propertySetName Then
                    For Each objProperty In objProperties
                        If objProperty.Name = propertyName Then
                            Return Convert.ToString(objProperty.Value)
                        End If
                    Next
                End If
            Next
        Catch
        Finally
            If Not objProperty Is Nothing Then
                Marshal.ReleaseComObject(objProperty)
                objProperty = Nothing
            End If

            If Not objProperties Is Nothing Then
                Marshal.ReleaseComObject(objProperties)
                objProperties = Nothing
            End If

            If Not objPropertySets Is Nothing Then
                objPropertySets.Close()
                Marshal.ReleaseComObject(objPropertySets)
                objPropertySets = Nothing
            End If
        End Try

        Return Nothing
    End Function
End Module
