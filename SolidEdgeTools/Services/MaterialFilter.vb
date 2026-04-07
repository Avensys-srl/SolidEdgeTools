Imports System.Collections.Generic

Public Module MaterialFilter

    Public Function MatchesSelectedMaterial(itemMaterial As String,
                                            selectedMaterials As IEnumerable(Of String)) As Boolean

        If String.IsNullOrEmpty(itemMaterial) Then
            Return False
        End If

        For Each material As String In selectedMaterials
            If itemMaterial.Contains(material) Then
                Return True
            End If
        Next

        Return False
    End Function
End Module
