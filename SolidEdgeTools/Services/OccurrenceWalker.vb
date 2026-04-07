Public Class OccurrenceWalker

    Public Function Walk(occurrences As SolidEdgeAssembly.Occurrences,
                         includeSubAssemblies As Boolean,
                         visitor As Func(Of SolidEdgeAssembly.Occurrence, Boolean)) As Boolean

        For Each item As SolidEdgeAssembly.Occurrence In occurrences
            If Not visitor(item) Then
                Return False
            End If

            If includeSubAssemblies AndAlso item.Type = SolidEdgeFramework.ObjectType.igSubAssembly Then
                If Not Walk(item.OccurrenceDocument.Occurrences, includeSubAssemblies, visitor) Then
                    Return False
                End If
            End If
        Next

        Return True
    End Function
End Class
