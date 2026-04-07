Imports System.IO

Public Class BomService

    Private ReadOnly _assemblies As New Dictionary(Of String, BOMAssembly)
    Private ReadOnly _propertyReader As Func(Of String, String, String, String)

    Public Sub New(propertyReader As Func(Of String, String, String, String))
        _propertyReader = propertyReader
    End Sub

    Public Function Build(rootAssemblyPath As String,
                          occurrences As SolidEdgeAssembly.Occurrences) As BOMAssembly

        Dim rootAssembly = AddAssembly(Nothing, rootAssemblyPath)

        Populate(rootAssembly, occurrences)

        Return rootAssembly
    End Function

    Public Function ToSupplierArray(bomAssembly As BOMAssembly,
                                    prefix As String,
                                    materialMatcher As Func(Of String, Boolean)) As Array

        Dim index As Integer = 0
        Dim flatBoms As New Dictionary(Of String, BOMItem)

        UpdateCount(bomAssembly, bomAssembly.Count)
        Flatten(flatBoms, bomAssembly)

        Dim values(flatBoms.Count, 4) As String
        values.SetValue("Nome File", index, 0)
        values.SetValue("Spessore", index, 1)
        values.SetValue("Materiale", index, 2)
        values.SetValue("Quantità", index, 3)
        index += 1

        For Each item As BOMItem In flatBoms.Values
            If materialMatcher(item.Material) Then
                values.SetValue(prefix + Path.GetFileNameWithoutExtension(item.Name), index, 0)
                values.SetValue(If(item.Thickness, ""), index, 1)
                values.SetValue(If(item.Material, ""), index, 2)
                values.SetValue(item.Count.ToString(), index, 3)
                index += 1
            End If
        Next

        Return values
    End Function

    Public Function ToPropertyArray(bomAssembly As BOMAssembly,
                                    materialMatcher As Func(Of String, Boolean)) As Array

        Dim index As Integer = 0
        Dim flatBoms As New Dictionary(Of String, BOMItem)

        UpdateCount(bomAssembly, bomAssembly.Count)
        Flatten(flatBoms, bomAssembly)

        Dim values(flatBoms.Count, 8) As String
        values.SetValue("Nome File", index, 0)
        values.SetValue("Spessore", index, 1)
        values.SetValue("Materiale", index, 2)
        values.SetValue("Quantità", index, 3)
        values.SetValue("Raggio di piega", index, 4)
        values.SetValue("Fattore Neutro", index, 5)
        values.SetValue("Revisione", index, 6)
        values.SetValue("Ultimo Salvataggio", index, 7)
        index += 1

        For Each item As BOMItem In flatBoms.Values
            If materialMatcher(item.Material) Then
                values.SetValue(Path.GetFileNameWithoutExtension(item.Name), index, 0)
                values.SetValue(If(item.Thickness, ""), index, 1)
                values.SetValue(If(item.Material, ""), index, 2)
                values.SetValue(item.Count.ToString(), index, 3)
                values.SetValue(If(item.BendRadius, ""), index, 4)
                values.SetValue(If(item.NeutralFactor, ""), index, 5)
                values.SetValue(If(item.RevisionNumber, ""), index, 6)
                values.SetValue(If(item.LastSaveDate, ""), index, 7)
                index += 1
            End If
        Next

        Return values
    End Function

    Private Function AddAssembly(parentAssembly As BOMAssembly, name As String) As BOMAssembly
        Dim bomAssembly As New BOMAssembly() With {
            .Name = name,
            .Count = 1
        }

        _assemblies.Add(name, bomAssembly)

        If Not parentAssembly Is Nothing Then
            parentAssembly.Items.Add(bomAssembly)
        End If

        Return bomAssembly
    End Function

    Private Function AddItem(parentAssembly As BOMAssembly, name As String) As BOMItem
        Dim bomItem As New BOMItem() With {
            .Name = name,
            .Count = 1,
            .LastSaveDate = SafeReadProperty(name, "SummaryInformation", "Last Save Date"),
            .RevisionNumber = SafeReadProperty(name, "SummaryInformation", "Revision Number"),
            .Material = SafeReadProperty(name, "MechanicalModeling", "Material"),
            .Thickness = SafeReadProperty(name, "Custom", "Material Thickness"),
            .BendRadius = SafeReadProperty(name, "Custom", "Bend Radius"),
            .NeutralFactor = SafeReadProperty(name, "Custom", "Neutral Factor")
        }

        parentAssembly.Items.Add(bomItem)

        Return bomItem
    End Function

    Private Sub Populate(parentAssembly As BOMAssembly,
                         occurrences As SolidEdgeAssembly.Occurrences)

        For Each item As SolidEdgeAssembly.Occurrence In occurrences
            If item.Type = SolidEdgeFramework.ObjectType.igSubAssembly Then
                Dim foundAssembly As BOMAssembly = Nothing

                If _assemblies.TryGetValue(item.OccurrenceFileName, foundAssembly) Then
                    foundAssembly.IncCount()
                Else
                    foundAssembly = AddAssembly(parentAssembly, item.OccurrenceFileName)
                    Populate(foundAssembly, item.OccurrenceDocument.Occurrences)
                End If
            Else
                Dim foundItem = FindItemByName(parentAssembly.Items, item.OccurrenceFileName)

                If Not foundItem Is Nothing Then
                    foundItem.Count += 1
                Else
                    AddItem(parentAssembly, item.OccurrenceFileName)
                End If
            End If
        Next
    End Sub

    Private Sub Flatten(flatBoms As Dictionary(Of String, BOMItem),
                        item As BOMItem)

        If TypeOf item Is BOMAssembly Then
            For Each subItem As BOMItem In item.Items
                Flatten(flatBoms, subItem)
            Next
        ElseIf flatBoms.ContainsKey(item.Name) Then
            flatBoms(item.Name).Count += item.Count
        Else
            flatBoms.Add(item.Name, New BOMItem() With {
                .Name = item.Name,
                .Thickness = item.Thickness,
                .Material = item.Material,
                .LastSaveDate = item.LastSaveDate,
                .RevisionNumber = item.RevisionNumber,
                .Count = item.Count,
                .BendRadius = item.BendRadius,
                .NeutralFactor = item.NeutralFactor
            })
        End If
    End Sub

    Private Sub UpdateCount(item As BOMItem, count As Integer)
        If TypeOf item Is BOMAssembly Then
            For Each subItem As BOMItem In item.Items
                UpdateCount(subItem, count * item.Count)
            Next
        Else
            item.Count *= count
        End If
    End Sub

    Private Function FindItemByName(items As List(Of BOMItem), name As String) As BOMItem
        For Each item As BOMItem In items
            If TypeOf item Is BOMItem AndAlso item.Name = name Then
                Return item
            End If
        Next

        Return Nothing
    End Function

    Private Function SafeReadProperty(path As String,
                                      propertySetName As String,
                                      propertyName As String) As String
        Try
            Return _propertyReader(path, propertySetName, propertyName)
        Catch
            Return ""
        End Try
    End Function
End Class
