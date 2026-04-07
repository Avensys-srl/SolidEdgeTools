Public Class ConfigurationEngine

    Public Function Build(input As ConfigurationInputModel) As ProductConfiguration
        Dim materialSelection As New MaterialSelectionOptions()
        materialSelection.SelectedMaterials.AddRange(input.SelectedMaterials)

        Return New ProductConfiguration() With {
            .ApplicationOptions = New SolidEdgeApplicationOptions() With {
                .MakeVisible = input.MakeApplicationVisible
            },
            .MaterialSelection = materialSelection,
            .IncludeSubAssemblies = input.IncludeSubAssemblies,
            .ProjectIdentity = New ProjectIdentity() With {
                .ProjectName = input.ProjectName,
                .Revision = input.Revision,
                .DocumentNumber = input.DocumentNumber
            },
            .Unit = New UnitModel() With {
                .Prefix = input.Prefix,
                .Configuration = input.Prefix,
                .Scale = input.Scale,
                .SelectedMaterials = New List(Of String)(input.SelectedMaterials)
            }
        }
    End Function

    Public Function CreateApplicationOptions(configuration As ProductConfiguration) As SolidEdgeApplicationOptions
        Return configuration.ApplicationOptions
    End Function

    Public Function CreateMaterialSelectionOptions(configuration As ProductConfiguration) As MaterialSelectionOptions
        Return configuration.MaterialSelection
    End Function

    Public Function CreateBomExportOptions(configuration As ProductConfiguration) As BomExportOptions
        Return New BomExportOptions() With {
            .Prefix = configuration.Unit.Prefix,
            .MaterialSelection = configuration.MaterialSelection
        }
    End Function

    Public Function CreateNeutralExportOptions(configuration As ProductConfiguration,
                                               exportType As String) As NeutralExportOptions
        Return New NeutralExportOptions() With {
            .Prefix = configuration.Unit.Prefix,
            .ExportType = exportType,
            .MaterialSelection = configuration.MaterialSelection
        }
    End Function

    Public Function CreateFlatDxfExportOptions(configuration As ProductConfiguration) As FlatDxfExportOptions
        Return New FlatDxfExportOptions() With {
            .Prefix = configuration.Unit.Prefix,
            .IncludeSubAssemblies = configuration.IncludeSubAssemblies,
            .MaterialSelection = configuration.MaterialSelection
        }
    End Function

    Public Function CreateImageExportOptions(configuration As ProductConfiguration) As ImageExportOptions
        Return New ImageExportOptions() With {
            .Prefix = configuration.Unit.Prefix,
            .IncludeSubAssemblies = configuration.IncludeSubAssemblies,
            .MaterialSelection = configuration.MaterialSelection
        }
    End Function

    Public Function CreateDraftGenerationOptions(configuration As ProductConfiguration) As DraftGenerationOptions
        Return New DraftGenerationOptions() With {
            .Prefix = configuration.Unit.Prefix,
            .Scale = configuration.Unit.Scale,
            .MaterialSelection = configuration.MaterialSelection
        }
    End Function

    Public Function CreateProjectCodingOptions(configuration As ProductConfiguration) As ProjectCodingOptions
        Return New ProjectCodingOptions() With {
            .ProjectName = configuration.ProjectIdentity.ProjectName,
            .Revision = configuration.ProjectIdentity.Revision,
            .DocumentNumber = configuration.ProjectIdentity.DocumentNumber
        }
    End Function
End Class
