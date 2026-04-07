Public Class SolidEdgeWorkflowService

    Public Function ExecuteWithAssembly(asmFilePath As String,
                                        applicationOptions As SolidEdgeApplicationOptions,
                                        displayAlerts As Boolean,
                                        work As Func(Of SolidEdgeFramework.Application, SolidEdgeAssembly.AssemblyDocument, Boolean)) As Boolean

        Dim session As SolidEdgeSessionContext = Nothing
        Dim seApplication As SolidEdgeFramework.Application = Nothing
        Dim seDocuments As SolidEdgeFramework.Documents = Nothing
        Dim seAssembly As SolidEdgeAssembly.AssemblyDocument = Nothing

        Try
            session = SolidEdgeSessionHelpers.OpenApplication(applicationOptions.MakeVisible)
            seApplication = session.Application
            seApplication.DisplayAlerts = displayAlerts

            seDocuments = seApplication.Documents
            seAssembly = seDocuments.Open(asmFilePath)

            Return work(seApplication, seAssembly)
        Finally
            SolidEdgeSessionHelpers.ReleaseCOMReference(seAssembly)
            SolidEdgeSessionHelpers.ReleaseCOMReference(seDocuments)
            SolidEdgeSessionHelpers.CloseApplication(session, True)
        End Try
    End Function

    Public Function ExecuteWithApplication(applicationOptions As SolidEdgeApplicationOptions,
                                           displayAlerts As Boolean,
                                           work As Func(Of SolidEdgeFramework.Application, Boolean)) As Boolean

        Dim session As SolidEdgeSessionContext = Nothing
        Dim seApplication As SolidEdgeFramework.Application = Nothing

        Try
            session = SolidEdgeSessionHelpers.OpenApplication(applicationOptions.MakeVisible)
            seApplication = session.Application
            seApplication.DisplayAlerts = displayAlerts

            Return work(seApplication)
        Finally
            SolidEdgeSessionHelpers.CloseApplication(session, True)
        End Try
    End Function
End Class
