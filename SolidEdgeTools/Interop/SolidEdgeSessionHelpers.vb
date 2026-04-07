Imports System.Runtime.InteropServices

Public Module SolidEdgeSessionHelpers

    Public Function OpenApplication(makeVisible As Boolean) As SolidEdgeFramework.Application
        Dim seApplication As SolidEdgeFramework.Application = Nothing

        SolidEdgeCommunity.OleMessageFilter.Register()
        seApplication = SolidEdgeCommunity.SolidEdgeUtils.Connect(True, True)
        seApplication.Visible = makeVisible

        Return seApplication
    End Function

    Public Sub CloseApplication(ByRef seApplication As SolidEdgeFramework.Application, quit As Boolean)
        If seApplication Is Nothing Then
            Return
        End If

        If quit Then
            seApplication.Quit()
        End If

        SolidEdgeCommunity.OleMessageFilter.Unregister()
    End Sub

    Public Sub ReleaseCOMReference(ByRef comObject As Object)
        If comObject Is Nothing Then
            Return
        End If

        Marshal.ReleaseComObject(comObject)
        comObject = Nothing
    End Sub
End Module
