Imports System.Runtime.InteropServices

Public Class SolidEdgeSessionContext
    Public Property Application As SolidEdgeFramework.Application
    Public Property StartedByTool As Boolean
End Class

Public Module SolidEdgeSessionHelpers

    Public Function OpenApplication(makeVisible As Boolean) As SolidEdgeSessionContext
        Dim session As New SolidEdgeSessionContext()

        SolidEdgeCommunity.OleMessageFilter.Register()
        Try
            session.Application = DirectCast(Marshal.GetActiveObject("SolidEdge.Application"), SolidEdgeFramework.Application)
            session.StartedByTool = False
        Catch ex As Exception
            session.Application = SolidEdgeCommunity.SolidEdgeUtils.Connect(True, True)
            session.StartedByTool = True
        End Try

        session.Application.Visible = makeVisible

        Return session
    End Function

    Public Sub CloseApplication(ByRef session As SolidEdgeSessionContext, quit As Boolean)
        If session Is Nothing OrElse session.Application Is Nothing Then
            Return
        End If

        If quit AndAlso session.StartedByTool Then
            session.Application.Quit()
        End If

        ReleaseCOMReference(session.Application)
        session = Nothing

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
