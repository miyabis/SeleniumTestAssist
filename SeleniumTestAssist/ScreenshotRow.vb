
Imports System.IO

Public Class ScreenshotRow

    Public Property TestClassName As String
    Public Property TestMethodName As String
    Public Property FullPath As String
    Public Property Title As String
    Public Property Count As Integer
    Public Property Note As String
    Public Property WindowTitle As String

    Public ReadOnly Property Filename As String
        Get
            Return New FileInfo(Me.FullPath).Name
        End Get
    End Property

End Class
