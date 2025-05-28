Imports Laila.Shell.Interop.Items

Public Class Home_RecentItemLink
    Inherits Link

    Public Sub New(shellItem2 As IShellItem2, logicalParent As Folder, doKeepAlive As Boolean, doHookUpdates As Boolean, threadId As Integer?, Optional pidl As Pidl = Nothing)
        MyBase.New(shellItem2, logicalParent, doKeepAlive, doHookUpdates, threadId, pidl)
    End Sub

    Public Overrides ReadOnly Property FullPath As String
        Get
            Return MyBase.TargetFullPath
        End Get
    End Property
End Class
