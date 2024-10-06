Namespace Events
    Public Class RequestSetSelectedFolderEventArgs
        Public Property RequestedFolder As Folder
        Public Property Callback As Action(Of Folder)
    End Class
End Namespace