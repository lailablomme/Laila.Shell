Namespace Events
    Public Class FolderEventArgs
        Inherits EventArgs

        Public Property Folder As Folder

        Public Sub New(folder As Folder)
            Me.Folder = folder
        End Sub
    End Class
End Namespace