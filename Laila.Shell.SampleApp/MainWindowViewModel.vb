Public Class MainWindowViewModel
    Inherits NotifyPropertyChangedBase

    Private _view As MainWindow
    Private _folderName As String

    Public Sub New(view As MainWindow)
        _view = view
    End Sub

    Public Property FolderName As String
        Get
            Return _folderName
        End Get
        Set(value As String)
            SetValue(_folderName, value)
        End Set
    End Property
End Class
