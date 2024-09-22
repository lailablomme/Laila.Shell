Public Class MainWindowViewModel
    Inherits NotifyPropertyChangedBase

    Private _view As MainWindow
    Private _folderName As String
    Private _logicalParent As Folder

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

    Public Property LogicalParent As Folder
        Get
            Return _logicalParent
        End Get
        Set(value As Folder)
            SetValue(_logicalParent, value)
        End Set
    End Property
End Class
