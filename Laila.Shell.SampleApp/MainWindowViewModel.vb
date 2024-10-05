Public Class MainWindowViewModel
    Inherits NotifyPropertyChangedBase

    Private _view As MainWindow
    Private _folder As Folder

    Public Sub New(view As MainWindow)
        _view = view
    End Sub

    Public Property Folder As Folder
        Get
            Return _folder
        End Get
        Set(value As Folder)
            SetValue(_folder, value)
        End Set
    End Property
End Class
