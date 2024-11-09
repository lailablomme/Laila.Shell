Imports System.Collections.ObjectModel
Imports System.ComponentModel

Public Class MainWindowViewModel
    Inherits NotifyPropertyChangedBase

    Private _view As MainWindow

    Public Sub New(view As MainWindow)
        _view = view
    End Sub
End Class
