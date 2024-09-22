Imports System.Collections.ObjectModel
Imports System.ComponentModel

Class MainWindow
    Private _model As MainWindowViewModel
    Public Sub New()
        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        _model = New MainWindowViewModel(Me)
        Me.DataContext = _model
    End Sub
End Class
