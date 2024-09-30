Imports System.Collections.ObjectModel
Imports System.ComponentModel
Imports System.Runtime.InteropServices
Imports System.Runtime.InteropServices.ComTypes
Imports System.Windows.Forms
Imports System.Windows.Interop

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
