Imports System.Collections.ObjectModel
Imports System.Windows

Namespace DebugTools
    Public Class DebugWindow
        Public Sub New()
            ' This call is required by the designer.
            InitializeComponent()

            ' Add any initialization after the InitializeComponent() call.
            Me.DataContext = Me
            listView.ItemsSource = Shell.ItemsCache.ToList()
        End Sub

        Private Sub listView_PreviewMouseDoubleClick(sender As Object, e As Input.MouseButtonEventArgs)
            listView.ItemsSource = Nothing
            listView.ItemsSource = Shell.ItemsCache.ToList()
        End Sub
    End Class
End Namespace