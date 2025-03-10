Imports System.Collections.ObjectModel
Imports System.Windows
Imports System.Windows.Controls
Imports Laila.Shell.Controls.TabControl
Imports Laila.Shell.Helpers

Namespace Controls
    Public Class TabItemCloseButton
        Inherits Button

        Public Sub New()
            AddHandler Me.Click,
                Sub(s As Object, e As EventArgs)
                    Dim tabItem As TabItem = UIHelper.GetParentOfType(Of TabItem)(Me)
                    Dim tabControl As TabControl = UIHelper.GetParentOfType(Of TabControl)(tabItem)
                    Dim items As ObservableCollection(Of Object) = tabControl.ItemsSource
                    If tabControl.Items.Count > 1 Then
                        items.Remove(tabItem.DataContext)
                    Else
                        Window.GetWindow(Me).Close()
                    End If
                End Sub
        End Sub
    End Class
End Namespace