Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Media.Imaging

Namespace Controls
    Public Class ViewMenu
        Inherits ContextMenu

        Public Shared ReadOnly FolderProperty As DependencyProperty = DependencyProperty.Register("Folder", GetType(Folder), GetType(ViewMenu), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))

        Private _isCheckingInternally As Boolean

        Protected Overrides Sub OnOpened(e As RoutedEventArgs)
            Me.Items.Clear()

            Me.AddItems(Me.Items)

            MyBase.OnOpened(e)
        End Sub

        Public Sub AddItems(menu As ItemCollection)
            For Each item In Shell.FolderViews
                Dim viewSubMenuItem As MenuItem = New MenuItem() With {
                    .Header = item.Key,
                    .Icon = New Image() With {.Source = New BitmapImage(New Uri(item.Value.Item1, UriKind.Absolute))},
                    .Tag = "View:" & item.Key,
                    .IsCheckable = True,
                    .IsChecked = item.Key = Me.Folder.View
                }
                AddHandler viewSubMenuItem.Checked,
                    Sub(s2 As Object, e2 As EventArgs)
                        If Not _isCheckingInternally Then
                            Me.Folder.View = item.Key
                        End If
                    End Sub
                AddHandler viewSubMenuItem.Unchecked,
                    Sub(s2 As Object, e2 As EventArgs)
                        If Not _isCheckingInternally Then
                            _isCheckingInternally = True
                            viewSubMenuItem.IsChecked = True
                            _isCheckingInternally = False
                        End If
                    End Sub
                menu.Add(viewSubMenuItem)
            Next
        End Sub

        Public Property Folder As Folder
            Get
                Return GetValue(FolderProperty)
            End Get
            Set(value As Folder)
                SetValue(FolderProperty, value)
            End Set
        End Property
    End Class
End Namespace