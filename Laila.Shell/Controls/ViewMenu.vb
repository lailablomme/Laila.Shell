Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Media
Imports System.Windows.Media.Imaging

Namespace Controls
    Public Class ViewMenu
        Inherits ContextMenu

        Public Shared ReadOnly FolderProperty As DependencyProperty = DependencyProperty.Register("Folder", GetType(Folder), GetType(ViewMenu), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly MenuStyleProperty As DependencyProperty = DependencyProperty.Register("MenuStyle", GetType(ViewMenuStyle), GetType(ViewMenu), New FrameworkPropertyMetadata(ViewMenuStyle.Toolbar, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))

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

            If Me.MenuStyle = ViewMenuStyle.RightClickMenu Then
                menu.Add(New Separator())
                Dim expandAllGroupsMenuItem As MenuItem = New MenuItem() With {
                    .Header = "Expand all groups",
                    .Icon = New Image() With {.Source = New ImageSourceConverter().ConvertFromInvariantString("pack://application:,,,/Laila.Shell;component/Images/expandall16.png")}
                }
                AddHandler expandAllGroupsMenuItem.Click,
                    Sub(s As Object, e As EventArgs)
                        Me.Folder.TriggerExpandAllGroups()
                    End Sub
                menu.Add(expandAllGroupsMenuItem)
                Dim collapseAllGroupsMenuItem As MenuItem = New MenuItem() With {
                    .Header = "Collapse all groups",
                    .Icon = New Image() With {.Source = New ImageSourceConverter().ConvertFromInvariantString("pack://application:,,,/Laila.Shell;component/Images/collapseall16.png")}
                }
                AddHandler collapseAllGroupsMenuItem.Click,
                    Sub(s As Object, e As EventArgs)
                        Me.Folder.TriggerCollapseAllGroups()
                    End Sub
                menu.Add(collapseAllGroupsMenuItem)
            End If

            If Me.MenuStyle = ViewMenuStyle.Toolbar Then
                menu.Add(New Separator())
                Dim viewSubMenuItem As MenuItem = New MenuItem() With {
                    .Header = "View"
                }
                menu.Add(viewSubMenuItem)
                Dim checkBoxesForItemsMenuItem As MenuItem = New MenuItem() With {
                    .Header = "Checkboxes for items",
                    .Icon = New Image() With {.Source = New ImageSourceConverter().ConvertFromInvariantString("pack://application:,,,/Laila.Shell;component/Images/filecheck16.png")},
                    .IsCheckable = True,
                    .IsChecked = Shell.Settings.DoShowCheckBoxesToSelect
                }
                AddHandler checkBoxesForItemsMenuItem.Click,
                    Sub(s As Object, e As EventArgs)
                        Shell.Settings.DoShowCheckBoxesToSelect = checkBoxesForItemsMenuItem.IsChecked
                    End Sub
                viewSubMenuItem.Items.Add(checkBoxesForItemsMenuItem)
                Dim fileNameExtensionsMenuItem As MenuItem = New MenuItem() With {
                    .Header = "Filename extensions",
                    .Icon = New Image() With {.Source = New ImageSourceConverter().ConvertFromInvariantString("pack://application:,,,/Laila.Shell;component/Images/fileext16.png")},
                    .IsCheckable = True,
                    .IsChecked = Not Shell.Settings.DoHideKnownFileExtensions
                }
                AddHandler fileNameExtensionsMenuItem.Click,
                    Sub(s As Object, e As EventArgs)
                        Shell.Settings.DoHideKnownFileExtensions = Not fileNameExtensionsMenuItem.IsChecked
                    End Sub
                viewSubMenuItem.Items.Add(fileNameExtensionsMenuItem)
            End If
        End Sub

        Public Property Folder As Folder
            Get
                Return GetValue(FolderProperty)
            End Get
            Set(value As Folder)
                SetValue(FolderProperty, value)
            End Set
        End Property

        Public Property MenuStyle As ViewMenuStyle
            Get
                Return GetValue(MenuStyleProperty)
            End Get
            Set(value As ViewMenuStyle)
                SetValue(MenuStyleProperty, value)
            End Set
        End Property
    End Class
End Namespace