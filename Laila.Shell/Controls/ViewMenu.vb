﻿Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Input
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
            For Each item In Me.Folder.Views
                Dim viewSubMenuItem As MenuItem = New MenuItem() With {
                    .Header = item.Title,
                    .Icon = New Image() With {.Source = New BitmapImage(item.IconUri)},
                    .Tag = "View:" & item.Guid.ToString(),
                    .IsCheckable = True,
                    .IsChecked = item.Guid = Me.Folder.ActiveView
                }
                AddHandler viewSubMenuItem.Checked,
                    Sub(s2 As Object, e2 As EventArgs)
                        If Not _isCheckingInternally Then
                            Using Shell.OverrideCursor(Cursors.Wait)
                                Me.Folder.ActiveView = item.Guid
                            End Using
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

            If Me.MenuStyle = ViewMenuStyle.RightClickMenu AndAlso Not String.IsNullOrWhiteSpace(Me.Folder.ItemsGroupByPropertyName) Then
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
                Dim compactModeMenuItem As MenuItem = New MenuItem() With {
                    .Header = "Compact mode",
                    .Icon = New Image() With {.Source = New ImageSourceConverter().ConvertFromInvariantString("pack://application:,,,/Laila.Shell;component/Images/compactmode16.png")},
                    .IsCheckable = True,
                    .IsChecked = Shell.Settings.IsCompactMode
                }
                AddHandler compactModeMenuItem.Click,
                    Sub(s As Object, e As EventArgs)
                        Shell.Settings.IsCompactMode = compactModeMenuItem.IsChecked
                    End Sub
                viewSubMenuItem.Items.Add(compactModeMenuItem)
                viewSubMenuItem.Items.Add(New Separator())
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
                Dim hiddenItemsMenuItem As MenuItem = New MenuItem() With {
                    .Header = "Hidden items",
                    .Icon = New Image() With {.Source = New ImageSourceConverter().ConvertFromInvariantString("pack://application:,,,/Laila.Shell;component/Images/eye16.png")},
                    .IsCheckable = True,
                    .IsChecked = Shell.Settings.DoShowHiddenFilesAndFolders
                }
                AddHandler hiddenItemsMenuItem.Click,
                    Sub(s As Object, e As EventArgs)
                        Shell.Settings.DoShowHiddenFilesAndFolders = hiddenItemsMenuItem.IsChecked
                    End Sub
                viewSubMenuItem.Items.Add(hiddenItemsMenuItem)
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