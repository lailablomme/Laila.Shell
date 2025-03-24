﻿Imports System.Windows.Controls
Imports System.Windows.Media.Imaging
Imports Laila.Shell.Interop.Items

Namespace Controls
    Public Class RightClickMenu
        Inherits BaseMenu

        Protected Overrides Async Function AddItems() As Task
            ' add right click menu items
            If Not Me.SelectedItems Is Nothing AndAlso Me.SelectedItems.Count = 1 AndAlso TypeOf Me.SelectedItems(0) Is Folder Then
                Await CType(Me.SelectedItems(0), Folder).AddRightClickMenuItems(Me)
            Else
                Await Me.Folder.AddRightClickMenuItems(Me)
            End If

            ' add view, sort and group by menus
            If Me.SelectedItems Is Nothing OrElse Me.SelectedItems.Count = 0 Then
                Dim insertIndex As Integer = 0

                Dim viewMenu As ViewMenu = New ViewMenu() With {.Folder = Me.Folder, .MenuStyle = ViewMenuStyle.RightClickMenu}
                Dim viewMenuItem As MenuItem = New MenuItem() With {
                    .Header = "View",
                    .Icon = New Image() With {.Source = New BitmapImage(New Uri("pack://application:,,,/Laila.Shell;component/Images/view16.png", UriKind.Absolute))}
                }
                viewMenu.AddItems(viewMenuItem.Items)
                Me.Items.Insert(0, viewMenuItem)
                insertIndex += 1

                If Not Me.Folder.ShellItem2 Is Nothing Then
                    Dim sortMenu As SortMenu = New SortMenu() With {.Folder = Me.Folder}

                    If Me.Folder.CanSort Then
                        Dim sortMenuItem As MenuItem = New MenuItem() With {
                            .Header = "Sort",
                            .Icon = New Image() With {.Source = New BitmapImage(New Uri("pack://application:,,,/Laila.Shell;component/Images/sort16.png", UriKind.Absolute))}
                        }
                        sortMenu.AddSortItems(sortMenuItem.Items)
                        Me.Items.Insert(insertIndex, sortMenuItem)
                        insertIndex += 1
                    End If

                    If Me.Folder.CanGroupBy Then
                        Dim groupByMenuItem As MenuItem = New MenuItem() With {
                            .Header = "Group by"
                        }
                        sortMenu.AddGroupByItems(groupByMenuItem.Items)
                        Me.Items.Insert(insertIndex, groupByMenuItem)
                        insertIndex += 1
                    End If
                End If

                If insertIndex > 0 Then
                    Me.Items.Insert(insertIndex, New Separator())
                End If
            End If
        End Function
    End Class
End Namespace