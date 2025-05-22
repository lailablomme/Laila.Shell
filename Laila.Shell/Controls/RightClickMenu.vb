Imports System.Windows.Controls
Imports System.Windows.Media.Imaging
Imports Laila.Shell.Interop.Items

Namespace Controls
    Public Class RightClickMenu
        Inherits BaseContextMenu

        Protected Overrides Function AddItems() As Task
            ' add right click menu items
            If Not Me.SelectedItems Is Nothing AndAlso Me.SelectedItems.Count = 1 AndAlso TypeOf Me.SelectedItems(0) Is Folder Then
                CType(Me.SelectedItems(0), Folder).AddRightClickMenuItems(Me)
            Else
                Me.Folder.AddRightClickMenuItems(Me)
            End If

            Dim hasItems As Boolean = Me.Items.Count > 0

            ' add view, sort and group by menus
            If Me.SelectedItems Is Nothing OrElse Me.SelectedItems.Count = 0 Then
                Dim insertIndex As Integer = 0

                Dim viewMenu As ViewMenu = New ViewMenu() With {.Folder = Me.Folder, .MenuStyle = ViewMenuStyle.RightClickMenu}
                Dim viewMenuItem As MenuItem = New MenuItem() With {
                    .Header = My.Resources.Menu_View,
                    .Icon = New Image() With {.Source = System.Windows.Application.Current.TryFindResource($"{ResourcePrefix}ViewIcon")}
                }
                viewMenu.AddItems(viewMenuItem.Items, Me.ResourcePrefix)
                Me.Items.Insert(0, viewMenuItem)
                insertIndex += 1

                If Not Me.Folder.ShellItem2 Is Nothing Then
                    Dim sortMenu As SortMenu = New SortMenu() With {.Folder = Me.Folder}

                    If Me.Folder.CanSort Then
                        Dim sortMenuItem As MenuItem = New MenuItem() With {
                            .Header = My.Resources.Menu_Sort,
                            .Icon = New Image() With {.Source = System.Windows.Application.Current.TryFindResource($"{ResourcePrefix}SortIcon")}
                        }
                        sortMenu.AddSortItems(sortMenuItem.Items, Me.ResourcePrefix)
                        Me.Items.Insert(insertIndex, sortMenuItem)
                        insertIndex += 1
                    End If

                    If Me.Folder.CanGroupBy Then
                        Dim groupByMenuItem As MenuItem = New MenuItem() With {
                            .Header = My.Resources.Menu_GroupBy
                        }
                        sortMenu.AddGroupByItems(groupByMenuItem.Items)
                        Me.Items.Insert(insertIndex, groupByMenuItem)
                        insertIndex += 1
                    End If
                End If

                If insertIndex > 0 AndAlso hasItems Then
                    Me.Items.Insert(insertIndex, New Separator())
                End If
            End If

            Return Task.CompletedTask
        End Function
    End Class
End Namespace