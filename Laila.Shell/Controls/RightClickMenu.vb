Imports System.Windows.Controls
Imports System.Windows.Media.Imaging

Namespace Controls
    Public Class RightClickMenu
        Inherits BaseMenu

        Protected Overrides Sub AddItems()
            Dim osver As Version = Environment.OSVersion.Version
            Dim isWindows11 As Boolean = osver.Major = 10 AndAlso osver.Minor = 0 AndAlso osver.Build >= 22000

            ' add menu items
            Dim menuItems As List(Of Control) = getMenuItems()
            Dim lastMenuItem As Control
            For Each item In menuItems
                Dim verb As String = If(Not item.Tag Is Nothing, CType(item.Tag, Tuple(Of Integer, String)).Item2, Nothing)
                Select Case verb
                    Case "copy", "cut", "paste", "delete", "pintohome", "rename"
                        ' don't add these
                    Case Else
                        Dim isNotDoubleSeparator As Boolean = Not (TypeOf item Is Separator AndAlso
                                (Not lastMenuItem Is Nothing AndAlso TypeOf lastMenuItem Is Separator))
                        Dim isNotInitialSeparator As Boolean = Not (TypeOf item Is Separator AndAlso Me.Items.Count = 0)
                        Dim isNotDoubleOneDriveItem As Boolean = verb Is Nothing OrElse
                                Not (isWindows11 AndAlso
                                    (verb.StartsWith("{5250E46F-BB09-D602-5891-F476DC89B70") _
                                     OrElse verb.StartsWith("{1FA0E654-C9F2-4A1F-9800-B9A75D744B0") _
                                     OrElse verb = "MakeAvailableOffline" _
                                     OrElse verb = "MakeAvailableOnline"))
                        If isNotDoubleSeparator AndAlso isNotInitialSeparator AndAlso isNotDoubleOneDriveItem Then
                            Me.Items.Add(item)
                            lastMenuItem = item
                        End If
                End Select
            Next

            ' add buttons
            Dim hasPaste As Boolean =
                Not Me.IsDefaultOnly _
                AndAlso (Me.SelectedItems Is Nothing OrElse Me.SelectedItems.Count = 0) _
                AndAlso Clipboard.CanPaste(Me.Folder)

            Dim menuItem As MenuItem = menuItems.FirstOrDefault(Function(i) If(Not i.Tag Is Nothing, CType(i.Tag, Tuple(Of Integer, String)).Item2, Nothing) = "cut")
            If Not menuItem Is Nothing Then Me.Buttons.Add(MakeButton(menuItem.Tag, menuItem.Header.ToString().Replace("_", "")))
            menuItem = menuItems.FirstOrDefault(Function(i) If(Not i.Tag Is Nothing, CType(i.Tag, Tuple(Of Integer, String)).Item2, Nothing) = "copy")
            If Not menuItem Is Nothing Then Me.Buttons.Add(MakeButton(menuItem.Tag, menuItem.Header.ToString().Replace("_", "")))
            menuItem = menuItems.FirstOrDefault(Function(i) If(Not i.Tag Is Nothing, CType(i.Tag, Tuple(Of Integer, String)).Item2, Nothing) = "paste")
            If Not menuItem Is Nothing Then Me.Buttons.Add(MakeButton(menuItem.Tag, menuItem.Header.ToString().Replace("_", ""))) _
                Else If hasPaste Then Me.Buttons.Add(MakeButton(New Tuple(Of Integer, String)(-1, "paste"), "Paste"))
            menuItem = menuItems.FirstOrDefault(Function(i) If(Not i.Tag Is Nothing, CType(i.Tag, Tuple(Of Integer, String)).Item2, Nothing) = "rename")
            If Not Me.SelectedItems Is Nothing AndAlso Me.SelectedItems.Count = 1 AndAlso Me.SelectedItems.All(Function(i) i.Attributes.HasFlag(SFGAO.CANRENAME)) Then _
                Me.Buttons.Add(MakeButton(New Tuple(Of Integer, String)(-1, "rename"), "Rename"))
            menuItem = menuItems.FirstOrDefault(Function(i) If(Not i.Tag Is Nothing, CType(i.Tag, Tuple(Of Integer, String)).Item2, Nothing) = "delete")
            If Not menuItem Is Nothing Then Me.Buttons.Add(MakeButton(menuItem.Tag, menuItem.Header.ToString().Replace("_", "")))
            If Not Me.SelectedItems Is Nothing AndAlso Me.SelectedItems.Count = 1 Then
                Dim test As Item = Item.FromParsingName(Me.SelectedItems(0).FullPath, Nothing)
                If Not test Is Nothing Then ' this won't work for all items 
                    test.Dispose()
                    Dim isPinned As Boolean = PinnedItems.GetIsPinned(Me.SelectedItems(0).FullPath)
                    Me.Buttons.Add(MakeToggleButton(New Tuple(Of Integer, String)(-1, "laila.shell.(un)pin"),
                                                        If(isPinned, "Unpin item", "Pin item"), isPinned))
                End If
            End If

            ' add view, sort and group by menus
            If Me.SelectedItems Is Nothing OrElse Me.SelectedItems.Count = 0 Then
                Dim viewMenu As ViewMenu = New ViewMenu() With {.Folder = Me.Folder}
                Dim viewMenuItem As MenuItem = New MenuItem() With {
                    .Header = "View",
                    .Icon = New Image() With {.Source = New BitmapImage(New Uri("pack://application:,,,/Laila.Shell;component/Images/view16.png", UriKind.Absolute))}
                }
                viewMenu.AddItems(viewMenuItem.Items)

                Dim sortMenu As SortMenu = New SortMenu() With {.Folder = Me.Folder}
                Dim sortMenuItem As MenuItem = New MenuItem() With {
                    .Header = "Sort",
                    .Icon = New Image() With {.Source = New BitmapImage(New Uri("pack://application:,,,/Laila.Shell;component/Images/sort16.png", UriKind.Absolute))}
                }
                sortMenu.AddSortItems(sortMenuItem.Items)

                Dim groupByMenuItem As MenuItem = New MenuItem() With {
                    .Header = "Group by"
                }
                sortMenu.AddGroupByItems(groupByMenuItem.Items)

                Me.Items.Insert(0, viewMenuItem)
                Me.Items.Insert(1, sortMenuItem)
                Me.Items.Insert(2, groupByMenuItem)
                If Me.Items.Count > 3 Then
                    Me.Items.Insert(3, New Separator())
                End If
            End If
        End Sub
    End Class
End Namespace