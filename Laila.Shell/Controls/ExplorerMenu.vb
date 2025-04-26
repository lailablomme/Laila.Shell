Imports System.IO
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Windows.Controls
Imports System.Windows.Controls.Primitives
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports System.Windows.Input
Imports System.Windows.Media.Imaging
Imports Laila.Shell.Events
Imports Laila.Shell.Helpers
Imports Laila.Shell.Interop
Imports Laila.Shell.Interop.ContextMenu
Imports Laila.Shell.Interop.DragDrop
Imports Laila.Shell.Interop.Folders
Imports Laila.Shell.Interop.Functions
Imports Laila.Shell.Interop.Items
Imports Laila.Shell.Interop.Properties
Imports Laila.Shell.Interop.Windows
Imports Laila.Shell.WinRT.Interface
Imports Microsoft.VisualBasic.Devices
Imports Microsoft.Win32
Imports SHDocVw

Namespace Controls
    Public Class ExplorerMenu
        Inherits BaseMenu

        Private Const COMMANDSTORE_PATH As String = "Software\Microsoft\Windows\CurrentVersion\Explorer\CommandStore\shell"

        Private _menuItems As List(Of MenuItemData) = New List(Of MenuItemData)()
        Private _helper As Helpers.ComHelper = New ComHelper()

        Protected Overrides Sub Make(folder As Folder, items As IEnumerable(Of Item), isDefaultOnly As Boolean)
            _thread.Run(
                Sub()
                    Dim assembly As Assembly = Assembly.LoadFrom("Laila.Shell.WinRT.dll")
                    Dim type As Type = assembly.GetType("Laila.Shell.WinRT.ExplorerMenuHelper")
                    Dim instance As Object = Activator.CreateInstance(type)

                    ' add cloud menu item
                    addCloudMenuItem(type, instance, folder, items)

                    ' add file explorer context menu items
                    addFileExplorerContextMenuItems(type, instance, folder, items)

                    ' add show more options command
                    If _menuItems.Count > 0 Then
                        _menuItems.Add(New MenuItemData() With {.Header = "-----"})
                    End If
                    Dim showMoreOptionsMenuItem As New MenuItemData() With {
                        .Header = "Show more options...",
                        .IsEnabled = True
                    }
                    Dim action As Action =
                        Sub()
                            UIHelper.OnUIThread(
                                Sub()
                                    Dim rightClickMenu As RightClickMenu = New RightClickMenu()  ' 
                                    rightClickMenu.Folder = folder
                                    rightClickMenu.SelectedItems = items
                                    rightClickMenu.IsOpen = True
                                End Sub)
                        End Sub
                    showMoreOptionsMenuItem.Tag = New Tuple(Of Integer, String, Object)(-1, Guid.NewGuid().ToString(), action)
                    _menuItems.Add(showMoreOptionsMenuItem)
                End Sub)
        End Sub

        Private Sub addCloudMenuItem(type As Type, instance As Object, folder As Folder, items As IEnumerable(Of Item))
            ' get cloud verbs
            Dim methodInfo As MethodInfo = type.GetMethod("GetCloudVerbs")
            Dim t As Task(Of Tuple(Of String, String, List(Of ExplorerCommandVerbInfo))) =
                methodInfo.Invoke(instance, {folder.FullPath})
            Dim result As Tuple(Of String, String, List(Of ExplorerCommandVerbInfo)) = t.Result

            If Not result Is Nothing Then
                ' add cloud menu
                Dim cloudMenuItem As New MenuItemData() With {
                    .Header = result.Item1,
                    .IsEnabled = True,
                    .Items = New List(Of MenuItemData)()
                }
                If Not String.IsNullOrWhiteSpace(result.Item2) Then
                    cloudMenuItem.Icon = New BitmapImage(New Uri(result.Item2))
                    cloudMenuItem.Icon.Freeze()
                End If

                ' add commands
                Dim array As IShellItemArray = Nothing
                Dim pidls As List(Of Pidl) = items.Select(Function(i) i.Pidl.Clone()).ToList()
                Functions.SHCreateShellItemArrayFromIDLists(pidls.Count, pidls.Select(Function(p) p.AbsolutePIDL).ToArray(), array)

                For Each item In result.Item3
                    Dim command As IExplorerCommand = _helper.MakeComObject(item.ComServerPath, item.ClsId, GetType(IExplorerCommand).GUID)
                    Dim flags As EXPCMDFLAGS
                    Dim menuItem As MenuItemData = getMenuItem(command, array, flags)
                    If Not menuItem Is Nothing Then
                        If flags.HasFlag(EXPCMDFLAGS.ECF_SEPARATORBEFORE) Then
                            cloudMenuItem.Items.Add(New MenuItemData() With {.Header = "-----"})
                        End If
                        cloudMenuItem.Items.Add(menuItem)
                        If flags.HasFlag(EXPCMDFLAGS.ECF_SEPARATORAFTER) Then
                            cloudMenuItem.Items.Add(New MenuItemData() With {.Header = "-----"})
                        End If
                    End If
                Next

                If cloudMenuItem.Items.Count > 0 Then
                    If _menuItems.Count > 0 Then
                        _menuItems.Add(New MenuItemData() With {.Header = "-----"})
                    End If
                    _menuItems.Add(cloudMenuItem)
                End If
            End If
        End Sub

        Private Sub addFileExplorerContextMenuItems(type As Type, instance As Object, folder As Folder, items As IEnumerable(Of Item))
            ' get file explorer verbs
            Dim methodInfo As MethodInfo = type.GetMethod("GetFileExplorerVerbs")
            Dim t As Task(Of List(Of ExplorerCommandVerbInfo)) = methodInfo.Invoke(instance, {})
            Dim result As List(Of ExplorerCommandVerbInfo) = t.Result

            If Not result Is Nothing Then
                ' filter commands
                Dim isBackground As Boolean = items Is Nothing OrElse items.Count = 0
                Dim filtered As List(Of ExplorerCommandVerbInfo) = New List(Of ExplorerCommandVerbInfo)()
                For Each item In result
                    If (item.ItemType = "Directory\Background" AndAlso isBackground) _
                        OrElse (item.ItemType = "*" AndAlso items.ToList().Exists(Function(i) Not TypeOf i Is Folder)) _
                        OrElse (item.ItemType = "Directory" AndAlso items.ToList().Exists(Function(i) TypeOf i Is Folder)) _
                        OrElse (item.ItemType = "Drive" AndAlso items.ToList().Exists(Function(i) i.IsDrive)) _
                        OrElse (items.ToList().Exists(Function(i) item.ItemType.Equals(IO.Path.GetExtension(i.FullPath).ToLower()))) Then
                        If Not filtered.Exists(Function(f) f.ClsId = item.ClsId AndAlso f.ComServerPath = item.ComServerPath) Then
                            filtered.Add(item)
                        End If
                    End If
                Next

                ' get item array
                Dim array As IShellItemArray = Nothing
                Dim pidls As List(Of Pidl) = items.Select(Function(i) i.Pidl.Clone()).ToList()
                Functions.SHCreateShellItemArrayFromIDLists(pidls.Count, pidls.Select(Function(p) p.AbsolutePIDL).ToArray(), array)

                ' group commands
                Dim tempRoot As List(Of MenuItemData) = New List(Of MenuItemData)()
                For Each item In filtered.Select(Function(f) $"{f.ApplicationName}\{f.PackageId}\{f.ApplicationId}").Distinct().OrderBy(Function(f) f)
                    Dim thisAppCommands As List(Of ExplorerCommandVerbInfo) = filtered.Where(Function(f) $"{f.ApplicationName}\{f.PackageId}\{f.ApplicationId}" = item).ToList()
                    Dim addTo As List(Of MenuItemData) = Nothing
                    Dim appMenuItem As MenuItemData = Nothing

                    If thisAppCommands.Count = 1 Then
                        addTo = tempRoot
                    Else
                        appMenuItem = New MenuItemData() With {
                            .ApplicationName = thisAppCommands(0).ApplicationName,
                            .Header = thisAppCommands(0).ApplicationName,
                            .Icon = New BitmapImage(New Uri(thisAppCommands(0).ApplicationIconPath)),
                            .IsEnabled = True,
                            .Items = New List(Of MenuItemData)()
                        }
                        appMenuItem.Icon.Freeze()
                        addTo = New List(Of MenuItemData)()
                    End If

                    For Each commandItem In thisAppCommands
                        Dim command As IExplorerCommand = _helper.MakeComObject(commandItem.ComServerPath, commandItem.ClsId, GetType(IExplorerCommand).GUID)
                        Dim flags As EXPCMDFLAGS
                        Dim menuItem As MenuItemData = getMenuItem(command, array, flags)
                        If Not menuItem Is Nothing Then
                            If Not menuItem.Items Is Nothing AndAlso menuItem.Items.Count = 1 _
                                AndAlso menuItem.Items(0).Header?.Equals(menuItem.Header) Then
                                ' how is this possible that i have two apps who get this wrong?!
                                menuItem = menuItem.Items(0)
                            End If
                            menuItem.ApplicationName = commandItem.ApplicationName
                            addTo.Add(menuItem)
                        End If
                    Next

                    If Not addTo.Equals(tempRoot) AndAlso addTo.Count > 0 Then
                        appMenuItem.Items.AddRange(addTo.OrderBy(Function(i) i.Header))
                        tempRoot.Add(appMenuItem)
                    End If
                Next

                If tempRoot.Count > 0 Then
                    If _menuItems.Count > 0 Then
                        _menuItems.Add(New MenuItemData() With {.Header = "-----"})
                    End If
                    _menuItems.AddRange(tempRoot.OrderBy(Function(i) $"{i.ApplicationName}\{i.Header}"))
                End If
            End If
        End Sub

        Private Function getMenuItem(command As IExplorerCommand, array As IShellItemArray, ByRef flags As EXPCMDFLAGS) As MenuItemData
            Dim state As EXPCMDSTATE = 0
            Dim h As HRESULT = command.GetState(array, True, state)
            If state <> EXPCMDSTATE.ECS_HIDDEN Then
                Dim menuItem As MenuItemData = New MenuItemData() With {.IsEnabled = state = EXPCMDSTATE.ECS_ENABLED}
                command.GetFlags(flags)
                If flags.HasFlag(EXPCMDFLAGS.ECF_ISSEPARATOR) Then
                    menuItem.Header = "-----"
                Else
                    command.GetTitle(array, menuItem.Header)
                    menuItem.Header = menuItem.Header.Replace("&&", "{{AMP}}").Replace("&", "_").Replace("{{AMP}}", "&")
                    Dim iconResource As String = Nothing
                    command.GetIcon(array, iconResource)
                    If Not String.IsNullOrWhiteSpace(iconResource) Then
                        menuItem.Icon = ImageHelper.ExtractIcon(iconResource)
                    End If
                    Dim guid As Guid = Guid.Empty
                    command.GetCanonicalName(guid)
                    Dim action As Action =
                        Sub()
                            Dim bindCtx As ComTypes.IBindCtx = Nothing
                            Functions.CreateBindCtx(0, bindCtx)
                            command.Invoke(array, Marshal.GetIUnknownForObject(bindCtx))
                        End Sub
                    menuItem.Tag = New Tuple(Of Integer, String, Object)(-1, guid.ToString(), action)
                    Dim enumExplorerCommand As IEnumExplorerCommand = Nothing
                    command.EnumSubCommands(enumExplorerCommand)
                    If Not enumExplorerCommand Is Nothing Then
                        menuItem.Items = New List(Of MenuItemData)()
                        Dim fetched As UInt32 = 1, cmd As IExplorerCommand = Nothing
                        While enumExplorerCommand.Next(1, cmd, fetched) = HRESULT.S_OK AndAlso fetched > 0
                            Dim subMenuItem As MenuItemData = getMenuItem(cmd, array, flags)
                            If Not subMenuItem Is Nothing Then
                                If flags.HasFlag(EXPCMDFLAGS.ECF_SEPARATORBEFORE) Then
                                    menuItem.Items.Add(New MenuItemData() With {.Header = "-----"})
                                End If
                                menuItem.Items.Add(subMenuItem)
                                If flags.HasFlag(EXPCMDFLAGS.ECF_SEPARATORAFTER) Then
                                    menuItem.Items.Add(New MenuItemData() With {.Header = "-----"})
                                End If
                            End If
                        End While
                    End If
                End If
                Return menuItem
            Else
                Return Nothing
            End If
        End Function

        Public Overrides Function InvokeCommand(id As Tuple(Of Integer, String, Object)) As Task
            Dim e As CommandInvokedEventArgs = Nothing

            UIHelper.OnUIThread(
                Sub()
                    Folder = Me.Folder
                    SelectedItems = Me.SelectedItems

                    e = New CommandInvokedEventArgs() With {
                        .Id = id.Item1,
                        .Verb = id.Item2
                    }
                    Dim c As Control = Me.Buttons.Cast(Of Control).FirstOrDefault(Function(i) id.Equals(i.Tag))
                    If TypeOf c Is ToggleButton Then e.IsChecked = CType(c, ToggleButton).IsChecked
                    RaiseCommandInvoked(e)
                End Sub)

            If Not e.IsHandled Then
                _thread.Add(
                    Sub()
                        CType(id.Item3, Action).Invoke()
                    End Sub)
            End If
        End Function

        Protected Overrides Function AddItems() As Task
            For Each item In _menuItems
                Me.Items.Add(getMenuItem(item))
            Next
            Return Task.CompletedTask
        End Function

        Private Function getMenuItem(item As MenuItemData) As Control
            If item.Header = "-----" Then
                Return New Separator()
            Else
                Dim menuItem As New MenuItem() With {
                    .Header = item.Header,
                    .IsEnabled = item.IsEnabled,
                    .Icon = New Image() With {.Source = item.Icon},
                    .Tag = item.Tag
                }
                If Not item.Items Is Nothing Then
                    For Each subItem In item.Items
                        menuItem.Items.Add(getMenuItem(subItem))
                    Next
                End If
                AddHandler menuItem.Click, AddressOf menuItem_Click
                Return menuItem
            End If
        End Function

        Protected Overrides Function DoRenameAfter(Tag As Tuple(Of Integer, String, Object)) As Boolean
            Return True
        End Function

        Protected Overrides Sub Dispose(disposing As Boolean)
            MyBase.Dispose(disposing)

            If Not _helper Is Nothing Then
                _helper.Dispose()
                _helper = Nothing
            End If
        End Sub
    End Class
End Namespace