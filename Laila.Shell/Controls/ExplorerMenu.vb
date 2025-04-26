Imports System.IO
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Windows.Controls
Imports System.Windows.Controls.Primitives
Imports System.Windows.Input
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
                    ' get cloud verbs
                    Dim assembly As Assembly = Assembly.LoadFrom("Laila.Shell.WinRT.dll")
                    Dim type As Type = assembly.GetType("Laila.Shell.WinRT.ExplorerMenuHelper")
                    Dim methodInfo As MethodInfo = type.GetMethod("GetCloudVerbs")
                    Dim instance As Object = Activator.CreateInstance(type)
                    Dim t As Task(Of Tuple(Of String, String, List(Of ExplorerCommandVerbInfo))) =
                        methodInfo.Invoke(instance, {folder.FullPath})
                    Dim result As Tuple(Of String, String, List(Of ExplorerCommandVerbInfo)) = t.Result

                    ' add cloud menu
                    Dim cloudMenuItem As New MenuItemData() With {
                        .Header = result.Item1,
                        .Items = New List(Of MenuItemData)()
                    }
                    If Not String.IsNullOrWhiteSpace(result.Item2) Then
                        cloudMenuItem.Icon = ImageHelper.ExtractIcon(result.Item2)
                    End If
                    _menuItems.Add(cloudMenuItem)

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
                End Sub)
        End Sub


        Private Function getMenuItem(command As IExplorerCommand, array As IShellItemArray, ByRef flags As EXPCMDFLAGS) As MenuItemData
            Dim state As EXPCMDSTATE = 0
            Dim h As HRESULT = command.GetState(array, True, state)
            If state = EXPCMDSTATE.ECS_ENABLED Then
                Dim menuItem As MenuItemData = New MenuItemData()
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