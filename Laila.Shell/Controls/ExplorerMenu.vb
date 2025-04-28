Imports System.IO
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports System.Security.Claims
Imports System.Text
Imports System.Windows.Controls
Imports System.Windows.Controls.Primitives
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports System.Windows.Input
Imports System.Windows.Media.Imaging
Imports Laila.Shell.Events
Imports Laila.Shell.Helpers
Imports Laila.Shell.Interop
Imports Laila.Shell.Interop.COM
Imports Laila.Shell.Interop.ContextMenu
Imports Laila.Shell.Interop.DragDrop
Imports Laila.Shell.Interop.Folders
Imports Laila.Shell.Interop.Functions
Imports Laila.Shell.Interop.Items
Imports Laila.Shell.Interop.Properties
Imports Laila.Shell.Interop.Windows
Imports Laila.Shell.WinRT
Imports Laila.Shell.WinRT.Interface
Imports Microsoft.VisualBasic.Devices
Imports Microsoft.Win32
Imports SHDocVw

Namespace Controls
    Public Class ExplorerMenu
        Inherits BaseContextMenu

        Private Const COMMANDSTORE_PATH As String = "Software\Microsoft\Windows\CurrentVersion\Explorer\CommandStore\shell"

        Private _contextMenuItems As List(Of MenuItemData) = New List(Of MenuItemData)()
        Private _menuItems As List(Of MenuItemData) = New List(Of MenuItemData)()
        Private _helper As Helpers.ComHelper = New ComHelper()

        Protected Overrides Sub Make(folder As Folder, items As IEnumerable(Of Item), isDefaultOnly As Boolean)
            If _wasMade Then Return

            _activeItems = If(Not Me.SelectedItems Is Nothing AndAlso Me.SelectedItems.Count > 0,
                Me.SelectedItems.ToList(), New List(Of Item) From {Me.Folder})

            ' get regular context menu items
            MyBase.Make(Me.Folder, Me.SelectedItems, Me.IsDefaultOnly)
            Dim tcs4 As New TaskCompletionSource(Of List(Of MenuItemData))
            _thread.Run(
                Sub()
                    tcs4.SetResult(getMenuItemData(_hMenu, -1))
                End Sub)
            _contextMenuItems = tcs4.Task.Result

            ' build menu
            _thread.Run(
                Sub()
                    ' instantiate the helper
                    Dim assembly As Assembly = Assembly.LoadFrom(IO.Path.Combine(IO.Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "Laila.Shell.WinRT.dll"))
                    Dim type As Type = assembly.GetType("Laila.Shell.WinRT.ExplorerMenuHelper")
                    Dim instance As Object = Activator.CreateInstance(type)

                    ' provide easy access to ResolveMsResourceFromPackage function in helper
                    Dim methodInfo As MethodInfo = type.GetMethod("ResolveMsResourceFromPackage")
                    Dim resolveMsResourceFromPackage As Func(Of String, String, String) =
                        Function(resourceReference As String, packageName As String) As String
                            Return methodInfo.Invoke(instance, {resourceReference, packageName})
                        End Function

                    ' get shell item array
                    Dim array As IShellItemArray = Nothing
                    If items.Count > 0 Then
                        Dim pidls As List(Of Pidl) = items.Select(Function(i) i.Pidl.Clone()).ToList()
                        Functions.SHCreateShellItemArrayFromIDLists(pidls.Count, pidls.Select(Function(p) p.AbsolutePIDL).ToArray(), array)
                    End If

                    ' add standard commands
                    addFromContextMenu("open", Nothing)
                    addFromContextMenu("openas", "Windows.OpenWith")
                    addFromContextMenu("empty", "Windows.RecycleBin.Empty")
                    addFromCommandStore("Windows.ModernShare", folder, items, array, resolveMsResourceFromPackage)
                    addFromContextMenu("undelete", "Windows.RecycleBin.RestoreItems")
                    addFromCommandStore("Windows.CompressTo", folder, items, array, resolveMsResourceFromPackage)
                    addFromContextMenu("copyaspath", "Windows.copyaspath")
                    addFromContextMenu("format", "Windows.DiskFormat")
                    addFromContextMenu("PinToStartScreen", "Windows.pintostartscreen")
                    addFromContextMenu("New", "Windows.newitem")
                    addFromContextMenu("properties", "Windows.properties")

                    ' add cloud menu item
                    addCloudMenuItem(type, instance, folder, array, resolveMsResourceFromPackage)

                    ' add file explorer context menu items
                    addFileExplorerContextMenuItems(type, instance, folder, items, array, resolveMsResourceFromPackage)

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

        Private Sub addFromContextMenu(verb As String, commandForIcon As String)
            Dim menuItem As MenuItemData = _contextMenuItems.FirstOrDefault(Function(i) Not i.Tag Is Nothing AndAlso CType(i.Tag, Tuple(Of Integer, String, Object)).Item2 = verb)
            If Not menuItem Is Nothing Then
                If Not commandForIcon Is Nothing Then
                    Using commandStoreKey As RegistryKey = Registry.LocalMachine.OpenSubKey(COMMANDSTORE_PATH)
                        Using subKey As RegistryKey = commandStoreKey.OpenSubKey(commandForIcon)
                            If Not subKey Is Nothing Then
                                ' get icon  
                                Dim icon As String = CType(subKey.GetValue("Icon"), String)
                                If Not icon Is Nothing Then
                                    menuItem.Icon = ImageHelper.ExtractIcon(icon)
                                End If
                            End If
                        End Using
                    End Using
                End If

                _menuItems.Add(menuItem)
            End If
        End Sub

        Private Sub addFromCommandStore(name As String, folder As Folder, items As IEnumerable(Of Item), array As IShellItemArray,
                                        resolveMsResourceFromPackage As Func(Of String, String, String))
            Using commandStoreKey As RegistryKey = Registry.LocalMachine.OpenSubKey(COMMANDSTORE_PATH)
                Using subKey As RegistryKey = commandStoreKey.OpenSubKey(name)
                    If Not subKey Is Nothing Then
                        ' get icon  
                        Dim icon As String = CType(subKey.GetValue("Icon"), String)

                        ' get localized text
                        Dim muiVerb As String = CType(subKey.GetValue("MUIVerb"), String)
                        If muiVerb?.StartsWith("@") Then
                            Dim sb As New StringBuilder(256)
                            Dim h = Functions.SHLoadIndirectString(muiVerb, sb, CUInt(sb.Capacity), IntPtr.Zero)
                            muiVerb = sb.ToString()
                        End If

                        If Not subKey.GetValue("ExplorerCommandHandler") Is Nothing Then
                            Try
                                ' instantiate explorer command
                                Dim handlerGUID As Guid = New Guid(CType(subKey.GetValue("ExplorerCommandHandler"), String))
                                Dim command As IExplorerCommand = Nothing
                                Functions.CoCreateInstance(handlerGUID, IntPtr.Zero, ClassContext.LocalServer Or ClassContext.InProcServer, GetType(IExplorerCommand).GUID, command)

                                ' init?
                                Dim init As IShellExtInit = TryCast(command, IShellExtInit)
                                If Not init Is Nothing Then
                                    Dim dobj As IDataObject_PreserveSig = Clipboard.GetDataObjectFor(folder, items)
                                    init.Initialize(folder.Pidl.AbsolutePIDL, dobj, IntPtr.Zero)
                                End If

                                ' set site?
                                Dim s As IObjectWithSite = TryCast(command, IObjectWithSite)
                                If Not s Is Nothing Then
                                    Dim sp As MockShellSite = New MockShellSite()
                                    Dim h As HRESULT = s.SetSite(sp)
                                End If

                                ' get menu item
                                If Not command Is Nothing Then
                                    Dim flags As EXPCMDFLAGS = 0
                                    Dim menuItem As MenuItemData = getMenuItem(command, array, flags, resolveMsResourceFromPackage)
                                    If Not menuItem Is Nothing Then
                                        If String.IsNullOrWhiteSpace(menuItem.Header) Then
                                            menuItem.Header = If(muiVerb, "")
                                        End If
                                        menuItem.Header = menuItem.Header
                                        If menuItem.Icon Is Nothing AndAlso Not icon Is Nothing Then
                                            menuItem.Icon = ImageHelper.ExtractIcon(icon)
                                        End If
                                        _menuItems.Add(menuItem)
                                    End If
                                End If
                            Catch ex As Exception
                            End Try
                        ElseIf Not subKey.GetValue("CommandStateHandler") Is Nothing Then
                            Dim handlerGUID As Guid = New Guid(CType(subKey.GetValue("CommandStateHandler"), String))
                            Dim commandType As Type = Type.GetTypeFromCLSID(handlerGUID)
                            Dim obj As Object = Activator.CreateInstance(commandType)
                            Dim commandState As IExplorerCommandState = TryCast(obj, IExplorerCommandState)
                            Dim command As IExecuteCommand = TryCast(obj, IExecuteCommand)
                            Dim init As IInitializeCommand = TryCast(obj, IInitializeCommand)
                            Dim owsel As IObjectWithSelection = TryCast(obj, IObjectWithSelection)
                            If Not owsel Is Nothing Then
                                Dim h As HRESULT = owsel.SetSelection(array)
                                Debug.WriteLine("SetSelection=" & h.ToString())
                            End If
                            If Not init Is Nothing Then
                                Dim verbName As String = CType(subKey.GetValue("VerbName"), String)
                                If verbName Is Nothing Then verbName = name
                                Dim bag As IPropertyBag = Nothing
                                Dim h As HRESULT = Functions.PSCreateMemoryPropertyStore(GetType(IPropertyBag).GUID, bag)
                                For Each valueName In subKey.GetValueNames()
                                    Dim value As String = CType(subKey.GetValue(valueName), String)
                                    If Not value Is Nothing Then
                                        Dim propValue As PROPVARIANT = New PROPVARIANT()
                                        propValue.SetValue(value)
                                        bag.Write(valueName, propValue)
                                    End If
                                Next
                                Debug.WriteLine("PSCreateMemoryPropertyStore=" & h.ToString())
                                h = init.Initialize(verbName, bag)
                                Debug.WriteLine("Initialize " & verbName & "=" & h.ToString())
                            End If
                            If Not commandState Is Nothing Then
                                Dim state As EXPCMDSTATE = 0
                                Dim h As HRESULT = commandState.GetState(Nothing, True, state)
                                Debug.WriteLine("GetState=" & h.ToString() & "--" & state.ToString())
                                If state <> 100 Then
                                    _menuItems.Add(New MenuItemData() With {
                                        .Header = name & "-" & If(muiVerb, ""),
                                        .Tag = handlerGUID
                                    })
                                    Dim a As Action =
                                            Sub()
                                                command.Execute()
                                            End Sub
                                    _menuItems(_menuItems.Count - 1).Tag = a
                                    If Not icon Is Nothing Then
                                        _menuItems(_menuItems.Count - 1).Icon = ImageHelper.ExtractIcon(icon)
                                    End If
                                End If
                            End If
                        End If
                    End If
                End Using
            End Using
        End Sub

        Private Sub addCloudMenuItem(type As Type, instance As Object, folder As Folder, array As IShellItemArray,
                                     resolveMsResourceFromPackage As Func(Of String, String, String))
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
                For Each item In result.Item3
                    Dim command As IExplorerCommand = _helper.MakeComObject(item.ComServerPath, item.ClsId, GetType(IExplorerCommand).GUID)
                    Dim flags As EXPCMDFLAGS
                    Dim menuItem As MenuItemData = getMenuItem(command, array, flags, resolveMsResourceFromPackage)
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

        Private Sub addFileExplorerContextMenuItems(type As Type, instance As Object, folder As Folder, items As IEnumerable(Of Item), array As IShellItemArray,
                                                    resolveMsResourceFromPackage As Func(Of String, String, String))
            ' get file explorer verbs
            Dim methodInfo As MethodInfo = type.GetMethod("GetFileExplorerVerbs")
            Dim t As Task(Of List(Of ExplorerCommandVerbInfo)) = methodInfo.Invoke(instance, {})
            Dim result As List(Of ExplorerCommandVerbInfo) = t.Result

            If Not result Is Nothing Then
                ' filter commands
                Dim isBackground As Boolean = (items Is Nothing OrElse items.Count = 0) AndAlso folder.Attributes.HasFlag(SFGAO.FILESYSTEM)
                Dim filtered As List(Of ExplorerCommandVerbInfo) = New List(Of ExplorerCommandVerbInfo)()
                For Each item In result
                    If (item.ItemType = "Directory\Background" AndAlso isBackground) _
                        OrElse (item.ItemType = "*" AndAlso items.ToList().Exists(Function(i) Not TypeOf i Is Folder)) _
                        OrElse (item.ItemType = "Directory" AndAlso items.ToList().Exists(Function(i) TypeOf i Is Folder AndAlso i.Attributes.HasFlag(SFGAO.FILESYSTEM))) _
                        OrElse (item.ItemType = "Drive" AndAlso items.ToList().Exists(Function(i) i.IsDrive)) _
                        OrElse (items.ToList().Exists(Function(i) item.ItemType.Equals(IO.Path.GetExtension(i.FullPath).ToLower()))) Then
                        If Not filtered.Exists(Function(f) f.ClsId = item.ClsId AndAlso f.ComServerPath = item.ComServerPath) Then
                            filtered.Add(item)
                        End If
                    End If
                Next

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
                        Dim menuItem As MenuItemData = getMenuItem(command, array, flags, resolveMsResourceFromPackage)
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

        Private Function getMenuItem(command As IExplorerCommand, array As IShellItemArray, ByRef flags As EXPCMDFLAGS,
                                     resolveMsResourceFromPackage As Func(Of String, String, String)) As MenuItemData
            Dim state As EXPCMDSTATE = 0
            Dim h As HRESULT = command.GetState(array, True, state)
            If state <> EXPCMDSTATE.ECS_HIDDEN Then
                Dim menuItem As MenuItemData = New MenuItemData() With {.IsEnabled = state = EXPCMDSTATE.ECS_ENABLED}
                command.GetFlags(flags)
                If flags.HasFlag(EXPCMDFLAGS.ECF_ISSEPARATOR) Then
                    menuItem.Header = "-----"
                Else
                    command.GetTitle(array, menuItem.Header)
                    menuItem.Header = If(menuItem.Header, "").Replace("&&", "{{AMP}}").Replace("&", "_").Replace("{{AMP}}", "&")
                    Dim iconResource As String = Nothing
                    command.GetIcon(array, iconResource)
                    If Not String.IsNullOrWhiteSpace(iconResource) Then
                        If iconResource.StartsWith("@") Then
                            Dim cleaned As String = iconResource.Substring(iconResource.IndexOf("{") + 1)
                            cleaned = cleaned.Substring(0, cleaned.IndexOf("}"))
                            Dim parts() As String = cleaned.Split("?")
                            Dim iconFileName As String = resolveMsResourceFromPackage(parts(1), parts(0))
                            If IO.File.Exists(iconFileName) Then
                                menuItem.Icon = New BitmapImage(New Uri(iconFileName, UriKind.Absolute))
                            End If
                        Else
                            menuItem.Icon = ImageHelper.ExtractIcon(iconResource)
                        End If
                        If Not menuItem.Icon Is Nothing Then
                            menuItem.Icon.Freeze()
                        End If
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
                        Dim fetched As UInt32, cmd(1) As IExplorerCommand
                        enumExplorerCommand.Next(1, cmd, fetched)
                        While fetched > 0
                            Dim flags2 As EXPCMDFLAGS = 0
                            Dim subMenuItem As MenuItemData = getMenuItem(cmd(0), array, flags2, resolveMsResourceFromPackage)
                            If Not subMenuItem Is Nothing Then
                                If flags2.HasFlag(EXPCMDFLAGS.ECF_SEPARATORBEFORE) Then
                                    menuItem.Items.Add(New MenuItemData() With {.Header = "-----"})
                                End If
                                menuItem.Items.Add(subMenuItem)
                                If flags2.HasFlag(EXPCMDFLAGS.ECF_SEPARATORAFTER) Then
                                    menuItem.Items.Add(New MenuItemData() With {.Header = "-----"})
                                End If
                            End If
                            enumExplorerCommand.Next(1, cmd, fetched)
                        End While
                    End If
                End If
                Return menuItem
            Else
                Return Nothing
            End If
        End Function

        Public Overrides Async Function InvokeCommand(id As Tuple(Of Integer, String, Object)) As Task
            If Not id.Item3 Is Nothing Then
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
            Else
                Await MyBase.InvokeCommand(id)
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

        <ComImport>
        <Guid("00000114-0000-0000-C000-000000000046")>
        <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
        Public Interface IOleWindow
            <PreserveSig>
            Function GetWindow(<Out> ByRef phwnd As IntPtr) As Integer

            <PreserveSig>
            Function ContextSensitiveHelp(<[In]> fEnterMode As Boolean) As Integer
        End Interface
        <ComImport>
        <Guid("831D1DFF-7F57-4720-87E4-CB57D6214428")>
        <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
        Public Interface IInvocationLocation
            ' HRESULT Proc3([out] int* p0);
            <PreserveSig>
            Function GetInvocationType(<Out> ByRef result As Integer) As Integer

            ' HRESULT Proc4([in] int p0);
            <PreserveSig>
            Function SetInvocationType(<[In]> value As Integer) As Integer
        End Interface


        <ComVisible(True), Guid("a985d29a-81ef-41b2-8440-457c38ef959c"), ProgId("Laila.Shell.Helpers.WpfDragTargetProxyc")>
        <ClassInterface(ClassInterfaceType.None)>
        Public Class MockIL
            Implements IInvocationLocation

            Public Function GetInvocationType(<Out> ByRef result As Integer) As Integer Implements IInvocationLocation.GetInvocationType
                Return 1
            End Function

            Public Function SetInvocationType(<[In]> value As Integer) As Integer Implements IInvocationLocation.SetInvocationType
                'Throw New NotImplementedException()
            End Function
        End Class

        <ComVisible(True), Guid("a985d29a-81ef-41b2-8440-457c38ef959c"), ProgId("Laila.Shell.Helpers.WpfDragTargetProxyc")>
        <ClassInterface(ClassInterfaceType.None)>
        Public Class Mockow
            Implements IOleWindow

            Public Function GetWindow(<Out> ByRef phwnd As IntPtr) As Integer Implements IOleWindow.GetWindow
                phwnd = Shell._hwnd
            End Function

            Public Function ContextSensitiveHelp(<[In]> fEnterMode As Boolean) As Integer Implements IOleWindow.ContextSensitiveHelp
                ' Throw New NotImplementedException()
            End Function
        End Class


        <ComVisible(True), Guid("a985d29a-81ef-41b2-8440-457c38ef959c"), ProgId("Laila.Shell.Helpers.WpfDragTargetProxyc")>
        <ClassInterface(ClassInterfaceType.None)>
        Public Class MockShellSite
            Implements Interop.Items.IServiceProvider

            Public Function QueryService(ByRef guidService As Guid, ByRef riid As Guid, ByRef ppvObject As IntPtr) As Integer Implements IServiceProvider.QueryService
                ' You can return your own IShellItemArray, IShellView, etc. here
                If riid = New Guid("831D1DFF-7F57-4720-87E4-CB57D6214428") Then
                    ppvObject = Marshal.GetComInterfaceForObject(New MockIL(), GetType(IInvocationLocation))
                    Debug.WriteLine("QueryServiceOK " & guidService.ToString() & " " & riid.ToString())
                    Return HRESULT.S_OK
                End If
                If riid = New Guid("00000114-0000-0000-C000-000000000046") Then
                    ppvObject = Marshal.GetComInterfaceForObject(New Mockow(), GetType(IOleWindow))
                    Debug.WriteLine("QueryServiceOK " & guidService.ToString() & " " & riid.ToString())
                    Return HRESULT.S_OK
                End If
                'Dim shellWindows As New ShellWindows()
                'Dim sp As IServiceProvider = Nothing
                '' Loop through all windows
                'For i As Integer = 0 To shellWindows.Count - 1
                '    Dim window As Object = shellWindows.Item(i)

                '    ' Try to cast to IServiceProvider
                '    sp = TryCast(window, IServiceProvider)
                '    If sp IsNot Nothing Then
                '        Exit For
                '    End If
                'Next
                'Return sp.QueryService(guidService, riid, ppvObject)
                ppvObject = IntPtr.Zero
                Debug.WriteLine("QueryService " & guidService.ToString() & " " & riid.ToString())
                Return HRESULT.E_NOINTERFACE
            End Function
        End Class

        Protected Overrides Sub Dispose(disposing As Boolean)
            MyBase.Dispose(disposing)

            If Not _helper Is Nothing Then
                _helper.Dispose()
                _helper = Nothing
            End If
        End Sub
    End Class
End Namespace