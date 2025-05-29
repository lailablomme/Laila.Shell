Imports System.Reflection
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Controls.Primitives
Imports System.Windows.Interop
Imports System.Windows.Media
Imports System.Windows.Media.Imaging
Imports Laila.Shell.Events
Imports Laila.Shell.Helpers
Imports Laila.Shell.Interop
Imports Laila.Shell.Interop.COM
Imports Laila.Shell.Interop.ContextMenu
Imports Laila.Shell.Interop.DragDrop
Imports Laila.Shell.Interop.Items
Imports Laila.Shell.Interop.Properties
Imports Laila.Shell.WinRT.Interface
Imports Microsoft.Win32

Namespace Controls
    Public Class ExplorerMenu
        Inherits BaseContextMenu

        Private Const COMMANDSTORE_PATH As String = "Software\Microsoft\Windows\CurrentVersion\Explorer\CommandStore\shell"

        Private _contextMenuItems As List(Of MenuItemData) = New List(Of MenuItemData)()
        Private _menuItems As List(Of MenuItemData) = New List(Of MenuItemData)()
        Private _helper As Helpers.ComHelper = New ComHelper()
        Private _arrayCloudItems As IShellItemArray = Nothing
        Private _arrayFileExplorerItems As IShellItemArray = Nothing
        Private _resourcePrefix As String

        Protected Overrides Sub Make(folder As Folder, items As IEnumerable(Of Item), isDefaultOnly As Boolean)
            If _wasMade Then Return

            _resourcePrefix = Me.ResourcePrefix

            _activeItems = If(Not Me.SelectedItems Is Nothing AndAlso Me.SelectedItems.Count > 0,
                Me.SelectedItems.ToList(), New List(Of Item) From {Me.Folder})

            ' get regular context menu items
            MyBase.Make(Me.Folder, Me.SelectedItems, Me.IsDefaultOnly)
            Dim tcs4 As New TaskCompletionSource(Of List(Of MenuItemData))
            _thread.Run(
                Sub()
                    tcs4.SetResult(getMenuItemData(_hMenu, -1, True))
                End Sub)
            _contextMenuItems = tcs4.Task.Result

            ' build menu
            _thread.Run(
                Sub()
                    ' instantiate the helper
                    Dim assembly As Assembly = Assembly.LoadFrom(IO.Path.Combine(IO.Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "Laila.Shell.WinRT.dll"))
                    Dim explorerMenuHelperType As Type = assembly.GetType("Laila.Shell.WinRT.ExplorerMenuHelper")
                    Dim explorerMenuHelper As Object = Activator.CreateInstance(explorerMenuHelperType)

                    ' provide easy access to ResolveMsResourceFromPackage function in helper
                    Dim methodInfo As MethodInfo = explorerMenuHelperType.GetMethod("ResolveMsResourceFromPackage")
                    Dim resolveMsResourceFromPackage As Func(Of String, String, String, String) =
                        Function(resourceReference As String, packageFullName As String, packageName As String) As String
                            Return methodInfo.Invoke(explorerMenuHelper, {resourceReference, packageFullName, packageName})
                        End Function

                    ' get shell item arrays
                    Dim pidls As List(Of Pidl) = Nothing
                    Try
                        If Not items Is Nothing AndAlso items.Count > 0 Then
                            pidls = items.Select(Function(i) i.Pidl.Clone()).ToList()
                            Functions.SHCreateShellItemArrayFromIDLists(pidls.Count, pidls.Select(Function(p) p.AbsolutePIDL).ToArray(), _arrayCloudItems)
                            Functions.SHCreateShellItemArrayFromIDLists(pidls.Count, pidls.Select(Function(p) p.AbsolutePIDL).ToArray(), _arrayFileExplorerItems)
                        Else
                            pidls = {folder}.Select(Function(i) i.Pidl.Clone()).ToList()
                            Functions.SHCreateShellItemArrayFromIDLists(pidls.Count, pidls.Select(Function(p) p.AbsolutePIDL).ToArray(), _arrayCloudItems)
                        End If
                    Finally
                        If Not pidls Is Nothing Then
                            For Each pidl In pidls
                                pidl.Dispose()
                            Next
                        End If
                    End Try

                    ' add standard commands
                    Dim isBackground As Boolean = (items Is Nothing OrElse items.Count = 0) AndAlso folder.Attributes.HasFlag(SFGAO.FILESYSTEM)
                    Dim isMount As Boolean = False
                    If Not isBackground _
                        AndAlso (items?.ToList().All(Function(i) TypeOf i Is Folder) _
                                 OrElse items?.ToList().All(Function(i) i.IsDrive) _
                                 OrElse items?.ToList().All(Function(i) IO.Path.GetExtension(items(0).FullPath).ToLower().Equals(IO.Path.GetExtension(i.FullPath).ToLower()))) Then
                        addFromContextMenu("mount", "Windows.mount")
                        isMount = _menuItems.Count > 0
                        If Not isMount Then
                            addFromContextMenu("open", Nothing, "Enter")
                            If _menuItems.Count > 0 Then
                                If TypeOf items(0) Is Folder Or items(0).IsDrive Then
                                    Try
                                        _menuItems(0).Icon = ImageHelper.GetApplicationIcon(Assembly.GetEntryAssembly().Location)
                                    Catch ex As Exception
                                        ' Protect against invalid image
                                    End Try
                                Else
                                    addFromCommandStore("Windows.OpenWith", folder, items, _arrayFileExplorerItems, resolveMsResourceFromPackage)
                                    If _menuItems.Count > 1 Then
                                        If Not _menuItems(1).Items(0).Icon Is Nothing Then
                                            _menuItems(0).Icon = _menuItems(1).Items(0).Icon
                                        End If
                                        _menuItems.RemoveAt(1)
                                    End If
                                    If _menuItems(0).Icon Is Nothing Then
                                        _menuItems(0).Icon = items(0).AssociatedApplicationIcon(48)
                                    End If
                                End If
                            End If
                        End If
                    End If
                    addFromContextMenu("openas", "Windows.OpenWith")
                    addFromContextMenu("empty", "Windows.RecycleBin.Empty")
                    addFromContextMenu("Windows.ModernShare", "Windows.ModernShare")
                    addFromContextMenu("eject", "Windows.Eject")
                    addFromContextMenu("PinToStartScreen", "Windows.pintostartscreen")
                    addFromContextMenu("Install", "Windows.Install")
                    addFromContextMenu("rotate90", "Windows.rotate90")
                    addFromContextMenu("rotate270", "Windows.rotate270")
                    addFromContextMenu("Windows.PowerShell.Run", "Windows.MultiVerb.PowerShell")
                    addFromContextMenu("opencontaining", "Windows.Shortcut.opencontaining")
                    addFromContextMenu("undelete", "Windows.RecycleBin.RestoreItems")
                    addFromContextMenu("extract", "Windows.CompressedFile.extract")
                    addFromContextMenu("setdesktopwallpaper", "Windows.setdesktopwallpaper")
                    addFromCommandStore("Windows.CompressTo", folder, items, _arrayFileExplorerItems, resolveMsResourceFromPackage)
                    addFromContextMenu("copyaspath", "Windows.copyaspath", "Ctrl-Shift-C")
                    addFromContextMenu("format", "Windows.DiskFormat")
                    addFromContextMenu("connectNetworkDrive", "Windows.connectNetworkDrive")
                    addFromContextMenu("disconnectNetworkDrive", "Windows.DisconnectNetworkDrive")
                    addFromContextMenu("New", "Windows.newitem")
                    addFromContextMenu("properties", "Windows.properties", "Alt+Enter")

                    ' add cloud menu items
                    If _menuItems.Count > 0 Then
                        _menuItems.Add(New MenuItemData() With {.Header = "-----"})
                    End If
                    addFromContextMenu("MakeAvailableOffline", "cloud_download")
                    addFromContextMenu("MakeAvailableOnline", "cloud")
                    addCloudMenuItem(explorerMenuHelperType, explorerMenuHelper, folder, items, _arrayCloudItems, resolveMsResourceFromPackage)

                    ' add file explorer context menu items
                    addFileExplorerContextMenuItems(explorerMenuHelperType, explorerMenuHelper, folder, items, _arrayFileExplorerItems, resolveMsResourceFromPackage)

                    If isMount AndAlso _menuItems.Count > 1 Then
                        _menuItems.Insert(1, New MenuItemData() With {.Header = "-----"})
                    End If

                    ' add show more options command
                    If _contextMenuItems.Count > 0 Then
                        If _menuItems.Count > 0 Then
                            _menuItems.Add(New MenuItemData() With {.Header = "-----"})
                        End If
                        Dim icon As ImageSource = System.Windows.Application.Current.TryFindResource($"{_resourcePrefix}MoreOptionsIcon")
                        icon.Freeze()
                        Dim showMoreOptionsMenuItem As New MenuItemData() With {
                            .Header = My.Resources.Menu_ShowMoreOptions,
                            .IsEnabled = True,
                            .Icon = icon
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
                    End If
                End Sub)
        End Sub

        Private Sub addFromContextMenu(verb As String, commandForIcon As String, Optional shortcutKeyText As String = Nothing)
            Dim menuItem As MenuItemData = _contextMenuItems.FirstOrDefault(Function(i) Not i.Tag Is Nothing AndAlso CType(i.Tag, Tuple(Of Integer, String, Object)).Item2 = verb)
            If Not menuItem Is Nothing Then
                If Not commandForIcon Is Nothing Then
                    Using commandStoreKey As RegistryKey = Registry.LocalMachine.OpenSubKey(COMMANDSTORE_PATH)
                        Using subKey As RegistryKey = commandStoreKey.OpenSubKey(commandForIcon)
                            If Not subKey Is Nothing Then
                                ' get icon  
                                Dim icon As String = CType(subKey.GetValue("Icon"), String)
                                If Not icon Is Nothing Then
                                    menuItem.Icon = ImageHelper.ExtractIcon(icon, False)
                                End If
                            End If
                        End Using
                    End Using
                    If menuItem.Icon Is Nothing Then
                        Try
                            menuItem.Icon = System.Windows.Application.Current.TryFindResource($"{_resourcePrefix}{commandForIcon}Icon")
                            menuItem.Icon?.Freeze()
                        Catch ex As Exception
                        End Try
                    End If
                End If

                menuItem.FontWeight = FontWeights.Regular
                menuItem.ShortcutKeyText = shortcutKeyText
                _menuItems.Add(menuItem)
            End If
        End Sub

        Private Sub addFromCommandStore(name As String, folder As Folder, items As IEnumerable(Of Item), array As IShellItemArray,
                                        resolveMsResourceFromPackage As Func(Of String, String, String, String))
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
                                    Dim sp As MockServiceProvider = New MockServiceProvider()
                                    Dim h As HRESULT = s.SetSite(sp)
                                End If

                                ' get menu item
                                If Not command Is Nothing Then
                                    Dim flags As EXPCMDFLAGS = 0
                                    Dim menuItem As MenuItemData = getMenuItem(command, array, flags, folder, items, resolveMsResourceFromPackage)
                                    If Not menuItem Is Nothing Then
                                        If String.IsNullOrWhiteSpace(menuItem.Header) Then
                                            menuItem.Header = If(muiVerb, "")
                                        End If
                                        menuItem.Header = menuItem.Header
                                        If menuItem.Icon Is Nothing AndAlso Not icon Is Nothing Then
                                            menuItem.Icon = ImageHelper.ExtractIcon(icon, False)
                                        End If
                                        If menuItem.Icon Is Nothing Then
                                            Try
                                                menuItem.Icon = System.Windows.Application.Current.TryFindResource($"{_resourcePrefix}{name}Icon")
                                                menuItem.Icon?.Freeze()
                                            Catch ex As Exception
                                            End Try
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
                                If state <> EXPCMDSTATE.ECS_HIDDEN Then
                                    _menuItems.Add(New MenuItemData() With {
                                        .Header = name & "-" & If(muiVerb, ""),
                                        .IsEnabled = state = EXPCMDSTATE.ECS_ENABLED
                                    })
                                    Dim action As Action =
                                        Sub()
                                            command.Execute()
                                        End Sub
                                    Dim canonicalName As String = If(subKey.GetValue("CanonicalName"), handlerGUID.ToString())
                                    _menuItems(_menuItems.Count - 1).Tag = New Tuple(Of Integer, String, Object)(-1, canonicalName, action)
                                    If Not icon Is Nothing Then
                                        _menuItems(_menuItems.Count - 1).Icon = ImageHelper.ExtractIcon(icon, False)
                                    End If
                                End If
                            End If
                        End If
                    End If
                End Using
            End Using
        End Sub

        Private Sub addCloudMenuItem(type As Type, instance As Object, folder As Folder, items As IEnumerable(Of Item), array As IShellItemArray,
                                     resolveMsResourceFromPackage As Func(Of String, String, String, String))
            ' get parent folder
            Dim parentFolder As Folder = Nothing
            If Not items Is Nothing AndAlso items.Count > 0 Then
                parentFolder = items(0).Parent
            Else
                parentFolder = folder.Parent
                items = New List(Of Item) From {folder}
            End If

            ' cannot get cloud verbs if parent folder is nothing or if all items are not filesystem items or if items have multiple different parent folders
            If parentFolder Is Nothing _
                OrElse (Not items Is Nothing AndAlso items.Count > 0 AndAlso items.All(Function(i) Not i.Attributes.HasFlag(SFGAO.FILESYSTEM))) _
                OrElse (Not items Is Nothing AndAlso items.Count > 0 AndAlso items.Any(Function(i) Not parentFolder.Equals(i.Parent))) Then Return

            ' get cloud verbs
            Dim methodInfo As MethodInfo = type.GetMethod("GetCloudVerbs")
            Dim t As Task(Of Tuple(Of String, String, List(Of ExplorerCommandVerbInfo))) =
            methodInfo.Invoke(instance, {parentFolder.FullPath})
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
                    Dim menuItem As MenuItemData = getMenuItem(command, array, flags, folder, items, resolveMsResourceFromPackage)
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
                    _menuItems.Add(cloudMenuItem)
                End If
            End If
        End Sub

        Private Sub addFileExplorerContextMenuItems(type As Type, instance As Object, folder As Folder, items As IEnumerable(Of Item), array As IShellItemArray,
                                                    resolveMsResourceFromPackage As Func(Of String, String, String, String))
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
                        OrElse (item.ItemType = "*" AndAlso items?.ToList().Exists(Function(i) Not TypeOf i Is Folder)) _
                        OrElse (item.ItemType = "Directory" AndAlso items?.ToList().Exists(Function(i) TypeOf i Is Folder AndAlso i.Attributes.HasFlag(SFGAO.FILESYSTEM))) _
                        OrElse (item.ItemType = "Drive" AndAlso items?.ToList().Exists(Function(i) i.IsDrive)) _
                        OrElse (items?.ToList().Exists(Function(i) item.ItemType.Equals(IO.Path.GetExtension(i.FullPath).ToLower()))) Then
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
                        Dim icon As ImageSource = Nothing
                        Try
                            icon = If(Not String.IsNullOrWhiteSpace(thisAppCommands(0).ApplicationIconPath) _
                                        AndAlso IO.File.Exists(thisAppCommands(0).ApplicationIconPath),
                                      trimTransparentBorders(New BitmapImage(New Uri(thisAppCommands(0).ApplicationIconPath))),
                                      Nothing)
                            If Not icon Is Nothing Then icon.Freeze()
                        Catch ex As Exception
                            ' protect against invalid image
                        End Try

                        appMenuItem = New MenuItemData() With {
                            .ApplicationName = thisAppCommands(0).ApplicationName,
                            .Header = thisAppCommands(0).ApplicationName,
                            .Icon = icon,
                            .IsEnabled = True,
                            .Items = New List(Of MenuItemData)(),
                            .IsSubMenu = True
                        }
                        addTo = New List(Of MenuItemData)()
                    End If

                    For Each commandItem In thisAppCommands
                        Dim command As IExplorerCommand = _helper.MakeComObject(commandItem.ComServerPath, commandItem.ClsId, GetType(IExplorerCommand).GUID)
                        Dim flags As EXPCMDFLAGS
                        Dim menuItem As MenuItemData = getMenuItem(command, array, flags, folder, items, resolveMsResourceFromPackage)
                        If Not menuItem Is Nothing Then
                            If Not menuItem.Items Is Nothing AndAlso menuItem.Items.Count = 1 _
                                AndAlso menuItem.Items(0).Header?.Equals(menuItem.Header) Then
                                ' how is this possible that i have two apps who get this wrong?!
                                menuItem = menuItem.Items(0)
                            End If
                            menuItem.ApplicationName = commandItem.ApplicationName
                            If menuItem.Icon Is Nothing AndAlso IO.File.Exists(commandItem.ApplicationIconPath) Then
                                Try
                                    menuItem.Icon = trimTransparentBorders(New BitmapImage(New Uri(commandItem.ApplicationIconPath, UriKind.Absolute)))
                                    If Not menuItem.Icon Is Nothing Then menuItem.Icon.Freeze()
                                Catch ex As Exception
                                    ' Protect against invalid image
                                End Try
                            End If
                            If menuItem.Icon Is Nothing _
                                AndAlso Not String.IsNullOrWhiteSpace(commandItem.ApplicationExecutable) _
                                AndAlso Not String.IsNullOrWhiteSpace(commandItem.InstalledPath) Then
                                Dim fullExePath As String = IO.Path.Combine(commandItem.InstalledPath, commandItem.ApplicationExecutable)
                                While Not IO.File.Exists(fullExePath) _
                                    AndAlso Not String.IsNullOrWhiteSpace(IO.Path.GetDirectoryName(IO.Path.GetDirectoryName(fullExePath)))
                                    fullExePath = IO.Path.Combine(IO.Path.GetDirectoryName(IO.Path.GetDirectoryName(fullExePath)), commandItem.ApplicationExecutable)
                                End While
                                If IO.File.Exists(fullExePath) Then
                                    menuItem.Icon = ImageHelper.GetApplicationIcon(fullExePath)
                                End If
                            End If
                            addTo.Add(menuItem)
                        End If
                    Next

                    If Not addTo.Equals(tempRoot) AndAlso addTo.Count > 0 Then
                        appMenuItem.Items.AddRange(addTo.OrderBy(Function(i) i.Header))
                        tempRoot.Add(appMenuItem)
                    End If
                Next

                If tempRoot.Count > 0 Then
                    If _menuItems.Count > 0 AndAlso _menuItems(_menuItems.Count - 1).Header <> "-----" Then
                        _menuItems.Add(New MenuItemData() With {.Header = "-----"})
                    End If
                    _menuItems.AddRange(tempRoot.OrderBy(Function(i) $"{i.ApplicationName}\{i.Header}"))
                End If
            End If
        End Sub

        Private Function getMenuItem(command As IExplorerCommand, array As IShellItemArray, ByRef flags As EXPCMDFLAGS,
                                     folder As Folder, items As IEnumerable(Of Item),
                                     resolveMsResourceFromPackage As Func(Of String, String, String, String)) As MenuItemData
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
                            Dim iconFileName As String = resolveMsResourceFromPackage(parts(1), parts(0), Nothing)
                            If IO.File.Exists(iconFileName) Then
                                Try
                                    menuItem.Icon = putOnBlueBackground(trimTransparentBorders(New BitmapImage(New Uri(iconFileName, UriKind.Absolute))))
                                Catch ex As Exception
                                    ' Protect against invalid image
                                End Try
                            End If
                        Else
                            menuItem.Icon = ImageHelper.ExtractIcon(iconResource, False)
                        End If
                        If Not menuItem.Icon Is Nothing Then
                            menuItem.Icon.Freeze()
                        End If
                    End If
                    Dim guid As Guid = Guid.Empty
                    command.GetCanonicalName(guid)

                    ' init?
                    Dim init As IShellExtInit = TryCast(command, IShellExtInit)
                    If Not init Is Nothing Then
                        Dim dobj As IDataObject_PreserveSig = Clipboard.GetDataObjectFor(folder, items)
                        init.Initialize(folder.Pidl.AbsolutePIDL, dobj, IntPtr.Zero)
                    End If

                    ' set site?
                    Dim s As IObjectWithSite = TryCast(command, IObjectWithSite)
                    If Not s Is Nothing Then
                        Dim sp As MockServiceProvider = New MockServiceProvider()
                        Dim h2 As HRESULT = s.SetSite(sp)
                    End If

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
                            Dim subMenuItem As MenuItemData = getMenuItem(cmd(0), array, flags2, folder, items, resolveMsResourceFromPackage)
                            If Not subMenuItem Is Nothing Then
                                menuItem.IsSubMenu = True
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

        Private Function putOnBlueBackground(source As BitmapSource) As BitmapSource
            Dim size As Integer = If(source.Width > source.Height, source.Width, source.Height)
            Dim pixelSize As Integer = If(source.PixelWidth > source.PixelHeight, source.PixelWidth, source.PixelHeight)

            Dim dv As New DrawingVisual()
            Using dc As DrawingContext = dv.RenderOpen()
                ' Fill the background with blue
                dc.DrawRectangle(Brushes.Blue, Nothing, New Rect(0, 0, size, size))

                ' Draw the image onto the bitmap 
                dc.DrawImage(source, New Rect((size - source.Width) / 2, (size - source.Height) / 2, source.Width, source.Height))
            End Using

            ' 4. Render the DrawingVisual to a RenderTargetBitmap
            Dim rtb As New RenderTargetBitmap(pixelSize, pixelSize, source.DpiX, source.DpiY, PixelFormats.Pbgra32)
            rtb.Render(dv)

            Return rtb
        End Function

        Private Function trimTransparentBorders(source As BitmapSource) As BitmapSource
            ' Convert to BGRA32 if needed
            If source.Format <> PixelFormats.Bgra32 Then
                source = New FormatConvertedBitmap(source, PixelFormats.Bgra32, Nothing, 0)
            End If

            Dim width = source.PixelWidth
            Dim height = source.PixelHeight
            Dim stride = width * 4
            Dim pixels(width * height * 4 - 1) As Byte
            source.CopyPixels(pixels, stride, 0)

            ' Find bounding box of non-transparent pixels
            Dim minX = width, minY = height, maxX = 0, maxY = 0
            Dim foundPixel = False

            For y = 0 To height - 1
                For x = 0 To width - 1
                    Dim index = y * stride + x * 4
                    Dim alpha = pixels(index + 3)
                    If alpha <> 0 Then
                        foundPixel = True
                        If x < minX Then minX = x
                        If y < minY Then minY = y
                        If x > maxX Then maxX = x
                        If y > maxY Then maxY = y
                    End If
                Next
            Next

            If Not foundPixel Then
                ' Entirely transparent image — return a 1x1 transparent bitmap with original DPI
                Dim empty = New WriteableBitmap(1, 1, source.DpiX, source.DpiY, PixelFormats.Bgra32, Nothing)
                empty.Lock()
                empty.Unlock()
                Return empty
            End If

            ' Define crop region
            Dim cropWidth = maxX - minX + 1
            Dim cropHeight = maxY - minY + 1
            Dim cropStride = cropWidth * 4
            Dim cropPixels(cropWidth * cropHeight * 4 - 1) As Byte

            ' Copy pixels from the crop region
            For y = 0 To cropHeight - 1
                Dim srcIndex = (minY + y) * stride + minX * 4
                Dim dstIndex = y * cropStride
                Array.Copy(pixels, srcIndex, cropPixels, dstIndex, cropStride)
            Next

            ' Create new WriteableBitmap with the original DPI
            Dim result = New WriteableBitmap(cropWidth, cropHeight, source.DpiX, source.DpiY, PixelFormats.Bgra32, Nothing)
            result.WritePixels(New Int32Rect(0, 0, cropWidth, cropHeight), cropPixels, cropStride, 0)
            Return result
        End Function

        Private Function isEmpty(source As BitmapImage) As Boolean
            Dim width = source.PixelWidth
            Dim height = source.PixelHeight
            Dim stride = width * 4
            Dim pixels(width * height * 4 - 1) As Byte
            source.CopyPixels(pixels, stride, 0)

            ' Find bounding box of non-transparent pixels
            Dim minX = width, minY = height, maxX = 0, maxY = 0
            Dim foundPixel = False

            For y = 0 To height - 1
                For x = 0 To width - 1
                    Dim index = y * stride + x * 4
                    Dim alpha = pixels(index + 3)
                    If alpha <> 0 Then
                        foundPixel = True
                        If x < minX Then minX = x
                        If y < minY Then minY = y
                        If x > maxX Then maxX = x
                        If y > maxY Then maxY = y
                    End If
                Next
            Next

            Return Not foundPixel
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
                    ' Or just deactivate focus
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
            ' add right click menu items
            If Not Me.SelectedItems Is Nothing AndAlso Me.SelectedItems.Count = 1 AndAlso TypeOf Me.SelectedItems(0) Is Folder Then
                CType(Me.SelectedItems(0), Folder).AddExplorerMenuItems(Me)
            Else
                Me.Folder.AddExplorerMenuItems(Me)
            End If

            Dim hasItems As Boolean = Me.Items.Count > 0

            ' add view, sort and group by menus
            If Me.SelectedItems Is Nothing OrElse Me.SelectedItems.Count = 0 Then
                Dim insertIndex As Integer = 0

                Dim viewMenu As ViewMenu = New ViewMenu() With {.Folder = Me.Folder, .MenuStyle = ViewMenuStyle.RightClickMenu}
                Dim viewMenuItem As MenuItem = New MenuItem() With {
                    .Header = My.Resources.Menu_View,
                    .Icon = New Image() With {.Source = System.Windows.Application.Current.TryFindResource($"{_resourcePrefix}ViewIcon")}
                }
                viewMenu.AddItems(viewMenuItem.Items, _resourcePrefix)
                Me.Items.Insert(0, viewMenuItem)
                insertIndex += 1

                If Not Me.Folder.ShellItem2 Is Nothing Then
                    Dim sortMenu As SortMenu = New SortMenu() With {.Folder = Me.Folder}

                    If Me.Folder.CanSort Then
                        Dim sortMenuItem As MenuItem = New MenuItem() With {
                            .Header = My.Resources.Menu_Sort,
                            .Icon = New Image() With {.Source = System.Windows.Application.Current.TryFindResource($"{_resourcePrefix}SortIcon")}
                        }
                        sortMenu.AddSortItems(sortMenuItem.Items, _resourcePrefix)
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

        Public Overrides Function GetMenuItems() As List(Of Control)
            Return MyBase.getMenuItems(_menuItems)
        End Function

        Protected Overrides Function DoRenameAfter(tag As Tuple(Of Integer, String, Object)) As Boolean
            Return tag.Item2 <> "paste"
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
            <PreserveSig>
            Function GetInvocationType(<Out> ByRef result As Integer) As Integer

            <PreserveSig>
            Function SetInvocationType(<[In]> value As Integer) As Integer
        End Interface

        <ComVisible(True), Guid("2c62e9b0-cc2e-4c95-8717-e7b46340983f"), ProgId("Laila.Shell.Controls.ExplorerMenu.MockInvocationLocation")>
        <ClassInterface(ClassInterfaceType.None)>
        Public Class MockInvocationLocation
            Implements IInvocationLocation

            Public Function GetInvocationType(<Out> ByRef result As Integer) As Integer Implements IInvocationLocation.GetInvocationType
                Return &HFFFFFF
            End Function

            Public Function SetInvocationType(<[In]> value As Integer) As Integer Implements IInvocationLocation.SetInvocationType
                Return HRESULT.S_OK
            End Function
        End Class

        <ComVisible(True), Guid("bb07f6f9-2573-4be5-9f14-7a4fa6a90f48"), ProgId("Laila.Shell.Controls.ExplorerMenu.MockOleWindow")>
        <ClassInterface(ClassInterfaceType.None)>
        Public Class MockOleWindow
            Implements IOleWindow

            Public Function GetWindow(<Out> ByRef phwnd As IntPtr) As Integer Implements IOleWindow.GetWindow
                phwnd = New WindowInteropHelper(If(New List(Of Window)(
                                                      System.Windows.Application.Current.Windows.Cast(Of Window)()) _
                                                          .Find(Function(w) w.IsActive),
                                                   System.Windows.Application.Current.MainWindow)).Handle
                Return HRESULT.S_OK
            End Function

            Public Function ContextSensitiveHelp(<[In]> fEnterMode As Boolean) As Integer Implements IOleWindow.ContextSensitiveHelp
                Return HRESULT.S_OK
            End Function
        End Class

        <ComVisible(True), Guid("6e6c5e4a-08f3-43cf-84d2-0df5ae6f6d88"), ProgId("Laila.Shell.Controls.ExplorerMenu.MockServiceProvider")>
        <ClassInterface(ClassInterfaceType.None)>
        Public Class MockServiceProvider
            Implements Interop.Items.IServiceProvider

            Public Function QueryService(ByRef guidService As Guid, ByRef riid As Guid, ByRef ppvObject As IntPtr) As Integer Implements IServiceProvider.QueryService
                If riid = New Guid("831D1DFF-7F57-4720-87E4-CB57D6214428") Then
                    ppvObject = Marshal.GetComInterfaceForObject(New MockInvocationLocation(), GetType(IInvocationLocation))
                    Debug.WriteLine("QueryServiceOK " & guidService.ToString() & " " & riid.ToString())
                    Return HRESULT.S_OK
                End If
                If riid = New Guid("00000114-0000-0000-C000-000000000046") Then
                    ppvObject = Marshal.GetComInterfaceForObject(New MockOleWindow(), GetType(IOleWindow))
                    Debug.WriteLine("QueryServiceOK " & guidService.ToString() & " " & riid.ToString())
                    Return HRESULT.S_OK
                End If

                ppvObject = IntPtr.Zero
                Debug.WriteLine("MockServiceProvider.QueryService " & guidService.ToString() & " " & riid.ToString())
                Return HRESULT.E_NOINTERFACE
            End Function
        End Class

        Protected Overrides Sub Dispose(disposing As Boolean)
            MyBase.Dispose(disposing)

            If Not _helper Is Nothing Then
                _helper.Dispose()
                _helper = Nothing
            End If
            If Not _arrayCloudItems Is Nothing Then
                Marshal.ReleaseComObject(_arrayCloudItems)
                _arrayCloudItems = Nothing
            End If
            If Not _arrayFileExplorerItems Is Nothing Then
                Marshal.ReleaseComObject(_arrayFileExplorerItems)
                _arrayFileExplorerItems = Nothing
            End If
        End Sub
    End Class
End Namespace