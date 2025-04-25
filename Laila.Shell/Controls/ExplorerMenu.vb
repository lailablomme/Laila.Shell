Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Windows.Controls
Imports System.Windows.Input
Imports Laila.Shell.Helpers
Imports Laila.Shell.Interop
Imports Laila.Shell.Interop.ContextMenu
Imports Laila.Shell.Interop.DragDrop
Imports Laila.Shell.Interop.Folders
Imports Laila.Shell.Interop.Items
Imports Laila.Shell.Interop.Properties
Imports Laila.Shell.Interop.Windows
Imports Microsoft.Win32
Imports SHDocVw

Namespace Controls
    Public Class ExplorerMenu
        Inherits BaseMenu

        Private Const COMMANDSTORE_PATH As String = "Software\Microsoft\Windows\CurrentVersion\Explorer\CommandStore\shell"

        Private _menuItems As List(Of MenuItemData) = New List(Of MenuItemData)()

        Protected Overrides Sub Make(folder As Folder, items As IEnumerable(Of Item), isDefaultOnly As Boolean)
            Dim array As IShellItemArray = Nothing
            Dim pidls As List(Of Pidl) = items.Select(Function(i) i.Pidl.Clone()).ToList()
            Functions.SHCreateShellItemArrayFromIDLists(pidls.Count, pidls.Select(Function(p) p.AbsolutePIDL).ToArray(), array)

            Dim grfKeyState As MK = 0
            If Keyboard.Modifiers.HasFlag(ModifierKeys.Shift) Then grfKeyState = grfKeyState Or MK.MK_SHIFT
            If Keyboard.Modifiers.HasFlag(ModifierKeys.Control) Then grfKeyState = grfKeyState Or MK.MK_CONTROL

            Using commandStoreKey As RegistryKey = Registry.LocalMachine.OpenSubKey(COMMANDSTORE_PATH)
                For Each subKeyName As String In commandStoreKey.GetSubKeyNames()
                    Debug.WriteLine(subKeyName)
                    If Not subKeyName = "Windows.ModernShareFlyout" _
                        AndAlso Not subKeyName = "Windows.newitem" _
                        AndAlso Not subKeyName = "Windows.RibbonShare" _
                        AndAlso Not subKeyName = "Windows.Share" _
                                         AndAlso Not subKeyName = "Windows.SendToMyPhone" _
       AndAlso Not subKeyName = "Windows.SharePrivate" Then
                        Using subKey As RegistryKey = commandStoreKey.OpenSubKey(subKeyName)
                            If subKey IsNot Nothing Then
                                Dim icon As String = CType(subKey.GetValue("Icon"), String)
                                Dim muiVerb As String = CType(subKey.GetValue("MUIVerb"), String)
                                If muiVerb?.StartsWith("@") Then
                                    Dim sb As New StringBuilder(256)
                                    Dim h = Functions.SHLoadIndirectString(muiVerb, sb, CUInt(sb.Capacity), IntPtr.Zero)
                                    muiVerb = sb.ToString()
                                End If

                                Dim handlerGUID As Guid
                                If Not subKey.GetValue("ExplorerCommandHandler") Is Nothing Then
                                    Try
                                        handlerGUID = New Guid(CType(subKey.GetValue("ExplorerCommandHandler"), String))
                                        Dim commandType As Type = Type.GetTypeFromCLSID(handlerGUID)
                                        Dim command As IExplorerCommand = TryCast(Activator.CreateInstance(commandType), IExplorerCommand)
                                        '   Dim iol As System.Runtime.InteropServices.ComTypes.iolewindow = TryCast(command, IShellExtInit)
                                        Dim init As IShellExtInit = TryCast(command, IShellExtInit)
                                        If Not init Is Nothing Then
                                            Dim dobj As IDataObject_PreserveSig = Clipboard.GetDataObjectFor(folder, items)
                                            init.Initialize(folder.Pidl.AbsolutePIDL, dobj, IntPtr.Zero)
                                        End If
                                        Dim s As IObjectWithSite = TryCast(command, IObjectWithSite)
                                        If Not s Is Nothing Then
                                            Dim sp As MockShellSite = New MockShellSite()
                                            'Dim shellView As IShellView = Nothing
                                            '' If Not disposedValue AndAlso Not Me.ShellFolder Is Nothing Then
                                            'folder.ShellFolder.CreateViewObject(IntPtr.Zero, Guids.IID_IShellView, shellView)
                                            ''  End If
                                            Dim h As HRESULT = s.SetSite(sp)
                                            Debug.WriteLine(h.ToString())
                                        End If
                                        If Not command Is Nothing Then
                                            Dim menuItem As MenuItemData = getMenuItem(command, array)
                                            If Not menuItem Is Nothing Then
                                                If String.IsNullOrWhiteSpace(menuItem.Header) Then
                                                    menuItem.Header = If(muiVerb, "")
                                                End If
                                                menuItem.Header = subKeyName & "-" & menuItem.Header
                                                If Not icon Is Nothing Then
                                                    menuItem.Icon = ImageHelper.ExtractIcon(icon)
                                                End If
                                                Dim a As Action =
                                                    Sub()
                                                        Dim bindCtx As ComTypes.IBindCtx = Nothing
                                                        Functions.CreateBindCtx(0, bindCtx)
                                                        command.Invoke(array, Marshal.GetIUnknownForObject(bindCtx))
                                                    End Sub
                                                menuItem.Tag = a
                                                _menuItems.Add(menuItem)
                                            End If
                                        End If
                                    Catch ex As Exception
                                    End Try
                                ElseIf Not subKey.GetValue("CommandStateHandler") Is Nothing Then
                                    handlerGUID = New Guid(CType(subKey.GetValue("CommandStateHandler"), String))
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
                                        If verbName Is Nothing Then verbName = subKeyName
                                        'verbName = subKeyName
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
                                        Dim h As HRESULT = commandState.GetState(Nothing, grfKeyState, state)
                                        Debug.WriteLine("GetState=" & h.ToString() & "--" & state.ToString())
                                        If state <> 100 Then
                                            _menuItems.Add(New MenuItemData() With {
                                                .Header = subKeyName & "-" & If(muiVerb, ""),
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
                    End If
                Next
            End Using

            'Dim provider As IExplorerCommandProvider = TryCast(folder.ShellFolder, IExplorerCommandProvider)
            'Dim enumerator As IEnumExplorerCommand = Nothing
            'If provider.GetCommands(IntPtr.Zero, GetType(IEnumExplorerCommand).GUID, enumerator) = HRESULT.S_OK Then
            '    Dim array As IShellItemArray = Nothing
            '    Dim pidls As List(Of Pidl) = items.Select(Function(i) i.Pidl).ToList()
            '    Functions.SHCreateShellItemArrayFromIDLists(pidls.Count, pidls.Select(Function(p) p.AbsolutePIDL).ToArray(), array)

            '    Dim grfKeyState As MK = 0
            '    If Keyboard.Modifiers.HasFlag(ModifierKeys.Shift) Then grfKeyState = grfKeyState Or MK.MK_SHIFT
            '    If Keyboard.Modifiers.HasFlag(ModifierKeys.Control) Then grfKeyState = grfKeyState Or MK.MK_CONTROL

            '    Dim command As IExplorerCommand = Nothing, fetched As UInt32
            '    Dim h As HRESULT = enumerator.Next(1, command, fetched)
            '    While h = HRESULT.S_OK AndAlso fetched > 0
            '        Dim state As EXPCMDSTATE = 0
            '        command.GetState(array, grfKeyState, state)
            '        Dim title As String = Nothing
            '        command.GetTitle(array, title)
            '        Debug.WriteLine(title & "  " & state.ToString())
            '        _menuItems.Add(New MenuItemData() With {
            '            .Header = "test"
            '        })
            '        h = enumerator.Next(1, command, fetched)
            '    End While
            'End If
            'Dim sb As New StringBuilder(256)
            'Dim hr = Functions.SHLoadIndirectString("@shell32.dll,-10564", sb, CUInt(sb.Capacity), IntPtr.Zero)
            'Dim t As Type = Type.GetTypeFromCLSID(New Guid("93a6a532-9396-45c2-9538-35a0a949e51a"))
            'Dim o As Object = Activator.CreateInstance(t)
            'Dim cm As IContextMenu = TryCast(o, IContextMenu)
        End Sub

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

        <ComVisible(True)>
        <ClassInterface(ClassInterfaceType.None)>
        Public Class MockShellBrowser
            Implements IShellBrowser

            Public Function GetWindow(ByRef hwnd As IntPtr) As Integer Implements IShellBrowser.GetWindow
                Throw New NotImplementedException()
            End Function

            Public Function ContextSensitiveHelp(fEnterMode As Integer) As Integer Implements IShellBrowser.ContextSensitiveHelp
                Throw New NotImplementedException()
            End Function

            Public Function InsertMenusSB(hmenuShared As IntPtr, lpMenuWidths As IntPtr) As Integer Implements IShellBrowser.InsertMenusSB
                Throw New NotImplementedException()
            End Function

            Public Function SetMenuSB(hmenuShared As IntPtr, holemenuRes As IntPtr, hwndActiveObject As IntPtr) As Integer Implements IShellBrowser.SetMenuSB
                Throw New NotImplementedException()
            End Function

            Public Function RemoveMenusSB(hmenuShared As IntPtr) As Integer Implements IShellBrowser.RemoveMenusSB
                Throw New NotImplementedException()
            End Function

            Public Function SetStatusTextSB(pszStatusText As IntPtr) As Integer Implements IShellBrowser.SetStatusTextSB
                Throw New NotImplementedException()
            End Function

            Public Function EnableModelessSB(fEnable As Boolean) As Integer Implements IShellBrowser.EnableModelessSB
                Throw New NotImplementedException()
            End Function

            Public Function TranslateAcceleratorSB(pmsg As IntPtr, wID As Short) As Integer Implements IShellBrowser.TranslateAcceleratorSB
                Throw New NotImplementedException()
            End Function

            Public Function BrowseObject(pidl As IntPtr, <MarshalAs(UnmanagedType.U4)> wFlags As UInteger) As Object Implements IShellBrowser.BrowseObject
                Throw New NotImplementedException()
            End Function

            Public Function GetViewStateStream(grfMode As UInteger, ppStrm As IntPtr) As Integer Implements IShellBrowser.GetViewStateStream
                Throw New NotImplementedException()
            End Function

            Public Function GetControlWindow(id As UInteger, ByRef phwnd As IntPtr) As Integer Implements IShellBrowser.GetControlWindow
                Throw New NotImplementedException()
            End Function

            Public Function SendControlMsg(id As UInteger, uMsg As UInteger, wParam As UInteger, lParam As UInteger, pret As IntPtr) As Integer Implements IShellBrowser.SendControlMsg
                Throw New NotImplementedException()
            End Function

            Public Function QueryActiveShellView(<MarshalAs(UnmanagedType.Interface)> ByRef ppshv As IShellView) As Integer Implements IShellBrowser.QueryActiveShellView
                Throw New NotImplementedException()
            End Function

            Public Function OnViewWindowActive(<MarshalAs(UnmanagedType.Interface)> pshv As IShellView) As Integer Implements IShellBrowser.OnViewWindowActive
                Throw New NotImplementedException()
            End Function

            Public Function SetToolbarItems(lpButtons As IntPtr, nButtons As UInteger, uFlags As UInteger) As Integer Implements IShellBrowser.SetToolbarItems
                Throw New NotImplementedException()
            End Function
        End Class

        Private Function getMenuItem(command As IExplorerCommand, array As IShellItemArray) As MenuItemData
            Dim grfKeyState As MK = 0
            If Keyboard.Modifiers.HasFlag(ModifierKeys.Shift) Then grfKeyState = grfKeyState Or MK.MK_SHIFT
            If Keyboard.Modifiers.HasFlag(ModifierKeys.Control) Then grfKeyState = grfKeyState Or MK.MK_CONTROL
            Dim state As EXPCMDSTATE = 0
            Dim h As HRESULT = command.GetState(array, True, state)
            Debug.WriteLine("GetState=" & h.ToString() & "--" & state.ToString())
            If state <> 100 Then ' = EXPCMDSTATE.ECS_ENABLED Then
                Dim title As String = Nothing
                command.GetTitle(array, title)
                If Not String.IsNullOrWhiteSpace(title) Then
                    Return New MenuItemData() With {
                        .Header = title
                    }
                Else
                    Return New MenuItemData() With {
                        .Header = ""
                    }
                End If
            Else
                Return Nothing
            End If
        End Function

        Protected Overrides Sub menuItem_Click(c As Control, e2 As EventArgs)
            Dim a As Action = c.Tag
            a()
            Me.IsOpen = False
        End Sub

        Public Overrides Function InvokeCommand(id As Tuple(Of Integer, String)) As Task
            Throw New NotImplementedException()
        End Function

        Protected Overrides Function AddItems() As Task
            For Each item In _menuItems
                Dim menuItem As New MenuItem() With {
                    .Header = item.Header,
                    .Icon = New Image() With {.Source = item.Icon},
                    .Tag = item.Tag
                }
                Items.Add(menuItem)
                AddHandler menuItem.Click, AddressOf menuItem_Click
            Next
            Return Task.CompletedTask
        End Function

        Protected Overrides Function DoRenameAfter(Tag As Tuple(Of Integer, String)) As Boolean
            Throw New NotImplementedException()
        End Function
    End Class
End Namespace