Imports System.IO
Imports System.Text
Imports System.Windows.Controls
Imports System.Windows.Input
Imports Laila.Shell.Interop
Imports Laila.Shell.Interop.ContextMenu
Imports Laila.Shell.Interop.Items
Imports Laila.Shell.Interop.Properties
Imports Laila.Shell.Interop.Windows
Imports Microsoft.Win32

Namespace Controls
    Public Class ExplorerMenu
        Inherits BaseMenu

        Private Const COMMANDSTORE_PATH As String = "Software\Microsoft\Windows\CurrentVersion\Explorer\CommandStore\shell"

        Private _menuItems As List(Of MenuItemData) = New List(Of MenuItemData)()

        Protected Overrides Sub Make(folder As Folder, items As IEnumerable(Of Item), isDefaultOnly As Boolean)
            Dim array As IShellItemArray = Nothing
            Dim pidls As List(Of Pidl) = items.Select(Function(i) i.Pidl).ToList()
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
                        AndAlso Not subKeyName = "Windows.SharePrivate" Then
                        Using subKey As RegistryKey = commandStoreKey.OpenSubKey(subKeyName)
                            If subKey IsNot Nothing Then
                                Dim handlerGUID As Guid
                                If Not subKey.GetValue("ExplorerCommandHandler") Is Nothing Then
                                    Try
                                        handlerGUID = New Guid(CType(subKey.GetValue("ExplorerCommandHandler"), String))
                                        Dim commandType As Type = Type.GetTypeFromCLSID(handlerGUID)
                                        Dim command As IExplorerCommand = TryCast(Activator.CreateInstance(commandType), IExplorerCommand)
                                        If Not command Is Nothing Then
                                            Dim menuItem As MenuItemData = getMenuItem(command, array)
                                            If Not menuItem Is Nothing Then
                                                menuItem.Header = subKeyName & "-" & menuItem.Header
                                                _menuItems.Add(menuItem)
                                            End If
                                        End If
                                    Catch ex As Exception
                                    End Try
                                ElseIf Not subKey.GetValue("CommandStateHandler") Is Nothing AndAlso Not subKey.GetValue("MUIVerb") Is Nothing Then
                                    handlerGUID = New Guid(CType(subKey.GetValue("CommandStateHandler"), String))
                                    Dim muiVerb As String = CType(subKey.GetValue("MUIVerb"), String)
                                    If muiVerb.StartsWith("@") Then
                                        Dim sb As New StringBuilder(256)
                                        Dim h = Functions.SHLoadIndirectString(muiVerb, sb, CUInt(sb.Capacity), IntPtr.Zero)
                                        muiVerb = sb.ToString()
                                    End If
                                    Dim commandType As Type = Type.GetTypeFromCLSID(handlerGUID)
                                    Dim obj As Object = Activator.CreateInstance(commandType)
                                    Dim command As IExplorerCommandState = TryCast(obj, IExplorerCommandState)
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
                                    If Not command Is Nothing Then
                                        Dim state As EXPCMDSTATE = 0
                                        Dim h As HRESULT = command.GetState(Nothing, grfKeyState, state)
                                        Debug.WriteLine("GetState=" & h.ToString() & "--" & state.ToString())
                                        If state = EXPCMDSTATE.ECS_ENABLED Then
                                            _menuItems.Add(New MenuItemData() With {
                                                .Header = subKeyName & "-" & muiVerb,
                                                .Tag = handlerGUID
                                            })
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

        Private Function getMenuItem(command As IExplorerCommand, array As IShellItemArray) As MenuItemData
            Dim grfKeyState As MK = 0
            If Keyboard.Modifiers.HasFlag(ModifierKeys.Shift) Then grfKeyState = grfKeyState Or MK.MK_SHIFT
            If Keyboard.Modifiers.HasFlag(ModifierKeys.Control) Then grfKeyState = grfKeyState Or MK.MK_CONTROL
            Dim state As EXPCMDSTATE = 0
            command.GetState(array, grfKeyState, state)
            Debug.WriteLine("GetState=" & state.ToString())
            If state = EXPCMDSTATE.ECS_ENABLED Then
                Dim title As String = Nothing
                command.GetTitle(array, title)
                If Not String.IsNullOrWhiteSpace(title) Then
                    Return New MenuItemData() With {
                        .Header = title
                    }
                Else
                    Return Nothing
                End If
            Else
                Return Nothing
            End If
        End Function

        Public Overrides Function InvokeCommand(id As Tuple(Of Integer, String)) As Task
            Throw New NotImplementedException()
        End Function

        Protected Overrides Function AddItems() As Task
            For Each item In _menuItems
                Dim menuItem As New MenuItem() With {
                    .Header = item.Header,
                    .Tag = item.Tag
                }
                Items.Add(menuItem)
            Next
            Return Task.CompletedTask
        End Function

        Protected Overrides Function DoRenameAfter(Tag As Tuple(Of Integer, String)) As Boolean
            Throw New NotImplementedException()
        End Function
    End Class
End Namespace