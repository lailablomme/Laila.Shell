'Imports System.IO
'Imports System.Reflection
'Imports System.Runtime.InteropServices
'Imports System.Text
'Imports System.Windows.Controls
'Imports System.Windows.Input
'Imports Laila.Shell.Helpers
'Imports Laila.Shell.Interop
'Imports Laila.Shell.Interop.ContextMenu
'Imports Laila.Shell.Interop.DragDrop
'Imports Laila.Shell.Interop.Folders
'Imports Laila.Shell.Interop.Functions
'Imports Laila.Shell.Interop.Items
'Imports Laila.Shell.Interop.Properties
'Imports Laila.Shell.Interop.Windows
'Imports Laila.Shell.WinRT.Interface
'Imports Microsoft.Win32
'Imports SHDocVw

'Namespace Controls
'    Public Class ExplorerMenu
'        Inherits BaseMenu

'        Private Const COMMANDSTORE_PATH As String = "Software\Microsoft\Windows\CurrentVersion\Explorer\CommandStore\shell"

'        Private _menuItems As List(Of MenuItemData) = New List(Of MenuItemData)()

'        Protected Overrides Sub Make(folder As Folder, items As IEnumerable(Of Item), isDefaultOnly As Boolean)
'            Shell.GlobalThreadPool.Run(
'                Sub()
'                    ' get cloud verbs
'                    Dim assembly As Assembly = Assembly.LoadFrom("Laila.Shell.WinRT.dll")
'                    Dim type As Type = assembly.GetType("Laila.Shell.WinRT.ExplorerMenuHelper")
'                    Dim methodInfo As MethodInfo = type.GetMethod("GetCloudVerbs")
'                    Dim instance As Object = Activator.CreateInstance(type)
'                    Dim t As Task(Of Tuple(Of String, String, List(Of ExplorerCommandVerbInfo))) =
'                        methodInfo.Invoke(instance, {folder.FullPath})
'                    Dim result As Tuple(Of String, String, List(Of ExplorerCommandVerbInfo)) = t.Result

'                    ' add cloud menu
'                    Dim cloudMenuItem As New MenuItemData() With {
'                        .Header = result.Item1
'                    }
'                    If Not String.IsNullOrWhiteSpace(result.Item2) Then
'                        cloudMenuItem.Icon = ImageHelper.ExtractIcon(result.Item2)
'                    End If
'                    _menuItems.Add(cloudMenuItem)

'                    ' add commands
'                    Dim array As IShellItemArray = Nothing
'                    Dim pidls As List(Of Pidl) = items.Select(Function(i) i.Pidl.Clone()).ToList()
'                    Functions.SHCreateShellItemArrayFromIDLists(pidls.Count, pidls.Select(Function(p) p.AbsolutePIDL).ToArray(), array)
'                    For Each item In result.Item3
'                        Helpers.ComHelper.RunWithComObject(item.ComServerPath, item.ClsId, GetType(IExplorerCommand).GUID,
'                    Sub(obj As Object)
'                        Dim ows As IInitializeWithItem = TryCast(obj, IInitializeWithItem)
'                        _menuItems.Add(getMenuItem(obj, array))
'                    End Sub)
'                    Next
'                End Sub)
'        End Sub

'        <ComImport>
'        <Guid("00000114-0000-0000-C000-000000000046")>
'        <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
'        Public Interface IOleWindow
'            <PreserveSig>
'            Function GetWindow(<Out> ByRef phwnd As IntPtr) As Integer

'            <PreserveSig>
'            Function ContextSensitiveHelp(<[In]> fEnterMode As Boolean) As Integer
'        End Interface
'        <ComImport>
'        <Guid("831D1DFF-7F57-4720-87E4-CB57D6214428")>
'        <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
'        Public Interface IInvocationLocation
'            ' HRESULT Proc3([out] int* p0);
'            <PreserveSig>
'            Function GetInvocationType(<Out> ByRef result As Integer) As Integer

'            ' HRESULT Proc4([in] int p0);
'            <PreserveSig>
'            Function SetInvocationType(<[In]> value As Integer) As Integer
'        End Interface


'        <ComVisible(True), Guid("a985d29a-81ef-41b2-8440-457c38ef959c"), ProgId("Laila.Shell.Helpers.WpfDragTargetProxyc")>
'        <ClassInterface(ClassInterfaceType.None)>
'        Public Class MockIL
'            Implements IInvocationLocation

'            Public Function GetInvocationType(<Out> ByRef result As Integer) As Integer Implements IInvocationLocation.GetInvocationType
'                Return 1
'            End Function

'            Public Function SetInvocationType(<[In]> value As Integer) As Integer Implements IInvocationLocation.SetInvocationType
'                'Throw New NotImplementedException()
'            End Function
'        End Class

'        <ComVisible(True), Guid("a985d29a-81ef-41b2-8440-457c38ef959c"), ProgId("Laila.Shell.Helpers.WpfDragTargetProxyc")>
'        <ClassInterface(ClassInterfaceType.None)>
'        Public Class Mockow
'            Implements IOleWindow

'            Public Function GetWindow(<Out> ByRef phwnd As IntPtr) As Integer Implements IOleWindow.GetWindow
'                phwnd = Shell._hwnd
'            End Function

'            Public Function ContextSensitiveHelp(<[In]> fEnterMode As Boolean) As Integer Implements IOleWindow.ContextSensitiveHelp
'                ' Throw New NotImplementedException()
'            End Function
'        End Class


'        <ComVisible(True), Guid("a985d29a-81ef-41b2-8440-457c38ef959c"), ProgId("Laila.Shell.Helpers.WpfDragTargetProxyc")>
'        <ClassInterface(ClassInterfaceType.None)>
'        Public Class MockShellSite
'            Implements Interop.Items.IServiceProvider

'            Public Function QueryService(ByRef guidService As Guid, ByRef riid As Guid, ByRef ppvObject As IntPtr) As Integer Implements IServiceProvider.QueryService
'                ' You can return your own IShellItemArray, IShellView, etc. here
'                If riid = New Guid("831D1DFF-7F57-4720-87E4-CB57D6214428") Then
'                    ppvObject = Marshal.GetComInterfaceForObject(New MockIL(), GetType(IInvocationLocation))
'                    Debug.WriteLine("QueryServiceOK " & guidService.ToString() & " " & riid.ToString())
'                    Return HRESULT.S_OK
'                End If
'                If riid = New Guid("00000114-0000-0000-C000-000000000046") Then
'                    ppvObject = Marshal.GetComInterfaceForObject(New Mockow(), GetType(IOleWindow))
'                    Debug.WriteLine("QueryServiceOK " & guidService.ToString() & " " & riid.ToString())
'                    Return HRESULT.S_OK
'                End If
'                'Dim shellWindows As New ShellWindows()
'                'Dim sp As IServiceProvider = Nothing
'                '' Loop through all windows
'                'For i As Integer = 0 To shellWindows.Count - 1
'                '    Dim window As Object = shellWindows.Item(i)

'                '    ' Try to cast to IServiceProvider
'                '    sp = TryCast(window, IServiceProvider)
'                '    If sp IsNot Nothing Then
'                '        Exit For
'                '    End If
'                'Next
'                'Return sp.QueryService(guidService, riid, ppvObject)
'                ppvObject = IntPtr.Zero
'                Debug.WriteLine("QueryService " & guidService.ToString() & " " & riid.ToString())
'                Return HRESULT.E_NOINTERFACE
'            End Function
'        End Class

'        <ComVisible(True)>
'        <ClassInterface(ClassInterfaceType.None)>
'        Public Class MockShellBrowser
'            Implements IShellBrowser

'            Public Function GetWindow(ByRef hwnd As IntPtr) As Integer Implements IShellBrowser.GetWindow
'                Throw New NotImplementedException()
'            End Function

'            Public Function ContextSensitiveHelp(fEnterMode As Integer) As Integer Implements IShellBrowser.ContextSensitiveHelp
'                Throw New NotImplementedException()
'            End Function

'            Public Function InsertMenusSB(hmenuShared As IntPtr, lpMenuWidths As IntPtr) As Integer Implements IShellBrowser.InsertMenusSB
'                Throw New NotImplementedException()
'            End Function

'            Public Function SetMenuSB(hmenuShared As IntPtr, holemenuRes As IntPtr, hwndActiveObject As IntPtr) As Integer Implements IShellBrowser.SetMenuSB
'                Throw New NotImplementedException()
'            End Function

'            Public Function RemoveMenusSB(hmenuShared As IntPtr) As Integer Implements IShellBrowser.RemoveMenusSB
'                Throw New NotImplementedException()
'            End Function

'            Public Function SetStatusTextSB(pszStatusText As IntPtr) As Integer Implements IShellBrowser.SetStatusTextSB
'                Throw New NotImplementedException()
'            End Function

'            Public Function EnableModelessSB(fEnable As Boolean) As Integer Implements IShellBrowser.EnableModelessSB
'                Throw New NotImplementedException()
'            End Function

'            Public Function TranslateAcceleratorSB(pmsg As IntPtr, wID As Short) As Integer Implements IShellBrowser.TranslateAcceleratorSB
'                Throw New NotImplementedException()
'            End Function

'            Public Function BrowseObject(pidl As IntPtr, <MarshalAs(UnmanagedType.U4)> wFlags As UInteger) As Object Implements IShellBrowser.BrowseObject
'                Throw New NotImplementedException()
'            End Function

'            Public Function GetViewStateStream(grfMode As UInteger, ppStrm As IntPtr) As Integer Implements IShellBrowser.GetViewStateStream
'                Throw New NotImplementedException()
'            End Function

'            Public Function GetControlWindow(id As UInteger, ByRef phwnd As IntPtr) As Integer Implements IShellBrowser.GetControlWindow
'                Throw New NotImplementedException()
'            End Function

'            Public Function SendControlMsg(id As UInteger, uMsg As UInteger, wParam As UInteger, lParam As UInteger, pret As IntPtr) As Integer Implements IShellBrowser.SendControlMsg
'                Throw New NotImplementedException()
'            End Function

'            Public Function QueryActiveShellView(<MarshalAs(UnmanagedType.Interface)> ByRef ppshv As IShellView) As Integer Implements IShellBrowser.QueryActiveShellView
'                Throw New NotImplementedException()
'            End Function

'            Public Function OnViewWindowActive(<MarshalAs(UnmanagedType.Interface)> pshv As IShellView) As Integer Implements IShellBrowser.OnViewWindowActive
'                Throw New NotImplementedException()
'            End Function

'            Public Function SetToolbarItems(lpButtons As IntPtr, nButtons As UInteger, uFlags As UInteger) As Integer Implements IShellBrowser.SetToolbarItems
'                Throw New NotImplementedException()
'            End Function
'        End Class

'        Private Function getMenuItem(command As IExplorerCommand, array As IShellItemArray) As MenuItemData
'            Dim grfKeyState As MK = 0
'            If Keyboard.Modifiers.HasFlag(ModifierKeys.Shift) Then grfKeyState = grfKeyState Or MK.MK_SHIFT
'            If Keyboard.Modifiers.HasFlag(ModifierKeys.Control) Then grfKeyState = grfKeyState Or MK.MK_CONTROL
'            Dim state As EXPCMDSTATE = 0
'            Dim h As HRESULT = command.GetState(array, True, state)
'            Dim ws As IObjectWithSelection = TryCast(command, IObjectWithSelection)
'            Debug.WriteLine("GetState=" & h.ToString() & "--" & state.ToString())
'            If state <> 100 Then ' = EXPCMDSTATE.ECS_ENABLED Then
'                Dim ieec As IEnumExplorerCommand = Nothing
'                command.EnumSubCommands(ieec)
'                Dim f As UInt32
'                command.GetFlags(f)
'                'If ieec IsNot Nothing Then
'                '    ieec.Next(1, command, f)
'                'End If
'                Dim title As String = Nothing
'                If Not command Is Nothing Then command.GetTitle(array, title)
'                If Not String.IsNullOrWhiteSpace(title) Then
'                    Return New MenuItemData() With {
'                        .Header = title & f
'                    }
'                Else
'                    Return New MenuItemData() With {
'                        .Header = "" & state.ToString() & f
'                    }
'                End If
'            Else
'                Return Nothing
'            End If
'        End Function

'        Protected Overrides Sub menuItem_Click(c As Control, e2 As EventArgs)
'            Dim a As Action = c.Tag
'            a()
'            Me.IsOpen = False
'        End Sub

'        Public Overrides Function InvokeCommand(id As Tuple(Of Integer, String, Object)) As Task
'            Throw New NotImplementedException()
'        End Function

'        Protected Overrides Function AddItems() As Task
'            For Each item In _menuItems
'                Dim menuItem As New MenuItem() With {
'                    .Header = item.Header,
'                    .Icon = New Image() With {.Source = item.Icon},
'                    .Tag = item.Tag
'                }
'                Items.Add(menuItem)
'                AddHandler menuItem.Click, AddressOf menuItem_Click
'            Next
'            Return Task.CompletedTask
'        End Function

'        Protected Overrides Function DoRenameAfter(Tag As Tuple(Of Integer, String, Object)) As Boolean
'            Throw New NotImplementedException()
'        End Function
'    End Class
'End Namespace