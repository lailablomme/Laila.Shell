'Imports System.Runtime.InteropServices
'Imports Laila.Shell.Controls.FolderView

'Public Class Snippets
'    If CType(e.NewValue, Folder).FullPath.Contains("G:") Then '::{8E908FC9-BECC-40F6-915B-F4CA0E70D03D}") Then
'    Dim ww As System.Windows.Forms.Form = New System.Windows.Forms.Form()
'                    ww.Show()
'                    Dim ptr As IntPtr, shellFolder As IShellFolder, shellView As IShellView
'    Dim h As HRESULT
'                    shellFolder = CType(e.NewValue, Folder).ShellFolder
'                    Try
'                        h = shellFolder.CreateViewObject(IntPtr.Zero, GetType(IShellView).GUID, ptr)
'                        shellView = Marshal.GetTypedObjectForIUnknown(ptr, GetType(IShellView))
'                    Finally
'    If Not IntPtr.Zero.Equals(ptr) Then
'                            Marshal.Release(ptr)
'                        End If
'    End Try
'    Dim sb As ShellBrowser = New ShellBrowser(ww.Handle)
'    Dim fs As FOLDERSETTINGS
'                    fs.fFlags = 0
'                    fs.ViewMode = -1
'                    Dim r As WIN32RECT = New WIN32RECT()
'                    r.Top = 20 : r.Left = 20 : r.Right = 500 : r.Bottom = 500
'                    shellView.CreateViewWindow(Nothing, fs, sb, r, ww.Handle)
'                    shellView.UIActivate(1)

'                    'Dim obj As IExplorerHost
'                    'h 'h = Functions.CoCreateInstance(New Guid("5BD95610-9434-43C2-886C-57852CC8A120"), IntPtr.Zero, 4 Or 1, GetType(IExplorerHost).GUID, obj)
'                    '    'Dim c As SFV_CREATE
'                    '    shellFolder.BindToObject(CType(e.NewValue, Folder).Pidl.RelativePIDL, IntPtr.Zero, New Guid("93CB110F-9189-4349-BD9F-392D9A4D0096"), ptr)
'                    ' = Activator.CreateInstance(Type.GetTypeFromCLSID(New Guid("5BD95610-9434-43C2-886C-57852CC8A120")))
'                    '    Dim uiel As IDUIElementProviderInit
'                    '    uiel = TryCast(obj, IDUIElementProviderInit)
'                    '    If Not uiel Is Nothing Then
'                    '        Dim i As Integer
'                    '        hh = uiel.Proc3(ww.Handle)
'                    '        Debug.WriteLine("proc3" & hh.ToString() & "  " & i.ToString())
'                    '        hh = uiel.Proc4(i)
'                    '        Debug.WriteLine("proc4" & hh.ToString() & "  " & i.ToString())
'                    '    End If
'                    'End If
'                    'c.cbSize = Marshal.SizeOf(Of SFV_CREATE)
'                    'c.pshf = shellFolder
'                    '' c.
'                    'shellFolder = CType(e.NewValue, Folder).ShellFolder
'                    ''    Dim shellView As IShellView
'                    '' Dim pf As IPersist
'                    ''Dim propertyBagPtr As IntPtr, propertyBag As IPropertyBag
'                    ''Functions.PSCreateMemoryPropertyStore(GetType(IPropertyBag).GUID, propertyBagPtr)
'                    ''propertyBag = Marshal.GetTypedObjectForIUnknown(propertyBagPtr, GetType(IPropertyBag))
'                    ''Dim var As PROPVARIANT
'                    ''var.vt = VarEnum.VT_UI4
'                    ''var.union.intVal = &H64
'                    ''Dim var2 As PROPVARIANT
'                    ''var2.vt = VarEnum.VT_LPWSTR
'                    ''var2.union.ptr = Marshal.StringToCoTaskMemUni("C:\Windows\System32\netcenter.dll")
'                    ''var2.SetValue("C:\Windows\System32\netcenter.dll")
'                    ''propertyBag.Write("ResourceDLL", var2) '  STR_ENUM_ITEMS_FLAGS 
'                    ''propertyBag.Write("ResourceID", var) '  STR_ENUM_ITEMS_FLAGS 
'                    ''Dim ipro As IPersistPropertyBag = shellFolder
'                    '''  ipro.InitNew()
'                    '''hh = ipro.Save(propertyBag, True, True)
'                    ''hh = ipro.Load(propertyBag, Nothing)
'                    '''ipro.InitNew()
'                    ''hh = ipro.Load(propertyBag, Nothing)
'                    ''Dim oo As UInt32
'                    ''propertyBag.Read("ResourceID", oo, Nothing)
'                    'Dim ww As System.Windows.Forms.Form = New Forms.Form()
'                    ''''Dim pan As System.Windows.Forms.Panel = New System.Windows.Forms.Panel()
'                    ''''wfh.Child = pan
'                    ''''pan.BackColor = System.Drawing.Color.Red
'                    '''' w.Show()
'                    'ww.Show()
'                    'Dim r As WIN32RECT, hwd As IntPtr = ww.Handle, xx As SW, btc As btc = New btc()
'                    'r.Top = 10 : r.Left = 10 : r.Right = 1000 : r.Bottom = 500
'                    '''Functions.ShowWindow()
'                    'Dim iii = 2147483320
'                    'While iii > 2147483320 - 10
'                    '    hh = obj.Proc3(CType(e.NewValue, Folder).Pidl.AbsolutePIDL, 0, r, iii, hwd, IntPtr.Zero, Nothing) ' btc)
'                    '    ''  hh = obj.Proc4(CType(e.NewValue, Folder).Pidl.AbsolutePIDL, &HFFFFFFFF, r, SW.SHOW)
'                    '    Debug.WriteLine(iii & "   hh=" & hh.ToString())
'                    '    iii -= 1
'                    'End While
'                    'obj.Initialize()
'                    'obj.CreateViewWindow(r, ww.Handle)
'                    'Dim wb As System.Windows.Forms.WebBrowser = New Forms.WebBrowser()
'                    'ww.Controls.Add(wb)
'                    'wb.Navigate("shell:::{8E908FC9-BECC-40F6-915B-F4CA0E70D03D}")

'                    'Dim eb As IExplorerBrowser = Activator.CreateInstance(Type.GetTypeFromCLSID(Guids.CLSID_ExplorerBrowser))
'                    'Dim fs As FOLDERSETTINGS
'                    'fs.ViewMode = -1
'                    'Dim r As WIN32RECT
'                    'r.Top = 20 : r.Left = 20 : r.Right = 500 : r.Bottom = 500
'                    'eb.Initialize(ww.Handle, r, fs)
'                    ''eb.BrowseToIDList(CType(e.NewValue, Folder).Pidl.AbsolutePIDL, SBSP.ABSOLUTE Or SBSP.OPENMODE Or SBSP.TRUSTEDFORACTIVEX)
'                    'eb.BrowseToObject(Shell.Desktop.ShellItem2, 0)
'                    'Dim ty As Type = Type.GetTypeFromCLSID(New Guid("8E908FC9-BECC-40F6-915B-F4CA0E70D03D"))
'                    'Dim sf22 = CTypeDynamic(shellFolder, ty)

'                    '''hh = Functions.SHCreateShellFolderView(c, shellView)
'                    'Dim g As Guid = New Guid() ' "137E7700-3573-11CF-AE69-08002B2E1262")
'                    'If hh = HRESULT.S_OK AndAlso Not shellView Is Nothing Then
'                    '    For x = Int32.MinValue To Int32.MaxValue
'                    '        hh = shellView.GetView(g, x)
'                    '        If hh = HRESULT.S_OK Then
'                    '            Debug.WriteLine("GetView returned " & hh.ToString()) ' & "   & Guid=" & g.ToString())
'                    '        End If
'                    '    Next
'                    'End If
'                    ''Dim r As WIN32RECT = New WIN32RECT()
'                    ''r.Top = 20 : r.Left = 20 : r.Right = 500 : r.Bottom = 500
'                    ''Dim fs As FOLDERSETTINGS = New FOLDERSETTINGS()
'                    ''fs.ViewMode = -1
'                    ''Dim hwnd As IntPtr
'                    '''Dim w As Window = New Window()
'                    '''Dim wfh As WindowsFormsHost = New WindowsFormsHost()

'                    '''w.Content = wfh
'                    ''Dim g As Guid '= New Guid("88E39E80-3578-11CF-AE69-08002B2E1262")

'                    'Dim sb As ShellBrowser = New ShellBrowser(ww.Handle)
'                    'Dim sv2 As IShellView2 = shellView
'                    ''Dim h As HRESULT, iii As IntPtr
'                    '''iii = Marshal.AllocHGlobal(1024)
'                    '''' For xx As UInt32 = 0 To UInt32.MaxValue
'                    '''h = sv2.GetView(iii, -1)
'                    ''''  If h = HRESULT.S_OK Then
'                    '''Debug.WriteLine("GetView returned " & h.ToString() & "   & guid=" & iii.ToString())
'                    '''h = sv2.GetView(iii, -2)
'                    ''''  If h = HRESULT.S_OK Then
'                    '''Debug.WriteLine("GetView returned " & h.ToString() & "   & guid=" & iii.ToString())
'                    '''      End If
'                    ''' Next
'                    ''Dim sv2pa As SV2CVW2_PARAMS
'                    ''sv2pa.hwndView = ww.Handle
'                    ''sv2pa.prcView = r
'                    ''sv2pa.cbSize = Marshal.SizeOf(Of SV2CVW2_PARAMS)
'                    ''sv2pa.pfs = fs
'                    ''sv2pa.psbOwner = sb
'                    ''sv2pa.pvid = g
'                    ''For ii = 1 To 1
'                    ''    Try
'                    ''        ' shellFolder = CType(e.NewValue, Folder).ShellFolder
'                    ''        'hh = shellFolder.BindToObject(CType(e.NewValue, Folder).Pidl.RelativePIDL, 0, GetType(IShellView).GUID, ptr)
'                    ''        hh = shellFolder.CreateViewObject(ww.Handle, GetType(IShellView).GUID, ptr)
'                    ''        shellView = Marshal.GetTypedObjectForIUnknown(ptr, GetType(IShellView))
'                    ''    Finally
'                    ''        If Not IntPtr.Zero.Equals(ptr) Then
'                    ''            Marshal.Release(ptr)
'                    ''        End If
'                    ''    End Try
'                    'sv2 = shellView
'                    'hh = sv2.CreateViewWindow3(sb, Nothing, 0, 0, 0, -1, IntPtr.Zero, r, ww.Handle)
'                    'shellView.UIActivate(1)
'                    '    System.Windows.MessageBox.Show("test")
'                    '    sv2.DestroyViewWindow()

'                    'Next
'                    '' Dim h As HRESULT = shellView.CreateViewWindow(Nothing, fs, sb, r, hwnd)
'                    'Debug.WriteLine("CreateViewWindow returned " & h.ToString() & "   & hwnd=" & hwnd.ToString())
'                    '' Functions.SetParent(hwnd, ww.Handle)
'                    '' Functions.MoveWindow(hwnd, 0, 0, ww.Width, ww.Height, True)

'                    ''Dim ocp As IOpenControlPanel = Activator.CreateInstance(Type.GetTypeFromCLSID(New Guid("06622D85-6856-4460-8DE1-A81921B41C4B")))
'                    ''Dim path As StringBuilder = New StringBuilder()
'                    ''path.Append(New String("x", 1024))
'                    ''ocp.GetPath("{8E908FC9-BECC-40F6-915B-F4CA0E70D03D}", path, 1024)
'                    ''Debug.WriteLine("path=" & path.ToString())
'                End If
'    <ComImport, Guid("B196B283-BAB4-101A-B69C-00AA00341D07"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
'    Public Interface IProvideClassInfo
'        Sub GetClassInfo(<Out, MarshalAs(UnmanagedType.Interface)> ByRef typeInfo As Object)
'    End Interface

'    <ComImport, InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("FF9CE68E-060D-4120-81F9-F3061C703225")>
'    Private Interface IControlPanelProcessExplorerHost
'        Sub Initialize()
'        Sub CreateViewWindow(ByRef prcView As WIN32RECT, phWnd As IntPtr)
'    End Interface

'    <ComImport, Guid("00020400-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
'    Public Interface IDispatch
'        Sub GetTypeInfoCount(ByRef pctinfo As UInteger)
'        Sub GetTypeInfo(<MarshalAs(UnmanagedType.U4)> iTInfo As UInteger, <MarshalAs(UnmanagedType.U4)> lcid As UInteger, <Out, MarshalAs(UnmanagedType.Interface)> ByRef ppTInfo As Object)
'        Sub GetIDsOfNames(ByRef riid As Guid, <MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.LPWStr)> rgszNames As String(), cNames As UInteger, lcid As UInteger, <MarshalAs(UnmanagedType.LPArray)> rgDispId As Integer())
'        Sub Invoke(dispIdMember As Integer, ByRef riid As Guid, lcid As UInteger, wFlags As UShort, ByRef pDispParams As Object, ByRef pVarResult As Object, ByRef pExcepInfo As Object, ByRef puArgErr As Object)
'    End Interface
'    <ComImport>
'    <Guid("00000114-0000-0000-C000-000000000046")>
'    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
'    Public Interface IOleWindow
'        ''' <summary>
'        ''' Retrieves a handle to one of the windows participating in in-place activation (frame or document window).
'        ''' </summary>
'        ''' <param name="phwnd">When this method returns, contains the window handle.</param>
'        <PreserveSig>
'        Function GetWindow(<Out> ByRef phwnd As IntPtr) As Integer

'        ''' <summary>
'        ''' Determines whether context-sensitive help mode should be entered during an in-place activation session.
'        ''' </summary>
'        ''' <param name="fEnterMode">True to enter context-sensitive help mode, False to exit it.</param>
'        <PreserveSig>
'        Function ContextSensitiveHelp(<[In]> fEnterMode As Boolean) As Integer
'    End Interface
'    <ComImport(), InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("FF9CE68E-060D-4120-81F9-F3061C703225")>
'    Public Interface IExplorerHost
'        <PreserveSig()>
'        Function Proc3(
'        <[In]()> ByVal p0 As IntPtr, ' FC_USER_MARSHAL* (probably marshaled data)
'        <[In]()> ByVal p1 As Integer,
'        <[In]()> ByVal p2 As WIN32RECT, ' Custom struct, you'd need to define it
'        <[In]()> ByVal p3 As SW,
'        <[In]()> ByVal p4 As IntPtr, ' HWND
'        <[In]()> ByVal p5 As IntPtr, ' IUnknown*
'        <[In]()> ByVal p6 As IBrowserThreadHandshake) As HRESULT  ' IBrowserThreadHandshake*

'        <PreserveSig()>
'        Function Proc4(
'        <[In]()> ByVal p0 As IntPtr, ' FC_USER_MARSHAL* (probably marshaled data)
'        <[In]()> ByVal p1 As Integer,
'        <[In]()> ByVal p2 As WIN32RECT, ' Struct_94* (pointer to another custom structure)
'        <[In]()> ByVal p3 As SW) As Integer
'    End Interface
'    <ComImport(), InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("E1F5EC9F-F933-492B-A242-C3CDAC6ECFBD")>
'    Public Interface IDUIElementProviderInit
'        Function Proc3(ByVal p0 As Integer) As Integer 'HRESULT is translated to Integer in VB.NET
'        Function Proc4(ByRef p0 As Integer) As Integer 'HRESULT is translated to Integer in VB.NET
'    End Interface

'    <ComImport(), InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("6ACA2690-0DE1-46C9-A1C5-9338A22BF12B")>
'    Public Interface IBrowserThreadHandshake
'        <PreserveSig>
'        Function Proc3(<Out> ByRef p0 As IntPtr) As Integer

'        <PreserveSig>
'        Function Proc4(<Out> ByRef p0 As ComTypes.IStream) As Integer

'        <PreserveSig>
'        Function Proc5(<[In]> p0 As Integer) As Integer

'        <PreserveSig>
'        Function Proc6() As Integer

'        <PreserveSig>
'        Function Proc7() As Integer
'    End Interface


'    Public Class btc
'        Implements IBrowserThreadHandshake

'        Public Function Proc3(<Out> ByRef p0 As IntPtr) As Integer Implements IBrowserThreadHandshake.Proc3
'            Return HRESULT.S_OK
'        End Function

'        Public Function Proc4(<Out> ByRef p0 As ComTypes.IStream) As Integer Implements IBrowserThreadHandshake.Proc4
'            ' Create a MemoryStream or other Stream implementation
'            'Dim memoryStream As New MemoryStream()

'            '' Write some sample data to the stream (optional)
'            'Dim data As Byte() = Encoding.UTF8.GetBytes("Sample Data")
'            'memoryStream.Write(data, 0, data.Length)

'            '' Reset the position to the start of the stream before returning
'            'memoryStream.Seek(0, SeekOrigin.Begin)

'            ' Marshal the MemoryStream to IStream
'            Dim iStream As IntPtr, ptr As IntPtr = Marshal.StringToHGlobalAnsi(Shell.Desktop.FullPath) '= CType(Marshal.GetIUnknownForObject(memoryStream), ComTypes.IStream)
'            'Marshal.Copy(Shell.Desktop.Pidl.Bytes, 0, ptr, Shell.Desktop.Pidl.Bytes.Length)
'            iStream = Functions.SHCreateMemStream(ptr, Shell.Desktop.FullPath.Length)
'            'iStream = Marshal.GetObjectForIUnknown((Marshal.GetComInterfaceForObject(memoryStream, GetType(ComTypes.IStream)))) ', ComTypes.IStream)

'            ' Return the IStream (via pointer to IStream)
'            ' Assuming the output parameter is a pointer to IStream, return the pointer to the caller
'            p0 = Marshal.GetObjectForIUnknown(iStream)
'            p0.Seek(Shell.Desktop.FullPath.Length, 0, IntPtr.Zero)

'            Return 0 ' Success (HRESULT S_OK)
'        End Function

'        Public Function Proc5(<[In]> p0 As Integer) As Integer Implements IBrowserThreadHandshake.Proc5
'            Debug.WriteLine("got " & p0)
'            Return HRESULT.S_OK
'        End Function

'        Public Function Proc6() As Integer Implements IBrowserThreadHandshake.Proc6
'            Return HRESULT.S_OK
'        End Function

'        Public Function Proc7() As Integer Implements IBrowserThreadHandshake.Proc7
'            Return HRESULT.S_OK
'        End Function
'    End Class


'    Public Class ShellBrowser
'        Implements IShellBrowser
'        Private _hp As IntPtr
'        Public Sub New(hwndParent As IntPtr)
'            _hp = hwndParent
'        End Sub
'        Public Function GetWindow(ByRef hwnd As IntPtr) As Integer Implements IShellBrowser.GetWindow
'            hwnd = _hp
'            Return HRESULT.S_OK
'        End Function

'        Public Function ContextSensitiveHelp(fEnterMode As Integer) As Integer Implements IShellBrowser.ContextSensitiveHelp
'            Throw New NotImplementedException()
'        End Function

'        Public Function InsertMenusSB(hmenuShared As IntPtr, lpMenuWidths As IntPtr) As Integer Implements IShellBrowser.InsertMenusSB
'            'Throw New NotImplementedException()
'            Return HRESULT.E_NOTIMPL
'        End Function

'        Public Function SetMenuSB(hmenuShared As IntPtr, holemenuRes As IntPtr, hwndActiveObject As IntPtr) As Integer Implements IShellBrowser.SetMenuSB
'            'Throw New NotImplementedException()
'            Return HRESULT.E_NOTIMPL
'        End Function

'        Public Function RemoveMenusSB(hmenuShared As IntPtr) As Integer Implements IShellBrowser.RemoveMenusSB
'            'Throw New NotImplementedException()
'            Return HRESULT.E_NOTIMPL
'        End Function

'        Public Function SetStatusTextSB(pszStatusText As IntPtr) As Integer Implements IShellBrowser.SetStatusTextSB
'            Throw New NotImplementedException()
'        End Function

'        Public Function EnableModelessSB(fEnable As Boolean) As Integer Implements IShellBrowser.EnableModelessSB
'            Throw New NotImplementedException()
'        End Function

'        Public Function TranslateAcceleratorSB(pmsg As IntPtr, wID As Short) As Integer Implements IShellBrowser.TranslateAcceleratorSB
'            Throw New NotImplementedException()
'        End Function

'        Public Function BrowseObject(pidl As IntPtr, <MarshalAs(UnmanagedType.U4)> wFlags As UInteger) As Object Implements IShellBrowser.BrowseObject
'            Throw New NotImplementedException()
'        End Function

'        Public Function GetViewStateStream(grfMode As UInteger, ppStrm As IntPtr) As Integer Implements IShellBrowser.GetViewStateStream
'            Throw New NotImplementedException()
'        End Function

'        Public Function GetControlWindow(id As UInteger, ByRef phwnd As IntPtr) As Integer Implements IShellBrowser.GetControlWindow
'            Throw New NotImplementedException()
'        End Function

'        Public Function SendControlMsg(id As UInteger, uMsg As UInteger, wParam As UInteger, lParam As UInteger, pret As IntPtr) As Integer Implements IShellBrowser.SendControlMsg
'            Throw New NotImplementedException()
'        End Function

'        Public Function QueryActiveShellView(<MarshalAs(UnmanagedType.Interface)> ByRef ppshv As IShellView) As Integer Implements IShellBrowser.QueryActiveShellView
'            Throw New NotImplementedException()
'        End Function

'        Public Function OnViewWindowActive(<MarshalAs(UnmanagedType.Interface)> pshv As IShellView) As Integer Implements IShellBrowser.OnViewWindowActive
'            Return HRESULT.E_NOTIMPL
'        End Function

'        Public Function SetToolbarItems(lpButtons As IntPtr, nButtons As UInteger, uFlags As UInteger) As Integer Implements IShellBrowser.SetToolbarItems
'            Throw New NotImplementedException()
'        End Function
'    End Class

'End Class
