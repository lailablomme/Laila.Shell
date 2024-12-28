'Imports Laila.Shell.Controls.FolderView
'Imports System.Runtime.InteropServices
'Imports System.Windows

'Public Class HostShellView
'    Public Sub HostShellView()
'        Dim ptr As IntPtr, shellFolder As IShellFolder, shellView As IShellView
'        Try
'            shellFolder = CType(e.NewValue, Folder).ShellFolder
'            shellFolder.CreateViewObject(IntPtr.Zero, Guids.IID_IShellView, ptr)
'            shellView = Marshal.GetTypedObjectForIUnknown(ptr, GetType(IShellView))
'        Finally
'            If Not IntPtr.Zero.Equals(ptr) Then
'                Marshal.Release(ptr)
'            End If
'        End Try
'        Dim r As WIN32RECT = New WIN32RECT()
'        r.Top = 20 : r.Left = 20 : r.Right = 500 : r.Bottom = 500
'        Dim fs As FOLDERSETTINGS = New FOLDERSETTINGS()
'        fs.ViewMode = -1
'        Dim hwnd As IntPtr
'        'Dim w As Window = New Window()
'        'Dim wfh As WindowsFormsHost = New WindowsFormsHost()

'        'w.Content = wfh
'        Dim ww As System.Windows.Forms.Form = New Forms.Form()
'        'Dim pan As System.Windows.Forms.Panel = New System.Windows.Forms.Panel()
'        'wfh.Child = pan
'        'pan.BackColor = System.Drawing.Color.Red
'        ' w.Show()
'        ww.Show()
'        Dim sb As ShellBrowser = New ShellBrowser(ww.Handle)
'        Dim h As HRESULT = shellView.CreateViewWindow(Nothing, fs, sb, r, hwnd)
'        Debug.WriteLine("CreateViewWindow returned " & h.ToString() & "   & hwnd=" & hwnd.ToString())
'        ' Functions.SetParent(hwnd, ww.Handle)
'        ' Functions.MoveWindow(hwnd, 0, 0, ww.Width, ww.Height, True)
'        shellView.UIActivate(1)

'    End Sub
'End Class
