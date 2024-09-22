Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Threading
Imports System.Windows
Imports System.Windows.Interop

Public Class Shell
    Private Shared _desktop As Folder

    Public Shared Property DoInsertLevelUpFolder As Boolean = False

    Public Shared Event Notification(sender As Object, e As NotificationEventArgs)

    Private Shared _hNotify As UInt32
    Private Shared _w As Window
    Public Shared cm2 As IContextMenu2
    Public Shared cm3 As IContextMenu3
    Public Shared _hwnd As IntPtr

    Shared Sub New()
        Dim entry(0) As SHChangeNotifyEntry
        entry(0).pIdl = IntPtr.Zero
        entry(0).Recursively = True

        AddHandler Application.Current.MainWindow.Loaded,
            Sub(sender As Object, e As EventArgs)
                _w = New Window()
                _w.Owner = Application.Current.MainWindow
                _w.Left = Int32.MinValue
                _w.Top = Int32.MinValue
                _w.Width = 1
                _w.Height = 1
                _w.ShowInTaskbar = False
                _w.Show()

                Dim hwnd As IntPtr = New WindowInteropHelper(_w).Handle
                Dim source As HwndSource = HwndSource.FromHwnd(hwnd)
                _hwnd = hwnd
                source.AddHook(AddressOf HwndHook)

                _hNotify = Functions.SHChangeNotifyRegister(
                    hwnd,
                    SHCNRF.NewDelivery Or SHCNRF.InterruptLevel Or SHCNRF.ShellLevel,
                    SHCNE.ALLEVENTS,
                    WM.USER + 1,
                    1,
                    entry)

                AddHandler Application.Current.Exit,
                    Sub(s2 As Object, e2 As ExitEventArgs)
                        Functions.SHChangeNotifyDeregister(_hNotify)
                    End Sub
            End Sub
    End Sub

    Public Shared Function HwndHook(hwnd As IntPtr, msg As Integer, wParam As IntPtr, lParam As IntPtr, ByRef handled As Boolean) As IntPtr
        Dim hr As Integer
        If msg = WM.INITMENUPOPUP Or msg = WM.MEASUREITEM Or msg = WM.DRAWITEM Then
            If msg = WM.DRAWITEM Or msg = WM.MEASUREITEM Then
                Dim i As Int16 = 9
            End If
            If cm2 IsNot Nothing Then
                hr = cm2.HandleMenuMsg(msg, wParam, lParam)
                If hr = 0 Then
                    Return IntPtr.Zero
                End If
            ElseIf msg = WM.MENUCHAR Then
                If cm3 IsNot Nothing Then
                    hr = cm3.HandleMenuMsg2(msg, wParam, lParam, IntPtr.Zero)
                    If hr = 0 Then
                        Return IntPtr.Zero
                    End If
                End If
            End If
        End If
        If msg = WM.USER + 1 Then
            Dim pppidl As IntPtr = IntPtr.Zero
            Dim lEvent As SHCNE = 0
            Dim hLock As IntPtr = Functions.SHChangeNotification_Lock(wParam, lParam, pppidl, lEvent)

            If hLock <> IntPtr.Zero Then
                Dim pidl1 As IntPtr = Marshal.ReadIntPtr(pppidl)
                pppidl = IntPtr.Add(pppidl, IntPtr.Size)
                Dim pidl2 As IntPtr = Marshal.ReadIntPtr(pppidl)

                Dim path1 As String, item1 As IShellItem2
                If Not IntPtr.Zero.Equals(pidl1) Then
                    Dim path As StringBuilder = New StringBuilder(260)
                    Functions.SHGetPathFromIDList(pidl1, path)
                    path1 = path.ToString()
                    Debug.WriteLine(path1.ToString())
                    item1 = Item.GetIShellItem2FromParsingName(path1.ToString())
                End If

                Dim path2 As String, item2 As IShellItem2
                If Not IntPtr.Zero.Equals(pidl2) Then
                    Dim path As StringBuilder = New StringBuilder(260)
                    Functions.SHGetPathFromIDList(pidl2, path)
                    path2 = path.ToString()
                    Debug.WriteLine(path2.ToString())
                    item2 = Item.GetIShellItem2FromParsingName(path2.ToString())
                End If

                RaiseEvent Notification(Nothing, New NotificationEventArgs() With {
                    .Item1Path = path1,
                    .Item1 = item1,
                    .Item2Path = path2,
                    .Item2 = item2,
                    .[Event] = lEvent
                })

                Functions.SHChangeNotification_Unlock(hLock)
            End If
        End If
    End Function


    Public Shared ReadOnly Property Desktop As Folder
        Get
            If _desktop Is Nothing Then
                _desktop = Folder.FromKnownFolderGuid(Guids.KnownFolder_Desktop, Nothing)
            End If

            Return _desktop
        End Get
    End Property
End Class
