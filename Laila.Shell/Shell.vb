Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Threading
Imports System.Windows
Imports System.Windows.Interop

Public Class Shell
    Private Shared _desktop As Folder

    Public Shared Event Notification(sender As Object, e As NotificationEventArgs)
    Friend Shared Event UpdateDirCompleted()

    Friend Shared Property UpdateDirCounter As Integer

    Private Shared _hNotify As UInt32
    Private Shared _w As Window
    Private Shared _specialFolders As Dictionary(Of String, Folder) = New Dictionary(Of String, Folder)()

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

        _specialFolders.Add("Documents", Folder.FromParsingName("shell:::{d3162b92-9365-467a-956b-92703aca08af}", Nothing, Nothing))
        _specialFolders.Add("Pictures", Folder.FromParsingName("shell:::{24ad3ad4-a569-4530-98e1-ab02f9417aa8}", Nothing, Nothing))
        _specialFolders.Add("Downloads", Folder.FromParsingName("shell:::{088e3905-0323-4b02-9826-5d99428e115f}", Nothing, Nothing))
        _specialFolders.Add("Videos", Folder.FromParsingName("shell:::{A0953C92-50DC-43bf-BE83-3742FED03C9C}", Nothing, Nothing))
        _specialFolders.Add("Music", Folder.FromParsingName("shell:::{1CF1260C-4DD0-4ebb-811F-33C572699FDE}", Nothing, Nothing))
        _specialFolders.Add("Favorites", Folder.FromParsingName("shell:::{323CA680-C24D-4099-B94D-446DD2D7249E}", Nothing, Nothing))
    End Sub

    Public Shared Function HwndHook(hwnd As IntPtr, msg As Integer, wParam As IntPtr, lParam As IntPtr, ByRef handled As Boolean) As IntPtr
        If msg = WM.USER + 1 Then
            Dim pppidl As IntPtr = IntPtr.Zero
            Dim lEvent As SHCNE = 0
            Dim hLock As IntPtr = Functions.SHChangeNotification_Lock(wParam, lParam, pppidl, lEvent)

            If hLock <> IntPtr.Zero Then
                Debug.WriteLine(lEvent.ToString())

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

    Public Shared Sub InvokeUpdateDirCompleted()
        RaiseEvent UpdateDirCompleted()
    End Sub

    Public Shared ReadOnly Property SpecialFolders As Dictionary(Of String, Folder)
        Get
            Return _specialFolders
        End Get
    End Property

    Public Shared ReadOnly Property Desktop As Folder
        Get
            If _desktop Is Nothing Then
                _desktop = Folder.FromKnownFolderGuid(Guids.KnownFolder_Desktop, Nothing)
            End If

            Return _desktop
        End Get
    End Property
End Class
