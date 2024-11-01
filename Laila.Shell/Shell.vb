Imports System.Drawing
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Threading
Imports System.Windows
Imports System.Windows.Interop
Imports Laila.Shell.Controls
Imports Laila.Shell.Events
Imports Laila.Shell.Helpers
Imports Microsoft.Win32

Public Class Shell
    Private Shared _desktop As Folder

    Public Shared Event Notification(sender As Object, e As NotificationEventArgs)
    Friend Shared Event FolderNotification(sender As Object, e As FolderNotificationEventArgs)

    Private Shared _hNotify As UInt32
    Friend Shared _w As Window
    Private Shared _specialFolders As Dictionary(Of String, Folder) = New Dictionary(Of String, Folder)()
    Public Shared _hwnd As IntPtr
    Private Shared _folderViews As Dictionary(Of String, Type) = New Dictionary(Of String, Type)()

    Shared Sub New()
        Functions.OleInitialize(IntPtr.Zero)

        Dim entry(0) As SHChangeNotifyEntry
        entry(0).pIdl = IntPtr.Zero
        entry(0).Recursively = True

        AddHandler Application.Current.MainWindow.Loaded,
            Async Sub(sender As Object, e As EventArgs)
                _w = New Window()
                _w.Owner = Application.Current.MainWindow
                _w.Left = Int32.MinValue
                _w.Top = Int32.MinValue
                _w.WindowStyle = WindowStyle.None
                _w.Width = 1920
                _w.Height = 1080
                _w.ShowInTaskbar = False
                _w.Title = "Hidden Window"
                _w.Show()

                Dim hwnd As IntPtr = New WindowInteropHelper(_w).Handle
                Dim source As HwndSource = HwndSource.FromHwnd(hwnd)
                source.AddHook(AddressOf HwndHook)
                _hwnd = hwnd

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
                        Functions.OleUninitialize()
                    End Sub
            End Sub

        Dim addSpecialFolder As Action(Of String, Item) =
            Sub(name As String, item As Item)
                If Not item Is Nothing AndAlso TypeOf item Is Folder Then
                    _specialFolders.Add(name, item)
                End If
            End Sub
        addSpecialFolder("Home", Folder.FromParsingName("shell:::{679f85cb-0220-4080-b29b-5540cc05aab6}", Nothing))
        addSpecialFolder("Desktop", Folder.FromParsingName("shell:::{" & Guids.KnownFolder_Desktop.ToString() & "}", Nothing))
        addSpecialFolder("Documents", Folder.FromParsingName("shell:::{d3162b92-9365-467a-956b-92703aca08af}", Nothing))
        addSpecialFolder("Pictures", Folder.FromParsingName("shell:::{24ad3ad4-a569-4530-98e1-ab02f9417aa8}", Nothing))
        addSpecialFolder("Downloads", Folder.FromParsingName("shell:::{088e3905-0323-4b02-9826-5d99428e115f}", Nothing))
        addSpecialFolder("Videos", Folder.FromParsingName("shell:::{A0953C92-50DC-43bf-BE83-3742FED03C9C}", Nothing))
        addSpecialFolder("Music", Folder.FromParsingName("shell:::{1CF1260C-4DD0-4ebb-811F-33C572699FDE}", Nothing))
        addSpecialFolder("Favorites", Folder.FromParsingName("shell:::{323CA680-C24D-4099-B94D-446DD2D7249E}", Nothing))
        addSpecialFolder("This computer", Folder.FromParsingName("shell:::{20D04FE0-3AEA-1069-A2D8-08002B30309D}", Nothing))
        addSpecialFolder("Network", Folder.FromParsingName("shell:::{F02C1A0D-BE21-4350-88B0-7367FC96EF3C}", Nothing))
        addSpecialFolder("Gallery", Folder.FromParsingName("shell:::{E88865EA-0E1C-4E20-9AA6-EDCD0212C87C}", Nothing))
        addSpecialFolder("Recycle Bin", Folder.FromParsingName("shell:::{645FF040-5081-101B-9F08-00AA002F954E}", Nothing))
        addSpecialFolder("Recent", CType(Folder.FromParsingName("%APPDATA%\Microsoft\Windows", Nothing), Folder) _
                                         .GetItems().First(Function(i) i.FullPath.EndsWith("\Recent")))
        addSpecialFolder("OneDrive", Folder.FromParsingName("shell:::{018D5C66-4533-4307-9B53-224DE2ED1FE6}", Nothing))
        addSpecialFolder("OneDrive Business", Folder.FromParsingName("shell:::{04271989-C4D2-BEC7-A521-3DF166FAB4BA}", Nothing))
        '_specialFolders.Add("Windows Tools", Folder.FromParsingName("shell:::{D20EA4E1-3957-11D2-A40B-0C5020524153}", Nothing))
        _specialFolders.Add("Libraries", Folder.FromParsingName("shell:::{031E4825-7B94-4DC3-B131-E946B44C8DD5}", Nothing))
        '_specialFolders.Add("User Pinned", Folder.FromParsingName("shell:::{1F3427C8-5C10-4210-AA03-2EE45287D668}", Nothing))
        '_specialFolders.Add("Control Panel", Folder.FromParsingName("shell:::{26EE0668-A00A-44D7-9371-BEB064C98683}", Nothing))
        _specialFolders.Add("Devices and Printers", Folder.FromParsingName("shell:::{A8A91A66-3A7D-4424-8D24-04E180695C7A}", Nothing))
        '_specialFolders.Add("All Tasks", Folder.FromParsingName("shell:::{ED7BA470-8E54-465E-825C-99712043E01C}", Nothing))
        '_specialFolders.Add("Applications", Folder.FromParsingName("shell:::{4234d49b-0245-4df3-b780-3893943456e1}", Nothing))
        addSpecialFolder("Frequent Folders", Folder.FromParsingName("shell:::{3936E9E4-D92C-4EEE-A85A-BC16D5EA0819}", Nothing))
        addSpecialFolder("User Profile", Folder.FromParsingName("%USERPROFILE%", Nothing))
        '_specialFolders.Add("Installed Updates", Folder.FromParsingName("shell:::{d450a8a1-9568-45c7-9c0e-b4f9fb4537bd}", Nothing))
        '_specialFolders.Add("Network Connections", Folder.FromParsingName("shell:::{7007ACC7-3202-11D1-AAD2-00805FC1270E}", Nothing))
        '_specialFolders.Add("Programs and Features", Folder.FromParsingName("shell:::{7b81be6a-ce2b-4676-a29e-eb907a5126c5}", Nothing))
        '_specialFolders.Add("Public", Folder.FromParsingName("shell:::{4336a54d-038b-4685-ab02-99bb52d3fb8b}", Nothing))
        '_specialFolders.Add("Recent Items", Folder.FromParsingName("shell:::{4564b25e-30cd-4787-82ba-39e73a750b14}", Nothing))

        FolderViews.Add("Large icons", GetType(LargeIconsView))
        FolderViews.Add("Normal icons", GetType(NormalIconsView))
        FolderViews.Add("Details", GetType(DetailsView))
    End Sub

    Public Shared Function HwndHook(hwnd As IntPtr, msg As Integer, wParam As IntPtr, lParam As IntPtr, ByRef handled As Boolean) As IntPtr
        'Debug.WriteLine(CType(msg, WM).ToString())
        If msg = WM.USER + 1 Then
            Dim pppidl As IntPtr = IntPtr.Zero
            Dim lEvent As SHCNE = 0
            Dim hLock As IntPtr = Functions.SHChangeNotification_Lock(wParam, lParam, pppidl, lEvent)

            If hLock <> IntPtr.Zero Then
                Debug.WriteLine(lEvent.ToString())

                Dim pidl1 As IntPtr = Marshal.ReadIntPtr(pppidl)
                pppidl = IntPtr.Add(pppidl, IntPtr.Size)
                Dim pidl2 As IntPtr = Marshal.ReadIntPtr(pppidl)

                Dim path1 As String
                If Not IntPtr.Zero.Equals(pidl1) Then
                    Dim path As StringBuilder = New StringBuilder(260)
                    Functions.SHGetPathFromIDList(pidl1, path)
                    path1 = path.ToString()
                    Debug.WriteLine(pidl1.ToString() & "/" & path1.ToString())
                End If

                Dim path2 As String
                If Not IntPtr.Zero.Equals(pidl2) Then
                    Dim path As StringBuilder = New StringBuilder(260)
                    Functions.SHGetPathFromIDList(pidl2, path)
                    path2 = path.ToString()
                    Debug.WriteLine(pidl2.ToString() & "/" & path2.ToString())
                End If

                RaiseEvent Notification(Nothing, New NotificationEventArgs() With {
                    .Item1Path = path1,
                    .Item2Path = path2,
                    .[Event] = lEvent
                })

                Functions.SHChangeNotification_Unlock(hLock)
            End If
        ElseIf msg = 49252 Then
            Dim dragImage As SHDRAGIMAGE = Marshal.PtrToStructure(Of SHDRAGIMAGE)(lParam)
            dragImage.sizeDragImage.Width = Drag._bitmap.Width
            dragImage.sizeDragImage.Height = Drag._bitmap.Height
            dragImage.ptOffset.x = Drag.ICON_SIZE / 2
            dragImage.ptOffset.y = Drag.ICON_SIZE / 2
            dragImage.hbmpDragImage = Drag._bitmap.GetHbitmap()
            dragImage.crColorKey = System.Drawing.Color.Purple.ToArgb()
            Marshal.StructureToPtr(Of SHDRAGIMAGE)(dragImage, lParam, False)
        End If
    End Function

    Public Shared ReadOnly Property SpecialFolders As Dictionary(Of String, Folder)
        Get
            Return _specialFolders
        End Get
    End Property

    Public Shared ReadOnly Property FolderViews As Dictionary(Of String, Type)
        Get
            Return _folderViews
        End Get
    End Property

    Public Shared ReadOnly Property Desktop As Folder
        Get
            If _desktop Is Nothing Then
                _desktop = Shell.SpecialFolders("Desktop")
            End If

            Return _desktop
        End Get
    End Property

    Friend Shared Sub RaiseFolderNotificationEvent(sender As Object, e As FolderNotificationEventArgs)
        RaiseEvent FolderNotification(sender, e)
    End Sub

    Friend Shared Sub RaiseNotificationEvent(sender As Object, e As NotificationEventArgs)
        RaiseEvent Notification(sender, e)
    End Sub
End Class
