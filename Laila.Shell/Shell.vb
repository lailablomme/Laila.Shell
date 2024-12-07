Imports System.Collections.Concurrent
Imports System.Runtime.InteropServices
Imports System.Threading
Imports System.Windows
Imports System.Windows.Input
Imports System.Windows.Interop
Imports Laila.Shell.Controls
Imports Laila.Shell.Events
Imports Laila.Shell.Helpers

Public Class Shell
    Private Shared _desktop As Folder

    Public Shared Event Notification(sender As Object, e As NotificationEventArgs)
    Friend Shared Event FolderNotification(sender As Object, e As FolderNotificationEventArgs)
    Public Shared Event ShuttingDown As EventHandler

    Public Shared SlowTaskQueue As New BlockingCollection(Of Action)
    Public Shared PriorityTaskQueue As New BlockingCollection(Of Action)
    Public Shared FolderTaskQueue As New BlockingCollection(Of Action)
    Private Shared _threads As List(Of Thread) = New List(Of Thread)()

    Public Shared IsSpecialFoldersReady As ManualResetEvent = New ManualResetEvent(False)
    Private Shared _shutDownTokensSource As CancellationTokenSource = New CancellationTokenSource()
    Public Shared ShuttingDownToken As CancellationToken = _shutDownTokensSource.Token

    Private Shared _hNotify As UInt32
    Friend Shared _w As Window
    Public Shared _hwnd As IntPtr

    Private Shared _specialFolders As Dictionary(Of String, Folder) = New Dictionary(Of String, Folder)()
    Private Shared _folderViews As Dictionary(Of String, Tuple(Of String, Type)) = New Dictionary(Of String, Tuple(Of String, Type))()

    Private Shared _itemsCacheLock As Object = New Object()
    Private Shared _itemsCache As List(Of Tuple(Of Item, DateTime)) = New List(Of Tuple(Of Item, DateTime))()
    Private Shared _isDebugVisible As Boolean = False
    Friend Shared _debugWindow As DebugTools.DebugWindow

    Private Shared _overrideCursorFunc As Func(Of Cursor, IDisposable) =
        Function(cursor As Cursor) As IDisposable
            Return New OverrideCursor(cursor)
        End Function

    Shared Sub New()
        Functions.OleInitialize(IntPtr.Zero)

        For i = 1 To Environment.ProcessorCount
            Dim staThread As Thread = New Thread(
                Sub()
                    Try
                        ' Process tasks from the queue
                        For Each task In PriorityTaskQueue.GetConsumingEnumerable(ShuttingDownToken)
                            task.Invoke()
                        Next
                    Catch ex As OperationCanceledException
                        Debug.WriteLine("PriorityTaskQueue was canceled.")
                    End Try
                End Sub)
            staThread.SetApartmentState(ApartmentState.STA)
            staThread.Start()
            _threads.Add(staThread)
        Next
        For i = 1 To 35
            Dim staThread As Thread = New Thread(
                Sub()
                    Try
                        ' Process tasks from the queue
                        For Each task In SlowTaskQueue.GetConsumingEnumerable(ShuttingDownToken)
                            task.Invoke()
                        Next
                    Catch ex As OperationCanceledException
                        Debug.WriteLine("SlowTaskQueue was canceled.")
                    End Try
                End Sub)
            staThread.SetApartmentState(ApartmentState.STA)
            staThread.Start()
            _threads.Add(staThread)
        Next
        Dim staThread2 As Thread = New Thread(
            Sub()
                Try
                    ' Process tasks from the queue
                    Functions.OleInitialize(IntPtr.Zero)
                    For Each task In FolderTaskQueue.GetConsumingEnumerable(ShuttingDownToken)
                        task.Invoke()
                    Next
                Catch ex As OperationCanceledException
                    Debug.WriteLine("FolderTaskQueue was canceled.")
                End Try
            End Sub)
        staThread2.SetApartmentState(ApartmentState.STA)
        staThread2.Start()
        _threads.Add(staThread2)
        Dim staThread3 As Thread = New Thread(
            Sub()
                Try
                    While Not ShuttingDownToken.IsCancellationRequested
                        For Each item In _itemsCache.ToList()
                            If ShuttingDownToken.IsCancellationRequested Then Exit For
                            If Not item Is Nothing AndAlso Not item.Item1 Is Nothing AndAlso DateTime.Now.Subtract(item.Item2).TotalMilliseconds > 2000 Then
                                item.Item1.MaybeDispose()
                            End If
                            Thread.Sleep(25)
                        Next
                        Thread.Sleep(10000)
                    End While
                Catch ex As OperationCanceledException
                    Debug.WriteLine("Disposing thread was canceled.")
                End Try
            End Sub)
        staThread3.SetApartmentState(ApartmentState.STA)
        staThread3.Start()
        _threads.Add(staThread3)

        Dim entry(0) As SHChangeNotifyEntry
        entry(0).pIdl = IntPtr.Zero
        entry(0).Recursively = True

        _w = New Window()
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

        EventManager.RegisterClassHandler(GetType(Window), Window.LoadedEvent, New RoutedEventHandler(AddressOf window_Loaded))

        Dim addSpecialFolder As Action(Of String, Item) =
            Sub(name As String, item As Item)
                If Not item Is Nothing AndAlso TypeOf item Is Folder Then
                    _specialFolders.Add(name, item)
                End If
            End Sub

        Shell.FolderTaskQueue.Add(
            Sub()
                addSpecialFolder("Home", Folder.FromParsingName("shell:::{679f85cb-0220-4080-b29b-5540cc05aab6}", Nothing, False))
                addSpecialFolder("Desktop", Folder.FromDesktop())
                addSpecialFolder("Documents", Folder.FromParsingName("shell:::{d3162b92-9365-467a-956b-92703aca08af}", Nothing, False))
                addSpecialFolder("Pictures", Folder.FromParsingName("shell:::{24ad3ad4-a569-4530-98e1-ab02f9417aa8}", Nothing, False))
                addSpecialFolder("Downloads", Folder.FromParsingName("shell:::{088e3905-0323-4b02-9826-5d99428e115f}", Nothing, False))
                addSpecialFolder("Videos", Folder.FromParsingName("shell:::{A0953C92-50DC-43bf-BE83-3742FED03C9C}", Nothing, False))
                addSpecialFolder("Music", Folder.FromParsingName("shell:::{1CF1260C-4DD0-4ebb-811F-33C572699FDE}", Nothing, False))
                addSpecialFolder("Favorites", Folder.FromParsingName("shell:::{323CA680-C24D-4099-B94D-446DD2D7249E}", Nothing, False))
                addSpecialFolder("This computer", Folder.FromParsingName("shell:::{20D04FE0-3AEA-1069-A2D8-08002B30309D}", Nothing, False))
                addSpecialFolder("Network", Folder.FromParsingName("shell:::{F02C1A0D-BE21-4350-88B0-7367FC96EF3C}", Nothing, False))
                addSpecialFolder("Gallery", Folder.FromParsingName("shell:::{E88865EA-0E1C-4E20-9AA6-EDCD0212C87C}", Nothing, False))
                addSpecialFolder("Recycle Bin", Folder.FromParsingName("shell:::{645FF040-5081-101B-9F08-00AA002F954E}", Nothing, False))
                'addSpecialFolder("Recent", CType(Folder.FromParsingName(Environment.ExpandEnvironmentVariables("%APPDATA%\Microsoft\Windows"), Nothing), Folder) _
                '                                 .GetItems().First(Function(i) i.FullPath.EndsWith("\Recent")))
                addSpecialFolder("OneDrive", Folder.FromParsingName("shell:::{018D5C66-4533-4307-9B53-224DE2ED1FE6}", Nothing, False))
                addSpecialFolder("OneDrive Business", Folder.FromParsingName("shell:::{04271989-C4D2-BEC7-A521-3DF166FAB4BA}", Nothing, False))
                '_specialFolders.Add("Windows Tools", Folder.FromParsingName("shell:::{D20EA4E1-3957-11D2-A40B-0C5020524153}", Nothing))
                _specialFolders.Add("Libraries", Folder.FromParsingName("shell:::{031E4825-7B94-4DC3-B131-E946B44C8DD5}", Nothing, False))
                '_specialFolders.Add("User Pinned", Folder.FromParsingName("shell:::{1F3427C8-5C10-4210-AA03-2EE45287D668}", Nothing))
                '_specialFolders.Add("Control Panel", Folder.FromParsingName("shell:::{26EE0668-A00A-44D7-9371-BEB064C98683}", Nothing))
                _specialFolders.Add("Devices and Printers", Folder.FromParsingName("shell:::{A8A91A66-3A7D-4424-8D24-04E180695C7A}", Nothing, False))
                '_specialFolders.Add("All Tasks", Folder.FromParsingName("shell:::{ED7BA470-8E54-465E-825C-99712043E01C}", Nothing))
                '_specialFolders.Add("Applications", Folder.FromParsingName("shell:::{4234d49b-0245-4df3-b780-3893943456e1}", Nothing))
                addSpecialFolder("Frequent Folders", Folder.FromParsingName("shell:::{3936E9E4-D92C-4EEE-A85A-BC16D5EA0819}", Nothing, False))
                addSpecialFolder("User Profile", Folder.FromParsingName(Environment.ExpandEnvironmentVariables("%USERPROFILE%"), Nothing, False))
                '_specialFolders.Add("Installed Updates", Folder.FromParsingName("shell:::{d450a8a1-9568-45c7-9c0e-b4f9fb4537bd}", Nothing))
                '_specialFolders.Add("Network Connections", Folder.FromParsingName("shell:::{7007ACC7-3202-11D1-AAD2-00805FC1270E}", Nothing))
                _specialFolders.Add("Programs and Features", Folder.FromParsingName("shell:::{7b81be6a-ce2b-4676-a29e-eb907a5126c5}", Nothing, False))
                '_specialFolders.Add("Public", Folder.FromParsingName("shell:::{4336a54d-038b-4685-ab02-99bb52d3fb8b}", Nothing))
                '_specialFolders.Add("Recent Items", Folder.FromParsingName("shell:::{4564b25e-30cd-4787-82ba-39e73a750b14}", Nothing))

                Shell.IsSpecialFoldersReady.Set()
            End Sub)

        FolderViews.Add("Extra large icons", New Tuple(Of String, Type)("pack://application:,,,/Laila.Shell;component/Images/extralargeicons16.png", GetType(ExtraLargeIconsView)))
        FolderViews.Add("Large icons", New Tuple(Of String, Type)("pack://application:,,,/Laila.Shell;component/Images/largeicons16.png", GetType(LargeIconsView)))
        FolderViews.Add("Normal icons", New Tuple(Of String, Type)("pack://application:,,,/Laila.Shell;component/Images/normalicons16.png", GetType(NormalIconsView)))
        FolderViews.Add("Small icons", New Tuple(Of String, Type)("pack://application:,,,/Laila.Shell;component/Images/smallicons16.png", GetType(SmallIconsView)))
        FolderViews.Add("List", New Tuple(Of String, Type)("pack://application:,,,/Laila.Shell;component/Images/list16.png", GetType(ListView)))
        FolderViews.Add("Details", New Tuple(Of String, Type)("pack://application:,,,/Laila.Shell;component/Images/details16.png", GetType(DetailsView)))
        FolderViews.Add("Tiles", New Tuple(Of String, Type)("pack://application:,,,/Laila.Shell;component/Images/tiles16.png", GetType(TileView)))
        FolderViews.Add("Content", New Tuple(Of String, Type)("pack://application:,,,/Laila.Shell;component/Images/content16.png", GetType(ContentView)))

        If _isDebugVisible Then
            _debugWindow = New DebugTools.DebugWindow()
            _debugWindow.Show()
        End If
    End Sub

    Private Shared Sub window_Loaded(sender As Object, e As EventArgs)
        RemoveHandler CType(sender, Window).Closed, AddressOf window_Closed
        AddHandler CType(sender, Window).Closed, AddressOf window_Closed
    End Sub

    Private Shared Sub window_Closed(sender As Object, e As EventArgs)
        If System.Windows.Application.Current.ShutdownMode = ShutdownMode.OnLastWindowClose _
            AndAlso System.Windows.Application.Current.Windows.Cast(Of Window) _
                .Where(Function(w) Not w.GetType().ToString().StartsWith("Microsoft.VisualStudio.")).Count <= 1 Then
            Shell.Shutdown()
        End If
    End Sub

    Public Shared Sub Shutdown()
        If Not Shell.ShuttingDownToken.IsCancellationRequested Then
            Functions.SHChangeNotifyDeregister(_hNotify)

            Dim list As List(Of Tuple(Of Item, DateTime))
            SyncLock _itemsCacheLock
                list = Shell.ItemsCache.ToList()
            End SyncLock
            While Not list.Count = 0
                For Each item In list
                    item.Item1.Dispose()
                Next
                SyncLock _itemsCacheLock
                    list = Shell.ItemsCache.ToList()
                End SyncLock
            End While

            RaiseEvent ShuttingDown(Nothing, New EventArgs())

            _shutDownTokensSource.Cancel()

            Functions.OleUninitialize()

            _w.Close()
        End If
    End Sub

    Public Shared Function HwndHook(hwnd As IntPtr, msg As Integer, wParam As IntPtr, lParam As IntPtr, ByRef handled As Boolean) As IntPtr
        'Debug.WriteLine(CType(msg, WM).ToString())
        If msg = WM.USER + 1 Then
            Dim pppidl As IntPtr = IntPtr.Zero
            Dim lEvent As SHCNE = 0
            Dim hLock As IntPtr = Functions.SHChangeNotification_Lock(wParam, lParam, pppidl, lEvent)

            If hLock <> IntPtr.Zero Then
                Dim pidl1 As IntPtr = Marshal.ReadIntPtr(pppidl)
                pppidl = IntPtr.Add(pppidl, IntPtr.Size)
                Dim pidl2 As IntPtr = Marshal.ReadIntPtr(pppidl)

                Dim e As NotificationEventArgs = New NotificationEventArgs() With {
                    .[Event] = lEvent
                }

                Debug.WriteLine(lEvent.ToString() & "  w=" & wParam.ToString() & "  l=" & lParam.ToString())

                If Not IntPtr.Zero.Equals(pidl1) Then
                    e.Item1Pidl = New Pidl(pidl1)
                    Using i = Item.FromPidl(pidl1, Nothing, True)
                        Debug.WriteLine(BitConverter.ToString(i.Pidl.Bytes) & vbCrLf & i.DisplayName & " (" & i.FullPath & ")")
                    End Using
                End If
                If Not IntPtr.Zero.Equals(pidl2) Then
                    e.Item2Pidl = New Pidl(pidl2)
                    Using i = Item.FromPidl(pidl1, Nothing, True)
                        Debug.WriteLine(BitConverter.ToString(i.Pidl.Bytes) & vbCrLf & i.DisplayName & " (" & i.FullPath & ")")
                    End Using
                End If

                RaiseEvent Notification(Nothing, e)

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

    Public Shared Function GetSpecialFolder(id As String) As Folder
        Shell.IsSpecialFoldersReady.WaitOne()
        Return _specialFolders(id)
    End Function

    Public Shared Function GetSpecialFolders() As Dictionary(Of String, Folder)
        Shell.IsSpecialFoldersReady.WaitOne()
        Return _specialFolders
    End Function

    Public Shared ReadOnly Property FolderViews As Dictionary(Of String, Tuple(Of String, Type))
        Get
            Return _folderViews
        End Get
    End Property

    Public Shared ReadOnly Property Desktop As Folder
        Get
            If _desktop Is Nothing Then
                _desktop = Shell.GetSpecialFolder("Desktop")
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

    Public Shared ReadOnly Property ItemsCache As List(Of Tuple(Of Item, DateTime))
        Get
            Return _itemsCache
        End Get
    End Property

    Public Shared Sub AddToItemsCache(item As Item)
        SyncLock _itemsCacheLock
            _itemsCache.Add(New Tuple(Of Item, Date)(item, DateTime.Now))
        End SyncLock
    End Sub

    Public Shared Sub RemoveFromItemsCache(item As Item)
        SyncLock _itemsCacheLock
            _itemsCache.Remove(_itemsCache.FirstOrDefault(Function(i) item.Equals(i.Item1)))
        End SyncLock
    End Sub

    Public Shared Property OverrideCursor As Func(Of Cursor, IDisposable)
        Get
            Return _overrideCursorFunc
        End Get
        Set(value As Func(Of Cursor, IDisposable))
            _overrideCursorFunc = value
        End Set
    End Property
End Class
