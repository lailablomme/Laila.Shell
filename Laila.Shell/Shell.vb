Imports System.Collections.Concurrent
Imports System.IO
Imports System.Runtime.CompilerServices
Imports System.Runtime.InteropServices
Imports System.Security.AccessControl
Imports System.Threading
Imports System.Windows
Imports System.Windows.Input
Imports System.Windows.Interop
Imports Laila.Shell.Controls
Imports Laila.Shell.Events
Imports Laila.Shell.Helpers
Imports Laila.Shell.Interfaces
Imports Laila.Shell.Interop
Imports Laila.Shell.Interop.Items
Imports Laila.Shell.Interop.Properties
Imports Laila.Shell.Interop.Windows
Imports Shell32

Public Class Shell
    Private Shared _desktop As Folder

    Public Shared Event ClipboardChanged As EventHandler

    Public Shared NotificationMainThread As Helpers.ThreadPool
    Public Shared NotificationThreadPool As Helpers.ThreadPool
    Public Shared GlobalThreadPool As Helpers.ThreadPool
    Public Shared DisposerThreadPool As Helpers.ThreadPool
    Private Shared _threads As List(Of Thread) = New List(Of Thread)()
    Private Shared _disposerLoopThread As Thread
    Private Shared _unhookUpdatesThread As Thread
    Private Shared _disposerBlockLock As SemaphoreSlim = New SemaphoreSlim(1, 1)
    Private Shared _disposerBlock As Integer = 0

    Public Shared IsSpecialFoldersReady As ManualResetEvent = New ManualResetEvent(False)
    Private Shared _shutDownTokensSource As CancellationTokenSource = New CancellationTokenSource()
    Public Shared ShuttingDownToken As CancellationToken = _shutDownTokensSource.Token

    Private Shared _mainWindow As Window
    Friend Shared _w As Window
    Public Shared _hwnd As IntPtr

    Private Shared _specialFolders As Dictionary(Of SpecialFolders, Folder) = New Dictionary(Of SpecialFolders, Folder)()
    Private Shared _customFolders As List(Of CustomFolder) = New List(Of CustomFolder)()
    Private Shared _customPropertiesByCanonicalName As Dictionary(Of String, Type) = New Dictionary(Of String, Type)()
    Private Shared _customPropertiesByKey As Dictionary(Of String, Type) = New Dictionary(Of String, Type)()

    Private Shared _notificationSubscribersLock As Object = New Object()
    Private Shared _notificationSubscribers As List(Of IProcessNotifications) = New List(Of IProcessNotifications)()
    Private Shared _threadPoolCacheLock As Object = New Object()
    Private Shared _threadPoolCache As List(Of Helpers.ThreadPool) = New List(Of Helpers.ThreadPool)()
    Private Shared _menuCacheLock As Object = New Object()
    Private Shared _menuCache As List(Of BaseMenu) = New List(Of BaseMenu)()
    Friend Shared _itemsCacheLock As SemaphoreSlim = New SemaphoreSlim(1, 1)
    Private Shared _itemsCache As List(Of Tuple(Of Item, DateTime)) = New List(Of Tuple(Of Item, DateTime))()
    Private Shared _fileSystemCache As Hashtable = New Hashtable()
    Private Shared _fileSystemCacheCount As Hashtable = New Hashtable()
    Private Shared _fileSystemCacheLock As SemaphoreSlim = New SemaphoreSlim(1, 1)
    Private Shared _controlCacheLock As Object = New Object()
    Private Shared _controlCache As List(Of IDisposable) = New List(Of IDisposable)()
    Private Shared _isDebugVisible As Boolean = False
    Friend Shared _debugWindow As DebugTools.DebugWindow

    Private Shared _nextClipboardViewer As IntPtr

    Private Shared _overrideCursorFunc As Func(Of Cursor, IDisposable) =
        Function(cursor As Cursor) As IDisposable
            Return New OverrideCursor(cursor)
        End Function

    Private Shared _settings As Settings = New Settings()

    Shared Sub New()
        ' watch for windows being loaded so we can gracefully shutdown when they're closed
        EventManager.RegisterClassHandler(GetType(Window), Window.LoadedEvent, New RoutedEventHandler(AddressOf window_Loaded))

        If Application.Current.ShutdownMode = ShutdownMode.OnExplicitShutdown Then
            ' hook application exit
            AddHandler Application.Current.Exit,
                Sub(s As Object, e As ExitEventArgs)
                    Shutdown()
                End Sub
        End If

        If Not Application.Current.MainWindow Is Nothing Then
            RemoveHandler CType(Application.Current.MainWindow, Window).Closed, AddressOf window_Closed
            RemoveHandler CType(Application.Current.MainWindow, Window).Closing, AddressOf window_Closing
            AddHandler CType(Application.Current.MainWindow, Window).Closed, AddressOf window_Closed
            AddHandler CType(Application.Current.MainWindow, Window).Closing, AddressOf window_Closing
        End If

        ' initialize com & ole
        Functions.OleInitialize(IntPtr.Zero)

        ImageHelper.Load()

        Shell.NotificationMainThread = New Helpers.ThreadPool(1)
        Shell.NotificationThreadPool = New Helpers.ThreadPool(75)
        Shell.GlobalThreadPool = New Helpers.ThreadPool(100)
        Shell.DisposerThreadPool = New Helpers.ThreadPool(10)

        ' thread for disposing items
        _disposerLoopThread = New Thread(
            Sub()
                Try
                    ' while we're not shutting down...
                    While Not ShuttingDownToken.IsCancellationRequested
                        ' take a snapshot of the shellitem cache...
                        Dim list As List(Of Tuple(Of Item, DateTime))
                        If _disposerBlock = 0 Then
                            _itemsCacheLock.Wait()
                            Try
                                list = Shell.ItemsCache.ToList()
                            Finally
                                _itemsCacheLock.Release()
                            End Try

                            Dim i As Long = 0
                            ' ...and go through it
                            For Each item In list
                                If _disposerBlock OrElse ShuttingDownToken.IsCancellationRequested Then Exit For

                                ' try to dispose the item
                                If Not item Is Nothing AndAlso Not item.Item1 Is Nothing _
                                    AndAlso DateTime.Now.Subtract(item.Item2).TotalMilliseconds > 10000 Then
                                    item.Item1.MaybeDispose()
                                End If

                                ' don't hog the process
                                i += 1
                                If i Mod 25 = 0 Then
                                    Thread.Sleep(250)
                                End If
                            Next
                        End If

                        Thread.Sleep(5000)
                    End While
                Catch ex As OperationCanceledException
                    Debug.WriteLine("Disposing thread was canceled.")
                End Try
            End Sub)
        _disposerLoopThread.IsBackground = True
        _disposerLoopThread.SetApartmentState(ApartmentState.MTA)
        _disposerLoopThread.Start()
        _threads.Add(_disposerLoopThread)

        ' make window to receive SHChangeNotify messages and for building
        ' the drag and drop image
        _w = New Window()
        _w.Left = Int32.MinValue
        _w.Top = Int32.MinValue
        _w.WindowStyle = WindowStyle.None
        _w.Width = 1920
        _w.Height = 1080
        _w.ShowInTaskbar = False
        _w.Title = "Messaging Window"
        _w.Show()

        ' add hook
        Dim hwnd As IntPtr = New WindowInteropHelper(_w).Handle
        Dim source As HwndSource = HwndSource.FromHwnd(hwnd)
        source.AddHook(AddressOf hwndHook)
        _hwnd = hwnd

        ' start listening for clipboard changes
        _nextClipboardViewer = Functions.SetClipboardViewer(hwnd)

        _customFolders.Add(New CustomFolder() With {
            .FullPath = "::{b8b10b36-5c36-4f45-ae9a-79f0297d64e1}",
            .ReplacesFullPath = "::{679f85cb-0220-4080-b29b-5540cc05aab6}",
            .Type = GetType(HomeFolder)
        })
        _customFolders.Add(New CustomFolder() With {
            .FullPath = "::{b8b10b36-5c36-4f45-ae9a-79f0297d64e1}",
            .ReplacesFullPath = "::{F874310E-B6B7-47DC-BC84-B9E6B38F5903}",
            .Type = GetType(HomeFolder)
        })
        _customPropertiesByCanonicalName = New Dictionary(Of String, Type) From {
            {System_StorageProviderUIStatusProperty.CanonicalName, GetType(System_StorageProviderUIStatusProperty)},
            {Home_CategoryProperty.CanonicalName, GetType(Home_CategoryProperty)}
        }
        _customPropertiesByKey = New Dictionary(Of String, Type) From {
            {System_StorageProviderUIStatusProperty.Key.ToString(), GetType(System_StorageProviderUIStatusProperty)},
            {Home_CategoryProperty.Key.ToString(), GetType(Home_CategoryProperty)}
        }

        Shell.GlobalThreadPool.Add(
            Sub()
                ' add special folders
                Shell.AddSpecialFolder(SpecialFolders.Desktop, Folder.FromDesktop())
                Shell.AddSpecialFolder(SpecialFolders.Home, Folder.FromParsingName("shell:::{b8b10b36-5c36-4f45-ae9a-79f0297d64e1}", Nothing, False))
                Shell.AddSpecialFolder(SpecialFolders.Documents, Folder.FromParsingName("shell:::{d3162b92-9365-467a-956b-92703aca08af}", Nothing, False))
                Shell.AddSpecialFolder(SpecialFolders.Pictures, Folder.FromParsingName("shell:::{24ad3ad4-a569-4530-98e1-ab02f9417aa8}", Nothing, False))
                Shell.AddSpecialFolder(SpecialFolders.Downloads, Folder.FromParsingName("shell:::{088e3905-0323-4b02-9826-5d99428e115f}", Nothing, False))
                Shell.AddSpecialFolder(SpecialFolders.Videos, Folder.FromParsingName("shell:::{A0953C92-50DC-43bf-BE83-3742FED03C9C}", Nothing, False))
                Shell.AddSpecialFolder(SpecialFolders.Music, Folder.FromParsingName("shell:::{1CF1260C-4DD0-4ebb-811F-33C572699FDE}", Nothing, False))
                Shell.AddSpecialFolder(SpecialFolders.Favorites, Folder.FromParsingName("shell:::{323CA680-C24D-4099-B94D-446DD2D7249E}", Nothing, False))
                Shell.AddSpecialFolder(SpecialFolders.ThisPc, Folder.FromParsingName("shell:::{20D04FE0-3AEA-1069-A2D8-08002B30309D}", Nothing, False))
                Shell.AddSpecialFolder(SpecialFolders.Network, Folder.FromParsingName("shell:::{F02C1A0D-BE21-4350-88B0-7367FC96EF3C}", Nothing, False))
                Shell.AddSpecialFolder(SpecialFolders.Gallery, Folder.FromParsingName("shell:::{E88865EA-0E1C-4E20-9AA6-EDCD0212C87C}", Nothing, False))
                Shell.AddSpecialFolder(SpecialFolders.RecycleBin, Folder.FromParsingName("shell:::{645FF040-5081-101B-9F08-00AA002F954E}", Nothing, False))
                Shell.AddSpecialFolder(SpecialFolders.Recent, Folder.FromParsingName(Environment.ExpandEnvironmentVariables("%APPDATA%\Microsoft\Windows\Recent"), Nothing, False))
                '                                 .GetItems().First(Function(i) i.FullPath.EndsWith("\Recent")))
                Shell.AddSpecialFolder(SpecialFolders.OneDrive, Folder.FromParsingName("shell:::{018D5C66-4533-4307-9B53-224DE2ED1FE6}", Nothing, False))
                Shell.AddSpecialFolder(SpecialFolders.OneDriveBusiness, Folder.FromParsingName("shell:::{04271989-C4D2-BEC7-A521-3DF166FAB4BA}", Nothing, False))
                Shell.AddSpecialFolder(SpecialFolders.WindowsTools, Folder.FromParsingName("shell:::{D20EA4E1-3957-11D2-A40B-0C5020524153}", Nothing, False))
                Shell.AddSpecialFolder(SpecialFolders.Libraries, Folder.FromParsingName("shell:::{031E4825-7B94-4DC3-B131-E946B44C8DD5}", Nothing, False))
                'addSpecialFolder("User Pinned", Folder.FromParsingName("shell:::{1F3427C8-5C10-4210-AA03-2EE45287D668}", Nothing, False))
                Shell.AddSpecialFolder(SpecialFolders.ControlPanel, Folder.FromParsingName("shell:::{26EE0668-A00A-44D7-9371-BEB064C98683}", Nothing))
                Shell.AddSpecialFolder(SpecialFolders.DevicesAndPrinters, Folder.FromParsingName("shell:::{A8A91A66-3A7D-4424-8D24-04E180695C7A}", Nothing, False))
                Shell.AddSpecialFolder(SpecialFolders.AllTasks, Folder.FromParsingName("shell:::{ED7BA470-8E54-465E-825C-99712043E01C}", Nothing, False))
                '_specialFolders.Add("Applications", Folder.FromParsingName("shell:::{4234d49b-0245-4df3-b780-3893943456e1}", Nothing))
                'Shell.AddSpecialFolder("Frequent Folders", Folder.FromParsingName("shell:::{3936E9E4-D92C-4EEE-A85A-BC16D5EA0819}", Nothing, False))
                Shell.AddSpecialFolder(SpecialFolders.UserProfile, Folder.FromParsingName(Environment.ExpandEnvironmentVariables("%USERPROFILE%"), Nothing, False))
                '_specialFolders.Add("Installed Updates", Folder.FromParsingName("shell:::{d450a8a1-9568-45c7-9c0e-b4f9fb4537bd}", Nothing))
                'addSpecialFolder("Network Connections", Folder.FromParsingName("shell:::{7007ACC7-3202-11D1-AAD2-00805FC1270E}", Nothing, False))
                Shell.AddSpecialFolder(SpecialFolders.ProgramsAndFeatures, Folder.FromParsingName("shell:::{7b81be6a-ce2b-4676-a29e-eb907a5126c5}", Nothing, False))
                '_specialFolders.Add("Public", Folder.FromParsingName("shell:::{4336a54d-038b-4685-ab02-99bb52d3fb8b}", Nothing))
                '_specialFolders.Add("Recent Items", Folder.FromParsingName("shell:::{4564b25e-30cd-4787-82ba-39e73a750b14}", Nothing))

                Shell.IsSpecialFoldersReady.Set()

                Shell.StartListening(Shell.Desktop)
            End Sub)

        ' show debug window?
        If _isDebugVisible Then
            _debugWindow = New DebugTools.DebugWindow()
            _debugWindow.Show()
        End If
    End Sub

    Private Shared Sub window_Loaded(sender As Object, e As EventArgs)
        ' whenever a window is loaded, we'll monitor for it closing so we can gracefully shutdown
        RemoveHandler CType(sender, Window).Closed, AddressOf window_Closed
        RemoveHandler CType(sender, Window).Closing, AddressOf window_Closing
        AddHandler CType(sender, Window).Closed, AddressOf window_Closed
        AddHandler CType(sender, Window).Closing, AddressOf window_Closing
    End Sub

    Private Shared Sub window_Closing(sender As Object, e As EventArgs)
        ' a window is about to close - was it the main window?
        _mainWindow = Application.Current.MainWindow
    End Sub

    Private Shared Sub window_Closed(sender As Object, e As EventArgs)
        If System.Windows.Application.Current.ShutdownMode = ShutdownMode.OnLastWindowClose _
            AndAlso System.Windows.Application.Current.Windows.Cast(Of Window) _
                .Where(Function(w) Not w.GetType().ToString().StartsWith("Microsoft.VisualStudio.")).Count <= 1 Then
            ' if ShutDownMode is OnLastWindowClose and this is the last window closing...
            Shell.Shutdown()
        ElseIf System.Windows.Application.Current.ShutdownMode = ShutdownMode.OnMainWindowClose _
            AndAlso sender.Equals(_mainWindow) Then
            ' if ShutDownMode is OnMainWindowClose and this is the main window closing...
            Shell.Shutdown()
        End If
    End Sub

    ''' <summary>
    ''' Shut down Laila.Shell.
    ''' </summary>
    Public Shared Sub Shutdown()
        If Not Shell.ShuttingDownToken.IsCancellationRequested Then
            ' notify we're shutting down to quit all ongoing processes 
            ' and prevent this function from getting called twice
            _shutDownTokensSource.Cancel()

            ' stop listening for clipboard changes
            Functions.ChangeClipboardChain(_hwnd, _nextClipboardViewer)

            ' stop receiving and handling notifications
            SyncLock _listenersLock
                For Each item In _listenerhNotifies.ToList()
                    Functions.SHChangeNotifyDeregister(item.Value)
                    _listenerhNotifies.Remove(item.Key)
                    _listenerCount.Remove(item.Key)
                Next
                For Each item In _listenerFileSystemWatchers.ToList()
                    item.Value.Dispose()
                    _listenerFileSystemWatchers.Remove(item.Key)
                Next
            End SyncLock

            Shell.NotificationMainThread.DisposeAndWait()
            Shell.NotificationThreadPool.DisposeAndWait()

            ' stop monitoring changes to settings so we can properly close the registry keys
            Shell.Settings.StopMonitoring()

            ' we need to clean up the menus because there is a hMenu which would block proper application 
            ' shutdown if it wasn't disposed of
            SyncLock _menuCacheLock
                For Each item In Shell.MenuCache.ToList()
                    item.Dispose()
                Next
            End SyncLock

            ' we clean up the controls, especially to revoke the drag and drop
            Dim clist As List(Of IDisposable)
            SyncLock _controlCache
                clist = _controlCache.ToList()
            End SyncLock
            While Not clist.Count = 0
                For Each item In clist
                    item.Dispose()
                Next
                SyncLock _controlCache
                    clist = _controlCache.ToList()
                End SyncLock
            End While

            ' items need not really be disposed of, because they have no managed resources,
            ' but we do it anyway, in case that changes in the future
            Dim c As Integer = 0
            _itemsCacheLock.Wait()
            Try
                For Each item In _itemsCache.Select(Function(i) i.Item1).ToList()
                    item.Dispose()
                    c += 1
                Next
            Finally
                _itemsCacheLock.Release()
            End Try

            Dim poolList As List(Of Helpers.ThreadPool) = Nothing
            SyncLock _threadPoolCacheLock
                poolList = Shell.ThreadPoolCache.ToList()
            End SyncLock
            For Each item In poolList
                item.DisposeAndWait()
            Next

            ' uninitialize ole
            Functions.OleUninitialize()

            ' close messaging window to allow the application to close when using ShutdownMode.OnLastWindowClose
            Debug.WriteLine("Closing messaging window")
            If Not _w Is Nothing Then _w.Close()
        End If
    End Sub

    Private Shared Function hwndHook(hwnd As IntPtr, msg As Integer, wParam As IntPtr, lParam As IntPtr, ByRef handled As Boolean) As IntPtr
        'Debug.WriteLine(CType(msg, WM).ToString())
        Select Case msg
            Case WM.USER + 1
                ' we received an SHChangeNotify message - go get the data
                Dim pppidl As IntPtr = IntPtr.Zero
                Dim lEvent As SHCNE = 0
                Dim hLock As IntPtr = Functions.SHChangeNotification_Lock(wParam, lParam, pppidl, lEvent)

                If hLock <> IntPtr.Zero Then
                    ' read pidls
                    Dim pidl1 As IntPtr = Marshal.ReadIntPtr(pppidl)
                    pppidl = IntPtr.Add(pppidl, IntPtr.Size)
                    Dim pidl2 As IntPtr = Marshal.ReadIntPtr(pppidl)

                    Shell.NotificationMainThread.Add(
                        Sub()
                            ' make eventargs
                            Dim e As NotificationEventArgs = New NotificationEventArgs() With {
                                .[Event] = lEvent
                            }

                            Dim text As String = lEvent.ToString() & "  w=" & wParam.ToString() & "  l=" & lParam.ToString() & Environment.NewLine

                            Select Case lEvent
                                Case SHCNE.ATTRIBUTES, SHCNE.CREATE, SHCNE.DELETE, SHCNE.DRIVEADD,
                                     SHCNE.DRIVEREMOVED, SHCNE.MEDIAINSERTED, SHCNE.MEDIAREMOVED,
                                     SHCNE.MKDIR, SHCNE.NETSHARE, SHCNE.NETUNSHARE, SHCNE.RENAMEFOLDER, SHCNE.RENAMEITEM,
                                     SHCNE.RMDIR, SHCNE.SERVERDISCONNECT, SHCNE.UPDATEDIR, SHCNE.UPDATEIMAGE, SHCNE.UPDATEITEM
                                    If Not IntPtr.Zero.Equals(pidl1) Then
                                        Using p As Pidl = New Pidl(pidl1).Clone()
                                            e.Item1 = Item.FromPidl(p, Nothing, False, False)
                                            If Not e.Item1 Is Nothing Then
                                                text &= BitConverter.ToString(e.Item1.Pidl.Bytes) & Environment.NewLine & e.Item1.DisplayName & " (" & e.Item1.FullPath & ")" & Environment.NewLine
                                            Else
                                                text &= BitConverter.ToString(p.Bytes) & Environment.NewLine & "(item not found)" & Environment.NewLine
                                            End If
                                        End Using
                                    End If
                                    If Not IntPtr.Zero.Equals(pidl2) Then
                                        Using p As Pidl = New Pidl(pidl2).Clone()
                                            e.Item2 = Item.FromPidl(p, Nothing, False, False)
                                            If Not e.Item2 Is Nothing Then
                                                text &= BitConverter.ToString(e.Item2.Pidl.Bytes) & Environment.NewLine & e.Item2.DisplayName & " (" & e.Item2.FullPath & ")" & Environment.NewLine
                                            Else
                                                text &= BitConverter.ToString(p.Bytes) & Environment.NewLine & "(item not found)" & Environment.NewLine
                                            End If
                                        End Using
                                    End If
                            End Select

                            ' notify children?
                            Dim isFileSystemNotification As Boolean =
                                e.Event = SHCNE.MKDIR OrElse e.Event = SHCNE.CREATE OrElse e.Event = SHCNE.RMDIR _
                                OrElse e.Event = SHCNE.DELETE OrElse e.Event = SHCNE.RENAMEFOLDER OrElse e.Event = SHCNE.RENAMEITEM

                            Dim isFileSystemItem As Boolean = True
                            If isFileSystemNotification Then
                                Dim parent As Item = e.Item1
                                While Not parent Is Nothing AndAlso isFileSystemItem
                                    isFileSystemItem = isFileSystemItem _
                                        AndAlso parent.Attributes.HasFlag(SFGAO.FILESYSTEM) _
                                        AndAlso e.Item1.Attributes.HasFlag(SFGAO.STORAGEANCESTOR)
                                    If isFileSystemItem Then parent = parent.Parent
                                End While
                            End If
                            If e.Item1 Is Nothing OrElse Not isFileSystemNotification OrElse Not isFileSystemItem Then
                                Debug.Write(text)

                                ' notify children
                                notifySubscribers(e)
                            End If

                            If Not e.IsHandled1 AndAlso Not e.Item1 Is Nothing Then e.Item1.Dispose()
                            If Not e.IsHandled2 AndAlso Not e.Item2 Is Nothing Then e.Item2.Dispose()

                            ' unlock
                            Functions.SHChangeNotification_Unlock(hLock)
                        End Sub)
                End If
            Case WM.SETTINGCHANGE
                Shell.Settings.OnSettingChange()
            Case WM.DRAWCLIPBOARD
                RaiseEvent ClipboardChanged(Nothing, New EventArgs())
            Case WM.CHANGECBCHAIN
                If wParam = _nextClipboardViewer Then
                    _nextClipboardViewer = lParam
                ElseIf _nextClipboardViewer <> IntPtr.Zero Then
                    Functions.SendMessage(_nextClipboardViewer, msg, wParam, lParam)
                End If
        End Select
    End Function

    Private Shared Sub notifySubscribers(e As NotificationEventArgs)
        Dim list As List(Of IProcessNotifications) = Nothing
        SyncLock _notificationSubscribersLock
            list = _notificationSubscribers.Where(Function(i) i.IsProcessingNotifications).ToList()
        End SyncLock
        If list.Count > 0 Then
            Dim size As Integer = Math.Max(1, Math.Min(list.Count / 10, 250))
            Dim chuncks()() As IProcessNotifications = list.Chunk(list.Count / size).ToArray()
            Dim tcses As List(Of TaskCompletionSource) = New List(Of TaskCompletionSource)()

            ' threads for refreshing
            For i = 0 To chuncks.Count - 1
                Dim j As Integer = i
                Dim tcs As TaskCompletionSource = New TaskCompletionSource()
                tcses.Add(tcs)
                Shell.NotificationThreadPool.Add(
                    Sub()
                        ' Process tasks from the queue
                        For Each item In chuncks(j)
                            If Shell.ShuttingDownToken.IsCancellationRequested Then Exit For
                            item?.ProcessNotification(e)
                        Next

                        tcs.SetResult()
                    End Sub)
            Next

            Task.WaitAll(tcses.Select(Function(tcs) tcs.Task).ToArray(), Shell.ShuttingDownToken)
        End If
    End Sub

    Public Shared Sub BlockDisposer(value As Boolean)
        _disposerBlockLock.Wait()
        Try
            If value Then
                _disposerBlock += 1
            Else
                _disposerBlock -= 1
            End If
            If _disposerBlock < 0 Then _disposerBlock = 0
        Finally
            _disposerBlockLock.Release()
        End Try
    End Sub

    Public Shared Sub AddSpecialFolder(specialFolder As SpecialFolders, item As Item)
        If Not item Is Nothing AndAlso TypeOf item Is Folder Then
            _specialFolders.Add(specialFolder, item)
        End If
    End Sub

    ''' <summary>
    ''' Gets the given special folder.
    ''' </summary>
    ''' <param name="id">The id of the special folder</param>
    ''' <returns>A folder object for the special folder</returns>
    Public Shared Function GetSpecialFolder(specialFolder As SpecialFolders) As Folder
        Shell.IsSpecialFoldersReady.WaitOne()
        If _specialFolders.ContainsKey(specialFolder) Then
            Return _specialFolders(specialFolder)
        Else
            Return Nothing
        End If
    End Function

    ''' <summary>
    ''' Gets a dictionary of all special folders.
    ''' </summary>
    ''' <returns>A dictionary containing all special folders</returns>
    Public Shared Function GetSpecialFolders() As Dictionary(Of SpecialFolders, Folder)
        Shell.IsSpecialFoldersReady.WaitOne()
        Return _specialFolders
    End Function

    Public Shared ReadOnly Property CustomFolders As List(Of CustomFolder)
        Get
            Return _customFolders
        End Get
    End Property

    Public Shared Function GetCustomProperty(key As PROPERTYKEY) As Type
        Dim t As Type = Nothing
        If _customPropertiesByKey.TryGetValue(key.ToString(), t) Then
            Return t
        Else
            Return Nothing
        End If
    End Function

    Public Shared Function GetCustomProperty(canonicalName As String) As Type
        Dim t As Type = Nothing
        If _customPropertiesByCanonicalName.TryGetValue(canonicalName, t) Then
            Return t
        Else
            Return Nothing
        End If
    End Function

    ''' <summary>
    ''' Gets the desktop folder, the root.
    ''' </summary>
    ''' <returns>A Folder object for the desktop folder</returns>
    Public Shared ReadOnly Property Desktop As Folder
        Get
            If _desktop Is Nothing Then
                _desktop = Shell.GetSpecialFolder(SpecialFolders.Desktop)
            End If

            Return _desktop
        End Get
    End Property

    ''' <summary>
    ''' Gets an object for overriding the mouse cursor.
    ''' </summary>
    ''' <returns>An IDiposable that overrides the mouse cursor and sets it back when disposed</returns>
    Public Shared Property OverrideCursor As Func(Of Cursor, IDisposable)
        Get
            Return _overrideCursorFunc
        End Get
        Set(value As Func(Of Cursor, IDisposable))
            _overrideCursorFunc = value
        End Set
    End Property

    Public Shared ReadOnly Property Settings As Settings
        Get
            Return _settings
        End Get
    End Property

    Friend Shared ReadOnly Property ItemsCache As List(Of Tuple(Of Item, DateTime))
        Get
            Return _itemsCache
        End Get
    End Property

    Friend Shared Sub AddToItemsCache(item As Item)
        _itemsCacheLock.Wait()
        Try
            _itemsCache.Add(New Tuple(Of Item, Date)(item, DateTime.Now))
            If item.Attributes.HasFlag(SFGAO.FILESYSTEM) AndAlso item.Attributes.HasFlag(SFGAO.STORAGEANCESTOR) Then
                AddToFileSystemCache(item)
            End If
        Finally
            _itemsCacheLock.Release()
        End Try
    End Sub

    Friend Shared Sub RemoveFromItemsCache(item As Item)
        _itemsCacheLock.Wait()
        Try
            _itemsCache.Remove(_itemsCache.FirstOrDefault(Function(i) item.Equals(i.Item1)))
            If item.Attributes.HasFlag(SFGAO.FILESYSTEM) AndAlso item.Attributes.HasFlag(SFGAO.STORAGEANCESTOR) Then
                RemoveFromFileSystemCache(item.FullPath)
            End If
        Finally
            _itemsCacheLock.Release()
        End Try
    End Sub

    Friend Shared Sub AddToFileSystemCache(item As Item)
        _fileSystemCacheLock.Wait()
        Try
            Dim fullPath As String = item?.FullPath
            If item.Attributes.HasFlag(SFGAO.FILESYSTEM) AndAlso item.Attributes.HasFlag(SFGAO.STORAGEANCESTOR) _
                AndAlso Not String.IsNullOrWhiteSpace(fullPath) Then
                If Not _fileSystemCache.ContainsKey(fullPath) Then
                    _fileSystemCache.Add(fullPath, item)
                    _fileSystemCacheCount.Add(fullPath, 1)
                Else
                    Dim count As Integer = _fileSystemCacheCount(fullPath)
                    _fileSystemCacheCount(fullPath) = count + 1
                End If
            End If
        Finally
            _fileSystemCacheLock.Release()
        End Try
    End Sub

    Friend Shared Sub RemoveFromFileSystemCache(fullPath As String)
        _fileSystemCacheLock.Wait()
        Try
            If Not String.IsNullOrWhiteSpace(fullPath) Then
                If _fileSystemCacheCount(fullPath) = 1 Then
                    _fileSystemCache.Remove(fullPath)
                    _fileSystemCacheCount.Remove(fullPath)
                ElseIf _fileSystemCacheCount.Contains(fullPath) Then
                    Dim count As Integer = _fileSystemCacheCount(fullPath)
                    _fileSystemCacheCount(fullPath) = count - 1
                End If
            End If
        Finally
            _fileSystemCacheLock.Release()
        End Try
    End Sub

    Friend Shared Sub UpdateFileSystemCache(oldFullPath As String, item As Item)
        _itemsCacheLock.Wait()
        Try
            RemoveFromFileSystemCache(oldFullPath)
            If item.Attributes.HasFlag(SFGAO.FILESYSTEM) AndAlso item.Attributes.HasFlag(SFGAO.STORAGEANCESTOR) Then
                AddToFileSystemCache(item)
            End If
        Finally
            _itemsCacheLock.Release()
        End Try
    End Sub

    Friend Shared Sub AddToControlCache(item As IDisposable)
        SyncLock _controlCacheLock
            _controlCache.Add(item)
        End SyncLock
    End Sub

    Friend Shared Sub RemoveFromControlCache(item As IDisposable)
        SyncLock _controlCacheLock
            _controlCache.Remove(item)
        End SyncLock
    End Sub

    Friend Shared ReadOnly Property ThreadPoolCache As List(Of Helpers.ThreadPool)
        Get
            Return _threadPoolCache
        End Get
    End Property

    Friend Shared Sub AddToThreadPoolCache(item As Helpers.ThreadPool)
        SyncLock _threadPoolCacheLock
            _threadPoolCache.Add(item)
        End SyncLock
    End Sub

    Friend Shared Sub RemoveFromThreadPoolCache(item As Helpers.ThreadPool)
        SyncLock _threadPoolCacheLock
            _threadPoolCache.Remove(item)
        End SyncLock
    End Sub

    Friend Shared ReadOnly Property MenuCache As List(Of BaseMenu)
        Get
            Return _menuCache
        End Get
    End Property

    Friend Shared Sub AddToMenuCache(item As BaseMenu)
        SyncLock _menuCacheLock
            _menuCache.Add(item)
        End SyncLock
    End Sub

    Friend Shared Sub RemoveFromMenuCache(item As BaseMenu)
        SyncLock _menuCacheLock
            _menuCache.Remove(item)
        End SyncLock
    End Sub

    Private Shared _listenersLock As Object = New Object()
    Private Shared _listenerhNotifies As Dictionary(Of String, UInt32) = New Dictionary(Of String, UInt32)()
    Private Shared _listenerCount As Dictionary(Of String, Integer) = New Dictionary(Of String, Integer)()
    Private Shared _listenerFileSystemWatchers As Dictionary(Of String, FileSystemWatcher) = New Dictionary(Of String, FileSystemWatcher)()

    Public Shared Sub StartListening(folder As Folder)
        If Shell.ShuttingDownToken.IsCancellationRequested Then
            Return
        End If

        Dim mustStart As Boolean
        SyncLock _listenersLock
            mustStart = Not _listenerCount.ContainsKey(folder.FullPath)
        End SyncLock

        If mustStart Then
            ' start receiving notifications
            Dim entry(0) As SHChangeNotifyEntry
            entry(0).pIdl = folder.Pidl.AbsolutePIDL
            entry(0).Recursively = True  ' folder.Pidl?.ToString.Equals("00-00")

            Dim fswNotify As Action(Of NotificationEventArgs) =
                Sub(e As NotificationEventArgs)
                    Dim text As String = "FSW: " & e.Event.ToString() & Environment.NewLine
                    If Not e.Item1 Is Nothing Then
                        text &= If(Not e.Item1.Pidl Is Nothing, BitConverter.ToString(e.Item1.Pidl.Bytes), "PIDL not available") & Environment.NewLine & e.Item1.DisplayName & " (" & e.Item1.FullPath & ")" & Environment.NewLine
                    Else
                        text &= "Item1 not found -- cancelling notification"
                    End If
                    If Not e.Item2 Is Nothing Then
                        text &= If(Not e.Item2.Pidl Is Nothing, BitConverter.ToString(e.Item2.Pidl.Bytes), "PIDL not available") & Environment.NewLine & e.Item2.DisplayName & " (" & e.Item2.FullPath & ")" & Environment.NewLine
                    ElseIf e.Event = SHCNE.RENAMEITEM OrElse e.Event = SHCNE.RENAMEFOLDER Then
                        text &= "Item2 not found -- cancelling notification"
                    End If
                    Debug.Write(text)
                    If Not e.Item1 Is Nothing AndAlso
                        (Not e.Item2 Is Nothing OrElse (e.Event <> SHCNE.RENAMEITEM AndAlso e.Event <> SHCNE.RENAMEFOLDER)) Then
                        ' notify components
                        notifySubscribers(e)

                        ' dispose of items
                        If Not e.IsHandled1 AndAlso Not e.Item1 Is Nothing Then e.Item1.Dispose()
                        If Not e.IsHandled2 AndAlso Not e.Item2 Is Nothing Then e.Item2.Dispose()
                    End If
                End Sub

            Dim fsw As FileSystemWatcher = Nothing, hNotify As UInt32?
            If folder.Attributes.HasFlag(SFGAO.STORAGEANCESTOR) AndAlso folder.Attributes.HasFlag(SFGAO.FILESYSTEM) Then
                fsw = New FileSystemWatcher(folder.FullPath)
                AddHandler fsw.Created,
                    Sub(s As Object, e As FileSystemEventArgs)
                        Shell.NotificationMainThread.Add(
                            Sub()
                                ' make eventargs
                                Dim e2 As NotificationEventArgs = New NotificationEventArgs()
                                e2.Event = SHCNE.CREATE
                                e2.Item1 = Item.FromParsingName(e.FullPath, Nothing, False, False)
                                If If(e2.Item1?.Attributes.HasFlag(SFGAO.STORAGEANCESTOR), False) Then fswNotify(e2)
                            End Sub)
                    End Sub
                AddHandler fsw.Deleted,
                    Sub(s As Object, e As FileSystemEventArgs)
                        Shell.NotificationMainThread.Add(
                            Sub()
                                ' make eventargs
                                Dim e2 As NotificationEventArgs = New NotificationEventArgs()
                                e2.Event = SHCNE.DELETE
                                Dim i As Item = Nothing
                                _fileSystemCacheLock.Wait()
                                Try
                                    If _fileSystemCache.ContainsKey(e.FullPath) Then
                                        i = _fileSystemCache(e.FullPath)
                                    End If
                                Finally
                                    _fileSystemCacheLock.Release()
                                End Try
                                If If(i?.Attributes.HasFlag(SFGAO.STORAGEANCESTOR), False) Then
                                    e2.Item1 = i?.Clone()
                                    If e2.Item1 Is Nothing Then
                                        e2.Item1 = New Item(e.FullPath)
                                    End If
                                    fswNotify(e2)
                                End If
                            End Sub)
                    End Sub
                AddHandler fsw.Renamed,
                    Sub(s As Object, e As RenamedEventArgs)
                        Shell.NotificationMainThread.Add(
                            Sub()
                                ' make eventargs
                                Dim e2 As NotificationEventArgs = New NotificationEventArgs()
                                e2.Event = SHCNE.RENAMEITEM
                                Dim i As Item = Nothing
                                _fileSystemCacheLock.Wait()
                                Try
                                    If _fileSystemCache.ContainsKey(e.OldFullPath) Then
                                        i = _fileSystemCache(e.OldFullPath)
                                    End If
                                Finally
                                    _fileSystemCacheLock.Release()
                                End Try
                                If If(i?.Attributes.HasFlag(SFGAO.STORAGEANCESTOR), False) Then
                                    e2.Item1 = i?.Clone()
                                    If e2.Item1 Is Nothing Then
                                        e2.Item1 = New Item(e.OldFullPath)
                                    End If
                                    e2.Item2 = Item.FromParsingName(e.FullPath, Nothing, False, False)
                                    If TypeOf e2.Item2 Is Folder Then e2.Event = SHCNE.RENAMEFOLDER
                                    If e2.Item2 Is Nothing Then e2.Event = SHCNE.DELETE
                                    If e2.Event = SHCNE.DELETE OrElse If(e2.Item2?.Attributes.HasFlag(SFGAO.STORAGEANCESTOR), False) Then
                                        fswNotify(e2)
                                    End If
                                End If
                            End Sub)
                    End Sub

                Debug.WriteLine($"Start monitoring by FSW {folder.FullPath}")
                Try
                    fsw.EnableRaisingEvents = True
                Catch ex As Exception
                    Debug.WriteLine(ex.Message)
                    fsw.Dispose()
                    fsw = Nothing
                End Try
            End If

            hNotify =
                Functions.SHChangeNotifyRegister(
                    _hwnd,
                    SHCNRF.NewDelivery Or SHCNRF.InterruptLevel Or SHCNRF.ShellLevel,
                    SHCNE.ALLEVENTS,
                    WM.USER + 1,
                    1,
                    entry)

            SyncLock _listenersLock
                If hNotify.HasValue Then _listenerhNotifies.Add(folder.FullPath, hNotify)
                _listenerCount.Add(folder.FullPath, 1)
                If Not fsw Is Nothing Then _listenerFileSystemWatchers.Add(folder.FullPath, fsw)
            End SyncLock
        Else
            SyncLock _listenersLock
                _listenerCount(folder.FullPath) += 1
            End SyncLock
        End If
    End Sub

    Public Shared Sub StopListening(folder As Folder)
        SyncLock _listenersLock
            Dim count As Integer
            If _listenerCount.TryGetValue(folder.FullPath, count) Then
                If count = 1 Then
                    ' stop receiving notifications
                    If _listenerhNotifies.ContainsKey(folder.FullPath) Then
                        Functions.SHChangeNotifyDeregister(_listenerhNotifies(folder.FullPath))
                        _listenerhNotifies.Remove(folder.FullPath)
                    End If
                    If _listenerFileSystemWatchers.ContainsKey(folder.FullPath) Then
                        _listenerFileSystemWatchers(folder.FullPath).Dispose()
                        _listenerFileSystemWatchers.Remove(folder.FullPath)
                    End If
                    _listenerCount.Remove(folder.FullPath)
                Else
                    _listenerCount(folder.FullPath) -= 1
                End If
            End If
        End SyncLock
    End Sub

    Public Shared Sub SubscribeToNotifications(item As IProcessNotifications)
        SyncLock _notificationSubscribersLock
            _notificationSubscribers.Add(item)
        End SyncLock
    End Sub

    Public Shared Sub UnsubscribeFromNotifications(item As IProcessNotifications)
        SyncLock _notificationSubscribersLock
            _notificationSubscribers.Remove(item)
        End SyncLock
    End Sub

    Public Class CustomFolder
        Public Property FullPath As String
        Public Property ReplacesFullPath As String
        Public Property Type As Type
    End Class
End Class
