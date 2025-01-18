Imports System.Collections.Concurrent
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Threading
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Controls.Primitives
Imports System.Windows.Input
Imports System.Windows.Media
Imports System.Windows.Media.Imaging
Imports Laila.Shell.Events
Imports Laila.Shell.Helpers

Namespace Controls
    Public MustInherit Class BaseMenu
        Inherits ContextMenu
        Implements IDisposable

        Public Shared ReadOnly FolderProperty As DependencyProperty = DependencyProperty.Register("Folder", GetType(Folder), GetType(BaseMenu), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly SelectedItemsProperty As DependencyProperty = DependencyProperty.Register("SelectedItems", GetType(IEnumerable(Of Item)), GetType(BaseMenu), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, AddressOf OnSelectedItemsChanged))
        Public Shared ReadOnly IsDefaultOnlyProperty As DependencyProperty = DependencyProperty.Register("IsDefaultOnly", GetType(Boolean), GetType(BaseMenu), New FrameworkPropertyMetadata(False, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))

        Public Event CommandInvoked(sender As Object, e As CommandInvokedEventArgs)
        Public Event RenameRequest(sender As Object, e As RenameRequestEventArgs)

        Public Property DefaultId As Tuple(Of Integer, String)

        Private _contextMenu As IContextMenu
        Private _contextMenu2 As IContextMenu2
        Private _contextMenu3 As IContextMenu3
        Private _hMenu As IntPtr
        Private _invokedId As Tuple(Of Integer, String)
        Private _renameRequestTimer As Timer
        Private _isWaitingForCreate As Boolean
        Private _wasMade As Boolean
        Private disposedValue As Boolean
        Private _staThread2 As Thread
        Private _taskQueue As New BlockingCollection(Of Action)
        Private _disposeTokensSource As CancellationTokenSource = New CancellationTokenSource()
        Private _disposeToken As CancellationToken = _disposeTokensSource.Token

        Public Sub New()
            Shell.AddToMenuCache(Me)

            _staThread2 = New Thread(
                Sub()
                    Try
                        ' Process tasks from the queue
                        Functions.OleInitialize(IntPtr.Zero)
                        For Each task In _taskQueue.GetConsumingEnumerable(_disposeToken)
                            task.Invoke()
                        Next
                    Catch ex As OperationCanceledException
                        Debug.WriteLine("Menu TaskQueue was canceled.")
                    End Try
                    Functions.OleUninitialize()
                End Sub)
            _staThread2.IsBackground = True
            _staThread2.SetApartmentState(ApartmentState.STA)
            _staThread2.Start()

            AddHandler Shell.Notification, AddressOf shell_Notification
        End Sub

        Protected MustOverride Sub AddItems()

        Protected Overrides Sub OnOpened(e As RoutedEventArgs)
            Using Shell.OverrideCursor(Cursors.Wait)
                Me.Make()
                MyBase.OnOpened(e)
            End Using
        End Sub

        Protected Overridable Function DoRenameAfter(Tag As Tuple(Of Integer, String)) As Boolean
            Dim newMenuItem As MenuItem = Me.Items.Cast(Of Control) _
                .FirstOrDefault(Function(c) TypeOf c Is MenuItem _
                    AndAlso Not c.Tag Is Nothing _
                    AndAlso CType(c.Tag, Tuple(Of Integer, String)).Item2 = "New")

            Return Not newMenuItem Is Nothing AndAlso
                Not newMenuItem.Items.Cast(Of Control).FirstOrDefault(
                    Function(c) _invokedId.Equals(c.Tag)) Is Nothing
        End Function

        Public Sub Make()
            If _wasMade Then Return

            makeContextMenu(Me.Folder, Me.SelectedItems, Me.IsDefaultOnly)

            Dim osver As Version = Environment.OSVersion.Version
            Dim isWindows11 As Boolean = osver.Major = 10 AndAlso osver.Minor = 0 AndAlso osver.Build >= 22000

            ' make our menu
            Me.Items.Clear()
            Me.Buttons.Clear()

            Me.AddItems()

            AddHandler Me.Closed,
                Sub(s As Object, e As EventArgs)
                    If Not _invokedId Is Nothing Then
                        If Me.DoRenameAfter(_invokedId) Then
                            initializeRenameRequest()
                        End If
                        Me.InvokeCommand(_invokedId)
                        _invokedId = Nothing
                    End If
                End Sub

            ' wire items
            Dim wireMenuItems As Action(Of List(Of Control)) =
                Sub(wireItems As List(Of Control))
                    For Each c As Control In wireItems
                        If (TypeOf c Is MenuItem AndAlso CType(c, MenuItem).Items.Count = 0) _
                        OrElse TypeOf c Is ButtonBase Then
                            If TypeOf c.Tag Is Tuple(Of Integer, String) Then
                                If TypeOf c Is System.Windows.Controls.Button Then
                                    AddHandler CType(c, System.Windows.Controls.Button).Click, AddressOf menuItem_Click
                                ElseIf TypeOf c Is ToggleButton Then
                                    AddHandler CType(c, ToggleButton).Click, AddressOf menuItem_Click
                                ElseIf TypeOf c Is MenuItem Then
                                    AddHandler CType(c, MenuItem).Click, AddressOf menuItem_Click
                                End If
                            End If
                        ElseIf TypeOf c Is MenuItem Then
                            wireMenuItems(CType(c, MenuItem).Items.Cast(Of Control).ToList())
                        End If
                    Next
                End Sub
            wireMenuItems(Me.Items.Cast(Of Control).ToList())
            wireMenuItems(Me.Buttons.Cast(Of Control).ToList())

            _wasMade = True
        End Sub

        Protected Function MakeButton(tag As Tuple(Of Integer, String), toolTip As String) As Button
            Dim button As Button = New Button()
            Dim image As Image = New Image()
            image.Width = 16
            image.Height = 16
            image.Margin = New Thickness(2)
            Select Case tag.Item2
                Case "copy" : image.Source = New ImageSourceConverter().ConvertFromString("pack://application:,,,/Laila.Shell;component/Images/copy16.png")
                Case "cut" : image.Source = New ImageSourceConverter().ConvertFromString("pack://application:,,,/Laila.Shell;component/Images/cut16.png")
                Case "paste" : image.Source = New ImageSourceConverter().ConvertFromString("pack://application:,,,/Laila.Shell;component/Images/paste16.png")
                Case "rename" : image.Source = New ImageSourceConverter().ConvertFromString("pack://application:,,,/Laila.Shell;component/Images/rename16.png")
                Case "delete" : image.Source = New ImageSourceConverter().ConvertFromString("pack://application:,,,/Laila.Shell;component/Images/delete16.png")
            End Select
            button.Content = image
            button.ToolTip = toolTip
            button.Tag = tag
            Return button
        End Function

        Protected Function MakeToggleButton(tag As Tuple(Of Integer, String), toolTip As String, isChecked As Boolean) As ToggleButton
            Dim button As ToggleButton = New ToggleButton()
            Dim image As Image = New Image()
            image.Width = 16
            image.Height = 16
            image.Margin = New Thickness(2)
            Select Case tag.Item2
                Case "laila.shell.(un)pin" : image.Source = New ImageSourceConverter().ConvertFromString("pack://application:,,,/Laila.Shell;component/Images/pin16.png")
            End Select
            button.Content = image
            button.ToolTip = toolTip
            button.Tag = tag
            button.IsChecked = isChecked
            Return button
        End Function

        Protected Function getMenuItems() As List(Of Control)
            Return getMenuItems(_hMenu, -1)
        End Function

        Private Function getMenuItems(hMenu2 As IntPtr, parentIndex As Integer) As List(Of Control)
            Dim result As List(Of Control) = New List(Of Control)()

            If Not _contextMenu Is Nothing AndAlso Not IntPtr.Zero.Equals(_hMenu) Then
                Dim tcs2 As New TaskCompletionSource()
                _taskQueue.Add(
                    Sub()
                        Try
                            If parentIndex >= 0 Then
                                Dim lParam As Integer = (&HFFFF0000) Or (parentIndex And &HFFFF)
                                If Not _contextMenu3 Is Nothing Then
                                    Dim ptr3 As IntPtr
                                    Try
                                        Dim h As HRESULT = _contextMenu3.HandleMenuMsg2(WM.INITMENUPOPUP, hMenu2, lParam, ptr3)
                                        Debug.WriteLine("contextMenu3 returned" & h.ToString())
                                    Finally
                                        If Not IntPtr.Zero.Equals(ptr3) Then
                                            Marshal.Release(ptr3)
                                        End If
                                    End Try
                                ElseIf Not _contextMenu2 Is Nothing Then
                                    _contextMenu2.HandleMenuMsg(WM.INITMENUPOPUP, hMenu2, lParam)
                                End If
                            End If
                            tcs2.SetResult()
                        Catch ex As Exception
                            tcs2.SetException(ex)
                        End Try
                    End Sub)
                tcs2.Task.Wait()

                For i = 0 To Functions.GetMenuItemCount(hMenu2) - 1
                    Dim mii As MENUITEMINFO
                    mii.cbSize = CUInt(Marshal.SizeOf(mii))
                    mii.fMask = MIIM.MIIM_STRING
                    mii.dwTypeData = New String(" "c, 2050)
                    mii.cch = mii.dwTypeData.Length - 2
                    Functions.GetMenuItemInfo(hMenu2, i, True, mii)
                    Dim header As String = mii.dwTypeData.Substring(0, mii.cch)

                    Dim tcs3 As New TaskCompletionSource
                    Dim bitmapSource As BitmapSource

                    _taskQueue.Add(
                        Sub()
                            Try
                                mii = New MENUITEMINFO()
                                mii.cbSize = CUInt(Marshal.SizeOf(mii))
                                mii.fMask = MIIM.MIIM_BITMAP Or MIIM.MIIM_FTYPE Or MIIM.MIIM_CHECKMARKS
                                Functions.GetMenuItemInfo(hMenu2, i, True, mii)

                                If Not IntPtr.Zero.Equals(mii.hbmpItem) Then
                                    bitmapSource = Interop.Imaging.CreateBitmapSourceFromHBitmap(mii.hbmpItem, IntPtr.Zero, Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions())
                                    bitmapSource.Freeze()
                                ElseIf Not IntPtr.Zero.Equals(mii.hbmpUnchecked) Then
                                    bitmapSource = Interop.Imaging.CreateBitmapSourceFromHBitmap(mii.hbmpUnchecked, IntPtr.Zero, Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions())
                                    bitmapSource.Freeze()
                                Else
                                    bitmapSource = Nothing
                                End If
                                tcs3.SetResult()
                            Catch ex As Exception
                                tcs3.SetException(ex)
                            End Try
                        End Sub)

                    tcs3.Task.Wait()

                    If mii.fType = MFT.SEPARATOR Then
                        ' refuse initial and double separators
                        If Not result.Count = 0 AndAlso Not TypeOf result(result.Count - 1) Is Separator Then
                            result.Add(New Separator())
                        End If
                    Else
                        mii = New MENUITEMINFO()
                        mii.cbSize = CUInt(Marshal.SizeOf(mii))
                        mii.fMask = MIIM.MIIM_ID
                        Functions.GetMenuItemInfo(hMenu2, i, True, mii)

                        Dim tcs As New TaskCompletionSource, verb As String, id As Integer
                        _taskQueue.Add(
                            Sub()
                                Try
                                    Dim cmd As StringBuilder = New StringBuilder()
                                    cmd.Append(New String(" ", 2050))
                                    If mii.wID >= 1 AndAlso mii.wID <= 99999 Then
                                        id = mii.wID - 1
                                        _contextMenu.GetCommandString(id, GCS.VERBW, 0, cmd, 2048)
                                        If cmd.Length = 0 Then
                                            _contextMenu.GetCommandString(id, GCS.VERBA, 0, cmd, 2048)
                                        End If
                                        verb = cmd.ToString()
                                    End If
                                    tcs.SetResult()
                                Catch ex As Exception
                                    tcs.SetException(ex)
                                End Try
                            End Sub)
                        tcs.Task.Wait()

                        Debug.WriteLine(header & "  " & id & vbTab & verb.ToString())

                        mii = New MENUITEMINFO()
                        mii.cbSize = CUInt(Marshal.SizeOf(mii))
                        mii.fMask = MIIM.MIIM_SUBMENU Or MIIM.MIIM_STATE
                        Functions.GetMenuItemInfo(hMenu2, i, True, mii)

                        Dim menuItem As MenuItem = New MenuItem() With {
                            .Header = header.Replace("&", "_"),
                            .Icon = New Image() With {.Source = bitmapSource},
                            .Tag = New Tuple(Of Integer, String)(id, verb),
                            .IsEnabled = If(CType(mii.fState, MFS).HasFlag(MFS.MFS_DISABLED), False, True),
                            .FontWeight = If(CType(mii.fState, MFS).HasFlag(MFS.MFS_DEFAULT), FontWeights.Bold, FontWeights.Normal)
                        }

                        If CBool(mii.fState And MFS.MFS_DEFAULT) Then
                            Me.DefaultId = menuItem.Tag
                        End If

                        If Not IntPtr.Zero.Equals(mii.hSubMenu) Then
                            Dim subMenu As List(Of Control) = getMenuItems(mii.hSubMenu, i)
                            For Each subMenuItem In subMenu
                                menuItem.Items.Add(subMenuItem)
                            Next
                        End If

                        result.Add(menuItem)
                    End If
                Next

                ' remove trailing separators
                While result.Count > 0 AndAlso TypeOf result(result.Count - 1) Is Separator
                    result.RemoveAt(result.Count - 1)
                End While
            End If

            Return result
        End Function

        Private Sub makeContextMenu(folder As Folder, items As IEnumerable(Of Item), isDefaultOnly As Boolean)
            Dim parentFolderPidl As Pidl
            If TypeOf folder Is SearchFolder Then
                parentFolderPidl = Shell.Desktop.Pidl.Clone()
            ElseIf Not folder.Parent Is Nothing Then
                parentFolderPidl = folder.Parent.Pidl.Clone()
            End If

            Dim tcs As New TaskCompletionSource()
            _taskQueue.Add(
                Sub()
                    Dim folderPidl As Pidl
                    Dim itemPidls As Pidl()
                    Dim flags As Integer = CMF.CMF_NORMAL
                    Dim shellFolder As IShellFolder

                    SyncLock folder._shellItemLock
                        If Not folder.disposedValue AndAlso Not folder.ShellItem2 Is Nothing Then
                            shellFolder = Folder.GetIShellFolderFromIShellItem2(folder.ShellItem2)
                        End If
                    End SyncLock

                    Try
                        If Not items Is Nothing AndAlso items.Count > 0 Then
                            ' user clicked on an item
                            flags = flags Or CMF.CMF_ITEMMENU

                            folderPidl = folder.Pidl.Clone()
                            itemPidls = items.Select(Function(i) i.Pidl.Clone()).ToArray()

                            CType(shellFolder, IShellFolderForIContextMenu).GetUIObjectOf _
                                (IntPtr.Zero, itemPidls.Length, itemPidls.Select(Function(p) p.RelativePIDL).ToArray(), GetType(IContextMenu).GUID, 0, _contextMenu)
                        Else
                            ' user clicked on the background
                            If folder.FullPath = Shell.Desktop.FullPath Then
                                ' this is the desktop
                                folderPidl = Shell.Desktop.Pidl.Clone()
                                itemPidls = {Shell.Desktop.Pidl.Clone()}
                            Else
                                ' this is any other folder
                                folderPidl = parentFolderPidl
                                itemPidls = {folder.Pidl.Clone()}
                            End If

                            CType(shellFolder, IShellFolderForIContextMenu).CreateViewObject _
                                (IntPtr.Zero, GetType(IContextMenu).GUID, _contextMenu)
                        End If

                        If Not _contextMenu Is Nothing Then
                            _contextMenu2 = TryCast(_contextMenu, IContextMenu2)
                            _contextMenu3 = TryCast(_contextMenu, IContextMenu3)
                        End If

                        If Not _contextMenu Is Nothing Then
                            _hMenu = Functions.CreatePopupMenu()
                            flags = flags Or CMF.CMF_EXTENDEDVERBS Or CMF.CMF_EXPLORE Or CMF.CMF_CANRENAME
                            If isDefaultOnly Then flags = flags Or CMF.CMF_DEFAULTONLY
                            _contextMenu.QueryContextMenu(_hMenu, 0, 1, 99999, flags)

                            Dim shellExtInit As IShellExtInit, dataObject As ComTypes.IDataObject
                            Try
                                shellExtInit = TryCast(_contextMenu, IShellExtInit)
                                If Not shellExtInit Is Nothing Then
                                    Functions.SHCreateDataObject(folderPidl.AbsolutePIDL, itemPidls.Count,
                                                     itemPidls.Select(Function(p) p.RelativePIDL).ToArray(),
                                                     IntPtr.Zero, GetType(ComTypes.IDataObject).GUID, dataObject)
                                    shellExtInit.Initialize(folder.Pidl.AbsolutePIDL, dataObject, IntPtr.Zero)
                                End If
                            Finally
                                If Not dataObject Is Nothing Then
                                    Marshal.ReleaseComObject(dataObject)
                                End If
                            End Try
                        End If
                    Finally
                        folderPidl.Dispose()
                        For Each p In itemPidls
                            p.Dispose()
                        Next
                        If Not shellFolder Is Nothing Then
                            Marshal.ReleaseComObject(shellFolder)
                        End If
                    End Try
                    tcs.SetResult()
                End Sub)
            tcs.Task.Wait()

            If Not parentFolderPidl Is Nothing Then
                parentFolderPidl.Dispose()
            End If
        End Sub

        Private Sub menuItem_Click(c As Control, e2 As EventArgs)
            invokeCommandDelayed(c.Tag)
            Me.IsOpen = False
        End Sub

        Protected Sub invokeCommandDelayed(id As Tuple(Of Integer, String))
            _invokedId = id
        End Sub

        Public Sub InvokeCommand(id As Tuple(Of Integer, String))
            If id Is Nothing Then Return

            Make()

            Dim folder As Folder
            Dim selectedItems As IEnumerable(Of Item)
            Dim e As CommandInvokedEventArgs

            UIHelper.OnUIThread(
                Sub()
                    folder = Me.Folder
                    selectedItems = Me.SelectedItems

                    e = New CommandInvokedEventArgs() With {
                        .Id = id.Item1,
                        .Verb = id.Item2
                    }
                    Dim c As Control = Me.Buttons.Cast(Of Control).FirstOrDefault(Function(i) id.Equals(i.Tag))
                    If TypeOf c Is ToggleButton Then e.IsChecked = CType(c, ToggleButton).IsChecked
                    RaiseEvent CommandInvoked(Me, e)

                    Select Case id.Item2
                        Case "Windows.ModernShare"
                            Menus.DoShare(selectedItems)
                        Case "copy"
                            Clipboard.CopyFiles(selectedItems)
                        Case "cut"
                            Clipboard.CutFiles(selectedItems)
                    End Select
                End Sub)

            If Not e.IsHandled Then
                Dim thread As Thread = New Thread(New ThreadStart(
                    Sub()
                        Select Case id.Item2
                            Case "Windows.ModernShare"
                            Case "copy"
                            Case "cut"
                            Case "paste"
                                Clipboard.PasteFiles(folder)
                            Case "delete"
                                Dim dataObject As IDataObject
                                dataObject = Clipboard.GetDataObjectFor(folder, selectedItems.ToList())
                                Dim fo As IFileOperation = Activator.CreateInstance(Type.GetTypeFromCLSID(Guids.CLSID_FileOperation))
                                If Keyboard.Modifiers.HasFlag(ModifierKeys.Shift) Then fo.SetOperationFlags(FOF.FOFX_WANTNUKEWARNING)
                                fo.DeleteItems(dataObject)
                                fo.PerformOperations()
                                Marshal.ReleaseComObject(fo)
                                Marshal.ReleaseComObject(dataObject)
                            Case Else
                                Dim cmi As New CMInvokeCommandInfoEx
                                Debug.WriteLine("InvokeCommand " & id.Item1 & ", " & id.Item2)
                                If id.Item1 > 0 Then
                                    cmi.lpVerb = New IntPtr(id.Item1)
                                    cmi.lpVerbW = New IntPtr(id.Item1)
                                Else
                                    cmi.lpVerb = Marshal.StringToHGlobalAnsi(id.Item2)
                                    cmi.lpVerbW = Marshal.StringToHGlobalUni(id.Item2)
                                End If
                                'cmi.lpDirectory = "c:\temp"
                                'cmi.lpDirectoryW = "c:\temp"
                                cmi.fMask = CMIC.UNICODE Or CMIC.ASYNCOK
                                If Keyboard.Modifiers.HasFlag(ModifierKeys.Control) Then cmi.fMask = cmi.fMask Or CMIC.CONTROL_DOWN
                                If Keyboard.Modifiers.HasFlag(ModifierKeys.Shift) Then cmi.fMask = cmi.fMask Or CMIC.SHIFT_DOWN
                                cmi.nShow = SW.SHOWNORMAL
                                'cmi.hwnd = Shell._hwnd
                                cmi.cbSize = CUInt(Marshal.SizeOf(cmi))

                                Dim h As HRESULT = _contextMenu.InvokeCommand(cmi)
                                Debug.WriteLine("InvokeCommand returned " & h.ToString())
                        End Select
                    End Sub))

                thread.SetApartmentState(ApartmentState.STA)
                thread.Start()
            End If
        End Sub

        Private Sub initializeRenameRequest()
            _isWaitingForCreate = True

            If Not _renameRequestTimer Is Nothing Then
                _renameRequestTimer.Dispose()
            End If

            _renameRequestTimer = New Timer(New TimerCallback(
                Sub()
                    UIHelper.OnUIThread(
                        Sub()
                            _isWaitingForCreate = False
                            If Not _renameRequestTimer Is Nothing Then
                                _renameRequestTimer.Dispose()
                                _renameRequestTimer = Nothing
                            End If
                        End Sub)
                End Sub), Nothing, 2500, Timeout.Infinite)
        End Sub

        Protected Async Sub shell_Notification(sender As Object, e As NotificationEventArgs)
            If Not disposedValue Then
                Select Case e.Event
                    Case SHCNE.CREATE, SHCNE.MKDIR
                        If _isWaitingForCreate Then
                            Dim e2 As RenameRequestEventArgs = New RenameRequestEventArgs()
                            e2.FullPath = e.Item1.FullPath
                            Await Task.Delay(250)
                            UIHelper.OnUIThread(
                                Sub()
                                    RaiseEvent RenameRequest(Me, e2)
                                    If e2.IsHandled Then
                                        _isWaitingForCreate = False
                                        If Not _renameRequestTimer Is Nothing Then
                                            _renameRequestTimer.Dispose()
                                            _renameRequestTimer = Nothing
                                        End If
                                    End If
                                End Sub)
                        End If
                End Select
            End If
        End Sub

        Public Property Folder As Folder
            Get
                Return GetValue(FolderProperty)
            End Get
            Set(value As Folder)
                SetValue(FolderProperty, value)
            End Set
        End Property

        Public Property SelectedItems As IEnumerable(Of Item)
            Get
                Return GetValue(SelectedItemsProperty)
            End Get
            Set(value As IEnumerable(Of Item))
                SetValue(SelectedItemsProperty, value)
            End Set
        End Property

        Public Property IsDefaultOnly As Boolean
            Get
                Return GetValue(IsDefaultOnlyProperty)
            End Get
            Set(value As Boolean)
                SetValue(IsDefaultOnlyProperty, value)
            End Set
        End Property

        Shared Sub OnSelectedItemsChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
            Dim bm As BaseMenu = d
            If bm.IsOpen Then bm.IsOpen = False
        End Sub

        Protected Overridable Sub Dispose(disposing As Boolean)
            If Not disposedValue Then
                disposedValue = True

                If disposing Then
                    ' dispose managed state (managed objects)
                    RemoveHandler Shell.Notification, AddressOf shell_Notification
                End If

                ' free unmanaged resources (unmanaged objects) and override finalizer
                ' set large fields to null
                If Not _contextMenu Is Nothing Then
                    Marshal.ReleaseComObject(_contextMenu)
                End If
                If Not _contextMenu2 Is Nothing Then
                    Marshal.ReleaseComObject(_contextMenu2)
                End If
                If Not _contextMenu3 Is Nothing Then
                    Marshal.ReleaseComObject(_contextMenu3)
                End If
                If Not IntPtr.Zero.Equals(_hMenu) Then
                    Functions.DestroyMenu(_hMenu)
                End If

                _disposeTokensSource.Cancel()

                Shell.RemoveFromMenuCache(Me)
            End If
        End Sub

        Public Sub Dispose() Implements IDisposable.Dispose
            ' Do not change this code. Put cleanup code in 'Dispose(disposing As Boolean)' method
            Dispose(disposing:=True)
            GC.SuppressFinalize(Me)
        End Sub
    End Class
End Namespace