Imports Laila.Shell.Events
Imports Laila.Shell.Helpers
Imports Laila.Shell.Interfaces
Imports Laila.Shell.Interop
Imports Laila.Shell.Interop.Items
Imports System.Runtime.InteropServices
Imports System.Threading
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Controls.Primitives
Imports System.Windows.Input
Imports System.Windows.Media
Imports System.Windows.Media.Imaging

Namespace Controls
    Public MustInherit Class BaseMenu
        Inherits ContextMenu
        Implements IDisposable, IProcessNotifications

        Public Shared ReadOnly FolderProperty As DependencyProperty = DependencyProperty.Register("Folder", GetType(Folder), GetType(BaseMenu), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly SelectedItemsProperty As DependencyProperty = DependencyProperty.Register("SelectedItems", GetType(IEnumerable(Of Item)), GetType(BaseMenu), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, AddressOf OnSelectedItemsChanged))
        Public Shared ReadOnly IsDefaultOnlyProperty As DependencyProperty = DependencyProperty.Register("IsDefaultOnly", GetType(Boolean), GetType(BaseMenu), New FrameworkPropertyMetadata(False, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))

        Public Event CommandInvoked(sender As Object, e As CommandInvokedEventArgs)
        Public Event RenameRequest(sender As Object, e As RenameRequestEventArgs)

        Public Property DefaultId As Tuple(Of Integer, String, Object)
        Public Property IsProcessingNotifications As Boolean = True Implements IProcessNotifications.IsProcessingNotifications

        Protected _activeItems As List(Of Item)
        Protected _hbitmapsToDispose As HashSet(Of IntPtr) = New HashSet(Of IntPtr)()
        Protected _invokedId As Tuple(Of Integer, String, Object)
        Protected _isWaitingForCreate As Boolean
        Protected _renameRequestTimer As Timer
        Protected _thread As Helpers.ThreadPool
        Protected _wasMade As Boolean
        Private _makeLock As Object = New Object()
        Protected disposedValue As Boolean

        Public Sub New()
            Shell.AddToMenuCache(Me)

            _thread = New Helpers.ThreadPool(1)

            Shell.SubscribeToNotifications(Me)
        End Sub

        Protected MustOverride Async Function AddItems() As Task

        Protected Overrides Async Sub OnOpened(e As RoutedEventArgs)
            Using Shell.OverrideCursor(Cursors.Wait)
                UIHelper.OnUIThread(
                    Sub()
                    End Sub, Threading.DispatcherPriority.Render)

                Await Me.Make()

                If Me.Items.Count > 0 Then
                    MyBase.OnOpened(e)
                Else
                    Me.IsOpen = False
                End If
            End Using
        End Sub

        Protected MustOverride Function DoRenameAfter(Tag As Tuple(Of Integer, String, Object)) As Boolean

        Public Overrides Async Function Make() As Task
            Using Shell.OverrideCursor(Cursors.Wait)
                SyncLock _makeLock
                    If _wasMade Then Return

                    _activeItems = If(Not Me.SelectedItems Is Nothing AndAlso Me.SelectedItems.Count > 0,
                    Me.SelectedItems.ToList(), New List(Of Item) From {Me.Folder})

                    Make(Me.Folder, Me.SelectedItems, Me.IsDefaultOnly)

                    ' make our menu
                    Me.Items.Clear()
                    Me.Buttons.Clear()

                    _wasMade = True
                End SyncLock

                Await Me.AddItems()
            End Using

            AddHandler Me.Closed,
                Sub(s As Object, e As EventArgs)
                    If Not _invokedId Is Nothing Then
                        If Me.DoRenameAfter(_invokedId) Then
                            initializeRenameRequest()
                        End If
                        Dim __ = Me.InvokeCommand(_invokedId)
                        _invokedId = Nothing
                    End If
                End Sub

            ' wire items
            Dim wireMenuItems As Action(Of List(Of Control)) =
                Sub(wireItems As List(Of Control))
                    For Each c As Control In wireItems
                        If (TypeOf c Is MenuItem AndAlso CType(c, MenuItem).Items.Count = 0) _
                        OrElse TypeOf c Is ButtonBase Then
                            If TypeOf c.Tag Is Tuple(Of Integer, String, Object) Then
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
        End Function

        Protected MustOverride Overloads Sub Make(folder As Folder, items As IEnumerable(Of Item), isDefaultOnly As Boolean)

        Protected Overridable Sub menuItem_Click(c As Control, e2 As EventArgs)
            invokeCommandDelayed(c.Tag)
            Me.IsOpen = False
        End Sub

        Protected Sub invokeCommandDelayed(id As Tuple(Of Integer, String, Object))
            If _invokedId Is Nothing Then
                _invokedId = id
            End If
        End Sub

        Public MustOverride Async Function InvokeCommand(id As Tuple(Of Integer, String, Object)) As Task

        Protected Sub RaiseCommandInvoked(e As CommandInvokedEventArgs)
            RaiseEvent CommandInvoked(Me, e)
        End Sub

        Protected Sub initializeRenameRequest()
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

        Protected Friend Overridable Async Sub ProcessNotification(e As NotificationEventArgs) Implements IProcessNotifications.ProcessNotification
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
                    Shell.UnsubscribeFromNotifications(Me)
                End If

                ' free unmanaged resources (unmanaged objects) and override finalizer
                ' set large fields to null

                ' first wait for thread to finish
                _thread.Dispose()

                ' dispose of bitmaps
                For Each hbitmap In _hbitmapsToDispose
                    Functions.DeleteObject(hbitmap)
                Next
                _hbitmapsToDispose.Clear()

                ' remove from cache
                Shell.RemoveFromMenuCache(Me)
            End If
        End Sub

        Public Sub Dispose() Implements IDisposable.Dispose
            ' Do not change this code. Put cleanup code in 'Dispose(disposing As Boolean)' method
            Dispose(disposing:=True)
            GC.SuppressFinalize(Me)
        End Sub

        Protected Class MenuItemData
            Public Property ApplicationName As String
            Public Property Header As String
            Public Property Icon As ImageSource
            Public Property Tag As Object
            Public Property IsEnabled As Boolean
            Public Property FontWeight As FontWeight
            Public Property Items As List(Of MenuItemData)
            Public Property ShortcutKeyText As String
        End Class
    End Class
End Namespace