Imports System.Collections.Specialized
Imports System.ComponentModel
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Threading
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Data
Imports Laila.Shell.Controls.Parts
Imports Laila.Shell.Events
Imports Laila.Shell.Helpers
Imports Laila.Shell.Interfaces
Imports Laila.Shell.Interop
Imports Laila.Shell.Interop.DragDrop
Imports Laila.Shell.Interop.Items
Imports Laila.Shell.Interop.Properties

Namespace Controls
    Public Class FolderView
        Inherits Control
        Implements IDisposable, IProcessNotifications

        Public Shared ReadOnly FolderProperty As DependencyProperty = DependencyProperty.Register("Folder", GetType(Folder), GetType(FolderView), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, AddressOf OnFolderChanged))
        Public Shared ReadOnly SelectedItemsProperty As DependencyProperty = DependencyProperty.Register("SelectedItems", GetType(IEnumerable(Of Item)), GetType(FolderView), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, AddressOf OnSelectedItemsChanged))
        Public Shared ReadOnly IsSelectingProperty As DependencyProperty = DependencyProperty.Register("IsSelecting", GetType(Boolean), GetType(FolderView), New FrameworkPropertyMetadata(False, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly StatusTextProperty As DependencyProperty = DependencyProperty.Register("StatusText", GetType(String), GetType(FolderView), New FrameworkPropertyMetadata(String.Empty, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly SearchBoxProperty As DependencyProperty = DependencyProperty.Register("SearchBox", GetType(SearchBox), GetType(FolderView), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly NavigationProperty As DependencyProperty = DependencyProperty.Register("Navigation", GetType(Navigation), GetType(FolderView), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))

        Shared Sub New()
            DefaultStyleKeyProperty.OverrideMetadata(GetType(FolderView), New FrameworkPropertyMetadata(GetType(FolderView)))
        End Sub

        Public Shared ReadOnly ViewActivatedEvent As RoutedEvent =
            EventManager.RegisterRoutedEvent(
                "ViewActivated", RoutingStrategy.Bubble,
                GetType(RoutedEventHandler), GetType(FolderView))

        Public Custom Event ViewActivated As RoutedEventHandler
            AddHandler(value As RoutedEventHandler)
                Me.AddHandler(ViewActivatedEvent, value)
            End AddHandler
            RemoveHandler(value As RoutedEventHandler)
                Me.RemoveHandler(ViewActivatedEvent, value)
            End RemoveHandler
            RaiseEvent(sender As Object, e As EventArgs)

            End RaiseEvent
        End Event

        Public Property IsProcessingNotifications As Boolean = True Implements IProcessNotifications.IsProcessingNotifications

        Private _views As Dictionary(Of Guid, Control)
        Private _activeView As BaseFolderView
        Private _dropTarget As IDropTarget
        Private PART_Grid As Grid
        Private _isLoaded As Boolean
        Private disposedValue As Boolean

        Public Sub New()
            ' register us in the control cache so we get disposed when the app is shutting down
            Shell.AddToControlCache(Me)

            AddHandler Me.Loaded,
                Sub(s As Object, e As EventArgs)
                    If Not _isLoaded Then
                        _isLoaded = True

                        ' wait for everything to load
                        UIHelper.OnUIThread(
                            Sub()
                            End Sub, Threading.DispatcherPriority.Loaded)

                        ' display any available folder
                        If Not Me.Folder Is Nothing Then
                            changeView(Me.Folder.ActiveView, Me.Folder)
                        End If

                        ' register us as a drop target
                        _dropTarget = New ListViewDropTarget(Me)
                        WpfDragTargetProxy.RegisterDragDrop(Me, _dropTarget)

                        ' when the window closes: dispose ourselves
                        AddHandler Window.GetWindow(Me).Closed,
                            Sub(s2 As Object, e2 As EventArgs)
                                Me.Dispose()
                            End Sub
                    End If
                End Sub

            ' subscribe to file and folder change notifications
            Shell.SubscribeToNotifications(Me)
        End Sub

        Protected Friend Overridable Sub ProcessNotification(e As NotificationEventArgs) Implements IProcessNotifications.ProcessNotification
            Select Case e.Event
                Case SHCNE.RENAMEITEM, SHCNE.RENAMEFOLDER
                    ' this is for supporting file operations within .zip files 
                    ' from explorer or 7-zip
                    Dim selectedItem As Item = Nothing
                    If selectedItem Is Nothing Then
                        UIHelper.OnUIThread(
                            Sub()
                                selectedItem = Me.Folder
                            End Sub)
                    End If
                    ' is something is being renamed to the currently selected folder...
                    If Not selectedItem Is Nothing _
                        AndAlso ((Not selectedItem.Pidl Is Nothing AndAlso Not e.Item2.Pidl Is Nothing AndAlso e.Item2.Pidl.Equals(selectedItem.Pidl)) _
                            OrElse ((selectedItem.Pidl Is Nothing OrElse e.Item2.Pidl Is Nothing) _
                                AndAlso (If(e.Item2.FullPath?.Equals(selectedItem.FullPath), False) OrElse If(e.Item2.FullPath?.Equals(selectedItem.FullPath.Split("~")(0)), False)))) Then
                        Shell.GlobalThreadPool.Add(
                            Sub()
                                ' get the first available parent in case the current folder disappears
                                Dim f As Folder = Me.GetParentOfSelectionBefore(selectedItem)
                                If Not f Is Nothing Then
                                    Thread.Sleep(300) ' wait for .zip operations/folder refresh to complete

                                    ' get the newly created .zip folder
                                    Dim replacement As Item = f.Items.ToList().LastOrDefault(Function(i) Not i Is Nothing AndAlso Not i.disposedValue _
                                        AndAlso ((Not i.Pidl Is Nothing AndAlso Not e.Item2.Pidl Is Nothing AndAlso If(i.Pidl?.Equals(e.Item2.Pidl), False)) _
                                            OrElse (i.Pidl Is Nothing OrElse e.Item2.Pidl Is Nothing) _
                                                AndAlso If(i.FullPath?.Equals(e.Item2.FullPath), False)))
                                    If Not replacement Is Nothing AndAlso TypeOf replacement Is Folder Then
                                        ' new folder matching the current folder was found -- load it
                                        CType(replacement, Folder).GetItems()
                                        UIHelper.OnUIThread(
                                            Sub()
                                                CType(replacement, Folder).LastScrollOffset = CType(selectedItem, Folder).LastScrollOffset
                                                CType(replacement, Folder).LastScrollSize = CType(selectedItem, Folder).LastScrollSize
                                                Me.Folder = replacement
                                            End Sub)
                                    Else
                                        ' the current folder disappeared -- switch to parent
                                        UIHelper.OnUIThread(
                                            Sub()
                                                Me.Folder = f
                                            End Sub)
                                    End If
                                End If
                            End Sub)
                    End If
                Case SHCNE.DELETE, SHCNE.RMDIR, SHCNE.DRIVEREMOVED
                    ' this sets the current folder to the first available parent 
                    ' when the current folder gets deleted
                    Dim selectedItem As Item = Nothing
                    If selectedItem Is Nothing Then
                        UIHelper.OnUIThread(
                            Sub()
                                selectedItem = Me.Folder
                            End Sub)
                    End If
                    ' if the current folder was deleted...
                    If Not selectedItem Is Nothing _
                        AndAlso ((Not selectedItem.Pidl Is Nothing AndAlso Not e.Item1.Pidl Is Nothing AndAlso e.Item1.Pidl.Equals(selectedItem.Pidl)) _
                            OrElse ((selectedItem.Pidl Is Nothing OrElse e.Item1.Pidl Is Nothing) _
                                AndAlso If(e.Item1.FullPath?.Equals(selectedItem.FullPath), False))) Then
                        UIHelper.OnUIThread(
                            Sub()
                                ' get the first available parent  
                                Dim f As Folder = selectedItem.LogicalParent
                                If Not f Is Nothing Then
                                    Me.Folder = f ' load it
                                End If
                            End Sub)
                    End If
            End Select
        End Sub

        Private Function GetParentOfSelectionBefore(folder As Item, Optional selectedItem As Folder = Nothing) As Folder
            ' get current folder
            If selectedItem Is Nothing Then
                UIHelper.OnUIThread(
                    Sub()
                        selectedItem = Me.Folder
                    End Sub)
            End If
            If selectedItem Is Nothing Then
                ' no current folder -- give up
                Return Nothing
            ElseIf selectedItem?.Pidl?.Equals(folder.Pidl) OrElse selectedItem.FullPath?.Equals(folder.FullPath) Then
                ' we found the current folder -- switch to it's parent
                Return selectedItem.LogicalParent
            ElseIf Not selectedItem?.LogicalParent Is Nothing Then
                ' keep moving up until we find the current folder
                Return Me.GetParentOfSelectionBefore(folder, selectedItem.LogicalParent)
            Else
                ' we're out of options
                Return Nothing
            End If
        End Function

        Public Overrides Sub OnApplyTemplate()
            MyBase.OnApplyTemplate()

            Me.PART_Grid = Me.Template.FindName("PART_Grid", Me)

            _views = New Dictionary(Of Guid, Control)()
        End Sub

        Public Sub DoRename(item As Item)
            Me.ActiveView.DoRename(item)
        End Sub

        Public Async Function DoRename(fullPath As String) As Task(Of Boolean)
            Return Await Me.ActiveView.DoRename(fullPath)
        End Function

        Public Property ActiveView As BaseFolderView
            Get
                Return _activeView
            End Get
            Set(value As BaseFolderView)
                _activeView = value
            End Set
        End Property

        Public Overridable Property SelectedItems As IEnumerable(Of Item)
            Get
                Return GetValue(SelectedItemsProperty)
            End Get
            Set(value As IEnumerable(Of Item))
                SetCurrentValue(SelectedItemsProperty, value)
            End Set
        End Property

        Public Property Folder As Folder
            Get
                Return GetValue(FolderProperty)
            End Get
            Set(ByVal value As Folder)
                SetCurrentValue(FolderProperty, value)
            End Set
        End Property

        Public Property IsSelecting As Boolean
            Get
                Return GetValue(IsSelectingProperty)
            End Get
            Set(ByVal value As Boolean)
                SetCurrentValue(IsSelectingProperty, value)
            End Set
        End Property

        Public Property StatusText As String
            Get
                Return GetValue(StatusTextProperty)
            End Get
            Set(ByVal value As String)
                SetCurrentValue(StatusTextProperty, value)
            End Set
        End Property

        Public Property SearchBox As SearchBox
            Get
                Return GetValue(SearchBoxProperty)
            End Get
            Set(ByVal value As SearchBox)
                SetCurrentValue(SearchBoxProperty, value)
            End Set
        End Property

        Public Property Navigation As Navigation
            Get
                Return GetValue(NavigationProperty)
            End Get
            Set(ByVal value As Navigation)
                SetCurrentValue(NavigationProperty, value)
            End Set
        End Property

        Private Async Sub updateStatusText()
            Await Task.Delay(100)
            If Not Me.Folder Is Nothing Then
                Dim text As String = String.Format("{0} {1}", Me.Folder.Items.Count, If(Me.Folder.Items.Count = 1, "item", "items"))
                If Not Me.SelectedItems Is Nothing AndAlso Not Me.SelectedItems.Count = 0 Then
                    text &= String.Format("       {0} {1} selected", Me.SelectedItems.Count, If(Me.SelectedItems.Count = 1, "item", "items"))
                    Dim items As IEnumerable(Of Item) = Me.SelectedItems.ToList()
                    Dim tcs As New TaskCompletionSource()
                    Shell.GlobalThreadPool.Add(
                        Sub()
                            Dim size As UInt64 = 0
                            For Each item In items
                                size += item.PropertiesByCanonicalName("System.Size")?.Value
                            Next
                            If size > 0 Then
                                Dim propertyDescription As IPropertyDescription = Nothing, pkey As PROPERTYKEY
                                Try
                                    Functions.PSGetPropertyDescriptionByName("System.Size", GetType(IPropertyDescription).GUID, propertyDescription)
                                    propertyDescription.GetPropertyKey(pkey)
                                Finally
                                    If Not propertyDescription Is Nothing Then
                                        Marshal.ReleaseComObject(propertyDescription)
                                        propertyDescription = Nothing
                                    End If
                                End Try
                                Dim buffer As StringBuilder = New StringBuilder()
                                buffer.Append(New String(" ", 2050))
                                Dim val As PROPVARIANT = New PROPVARIANT()
                                val.SetValue(size)
                                Functions.PSFormatForDisplay(pkey, val, PropertyDescriptionFormatOptions.None, buffer, 2048)
                                text &= "   " & buffer.ToString()
                            End If
                            Dim storageProviderUIStatus As List(Of String) =
                                items.Where(Function(i) Not i.disposedValue AndAlso Not i.IsReadyForDispose).Select(Function(i) _
                                    i.PropertiesByCanonicalName("System.StorageProviderUIStatus")?.Text).ToList()
                            If storageProviderUIStatus.Count = 1 Then
                                Dim status As String = storageProviderUIStatus.FirstOrDefault(Function(i) Not String.IsNullOrWhiteSpace(i))
                                If Not status Is Nothing Then text &= "       " & status
                            End If
                            tcs.SetResult()
                        End Sub)
                    Await tcs.Task
                    Me.StatusText = text
                Else
                    Me.StatusText = text
                End If
            Else
                Me.StatusText = String.Empty
            End If
        End Sub

        Shared Sub OnFolderChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
            Dim fv As FolderView = d
            fv.SelectedItems = Nothing
            If Not e.OldValue Is Nothing Then
                RemoveHandler CType(e.OldValue, Folder).PropertyChanged, AddressOf fv.folder_PropertyChanged
                RemoveHandler CType(e.OldValue, Folder).Items.CollectionChanged, AddressOf fv.folder_Items_CollectionChanged
            End If
            If Not e.NewValue Is Nothing Then
                CType(e.NewValue, Folder).IsActiveInFolderView = True
                Dim folderViewState As FolderViewState = FolderViewState.FromFolder(CType(e.NewValue, Folder))
                If Not CType(e.NewValue, Folder).ActiveView.HasValue Then
                    CType(e.NewValue, Folder).ActiveView = folderViewState.ActiveView
                End If
                fv.changeView(CType(e.NewValue, Folder).ActiveView, e.NewValue)
                AddHandler CType(e.NewValue, Folder).PropertyChanged, AddressOf fv.folder_PropertyChanged
                AddHandler CType(e.NewValue, Folder).Items.CollectionChanged, AddressOf fv.folder_Items_CollectionChanged
            End If
            If Not e.OldValue Is Nothing Then
                CType(e.OldValue, Folder).IsActiveInFolderView = False
            End If
            fv.updateStatusText()
        End Sub

        Private Sub folder_Items_CollectionChanged(sender As Object, e As NotifyCollectionChangedEventArgs)
            updateStatusText()
        End Sub

        Private Sub folder_PropertyChanged(sender As Object, e As PropertyChangedEventArgs)
            Select Case e.PropertyName
                Case "ActiveView"
                    changeView(Me.Folder.ActiveView, Me.Folder)
            End Select
        End Sub

        Private Sub changeView(newValue As Guid, folder As Folder)
            If Not _views Is Nothing Then
                Dim selectedItems As IEnumerable(Of Item) = If(Not Me.SelectedItems Is Nothing, Me.SelectedItems.ToList(), Nothing)
                Dim hasFocus As Boolean = Me.IsKeyboardFocusWithin
                If Not Me.ActiveView Is Nothing Then
                    Me.ActiveView.Folder = Nothing
                    BindingOperations.ClearBinding(Me.ActiveView, BaseFolderView.SelectedItemsProperty)
                    BindingOperations.ClearBinding(Me.ActiveView, BaseFolderView.IsSelectingProperty)
                End If
                For Each v In _views.Values
                    v.SetValue(Panel.ZIndexProperty, 0)
                Next
                If Not _views.ContainsKey(newValue) Then
                    Me.ActiveView = Activator.CreateInstance(folder.Views.First(Function(v) v.Guid = newValue).Type)
                    Me.ActiveView.Host = Me
                    Me.ActiveView.SearchBox = Me.SearchBox
                    Me.ActiveView.Navigation = Me.Navigation
                    Me.ActiveView.Visibility = Visibility.Hidden
                    _views.Add(newValue, Me.ActiveView)
                    If Not Me.PART_Grid Is Nothing Then
                        Me.PART_Grid.Children.Add(Me.ActiveView)
                    End If
                End If
                Me.ActiveView = _views(newValue)
                Me.ActiveView.Visibility = Visibility.Visible
                Me.ActiveView.SetValue(Panel.ZIndexProperty, 1)
                If hasFocus Then Me.ActiveView.Focus()
                Me.ActiveView.Folder = folder
                Dim folderViewState As FolderViewState = FolderViewState.FromViewName(folder.FullPath)
                folderViewState.ActiveView = newValue
                folderViewState.Persist()
                BindingOperations.SetBinding(Me.ActiveView, BaseFolderView.SelectedItemsProperty, New Binding("SelectedItems") With {.Source = Me})
                BindingOperations.SetBinding(Me.ActiveView, BaseFolderView.IsSelectingProperty, New Binding("IsSelecting") With {.Source = Me})
                Me.SelectedItems = selectedItems
                Me.RaiseEvent(New RoutedEventArgs(ViewActivatedEvent))
            End If
        End Sub

        Shared Sub OnSelectedItemsChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
            Dim fv As FolderView = d
            fv.updateStatusText()
        End Sub

        Protected Overridable Sub Dispose(disposing As Boolean)
            If Not disposedValue Then
                If disposing Then
                    ' dispose managed state (managed objects)
                    Shell.UnsubscribeFromNotifications(Me)

                    If Not Me.Folder Is Nothing Then
                        Me.Folder.IsActiveInFolderView = False
                    End If

                    WpfDragTargetProxy.RevokeDragDrop(Me)

                    Shell.RemoveFromControlCache(Me)
                End If

                ' free unmanaged resources (unmanaged objects) and override finalizer
                ' set large fields to null
                disposedValue = True
            End If
        End Sub

        Public Sub Dispose() Implements IDisposable.Dispose
            ' Do not change this code. Put cleanup code in 'Dispose(disposing As Boolean)' method
            Dispose(disposing:=True)
            GC.SuppressFinalize(Me)
        End Sub
    End Class
End Namespace