Imports System.Collections.Specialized
Imports System.ComponentModel
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Data
Imports System.Windows.Media
Imports Laila.Shell.Behaviors
Imports Laila.Shell.Events
Imports Laila.Shell.Helpers

Namespace Controls
    Public Class FolderView
        Inherits Control
        Implements IDisposable

        Public Shared ReadOnly FolderProperty As DependencyProperty = DependencyProperty.Register("Folder", GetType(Folder), GetType(FolderView), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, AddressOf OnFolderChanged))
        Public Shared ReadOnly SelectedItemsProperty As DependencyProperty = DependencyProperty.Register("SelectedItems", GetType(IEnumerable(Of Item)), GetType(FolderView), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, AddressOf OnSelectedItemsChanged))
        Public Shared ReadOnly IsSelectingProperty As DependencyProperty = DependencyProperty.Register("IsSelecting", GetType(Boolean), GetType(FolderView), New FrameworkPropertyMetadata(False, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly StatusTextProperty As DependencyProperty = DependencyProperty.Register("StatusText", GetType(String), GetType(FolderView), New FrameworkPropertyMetadata(String.Empty, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))

        Shared Sub New()
            DefaultStyleKeyProperty.OverrideMetadata(GetType(FolderView), New FrameworkPropertyMetadata(GetType(FolderView)))
        End Sub

        Private _views As Dictionary(Of String, Control) = New Dictionary(Of String, Control)()
        Private _activeView As BaseFolderView
        Private _dropTarget As IDropTarget
        Private PART_Grid As Grid
        Private _isLoaded As Boolean
        Private disposedValue As Boolean

        Public Sub New()
            AddHandler Shell.ShuttingDown,
                Sub(s As Object, e As EventArgs)
                    Me.Dispose()
                End Sub

            AddHandler Me.Loaded,
                Sub(s As Object, e As EventArgs)
                    If Not _isLoaded Then
                        _isLoaded = True

                        _dropTarget = New ListViewDropTarget(Me)
                        WpfDragTargetProxy.RegisterDragDrop(Me, _dropTarget)
                    End If
                End Sub
        End Sub

        Public Overrides Sub OnApplyTemplate()
            MyBase.OnApplyTemplate()

            Me.PART_Grid = Me.Template.FindName("PART_Grid", Me)

            If Not Me.ActiveView Is Nothing Then
                Me.PART_Grid.Children.Add(Me.ActiveView)
            End If
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

        Private Async Sub updateStatusText()
            Await Task.Delay(100)
            If Not Me.Folder Is Nothing Then
                Dim text As String = String.Format("{0} {1}", Me.Folder.Items.Count, If(Me.Folder.Items.Count = 1, "item", "items"))
                If Not Me.SelectedItems Is Nothing AndAlso Not Me.SelectedItems.Count = 0 Then
                    text &= String.Format("       {0} {1} selected", Me.SelectedItems.Count, If(Me.SelectedItems.Count = 1, "item", "items"))
                    Dim items As IEnumerable(Of Item) = Me.SelectedItems.ToList()
                    Dim tcs As New TaskCompletionSource()
                    Shell.SlowTaskQueue.Add(
                        Sub()
                            Dim size As UInt64 = 0
                            For Each item In items
                                size += item.PropertiesByCanonicalName("System.Size")?.Value
                            Next
                            If size > 0 Then
                                Dim propertyDescription As IPropertyDescription, pkey As PROPERTYKEY
                                Try
                                    Functions.PSGetPropertyDescriptionByName("System.Size", GetType(IPropertyDescription).GUID, propertyDescription)
                                    propertyDescription.GetPropertyKey(pkey)
                                Finally
                                    If Not propertyDescription Is Nothing Then
                                        Marshal.ReleaseComObject(propertyDescription)
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
                                    i.PropertiesByCanonicalName("System.StorageProviderUIStatus").Text).ToList()
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

        Shared Async Sub OnFolderChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
            Dim fv As FolderView = d
            fv.SelectedItems = Nothing
            If Not e.OldValue Is Nothing Then
                RemoveHandler CType(e.OldValue, Folder).PropertyChanged, AddressOf fv.folder_PropertyChanged
                RemoveHandler CType(e.OldValue, Folder).Items.CollectionChanged, AddressOf fv.folder_Items_CollectionChanged
            End If
            If Not e.NewValue Is Nothing Then
                CType(e.NewValue, Folder).IsActiveInFolderView = True
                Dim folderViewState As FolderViewState = FolderViewState.FromViewName(CType(e.NewValue, Folder).FullPath)
                If String.IsNullOrWhiteSpace(CType(e.NewValue, Folder).View) Then
                    CType(e.NewValue, Folder).View = folderViewState.View
                End If
                fv.changeView(CType(e.NewValue, Folder).View, e.NewValue)
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
                Case "View"
                    changeView(Me.Folder.View, Me.Folder)
            End Select
        End Sub

        Private Sub changeView(newValue As String, folder As Folder)
            Dim selectedItems As IEnumerable(Of Item) = If(Not Me.SelectedItems Is Nothing, Me.SelectedItems.ToList(), Nothing)
            Dim hasFocus As Boolean
            If Not Me.ActiveView Is Nothing Then
                hasFocus = Me.ActiveView.IsKeyboardFocusWithin
                Me.ActiveView.Folder = Nothing
                BindingOperations.ClearBinding(Me.ActiveView, BaseFolderView.SelectedItemsProperty)
                BindingOperations.ClearBinding(Me.ActiveView, BaseFolderView.IsSelectingProperty)
            End If
            For Each v In _views.Values
                v.SetValue(Panel.ZIndexProperty, 0)
            Next
            If Not _views.ContainsKey(newValue) Then
                Me.ActiveView = Activator.CreateInstance(Shell.FolderViews(newValue).Item2)
                Me.ActiveView.Host = Me
                _views.Add(newValue, Me.ActiveView)
                If Not Me.PART_Grid Is Nothing Then
                    Me.PART_Grid.Children.Add(Me.ActiveView)
                End If
            Else
                Me.ActiveView = _views(newValue)
            End If
            Me.ActiveView.SetValue(Panel.ZIndexProperty, 1)
            If hasFocus Then Me.ActiveView.Focus()
            Me.ActiveView.Folder = folder
            Dim folderViewState As FolderViewState = FolderViewState.FromViewName(folder.FullPath)
            folderViewState.View = newValue
            folderViewState.Persist()
            BindingOperations.SetBinding(Me.ActiveView, BaseFolderView.SelectedItemsProperty, New Binding("SelectedItems") With {.Source = Me})
            BindingOperations.SetBinding(Me.ActiveView, BaseFolderView.IsSelectingProperty, New Binding("IsSelecting") With {.Source = Me})
            Me.SelectedItems = selectedItems
        End Sub

        Shared Async Sub OnSelectedItemsChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
            Dim fv As FolderView = d
            fv.updateStatusText()
        End Sub

        Protected Overridable Sub Dispose(disposing As Boolean)
            If Not disposedValue Then
                If disposing Then
                    ' dispose managed state (managed objects)
                    If Not Me.Folder Is Nothing Then
                        Me.Folder.IsActiveInFolderView = False
                    End If

                    WpfDragTargetProxy.RevokeDragDrop(Me)
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