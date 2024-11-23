Imports System.ComponentModel
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Data
Imports System.Windows.Media
Imports Laila.Shell.Events
Imports Laila.Shell.Helpers

Namespace Controls
    Public Class FolderView
        Inherits Control
        Implements IDisposable

        Public Shared ReadOnly FolderProperty As DependencyProperty = DependencyProperty.Register("Folder", GetType(Folder), GetType(FolderView), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, AddressOf OnFolderChanged))
        Public Shared ReadOnly SelectedItemsProperty As DependencyProperty = DependencyProperty.Register("SelectedItems", GetType(IEnumerable(Of Item)), GetType(FolderView), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly MenusProperty As DependencyProperty = DependencyProperty.Register("Menus", GetType(Menus), GetType(FolderView), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, AddressOf OnMenusChanged))

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
            AddHandler Me.Loaded,
                Sub(s As Object, e As EventArgs)
                    If Not _isLoaded Then
                        _isLoaded = True

                        _dropTarget = New ListViewDropTarget(Me)
                        WpfDragTargetProxy.RegisterDragDrop(Me, _dropTarget)
                    End If
                End Sub

            AddHandler System.Windows.Application.Current.MainWindow.Closed,
                Sub()
                    WpfDragTargetProxy.RevokeDragDrop(Me)
                End Sub
        End Sub

        Public Overrides Sub OnApplyTemplate()
            MyBase.OnApplyTemplate()

            Me.PART_Grid = Me.Template.FindName("PART_Grid", Me)

            If Not Me.ActiveView Is Nothing Then
                Me.PART_Grid.Children.Add(Me.ActiveView)
            End If
        End Sub

        Public Sub DoRename()
            Me.ActiveView.DoRename()
        End Sub

        Public Async Function DoRename(pidl As Pidl) As Task(Of Boolean)
            Return Await Me.ActiveView.DoRename(pidl)
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

        Public Property Menus As Menus
            Get
                Return GetValue(MenusProperty)
            End Get
            Set(ByVal value As Menus)
                SetCurrentValue(MenusProperty, value)
            End Set
        End Property

        Shared Sub OnMenusChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
            Dim fv As FolderView = d
            If Not e.OldValue Is Nothing Then
                RemoveHandler CType(e.OldValue, Menus).RenameRequest, AddressOf fv.menus_RenameRequest
            End If
            If Not e.NewValue Is Nothing Then
                AddHandler CType(e.NewValue, Menus).RenameRequest, AddressOf fv.menus_RenameRequest
            End If
        End Sub

        Private Async Sub menus_RenameRequest(sender As Object, e As RenameRequestEventArgs)
            e.IsHandled = Await Me.DoRename(e.Pidl)
        End Sub

        Shared Async Sub OnFolderChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
            Dim fv As FolderView = d
            fv.SelectedItems = Nothing
            If Not e.OldValue Is Nothing Then
                RemoveHandler CType(e.OldValue, Folder).PropertyChanged, AddressOf fv.folder_PropertyChanged
            End If
            If Not e.NewValue Is Nothing Then
                CType(e.NewValue, Folder).IsActiveInFolderView = True
                Dim folderViewState As FolderViewState = FolderViewState.FromViewName(CType(e.NewValue, Folder).FullPath)
                fv.changeView(folderViewState.View)
                CType(e.NewValue, Folder).View = folderViewState.View
                AddHandler CType(e.NewValue, Folder).PropertyChanged, AddressOf fv.folder_PropertyChanged
            End If
            If Not e.OldValue Is Nothing Then
                CType(e.OldValue, Folder).IsActiveInFolderView = False
            End If
        End Sub

        Private Sub folder_PropertyChanged(sender As Object, e As PropertyChangedEventArgs)
            Select Case e.PropertyName
                Case "View"
                    changeView(Me.Folder.View)
            End Select
        End Sub

        Private Sub changeView(newValue As String)
            Dim selectedItems As IEnumerable(Of Item) = If(Not Me.SelectedItems Is Nothing, Me.SelectedItems.ToList(), Nothing)
            Dim hasFocus As Boolean
            If Not Me.ActiveView Is Nothing Then
                hasFocus = Me.ActiveView.IsKeyboardFocusWithin
                BindingOperations.ClearBinding(Me.ActiveView, BaseFolderView.FolderProperty)
                BindingOperations.ClearBinding(Me.ActiveView, BaseFolderView.SelectedItemsProperty)
                BindingOperations.ClearBinding(Me.ActiveView, BaseFolderView.MenusProperty)
            End If
            For Each v In _views.Values
                v.SetValue(Panel.ZIndexProperty, 0)
            Next
            If Not _views.ContainsKey(newValue) Then
                Me.ActiveView = Activator.CreateInstance(Shell.FolderViews(newValue).Item2)
                'Me.ActiveView.Host = me
                _views.Add(newValue, Me.ActiveView)
                If Not Me.PART_Grid Is Nothing Then
                    Me.PART_Grid.Children.Add(Me.ActiveView)
                End If
            Else
                Me.ActiveView = _views(newValue)
            End If
            Me.ActiveView.SetValue(Panel.ZIndexProperty, 1)
            If hasFocus Then Me.ActiveView.Focus()
            BindingOperations.SetBinding(Me.ActiveView, BaseFolderView.FolderProperty, New Binding("Folder") With {.Source = Me})
            BindingOperations.SetBinding(Me.ActiveView, BaseFolderView.SelectedItemsProperty, New Binding("SelectedItems") With {.Source = Me})
            BindingOperations.SetBinding(Me.ActiveView, BaseFolderView.MenusProperty, New Binding("Menus") With {.Source = Me})
            Me.SelectedItems = selectedItems
        End Sub

        Protected Overridable Sub Dispose(disposing As Boolean)
            If Not disposedValue Then
                If disposing Then
                    ' dispose managed state (managed objects)
                    If Not Me.Folder Is Nothing Then
                        Me.Folder.IsActiveInFolderView = False
                    End If
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