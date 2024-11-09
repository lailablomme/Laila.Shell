Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Data
Imports System.Windows.Media
Imports Laila.Shell.Helpers

Namespace Controls
    Public Class FolderView
        Inherits Control
        Implements IDisposable

        Public Shared ReadOnly FolderProperty As DependencyProperty = DependencyProperty.Register("Folder", GetType(Folder), GetType(FolderView), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, AddressOf OnFolderChanged))
        Public Shared ReadOnly ViewProperty As DependencyProperty = DependencyProperty.Register("View", GetType(String), GetType(FolderView), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, AddressOf OnViewChanged))
        Public Shared ReadOnly SelectedItemsProperty As DependencyProperty = DependencyProperty.Register("SelectedItems", GetType(IEnumerable(Of Item)), GetType(FolderView), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))

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

        Shared Sub OnFolderChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
            If Not e.NewValue Is Nothing Then
                CType(e.NewValue, Folder).IsActiveInFolderView = True
            End If
            If Not e.OldValue Is Nothing Then
                Dim view As CollectionView = CollectionViewSource.GetDefaultView(e.OldValue.Items)
                view.SortDescriptions.Clear()
                CType(e.OldValue, Folder).IsActiveInFolderView = False
            End If
        End Sub

        Public Property View As String
            Get
                Return GetValue(ViewProperty)
            End Get
            Set(ByVal value As String)
                SetCurrentValue(ViewProperty, value)
            End Set
        End Property

        Shared Sub OnViewChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
            Dim fv As FolderView = d
            If Not e.OldValue Is Nothing Then
                BindingOperations.ClearBinding(fv._views(e.OldValue), BaseFolderView.FolderProperty)
                BindingOperations.ClearBinding(fv._views(e.OldValue), BaseFolderView.SelectedItemsProperty)
            End If
            For Each v In fv._views.Values
                v.SetValue(Panel.ZIndexProperty, 0)
            Next
            If Not fv._views.ContainsKey(e.NewValue) Then
                fv.ActiveView = Activator.CreateInstance(Shell.FolderViews(e.NewValue))
                fv.ActiveView.Host = fv
                fv._views.Add(e.NewValue, fv.ActiveView)
                If Not fv.PART_Grid Is Nothing Then
                    fv.PART_Grid.Children.Add(fv.ActiveView)
                End If
            Else
                fv.ActiveView = fv._views(e.NewValue)
            End If
            fv.ActiveView.SetValue(Panel.ZIndexProperty, 1)
            fv.Folder.LastScrollOffset = New Point()
            BindingOperations.SetBinding(fv.ActiveView, BaseFolderView.FolderProperty, New Binding("Folder") With {.Source = fv})
            BindingOperations.SetBinding(fv.ActiveView, BaseFolderView.SelectedItemsProperty, New Binding("SelectedItems") With {.Source = fv})
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