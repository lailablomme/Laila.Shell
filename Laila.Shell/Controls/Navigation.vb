﻿Imports System.Windows
Imports System.Windows.Input
Imports Laila.Shell.Helpers

Namespace Controls
    Public Class Navigation
        Inherits FrameworkElement
        Implements IDisposable

        Public Shared ReadOnly FolderProperty As DependencyProperty = DependencyProperty.Register("Folder", GetType(Folder), GetType(Navigation), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, AddressOf OnFolderChanged))
        Public Shared ReadOnly CanBackProperty As DependencyProperty = DependencyProperty.Register("CanBack", GetType(Boolean), GetType(Navigation), New FrameworkPropertyMetadata(False))
        Public Shared ReadOnly CanForwardProperty As DependencyProperty = DependencyProperty.Register("CanForward", GetType(Boolean), GetType(Navigation), New FrameworkPropertyMetadata(False))
        Public Shared ReadOnly CanUpProperty As DependencyProperty = DependencyProperty.Register("CanUp", GetType(Boolean), GetType(Navigation), New FrameworkPropertyMetadata(False, AddressOf OnCanUpChanged))
        Public Shared ReadOnly BackTextProperty As DependencyProperty = DependencyProperty.Register("BackText", GetType(String), GetType(Navigation), New FrameworkPropertyMetadata("Back"))
        Public Shared ReadOnly ForwardTextProperty As DependencyProperty = DependencyProperty.Register("ForwardText", GetType(String), GetType(Navigation), New FrameworkPropertyMetadata("Forward"))
        Public Shared ReadOnly UpTextProperty As DependencyProperty = DependencyProperty.Register("UpText", GetType(String), GetType(Navigation), New FrameworkPropertyMetadata("Up"))
        Public Shared ReadOnly RefreshTextProperty As DependencyProperty = DependencyProperty.Register("RefreshText", GetType(String), GetType(Navigation), New FrameworkPropertyMetadata("Refresh"))

        Private _isFolderChanging As Boolean
        Private _list As List(Of Folder)
        Private _pointer As Integer = -1
        Private disposedValue As Boolean
        Private _isLoaded As Boolean

        Shared Sub New()
            DefaultStyleKeyProperty.OverrideMetadata(GetType(Navigation), New FrameworkPropertyMetadata(GetType(Navigation)))
        End Sub

        Public Sub New()
            AddHandler Me.Loaded,
                Sub(s As Object, e As EventArgs)
                    If Not _isLoaded Then
                        _isLoaded = True
                        AddHandler Window.GetWindow(Me).Closed,
                            Sub(s2 As Object, e2 As EventArgs)
                                Me.Dispose()
                            End Sub
                    End If
                End Sub
        End Sub

        Public Property Folder As Folder
            Get
                Return GetValue(FolderProperty)
            End Get
            Set(ByVal value As Folder)
                SetCurrentValue(FolderProperty, value)
            End Set
        End Property

        Public Property CanBack As Boolean
            Get
                Return GetValue(CanBackProperty)
            End Get
            Set(ByVal value As Boolean)
                SetCurrentValue(CanBackProperty, value)
            End Set
        End Property

        Public Property CanForward As Boolean
            Get
                Return GetValue(CanForwardProperty)
            End Get
            Set(ByVal value As Boolean)
                SetCurrentValue(CanForwardProperty, value)
            End Set
        End Property

        Public Property CanUp As Boolean
            Get
                Return GetValue(CanUpProperty)
            End Get
            Set(ByVal value As Boolean)
                SetCurrentValue(CanUpProperty, value)
            End Set
        End Property

        Public Property BackText As String
            Get
                Return GetValue(BackTextProperty)
            End Get
            Set(ByVal value As String)
                SetCurrentValue(BackTextProperty, value)
            End Set
        End Property

        Public Property ForwardText As String
            Get
                Return GetValue(ForwardTextProperty)
            End Get
            Set(ByVal value As String)
                SetCurrentValue(ForwardTextProperty, value)
            End Set
        End Property

        Public Property UpText As String
            Get
                Return GetValue(UpTextProperty)
            End Get
            Set(ByVal value As String)
                SetCurrentValue(UpTextProperty, value)
            End Set
        End Property

        Public Property RefreshText As String
            Get
                Return GetValue(RefreshTextProperty)
            End Get
            Set(ByVal value As String)
                SetCurrentValue(RefreshTextProperty, value)
            End Set
        End Property

        Public Sub Back()
            Using Shell.OverrideCursor(Cursors.Wait)
                _isFolderChanging = True
                _pointer -= 1
                Me.Folder = _list(_pointer)
                _isFolderChanging = False
            End Using
        End Sub

        Public Sub Forward()
            Using Shell.OverrideCursor(Cursors.Wait)
                _isFolderChanging = True
                _pointer += 1
                Me.Folder = _list(_pointer)
                _isFolderChanging = False
            End Using
        End Sub

        Public Sub Up()
            Using Shell.OverrideCursor(Cursors.Wait)
                Dim existing As Folder = _list.FirstOrDefault(Function(i) If(Me.Folder?.LogicalParent?.Pidl?.Equals(i.Pidl), False))
                If Me.CanBack AndAlso Not existing Is Nothing Then
                    Me.Folder = existing
                Else
                    Me.Folder.LogicalParent.LastScrollOffset = New Point()
                    Me.Folder.LogicalParent.IsInHistory = True
                    Me.Folder = Me.Folder.LogicalParent
                End If
            End Using
        End Sub

        Friend Sub OnFolderChangedInternal()
            If Not _isFolderChanging Then
                If _pointer = -1 Then
                    _list = New List(Of Folder)() From {Me.Folder}
                ElseIf _pointer = _list.Count - 1 Then
                    _list.Add(Me.Folder)
                Else
                    Dim newList As List(Of Folder) = _list.Take(_pointer + 1).ToList()
                    For Each item In _list.Skip(_pointer + 1)
                        If Not newList.Exists(Function(i) item.Pidl?.Equals(i.Pidl)) Then
                            item.IsInHistory = False
                        End If
                    Next
                    _list = newList
                    _list.Add(Me.Folder)
                End If
                Me.Folder.IsInHistory = True
                _pointer = _list.Count - 1
            End If
            Me.CanBack = _pointer > 0
            Me.CanForward = _pointer < _list.Count - 1
            Me.CanUp = Not Me.Folder.LogicalParent Is Nothing
            Me.BackText = If(_pointer - 1 >= 0, String.Format(My.Resources.Navigation_BackText, _list(_pointer - 1).DisplayName), "")
            Me.ForwardText = If(_pointer + 1 <= _list.Count - 1, String.Format(My.Resources.Navigation_ForwardText, _list(_pointer + 1).DisplayName), "")
            Me.UpText = If(Not Me.Folder.LogicalParent Is Nothing, String.Format(My.Resources.Navigation_UpText, Me.Folder.LogicalParent.DisplayName), "")
            Me.RefreshText = If(Not Me.Folder Is Nothing, String.Format(My.Resources.Navigation_RefreshText, Me.Folder.DisplayName), "")
        End Sub

        Shared Sub OnFolderChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
            Dim h As Navigation = TryCast(d, Navigation)
            If Not e.NewValue Is Nothing Then
                h.OnFolderChangedInternal()
            End If
        End Sub

        Shared Sub OnCanUpChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
            Dim h As Navigation = TryCast(d, Navigation)
        End Sub

        Protected Overridable Sub Dispose(disposing As Boolean)
            If Not disposedValue Then
                If disposing Then
                    ' dispose managed state (managed objects)
                    For Each f In _list
                        f.IsInHistory = False
                    Next
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