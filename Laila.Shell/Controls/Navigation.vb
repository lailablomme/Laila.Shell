Imports System.Windows
Imports Laila.Shell.Helpers

Namespace Controls
    Public Class Navigation
        Inherits FrameworkElement

        Public Shared ReadOnly FolderProperty As DependencyProperty = DependencyProperty.Register("Folder", GetType(Folder), GetType(Navigation), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, AddressOf OnFolderChanged))
        Public Shared ReadOnly CanBackProperty As DependencyProperty = DependencyProperty.Register("CanBack", GetType(Boolean), GetType(Navigation), New FrameworkPropertyMetadata(False))
        Public Shared ReadOnly CanForwardProperty As DependencyProperty = DependencyProperty.Register("CanForward", GetType(Boolean), GetType(Navigation), New FrameworkPropertyMetadata(False))
        Public Shared ReadOnly CanUpProperty As DependencyProperty = DependencyProperty.Register("CanUp", GetType(Boolean), GetType(Navigation), New FrameworkPropertyMetadata(False, AddressOf OnCanUpChanged))
        Public Shared ReadOnly BackTextProperty As DependencyProperty = DependencyProperty.Register("BackText", GetType(String), GetType(Navigation), New FrameworkPropertyMetadata("Back"))
        Public Shared ReadOnly ForwardTextProperty As DependencyProperty = DependencyProperty.Register("ForwardText", GetType(String), GetType(Navigation), New FrameworkPropertyMetadata("Forward"))
        Public Shared ReadOnly UpTextProperty As DependencyProperty = DependencyProperty.Register("UpText", GetType(String), GetType(Navigation), New FrameworkPropertyMetadata("Up"))

        Private _isFolderChanging As Boolean
        Private _list As List(Of Folder)
        Private _pointer As Integer = -1

        Shared Sub New()
            DefaultStyleKeyProperty.OverrideMetadata(GetType(Navigation), New FrameworkPropertyMetadata(GetType(Navigation)))
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

        Public Sub Back()
            _isFolderChanging = True
            _pointer -= 1
            Me.Folder = _list(_pointer)
            _isFolderChanging = False
        End Sub

        Public Sub Forward()
            _isFolderChanging = True
            _pointer += 1
            Me.Folder = _list(_pointer)
            _isFolderChanging = False
        End Sub

        Public Sub Up()
            Me.Folder = Me.Folder.Parent
        End Sub

        Friend Sub OnFolderChangedInternal()
            If Not _isFolderChanging Then
                If _pointer = -1 Then
                    _list = New List(Of Folder)() From {Me.Folder}
                ElseIf _pointer = _list.Count - 1 Then
                    _list.Add(Me.Folder)
                Else
                    _list = _list.Take(_pointer + 1).ToList()
                    _list.Add(Me.Folder)
                End If
                _pointer = _list.Count - 1
            End If
            Me.CanBack = _pointer > 0
            Me.CanForward = _pointer < _list.Count - 1
            Me.CanUp = Not Me.Folder.Parent Is Nothing
            Me.BackText = If(_pointer - 1 >= 0, "Back to " & _list(_pointer - 1).DisplayName, "")
            Me.ForwardText = If(_pointer + 1 <= _list.Count - 1, "Forward to " & _list(_pointer + 1).DisplayName, "")
            Me.UpText = If(Not Me.Folder.Parent Is Nothing, "Up to " & Me.Folder.Parent.DisplayName, "")
        End Sub

        Shared Sub OnFolderChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
            Dim h As Navigation = TryCast(d, Navigation)
            h.OnFolderChangedInternal()
        End Sub

        Shared Sub OnCanUpChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
            Dim h As Navigation = TryCast(d, Navigation)
        End Sub
    End Class
End Namespace