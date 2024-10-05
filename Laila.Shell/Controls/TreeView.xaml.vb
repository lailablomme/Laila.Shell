Imports System.Windows
Imports Laila.Shell.ViewModels

Namespace Controls
    Public Class TreeView
        Public Shared ReadOnly FolderProperty As DependencyProperty = DependencyProperty.Register("Folder", GetType(Folder), GetType(TreeView), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, AddressOf OnFolderChanged))

        Private _model As TreeViewModel
        Friend Shared _isFolderChanging As Boolean

        Public Sub New()
            ' This call is required by the designer.
            InitializeComponent()

            ' Add any initialization after the InitializeComponent() call.
            _model = New TreeViewModel(Me)
            Me.DataContext = _model
        End Sub

        Public ReadOnly Property Model As TreeViewModel
            Get
                Return _model
            End Get
        End Property

        Public Property Folder As Folder
            Get
                Return GetValue(FolderProperty)
            End Get
            Set(ByVal value As Folder)
                Try
                    _isFolderChanging = True
                    SetCurrentValue(FolderProperty, value)
                Finally
                    _isFolderChanging = False
                End Try
            End Set
        End Property

        Shared Sub OnFolderChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
            Dim tv As TreeView = TryCast(d, TreeView)
            If Not tv._isFolderChanging Then
                tv.Model.SetSelectedFolder(e.NewValue)
            End If
        End Sub
    End Class
End Namespace