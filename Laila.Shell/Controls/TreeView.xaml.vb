Imports System.Windows
Imports Laila.Shell.ViewModels

Namespace Controls
    Public Class TreeView
        Public Shared ReadOnly SelectedFolderNameProperty As DependencyProperty = DependencyProperty.Register("SelectedFolderName", GetType(String), GetType(TreeView), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, AddressOf OnFolderNameChanged))

        Private _model As TreeViewModel

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

        Public Property SelectedFolderName As String
            Get
                Return GetValue(SelectedFolderNameProperty)
            End Get
            Set(ByVal value As String)
                SetCurrentValue(SelectedFolderNameProperty, value)
            End Set
        End Property

        Shared Sub OnFolderNameChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
            Dim tv As TreeView = TryCast(d, TreeView)
            tv.Model.SetSelectedFolder(e.NewValue)
        End Sub
    End Class
End Namespace