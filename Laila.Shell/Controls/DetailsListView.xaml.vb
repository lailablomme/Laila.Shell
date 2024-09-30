Imports System.Runtime.InteropServices
Imports System.Windows
Imports Laila.Shell.ViewModels

Namespace Controls
    Public Class DetailsListView
        Public Shared ReadOnly FolderNameProperty As DependencyProperty = DependencyProperty.Register("FolderName", GetType(String), GetType(DetailsListView), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, AddressOf OnFolderNameChanged))
        Public Shared ReadOnly LogicalParentProperty As DependencyProperty = DependencyProperty.Register("LogicalParent", GetType(Folder), GetType(DetailsListView), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))

        Private _model As DetailsListViewModel

        Public Sub New()
            ' This call is required by the designer.
            InitializeComponent()

            ' Add any initialization after the InitializeComponent() call.
            _model = New DetailsListViewModel(Me)
            Me.DataContext = _model
            _model.FolderName = Me.FolderName
        End Sub

        Public Property FolderName As String
            Get
                Return GetValue(FolderNameProperty)
            End Get
            Set(ByVal value As String)
                SetValue(FolderNameProperty, value)
            End Set
        End Property

        Shared Sub OnFolderNameChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
            Dim dlv As DetailsListView = TryCast(d, DetailsListView)
            dlv.Model.FolderName = e.NewValue
        End Sub

        Public Property LogicalParent As Folder
            Get
                Return GetValue(LogicalParentProperty)
            End Get
            Set(ByVal value As Folder)
                SetValue(LogicalParentProperty, value)
            End Set
        End Property

        Public ReadOnly Property Model As DetailsListViewModel
            Get
                Return _model
            End Get
        End Property
    End Class
End Namespace