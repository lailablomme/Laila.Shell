Imports System.Runtime.InteropServices
Imports System.Windows
Imports Laila.Shell.ViewModels

Namespace Controls
    Public Class DetailsListView
        Public Shared ReadOnly FolderProperty As DependencyProperty = DependencyProperty.Register("Folder", GetType(Folder), GetType(DetailsListView), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, AddressOf OnFolderChanged))

        Private _model As DetailsListViewModel

        Public Sub New()
            ' This call is required by the designer.
            InitializeComponent()

            ' Add any initialization after the InitializeComponent() call.
            _model = New DetailsListViewModel(Me)
            Me.DataContext = _model
        End Sub

        Public Property Folder As Folder
            Get
                Return GetValue(FolderProperty)
            End Get
            Set(ByVal value As Folder)
                SetCurrentValue(FolderProperty, value)
            End Set
        End Property

        Shared Sub OnFolderChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
            Dim dlv As DetailsListView = TryCast(d, DetailsListView)
            dlv.Model.Folder = e.NewValue
        End Sub

        Public ReadOnly Property Model As DetailsListViewModel
            Get
                Return _model
            End Get
        End Property
    End Class
End Namespace