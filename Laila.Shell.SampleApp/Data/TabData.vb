Imports Laila.Shell.Helpers

Namespace Data
    Public Class TabData
        Inherits NotifyPropertyChangedBase

        Private _folder As Folder
        Private _selectedItems As IEnumerable(Of Item)

        Public Sub New()
            Me.Folder = Shell.GetSpecialFolder(SpecialFolders.Home).Clone()
        End Sub

        Public Property Folder As Folder
            Get
                Return _folder
            End Get
            Set(value As Folder)
                SetValue(_folder, value)
            End Set
        End Property

        Public Property SelectedItems As IEnumerable(Of Item)
            Get
                Return _selectedItems
            End Get
            Set(value As IEnumerable(Of Item))
                SetValue(_selectedItems, value)
            End Set
        End Property

        Public Overrides Function ToString() As String
            Return Me.Folder?.DisplayName
        End Function
    End Class
End Namespace