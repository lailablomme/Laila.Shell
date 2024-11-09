Imports System.Collections.ObjectModel
Imports System.Windows

Namespace DebugTools
    Public Class DebugWindow
        Private _items As ObservableCollection(Of Item) = Shell.ItemsCache

        Public Sub New()
            ' This call is required by the designer.
            InitializeComponent()

            ' Add any initialization after the InitializeComponent() call.
            Me.DataContext = Me
        End Sub

        Public ReadOnly Property Items As ObservableCollection(Of Item)
            Get
                Return _items
            End Get
        End Property
    End Class
End Namespace