Imports Laila.Shell.Helpers

Namespace Controls.Parts
    Public Class PinnedItemsPlaceholderFolder
        Inherits DummyFolder

        Public Sub New()
            MyBase.New("-", Nothing, Nothing)

            _fullPath = "-placeholder-"
        End Sub
    End Class
End Namespace