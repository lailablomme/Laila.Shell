Public Class PinnedAndFrequentPlaceholderFolder
    Inherits DummyFolder

    Public Sub New()
        MyBase.New("-", Nothing, Nothing)

        _fullPath = "-placeholder-"
    End Sub
End Class
