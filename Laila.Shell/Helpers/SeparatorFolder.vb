Public Class SeparatorFolder
    Inherits DummyFolder

    Public Sub New()
        MyBase.New("-", Nothing, Nothing)

        _fullPath = "-separator-"
    End Sub
End Class
