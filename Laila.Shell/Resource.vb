Imports System.Windows.Input
Imports System.Windows.Markup

Public Class Resource
    Inherits MarkupExtension

    Public Property Key As String

    Public Sub New(key As String)
        Me.Key = key
    End Sub

    Public Overrides Function ProvideValue(serviceProvider As IServiceProvider) As Object
        Return My.Resources.ResourceManager.GetString(Key, My.Resources.Culture)
    End Function
End Class
