Imports System.ComponentModel
Imports System.Runtime.CompilerServices
Imports System.Windows

Public Class NotifyPropertyChangedBase
    Implements INotifyPropertyChanged

    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged

    Public Sub NotifyOfPropertyChange(propertyName As String)
        Dim appl As Application = Application.Current
        If Not appl Is Nothing Then
            appl.Dispatcher.Invoke(
                Sub()
                    RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(propertyName))
                End Sub)
        End If
    End Sub

    Protected Sub SetValue(Of T)(ByRef member As T, value As T, <CallerMemberName> Optional propertyName As String = "")
        member = value
        NotifyOfPropertyChange(propertyName)
    End Sub
End Class
