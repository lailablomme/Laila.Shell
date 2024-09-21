Imports System.ComponentModel
Imports System.Runtime.CompilerServices
Imports System.Windows

Public Class NotifyPropertyChangedBase
    Implements INotifyPropertyChanged

    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged

    Public Sub NotifyOfPropertyChange(propertyName As String)
        Application.Current.Dispatcher.Invoke(
            Sub()
                RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(propertyName))
            End Sub)
    End Sub

    Protected Sub SetValue(Of T)(ByRef member As T, value As T, <CallerMemberName> Optional propertyName As String = "")
        member = value
        NotifyOfPropertyChange(propertyName)
    End Sub
End Class
