Imports System.Windows.Input

Namespace ViewModels
    Public Class RelayCommand
        Implements ICommand

        Public Event CanExecuteChanged As EventHandler Implements ICommand.CanExecuteChanged

        Private _execute As Action(Of Object)
        Private _canExecute As Predicate(Of Object)

        Public Sub New(execute As Action(Of Object))
            Me.New(execute, Nothing)
        End Sub

        Public Sub New(execute As Action(Of Object), canExecute As Predicate(Of Object))
            _execute = execute
            _canExecute = canExecute
        End Sub

        Public Sub Execute(parameter As Object) Implements ICommand.Execute
            _execute(parameter)
        End Sub

        Public Function CanExecute(parameter As Object) As Boolean Implements ICommand.CanExecute
            Return If(_canExecute Is Nothing, True, _canExecute(parameter))
        End Function
    End Class
End Namespace