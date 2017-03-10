Public Class DelegateCommand
    Implements ICommand

    Private _canExecute As Func(Of Object, Boolean)
    Private _executeAction as Action(Of Object)
    Private _canExecuteCache as Boolean

    Public Sub New(ByVal executeAction As Action(Of Object), ByVal canExecute As Func(Of Object, Boolean))
        Me._executeAction = executeAction
        Me._canExecute = canExecute
    End Sub

    Public Function CanExecute(parameter As Object) As Boolean Implements ICommand.CanExecute
        Dim temp As Boolean = _canExecute(parameter)
        If _canExecuteCache <> temp Then
            _canExecuteCache = temp
            RaiseEvent CanExecuteChanged(Me, New EventArgs())
        End If
        Return _canExecuteCache
    End Function

    Public Sub Execute(parameter As Object) Implements ICommand.Execute
        _executeAction(parameter)
    End Sub

    Public Event CanExecuteChanged As EventHandler Implements ICommand.CanExecuteChanged

End Class
