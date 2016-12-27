Public Class DelegateCommand
    Implements ICommand

    Private m_canExecute As Func(Of Object, Boolean)
    Private m_ExecuteAction as Action(Of Object)
    Private m_canExecuteCache as Boolean

    Public Sub New(ByVal executeAction As Action(Of Object), ByVal canExecute As Func(Of Object, Boolean))
        Me.m_ExecuteAction = executeAction
        Me.m_canExecute = canExecute
    End Sub

    Public Function CanExecute(parameter As Object) As Boolean Implements ICommand.CanExecute
        Dim temp As Boolean = m_canExecute(parameter)
        If m_canExecuteCache <> temp Then
            m_canExecuteCache = temp
            RaiseEvent CanExecuteChanged(Me, New EventArgs())
        End If
        Return m_canExecuteCache
    End Function

    Public Sub Execute(parameter As Object) Implements ICommand.Execute
        m_ExecuteAction(parameter)
    End Sub

    Public Event CanExecuteChanged As EventHandler Implements ICommand.CanExecuteChanged

End Class
