Imports System.Collections.ObjectModel
Imports System.ComponentModel
Imports System.Data
Imports Microsoft.Office.Interop.Excel
Imports Microsoft.Win32

Public Class MainViewModel
    Implements INotifyPropertyChanged

    #Region "Fields"

    ' DataSet to hold all worksheets from excel file
    Private _sheetSet as New DataSet

    ' Backing fields for bound properties
    Private _fileLoaded as Boolean
    Private _gridView As New DataView
    Private _selectedTableIndex as Integer
    Private _tableList As New ObservableCollection(Of String)

    #End Region

    #Region "Properties"

    ' Keeps track of whether there is a file loaded, the ComboBox
    ' and DataGrid will be hidden when no file is loaded
    Public Property FileLoaded As Boolean
        Get
            Return Me._fileLoaded
        End Get
        Set(value As Boolean)
            If Me._fileLoaded.Equals(value) Then
                Return
            End If
            Me._fileLoaded = value
            NotifyPropertyChanged(NameOf(Me.FileLoaded))
        End Set
    End Property

     ' DataView that the DataGrid uses to display data
    Public Property GridView as DataView
        Get
            Return Me._gridView
        End Get
        Set(value as DataView)
            If Me._gridView.Equals(value) Then
                Return
            End If
            Me._gridView = value
            NotifyPropertyChanged(NameOf(Me.GridView))
        End Set
    End Property

    ' Keeps track of which table in the combo box is selected...
    Public Property SelectedTableIndex As Integer
        Get
            Return _selectedTableIndex
        End Get
        Set(value As Integer)
            If Me._selectedTableIndex = value Then
                Return
            End If
            Me._selectedTableIndex = value
            NotifyPropertyChanged(NameOf(Me.SelectedTableIndex))

            ' ...and sets the DataView to be of the selected DataTable index,
            ' provided it is within the range of the DataSet's Tables collection
            If(Me.SelectedTableIndex >=0 AndAlso Me.SelectedTableIndex < Me._sheetSet.Tables.Count)
                Me.GridView = Me._sheetSet.Tables(Me.SelectedTableIndex).DefaultView
            End If
        End Set
    End Property

    ' Provides a list of tables in the DataSet for the ComboBox
    ' to display, table name comes from the worksheet name
    Public Property TableList As ObservableCollection(Of String)
        Get
            Return Me._tableList
        End Get
        Set(value As ObservableCollection(Of String))
            If Me._tableList.Equals(value) Then
                Return
            End If
            Me._tableList = value
            NotifyPropertyChanged(NameOf(Me.TableList))
        End Set
    End Property

    ' Delegate that a WPF element can be mapped to,
    ' invokes the LoadExcelCommand method
    Public Property LoadExcelCommand as ICommand = New DelegateCommand(AddressOf LoadExcel, AddressOf CanLoadExcel)

    #End Region
    
    #Region "Methods"

    ' Invoked by the LoadExcelCommand delegate
    Private Sub LoadExcel()

        ' Prompt user for file
        Dim openFile = New OpenFileDialog
        openFile.Title = "Select an Excel File"
        openFile.Filter = "Excel Files|*.xls;*.xlsx|All Files|*.*"
        If openFile.ShowDialog() <> True
            Return
        End If

        Try
            ' Create new instance of excel, create new DataSet, and read file
            Dim xl As New Microsoft.Office.Interop.Excel.Application
            Dim xlBooks As Workbooks = xl.Workbooks
            Dim thisFile As Workbook = xlBooks.Open(openFile.FileName)
            Dim returnSet As New DataSet
            Dim newTableList As New List(Of String)

            ' Read each sheet in the file
            For s As Integer = 1 To thisFile.Sheets.Count

                ' Make a new DataTable to hold the values from the sheet
                Dim returnTable As New System.Data.DataTable
                Dim thisSheet As Worksheet = thisFile.Sheets(s)
                returnTable.TableName = thisSheet.Name
                Dim thisRange as Range = thisSheet.UsedRange

                ' Create columns in the new DataTable for each column in the sheet's used range
                For c As Integer = 1 To thisRange.Columns.Count
                    Dim newCol As New DataColumn
                    newCol.ColumnName = String.Format("Column{0}", c)
                    returnTable.Columns.Add(newCol)
                Next

                ' Read each row in the sheet, import values to the DataTable
                For r As Integer = 1 to thisRange.Rows.Count
                    Dim newRow As DataRow = returnTable.NewRow()

                    ' Read each column in the excel row, add values to the new DataRow
                    For c As Integer = 1 To thisRange.Columns.Count
                        newRow(c - 1) = thisRange.Cells(r, c).Value.ToString()
                    Next

                    ' Add the new DataRow to the DataTable and output progress to the Console
                    returnTable.Rows.Add(newRow)
                    Console.WriteLine(String.Format("Read {0} row(s) from sheet {1}.", r - 1, s))
                Next

                ' Add the new DataTable to the DataSet, and the new
                ' table name to the list of table names
                returnSet.Tables.Add(returnTable)
                newTableList.Add(returnTable.TableName)
            Next

            ' All done, let's close excel
            thisFile.Close()
            xlBooks.Close()
            xl.Quit()

            ' Store the DataSet we've loaded, as well as
            ' the list of table names, and set FileLoaded to True
            Me._sheetSet = returnSet
            Me.TableList = New ObservableCollection(Of String)(newTableList)
            Me.FileLoaded = True

            ' Display the first DataTable
            Me.SelectedTableIndex = 0
            Me.GridView = Me._sheetSet.Tables(Me.SelectedTableIndex).DefaultView
            Console.WriteLine("Done!")

        Catch ex As Exception
            Me.FileLoaded = False
            MessageBox.Show(String.Format(ex.Message), "Error Reading File")
        End Try

    End Sub

    ' Checks to see if we can perform the requested action, in this case
    ' just returns True. We could insert a check here to see if the excel 
    ' interop libraries are available before trying to open excel
    Private Function CanLoadExcel(ByVal param as Object)
        Return True
    End Function

    #End Region

    #Region "Interface Implementation Stuff"

    ' Notifies the WPF visual elements that a bound property has changed
    Private Sub NotifyPropertyChanged(ByVal propertyName As String)
        RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(propertyName))
    End Sub

    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged
    
    #End Region

End Class

' Defines a delegate that invokes an action
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