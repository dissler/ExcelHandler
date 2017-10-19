Imports System.Collections.ObjectModel
Imports System.ComponentModel
Imports System.Data
Imports Microsoft.Win32
Imports Excel = Microsoft.Office.Interop.Excel

Public Class MainViewModel
    Implements INotifyPropertyChanged

    #Region "Fields"

    ' DataSet to hold all worksheets from excel file.
    Private _sheetSet as New DataSet

    #End Region ' Fields

    #Region "Properties"

    ' DataView that the DataGrid uses to display data.
    Public ReadOnly Property GridView As DataView
        Get
            If Me._sheetSet.Tables.Count > 0
                Return Me._sheetSet.Tables(SelectedTableIndex).DefaultView
            End If
            Return New DataView()
        End Get
    End Property

    ' Keeps track of which table in the combo box is selected.
    Private _selectedTableIndex as Integer = 0

    Public Property SelectedTableIndex As Integer
        Get
            Return _selectedTableIndex
        End Get
        Set(value As Integer)
            Me._selectedTableIndex = value
            NotifyPropertyChanged("SelectedTableIndex")

            ' ...and sets the DataView to be of the selected DataTable
            NotifyPropertyChanged("GridView")
        End Set
    End Property

    ' Provides a list of tables in the DataSet for the ComboBox to display.
    ' Table name comes from the worksheet name.
    Private _tableList As New ObservableCollection(Of String)

    Public Property TableList As ObservableCollection(Of String)
        Get
            Return Me._tableList
        End Get
        Set(value As ObservableCollection(Of String))
            If Me._tableList.Equals(value) Then
                Return
            End If
            Me._tableList = value
            NotifyPropertyChanged("TableList")
        End Set
    End Property

    ' Keeps track of whether there is a file loaded; the ComboBox
    ' and DataGrid will be hidden when no file is loaded.
    Private _fileLoaded as Boolean

    Public Property FileLoaded As Boolean
        Get
            Return Me._fileLoaded
        End Get
        Set(value As Boolean)
            If Me._fileLoaded.Equals(value) Then
                Return
            End If
            Me._fileLoaded = value
            NotifyPropertyChanged("FileLoaded")
        End Set
    End Property

    #End Region ' Properties

    #Region "Methods"

    ' This is run when the Button is clicked
    Private Sub LoadExcel()

        ' Prompt user for file
        Dim openFile = New OpenFileDialog
        openFile.Title = "Select an Excel File"
        openFile.Filter = "Excel Files|*.xls;*.xlsx|All Files|*.*"
        If openFile.ShowDialog() <> True
            Return
        End If

        Try
            ' Open instance of excel, create new DataSet, and read file.
            Dim xl As New Excel.Application
            xl.DisplayAlerts = false
            xl.ScreenUpdating = false
            Dim xlBooks As Excel.Workbooks = xl.Workbooks
            Dim thisFile As Excel.Workbook = xlBooks.Open(openFile.FileName)
            Dim returnSet As New DataSet
            Dim newTableList As New List(Of String)

            ' Read each sheet in the file
            For s As Integer = 1 To thisFile.Sheets.Count

                ' Make a new DataTable to hold the values from the sheet.
                Dim returnTable As New DataTable
                Dim thisSheet As Excel.Worksheet = thisFile.Sheets(s)
                returnTable.TableName = thisSheet.Name
                Dim thisRange as Excel.Range = thisSheet.UsedRange

                ' Create columns in the new DataTable for each column in the sheet's used range.
                For c As Integer = 1 To thisRange.Columns.Count
                    Dim newCol As New DataColumn
                    newCol.ColumnName = String.Format("Column{0}", c)
                    returnTable.Columns.Add(newCol)
                Next

                ' Read each row in the sheet, import values to the DataTable.
                For r As Integer = 1 to thisRange.Rows.Count
                    Dim newRow As DataRow = returnTable.NewRow()

                    ' Read each column in the excel row, add values to the new DataRow
                    For c As Integer = 1 To thisRange.Columns.Count
                        Dim thisCell = thisRange.Cells(r, c).Value
                        If thisCell IsNot Nothing
                            newRow(c - 1) = thisRange.Cells(r, c).Value.ToString()
                        End If
                    Next

                    ' Add the new DataRow to the DataTable and output progress to the Console.
                    returnTable.Rows.Add(newRow)
                    Console.WriteLine(String.Format("Read {0}/{1} row(s) from sheet {2}/{3}.", r, thisRange.Rows.Count, s, thisFile.Sheets.Count))
                Next

                ' Add the new DataTable to the DataSet, and the new
                ' table name to the list of table names.
                returnSet.Tables.Add(returnTable)
                newTableList.Add(returnTable.TableName)
            Next

            ' All done, let's close excel.
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
            Console.WriteLine("Done!")

        Catch ex As Exception
            Me._fileLoaded = False
            MessageBox.Show(String.Format("{0}{1}{2}", ex.Message, vbCrLf, ex.StackTrace), "Error Reading File")
        End Try

    End Sub

    #End Region ' Methods

    #Region "UI Implementation Stuff"

    Public Property LoadExcelCommand as ICommand = New DelegateCommand(AddressOf LoadExcel, AddressOf CanLoadExcel)

    Private Function CanLoadExcel(ByVal param as Object)
        Return True
    End Function

    Private Sub NotifyPropertyChanged(ByVal propertyName As String)
        RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(propertyName))
    End Sub

    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged
    
    #End Region ' UI Implementation Stuff

End Class
