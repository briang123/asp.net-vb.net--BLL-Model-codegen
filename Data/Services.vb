Imports System.Data

Namespace Data

    Public Class Services

        Public Shared Sub MoveTableToNewDataSet(ByRef targetDataSet As DataSet, ByRef srcTable As DataTable)

            Try
                Dim dt As DataTable = srcTable.Clone
                dt.TableName = srcTable.TableName
                For Each row As DataRow In srcTable.Rows
                    dt.ImportRow(row)
                Next
                targetDataSet.Tables.Add(dt)
            Catch ex As Exception
                Throw
            End Try

        End Sub

        Public Shared Sub TruncateDataTable(ByRef sourceTable As DataTable)
            For i As Integer = sourceTable.Rows.Count - 1 To 0 Step -1
                sourceTable.Rows.RemoveAt(i)
            Next
        End Sub

        Public Shared Function RemoveDuplicateRows(ByVal sourceTable As DataTable, _
                                                    ByVal columnList As String, _
                                                    Optional ByVal sortExpr As String = "") As DataTable

            'Build my string array of column names in the data table
            Dim cols() As String = columnList.Split(","c)

            'Create array to store values in each field of cols() string array
            Dim vals() As String
            vals = DirectCast(cols.Clone, String())

            'Make a copy of our original data table to import distinct records
            Dim cloneTable As DataTable = sourceTable.Clone()

            'Sort our data table before iterating through data
            'If a sort expression is not supplied, then we will just
            'have our data table as it was passed into this function.
            Dim rows() As DataRow = sourceTable.Select("", sortExpr)

            'Initialize our vals() array with empty string values
            Dim i As Integer
            For i = cols.GetLowerBound(0) To cols.GetUpperBound(0)
                vals.SetValue(String.Empty, i)
            Next

            Dim result As Integer
            Dim different As Boolean = False
            'Loop through each data row
            Dim row As DataRow
            For Each row In rows

                'Loop through each column
                For i = cols.GetLowerBound(0) To cols.GetUpperBound(0)

                    'Check if each column value matches the previous rows column values
                    'A non-match will return -1, so we will need to check for it.
                    If vals(i).Equals(row(cols(i)).ToString) Then
                        result = 0
                    Else
                        result = -1
                    End If
                    'result = vals(i).Compare(vals(i), row(cols(i).ToString).ToString, True)

                    'Update our values array with the data in the current row
                    'This will be checked for the next time around
                    vals.SetValue(row(cols(i).ToString).ToString, i)

                    'If we are not on the first data row (empty string value) or 
                    'the current row does not match previous row then NO DUPE!!!
                    If vals.GetValue(i) Is String.Empty Or result < 0 Then
                        different = True
                    End If

                Next

                'If our NO DUPE flag is True then import the row into the 
                'new data table. We need to make sure we RESET our NO DUPE 
                'flag back to False; otherwise, all subsequent rows will 
                'be duped incorrectly
                If different = True Then
                    cloneTable.ImportRow(row)
                    different = False
                End If

            Next

            'Remember now that we imported the *FIRST* occurrence of each data row
            'in our data table, which was the reason for the sort expression. If 
            'you get incorrect results, you may want to check how you are sorting
            'the data in the data table.

            'Return our distinct values 
            Return cloneTable

        End Function

        Public Shared Function SafeSql(ByVal value As String) As String
            Return Chr(39) & value.Replace("'", "''") & Chr(39)
        End Function

        Public Shared Function GetString(ByVal value As Object) As String
            Try
                If value IsNot Nothing Then
                    If TypeOf value Is Integer Then
                        Return CType(value, Integer).ToString
                    Else
                        If value.ToString = "Null" Then
                            Return String.Empty
                        Else
                            Return value.ToString
                        End If
                    End If
                Else
                    Return String.Empty
                End If
            Catch aex As ApplicationException
                Throw New ApplicationException("An error occurred while trying to translate value into a String", aex)
            Catch ex As Exception
                Throw ex
            End Try
        End Function

        Public Shared Function GetInteger(ByVal value As Object) As Integer
            Try
                Dim l_value As Integer = 0
                If value IsNot Nothing Then
                    If TypeOf value Is Integer Then
                        l_value = value
                    Else
                        If value Is DBNull.Value Then
                            l_value = 0
                        ElseIf value.ToString = "" Then
                            l_value = 0
                        Else
                            l_value = CType(value, Integer)
                        End If
                    End If
                End If
                Return l_value
            Catch aex As ApplicationException
                Throw New ApplicationException("An error occurred while trying to translate value into an Integer", aex)
            Catch ex As Exception
                Throw ex
            End Try
        End Function

        Public Shared Function GetDate(ByVal value As Object) As Date
            Try
                Dim l_value As Date = Date.MinValue
                If value IsNot Nothing Then
                    If TypeOf value Is Date Then
                        l_value = value
                    ElseIf value Is DBNull.Value Then
                        l_value = Date.MinValue
                    Else
                        l_value = CType(value, Date)
                    End If
                End If
                Return l_value
            Catch aex As ApplicationException
                Throw New ApplicationException("An error occurred while trying to translate value into a Date", aex)
            Catch ex As Exception
                Throw ex
            End Try
        End Function

        Public Shared Function GetDouble(ByVal value As Object) As Double
            Try
                Dim l_value As Double = 0
                If value IsNot Nothing Then
                    If TypeOf value Is Double Then
                        l_value = value
                    Else
                        l_value = CType(value, Double)
                    End If
                End If
                Return l_value
            Catch aex As ApplicationException
                Throw New ApplicationException("An error occurred while trying to translate value into a Double", aex)
            Catch ex As Exception
                Throw ex
            End Try
        End Function

        Public Shared Function GetBoolean(ByVal value As Object) As Boolean
            Try
                Dim l_value As Boolean = False
                If value IsNot Nothing Then
                    If TypeOf value Is String Then
                        Dim s_value As String = String.Empty
                        s_value = value.ToString
                        Select Case s_value.Trim.ToUpper
                            Case "Y", "YES", "TRUE", "1"
                                l_value = True
                            Case "N", "NO", "FALSE", "0"
                                l_value = False
                        End Select
                    ElseIf TypeOf value Is Boolean Then
                        l_value = value
                    Else
                        l_value = CType(value, Boolean)
                    End If
                End If
                Return l_value
            Catch aex As ApplicationException
                Throw New ApplicationException("An error occurred while trying to translate value into a Boolean", aex)
            Catch ex As Exception
                Throw ex
            End Try
        End Function

        Public Shared Function GetNULLableString(ByVal val As Object) As String
            If val Is Nothing Then
                Return String.Empty
            ElseIf val.Equals(DBNull.Value) Or val.ToString = "&nbsp;" Then
                Return String.Empty
            Else
                Return val.ToString()
            End If
        End Function

        Public Shared Function GetNULLableInteger(ByVal val As Object) As Integer
            If val Is Nothing Then
                Return 0
            ElseIf val.Equals(DBNull.Value) Or val.Equals(String.Empty) Or val.ToString = "&nbsp;" Then
                Return 0
            Else
                Return Convert.ToInt32(val)
            End If
        End Function

        Public Shared Function GetNULLableBoolean(ByVal val As Object) As Boolean
            If val Is Nothing Then
                Return False
            ElseIf val.Equals(DBNull.Value) Or val.ToString = "&nbsp;" Then
                Return False
            Else
                Return Convert.ToBoolean(val)
            End If
        End Function

        Public Shared Function GetNULLableDouble(ByVal val As Object) As Double
            If val Is Nothing Then
                Return 0
            ElseIf val.Equals(DBNull.Value) Or val.Equals(String.Empty) Or val.ToString = "&nbsp;" Then
                Return 0
            Else
                Return Convert.ToDouble(val)
            End If
        End Function

        Public Shared Function GetNULLableDecimal(ByVal val As Object) As Decimal
            If val Is Nothing Then
                Return 0
            ElseIf val.Equals(DBNull.Value) Or val.Equals(String.Empty) Or val.ToString = "&nbsp;" Then
                Return 0
            Else
                Return Convert.ToDecimal(val)
            End If
        End Function

        Public Shared Function GetNULLableDateTime(ByVal val As Object) As DateTime
            If val Is Nothing Then
                Return Nothing
            ElseIf val.Equals(DBNull.Value) Or val.Equals(String.Empty) Or val.ToString = "&nbsp;" Then
                Return Nothing
            Else
                Return Convert.ToDateTime(val)
            End If
        End Function

        Public Shared Function GetNULLableStringParameter(ByVal iParm As IDbDataParameter, ByVal val As Object) As IDbDataParameter

            Dim localParm As IDbDataParameter = iParm
            If val.Equals(DBNull.Value) Or val.Equals(String.Empty) Then
                localParm.Value = DBNull.Value
            Else
                localParm.Value = Replace(val.ToString, "'", "''")
            End If

            Return localParm

        End Function

        Public Shared Function GetNULLableIntegerParameter(ByVal iParm As IDbDataParameter, ByVal val As Object) As IDbDataParameter

            Dim localParm As IDbDataParameter = iParm
            If val.Equals(DBNull.Value) Or val.Equals(String.Empty) Then
                localParm.Value = DBNull.Value
            Else
                localParm.Value = Convert.ToInt32(val)
            End If

            Return localParm

        End Function

        Public Shared Function GetNULLableDoubleParameter(ByVal iParm As IDbDataParameter, ByVal val As Object) As IDbDataParameter

            Dim localParm As IDbDataParameter = iParm
            If val.Equals(DBNull.Value) Or val.Equals(String.Empty) Then
                localParm.Value = DBNull.Value
            Else
                localParm.Value = Convert.ToDouble(val)
            End If

            Return localParm

        End Function

        Public Shared Function GetNULLableDecimalParameter(ByVal iParm As IDbDataParameter, ByVal val As Object) As IDbDataParameter

            Dim localParm As IDbDataParameter = iParm
            If val.Equals(DBNull.Value) Or val.Equals(String.Empty) Then
                localParm.Value = DBNull.Value
            Else
                localParm.Value = Convert.ToDecimal(val)
            End If

            Return localParm

        End Function

        Public Shared Function GetNULLableDateParameter(ByVal iParm As IDbDataParameter, ByVal val As Object) As IDbDataParameter

            Dim localParm As IDbDataParameter = iParm
            If val.Equals(DBNull.Value) Or val.Equals(String.Empty) Then
                localParm.Value = DBNull.Value
            Else
                localParm.Value = Convert.ToDateTime(val)
            End If

            Return localParm

        End Function

        Public Shared Function GetReaderNULLableValue(ByVal r As IDataReader, ByVal columnIndex As Integer, ByVal type As DbType, ByVal retText As Object) As Object
            Dim value As Object = Nothing
            If r.IsDBNull(columnIndex) = True Then
                Select Case type
                    Case DbType.Currency, DbType.Decimal, DbType.Double, DbType.Int32, DbType.Int16, DbType.Int64, DbType.Single
                        value = 0
                    Case DbType.String
                        value = String.Empty
                End Select
            Else
                value = r.GetValue(columnIndex)
            End If
            Return value
        End Function


        Public Shared Function NLTIsNumber(ByVal val As String) As Boolean
            Dim isNum As Boolean = True
            If val.Length > 0 Then
                For i As Integer = 1 To val.Length
                    If Char.IsNumber(val, i - 1) = False Then
                        isNum = False
                        Exit For
                    End If
                Next
            Else
                isNum = False
            End If
            Return isNum
        End Function

        Public Shared Function IsPastDate(ByVal val As String) As Boolean

            If IsDate(val) Then

                Dim mo As Integer = Month(CType(val, Date))
                Dim d As Integer = Day(CType(val, Date))
                Dim y As Integer = Year(CType(val, Date))

                Dim dateSelected As Date = New Date(y, mo, d)

                If Date.Today.Subtract(dateSelected).TotalDays > 0 Then
                    Return True
                Else
                    Return False
                End If
            Else
                Throw New InvalidCastException()
            End If

        End Function

        Public Shared Function BindToEnumKeyValue(ByVal names() As String, ByVal values As Array) As DataTable

            Dim dt As New DataTable
            dt.Columns.Add("Key", GetType(String))
            dt.Columns.Add("Value", GetType(Integer))

            Dim i As Integer = 0
            While i < names.Length
                Dim dr As DataRow = dt.NewRow
                dr("Key") = names(i)
                dr("Value") = CType(values.GetValue(i), Integer)
                dt.Rows.Add(dr)
                i += i
            End While
            Return dt

        End Function

        Public Shared Function AddReadOnlyLineBreaks(ByVal val As String) As String
            If val Is Nothing Then val = String.Empty
            Return RegularExpressions.Regex.Replace(val, vbCrLf, "<br/>")
        End Function

        Public Shared Function GetListFromConfig(ByVal key As String) As IList

            'NOTE:  If you see a blank item in whatever you're binding to, then check 
            '       the web.config key for a ";" at the end of the value list. No semi-colon at the end.
            Dim customList() As String = System.Configuration.ConfigurationManager.AppSettings(key).Split(";"c)
            Dim al As New ArrayList
            With al
                For i As Integer = 0 To customList.Length - 1
                    .Add(customList(i))
                Next
                .Sort()
            End With
            Return al
        End Function

        Public Shared Function GetHashFromConfig(ByVal key As String) As Hashtable

            'NOTE:  If you see a blank item in whatever you're binding to, then check 
            '       the web.config key for a ";" at the end of the value list. No semi-colon at the end.
            Dim customList() As String = System.Configuration.ConfigurationManager.AppSettings(key).Split(";"c)
            Dim ht As New Hashtable
            With ht
                For i As Integer = 0 To customList.Length - 1
                    Dim keyvalpair() As String = customList(i).Split("|"c)

                    Dim val1 As String = keyvalpair(0).ToString
                    Dim val2 As String = keyvalpair(1).ToString
                    .Add(val2, val1)
                Next
            End With
            Return ht
        End Function

        ''' <summary>
        ''' Convert a ASP.NET DataSet to a JSON object
        ''' </summary>
        ''' <param name="ds"></param>
        ''' <returns></returns>
        ''' <remarks>
        ''' Once you parse the JSON string to an object, you will be able to reference rows like this:
        ''' var jObj = $.parseJSON(msg.d);
        ''' jObj.TableName[rowindex].ColumnName 
        ''' NOTE: Make sure to pass the value through a format method of some sort depending on the data type;
        ''' otherwise, you may not see any data on the screen. (i.e. if null, then return '')
        '''</remarks>
        Public Shared Function DataSetToJSON(ByVal ds As DataSet) As String
            Dim dict As New Dictionary(Of String, Object)

            'loop through each datatable in the dataset
            For Each dt As DataTable In ds.Tables

                Dim rows As New List(Of Dictionary(Of String, Object))
                Dim row As Dictionary(Of String, Object)

                'loop through each row of datatable and add key/value pair
                For Each dr As DataRow In dt.Rows
                    row = New Dictionary(Of String, Object)
                    For Each col As DataColumn In dt.Columns
                        row.Add(col.ColumnName, dr(col))
                    Next

                    'add each row to the rows collection
                    rows.Add(row)
                Next

                'add each rows collection for each datatable to a dictionary object
                dict.Add(dt.TableName, rows)
            Next

            Dim json As New System.Web.Script.Serialization.JavaScriptSerializer()

            'serialize the dictionary object with all datatables so that it's returned as a json string
            Return json.Serialize(dict)
        End Function


        Public Shared Function GetJson(ByVal dt As DataTable) As String

            Dim serializer As System.Web.Script.Serialization.JavaScriptSerializer = New System.Web.Script.Serialization.JavaScriptSerializer()
            Dim rows As New List(Of Dictionary(Of String, Object))
            Dim row As Dictionary(Of String, Object)
            For Each dr As DataRow In dt.Rows
                row = New Dictionary(Of String, Object)
                For Each col As DataColumn In dt.Columns
                    row.Add(col.ColumnName, dr(col))
                Next
                rows.Add(row)
            Next
            Return serializer.Serialize(rows)

        End Function

        ''' <summary>
        ''' Giving the ability to customize our results with some additional settings about the data we are returning
        ''' </summary>
        ''' <param name="CurrentPage"></param>
        ''' <param name="RecordCount"></param>
        ''' <returns></returns>
        ''' <remarks>
        ''' Optional: We don't necessarily need to grab these settings, but this is how we can 
        ''' set it up and bring back data based on server side processing. 
        ''' </remarks>
        Public Shared Function GetPagingSettings(ByVal CurrentPage As Integer, _
                                                 ByVal RecordsPerPage As Integer, _
                                                 ByVal RecordCount As Integer) As DataTable

            'we are adding a new datatable to the dataset to hold table settings
            Dim dtSettings As New DataTable("Settings")
            With dtSettings.Columns
                .Add(New DataColumn("RecordCount", System.Type.GetType("System.String")))
                .Add(New DataColumn("PageCount", System.Type.GetType("System.String")))
                .Add(New DataColumn("RecordsPerPage", System.Type.GetType("System.String")))
                .Add(New DataColumn("CurrentPage", System.Type.GetType("System.String")))
                .Add(New DataColumn("StartRowIndex", System.Type.GetType("System.String")))
                .Add(New DataColumn("EndRowIndex", System.Type.GetType("System.String")))
            End With
            Dim dr As DataRow = dtSettings.NewRow

            If RecordCount > 0 Then

                Dim _recordsPerPage As Integer
                If CurrentPage > 0 Then
                    _recordsPerPage = RecordsPerPage 'CType(System.Configuration.ConfigurationManager.AppSettings("RecordsPerPage"), Integer)
                Else
                    _recordsPerPage = RecordCount
                End If

                dr("RecordCount") = RecordCount
                dr("RecordsPerPage") = _recordsPerPage

                Dim TotalPages As Integer = RecordCount / _recordsPerPage
                Dim partialPage As Integer = RecordCount Mod _recordsPerPage
                If partialPage > 0 Then
                    TotalPages += 1
                End If
                dr("PageCount") = TotalPages

                'make sure that current page is within the bounds of the pages we actually have.
                If CurrentPage > TotalPages OrElse CurrentPage = 0 Then
                    CurrentPage = 1
                End If
                dr("CurrentPage") = CurrentPage

                Dim startRowIndex As Integer
                If CurrentPage > 1 Then
                    startRowIndex = ((CurrentPage - 1) * _recordsPerPage)
                Else
                    startRowIndex = 0
                End If

                dr("StartRowIndex") = startRowIndex

                Dim endRowIndex As Integer
                If CurrentPage = TotalPages Then
                    If partialPage > 0 Then
                        endRowIndex = (startRowIndex + partialPage) - 1
                    Else
                        endRowIndex = (startRowIndex + _recordsPerPage) - 1
                    End If
                Else
                    endRowIndex = (startRowIndex + _recordsPerPage) - 1
                End If

                dr("EndRowIndex") = endRowIndex

            End If

            dtSettings.Rows.Add(dr)

            Return dtSettings

        End Function

    End Class

End Namespace