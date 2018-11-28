Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Imports System.Net

Namespace NewLeaf.Utility

    Public Class RunSql

        Private _logViewer As String
        Private _isPreview As Boolean
        Private _sqlfile As String
        Private _connectionString As String
        Private _timeout As Integer = 600
        Private _createProcs As Boolean = False
        Private _log As SortedList

        Private _database As String
        Private _server As String
        Private _userid As String
        Private _password As String

        Private Sub New()
        End Sub

        Public Sub New(ByVal SqlFile As String)
            _sqlfile = StripFileName(SqlFile)
            _isPreview = True
        End Sub

        Public Sub New(ByVal DataBaseName As String, ByVal ServerName As String, _
                        ByVal UserId As String, ByVal Password As String, _
                        ByVal SqlFile As String, ByVal CreateProcedures As Boolean, _
                        Optional ByVal PreviewOnly As Boolean = False)

            _isPreview = PreviewOnly
            _connectionString = String.Format("Server={0};User ID={1};Password={2};Initial Catalog={3}", _
                                                ServerName, UserId, Password, DataBaseName)

            _sqlfile = StripFileName(SqlFile)
            _createProcs = CreateProcedures

            If _isPreview Then
                Me.ExecutePreviewMode()
            Else
                Me.Execute()
            End If

        End Sub

        Public ReadOnly Property Log() As String
            Get
                Dim sb As New StringBuilder(String.Empty)
                For Each item As String In _log
                    sb.Append(item & "<br/>")
                Next
                Return sb.ToString
            End Get
        End Property

        Private Function StripFileName(ByVal path As String) As String
            Dim nPos As Integer
            nPos = path.LastIndexOf("\") + 1
            Return path.Substring(3, path.Length - nPos)
        End Function

        Public Function ExecutePreviewMode() As String

            Dim l_success As Boolean = True

            Dim __log As New StringBuilder(String.Empty)

            Try

                Dim uploadFile As String = "C:\CodeGenUploads\" & _sqlfile
                ' read file 
                Dim request As WebRequest = WebRequest.Create(uploadFile)
                Using sr As New StreamReader(request.GetResponse().GetResponseStream())

                    While Not sr.EndOfStream
                        Dim sb As New StringBuilder()

                        While Not sr.EndOfStream
                            Dim s As String = sr.ReadLine()
                            If s IsNot Nothing AndAlso s.ToUpper().Trim().Equals("GO") Then
                                Exit While
                            End If

                            sb.AppendLine(s)
                        End While

                        __log.Append(sb.ToString & "<br/>")
                    End While

                End Using

                __log.Append("Done reading " & _sqlfile & "<br/>")

            Catch ex As Exception
                l_success = False
                __log.Append(String.Format("An error occurred: {0}<br/>", ex.ToString))
            End Try

            Return __log.ToString

        End Function

        Public Function Execute() As String

            Dim l_success As Boolean = True

            _log = New SortedList
            Dim __log As New StringBuilder(String.Empty)

            Dim conn As SqlConnection = Nothing
            Try
                'Me.Response.Write([String].Format("Opening url {0}<BR>", fileUrl))

                ' read file 
                Dim request As WebRequest = WebRequest.Create(_sqlfile)
                Using sr As New StreamReader(request.GetResponse().GetResponseStream())

                    _log.Add(Now.ToLongTimeString, "Connecting to " & _database)
                    __log.Append("Connecting to " & _database)

                    ' Create new connection to database 
                    conn = New SqlConnection(_connectionString)

                    conn.Open()

                    While Not sr.EndOfStream
                        Dim sb As New StringBuilder()
                        Dim cmd As SqlCommand = conn.CreateCommand()

                        While Not sr.EndOfStream
                            Dim s As String = sr.ReadLine()
                            If s IsNot Nothing AndAlso s.ToUpper().Trim().Equals("GO") Then
                                Exit While
                            End If

                            sb.AppendLine(s)
                        End While

                        ' Execute T-SQL against the target database 
                        cmd.CommandText = sb.ToString()
                        cmd.CommandTimeout = _timeout

                        cmd.ExecuteNonQuery()
                    End While

                End Using

                _log.Add(Now.ToLongTimeString, _sqlfile & " file executed successfully")
                __log.Append(_sqlfile & " file executed successfully<br/>")
            Catch ex As Exception
                l_success = False
                _log.Add(Now.ToLongTimeString, String.Format("An error occurred: {0}", ex.ToString))
                __log.Append(String.Format("An error occurred: {0}<br/>", ex.ToString))
            Finally
                ' Close out the connection 
                If conn IsNot Nothing Then
                    Try
                        conn.Close()
                        conn.Dispose()
                    Catch e As Exception
                        l_success = False
                        _log.Add(Now.ToLongTimeString, String.Format("Could not close the connection. Error was {0}", e.ToString()))
                        __log.Append(String.Format("Could not close the connection. Error was {0}<br/>", e.ToString()))
                    End Try
                End If
            End Try

            Return __log.ToString

        End Function

    End Class

    Public Class CodeGen

#Region " PRIVATE/PUBLIC MEMBERS & ENUMERATIONS "

        Private MAX_FILES_TO_CREATE As Integer = 0
        Private _FilesCreated As Integer = 0
        Private _WriteToFile As Boolean = False
        Private _OverwriteExistingFiles As Boolean = False

        Enum METHOD_TYPE
            UPDATE
            INSERT
            QUERY
        End Enum

        Enum OBJECT_TYPE
            [Class] = 0
            [Interface] = 1
        End Enum

#End Region

        Private _database As String
        Private _tables As DataTable
        Private _outputfile As String
        Private _tablePrefix As String
        Private _procPrefix As String
        Private _projectName As String
        Private _zipFilePassword As String

        Private Sub New()
        End Sub

        Public Sub New(ByVal Database As String, _
                       ByVal OutputFile As String, _
                       ByVal TablePrefix As String, _
                       ByVal StoredProcedurePrefix As String, _
                       ByVal ProjectName As String, _
                       Optional ByVal ZipFilePassword As String = "")
            _database = Database
            _outputfile = OutputFile
            _tablePrefix = TablePrefix.ToLower()
            _procPrefix = StoredProcedurePrefix.ToLower()
            _projectName = ProjectName
            _zipFilePassword = ZipFilePassword
        End Sub

        Public Sub New(ByVal Database As String, _
                       ByVal Tables As DataTable, _
                       ByVal OutputFile As String, _
                       ByVal TablePrefix As String, _
                       ByVal StoredProcedurePrefix As String, _
                       ByVal ProjectName As String, _
                       Optional ByVal ZipFilePassword As String = "")

            Me.New(Database, OutputFile, TablePrefix, StoredProcedurePrefix, ProjectName, ZipFilePassword)

            _tables = Tables

        End Sub

        Public Function GetTables() As DataTable
            Dim connString As String = SiteProfile.GetConnectionString
            Dim data As New Global.Data.DataAccess(connString, Global.Data.DataAccess.DataProvider.SQL)
            Dim iCmd As IDbCommand = data.GetCommand("select * from nltcontrolpanel..tbldbmappings where lower(tablename) like '" & _tablePrefix & "%' order by tablename, ordinalposition asc", CommandType.Text, Global.Data.DataAccess.DataProvider.SQL)
            Dim dt As DataTable = data.GetDataTable(iCmd)
            Return dt
        End Function

        Public Function BuildClassAndInterfaceFiles(ByVal TableList As SortedList, ByRef totalOutput As StringBuilder) As Boolean
            Try

                Dim previewer As StringBuilder = New StringBuilder(String.Empty)
                previewer.AppendLine("The following classes and interfaces will need to be created:")

                'Dim totalOutput As StringBuilder = New StringBuilder(String.Empty)
                Dim columnNames As ArrayList = New ArrayList

                Dim sqlOutput As StringBuilder = New StringBuilder(String.Empty)
                Dim sqlColumnList As SortedList = New SortedList

                Dim dt As DataTable

                If _tables IsNot Nothing Then
                    dt = _tables
                Else
                    dt = Me.GetTables
                End If

                For Each tableName As String In TableList.Values
                    Dim classFileOutput As StringBuilder = New StringBuilder(String.Empty)
                    Dim interfaceFileOutput As StringBuilder = New StringBuilder(String.Empty)

                    Dim fetchParams As StringBuilder = New StringBuilder(String.Empty)
                    Dim saveParams As StringBuilder = New StringBuilder(String.Empty)
                    Dim iParmNamesAddedToCmd As StringBuilder = New StringBuilder(String.Empty)

                    Dim fillSets As StringBuilder = New StringBuilder(String.Empty)
                    Dim fillRowFilter As String = String.Empty
                    Dim SubNewSignature As String = String.Empty
                    Dim SubNewSequence As String = String.Empty

                    Dim dvTables As DataView = dt.DefaultView
                    dvTables.RowFilter = "TABLE_NAME = '" & tableName & "'"

                    Dim parsedTableName As String = tableName.Replace(_tablePrefix, "")
                    Dim interfaceName As String = "I" & parsedTableName

                    With totalOutput
                        .AppendLine("'#########################################################################################")
                        .AppendLine("")
                    End With

                    With interfaceFileOutput
                        .AppendLine("Public Interface " & interfaceName)
                        .AppendLine("")
                    End With

                    previewer.AppendLine("'" & parsedTableName & " --> " & interfaceName)

                    Dim propertyDeclaration As StringBuilder = New StringBuilder(String.Empty)
                    Dim classPropertyType As StringBuilder = New StringBuilder(String.Empty)
                    Dim SequenceColumn As String = String.Empty

                    For Each row As DataRowView In dvTables

                        Dim VbDataType As String = String.Empty
                        Dim SqlDataType As String = String.Empty

                        Dim columnName As String = Data.Services.GetNULLableString(row("COLUMN_NAME"))
                        Dim columnPosition As Integer = Data.Services.GetNULLableInteger(row("ORDINAL_POSITION"))
                        Dim maxCharacterLength As String = Data.Services.GetNULLableString(row("CHARACTER_MAXIMUM_LENGTH"))
                        Dim numericPrecision As String = Data.Services.GetNULLableString(row("NUMERIC_PRECISION"))
                        Dim numericScale As String = Data.Services.GetNULLableString(row("NUMERIC_SCALE"))
                        Dim dataType As String = Data.Services.GetNULLableString(row("DATA_TYPE")).ToLower
                        Dim parsedPropertyClass As String = columnName.Remove(columnName.Length - 2, 2)
                        Dim currentPropertyClassInterface As String = "I" & parsedPropertyClass
                        Dim privateMember As String = "_" & columnName.Substring(0, 1).ToLower & columnName.Substring(1, columnName.Length - 1)

                        If columnNames.Contains(columnName) = False Then
                            columnNames.Add(columnName)
                        End If

                        Select Case dataType
                            Case "bit"
                                VbDataType = "Boolean"
                                SqlDataType = "bit"
                            Case "char", "varchar", "text", "nvarchar"
                                VbDataType = "String"
                                SqlDataType = dataType & "(" & maxCharacterLength & ")"
                            Case "uniqueidentifier"
                                VbDataType = "System.Guid"
                                SqlDataType = "uniqueidentifier"
                            Case "int"
                                VbDataType = "Integer"
                                SqlDataType = "int"
                            Case "datetime"
                                VbDataType = "DateTime"
                                SqlDataType = "datetime"
                            Case "money"
                                VbDataType = "Double"
                                SqlDataType = "money"
                            Case "decimal"
                                VbDataType = "Decimal"
                                SqlDataType = "decimal(" & numericPrecision & "," & numericScale & ")"
                            Case Else
                                VbDataType = "Object"
                                SqlDataType = ""
                        End Select

                        Dim interfacePropertyType As String = String.Empty
                        Dim parsedPropertyClassAsVariable As String = "_" & parsedPropertyClass.Substring(0, 1).ToLower & parsedPropertyClass.Substring(1, parsedPropertyClass.Length - 1)
                        If columnName.ToUpper.EndsWith("ID") And columnPosition > 1 Then

                            interfacePropertyType = "    Property " & parsedPropertyClass & "() As " & currentPropertyClassInterface

                            Dim varName As String = "o" & parsedPropertyClass
                            fillSets.AppendLine("")
                            fillSets.AppendLine("                Dim " & varName & " As " & currentPropertyClassInterface & " = New " & parsedPropertyClass & "(Services.GetNULLableInteger(row(""" & parsedPropertyClass & "ID"")))")
                            fillSets.AppendLine("                Me." & parsedPropertyClassAsVariable & " = " & varName)
                            fillSets.AppendLine("")

                            propertyDeclaration.AppendLine("    Private " & parsedPropertyClassAsVariable & " As " & currentPropertyClassInterface)

                            saveParams.AppendLine("        " & Me.BuildDbParam(columnName, columnPosition, parsedPropertyClassAsVariable, maxCharacterLength, numericPrecision, numericScale, currentPropertyClassInterface, METHOD_TYPE.INSERT))
                        Else
                            interfacePropertyType = "    Property " & columnName & "() As " & VbDataType
                            propertyDeclaration.AppendLine("    Private " & privateMember & " As " & VbDataType)

                            'Exclude the primary key in this case because we do this in our Save() method by checking if the current variable for the class is "0" or not.
                            If privateMember.ToLower().Contains(parsedTableName.ToLower()) = False Then
                                saveParams.AppendLine("        " & Me.BuildDbParam(columnName, columnPosition, privateMember, maxCharacterLength, numericPrecision, numericScale, VbDataType, METHOD_TYPE.INSERT))
                            End If

                        End If
                        interfaceFileOutput.AppendLine(interfacePropertyType)

                        iParmNamesAddedToCmd.AppendLine("            .Add(iParm" & columnName & ")")

                        If columnPosition = 1 Then
                            fillRowFilter = "dv.RowFilter = """ & columnName & " = "" & Me." & privateMember & ".ToString"
                            SubNewSignature = "Sub New(ByVal " & columnName & " As " & VbDataType & ")"
                            SubNewSequence = "Me." & privateMember & " = " & columnName
                        End If

                        classPropertyType.AppendLine("")
                        If columnName.ToUpper.EndsWith("ID") And columnPosition > 1 Then

                            classPropertyType.AppendLine("    ''' <summary>")
                            classPropertyType.AppendLine("    ''' ")
                            classPropertyType.AppendLine("    ''' </summary>")
                            classPropertyType.AppendLine("    ''' <value></value>")
                            classPropertyType.AppendLine("    ''' <returns></returns>")
                            classPropertyType.AppendLine("    ''' <remarks></remarks>")
                            classPropertyType.AppendLine("    Property " & parsedPropertyClass & "() As " & currentPropertyClassInterface & " Implements " & interfaceName & "." & parsedPropertyClass)
                            classPropertyType.AppendLine("        Get")
                            classPropertyType.AppendLine("            Return " & parsedPropertyClassAsVariable)
                            classPropertyType.AppendLine("        End Get")
                            classPropertyType.AppendLine("        Set(ByVal value As " & currentPropertyClassInterface & ")")
                            classPropertyType.AppendLine("            value = " & parsedPropertyClassAsVariable)
                            classPropertyType.AppendLine("        End Set")
                            classPropertyType.AppendLine("    End Property")
                        Else
                            classPropertyType.AppendLine("    ''' <summary>")
                            classPropertyType.AppendLine("    ''' ")
                            classPropertyType.AppendLine("    ''' </summary>")
                            classPropertyType.AppendLine("    ''' <value></value>")
                            classPropertyType.AppendLine("    ''' <returns></returns>")
                            classPropertyType.AppendLine("    ''' <remarks></remarks>")
                            classPropertyType.AppendLine("    " & interfacePropertyType & " Implements " & interfaceName & "." & columnName)
                            classPropertyType.AppendLine("        Get")
                            classPropertyType.AppendLine("            Return " & privateMember)
                            classPropertyType.AppendLine("        End Get")
                            classPropertyType.AppendLine("        Set(ByVal value As " & VbDataType & ")")
                            classPropertyType.AppendLine("            value = " & privateMember)
                            classPropertyType.AppendLine("        End Set")
                            classPropertyType.AppendLine("    End Property")
                        End If

                        sqlColumnList.Add(columnPosition, columnName)

                    Next

                    Dim privateMemberVariableFromTableName As String = "_" & parsedTableName.Substring(0, 1).ToLower & parsedTableName.Substring(1, parsedTableName.Length - 1) & "ID"

                    Dim columnList As String = String.Empty
                    For Each item As String In sqlColumnList.Values
                        columnList += item & ", "
                    Next

                    sqlColumnList.Clear()

                    With interfaceFileOutput
                        .AppendLine("")
                        .AppendLine("    Sub Fill()")
                        .AppendLine("    Function GetList() As System.Data.DataTable")
                        .AppendLine("    Function Save() As Boolean")
                        .AppendLine("    Function Delete() As Boolean")
                        .AppendLine("")
                        .AppendLine("End Interface")
                        .AppendLine("")
                    End With

                    With classFileOutput
                        .AppendLine("#Region "" IMPORTS """)
                        .AppendLine("Imports Data")
                        .AppendLine("Imports System.Data")
                        .AppendLine("#End Region")
                        .AppendLine("")
                        .AppendLine("''' <summary>")
                        .AppendLine("''' ")
                        .AppendLine("''' </summary>")
                        .AppendLine("''' <remarks>Auto-generated by New Leaf Technologies, Inc. Code Generator on " & Now.ToString & "</remarks>")
                        .AppendLine("Public Class " & parsedTableName)
                        .AppendLine("   Implements " & interfaceName)
                        .AppendLine("")
                        .AppendLine("#Region "" PRIVATE/PUBLIC MEMBERS """)
                        .AppendLine("")
                        .AppendLine(propertyDeclaration.ToString)
                        .AppendLine("#End Region")
                        .AppendLine("")
                        .AppendLine("#Region "" PROPERTIES """)
                        .AppendLine("        " & classPropertyType.ToString)
                        .AppendLine("#End Region")
                        .AppendLine("")
                        .AppendLine("#Region "" INSTANTIATE """)
                        .AppendLine("")
                        .AppendLine("    Sub New()")
                        .AppendLine("    End Sub")
                        .AppendLine("")
                        .AppendLine("    " & SubNewSignature)
                        .AppendLine("        MyBase.New()")
                        .AppendLine("")
                        .AppendLine("        " & SubNewSequence)
                        .AppendLine("        Me.Fill()")
                        .AppendLine("")
                        .AppendLine("    End Sub")
                        .AppendLine("")
                        .AppendLine("#End Region")
                        .AppendLine("")
                        .AppendLine("#Region "" METHODS """)
                        .AppendLine("")
                        .AppendLine("    ''' <summary>")
                        .AppendLine("    ''' ")
                        .AppendLine("    ''' </summary>")
                        .AppendLine("    ''' <remarks></remarks>")
                        .AppendLine("    Public Sub Fill() Implements " & interfaceName & ".Fill")
                        .AppendLine("")
                        .AppendLine("        Try")
                        .AppendLine("            Dim dv As DataView = Me.GetList().DefaultView")
                        .AppendLine("            " & fillRowFilter)
                        .AppendLine("            For Each row As DataRowView In dv")
                        .Append(fillSets.ToString)
                        .AppendLine("            Next")
                        .AppendLine("        Catch ex As Exception")
                        .AppendLine("            Throw New NLTException(""Error retrieving and filling [" & parsedTableName.ToUpper & "] information"", ex, """ & parsedTableName & ".vb"", ""Sub Fill()"")")
                        .AppendLine("        End Try")
                        .AppendLine("")
                        .AppendLine("    End Sub")
                        .AppendLine("")
                        .AppendLine("    ''' <summary>")
                        .AppendLine("    ''' ")
                        .AppendLine("    ''' </summary>")
                        .AppendLine("    ''' <returns></returns>")
                        .AppendLine("    ''' <remarks></remarks>")
                        .AppendLine("    Public Function GetList() as System.Data.DataTable Implements " & interfaceName & ".GetList")
                        .AppendLine("")
                        .AppendLine("        Try")
                        .AppendLine("            Dim connString = SiteProfile.GetConnectionString()")
                        .AppendLine("            Dim data As New DataAccess(connString, DataAccess.DataProvider.SQL)")
                        .AppendLine("            Dim iCmd As IDbCommand = data.GetCommand(""" & _procPrefix & "Get" & parsedTableName & "ById"", CommandType.StoredProcedure, DataAccess.DataProvider.SQL)")
                        .AppendLine("            Return data.GetDataTable(iCmd)")
                        .AppendLine("        Catch ex As Exception")
                        .AppendLine("            Throw New NLTException(""Error retrieving [" & parsedTableName.ToUpper & "] information"", ex, """ & parsedTableName & ".vb"", ""Function GetList()"")")
                        .AppendLine("        End Try")
                        .AppendLine("")
                        .AppendLine("    End Function")
                        .AppendLine("")
                        .AppendLine("    ''' <summary>")
                        .AppendLine("    ''' ")
                        .AppendLine("    ''' </summary>")
                        .AppendLine("    ''' <returns></returns>")
                        .AppendLine("    ''' <remarks></remarks>")
                        .AppendLine("    Public Function Save() as Boolean Implements " & interfaceName & ".Save")
                        .AppendLine("")
                        .AppendLine("        Dim success As Boolean = False")
                        .AppendLine("        Dim iCmd As IDbCommand = Nothing")
                        .AppendLine("        Dim iParm" & parsedTableName & "ID As IDbDataParameter = Nothing")
                        .AppendLine("        Dim data As New DataAccess(SiteProfile.GetConnectionString, DataAccess.DataProvider.SQL)")
                        .AppendLine("        Dim dict As New System.Collections.Specialized.HybridDictionary")
                        .AppendLine("        Dim storedProc As String = String.Empty")
                        .AppendLine("")
                        .AppendLine("        If Me." & privateMemberVariableFromTableName & " = 0 Then")
                        .AppendLine("            storedProc = """ & _procPrefix & "Add" & parsedTableName & """")
                        .AppendLine("            " & Me.BuildDbParam(parsedTableName & "ID", 1, privateMemberVariableFromTableName, "4", "", "", "Integer", METHOD_TYPE.INSERT, True))
                        .AppendLine("        Else")
                        .AppendLine("            storedProc = """ & _procPrefix & "Update" & parsedTableName & """")
                        .AppendLine("            " & Me.BuildDbParam(parsedTableName & "ID", 1, privateMemberVariableFromTableName, "4", "", "", "Integer", METHOD_TYPE.UPDATE, True))
                        .AppendLine("        End If")
                        .AppendLine("")
                        .AppendLine("        iCmd = data.GetCommand(storedProc, CommandType.StoredProcedure, DataAccess.DataProvider.SQL)")
                        .AppendLine(saveParams.ToString)
                        .AppendLine("        With iCmd.Parameters")
                        .Append(iParmNamesAddedToCmd.ToString)
                        .AppendLine("        End With")
                        .AppendLine("")
                        .AppendLine("        dict.Add(dict.Count, iCmd)")
                        .AppendLine("")
                        .AppendLine("        Try")
                        .AppendLine("            success = data.ExecuteNonQuery(dict)")
                        .AppendLine("            If success = True Then")
                        .AppendLine("                If Me." & privateMemberVariableFromTableName & " = 0 Then")
                        .AppendLine("                    Me." & privateMemberVariableFromTableName & " = iParm" & parsedTableName & "ID.Value")
                        .AppendLine("                End If")
                        .AppendLine("            End If")
                        .AppendLine("            Return success")
                        .AppendLine("        Catch ex As Exception")
                        .AppendLine("            Throw New NLTException(""Error updating [" & parsedTableName.ToUpper & "]"", ex, """ & parsedTableName & ".vb"", ""Function Save() As Boolean"")")
                        .AppendLine("        Finally")
                        .AppendLine("            If iCmd.Connection.State <> ConnectionState.Closed Then")
                        .AppendLine("                iCmd.Connection.Close()")
                        .AppendLine("            End If")
                        .AppendLine("        End Try")
                        .AppendLine("")
                        .AppendLine("    End Function")
                        .AppendLine("")
                        .AppendLine("    ''' <summary>")
                        .AppendLine("    ''' ")
                        .AppendLine("    ''' </summary>")
                        .AppendLine("    ''' <returns></returns>")
                        .AppendLine("    ''' <remarks></remarks>")
                        .AppendLine("    Public Function Delete() as Boolean Implements " & interfaceName & ".Delete")
                        .AppendLine("")
                        .AppendLine("        Dim success As Boolean = False")
                        .AppendLine("        Dim iCmd As IDbCommand = Nothing")
                        .AppendLine("        Dim data As New DataAccess(SiteProfile.GetConnectionString, DataAccess.DataProvider.SQL)")
                        .AppendLine("        Dim dict As New System.Collections.Specialized.HybridDictionary")
                        .AppendLine("")
                        .AppendLine("        iCmd = data.GetCommand(""" & _procPrefix & "Delete" & parsedTableName & """, CommandType.StoredProcedure, DataAccess.DataProvider.SQL)")
                        .AppendLine("        " & Me.BuildDbParam(parsedTableName & "ID", 1, privateMemberVariableFromTableName, "4", "", "", "Integer", METHOD_TYPE.UPDATE, False))
                        .AppendLine("")
                        .AppendLine("        With iCmd.Parameters")
                        .AppendLine("            .Add(iParm" & parsedTableName & "ID)")
                        .AppendLine("        End With")
                        .AppendLine("")
                        .AppendLine("        dict.Add(dict.Count, iCmd)")
                        .AppendLine("")
                        .AppendLine("        Try")
                        .AppendLine("            success = data.ExecuteNonQuery(dict)")
                        .AppendLine("            Return success")
                        .AppendLine("        Catch ex As Exception")
                        .AppendLine("            Throw New NLTException(""Error deleting [" & parsedTableName.ToUpper & "]"", ex, """ & parsedTableName & ".vb"", ""Function Delete() As Boolean"")")
                        .AppendLine("        Finally")
                        .AppendLine("            If iCmd.Connection.State <> ConnectionState.Closed Then")
                        .AppendLine("                iCmd.Connection.Close()")
                        .AppendLine("            End If")
                        .AppendLine("        End Try")
                        .AppendLine("")
                        .AppendLine("    End Function")
                        .AppendLine("")
                        .AppendLine("#End Region")
                        .AppendLine("")
                        .AppendLine("End Class")
                        .AppendLine("")
                    End With

                    totalOutput.Append(interfaceFileOutput.ToString)
                    totalOutput.Append(classFileOutput.ToString)

                    Call SaveToFile(interfaceFileOutput.ToString, interfaceName & ".vb", OBJECT_TYPE.Interface)
                    Call SaveToFile(classFileOutput.ToString, parsedTableName & ".vb", OBJECT_TYPE.Class)

                Next

                Try
                    Using zip As Global.Ionic.Zip.ZipFile = New Global.Ionic.Zip.ZipFile
                        Dim dir As String = My.Request.PhysicalApplicationPath & "\Downloads\" & _projectName
                        zip.AddDirectory(dir, System.IO.Path.GetFileName(dir))
                        zip.Comment = "Autogenerated by New Leaf Technologies, Inc. Code Generator"
                        If Me._zipFilePassword.Equals(String.Empty) = False Then
                            zip.Password = Me._zipFilePassword
                        End If
                        zip.Save(_projectName & "_CodeFiles.zip")
                        File.Move(zip.Name, My.Request.PhysicalApplicationPath & "\Downloads\" & zip.Name)
                    End Using
                Catch ex As Exception
                    Throw ex
                End Try


                Dim folder As New System.IO.DirectoryInfo(My.Request.PhysicalApplicationPath & "\Downloads\" & _projectName)
                If folder.Exists Then
                    Try
                        folder.Delete(True)
                    Catch ex As Exception
                        Throw New ApplicationException("An error occurred while attempting to delete the project directory", ex)
                    End Try

                End If

            Catch ex As Exception
                Throw ex
            End Try

            Return True

        End Function

        Private Function GetNullableType(ByVal DataType As String, ByVal value As String) As String

            Select Case DataType
                Case "Boolean"
                    Return "Data.Services.GetNULLableBoolean(" & value & ")"
                Case "String", "System.Guid"
                    Return "Data.Services.GetNULLableString(" & value & ")"
                Case "Integer"
                    Return "Data.Services.GetNULLableInteger(" & value & ")"
                Case "DateTime"
                    Return "Data.Services.GetNULLableDateTime(" & value & ")"
                Case "Double"
                    Return "Data.Services.GetNULLableDouble(" & value & ")"
                Case "Decimal"
                    Return "Data.Services.GetNULLableDecimal(" & value & ")"
                Case Else
                    Return String.Empty
            End Select

        End Function

        Private Function BuildDbParam(ByVal ColumnName As String, _
                                        ByVal OrdinalPosition As Integer, _
                                        ByVal MemberVariable As String, _
                                        ByVal DataTypeLength As String, _
                                        ByVal NumericPrecision As String, _
                                        ByVal NumericScale As String, _
                                        ByVal DataType As String, _
                                        ByVal MethodType As METHOD_TYPE, _
                                        Optional ByVal IgnoreDeclaration As Boolean = False) As String

            Dim parameter As String = String.Empty
            If IgnoreDeclaration = True Then
                parameter = "iParm" & ColumnName & " = data.GetParameter(DataAccess.DataProvider.SQL, ""@" & ColumnName & """, "
            Else
                parameter = "Dim iParm" & ColumnName & " As IDbDataParameter = data.GetParameter(DataAccess.DataProvider.SQL, ""@" & ColumnName & """, "
            End If

            Select Case DataType.ToLower
                Case "boolean"
                    parameter += "DbType.Boolean, Me." & MemberVariable & ", 1, ParameterDirection.Input)"
                Case "system.guid"
                    parameter += "DbType.Guid, Me." & MemberVariable & ", Me." & MemberVariable & ".Length, ParameterDirection.Input)"
                Case "string"
                    parameter += "DbType.String, Me." & MemberVariable & ", " & DataTypeLength & ", ParameterDirection.Input)"
                Case "integer"
                    If MethodType = METHOD_TYPE.INSERT Then
                        If OrdinalPosition = 1 Then
                            parameter += "DbType.Int32, Nothing, 4, ParameterDirection.Output)"
                        Else
                            parameter += "DbType.Int32, Me." & MemberVariable & ", 4, ParameterDirection.Input)"
                        End If
                    Else
                        parameter += "DbType.Int32, Me." & MemberVariable & ", 4, ParameterDirection.Input)"
                    End If
                Case "datetime"
                    parameter += "DbType.DateTime, Me." & MemberVariable & ", Nothing, ParameterDirection.Input)"
                Case "double"
                    parameter += "DbType.Double, Me." & MemberVariable & ", Nothing, ParameterDirection.Input, " & NumericPrecision & ", " & NumericScale & ", " & ColumnName & ")"
                Case "decimal"
                    parameter += "DbType.Decimal, Me." & MemberVariable & ", Nothing, ParameterDirection.Input, " & NumericPrecision & ", " & NumericScale & ", " & ColumnName & ")"
                Case "money"
                    parameter += "DbType.Currency, Me." & MemberVariable & ", Nothing, ParameterDirection.Input, " & NumericPrecision & ", " & NumericScale & ", " & ColumnName & ")"
                Case Else
                    If MethodType = METHOD_TYPE.INSERT Or MethodType = METHOD_TYPE.UPDATE Then
                        If OrdinalPosition > 1 Then
                            If DataType.StartsWith("I") And MemberVariable.ToUpper.EndsWith("ID") = False Then 'Interface Type
                                parameter += "DbType.Int32, Me." & MemberVariable & "." & ColumnName & ", 4, ParameterDirection.Input)"
                            End If
                        Else
                            parameter += "DbType.Int32, Me." & MemberVariable & ", 4, ParameterDirection.Input)"
                        End If
                    Else
                        Return String.Empty
                    End If
            End Select

            Return parameter

        End Function

        Private Sub SaveToFile(ByVal FileContents As String, ByVal ObjectFileName As String, ByVal ObjectType As OBJECT_TYPE)

            If _FilesCreated < MAX_FILES_TO_CREATE Or Me.MAX_FILES_TO_CREATE = 0 Then

                Dim folder As System.IO.DirectoryInfo = Nothing
                Select Case ObjectType
                    Case OBJECT_TYPE.Class
                        folder = New System.IO.DirectoryInfo(My.Request.PhysicalApplicationPath & "\Downloads\" & _projectName & "\App_Code\")
                    Case OBJECT_TYPE.Interface
                        folder = New System.IO.DirectoryInfo(My.Request.PhysicalApplicationPath & "\Downloads\" & _projectName & "\App_Code\Interfaces\")
                End Select

                If folder.Exists = False Then
                    Try
                        folder.Create()
                    Catch ex As Exception
                        Throw New ApplicationException("An error occurred while attempting to create directorys to load project files in", ex)
                    End Try

                End If

                Dim FilePathSave As String = folder.ToString() & ObjectFileName
                Dim file As New System.IO.FileInfo(FilePathSave)
                If file.Exists Then
                    If Me._OverwriteExistingFiles = True Then
                        Try
                            file.Delete()
                            System.IO.File.WriteAllText(FilePathSave, FileContents)
                        Catch ex As Exception
                            Throw New ApplicationException("Error occurred while attempting to delete and add new project files.", ex)
                        End Try
                    End If
                Else
                    Try
                        System.IO.File.WriteAllText(FilePathSave, FileContents)
                    Catch ex As Exception
                        Throw New ApplicationException("An error occurred while attempting to create new project files", ex)
                    End Try


                End If

                Me._FilesCreated += 1
            End If


        End Sub
    End Class

End Namespace