#Region " IMPORTS "
Imports Data
Imports System.Data
Imports System.Diagnostics
#End Region

Namespace NewLeaf

    Public Class CodeGenerator

#Region " USAGE INSTRUCTIONS "
        ''' <summary>
        ''' CREATED BY: Brian Gaines
        ''' CREATED ON: 9/22/2008
        ''' 
        ''' This code generator was designed for a website project in Visual Studio 2005 running ASP.NET 2.0 using the VB.NET langauge. 
        ''' The database scripts were designed for SQL Server 2005; however, you should be able to make any modifications for earlier versions.
        ''' 
        ''' NOTES:
        ''' Be aware that this is not a be-all end-all code generation utility, so there will be corrections that you will need to make. 
        ''' I hope to cover what you may need to change in the next section below. I created this utility as a means for expediting my 
        ''' development efforts based on an object model or design pattern I like to code against. I wanted to eliminate the efforts I 
        ''' commonly put forth of mundane tasks like creating stored procedures for each and every table, creating interfaces that my 
        ''' class objects implement, declaring the member variables, instantiating other classes based on a property that references 
        ''' such class, and creating common methods across all my classes, which includes the tedious task of creating parameters.
        ''' 
        ''' In conjunction with this .NET code generator, you will also need to run the stored procedures and functions in the script 
        ''' section below "STORED PROCEDURE CODE GENERATOR SCRIPTS" on the database you want to generate your stored procedures for. 
        ''' These stored procedures and helper functions will auto-generate common stored procedures for each table that you have in 
        ''' a database. They will create the following stored procedures:
        ''' 1) sp__Get[TABLE_NAME]ById
        ''' 2) sp__Add[TABLE_NAME]
        ''' 3) sp__Update[TABLE_NAME]
        ''' 4) sp__Delete[TABLE_NAME]
        ''' 
        ''' where [TABLE_NAME] strips off the "tbl" prefix, but leaves the rest of the name in tact.
        ''' 
        ''' Disclaimer: The stored procedure auto-generator was not created by New Leaf Technologies, Inc; however, we have tailored 
        ''' them to work for our needs. 
        ''' 
        ''' Link to article:
        ''' http://tinyurl.com/dynamicSql-article
        ''' 
        ''' Link to sql code file:
        ''' http://tinyurl.com/dynamicSql
        ''' 
        ''' INSTRUCTIONS:
        ''' 1. Design and create your database with the following standards
        ''' a) Prefix your table name with "tbl" (i.e. tblVehicle)
        ''' b) Create primary and foreign key references as you normally would; however, make sure that you:
        '''    - Each table should have an identity column. I make this the primary key for the table
        '''    - Name the primary key for each table the same as the table, but remove "tbl" and add "ID" (i.e. VehicleID)
        '''    - Name your foreign keys on a relationship table the same as you would on the base table. These columns will join 
        '''      with each other. (i.e. tblVehicle.VehicleTypeID inner join tblVehicleType.VehicleTypeID)
        ''' c) DO NOT name any of your columns the same name as your table.
        ''' 2. Execute "sp_NLT_GenerateSqlProcs @bExecute = 1" stored procedure to create your stored procedures (4 per table)
        ''' 3. If you made any any changes to the convention, you will likely need to make some changes to this .NET code generation
        '''    file.
        ''' 4. Okay, the database is designed the way you want it with all the columns you want. Now it is time to generate your class
        '''    files.
        ''' 5. The first thing is to make sure that you have all the required files that I list out below
        ''' 6. This particular class file will need to be loaded in the project you want to create the class files in.
        ''' 7. Add the following code below to a place where you want to generate the class files from. I designed this to print out 
        '''    to a web page, so if this is for a windows application or other type of application, then modifications will need to be 
        '''    accommmodated for it.
        ''' 
        ''' NOTE: If you want to only generate a couple of files to see what they look like; thus testing the system, then I've added 
        '''       the ability for a maximum number of files to be created in the constant [MAX_FILES_TO_CREATE]. If it's set to 0, then 
        '''       all files will be created.
        ''' 
        ''' Argument 1: Create the files, if not and you just want to see what the generated code looks like, then set to False
        ''' Argument 2: Overwrite existing files if they exist; however, if set to True, this would like error in the beginning because 
        '''             changes or fixes would be likely to accommodate. (Or, you could just manually delete the files yourself).
        ''' 
        ''' Dim code As New NLT.CodeGenerator(True, True)
        ''' code.BuildClassAndInterfaceFiles()
        ''' 
        ''' 8. Run the code. It should only take a brief amount of time to generate all of your class and interface objects and when it
        '''    completes, an output should display on the page, which you are executing this code from. You can now stop debug mode 
        '''    and "Refresh" the project, you should see the new files in the "\App_Code\BLL\" and "\App_Code\BLL\Interfaces" folders.
        ''' 
        ''' 9. Review the output in each file and make any further adjustments to get it working the way you need it.
        ''' 
        ''' 10. You're going to love this one, but now that we have written out the *.vb files, compile, we see that we do not have
        '''    build errors, but Visual Studio.NET 2005 tells us that the build failed. I researched this and don't understand it myself. 
        '''    I did find this forum posting of others experiencing the same issue. The fix for me was to do the following:
        '''    a) Open each newly written out file created by the code generator and comment out ALL lines of code.
        '''    b) Recompile your application; the build should be successful.
        '''    c) Uncomment ALL lines of code you JUST commented out and recompile again.
        '''    d) You should now see that the website compiled successfully without errors.
        '''    e) If this does not work for you, you may want to look at the following link for ideas.
        '''    f) TIP: I know that by manually adding a new class from the templates in VS.NET, you can copy and paste the entire output from the 
        '''       previewer window into the new class file (removing the #Region..#End Region with the Imports statements); It will compile, but
        '''       everything resides in a single file, which is not our ideal situation, but it will get you moving forward.
        ''' http://www.velocityreviews.com/forums/t361938-website-wont-build-build-failed-with-no-errors-warnings-or-messages-vs2005.html
        ''' 
        ''' CLASS STRUCTURE:
        ''' The output of the classes are as such:
        ''' 1. One class and interface file for each database table
        ''' 2. Each class will have a set of properties that map to a column in a table
        ''' 3. If the column/property is a foreign key to another another table/class, then the property type will be of that type
        '''    (i.e Property VehicleType() as IVehicleType) and the constructor of the class will attempt to fetch an instance of that 
        '''    class based on the ID (i.e. VehicleTypeID) and attach it to the current class as a property. This is repeated for each
        '''    foreign key reference
        ''' 4. Each class has the following methods to interact with the database
        ''' a) Fill() - Fetch records (calls GetList()) and populate all the properties with information.
        ''' b) GetList() - Fetch records from the database based on the current ID
        ''' c) Save() - Insert or Update the database depending on whether the current class instance has data in it.
        ''' d) Delete() - Delete record based on the current ID
        ''' 
        ''' REQUIRED FILES:
        ''' DataAccess.vb - This is our DAL that we base all communications with the database through
        ''' IDataAccess.vb - This is the interface that the DataAccess file implements
        ''' Services.vb - These are some helper functions that are used by this utility
        ''' NLTException.vb - Custom exception class
        ''' NewLeafSettings.vb - Shared methods that retrieve information from web.config
        ''' </summary>
        ''' <remarks></remarks>
#End Region

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

#Region " INSTANTIATE "

        Sub New()
            Me._WriteToFile = False
        End Sub

        Sub New(ByVal CreateFiles As Boolean, ByVal OverwriteIfExists As Boolean)
            Me._WriteToFile = CreateFiles
            Me._OverwriteExistingFiles = OverwriteIfExists
        End Sub

#End Region

#Region " PRIVATE METHODS "

        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="FileContents"></param>
        ''' <param name="ObjectFileName"></param>
        ''' <param name="ObjectType"></param>
        ''' <remarks></remarks>
        Private Sub SaveToFile(ByVal FileContents As String, ByVal ObjectFileName As String, ByVal ObjectType As OBJECT_TYPE)

            If _FilesCreated < MAX_FILES_TO_CREATE Or Me.MAX_FILES_TO_CREATE = 0 Then

                Dim folder As System.IO.DirectoryInfo = Nothing
                Select Case ObjectType
                    Case OBJECT_TYPE.Class
                        folder = New System.IO.DirectoryInfo(My.Request.PhysicalApplicationPath & "\Private\Downloads\")
                    Case OBJECT_TYPE.Interface
                        folder = New System.IO.DirectoryInfo(My.Request.PhysicalApplicationPath & "\Private\Downloads\")
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

                            Debug.WriteLine("#####################################")
                            Debug.WriteLine("File Deleted on " & Now.ToString)
                            Debug.WriteLine(folder.ToString() & ObjectFileName)

                            System.IO.File.WriteAllText(FilePathSave, FileContents)

                            Debug.WriteLine("#####################################")
                            Debug.WriteLine("File Created on " & Now.ToString)
                            Debug.WriteLine(folder.ToString() & ObjectFileName)

                        Catch ex As Exception
                            Throw New ApplicationException("Error occurred while attempting to delete and add new project files.", ex)
                        End Try

                    Else
                        Debug.WriteLine("#####################################")
                        Debug.WriteLine("The following file was NOT created as a result of a parameter setting passed to the code generator class:")
                        Debug.WriteLine(folder.ToString() & ObjectFileName)
                    End If
                Else

                    Try
                        System.IO.File.WriteAllText(FilePathSave, FileContents)

                        Debug.WriteLine("#####################################")
                        Debug.WriteLine("File Created on " & Now.ToString)
                        Debug.WriteLine(folder.ToString() & ObjectFileName)

                    Catch ex As Exception
                        Throw New ApplicationException("An error occurred while attempting to create new project files", ex)
                    End Try


                End If

                Me._FilesCreated += 1
            End If


        End Sub

        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="ColumnName"></param>
        ''' <param name="OrdinalPosition"></param>
        ''' <param name="MemberVariable"></param>
        ''' <param name="DataTypeLength"></param>
        ''' <param name="NumericPrecision"></param>
        ''' <param name="NumericScale"></param>
        ''' <param name="DataType"></param>
        ''' <param name="MethodType"></param>
        ''' <param name="IgnoreDeclaration"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function BuildDbParam(ByVal ColumnName As String, ByVal OrdinalPosition As Integer, ByVal MemberVariable As String, ByVal DataTypeLength As String, ByVal NumericPrecision As String, ByVal NumericScale As String, ByVal DataType As String, ByVal MethodType As METHOD_TYPE, Optional ByVal IgnoreDeclaration As Boolean = False) As String

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

        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="DataType"></param>
        ''' <param name="value"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function GetNullableType(ByVal DataType As String, ByVal value As String) As String

            Select Case DataType
                Case "Boolean"
                    Return "Services.GetNULLableBoolean(" & value & ")"
                Case "String", "System.Guid"
                    Return "Services.GetNULLableString(" & value & ")"
                Case "Integer"
                    Return "Services.GetNULLableInteger(" & value & ")"
                Case "DateTime"
                    Return "Services.GetNULLableDateTime(" & value & ")"
                Case "Double"
                    Return "Services.GetNULLableDouble(" & value & ")"
                Case "Decimal"
                    Return "Services.GetNULLableDecimal(" & value & ")"
                Case Else
                    Return String.Empty
            End Select

        End Function

        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="Value"></param>
        ''' <remarks></remarks>
        Private Sub OutputDesignWrapper(ByVal Value As StringBuilder)

            Dim container As StringBuilder = New StringBuilder(String.Empty)
            Dim output As String = Value.ToString
            output = output.Replace(" ", "&nbsp;")
            output = output.Replace(vbCrLf, "<br/>")

            container.AppendFormat("<div style=""{0}"">{1}</div>", _
                                    "font-family:'Courier New';font-size:11px;padding:10;border-top::1px solid #666666;border-right:1px solid #666666;border-bottom:1px solid #666666;border-left:1px solid #666666;background-color:#eeeeee;", _
                                    output)

            HttpContext.Current.Response.Write(container.ToString)

        End Sub

#End Region

#Region " PUBLIC METHODS "

        Public Function GetDistinctTables() As DataTable
            Dim dt As DataTable = Me.GetTables()
            Dim dtTables As DataTable = dt.DefaultView.ToTable(True, "TABLE_NAME")
            Return dtTables
        End Function

        Public Function GetTables() As DataTable
            Dim connString As String = NewLeafSettings.GetConnectionString
            Dim data As New DataAccess(connString, DataAccess.DataProvider.SQL)
            Dim iCmd As IDbCommand = data.GetCommand("select * from information_schema.columns where lower(table_name) like 'tbl%' order by table_name, ordinal_position asc", CommandType.Text, DataAccess.DataProvider.SQL)
            Dim dt As DataTable = data.GetDataTable(iCmd)
            Return dt
        End Function


        Public Function BuildClassAndInterfaceFiles(ByVal TableList As SortedList) As Boolean
            Try

                Dim previewer As StringBuilder = New StringBuilder(String.Empty)
                previewer.AppendLine("The following classes and interfaces will need to be created:")

                Dim totalOutput As StringBuilder = New StringBuilder(String.Empty)
                Dim columnNames As ArrayList = New ArrayList

                Dim sqlOutput As StringBuilder = New StringBuilder(String.Empty)
                Dim sqlColumnList As SortedList = New SortedList

                Dim dt As DataTable = Me.GetTables

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

                    Dim parsedTableName As String = tableName.Replace("tbl", "")
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

                        Dim columnName As String = Services.GetNULLableString(row("COLUMN_NAME"))
                        Dim columnPosition As Integer = Services.GetNULLableInteger(row("ORDINAL_POSITION"))
                        Dim maxCharacterLength As String = Services.GetNULLableString(row("CHARACTER_MAXIMUM_LENGTH"))
                        Dim numericPrecision As String = Services.GetNULLableString(row("NUMERIC_PRECISION"))
                        Dim numericScale As String = Services.GetNULLableString(row("NUMERIC_SCALE"))
                        Dim dataType As String = Services.GetNULLableString(row("DATA_TYPE")).ToLower
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
                        .AppendLine("            Dim connString = NewLeafSettings.GetConnectionString()")
                        .AppendLine("            Dim data As New DataAccess(connString, DataAccess.DataProvider.SQL)")
                        .AppendLine("            Dim iCmd As IDbCommand = data.GetCommand(""sp__Get" & parsedTableName & "ById"", CommandType.StoredProcedure, DataAccess.DataProvider.SQL)")
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
                        .AppendLine("        Dim data As New DataAccess(NewLeafSettings.GetConnectionString, DataAccess.DataProvider.SQL)")
                        .AppendLine("        Dim dict As New System.Collections.Specialized.HybridDictionary")
                        .AppendLine("        Dim storedProc As String = String.Empty")
                        .AppendLine("")
                        .AppendLine("        If Me." & privateMemberVariableFromTableName & " = 0 Then")
                        .AppendLine("            storedProc = ""sp__Add" & parsedTableName & """")
                        .AppendLine("            " & Me.BuildDbParam(parsedTableName & "ID", 1, privateMemberVariableFromTableName, "4", "", "", "Integer", METHOD_TYPE.INSERT, True))
                        .AppendLine("        Else")
                        .AppendLine("            storedProc = ""sp__Update" & parsedTableName & """")
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
                        .AppendLine("        Dim data As New DataAccess(NewLeafSettings.GetConnectionString, DataAccess.DataProvider.SQL)")
                        .AppendLine("        Dim dict As New System.Collections.Specialized.HybridDictionary")
                        .AppendLine("")
                        .AppendLine("        iCmd = data.GetCommand(""sp__Delete" & parsedTableName & """, CommandType.StoredProcedure, DataAccess.DataProvider.SQL)")
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

                    If Me._WriteToFile = True Then
                        Call SaveToFile(interfaceFileOutput.ToString, interfaceName & ".vb", OBJECT_TYPE.Interface)
                        Call SaveToFile(classFileOutput.ToString, parsedTableName & ".vb", OBJECT_TYPE.Class)
                    End If
                Next

                Debug.Write(totalOutput.ToString)

            Catch ex As Exception
                Throw ex
            End Try

            Return True

        End Function



        Public Function BuildClassAndInterfaceFilePreviewer(ByVal TableList As SortedList) As String
            Try

                Dim previewer As StringBuilder = New StringBuilder(String.Empty)
                previewer.AppendLine("The following classes and interfaces will need to be created:")

                Dim totalOutput As StringBuilder = New StringBuilder(String.Empty)
                Dim columnNames As ArrayList = New ArrayList

                Dim sqlOutput As StringBuilder = New StringBuilder(String.Empty)
                Dim sqlColumnList As SortedList = New SortedList

                Dim dt As DataTable = Me.GetTables

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

                    Dim parsedTableName As String = tableName.Replace("tbl", "")
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

                        Dim columnName As String = Services.GetNULLableString(row("COLUMN_NAME"))
                        Dim columnPosition As Integer = Services.GetNULLableInteger(row("ORDINAL_POSITION"))
                        Dim maxCharacterLength As String = Services.GetNULLableString(row("CHARACTER_MAXIMUM_LENGTH"))
                        Dim numericPrecision As String = Services.GetNULLableString(row("NUMERIC_PRECISION"))
                        Dim numericScale As String = Services.GetNULLableString(row("NUMERIC_SCALE"))
                        Dim dataType As String = Services.GetNULLableString(row("DATA_TYPE")).ToLower
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

                            If columnPosition > 1 Then
                                'Exclude the primary key in this case because we do this in our Save() method by checking if the current variable for the class is "0" or not.
                                saveParams.AppendLine("        " & Me.BuildDbParam(columnName, columnPosition, privateMember, maxCharacterLength, numericPrecision, numericScale, VbDataType, METHOD_TYPE.INSERT))

                                fillSets.AppendLine("                Me." & privateMember & " = " & Me.GetNullableType(VbDataType, "row(""" & columnName & """)"))

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


                            classPropertyType.AppendLine("    Property " & parsedPropertyClass & "() As " & currentPropertyClassInterface & " Implements " & interfaceName & "." & parsedPropertyClass)
                            classPropertyType.AppendLine("        Get")
                            classPropertyType.AppendLine("            Return " & parsedPropertyClassAsVariable)
                            classPropertyType.AppendLine("        End Get")
                            classPropertyType.AppendLine("        Set(ByVal value As " & currentPropertyClassInterface & ")")
                            classPropertyType.AppendLine("            value = " & parsedPropertyClassAsVariable)
                            classPropertyType.AppendLine("        End Set")
                            classPropertyType.AppendLine("    End Property")
                        Else

                            classPropertyType.AppendLine(interfacePropertyType & " Implements " & interfaceName & "." & columnName)
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

                        .AppendLine("    Public Sub Fill() Implements " & interfaceName & ".Fill")
                        .AppendLine("")
                        .AppendLine("        Try")
                        .AppendLine("            Dim dt As DataTable = Me.GetList()")
                        .AppendLine("            For Each row As DataRow In dt.Rows()")
                        .Append(fillSets.ToString)
                        .AppendLine("            Next")
                        .AppendLine("        Catch ex As Exception")
                        .AppendLine("            Throw New NLTException(""Error retrieving and filling [" & parsedTableName.ToUpper & "] information"", ex, """ & parsedTableName & ".vb"", ""Sub Fill()"")")
                        .AppendLine("        End Try")
                        .AppendLine("")
                        .AppendLine("    End Sub")
                        .AppendLine("")
                        .AppendLine("    Public Function GetList() as System.Data.DataTable Implements " & interfaceName & ".GetList")
                        .AppendLine("")
                        .AppendLine("        Try")
                        .AppendLine("            Dim connString = NewLeafSettings.GetConnectionString()")
                        .AppendLine("            Dim data As New DataAccess(connString, DataAccess.DataProvider.SQL)")
                        .AppendLine("            Dim iCmd As IDbCommand = data.GetCommand(""sp__Get" & parsedTableName & "ById"", CommandType.StoredProcedure, DataAccess.DataProvider.SQL)")
                        .AppendLine("            " & Me.BuildDbParam(parsedTableName & "ID", 1, privateMemberVariableFromTableName, "4", "", "", "Integer", METHOD_TYPE.UPDATE, False))
                        .AppendLine("")
                        .AppendLine("            With iCmd.Parameters")
                        .AppendLine("               .Add(iParm" & parsedTableName & "ID)")
                        .AppendLine("            End With")
                        .AppendLine("")
                        .AppendLine("            Return data.GetDataTable(iCmd)")
                        .AppendLine("        Catch ex As Exception")
                        .AppendLine("            Throw New NLTException(""Error retrieving [" & parsedTableName.ToUpper & "] information"", ex, """ & parsedTableName & ".vb"", ""Function GetList()"")")
                        .AppendLine("        End Try")
                        .AppendLine("")
                        .AppendLine("    End Function")
                        .AppendLine("")
                        .AppendLine("    Public Function Save() as Boolean Implements " & interfaceName & ".Save")
                        .AppendLine("")
                        .AppendLine("        Dim success As Boolean = False")
                        .AppendLine("        Dim iCmd As IDbCommand = Nothing")
                        .AppendLine("        Dim iParm" & parsedTableName & "ID As IDbDataParameter = Nothing")
                        .AppendLine("        Dim data As New DataAccess(NewLeafSettings.GetConnectionString, DataAccess.DataProvider.SQL)")
                        .AppendLine("        Dim dict As New System.Collections.Specialized.HybridDictionary")
                        .AppendLine("        Dim storedProc As String = String.Empty")
                        .AppendLine("")
                        .AppendLine("        If Me." & privateMemberVariableFromTableName & " = 0 Then")
                        .AppendLine("            storedProc = ""sp__Add" & parsedTableName & "")
                        .AppendLine("            " & Me.BuildDbParam(parsedTableName & "ID", 1, privateMemberVariableFromTableName, "4", "", "", "Integer", METHOD_TYPE.INSERT, True))
                        .AppendLine("        Else")
                        .AppendLine("            storedProc = ""sp__Update" & parsedTableName & "")
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
                        .AppendLine("    Public Function Delete() as Boolean Implements " & interfaceName & ".Delete")
                        .AppendLine("")
                        .AppendLine("        Dim success As Boolean = False")
                        .AppendLine("        Dim iCmd As IDbCommand = Nothing")
                        .AppendLine("        Dim data As New DataAccess(NewLeafSettings.GetConnectionString, DataAccess.DataProvider.SQL)")
                        .AppendLine("        Dim dict As New System.Collections.Specialized.HybridDictionary")
                        .AppendLine("")
                        .AppendLine("        iCmd = data.GetCommand(""sp__Delete" & parsedTableName & """, CommandType.StoredProcedure, DataAccess.DataProvider.SQL)")
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

                totalOutput.Insert(0, previewer)

                Debug.Write(totalOutput.ToString)


                Dim container As StringBuilder = New StringBuilder(String.Empty)
                Dim output As String = totalOutput.ToString
                output = output.Replace(" ", "&nbsp;")
                output = output.Replace(vbCrLf, "<br/>")

                Return output

            Catch ex As Exception
                Throw ex
            End Try
        End Function

#End Region

    End Class

End Namespace

#Region " STORED PROCEDURE CODE GENERATOR SCRIPTS "

'    /****** Object:  StoredProcedure [dbo].[sp_NLT_MakeDeleteRecordProc]    Script Date: 09/22/2008 08:37:01 ******/
'    IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[sp_NLT_MakeDeleteRecordProc]') AND type in (N'P', N'PC'))
'    DROP PROCEDURE [dbo].[sp_NLT_MakeDeleteRecordProc]
'    GO
'    /****** Object:  StoredProcedure [dbo].[sp_NLT_MakeSelectRecordProc]    Script Date: 09/22/2008 08:37:01 ******/
'    IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[sp_NLT_MakeSelectRecordProc]') AND type in (N'P', N'PC'))
'    DROP PROCEDURE [dbo].[sp_NLT_MakeSelectRecordProc]
'    GO
'    /****** Object:  StoredProcedure [dbo].[sp_NLT_MakeUpdateRecordProc]    Script Date: 09/22/2008 08:37:01 ******/
'    IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[sp_NLT_MakeUpdateRecordProc]') AND type in (N'P', N'PC'))
'    DROP PROCEDURE [dbo].[sp_NLT_MakeUpdateRecordProc]
'    GO
'    /****** Object:  UserDefinedFunction [dbo].[fn__TableColumnInfo]    Script Date: 09/22/2008 08:37:01 ******/
'    IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[fn__TableColumnInfo]') AND type in (N'FN', N'IF', N'TF', N'FS', N'FT'))
'    DROP FUNCTION [dbo].[fn__TableColumnInfo]
'    GO
'    /****** Object:  StoredProcedure [dbo].[sp_NLT_MakeInsertRecordProc]    Script Date: 09/22/2008 08:37:01 ******/
'    IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[sp_NLT_MakeInsertRecordProc]') AND type in (N'P', N'PC'))
'    DROP PROCEDURE [dbo].[sp_NLT_MakeInsertRecordProc]
'    GO
'    /****** Object:  StoredProcedure [dbo].[sp_NLT_GenerateSqlProcs]    Script Date: 09/22/2008 08:37:01 ******/
'    IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[sp_NLT_GenerateSqlProcs]') AND type in (N'P', N'PC'))
'    DROP PROCEDURE [dbo].[sp_NLT_GenerateSqlProcs]
'    GO
'    /****** Object:  UserDefinedFunction [dbo].[fn__TableHasPrimaryKey]    Script Date: 09/22/2008 08:37:01 ******/
'    IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[fn__TableHasPrimaryKey]') AND type in (N'FN', N'IF', N'TF', N'FS', N'FT'))
'    DROP FUNCTION [dbo].[fn__TableHasPrimaryKey]
'    GO
'    /****** Object:  UserDefinedFunction [dbo].[fn__ColumnDefault]    Script Date: 09/22/2008 08:37:01 ******/
'    IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[fn__ColumnDefault]') AND type in (N'FN', N'IF', N'TF', N'FS', N'FT'))
'    DROP FUNCTION [dbo].[fn__ColumnDefault]
'    GO
'    /****** Object:  UserDefinedFunction [dbo].[fn__IsColumnPrimaryKey]    Script Date: 09/22/2008 08:37:01 ******/
'    IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[fn__IsColumnPrimaryKey]') AND type in (N'FN', N'IF', N'TF', N'FS', N'FT'))
'    DROP FUNCTION [dbo].[fn__IsColumnPrimaryKey]
'    GO
'    /****** Object:  UserDefinedFunction [dbo].[fn__CleanDefaultValue]    Script Date: 09/22/2008 08:37:01 ******/
'    IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[fn__CleanDefaultValue]') AND type in (N'FN', N'IF', N'TF', N'FS', N'FT'))
'    DROP FUNCTION [dbo].[fn__CleanDefaultValue]
'    GO
'    /****** Object:  UserDefinedFunction [dbo].[fn__CleanDefaultValue]    Script Date: 09/22/2008 08:37:01 ******/
'    SET ANSI_NULLS ON
'    GO
'    SET QUOTED_IDENTIFIER ON
'    GO
'    IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[fn__CleanDefaultValue]') AND type in (N'FN', N'IF', N'TF', N'FS', N'FT'))
'    BEGIN
'    execute dbo.sp_executesql @statement = N'
'    CREATE FUNCTION [dbo].[fn__CleanDefaultValue](@sDefaultValue varchar(4000))
'    RETURNS varchar(4000)
'    AS
'    BEGIN
'    	RETURN SubString(@sDefaultValue, 2, DataLength(@sDefaultValue)-2)
'    END



'' 
'    END
'    GO
'    /****** Object:  UserDefinedFunction [dbo].[fn__IsColumnPrimaryKey]    Script Date: 09/22/2008 08:37:01 ******/
'    SET ANSI_NULLS ON
'    GO
'    SET QUOTED_IDENTIFIER ON
'    GO
'    IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[fn__IsColumnPrimaryKey]') AND type in (N'FN', N'IF', N'TF', N'FS', N'FT'))
'    BEGIN
'    execute dbo.sp_executesql @statement = N'



'    CREATE FUNCTION [dbo].[fn__IsColumnPrimaryKey](@sTableName varchar(128), @nColumnName varchar(128))
'    RETURNS bit
'    AS
'    BEGIN
'    	DECLARE @nTableID int,
'    		@nIndexID int,
'    		@i int

'    	SET 	@nTableID = OBJECT_ID(@sTableName)

'    	SELECT 	@nIndexID = indid
'    	FROM 	sysindexes
'    	WHERE 	id = @nTableID
'    	 AND 	indid BETWEEN 1 And 254 
'    	 AND 	(status & 2048) = 2048

'    	IF @nIndexID Is Null
'    		RETURN 0

'    	IF @nColumnName IN
'    		(SELECT sc.[name]
'    		FROM 	sysindexkeys sik
'    			INNER JOIN syscolumns sc ON sik.id = sc.id AND sik.colid = sc.colid
'    		WHERE 	sik.id = @nTableID
'    		 AND 	sik.indid = @nIndexID)
'    	 BEGIN
'    		RETURN 1
'    	 END


'    	RETURN 0
'    END







'' 
'    END
'    GO
'    /****** Object:  UserDefinedFunction [dbo].[fn__TableHasPrimaryKey]    Script Date: 09/22/2008 08:37:01 ******/
'    SET ANSI_NULLS ON
'    GO
'    SET QUOTED_IDENTIFIER ON
'    GO
'    IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[fn__TableHasPrimaryKey]') AND type in (N'FN', N'IF', N'TF', N'FS', N'FT'))
'    BEGIN
'    execute dbo.sp_executesql @statement = N'
'    CREATE FUNCTION [dbo].[fn__TableHasPrimaryKey](@sTableName varchar(128))
'    RETURNS bit
'    AS
'    BEGIN
'    	DECLARE @nTableID int,
'    		@nIndexID int

'    	SET 	@nTableID = OBJECT_ID(@sTableName)

'    	SELECT 	@nIndexID = indid
'    	FROM 	sysindexes
'    	WHERE 	id = @nTableID
'    	 AND 	indid BETWEEN 1 And 254 
'    	 AND 	(status & 2048) = 2048

'    	IF @nIndexID IS NOT Null
'    		RETURN 1

'    	RETURN 0
'    END


'' 
'    END
'    GO
'    /****** Object:  StoredProcedure [dbo].[sp_NLT_GenerateSqlProcs]    Script Date: 09/22/2008 08:37:01 ******/
'    SET ANSI_NULLS ON
'    GO
'    SET QUOTED_IDENTIFIER ON
'    GO
'    IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[sp_NLT_GenerateSqlProcs]') AND type in (N'P', N'PC'))
'    BEGIN
'    EXEC dbo.sp_executesql @statement = N'

'    CREATE PROC [dbo].[sp_NLT_GenerateSqlProcs]
'    	@bExecute bit = 0
'    as
'    begin

'    	DECLARE @tableName sysname,
'    		@sql varchar(4000)

'    	DECLARE nlt_cursor CURSOR FOR

'    	select table_name 
'    	from information_schema.tables 
'    	where table_name like ''tbl%''

'    	OPEN nlt_cursor

'    	FETCH NEXT FROM nlt_cursor
'    	INTO @tableName

'    	WHILE @@FETCH_STATUS = 0
'    	BEGIN
'    		select @sql = ''EXECUTE sp_NLT_MakeSelectRecordProc @sTableName = '''''' + @tableName + '''''', @bExecute = '' + convert(char(1),@bExecute) + char(13) + char(10)
'    		select @sql = @sql + ''EXECUTE sp_NLT_MakeInsertRecordProc @sTableName = '''''' + @tableName + '''''', @bExecute = '' + convert(char(1),@bExecute) + char(13) + char(10)
'    		select @sql = @sql + ''EXECUTE sp_NLT_MakeUpdateRecordProc @sTableName = '''''' + @tableName + '''''', @bExecute = '' + convert(char(1),@bExecute) + char(13) + char(10)
'    		select @sql = @sql + ''EXECUTE sp_NLT_MakeDeleteRecordProc @sTableName = '''''' + @tableName + '''''', @bExecute = '' + convert(char(1),@bExecute) + char(13) + char(10)
'    		select @sql = @sql + char(13) + char(10)

'    		print @sql

'    		if (@bExecute = 1)
'    		begin
'    			exec(@sql)
'    		end

'    	   FETCH NEXT FROM nlt_cursor
'    	   INTO @tableName
'    	END

'    	CLOSE nlt_cursor
'    	DEALLOCATE nlt_cursor

'    end










'' 
'    END
'    GO
'    /****** Object:  UserDefinedFunction [dbo].[fn__ColumnDefault]    Script Date: 09/22/2008 08:37:01 ******/
'    SET ANSI_NULLS ON
'    GO
'    SET QUOTED_IDENTIFIER ON
'    GO
'    IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[fn__ColumnDefault]') AND type in (N'FN', N'IF', N'TF', N'FS', N'FT'))
'    BEGIN
'    execute dbo.sp_executesql @statement = N'
'    CREATE FUNCTION [dbo].[fn__ColumnDefault](@sTableName varchar(128), @sColumnName varchar(128))
'    RETURNS varchar(4000)
'    AS
'    BEGIN
'    	DECLARE @sDefaultValue varchar(4000)

'    	SELECT	@sDefaultValue = dbo.fn__CleanDefaultValue(COLUMN_DEFAULT)
'    	FROM	INFORMATION_SCHEMA.COLUMNS
'    	WHERE	TABLE_NAME = @sTableName
'    	 AND 	COLUMN_NAME = @sColumnName

'    	RETURN 	@sDefaultValue

'    END

'' 
'    END
'    GO
'    /****** Object:  StoredProcedure [dbo].[sp_NLT_MakeDeleteRecordProc]    Script Date: 09/22/2008 08:37:01 ******/
'    SET ANSI_NULLS ON
'    GO
'    SET QUOTED_IDENTIFIER OFF
'    GO
'    IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[sp_NLT_MakeDeleteRecordProc]') AND type in (N'P', N'PC'))
'    BEGIN
'    EXEC dbo.sp_executesql @statement = N'/*
'    exec sp_NLT_MakeSelectRecordProc ''tblAidStation''
'    exec sp_NLT_MakeInsertRecordProc ''tblAidStation''
'    exec sp_NLT_MakeUpdateRecordProc ''tblAidStation''
'    exec sp_NLT_MakeDeleteRecordProc ''tblAidStation''
'    */
'    CREATE PROC [dbo].[sp_NLT_MakeDeleteRecordProc]
'    	@sTableName varchar(128),
'    	@bExecute bit = 0
'    AS

'    IF dbo.fn__TableHasPrimaryKey(@sTableName) = 0
'     BEGIN
'    	RAISERROR (''Procedure cannot be created on a table with no primary key.'', 10, 1)
'    	RETURN
'     END

'    DECLARE	@sProcText varchar(8000),
'    	@sKeyFields varchar(2000),
'    	@sWhereClause varchar(2000),
'    	@sColumnName varchar(128),
'    	@nColumnID smallint,
'    	@bPrimaryKeyColumn bit,
'    	@nAlternateType int,
'    	@nColumnLength int,
'    	@nColumnPrecision int,
'    	@nColumnScale int,
'    	@IsNullable bit, 
'    	@IsIdentity int,
'    	@sTypeName varchar(128),
'    	@sDefaultValue varchar(4000),
'    	@sCRLF char(2),
'    	@sTAB char(1)

'    SET	@sTAB = char(9)
'    SET 	@sCRLF = char(13) + char(10)

'    SET 	@sProcText = ''''
'    SET 	@sKeyFields = ''''
'    SET	@sWhereClause = ''''

'    SET 	@sProcText = @sProcText + ''IF EXISTS(SELECT * FROM sysobjects WHERE name = ''''sp__Delete'' + substring(@sTableName,4,len(@sTableName)-3) + '''''')'' + @sCRLF
'    SET 	@sProcText = @sProcText + @sTAB + ''DROP PROC sp__Delete'' + substring(@sTableName,4,len(@sTableName)-3) + @sCRLF
'    IF @bExecute = 0
'    	SET 	@sProcText = @sProcText + ''GO'' + @sCRLF

'    SET 	@sProcText = @sProcText + @sCRLF

'    PRINT @sProcText

'    IF @bExecute = 1 
'    	EXEC (@sProcText)

'    SET 	@sProcText = ''''
'    SET 	@sProcText = @sProcText + ''----------------------------------------------------------------------------'' + @sCRLF
'    SET 	@sProcText = @sProcText + ''-- Delete a single record from '' + @sTableName + @sCRLF
'    SET 	@sProcText = @sProcText + ''----------------------------------------------------------------------------'' + @sCRLF
'    SET 	@sProcText = @sProcText + ''CREATE PROC sp__Delete'' + substring(@sTableName,4,len(@sTableName)-3) + @sCRLF

'    DECLARE crKeyFields cursor for
'    	SELECT	*
'    	FROM	dbo.fn__TableColumnInfo(@sTableName)
'    	ORDER BY 2

'    OPEN crKeyFields

'    FETCH 	NEXT 
'    FROM 	crKeyFields 
'    INTO 	@sColumnName, @nColumnID, @bPrimaryKeyColumn, @nAlternateType, 
'    	@nColumnLength, @nColumnPrecision, @nColumnScale, @IsNullable, 
'    	@IsIdentity, @sTypeName, @sDefaultValue

'    WHILE (@@FETCH_STATUS = 0)
'     BEGIN

'    	IF (@bPrimaryKeyColumn = 1)
'    	 BEGIN
'    		IF (@sKeyFields <> '''')
'    			SET @sKeyFields = @sKeyFields + '','' + @sCRLF 

'    		SET @sKeyFields = @sKeyFields + @sTAB + ''@'' + @sColumnName + '' '' + @sTypeName

'    		IF (@nAlternateType = 2) --decimal, numeric
'    			SET @sKeyFields =  @sKeyFields + ''('' + CAST(@nColumnPrecision AS varchar(3)) + '', '' 
'    					+ CAST(@nColumnScale AS varchar(3)) + '')''

'    		ELSE IF (@nAlternateType = 1) --character and binary
'    			SET @sKeyFields =  @sKeyFields + ''('' + CAST(@nColumnLength AS varchar(4)) +  '')''

'    		IF (@sWhereClause = '''')
'    			SET @sWhereClause = @sWhereClause + ''WHERE '' 
'    		ELSE
'    			SET @sWhereClause = @sWhereClause + '' AND '' 

'    		SET @sWhereClause = @sWhereClause + @sTAB + @sColumnName  + '' = @'' + @sColumnName + @sCRLF 
'    	 END

'    	FETCH 	NEXT 
'    	FROM 	crKeyFields 
'    	INTO 	@sColumnName, @nColumnID, @bPrimaryKeyColumn, @nAlternateType, 
'    		@nColumnLength, @nColumnPrecision, @nColumnScale, @IsNullable, 
'    		@IsIdentity, @sTypeName, @sDefaultValue
'     END

'    CLOSE crKeyFields
'    DEALLOCATE crKeyFields

'    SET 	@sProcText = @sProcText + @sKeyFields + @sCRLF
'    SET 	@sProcText = @sProcText + ''AS'' + @sCRLF
'    SET 	@sProcText = @sProcText + @sCRLF
'    SET 	@sProcText = @sProcText + ''DELETE	'' + @sTableName + @sCRLF
'    SET 	@sProcText = @sProcText + @sWhereClause
'    SET 	@sProcText = @sProcText + @sCRLF
'    IF @bExecute = 0
'    	SET 	@sProcText = @sProcText + ''GO'' + @sCRLF


'    PRINT @sProcText

'    IF @bExecute = 1 
'    	EXEC (@sProcText)









'' 
'    END
'    GO
'    /****** Object:  StoredProcedure [dbo].[sp_NLT_MakeSelectRecordProc]    Script Date: 09/22/2008 08:37:01 ******/
'    SET ANSI_NULLS ON
'    GO
'    SET QUOTED_IDENTIFIER OFF
'    GO
'    IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[sp_NLT_MakeSelectRecordProc]') AND type in (N'P', N'PC'))
'    BEGIN
'    EXEC dbo.sp_executesql @statement = N'



'    CREATE PROC [dbo].[sp_NLT_MakeSelectRecordProc]
'    	@sTableName varchar(128),
'    	@bExecute bit = 0
'    AS

'    IF dbo.fn__TableHasPrimaryKey(@sTableName) = 0
'     BEGIN
'    	RAISERROR (''Procedure cannot be created on a table with no primary key.'', 10, 1)
'    	RETURN
'     END

'    DECLARE	@sProcText varchar(8000),
'    	@sKeyFields varchar(2000),
'    	@sSelectClause varchar(2000),
'    	@sWhereClause varchar(2000),
'    	@sColumnName varchar(128),
'    	@nColumnID smallint,
'    	@bPrimaryKeyColumn bit,
'    	@nAlternateType int,
'    	@nColumnLength int,
'    	@nColumnPrecision int,
'    	@nColumnScale int,
'    	@IsNullable bit, 
'    	@IsIdentity int,
'    	@sTypeName varchar(128),
'    	@sDefaultValue varchar(4000),
'    	@sCRLF char(2),
'    	@sTAB char(1)

'    SET	@sTAB = char(9)
'    SET 	@sCRLF = char(13) + char(10)

'    SET 	@sProcText = ''''
'    SET 	@sKeyFields = ''''
'    SET	@sSelectClause = ''''
'    SET	@sWhereClause = ''''

'    SET 	@sProcText = @sProcText + ''IF EXISTS(SELECT * FROM sysobjects WHERE name = ''''sp__Get'' + substring(@sTableName,4,len(@sTableName)-3) + ''ById'''')'' + @sCRLF
'    SET 	@sProcText = @sProcText + @sTAB + ''DROP PROC sp__Get'' + substring(@sTableName,4,len(@sTableName)-3) + ''ById'' + @sCRLF
'    IF @bExecute = 0
'    	SET 	@sProcText = @sProcText + ''GO'' + @sCRLF

'    SET 	@sProcText = @sProcText + @sCRLF

'    PRINT @sProcText

'    IF @bExecute = 1 
'    	EXEC (@sProcText)

'    SET 	@sProcText = ''''
'    SET 	@sProcText = @sProcText + ''----------------------------------------------------------------------------'' + @sCRLF
'    SET 	@sProcText = @sProcText + ''-- Select a single record from '' + @sTableName + @sCRLF
'    SET 	@sProcText = @sProcText + ''----------------------------------------------------------------------------'' + @sCRLF
'    SET 	@sProcText = @sProcText + ''CREATE PROC sp__Get'' + substring(@sTableName,4,len(@sTableName)-3) + ''ById'' + @sCRLF

'    DECLARE crKeyFields cursor for
'    	SELECT	*
'    	FROM	dbo.fn__TableColumnInfo(@sTableName)
'    	ORDER BY 2

'    OPEN crKeyFields

'    FETCH 	NEXT 
'    FROM 	crKeyFields 
'    INTO 	@sColumnName, @nColumnID, @bPrimaryKeyColumn, @nAlternateType, 
'    	@nColumnLength, @nColumnPrecision, @nColumnScale, @IsNullable, 
'    	@IsIdentity, @sTypeName, @sDefaultValue

'    WHILE (@@FETCH_STATUS = 0)
'     BEGIN
'    	IF (@bPrimaryKeyColumn = 1)
'    	 BEGIN
'    		IF (@sKeyFields <> '''')
'    			SET @sKeyFields = @sKeyFields + '','' + @sCRLF 

'    		SET @sKeyFields = @sKeyFields + @sTAB + ''@'' + @sColumnName + '' '' + @sTypeName

'    		IF (@nAlternateType = 2) --decimal, numeric
'    			SET @sKeyFields =  @sKeyFields + ''('' + CAST(@nColumnPrecision AS varchar(3)) + '', '' 
'    					+ CAST(@nColumnScale AS varchar(3)) + '')''

'    		ELSE IF (@nAlternateType = 1) --character and binary
'    			SET @sKeyFields =  @sKeyFields + ''('' + CAST(@nColumnLength AS varchar(4)) +  '')''

'    		IF (@sWhereClause = '''')
'    			SET @sWhereClause = @sWhereClause + ''WHERE '' 
'    		ELSE
'    			SET @sWhereClause = @sWhereClause + '' AND '' 

'    		SET @sWhereClause = @sWhereClause + @sTAB + @sColumnName  + '' = @'' + @sColumnName + @sCRLF 
'    	 END

'    	IF (@sSelectClause = '''')
'    		SET @sSelectClause = @sSelectClause + ''SELECT''
'    	ELSE
'    		SET @sSelectClause = @sSelectClause + '','' + @sCRLF 

'    	SET @sSelectClause = @sSelectClause + @sTAB + @sColumnName 

'    	FETCH 	NEXT 
'    	FROM 	crKeyFields 
'    	INTO 	@sColumnName, @nColumnID, @bPrimaryKeyColumn, @nAlternateType, 
'    		@nColumnLength, @nColumnPrecision, @nColumnScale, @IsNullable, 
'    		@IsIdentity, @sTypeName, @sDefaultValue
'     END

'    CLOSE crKeyFields
'    DEALLOCATE crKeyFields

'    SET 	@sSelectClause = @sSelectClause + @sCRLF

'    SET 	@sProcText = @sProcText + @sKeyFields + @sCRLF
'    SET 	@sProcText = @sProcText + ''AS'' + @sCRLF
'    SET 	@sProcText = @sProcText + @sCRLF
'    SET 	@sProcText = @sProcText + @sSelectClause
'    SET 	@sProcText = @sProcText + ''FROM	'' + @sTableName + @sCRLF
'    SET 	@sProcText = @sProcText + @sWhereClause
'    SET 	@sProcText = @sProcText + @sCRLF
'    IF @bExecute = 0
'    	SET 	@sProcText = @sProcText + ''GO'' + @sCRLF


'    PRINT @sProcText

'    IF @bExecute = 1 
'    	EXEC (@sProcText)









'' 
'    END
'    GO
'    /****** Object:  StoredProcedure [dbo].[sp_NLT_MakeUpdateRecordProc]    Script Date: 09/22/2008 08:37:01 ******/
'    SET ANSI_NULLS ON
'    GO
'    SET QUOTED_IDENTIFIER ON
'    GO
'    IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[sp_NLT_MakeUpdateRecordProc]') AND type in (N'P', N'PC'))
'    BEGIN
'    EXEC dbo.sp_executesql @statement = N'








'    CREATE PROC [dbo].[sp_NLT_MakeUpdateRecordProc]
'    	@sTableName varchar(128),
'    	@bExecute bit = 0
'    AS

'    IF dbo.fn__TableHasPrimaryKey(@sTableName) = 0
'     BEGIN
'    	RAISERROR (''Procedure cannot be created on a table with no primary key.'', 10, 1)
'    	RETURN
'     END

'    DECLARE	@sProcText varchar(8000),
'    	@sKeyFields varchar(2000),
'    	@sSetClause varchar(2000),
'    	@sWhereClause varchar(2000),
'    	@sColumnName varchar(128),
'    	@nColumnID smallint,
'    	@bPrimaryKeyColumn bit,
'    	@nAlternateType int,
'    	@nColumnLength int,
'    	@nColumnPrecision int,
'    	@nColumnScale int,
'    	@IsNullable bit, 
'    	@IsIdentity int,
'    	@sTypeName varchar(128),
'    	@sDefaultValue varchar(4000),
'    	@sCRLF char(2),
'    	@sTAB char(1)

'    SET	@sTAB = char(9)
'    SET 	@sCRLF = char(13) + char(10)

'    SET 	@sProcText = ''''
'    SET 	@sKeyFields = ''''
'    SET	@sSetClause = ''''
'    SET	@sWhereClause = ''''

'    SET 	@sProcText = @sProcText + ''IF EXISTS(SELECT * FROM sysobjects WHERE name = ''''sp__Update'' + substring(@sTableName,4,len(@sTableName)-3) + '''''')'' + @sCRLF
'    SET 	@sProcText = @sProcText + @sTAB + ''DROP PROC sp__Update'' + substring(@sTableName,4,len(@sTableName)-3) + @sCRLF
'    IF @bExecute = 0
'    	SET 	@sProcText = @sProcText + ''GO'' + @sCRLF

'    SET 	@sProcText = @sProcText + @sCRLF

'    PRINT @sProcText

'    IF @bExecute = 1 
'    	EXEC (@sProcText)

'    SET 	@sProcText = ''''
'    SET 	@sProcText = @sProcText + ''----------------------------------------------------------------------------'' + @sCRLF
'    SET 	@sProcText = @sProcText + ''-- Update a single record in '' + @sTableName + @sCRLF
'    SET 	@sProcText = @sProcText + ''----------------------------------------------------------------------------'' + @sCRLF
'    SET 	@sProcText = @sProcText + ''CREATE PROC sp__Update'' + substring(@sTableName,4,len(@sTableName)-3) + @sCRLF

'    DECLARE crKeyFields cursor for
'    	SELECT	*
'    	FROM	dbo.fn__TableColumnInfo(@sTableName)
'    	ORDER BY 2

'    OPEN crKeyFields


'    FETCH 	NEXT 
'    FROM 	crKeyFields 
'    INTO 	@sColumnName, @nColumnID, @bPrimaryKeyColumn, @nAlternateType, 
'    	@nColumnLength, @nColumnPrecision, @nColumnScale, @IsNullable, 
'    	@IsIdentity, @sTypeName, @sDefaultValue

'    WHILE (@@FETCH_STATUS = 0)
'     BEGIN
'    	IF (@sKeyFields <> '''')
'    		SET @sKeyFields = @sKeyFields + '','' + @sCRLF 

'    	SET @sKeyFields = @sKeyFields + @sTAB + ''@'' + @sColumnName + '' '' + @sTypeName

'    	IF (@nAlternateType = 2) --decimal, numeric
'    		SET @sKeyFields =  @sKeyFields + ''('' + CAST(@nColumnPrecision AS varchar(3)) + '', '' 
'    				+ CAST(@nColumnScale AS varchar(3)) + '')''

'    	ELSE IF (@nAlternateType = 1) --character and binary
'    		SET @sKeyFields =  @sKeyFields + ''('' + CAST(@nColumnLength AS varchar(4)) +  '')''

'    	IF (@bPrimaryKeyColumn = 1)
'    	 BEGIN
'    		IF (@sWhereClause = '''')
'    			SET @sWhereClause = @sWhereClause + ''WHERE '' 
'    		ELSE
'    			SET @sWhereClause = @sWhereClause + '' AND '' 

'    		SET @sWhereClause = @sWhereClause + @sTAB + @sColumnName  + '' = @'' + @sColumnName + @sCRLF 
'    	 END
'    	ELSE
'    		IF (@IsIdentity = 0)
'    		 BEGIN
'    			IF (@sSetClause = '''')
'    				SET @sSetClause = @sSetClause + ''SET''
'    			ELSE
'    				SET @sSetClause = @sSetClause + '','' + @sCRLF 
'    			SET @sSetClause = @sSetClause + @sTAB + @sColumnName  + '' = ''
'    			IF (@sTypeName = ''timestamp'')
'    				SET @sSetClause = @sSetClause + ''NULL''
'    			ELSE IF (@sDefaultValue IS NOT NULL)
'    				SET @sSetClause = @sSetClause + ''COALESCE(@'' + @sColumnName + '', '' + @sDefaultValue + '')''
'    			ELSE
'    				SET @sSetClause = @sSetClause + ''@'' + @sColumnName 
'    		 END

'    	IF (@IsIdentity = 0)
'    	 BEGIN
'    		IF (@IsNullable = 1) OR (@sTypeName = ''timestamp'')
'    			SET @sKeyFields = @sKeyFields + '' = NULL''
'    	 END

'    	FETCH 	NEXT 
'    	FROM 	crKeyFields 
'    	INTO 	@sColumnName, @nColumnID, @bPrimaryKeyColumn, @nAlternateType, 
'    		@nColumnLength, @nColumnPrecision, @nColumnScale, @IsNullable, 
'    		@IsIdentity, @sTypeName, @sDefaultValue
'     END

'    CLOSE crKeyFields
'    DEALLOCATE crKeyFields

'    SET 	@sSetClause = @sSetClause + @sCRLF

'    SET 	@sProcText = @sProcText + @sKeyFields + @sCRLF
'    SET 	@sProcText = @sProcText + ''AS'' + @sCRLF
'    SET 	@sProcText = @sProcText + @sCRLF
'    SET 	@sProcText = @sProcText + ''UPDATE	'' + @sTableName + @sCRLF
'    SET 	@sProcText = @sProcText + @sSetClause
'    SET 	@sProcText = @sProcText + @sWhereClause
'    SET 	@sProcText = @sProcText + @sCRLF
'    IF @bExecute = 0
'    	SET 	@sProcText = @sProcText + ''GO'' + @sCRLF


'    PRINT @sProcText

'    IF @bExecute = 1 
'    	EXEC (@sProcText)











'' 
'    END
'    GO
'    /****** Object:  StoredProcedure [dbo].[sp_NLT_MakeInsertRecordProc]    Script Date: 09/22/2008 08:37:01 ******/
'    SET ANSI_NULLS ON
'    GO
'    SET QUOTED_IDENTIFIER ON
'    GO
'    IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[sp_NLT_MakeInsertRecordProc]') AND type in (N'P', N'PC'))
'    BEGIN
'    EXEC dbo.sp_executesql @statement = N'

'    CREATE PROC [dbo].[sp_NLT_MakeInsertRecordProc]
'    	@sTableName varchar(128),
'    	@bExecute bit = 0
'    AS

'    IF dbo.fn__TableHasPrimaryKey(@sTableName) = 0
'     BEGIN
'    	RAISERROR (''Procedure cannot be created on a table with no primary key.'', 10, 1)
'    	RETURN
'     END

'    DECLARE	@sProcText varchar(8000),
'    	@sKeyFields varchar(2000),
'    	@sAllFields varchar(2000),
'    	@sAllParams varchar(2000),
'    	@sWhereClause varchar(2000),
'    	@sColumnName varchar(128),
'    	@nColumnID smallint,
'    	@bPrimaryKeyColumn bit,
'    	@nAlternateType int,
'    	@nColumnLength int,
'    	@nColumnPrecision int,
'    	@nColumnScale int,
'    	@IsNullable bit, 
'    	@IsIdentity int,
'    	@HasIdentity int,
'    	@sTypeName varchar(128),
'    	@sDefaultValue varchar(4000),
'    	@sCRLF char(2),
'    	@sTAB char(1)

'    SET 	@HasIdentity = 0
'    SET	@sTAB = char(9)
'    SET 	@sCRLF = char(13) + char(10)
'    SET 	@sProcText = ''''
'    SET 	@sKeyFields = ''''
'    SET	@sAllFields = ''''
'    SET	@sWhereClause = ''''
'    SET	@sAllParams  = ''''

'    SET 	@sProcText = @sProcText + ''IF EXISTS(SELECT * FROM sysobjects WHERE name = ''''sp__Add'' + substring(@sTableName,4,len(@sTableName)-3) + '''''')'' + @sCRLF
'    SET 	@sProcText = @sProcText + @sTAB + ''DROP PROC sp__Add'' + substring(@sTableName,4,len(@sTableName)-3) + @sCRLF
'    IF @bExecute = 0
'    	SET 	@sProcText = @sProcText + ''GO'' + @sCRLF

'    SET 	@sProcText = @sProcText + @sCRLF

'    PRINT @sProcText

'    IF @bExecute = 1 
'    	EXEC (@sProcText)

'    SET 	@sProcText = ''''
'    SET 	@sProcText = @sProcText + ''----------------------------------------------------------------------------'' + @sCRLF
'    SET 	@sProcText = @sProcText + ''-- Insert a single record into '' + @sTableName + @sCRLF
'    SET 	@sProcText = @sProcText + ''----------------------------------------------------------------------------'' + @sCRLF
'    SET 	@sProcText = @sProcText + ''CREATE PROC sp__Add'' + substring(@sTableName,4,len(@sTableName)-3) + @sCRLF

'    DECLARE crKeyFields cursor for
'    	SELECT	*
'    	FROM	dbo.fn__TableColumnInfo(@sTableName)
'    	ORDER BY 2

'    OPEN crKeyFields


'    FETCH 	NEXT 
'    FROM 	crKeyFields 
'    INTO 	@sColumnName, @nColumnID, @bPrimaryKeyColumn, @nAlternateType, 
'    	@nColumnLength, @nColumnPrecision, @nColumnScale, @IsNullable, 
'    	@IsIdentity, @sTypeName, @sDefaultValue

'    WHILE (@@FETCH_STATUS = 0)
'     BEGIN
'    	IF (@IsIdentity = 0)
'    	 BEGIN
'    		IF (@sKeyFields <> '''')
'    			SET @sKeyFields = @sKeyFields + '','' + @sCRLF 

'    		SET @sKeyFields = @sKeyFields + @sTAB + ''@'' + @sColumnName + '' '' + @sTypeName

'    		IF (@sAllFields <> '''')
'    		 BEGIN
'    			SET @sAllParams = @sAllParams + '', ''
'    			SET @sAllFields = @sAllFields + '', ''
'    		 END

'    		IF (@sTypeName = ''timestamp'')
'    			SET @sAllParams = @sAllParams + ''NULL''
'    		ELSE IF (@sDefaultValue IS NOT NULL)
'    			SET @sAllParams = @sAllParams + ''COALESCE(@'' + @sColumnName + '', '' + @sDefaultValue + '')''
'    		ELSE
'    			SET @sAllParams = @sAllParams + ''@'' + @sColumnName 

'    		SET @sAllFields = @sAllFields + @sColumnName 

'    	 END
'    	ELSE
'    	 BEGIN
'    		SET @HasIdentity = 1
'    	 END

'    	IF (@nAlternateType = 2) --decimal, numeric
'    		SET @sKeyFields =  @sKeyFields + ''('' + CAST(@nColumnPrecision AS varchar(3)) + '', '' 
'    				+ CAST(@nColumnScale AS varchar(3)) + '')''

'    	ELSE IF (@nAlternateType = 1) --character and binary
'    		SET @sKeyFields =  @sKeyFields + ''('' + CAST(@nColumnLength AS varchar(4)) +  '')''

'    	IF (@IsIdentity = 0)
'    	 BEGIN
'    		IF (@sDefaultValue IS NOT NULL) OR (@IsNullable = 1) OR (@sTypeName = ''timestamp'')
'    			SET @sKeyFields = @sKeyFields + '' = NULL''
'    	 END

'    	FETCH 	NEXT 
'    	FROM 	crKeyFields 
'    	INTO 	@sColumnName, @nColumnID, @bPrimaryKeyColumn, @nAlternateType, 
'    		@nColumnLength, @nColumnPrecision, @nColumnScale, @IsNullable, 
'    		@IsIdentity, @sTypeName, @sDefaultValue
'     END

'    CLOSE crKeyFields
'    DEALLOCATE crKeyFields

'    SET 	@sProcText = @sProcText + @sKeyFields + @sCRLF
'    SET 	@sProcText = @sProcText + ''AS'' + @sCRLF
'    SET 	@sProcText = @sProcText + @sCRLF
'    SET 	@sProcText = @sProcText + ''INSERT '' + @sTableName + ''('' + @sAllFields + '')'' + @sCRLF
'    SET 	@sProcText = @sProcText + ''VALUES ('' + @sAllParams + '')'' + @sCRLF
'    SET 	@sProcText = @sProcText + @sCRLF

'    IF (@HasIdentity = 1)
'     BEGIN
'    	SET 	@sProcText = @sProcText + ''RETURN SCOPE_IDENTITY()'' + @sCRLF
'    	SET 	@sProcText = @sProcText + @sCRLF
'     END

'    IF @bExecute = 0
'    	SET 	@sProcText = @sProcText + ''GO'' + @sCRLF


'    PRINT @sProcText

'    IF @bExecute = 1 
'    	EXEC (@sProcText)







'' 
'    END
'    GO
'    /****** Object:  UserDefinedFunction [dbo].[fn__TableColumnInfo]    Script Date: 09/22/2008 08:37:01 ******/
'    SET ANSI_NULLS ON
'    GO
'    SET QUOTED_IDENTIFIER ON
'    GO
'    IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[fn__TableColumnInfo]') AND type in (N'FN', N'IF', N'TF', N'FS', N'FT'))
'    BEGIN
'    execute dbo.sp_executesql @statement = N'






'    CREATE FUNCTION [dbo].[fn__TableColumnInfo](@sTableName varchar(128))
'    RETURNS TABLE
'    AS
'    	RETURN
'    	SELECT	c.name AS sColumnName,
'    		c.colid AS nColumnID,
'    		dbo.fn__IsColumnPrimaryKey(@sTableName, c.name) AS bPrimaryKeyColumn,
'    		CASE 	WHEN t.name IN (''char'', ''varchar'', ''binary'', ''varbinary'', ''nchar'', ''nvarchar'') THEN 1
'    			WHEN t.name IN (''decimal'', ''numeric'') THEN 2
'    			ELSE 0
'    		END AS nAlternateType,
'    		c.length AS nColumnLength,
'    		c.prec AS nColumnPrecision,
'    		c.scale AS nColumnScale, 
'    		c.IsNullable, 
'    		SIGN(c.status & 128) AS IsIdentity,
'    		t.name as sTypeName,
'    		dbo.fn__ColumnDefault(@sTableName, c.name) AS sDefaultValue
'    	FROM	syscolumns c 
'    		INNER JOIN systypes t ON c.xtype = t.xtype and c.usertype = t.usertype
'    	WHERE	c.id = OBJECT_ID(@sTableName)








'' 
'    END
'    GO

#End Region
