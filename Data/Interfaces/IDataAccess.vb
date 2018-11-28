Imports System.Data

Namespace Data

    Public Interface IDataAccess
        Function GetConnection(ByVal connectionString As String, ByVal dataProvider As DataAccess.DataProvider) As IDbConnection
        Function GetTransaction(ByVal dataProvider As DataAccess.DataProvider) As IDbTransaction
        Function GetDataAdapter(ByVal query As String, ByVal dataProvider As DataAccess.DataProvider) As IDbDataAdapter
        Function GetCommand(ByVal query As String, ByVal commandType As CommandType, ByVal dataProvider As DataAccess.DataProvider) As IDbCommand
        Function GetDataReader(ByVal query As String, ByVal commandType As CommandType, ByVal dataProvider As DataAccess.DataProvider) As IDataReader
        Function GetDataList(ByVal query As String) As System.ComponentModel.IListSource
        Function GetDataList(ByVal iCmd As IDbCommand) As System.ComponentModel.IListSource
        Function ExecuteNonQuery(ByVal commandText As String, ByVal commandType As CommandType) As Boolean
        Function ExecuteNonQuery(ByVal dict As IDictionary) As Boolean
        Sub CloseConnection()
        Sub OpenConnection()
    End Interface

End Namespace