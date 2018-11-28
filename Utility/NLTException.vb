Public Class NLTException : Inherits System.Exception

#Region " ENUMERATIONS "
    Public Enum CCMTaskType
        DB_SAVE
        DB_DELETE
        DB_SELECT
        OBJ_CREATE
        MISC_ERROR
        PAGE_EVENT
    End Enum

#End Region

#Region " PRIVATE MEMBERS "
    Private _sourceFileName As String = ""
    Private _sourceMethodName As String = ""
    Private _userID As String = ""
#End Region

#Region " PROPERTIES "
    Property SourceFileName() As String
        Get
            Return _sourceFileName
        End Get
        Set(ByVal Value As String)
            _sourceFileName = Value
        End Set
    End Property
    Property SourceMethodName() As String
        Get
            Return _sourceMethodName
        End Get
        Set(ByVal Value As String)
            _sourceMethodName = Value
        End Set
    End Property
    Property UserID() As String
        Get
            Return _userID
        End Get
        Set(ByVal Value As String)
            _userID = Value.ToUpper
        End Set
    End Property
#End Region

#Region " CONSTRUCTORS "
    Private Sub New()
        MyBase.New()
    End Sub

    Public Sub New(ByVal message As String, ByVal inner As Exception, _
               ByVal sourceFileName As String, _
               ByVal sourceMethodName As String)
        MyBase.New(message, inner)
        _sourceFileName = sourceFileName
        _sourceMethodName = sourceMethodName
    End Sub
#End Region

#Region " METHODS "

#End Region

End Class
