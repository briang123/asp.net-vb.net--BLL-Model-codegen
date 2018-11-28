Imports Microsoft.VisualBasic
Imports System.Configuration
Imports System.Data

Public Class NewLeafSettings

#Region " DATABASE "

    Public Shared Function GetConnectionString() As String
        Return ConfigurationManager.ConnectionStrings("NLURConnectionString").ConnectionString
    End Function
	
#End Region

#Region " PASSWORD INFORMATION "

    Public Shared ReadOnly Property MaxLoginAttempts As Integer
        Get
            Return 5
        End Get
    End Property

    Public Shared ReadOnly Property MinRequiredPasswordLength As Integer
        Get
            Return 8
        End Get
    End Property

    Public Shared ReadOnly Property AcceptablePasswordSymbols As String
        Get
            Return "!@#$%^&*()"
        End Get
    End Property

    Public Shared ReadOnly Property MinRequiredPasswordCapitalLetters As Integer
        Get
            Return 1
        End Get
    End Property

    Public Shared ReadOnly Property MinRequiredPasswordLowerLetters As Integer
        Get
            Return 1
        End Get
    End Property

    Public Shared ReadOnly Property MinRequiredPasswordSymbols As Integer
        Get
            Return 1
        End Get
    End Property

    Public Shared ReadOnly Property MinRequiredPasswordNumbers As Integer
        Get
            Return 1
        End Get
    End Property

    Public Shared ReadOnly Property DefaultSecurityQuestion As String()
        Get
            Dim list As String() = New String() {"What is your favorite ice cream flavor?", _
                                                 "What is your favorite book?", _
                                                 "What is your favorite word?", _
                                                 "Where do you enjoy vacationing most?", _
                                                 "What is your favorite dessert?", _
                                                 "What was the name of your first pet?", _
                                                 "What is your favorite beverage?", _
                                                 "What is your favorite food?", _
                                                 "What is your favorite exercise?"}
            Return list
        End Get
    End Property

#End Region

#Region " ENUMERATIONS "

    Public Enum MESSAGE_TYPE
        INFO
        SUCCESS
        [ERROR]
    End Enum

#End Region

#Region " LISTS "



    Public Shared Function States() As SortedList(Of String, String)

        Dim list As New SortedList(Of String, String)

        list.Add("", "")
        list.Add("AL", "Alabama")
        list.Add("AK", "Alaska")
        list.Add("AZ", "Arizona")
        list.Add("AR", "Arkansas")
        list.Add("CA", "California")
        list.Add("CO", "Colorado")
        list.Add("CT", "Connecticut")
        list.Add("DE", "Delaware")
        list.Add("FL", "Florida")
        list.Add("GA", "Georgia")
        list.Add("HI", "Hawaii")
        list.Add("ID", "Idaho")
        list.Add("IL", "Illinois")
        list.Add("IN", "Indiana")
        list.Add("IA", "Iowa")
        list.Add("KS", "Kansas")
        list.Add("KY", "Kentucky")
        list.Add("LA", "Louisiana")
        list.Add("ME", "Maine")
        list.Add("MD", "Maryland")
        list.Add("MA", "Massachusetts")
        list.Add("MI", "Michigan")
        list.Add("MN", "Minnesota")
        list.Add("MS", "Mississippi")
        list.Add("MO", "Missouri")
        list.Add("MT", "Montana")
        list.Add("NE", "Nebraska")
        list.Add("NV", "Nevada")
        list.Add("NH", "New Hampshire")
        list.Add("NJ", "New Jersey")
        list.Add("NM", "New Mexico")
        list.Add("NY", "New York")
        list.Add("NC", "North Carolina")
        list.Add("ND", "North Dakota")
        list.Add("OH", "Ohio")
        list.Add("OK", "Oklahoma")
        list.Add("OR", "Oregon")
        list.Add("PA", "Pennsylvania")
        list.Add("RI", "Rhode Island")
        list.Add("SC", "South Carolina")
        list.Add("SD", "South Dakota")
        list.Add("TN", "Tennessee")
        list.Add("TX", "Texas")
        list.Add("UT", "Utah")
        list.Add("VT", "Vermont")
        list.Add("VA", "Virginia")
        list.Add("WA", "Washington")
        list.Add("DC", "Washington DC")
        list.Add("WV", "West Virginia")
        list.Add("WI", "Wisconsin")
        list.Add("WY", "Wyoming")

        Return list

    End Function

    Public Shared Function Genders() As SortedList(Of String, String)

        Dim list As New SortedList(Of String, String)

        list.Add("", "")
        list.Add("F", "Female")
        list.Add("M", "Male")

        Return list

    End Function

    Public Shared Function YesNoMaybe() As SortedList(Of String, String)

        Dim list As New SortedList(Of String, String)

        list.Add("", "")
        list.Add("Y", "YES")
        list.Add("M", "MAYBE")
        list.Add("N", "NO")

        Return list

    End Function

    Public Shared Function YesNo() As SortedList(Of String, String)

        Dim list As New SortedList(Of String, String)

        list.Add("", "")
        list.Add("Y", "YES")
        list.Add("N", "NO")

        Return list

    End Function

    Public Shared Function MembershipStatus() As IList(Of String)

        Dim list As New List(Of String)

        list.Add("")
        list.Add("PENDING APPROVAL")    'user requesting to join
        list.Add("APPROVED")            'approved to be part of group
        list.Add("DENIED")              'user denied membership
        list.Add("EXPIRED")             'user membership has expired
        list.Add("DELETED")             'user deleted their membership
        list.Add("BANNED")              'user has been banned from group

        Return list

    End Function

    Public Shared Function AdType() As IList(Of String)

        Dim list As New List(Of String)

        list.Add("")
        list.Add("AFFILIATE")
        list.Add("PRODUCT SPONSOR")
        list.Add("FINANCIAL SPONSOR")

        Return list

    End Function

    Public Shared Function TShirtSizes() As SortedList(Of String, String)

        Dim list As New SortedList(Of String, String)

        list.Add("", "")
        list.Add("XXS", "XX-SMALL")
        list.Add("XS", "X-SMALL")
        list.Add("S", "SMALL")
        list.Add("M", "MEDIUM")
        list.Add("L", "LARGE")
        list.Add("XL", "X-LARGE")
        list.Add("XXL", "XX-LARGE")

        Return list

    End Function

    Public Shared Function ContactMessageType() As IList(Of String)

        Dim list As New List(Of String)

        list.Add("")
        list.Add("GENERAL INFORMATION")
        list.Add("TECHNICAL SUPPORT")

        Return list

    End Function

    Public Shared Function ChicagolandRegion() As IList(Of String)

        Dim list As New List(Of String)

        list.Add("")
        list.Add("CHICAGO")
        list.Add("NORTH SUBURB")
        list.Add("NORTHWEST SUBURB")
        list.Add("NEAR WEST SUBURB")
        list.Add("WEST SUBURB")
        list.Add("SOUTH/SOUTHWEST SUBURB")
        list.Add("OTHER")

        Return list

    End Function

    Public Shared Function EventMessageSendList() As IList(Of String)

        Dim list As New List(Of String)

        list.Add("")
        list.Add("ATTENDING")
        list.Add("MAYBE ATTENDING")
        list.Add("NOT ATTENDING")
        list.Add("NOT RESPONDED")
        list.Add("EVERYONE")

        Return list

    End Function

    Public Shared Function MessageTemplateList() As SortedList(Of String, String)

        Dim list As New SortedList(Of String, String)

        list.Add("", "")
        list.Add("RSVP Reminder", "A reminder to RSVP in case you are on the fence or have not yet had time to do so.")
        list.Add("Event Reminder", "This is a reminder that [EVENT_NAME] is on [EVENT_DATE] at [EVENT_TIME]. You are receiving this email because you confirmed your attendance at this event. If you have changed your mind, please update your RSVP status on the website.")
        list.Add("Review (DRAFT) Results", "Please review the results and let me know if there are any necessary changes.")
        list.Add("Results Finalized", "The results from this event have been finalized and posted.")
        list.Add("Thank you", "Thank you to everyone who participated and/or volunteered at the event.")

        Return list

    End Function

#End Region

#Region " FUNCTIONS "

    Public Shared Function GetKeyValue(ByVal Prefix As String, ByVal Key As String) As DataTable

        If Not ConfigurationManager.AppSettings(Prefix.ToUpper & "_" & Key.ToUpper) Is Nothing Then
            Dim keyValues As String = ConfigurationManager.AppSettings(Prefix.ToUpper & "_" & Key.ToUpper)
            If Not keyValues.Equals(String.Empty) Then
                Dim allKeyValues As String() = keyValues.Split(";")
                Dim keyStringValue As String = String.Empty

                Dim dt As New DataTable
                dt.Columns.Add(New DataColumn("key"))
                dt.Columns.Add(New DataColumn("val"))

                For Each valueInList As String In allKeyValues
                    Dim kv As String() = valueInList.Split("|")
                    Dim dr As DataRow = dt.NewRow()
                    dr("key") = kv(0)
                    dr("val") = kv(1)
                    dt.Rows.Add(dr)
                Next
                Return dt
            Else
                Return Nothing
            End If
        Else
            Return Nothing
        End If

    End Function

#End Region

End Class
