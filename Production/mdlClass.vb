Imports System.IO
Imports SEAN.ConfigurationStoring
'Public Batch As New BatchInfo
Public Class BatchInfos
    Inherits List(Of BatchInfo)

    Public Overloads Function Find(tripNumber As String) As BatchInfo
        Return (From res In Me Where res.TripNo = tripNumber Select res).FirstOrDefault
    End Function
End Class
Public Class BatchInfo
    Public TripNo As String = ""
    Public ClientEmailDateTime As String = ""
    Public ClientEmail As String = ""
    Public DDCSentEmailDateTime As String = ""
    Public BillCount As String = ""
    Public Folder As String = ""
    Public Synergized As String = ""

    <Xml.Serialization.XmlIgnore> Public Filename As String
    <Xml.Serialization.XmlIgnore> Public DateFolder As String
    <Xml.Serialization.XmlIgnore> Public ReadOnly Property ForEntry As Integer
        Get
            Dim filecount As Integer = 0
            If DateFolder <> "" Then
                Dim _processed As String = DateFolder & "\PROCESSED"
                If Directory.Exists(_processed) Then
                    filecount = BillCount - (Query.Count + Billed.Count + Reject.Count + Ongoing.Count)
                End If
            End If
            Return filecount
        End Get
    End Property

    <Xml.Serialization.XmlIgnore> Public ReadOnly Property Query As List(Of String)
        Get
            Dim files As New List(Of String)
            If DateFolder <> "" Then
                Dim _query As String = DateFolder & "\PROCESSED\QUERY"
                If Directory.Exists(_query) Then
                    files.AddRange(Directory.GetFiles(_query, "*" & TripNo & "*.xml"))
                End If
            End If
            Return files
        End Get
    End Property

    <Xml.Serialization.XmlIgnore> Public ReadOnly Property Billed As List(Of String)
        Get
            Dim files As New List(Of String)
            If DateFolder <> "" Then
                Dim _finished As String = DateFolder & "\PROCESSED\FINISH"
                If Directory.Exists(_finished) Then
                    files.AddRange(Directory.GetFiles(_finished, "*" & TripNo & "*.xml"))
                End If
            End If
            Return files
        End Get
    End Property

    <Xml.Serialization.XmlIgnore> Public ReadOnly Property Reject As List(Of String)
        Get
            Dim files As New List(Of String)
            If DateFolder <> "" Then
                Dim _reject As String = DateFolder & "\PROCESSED\REJECT"
                If Directory.Exists(_reject) Then
                    files.AddRange(Directory.GetFiles(_reject, "*" & TripNo & "*.xml"))
                End If
            End If
            Return files
        End Get
    End Property

    <Xml.Serialization.XmlIgnore> Public ReadOnly Property Ongoing As List(Of String)
        Get
            Dim files As New List(Of String)
            If DateFolder <> "" Then
                Dim _ongoing As String = DateFolder & "\ONGOING"
                If Directory.Exists(_ongoing) Then
                    files.AddRange(Directory.GetFiles(_ongoing, "*" & TripNo & "*.xml"))
                End If
            End If
            Return files
        End Get
    End Property

    <Xml.Serialization.XmlIgnore> Public ReadOnly Property TA As String
        Get
            Dim tm As TimeSpan = (Now - Date.Parse(ClientEmailDateTime))
            Return String.Format("{0:00}:{1:00}", tm.Minutes, tm.Seconds)
        End Get
    End Property

    <Xml.Serialization.XmlIgnore> Public ReadOnly Property Time As Date
        Get
            Dim _time As String = DateFolder & "\TIME"
            If Directory.Exists(_time) Then
                Dim times As String() = Directory.GetFiles(_time)
                If times.Length > 0 Then
                    Return Date.Parse(Path.GetFileNameWithoutExtension(times(0)).Replace("#", ":"))
                End If
            End If
            Return Now
        End Get
    End Property

    Public Overrides Function ToString() As String
        Return TripNo
    End Function
End Class

Public Class BOLInfos
    Inherits List(Of BOLInfo)

    Public Shared Function collectProNumbers(username As String, dateFolder As String, status As BOLInfo.BOLStatus) As BOLInfo()
        Dim fld As String = ""
        Select Case status
            Case BOLInfo.BOLStatus.ONGOING
                fld = "\ONGOING"
            Case BOLInfo.BOLStatus.QUERY
                fld = "\PROCESSED\QUERY"
            Case BOLInfo.BOLStatus.FINISH
                fld = "\PROCESSED\FINISH"
        End Select
        Try
            Dim pros As String() = Directory.GetFiles(dateFolder & fld, "*" & username & ".XML")
            Return (From res In pros Select DirectCast(XmlSerialization.ReadFromFile(res, New BOLInfo), BOLInfo)).ToArray
        Catch ex As Exception

        End Try
      
    End Function

    Public Overloads Function Find(proNo As String, FBNo As String) As BOLInfo
        Return (From res In Me Where res.FBNo = FBNo And res.ProNo = proNo Select res).FirstOrDefault
    End Function
End Class
Public Class BOLInfo
    Public Batch As New BatchInfo
    Public Entry As New List(Of TimeInfo)
    Public Query As New List(Of QueryInfo)
    Public ProNo As String = ""
    Public FBNo As String = ""
    Public QueryToVALC As New List(Of QueryInfo)
    Public Username As String = ""
    Public Fullname As String = ""
    Public Status As String = "" 'Billed,Reject,Query,Answered,Query_Billed
    Public Remarks As String = ""

    Public Sub Write(status As BOLStatus)
        Dim fld As String = ""
        Select Case status
            Case BOLStatus.ONGOING
                fld = "\ONGOING"
            Case BOLStatus.QUERY
                fld = "\PROCESSED\QUERY"
            Case BOLStatus.FINISH
                fld = "\PROCESSED\FINISH"
        End Select

        XmlSerialization.WriteToFile(String.Format("{0}\{1}\{2}_{3}_{4}_{5}.XML", Batch.DateFolder, fld, Batch.Filename, ProNo, FBNo, Username), Me)
    End Sub

    Public Sub Delete(status As BOLStatus)
        Dim fld As String = ""
        Select Case status
            Case BOLStatus.ONGOING
                fld = "\ONGOING"
            Case BOLStatus.QUERY
                fld = "\PROCESSED\QUERY"
            Case BOLStatus.FINISH
                fld = "\PROCESSED\FINISH"
            Case BOLStatus.ANSWERED
                fld = "\PROCESSED\QUERY\ANSWERED"
        End Select

        File.Delete(String.Format("{0}\{1}\{2}_{3}_{4}_{5}.XML", Batch.DateFolder, fld, Batch.Filename, ProNo, FBNo, Username))
    End Sub

    Public Enum BOLStatus
        ONGOING = 0
        QUERY = 1
        FINISH = 2
        ANSWERED = 3
        REJECT = 4
    End Enum
End Class



Public Class TimeInfo
    Public Start As String = ""
    Public Endd As String = ""

    Public ReadOnly Property TimeElapsed As String
        Get
            Dim tm As TimeSpan = (Date.Parse(Endd) - Date.Parse(Start))
            Return String.Format("{0}:{1}", tm.Minutes, tm.Seconds)
        End Get
    End Property
End Class
Public Class QueryInfo
    Public Start As String = ""
    Public Endd As String = ""
    Public BillerQuery As String = ""
    Public QueryAnswer As String = ""
    Public ShipperName As String = ""
    Public ConsigneeName As String = ""
End Class
Public Class clsUserInfo
    Public Fullname As String
    Public Username As String
End Class

