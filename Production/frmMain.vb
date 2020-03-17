Imports System.IO
Imports SEAN.ConfigurationStoring
Public Class frmMain
    Private productionPath As String
    Private currentIdx As Integer
    Private batchContainer As BatchInfos
    Private bolContainer As BOLInfos
    Private bolHolder As BOLInfo
    Private dateFolder As String

    Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Text = Application.ProductName & " v" & Application.ProductVersion

        bolContainer = New BOLInfos
        batchContainer = New BatchInfos
        userInfo = New clsUserInfo() 'With {.Username = "DDCUSER109", .Fullname = "FERNANDEZ, SEAN IVAN M."}
        If login() = False Then
            Close()
            Exit Sub
        End If

        If Not Directory.Exists(My.Settings.ProductionPath) Then
            MsgBox("Production Path not found.")
            sfrmSettings.ShowDialog()
        End If
        productionPath = My.Settings.ProductionPath

        tm_Tick(tm, Nothing)
        tm.Enabled = True
        tm.Interval = 1000
    End Sub

    Private Sub collectPronumbers()
        Dim pros As New List(Of BOLInfo)
        Dim tmppros As BOLInfo() = BOLInfos.collectProNumbers(userInfo.Username, dateFolder, BOLInfo.BOLStatus.FINISH)
        pros.AddRange(BOLInfos.collectProNumbers(userInfo.Username, dateFolder, BOLInfo.BOLStatus.FINISH))
        pros.AddRange(BOLInfos.collectProNumbers(userInfo.Username, dateFolder, BOLInfo.BOLStatus.QUERY))
        pros.AddRange(BOLInfos.collectProNumbers(userInfo.Username, dateFolder, BOLInfo.BOLStatus.REJECT))
        For Each pro In pros
            If bolContainer.Find(pro.ProNo, pro.FBNo) Is Nothing Then
                bolContainer.Add(pro)
                With pro
                    dgvProduction.Rows.Add(.Batch.ClientEmailDateTime, .Batch.Folder, .Batch.TripNo, .ProNo, .FBNo, .Remarks, .Status, .Entry(0).Start, .Entry(0).Endd)
                End With
            End If
        Next
    End Sub

    Private Sub collectBatches()
        dateFolder = Directory.GetDirectories(productionPath, Now.ToString("yyyyMMdd"))(0)
        Dim batches As String() = Directory.GetFiles(dateFolder & "\BATCH")
        For Each _batch In batches
            Dim binf As BatchInfo = XmlSerialization.ReadFromFile(_batch, New BatchInfo)
            With binf
                .Filename = Path.GetFileNameWithoutExtension(_batch)
                .DateFolder = dateFolder
                If batchContainer.Find(binf.TripNo) Is Nothing AndAlso binf.ForEntry > 0 Then
                    batchContainer.Add(binf)
                    dgvBatches.Rows.Add(.Filename, binf.TripNo, .BillCount, .ForEntry, .Query.Count, .Billed.Count, .Reject.Count, .Ongoing.Count, .ClientEmailDateTime, .TA, .ClientEmail)
                End If
            End With
        Next
    End Sub

    Private Sub collectFeedback()
        dgvQueryAnswer.Rows.Clear()
        Dim _answeredQuery As String = dateFolder & "\PROCESSED\QUERY\ANSWERED\"
        If Directory.Exists(_answeredQuery) Then
            Dim bolpaths As String() = Directory.GetFiles(_answeredQuery, "*" & userInfo.Username & ".XML")
            Dim bols As New List(Of BOLInfo)
            For Each bolpath In bolpaths
                Dim bol As BOLInfo = XmlSerialization.ReadFromFile(bolpath, New BOLInfo)
                For Each query As QueryInfo In bol.Query
                    dgvQueryAnswer.Rows.Add(query.Endd, bol.Batch.TripNo, bol.ProNo, bol.FBNo, query.QueryAnswer)
                Next
            Next
        End If
    End Sub

    Private Sub refreshBatches()
        For i As Integer = 0 To batchContainer.Count - 1
            Dim binf As BatchInfo = batchContainer(i)
            If binf.ForEntry = 0 Then
                batchContainer.Remove(binf)
                dgvBatches.Rows.RemoveAt(i)
            Else
                With binf
                    dgvBatches.Rows(i).SetValues(.Filename, binf.TripNo, .BillCount, .ForEntry, .Query.Count, .Billed.Count, .Reject.Count, .Ongoing.Count, .ClientEmailDateTime, .TA, .ClientEmail)
                End With
            End If
        Next
    End Sub

    Private Function validateFields() As Boolean
        If tbProNumber.Text = "" Then
            MsgBox("Pro Number should not be blank.")
            Return False
        End If

        If tbFBNumber.Text = "" Then
            MsgBox("FB Number should not be blank.")
            Return False
        End If

        If tbFolder.Text = "" Then Return False
        'If Date.TryParse(tbStarttime.Text, New Date) = False Then
        '    MsgBox("Invalid Datetime format on Start Time")
        '    Return False
        'End If

        If bolContainer.Find(tbProNumber.Text, tbFBNumber.Text) IsNot Nothing Then
            MsgBox("Duplicate Pro Number or FB Number found")
            Return False
        End If
        Return True
    End Function
    Private Sub clearFields()
        tbTripNumber.Text = ""
        tbProNumber.Text = ""
        tbFBNumber.Text = ""
        tbFolder.Text = ""

        ' tbAudittime.Text = ""
        tbStarttime.Text = ""
        tbEndtime.Text = ""

        cbRemark.Text = ""
    End Sub

    Private Sub mnSettings_Click(sender As Object, e As EventArgs) Handles mnSettings.Click
        sfrmSettings.ShowDialog()
    End Sub

    Private Sub dgv_CurrentCellChanged(sender As Object, e As EventArgs) Handles dgvBatches.CurrentCellChanged, dgvBatches.GotFocus
        If dgvBatches.CurrentCell IsNot Nothing Then
            clearFields()
            Dim _curridx As Integer = dgvBatches.CurrentCell.RowIndex
            bolHolder = New BOLInfo
            bolHolder.Batch = batchContainer.Find(dgvBatches.Rows(_curridx).Cells(1).Value)
            tbTripNumber.Text = bolHolder.Batch.TripNo
            tbFolder.Text = bolHolder.Batch.Folder
        End If
    End Sub

    Private Sub dgvProduction_CurrentCellChanged(sender As Object, e As EventArgs) Handles dgvProduction.CurrentCellChanged, dgvProduction.GotFocus
        If dgvBatches.CurrentCell IsNot Nothing Then
            clearFields()
            Dim _curridx As Integer = dgvProduction.CurrentCell.RowIndex
            bolHolder = bolContainer.Find(dgvProduction.Rows(_curridx).Cells(3).Value, dgvProduction.Rows(_curridx).Cells(4).Value)
            tbTripNumber.Text = bolHolder.Batch.TripNo
            tbFolder.Text = bolHolder.Batch.Folder
            tbProNumber.Text = bolHolder.ProNo
            tbFBNumber.Text = bolHolder.FBNo
            cbRemark.Text = bolHolder.Remarks
        End If
    End Sub
    Private Sub mnAddBill_Click(sender As Object, e As EventArgs) Handles mnAddBill.Click
        If validateFields() Then
            With bolHolder
                .ProNo = tbProNumber.Text
                .FBNo = tbFBNumber.Text
                .Remarks = cbRemark.Text

                .Username = userInfo.Username
                .Fullname = userInfo.Fullname

                .Query = New List(Of QueryInfo)
                .QueryToVALC = New List(Of QueryInfo)

                .Entry = New List(Of TimeInfo)
                Using holder As New sfrmHolder(bolHolder, New TimeInfo With {.Start = bolHolder.Batch.Time.ToString("yyyy-MM-dd HH:mm:ss")}) 'Date.Parse(tbStarttime.Text)
                    .Write(BOLInfo.BOLStatus.ONGOING)
                    If holder.ShowDialog = Windows.Forms.DialogResult.OK Then
                        .Entry.Add(holder.TimeInf)
                        If holder.Query.BillerQuery <> "" Then
                            .Status = "QUERY"
                            .Query.Add(holder.Query)
                            .Write(BOLInfo.BOLStatus.QUERY)
                        ElseIf holder.QueryForVALC.BillerQuery <> "" Then
                            .Status = "QUERY"
                            .QueryToVALC.Add(holder.QueryForVALC)
                            .Write(BOLInfo.BOLStatus.QUERY)
                        Else
                            If .Status = "ANSWERED" Then
                                .Status = "QUERY_BILLED"
                            Else : .Status = "BILLED"
                            End If
                            .Write(BOLInfo.BOLStatus.FINISH)
                        End If
                    End If
                    .Delete(BOLInfo.BOLStatus.ONGOING)
                    .Delete(BOLInfo.BOLStatus.ANSWERED)
                    clearFields()
                End Using
            End With
        End If
    End Sub

    Private Sub tm_Tick(sender As Object, e As EventArgs) Handles tm.Tick
        tm.Enabled = False
        If Not bgRefresher.IsBusy Then bgRefresher.RunWorkerAsync()
    End Sub

    Private Sub bgRefresher_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles bgRefresher.DoWork
        Try
            dgvBatches.Invoke(Sub()
                                  collectBatches()
                                  refreshBatches()
                              End Sub)
            dgvQueryAnswer.Invoke(Sub() collectFeedback())
            dgvProduction.Invoke(Sub() collectPronumbers())
        Catch ex As Exception
            ' MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub bgRefresher_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles bgRefresher.RunWorkerCompleted
        tm.Enabled = True
    End Sub

    Private userInfo As clsUserInfo
    Private Sub mnRelogin_Click(sender As Object, e As EventArgs) Handles mnRelogin.Click
        login()
    End Sub
    Private Function login() As Boolean
        Using userLogin As New sfrmLogin(userInfo)
            If userLogin.ShowDialog() = Windows.Forms.DialogResult.OK Then
                userInfo = userLogin.User
                lbFullname.Text = userInfo.Fullname
                lbUsername.Text = userInfo.Username
                Return True
            Else : Return False
            End If
        End Using
    End Function

    Private Sub mnSaveBill_Click(sender As Object, e As EventArgs) Handles mnSaveBill.Click

    End Sub

    Private Sub mnSaveAsQuery_Click(sender As Object, e As EventArgs) Handles mnSaveAsQuery.Click

    End Sub

End Class
