Public Class sfrmHolder
    Public BOL As BOLInfo
    Public TimeInf As TimeInfo
    Public Query As New QueryInfo
    Public QueryForVALC As New QueryInfo
    Sub New(_bol As BOLInfo, _timeinf As TimeInfo)

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        TimeInf = _timeinf
        BOL = _bol
        With _bol
            lbUsername.Text = .Username
            lbFullname.Text = .Fullname

            lbProNumber.Text = .ProNo
            lbFBNumber.Text = .FBNo

            lbStarttime.Text = _timeinf.Start
            lbElapsetime.Text = _timeinf.Endd

            lbRemark.Text = .Remarks
        End With

        tmRefresher.Enabled = True
    End Sub

    Private Sub tm_Tick(sender As Object, e As EventArgs) Handles tmRefresher.Tick
        Me.Invoke(Sub() tmRefresher.Enabled = False)
        Dim tm As TimeSpan = (BOL.Batch.Time - Date.Parse(TimeInf.Start))
        lbElapsetime.Text = String.Format("{0:00}:{1:00}:{2:00}", tm.Hours, tm.Minutes, tm.Seconds)

        BOL.Batch.CheckTimeManagerActivity(5)
        Me.Invoke(Sub() tmRefresher.Enabled = True)
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        If MsgBox("Are you sure you want to save?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
            TimeInf.Endd = BOL.Batch.Time.ToString("yyyy-MM-dd HH:mm:ss")
            DialogResult = Windows.Forms.DialogResult.OK
            Close()
        End If
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        If MsgBox("Are you sure you want to cancel?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
            DialogResult = Windows.Forms.DialogResult.Cancel
            Close()
        End If
    End Sub

    Private Function enterQuery() As QueryInfo
        Using query As New sfrmQuery
            If query.ShowDialog() = Windows.Forms.DialogResult.OK Then
                Dim _qry As QueryInfo = New QueryInfo With {.BillerQuery = query.tbQuery.Text, .ShipperName = query.tbShipper.Text, .ConsigneeName = query.tbConsignee.Text}
                Return _qry
            End If
        End Using
        Return Nothing
    End Function

    Private Sub btnQuery_Click(sender As Object, e As EventArgs) Handles btnQuery.Click, btnQuerytoVALC.Click
        If MsgBox("Are you sure you want to save as query?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
            Dim _query As QueryInfo = enterQuery()
            If _query IsNot Nothing Then
                _query.Start = BOL.Batch.Time
                If sender.Equals(btnQuery) Then
                    Query = _query
                Else : QueryForVALC = _query
                End If
                TimeInf.Endd = BOL.Batch.Time.ToString("yyyy-MM-dd HH:mm:ss")
                DialogResult = Windows.Forms.DialogResult.OK
                Close()
            End If
        End If
    End Sub
End Class