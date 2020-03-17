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

        tm.Enabled = True
    End Sub

    Private Sub tm_Tick(sender As Object, e As EventArgs) Handles tm.Tick
        Dim tm As TimeSpan = (Now - Date.Parse(TimeInf.Start))
        lbElapsetime.Text = String.Format("{0}:{1}", tm.Minutes, tm.Seconds)
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        TimeInf.Endd = Now.ToString("yyyy-MM-dd HH:mm:ss")
        DialogResult = Windows.Forms.DialogResult.OK
        Close()
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        DialogResult = Windows.Forms.DialogResult.Cancel
        Close()
    End Sub

    Private Function enterQuery() As QueryInfo
        Using query As New sfrmQuery
            If query.ShowDialog() = Windows.Forms.DialogResult.OK Then
                Dim _qry As QueryInfo = New QueryInfo With {.QueryAnswer = query.tbQuery.Text, .ShipperName = query.tbShipper.Text, .ConsigneeName = query.tbConsignee.Text}
                Return _qry
            End If
        End Using
        Return Nothing
    End Function

    Private Sub btnQuery_Click(sender As Object, e As EventArgs) Handles btnQuery.Click, btnQuerytoVALC.Click
        Dim _query As QueryInfo = enterQuery()
        If _query IsNot Nothing Then
            _query.Start = BOL.Batch.Time
            If sender.Equals(btnQuery) Then
                Query = _query
            Else : QueryForVALC = _query
            End If
            TimeInf.Endd = Now.ToString("yyyy-MM-dd HH:mm:ss")
            DialogResult = Windows.Forms.DialogResult.OK
            Close()
        End If
    End Sub
End Class