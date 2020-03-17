Public Class sfrmLogin
    Public User As clsUserInfo
    Sub New(_user As clsUserInfo)
        InitializeComponent()

        User = New clsUserInfo
        tbFullname.Text = _user.Fullname
        tbUsername.Text = _user.Username
    End Sub
    Private Sub btnLogin_Click(sender As Object, e As EventArgs) Handles btnLogin.Click
        User.Fullname = tbFullname.Text
        User.Username = tbUsername.Text
        DialogResult = Windows.Forms.DialogResult.OK
        Close()
    End Sub

    Private Sub tbFullname_KeyDown(sender As Object, e As KeyEventArgs) Handles tbFullname.KeyDown
        If e.KeyCode = Keys.Enter Then
            tbUsername.Focus()
        End If
    End Sub
    Private Sub tbUsername_KeyDown(sender As Object, e As KeyEventArgs) Handles tbUsername.KeyDown
        If e.KeyCode = Keys.Enter Then
            btnLogin.PerformClick()
        End If
    End Sub
End Class