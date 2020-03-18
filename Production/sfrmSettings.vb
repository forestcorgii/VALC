Public Class sfrmSettings

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        My.Settings.ProductionPath = SEAN.General.ConvertToUNCPath(tbProductionPath.Text)
        My.Settings.Save()

        MsgBox("saved.")
        DialogResult = Windows.Forms.DialogResult.OK
        Close()
    End Sub

    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        tbProductionPath.Text = SEAN.General.ConvertToUNCPath(My.Settings.ProductionPath)
    End Sub

    Private Sub btnFindpath_Click(sender As Object, e As EventArgs) Handles btnFindpath.Click
        FBD.SelectedPath = My.Settings.ProductionPath
        If FBD.ShowDialog = Windows.Forms.DialogResult.OK Then
            tbProductionPath.Text = SEAN.General.ConvertToUNCPath(FBD.SelectedPath)
        End If
    End Sub

    Private Sub tbProductionPath_TextChanged(sender As Object, e As EventArgs) Handles tbProductionPath.TextChanged
        tbProductionPath.Text = SEAN.General.ConvertToUNCPath(tbProductionPath.Text)
    End Sub
End Class