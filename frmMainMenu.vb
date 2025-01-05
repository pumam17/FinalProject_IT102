Public Class frmMainMenu
    Private Sub btnPassengers_Click(sender As Object, e As EventArgs) Handles btnPassengers.Click
        Dim frmPLogin As New frmPLogin
        frmPLogin.Show()
        Me.Hide()
    End Sub

    Private Sub btnEmployees_Click(sender As Object, e As EventArgs) Handles btnEmployees.Click
        Dim frmELogin As New frmELogin
        frmELogin.Show()
        Me.Hide()
    End Sub

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        Close()
    End Sub
End Class