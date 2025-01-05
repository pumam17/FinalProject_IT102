Public Class frmAdminPAS
    Private Sub btnManagePilots_Click(sender As Object, e As EventArgs) Handles btnManagePilots.Click
        Dim frmAdminP As New frmAdminP
        frmAdminP.Show()
        Me.Hide()
    End Sub

    Private Sub btnManageAttendants_Click(sender As Object, e As EventArgs) Handles btnManageAttendants.Click
        Dim frmAdminA As New frmAdminA
        frmAdminA.Show()
        Me.Hide()
    End Sub

    Private Sub btnStats_Click(sender As Object, e As EventArgs) Handles btnStats.Click
        Dim frmAdminS As New frmAdminS
        frmAdminS.Show()
        Me.Hide()
    End Sub

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        Dim frmELogin As New frmELogin
        frmELogin.Show()
        Me.Hide()
    End Sub

    Private Sub btnFutureFlights_Click(sender As Object, e As EventArgs) Handles btnFutureFlights.Click
        Dim frmFutureFlights As New frmFutureFlights
        frmFutureFlights.Show()
        Me.Hide()
    End Sub
End Class