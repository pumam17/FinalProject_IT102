Public Class frmAdminP
    Private Sub btnAddPilot_Click(sender As Object, e As EventArgs) Handles btnAddPilot.Click
        Dim frmPilotsAdd As New frmPilotsAdd
        frmPilotsAdd.Show()
        Me.Hide()
    End Sub

    Private Sub btnDeletePilot_Click(sender As Object, e As EventArgs) Handles btnDeletePilot.Click
        Dim frmPilotsDelete As New frmPilotsDelete
        frmPilotsDelete.Show()
        Me.Hide()
    End Sub

    Private Sub btnAddPilotFlight_Click(sender As Object, e As EventArgs) Handles btnAddPilotFlight.Click
        Dim frmPilotsAddFlight As New frmPilotsAddFlight
        frmPilotsAddFlight.Show()
        Me.Hide()
    End Sub

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        Dim frmAdminPAS As New frmAdminPAS
        frmAdminPAS.Show()
        Me.Hide()
    End Sub
End Class