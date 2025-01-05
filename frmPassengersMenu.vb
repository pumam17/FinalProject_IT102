Public Class frmPassengersMenu
    Private Sub btnUpdatePassenger_Click(sender As Object, e As EventArgs) Handles btnUpdatePassenger.Click
        Dim frmPassengerUpdate As New frmPassengerUpdate
        frmPassengerUpdate.Show()
        Me.Hide()
    End Sub

    Private Sub btnAddFlight_Click(sender As Object, e As EventArgs) Handles btnAddFlight.Click
        Dim frmPAddFlight As New frmPAddFlight
        frmPAddFlight.Show()
        Me.Hide()
    End Sub

    Private Sub btnPassengerPastFlights_Click(sender As Object, e As EventArgs) Handles btnPassengerPastFlights.Click
        Dim frmPPFlights As New frmPPFlights
        frmPPFlights.Show()
        Me.Hide()
    End Sub

    Private Sub btnPassengerFutureFlights_Click(sender As Object, e As EventArgs) Handles btnPassengerFutureFlights.Click
        Dim frmPFFlights As New frmPFFlights
        frmPFFlights.Show()
        Me.Hide()
    End Sub

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        Dim frmPLogin As New frmPLogin
        frmPLogin.Show()
        Me.Hide()
    End Sub

End Class