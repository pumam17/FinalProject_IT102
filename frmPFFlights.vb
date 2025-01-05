Public Class frmPFFlights
    Private Sub frmPFFlights_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            Dim strSelect As String = ""
            Dim cmdSelect As OleDb.OleDbCommand        ' this will be used for our Select statement
            Dim drSourceTable As OleDb.OleDbDataReader ' this will be where our result set will
            ' open the DB
            If OpenDatabaseConnectionSQLServer() = False Then
                ' No, warn the user ...
                MessageBox.Show(Me, "Database connection error." & vbNewLine &
                                    "The application will now close.",
                                    Me.Text + " Error",
                                    MessageBoxButtons.OK, MessageBoxIcon.Error)
                ' and close the form/application
                Me.Close()
            End If
            ' Build the select statement 
            strSelect = "Select TF.intFlightID ,TF.dtmFlightDate, TF.strFlightNumber, TF.dtmTimeofDeparture, TF.dtmTimeOfLanding, TFAP.strAirportCity AS FromAirport, TTAP.strAirportCity AS ToAirport, TF.intMilesFlown, TPL.strPlaneNumber
                         From 
                         TFlights	    as TF JOIN TFlightPassengers  as TFP
                         ON TF.intFlightID = TFP.intFlightID
                         JOIN TPassengers as TP
	                     ON TP.intPassengerID = TFP.intPassengerID
                         JOIN TAirports as TFAP
                         ON TFAP.intAirportID = TF.intFromAirportID
                         JOIN TAirports as TTAP
                         ON TTAP.intAirportID = TF.intToAirportID
                         JOIN TPlanes as TPL
                         ON TPL.intPlaneID = TF.intPlaneID
                         WHERE TP.intPassengerID = " & strString & "and TF.dtmFlightDate > GETDATE() "
            cmdSelect = New OleDb.OleDbCommand(strSelect, m_conAdministrator)
            drSourceTable = cmdSelect.ExecuteReader
            'loop through result set and display in Listbox
            lstResultSet.Items.Add("Past flights for selected passenger")
            lstResultSet.Items.Add("=============================")
            While drSourceTable.Read()
                lstResultSet.Items.Add("  ")
                lstResultSet.Items.Add("Flight Date: " & vbTab & drSourceTable("dtmFlightDate"))
                lstResultSet.Items.Add("Flight Number: " & vbTab & drSourceTable("strFlightNumber"))
                lstResultSet.Items.Add("Time of Departure: " & vbTab & drSourceTable("dtmTimeofDeparture"))
                lstResultSet.Items.Add("TimeOfLanding: " & vbTab & drSourceTable("dtmTimeOfLanding"))
                lstResultSet.Items.Add("From Airport: " & vbTab & drSourceTable("FromAirport"))
                lstResultSet.Items.Add("To Airport " & vbTab & drSourceTable("ToAirport"))
                lstResultSet.Items.Add("Plane ID Number: " & vbTab & drSourceTable("strPlaneNumber"))
                lstResultSet.Items.Add("  ")
                lstResultSet.Items.Add("=============================")
            End While
            ' Clean up
            drSourceTable.Close()
            ' close the database connection
            CloseDatabaseConnection()
        Catch ex As Exception
            ' Log and display error message
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        Dim frmPassengerMenu As New frmPassengersMenu
        frmPassengerMenu.Show()
        Me.Hide()
    End Sub
End Class