Public Class frmAFFlights
    Private Sub frmAFlights_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            Dim cmdSelect As OleDb.OleDbCommand        ' this will be used for our Select statement
            Dim drSourceTable As OleDb.OleDbDataReader ' this will be where our result set will 
            Dim objParam As OleDb.OleDbParameter

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

            ' Retrieve all the records 
            cmdSelect = New OleDb.OleDbCommand("uspAttendantFutureFlights", m_conAdministrator)
            cmdSelect.CommandType = CommandType.StoredProcedure
            objParam = cmdSelect.Parameters.Add("@intAttendant_ID", OleDb.OleDbType.Integer)
            objParam.Direction = ParameterDirection.Input
            objParam.Value = strString
            drSourceTable = cmdSelect.ExecuteReader
            'loop through result set and display in Listbox
            lstResultSet.Items.Add("Future flights for selected attendant")
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
        Dim frmAttendantsMenu As New frmAttendantsMenu
        frmAttendantsMenu.Show()
        Me.Hide()
    End Sub

End Class