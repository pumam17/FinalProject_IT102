Public Class frmAdminS
    Private Sub frmAdminS_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            Dim strSelect As String
            Dim cmdSelect As OleDb.OleDbCommand            ' this will be used for our Select statement
            Dim drSourceTable As OleDb.OleDbDataReader     ' this will be where our result set will 


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

            cmdSelect = New OleDb.OleDbCommand("uspTotalPassengersInDB", m_conAdministrator)
            cmdSelect.CommandType = CommandType.StoredProcedure
            drSourceTable = cmdSelect.ExecuteReader
            drSourceTable.Read()
            If drSourceTable.IsDBNull(0) = True Then
                lblTotalCustomers.Text = 0
            Else
                lblTotalCustomers.Text = CInt(drSourceTable("TotalPassengers"))
            End If
            '-----------------------------------------------------------

            cmdSelect = New OleDb.OleDbCommand("uspTotalFlights", m_conAdministrator)
            cmdSelect.CommandType = CommandType.StoredProcedure
            drSourceTable = cmdSelect.ExecuteReader
            drSourceTable.Read()
            If drSourceTable.IsDBNull(0) = True Then
                lblTotalFlights.Text = 0
            Else
                lblTotalFlights.Text = CInt(drSourceTable("TotalFlights"))
            End If
            '-----------------------------------------------------------
            strSelect = "SELECT SUM(TF.intMilesFlown) / Count(TFP.intPassengerID) as Average " &
                        "FROM TFlightPassengers as TFP  JOIN TFlights as TF " &
                        "ON TF.intFlightID = TFP.intFlightID"
            cmdSelect = New OleDb.OleDbCommand(strSelect, m_conAdministrator)
            drSourceTable = cmdSelect.ExecuteReader
            drSourceTable.Read()
            If drSourceTable.IsDBNull(0) = True Then
                lblAvg.Text = 0
            Else
                lblAvg.Text = CInt(drSourceTable("Average"))
            End If
            '-----------------------------------------------------------

            strSelect = "SELECT Sum(TF.intMilesFlown) as TotalMiles, TP.strFirstName, TP.strLastName " &
                        "FROM TPilots as TP JOIN TPilotFlights as TPF " &
                        "ON TP.intPilotID = TPF.intPilotID " &
                        "JOIN TFlights as TF " &
                        "ON TF.intFlightID = TPF.intFlightID " &
                        "Group By TP.strFirstName, TP.strLastName " &
                        "Order By TP.strLastName asc"
            cmdSelect = New OleDb.OleDbCommand(strSelect, m_conAdministrator)
            drSourceTable = cmdSelect.ExecuteReader
            lstPilots.Items.Add("Total Flights Per Pilot")
            lstPilots.Items.Add("  ")
            lstPilots.Items.Add("=======================================")
            While drSourceTable.Read()
                lstPilots.Items.Add("  ")
                lstPilots.Items.Add("First Name: " & vbTab & drSourceTable("strFirstName"))
                lstPilots.Items.Add("Last Name: " & vbTab & drSourceTable("strLastName"))
                lstPilots.Items.Add("Has Flown : " & vbTab & drSourceTable("TotalMiles") & " Miles")
                lstPilots.Items.Add("  ")
                lstPilots.Items.Add("=======================================")

            End While
            '-----------------------------------------------------------

            strSelect = "SELECT Sum(TF.intMilesFlown) as TotalMiles, TA.strFirstName, TA.strLastName " &
                        "FROM TAttendants as TA JOIN TAttendantFlights as TAF " &
                        "ON TA.intAttendantID = TAF.intAttendantID " &
                        "JOIN TFlights as TF " &
                        "ON TF.intFlightID = TAF.intAttendantID " &
                        "Group By TA.strFirstName, TA.strLastName " &
                        "Order By TA.strLastName asc"
            cmdSelect = New OleDb.OleDbCommand(strSelect, m_conAdministrator)
            drSourceTable = cmdSelect.ExecuteReader
            lstAttendants.Items.Add("Total Flights Per Attendant")
            lstAttendants.Items.Add("  ")
            lstAttendants.Items.Add("=======================================")
            While drSourceTable.Read()
                lstAttendants.Items.Add("  ")
                lstAttendants.Items.Add("First Name: " & vbTab & drSourceTable("strFirstName"))
                lstAttendants.Items.Add("Last Name: " & vbTab & drSourceTable("strLastName"))
                lstAttendants.Items.Add("Has Flown : " & vbTab & drSourceTable("TotalMiles") & " Miles")
                lstAttendants.Items.Add("  ")
                lstAttendants.Items.Add("=======================================")

            End While

            drSourceTable.Close()
            CloseDatabaseConnection()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        Dim frmAdminPAS As New frmAdminPAS
        frmAdminPAS.Show()
        Me.Hide()
    End Sub

End Class