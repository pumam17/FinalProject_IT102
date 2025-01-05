Public Class frmPilotsAddFlight
    Private Sub frmPilotAddFlight_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            Dim strSelect As String = ""
            Dim cmdSelect As OleDb.OleDbCommand
            Dim drSourceTable As OleDb.OleDbDataReader
            Dim dt As DataTable = New DataTable
            Dim dts As DataTable = New DataTable

            If OpenDatabaseConnectionSQLServer() = False Then

                MessageBox.Show(Me, "Database connection error." & vbNewLine &
                                "The application will now close.",
                                Me.Text + " Error",
                                MessageBoxButtons.OK, MessageBoxIcon.Error)
                Me.Close()

            End If
            strSelect = "SELECT intPilotID, strFirstName + ' ' + strLastName as PilotName From TPilots"

            cmdSelect = New OleDb.OleDbCommand(strSelect, m_conAdministrator)
            drSourceTable = cmdSelect.ExecuteReader
            dt.Load(drSourceTable)

            cboPilots.ValueMember = "intPilotID"
            cboPilots.DisplayMember = "PilotName"
            cboPilots.DataSource = dt
            '------------------------------------------------------------------------
            strSelect = "SELECT TF.intFlightID, TF.strFlightNumber as Flight From TFlights as TF"

            cmdSelect = New OleDb.OleDbCommand(strSelect, m_conAdministrator)
            drSourceTable = cmdSelect.ExecuteReader
            dts.Load(drSourceTable)

            cboFlights.ValueMember = "intFlightID"
            cboFlights.DisplayMember = "Flight"
            cboFlights.DataSource = dts

            strSelect = "Select TF.intFlightID ,TF.dtmFlightDate, TF.strFlightNumber, TF.dtmTimeofDeparture, TF.dtmTimeOfLanding, TFAP.strAirportCity AS FromAirport, TTAP.strAirportCity AS ToAirport, TF.intMilesFlown, TPL.strPlaneNumber 
                         From 
                         TFlights	    as TF JOIN TAirports as TFAP
                         ON TFAP.intAirportID = TF.intFromAirportID
                         JOIN TAirports as TTAP
                         ON TTAP.intAirportID = TF.intToAirportID
                         JOIN TPlanes as TPL
                         ON TPL.intPlaneID = TF.intPlaneID"

            ' Retrieve all the records 
            cmdSelect = New OleDb.OleDbCommand(strSelect, m_conAdministrator)
            drSourceTable = cmdSelect.ExecuteReader
            'loop through result set and display in Listbox
            lstResultSet.Items.Add("All flights in dbFlyMe2theMoon")
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

    Private Sub btnSubmit_Click(sender As Object, e As EventArgs) Handles btnSubmit.Click

        Dim strSelect As String
        Dim cmdSelect As OleDb.OleDbCommand
        Dim drSourceTable As OleDb.OleDbDataReader
        Dim intNextPrimaryKey As Integer
        Dim strInsert As String
        Dim cmdInsert As OleDb.OleDbCommand
        Dim intRowsAffected As Integer
        Dim intSelectedFlight As Integer
        Dim intSelectedPilot As Integer
        intSelectedFlight = cboFlights.SelectedValue
        intSelectedPilot = cboPilots.SelectedValue
        If OpenDatabaseConnectionSQLServer() = False Then

            MessageBox.Show(Me, "Database connection error." & vbNewLine &
                           "The application will now close.",
                           Me.Text + " Error",
                           MessageBoxButtons.OK, MessageBoxIcon.Error)

            Me.Close()
        End If
        strSelect = "SELECT MAX(intPilotFlightID) + 1 AS intNextPrimaryKey" & " FROM TPilotFlights"


        cmdSelect = New OleDb.OleDbCommand(strSelect, m_conAdministrator)
        drSourceTable = cmdSelect.ExecuteReader
        drSourceTable.Read()
        If drSourceTable.IsDBNull(0) = True Then
            intNextPrimaryKey = 1
        Else
            intNextPrimaryKey = CInt(drSourceTable("intNextPrimaryKey"))
        End If
        strInsert = "INSERT INTO TPilotFlights (intPilotFlightID, intPilotID, intFlightID)
        VALUES (" & intNextPrimaryKey & ", '" & intSelectedPilot & "', '" & intSelectedFlight & "')"

        cmdInsert = New OleDb.OleDbCommand(strInsert, m_conAdministrator)
        intRowsAffected = cmdInsert.ExecuteNonQuery()
        If intRowsAffected > 0 Then
            MessageBox.Show("Flight has been added")
        End If
        CloseDatabaseConnection()
        Close()
        Dim frmPilotsAddFlight As New frmPilotsAddFlight
        frmPilotsAddFlight.Show()
        Exit Sub
    End Sub

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        Dim frmAdminP As New frmAdminP
        frmAdminP.Show()
        Me.Hide()
    End Sub
End Class