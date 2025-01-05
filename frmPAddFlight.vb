Public Class frmPAddFlight
    Dim blnBoolean As Boolean
    Private Sub frmAddCustomerFlight_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            Dim strSelect As String = ""
            Dim cmdSelect As OleDb.OleDbCommand
            Dim drSourceTable As OleDb.OleDbDataReader
            Dim dt As DataTable = New DataTable

            If OpenDatabaseConnectionSQLServer() = False Then

                MessageBox.Show(Me, "Database connection error." & vbNewLine &
                                "The application will now close.",
                                Me.Text + " Error",
                                MessageBoxButtons.OK, MessageBoxIcon.Error)
                Me.Close()

            End If
            strSelect = "SELECT TF.intFlightID, TF.strFlightNumber as Flight From TFlights as TF"

            cmdSelect = New OleDb.OleDbCommand(strSelect, m_conAdministrator)
            drSourceTable = cmdSelect.ExecuteReader
            dt.Load(drSourceTable)

            cboFlights.ValueMember = "intFlightID"
            cboFlights.DisplayMember = "Flight"
            cboFlights.DataSource = dt

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
        Dim strSeat As String
        Dim dblFlightCost As Double
        intSelectedFlight = cboFlights.SelectedValue
        strSeat = "1A"
        dblFlightCost = Calculate_FlightCost()

        If OpenDatabaseConnectionSQLServer() = False Then

            MessageBox.Show(Me, "Database connection error." & vbNewLine &
                           "The application will now close.",
                           Me.Text + " Error",
                           MessageBoxButtons.OK, MessageBoxIcon.Error)

            Me.Close()
        End If
        strSelect = "SELECT MAX(intFlightPassengerID) + 1 AS intNextPrimaryKey" & " FROM TFlightPassengers"


        cmdSelect = New OleDb.OleDbCommand(strSelect, m_conAdministrator)
        drSourceTable = cmdSelect.ExecuteReader
        drSourceTable.Read()
        If drSourceTable.IsDBNull(0) = True Then
            intNextPrimaryKey = 1
        Else
            intNextPrimaryKey = CInt(drSourceTable("intNextPrimaryKey"))
        End If
        strInsert = "INSERT INTO TFlightPassengers (intFlightPassengerID, intFlightID, intPassengerID, strSeat, dblFlightCost)
        VALUES (" & intNextPrimaryKey & ", '" & intSelectedFlight & "', '" & strString & "', '" & strSeat & "', '" & dblFlightCost & "')"

        cmdInsert = New OleDb.OleDbCommand(strInsert, m_conAdministrator)
        intRowsAffected = cmdInsert.ExecuteNonQuery()
        If intRowsAffected > 0 Then
            MessageBox.Show("Flight has been added")
        End If
        CloseDatabaseConnection()
        Close()
        Dim frmPAddFlight As New frmPAddFlight
        frmPAddFlight.Show()
        Exit Sub
    End Sub

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        Dim frmPassengerMenu As New frmPassengersMenu
        frmPassengerMenu.Show()
        Me.Hide()
    End Sub

    Private Sub cboFlights_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboFlights.SelectedIndexChanged
        If blnBoolean = True Then
            Dim dblReservedSeat As Double = 375
            Dim dblDesignated As Double = 250
            Dim dblSeatPrice As Double
            Dim dblDiscountPercentageTotal As Double
            Dim intMilesFlown As Integer
            Dim intTotalPassengersPerFlight As Integer
            Dim intPlaneType As Integer
            Dim intAirport As Integer
            Dim intPassengerDay As Integer
            Dim intRepeatCustomer As Integer
            Dim intPassengerAge As Integer
            Dim dblDaysInAYear As Double = 365.25
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

                'StoreProcedure: If Total Miles of flight is greater than 750 miles
                cmdSelect = New OleDb.OleDbCommand("uspMilesFlown", m_conAdministrator)
                cmdSelect.CommandType = CommandType.StoredProcedure
                objParam = cmdSelect.Parameters.Add("@intFlight_ID", OleDb.OleDbType.Integer)
                objParam.Direction = ParameterDirection.Input
                objParam.Value = cboFlights.SelectedValue
                drSourceTable = cmdSelect.ExecuteReader
                While drSourceTable.Read()
                    intMilesFlown = drSourceTable("intMilesFlown")
                End While
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

                'StoredProcedure: # of passengers on flight
                cmdSelect = New OleDb.OleDbCommand("uspTotalPassengersPerFlight", m_conAdministrator)
                cmdSelect.CommandType = CommandType.StoredProcedure
                objParam = cmdSelect.Parameters.Add("@intFlight_ID", OleDb.OleDbType.Integer)
                objParam.Direction = ParameterDirection.Input
                objParam.Value = cboFlights.SelectedValue
                drSourceTable = cmdSelect.ExecuteReader
                While drSourceTable.Read()
                    intTotalPassengersPerFlight = drSourceTable("TotalPassengersPerFlight")
                End While
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

                'StoredProcedure: Type of Plane
                cmdSelect = New OleDb.OleDbCommand("uspTypeOfPlane", m_conAdministrator)
                cmdSelect.CommandType = CommandType.StoredProcedure
                objParam = cmdSelect.Parameters.Add("@intFlight_ID", OleDb.OleDbType.Integer)
                objParam.Direction = ParameterDirection.Input
                objParam.Value = cboFlights.SelectedValue
                drSourceTable = cmdSelect.ExecuteReader
                While drSourceTable.Read()
                    intPlaneType = drSourceTable("intPlaneTypeID")
                End While
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

                'Stored Procedure: If landing in “MIA” for this flight  
                cmdSelect = New OleDb.OleDbCommand("uspAirPlaneDestination", m_conAdministrator)
                cmdSelect.CommandType = CommandType.StoredProcedure
                objParam = cmdSelect.Parameters.Add("@intFlight_ID", OleDb.OleDbType.Integer)
                objParam.Direction = ParameterDirection.Input
                objParam.Value = cboFlights.SelectedValue
                drSourceTable = cmdSelect.ExecuteReader
                While drSourceTable.Read()
                    intAirport = drSourceTable("Destination")
                End While
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

                'StoreProcedure: Calculate how old passenger is in days (ex: 366 days = 1 years old)
                cmdSelect = New OleDb.OleDbCommand("uspPassengerDay", m_conAdministrator)
                cmdSelect.CommandType = CommandType.StoredProcedure
                objParam = cmdSelect.Parameters.Add("@intPassenger_ID", OleDb.OleDbType.Integer)
                objParam.Direction = ParameterDirection.Input
                objParam.Value = strString
                drSourceTable = cmdSelect.ExecuteReader
                While drSourceTable.Read()
                    intPassengerDay = drSourceTable("PassengerDay")
                End While
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

                'StoredProcedure: Amount of flights passenger has recorded in DB
                cmdSelect = New OleDb.OleDbCommand("uspRepeatCustomer", m_conAdministrator)
                cmdSelect.CommandType = CommandType.StoredProcedure
                objParam = cmdSelect.Parameters.Add("@intPassenger_ID", OleDb.OleDbType.Integer)
                objParam.Direction = ParameterDirection.Input
                objParam.Value = strString
                drSourceTable = cmdSelect.ExecuteReader
                While drSourceTable.Read()
                    intRepeatCustomer = drSourceTable("RepeatCustomer")
                End While
                drSourceTable.Close()
                ' close the database connection
                CloseDatabaseConnection()
            Catch ex As Exception
                ' Log and display error message
                MessageBox.Show(ex.Message)
            End Try
            intPassengerAge = intPassengerDay / dblDaysInAYear
            dblSeatPrice = Calculate_SeatPrice(intMilesFlown, intTotalPassengersPerFlight, intPlaneType, intAirport, intPassengerAge, intRepeatCustomer)
            dblReservedSeat += dblSeatPrice
            dblDesignated += dblSeatPrice
            dblDiscountPercentageTotal = Calculate_Discount(intPassengerAge, intRepeatCustomer)
            dblReservedSeat = dblReservedSeat - (dblReservedSeat * dblDiscountPercentageTotal)
            dblDesignated = dblDesignated - (dblDesignated * dblDiscountPercentageTotal)
            lblReserved.Text = dblReservedSeat.ToString("c")
            lblDesignated.Text = dblDesignated.ToString("c")
            radReservedSeat.Enabled = True
            radDesignated.Enabled = True
        Else
            blnBoolean = True
        End If
    End Sub

    Private Sub radReservedSeat_CheckedChanged(sender As Object, e As EventArgs) Handles radReservedSeat.CheckedChanged
        btnSubmit.Enabled = True
    End Sub

    Private Sub radDesignated_CheckedChanged(sender As Object, e As EventArgs) Handles radDesignated.CheckedChanged
        btnSubmit.Enabled = True
    End Sub
    Private Function Calculate_FlightCost()
        Dim dblFlightCost As Double
        If radReservedSeat.Checked = True Then
            dblFlightCost = lblReserved.Text
        Else
            dblFlightCost = lblDesignated.Text
        End If
        Return dblFlightCost
    End Function

    Private Function Calculate_SeatPrice(ByVal intMilesFlown As Integer, ByVal intTotalPassengersPerFlight As Integer, ByVal intPlaneType As Integer, ByVal intAirport As Integer, ByVal intPassengerAge As Integer, ByVal intRepeatCustomer As Integer)
        Dim dblSeatPrice As Double
        If intMilesFlown > 750 Then
            dblSeatPrice += 50
        End If
        If intTotalPassengersPerFlight > 8 Then
            dblSeatPrice += 100
        ElseIf intTotalPassengersPerFlight < 4 Then
            dblSeatPrice -= 50
        End If
        'Airbus A350 
        If intPlaneType = 1 Then
            dblSeatPrice += 35
            'Boeing 747-8
        ElseIf intPlaneType = 2 Then
            dblSeatPrice -= 25
        End If
        'Miami
        If intAirport = 2 Then
            dblSeatPrice += 15
        End If
        Return dblSeatPrice
    End Function
    Private Function Calculate_Discount(ByVal intPassengerAge As Integer, ByVal intRepeatCustomer As Integer)
        Dim dblDiscountPercentageTotal As Double
        If intPassengerAge >= 65 Then
            dblDiscountPercentageTotal += 0.2
        ElseIf intPassengerAge <= 5 Then
            dblDiscountPercentageTotal += 0.65
        End If
        If intRepeatCustomer >= 10 Then
            dblDiscountPercentageTotal += 0.2
        ElseIf intRepeatCustomer >= 5 Then
            dblDiscountPercentageTotal += 0.1
        End If
        Return dblDiscountPercentageTotal
    End Function
End Class
