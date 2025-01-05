Public Class frmFutureFlights
    Dim dtmToday As String = Today
    Private Sub frmFutureFlights_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            Dim strSelect As String = ""
            Dim cmdSelect As OleDb.OleDbCommand
            Dim drSourceTable As OleDb.OleDbDataReader
            Dim dt As DataTable = New DataTable
            Dim ds As DataTable = New DataTable
            Dim dts As DataTable = New DataTable

            If OpenDatabaseConnectionSQLServer() = False Then

                MessageBox.Show(Me, "Database connection error." & vbNewLine &
                                "The application will now close.",
                                Me.Text + " Error",
                                MessageBoxButtons.OK, MessageBoxIcon.Error)
                Me.Close()

            End If
            strSelect = "SELECT intAirportID, strAirportCity From TAirports"
            cmdSelect = New OleDb.OleDbCommand(strSelect, m_conAdministrator)
            drSourceTable = cmdSelect.ExecuteReader
            dt.Load(drSourceTable)
            cboFromAirport.ValueMember = "intAirportID"
            cboFromAirport.DisplayMember = "strAirportCity"
            cboFromAirport.DataSource = dt

            strSelect = "SELECT intAirportID, strAirportCity From TAirports"
            cmdSelect = New OleDb.OleDbCommand(strSelect, m_conAdministrator)
            drSourceTable = cmdSelect.ExecuteReader
            ds.Load(drSourceTable)
            cboToAirport.ValueMember = "intAirportID"
            cboToAirport.DisplayMember = "strAirportCity"
            cboToAirport.DataSource = ds

            strSelect = "SELECT intPlaneID, strPlaneNumber From TPlanes"
            cmdSelect = New OleDb.OleDbCommand(strSelect, m_conAdministrator)
            drSourceTable = cmdSelect.ExecuteReader
            dts.Load(drSourceTable)
            cboPlane.ValueMember = "intPlaneID"
            cboPlane.DisplayMember = "strPlaneNumber"
            cboPlane.DataSource = dts
        Catch ex As Exception
            ' Log and display error message
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub btnSubmit_Click(sender As Object, e As EventArgs) Handles btnSubmit.Click
        Dim dtmFlightDate As String
        Dim intFlightNumber As Integer
        Dim dtmTimeofDeparture As String
        Dim dtmTimeOfLanding As String
        Dim intFromAirportID As Integer
        Dim intToAirportID As Integer
        Dim intMilesFlown As Integer
        Dim intPlaneID As Integer
        Dim blnValidation As Boolean = True
        Get_And_Validate_Input(dtmFlightDate, intFlightNumber, dtmTimeofDeparture, dtmTimeOfLanding, intFromAirportID, intToAirportID, intMilesFlown, intPlaneID, blnValidation)
        If blnValidation = True Then
            Dim cmdAddFlight As New OleDb.OleDbCommand()
            Dim intPKID As Integer
            Dim intRowsAffected As Integer

            If OpenDatabaseConnectionSQLServer() = False Then

                MessageBox.Show(Me, "Database connection error." & vbNewLine &
                               "The application will now close.",
                               Me.Text + " Error",
                               MessageBoxButtons.OK, MessageBoxIcon.Error)

                Me.Close()
            End If
            cmdAddFlight.CommandText = "EXECUTE uspAddFlight " & intPKID & ", '" & dtmFlightDate & "', " & intFlightNumber & ", '" & dtmTimeofDeparture & "', '" & dtmTimeOfLanding & "', " & intFromAirportID & ", " & intToAirportID & ", " & intMilesFlown & ", " & intPlaneID
            cmdAddFlight.CommandType = CommandType.StoredProcedure
            cmdAddFlight = New OleDb.OleDbCommand(cmdAddFlight.CommandText, m_conAdministrator)
            intRowsAffected = cmdAddFlight.ExecuteNonQuery()
            If intRowsAffected > 0 Then
                MessageBox.Show("Flight has been added")
            End If
            CloseDatabaseConnection()
        Else
            Exit Sub
        End If
    End Sub

    Private Sub Get_And_Validate_Input(ByRef dtmFlightDate As String, ByRef intFlightNumber As Integer, ByRef dtmTimeofDeparture As String, ByRef dtmTimeOfLanding As String, ByRef intFromAirportID As Integer, ByRef intToAirportID As Integer, ByRef intMilesFlown As Integer, ByRef intPlaneID As Integer, ByRef blnValidation As Boolean)
        Validate_FlightDate(dtmFlightDate, blnValidation)
        Validate_FlightNumber(intFlightNumber, blnValidation)
        dtmTimeofDeparture = TimeOfDepString()
        dtmTimeOfLanding = TimeOfLandString()
        intFromAirportID = cboFromAirport.SelectedValue
        intToAirportID = cboToAirport.SelectedValue
        Validate_MilesFlown(intMilesFlown, blnValidation)
        intPlaneID = cboPlane.SelectedValue
    End Sub

    Private Sub Validate_FlightDate(ByRef dtmFlightDate As String, ByRef blnValidation As Boolean)

        Dim strString As String
        If txtFlightDate.Text = String.Empty Then
            blnValidation = ValidationDate()
        Else
            dtmFlightDate = txtFlightDate.Text
            If dtmFlightDate.Length <> 10 Then
                blnValidation = ValidationDate()
            Else
                strString = dtmFlightDate.Substring(0, 2)
                If IsNumeric(strString) Then
                    If strString < 0 Then
                        blnValidation = ValidationDate()
                    End If
                    If strString > 12 Then
                        blnValidation = ValidationDate()
                    End If
                Else
                    blnValidation = ValidationDate()
                End If
                strString = dtmFlightDate.Substring(3, 2)
                If IsNumeric(strString) Then
                    If strString < 0 Then
                        blnValidation = ValidationDate()
                    End If
                    If strString > 31 Then
                        blnValidation = ValidationDate()
                    End If
                Else
                    blnValidation = ValidationDate()
                End If
                strString = dtmFlightDate.Substring(6, 4)
                If IsNumeric(strString) Then
                    If dtmFlightDate.IndexOf("/") <> 2 And 5 Then
                        blnValidation = ValidationDate()
                    End If
                End If
            End If
            If CDate(dtmFlightDate) < CDate(dtmToday) Then
                blnValidation = ValidationDate()
            End If
        End If
    End Sub

    Private Sub Validate_FlightNumber(ByRef intFlightNumber As Integer, ByRef blnValidation As Boolean)
        If Integer.TryParse(txtFlightNumber.Text, intFlightNumber) Then
            intFlightNumber = txtFlightNumber.Text
            If (intFlightNumber < 1) Or (intFlightNumber >= 1000) Then
                MessageBox.Show("User Input Must Be Greater than 0 and Less or Equal to 999")
                txtFlightNumber.Focus()
                Exit Sub
            End If
        Else
            MessageBox.Show("User Input Must Exist and Must be Numeric")
            blnValidation = False
            txtFlightNumber.Focus()
        End If
    End Sub

    Private Sub Validate_MilesFlown(ByRef intMilesFlown As Integer, ByRef blnValidation As Boolean)
        If Integer.TryParse(txtMilesFlown.Text, intMilesFlown) Then
            intMilesFlown = txtMilesFlown.Text
            If (intMilesFlown < 1) Or (intMilesFlown >= 10000) Then
                MessageBox.Show("User Input Must Be Greater than 0 and Less or Equal to 9999")
                txtMilesFlown.Focus()
                Exit Sub
            End If
        Else
            MessageBox.Show("User Input Must Exist and Must be Numeric")
            blnValidation = False
            txtMilesFlown.Focus()
        End If
    End Sub

    Private Function TimeOfDepString()
        Dim dtmTimeofDeparture As String = ""
        Dim intTime As Integer
        Dim strString As String = ":"
        Dim strZero As String = "0"
        intTime = cboTimeOfDepHours.Value
        If chkTimeOfDep.Checked = True Then
            intTime += 12
        End If
        If intTime <= 9 Then
            dtmTimeofDeparture = dtmTimeofDeparture.Insert(0, strZero)
            dtmTimeofDeparture = dtmTimeofDeparture.Insert(1, intTime)
        Else
            dtmTimeofDeparture = dtmTimeofDeparture.Insert(0, intTime.ToString)
        End If
        dtmTimeofDeparture = dtmTimeofDeparture.Insert(2, strString)
        intTime = cboTimeOfDepMinutes.Value
        If intTime <= 9 Then
            dtmTimeofDeparture = dtmTimeofDeparture.Insert(3, strZero)
            dtmTimeofDeparture = dtmTimeofDeparture.Insert(4, intTime)
        Else
            dtmTimeofDeparture = dtmTimeofDeparture.Insert(3, intTime.ToString)
        End If
        dtmTimeofDeparture = dtmTimeofDeparture.Insert(5, strString)
        strZero = "00"
        dtmTimeofDeparture = dtmTimeofDeparture.Insert(6, strZero)

        Return dtmTimeofDeparture
    End Function
    Private Function TimeOfLandString()
        Dim dtmTimeOfLanding As String = ""
        Dim intTime As Integer
        Dim strString As String = ":"
        Dim strZero As String = "0"
        intTime = cboTimeOfLandHours.Value
        If chkTimeOfLand.Checked = True Then
            intTime += 12
        End If
        If intTime <= 9 Then
            dtmTimeOfLanding = dtmTimeOfLanding.Insert(0, strZero)
            dtmTimeOfLanding = dtmTimeOfLanding.Insert(1, intTime)
        Else
            dtmTimeOfLanding = dtmTimeOfLanding.Insert(0, intTime.ToString)
        End If
        dtmTimeOfLanding = dtmTimeOfLanding.Insert(2, strString)
        intTime = cboTimeOfLandMinutes.Value
        If intTime <= 9 Then
            dtmTimeOfLanding = dtmTimeOfLanding.Insert(3, strZero)
            dtmTimeOfLanding = dtmTimeOfLanding.Insert(4, intTime)
        Else
            dtmTimeOfLanding = dtmTimeOfLanding.Insert(3, intTime.ToString)
        End If
        dtmTimeOfLanding = dtmTimeOfLanding.Insert(5, strString)
        strZero = "00"
        dtmTimeOfLanding = dtmTimeOfLanding.Insert(6, strZero)

        Return dtmTimeOfLanding
    End Function
    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        Dim frmAdminPAS As New frmAdminPAS
        frmAdminPAS.Show()
        Me.Hide()
    End Sub

    Private Function ValidationDate()
        Dim blnValidation As Boolean
        MessageBox.Show("Date must be after current day and vaild format 00/00/0000.")
        txtFlightDate.Focus()
        blnValidation = False
        Return blnValidation
    End Function

End Class