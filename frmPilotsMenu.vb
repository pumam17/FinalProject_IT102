Public Class frmPilotsMenu
    Private Sub frmPilotsMenu_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim strSelect As String
        Dim cmdSelect As OleDb.OleDbCommand
        Dim drSourceTable As OleDb.OleDbDataReader
        Dim dt As DataTable = New DataTable
        Dim objParam As OleDb.OleDbParameter

        Try
            If OpenDatabaseConnectionSQLServer() = False Then

                MessageBox.Show(Me, "Database connection error." & vbNewLine &
                                   "The application will now close.",
                                   Me.Text + " Error",
                                   MessageBoxButtons.OK, MessageBoxIcon.Error)

                Me.Close()
            End If
            strSelect = "SELECT intPilotRoleID, strPilotRole From TPilotRoles"

            cmdSelect = New OleDb.OleDbCommand(strSelect, m_conAdministrator)
            drSourceTable = cmdSelect.ExecuteReader
            dt.Load(drSourceTable)

            cboPilotRoles.ValueMember = "intPilotRoleID"
            cboPilotRoles.DisplayMember = "strPilotRole"
            cboPilotRoles.DataSource = dt

            '--------------------------------------------------------------------

            strSelect = "SELECT intPilotID, strFirstName, strLastName, strEmployeeID, dtmDateofHire, dtmDateofTermination, dtmDateofLicense, intPilotRoleID 
            FROM TPilots Where intPilotID = " & strString
            cmdSelect = New OleDb.OleDbCommand(strSelect, m_conAdministrator)
            drSourceTable = cmdSelect.ExecuteReader
            drSourceTable.Read()

            txtFristName.Text = drSourceTable("strFirstName")
            txtLastName.Text = drSourceTable("strLastName")
            txtEmployeeID.Text = drSourceTable("strEmployeeID")
            txtDateOfHire.Text = drSourceTable("dtmDateofHire")
            txtDateOfTermination.Text = drSourceTable("dtmDateofTermination")
            txtDateOfLicense.Text = drSourceTable("dtmDateofLicense")
            cboPilotRoles.SelectedValue = drSourceTable("intPilotRoleID")


            cmdSelect = New OleDb.OleDbCommand("uspLoginInfo", m_conAdministrator)
            cmdSelect.CommandType = CommandType.StoredProcedure
            objParam = cmdSelect.Parameters.Add("@strEmployeeRole", OleDb.OleDbType.VarChar)
            objParam.Direction = ParameterDirection.Input
            objParam.Value = "Pilot"
            objParam = cmdSelect.Parameters.Add("@intEmployeePK", OleDb.OleDbType.Integer)
            objParam.Direction = ParameterDirection.Input
            objParam.Value = strString
            drSourceTable = cmdSelect.ExecuteReader
            drSourceTable.Read()

            txtUsername.Text = drSourceTable("strEmployeeLoginID")
            txtPassword.Text = drSourceTable("strEmployeePassword")
            CloseDatabaseConnection()
        Catch ex As Exception
            ' Log and display error message
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub btnUpdatePilot_Click(sender As Object, e As EventArgs) Handles btnUpdatePilot.Click

        Dim strFirstName As String
        Dim strLastName As String
        Dim intEmployeeID As String
        Dim dtmDateOfHire As String
        Dim dtmDateOfTermination As String
        Dim dtmDateOfLicense As String
        Dim intPilotRoleID As Integer
        Dim strLogin As String
        Dim strPassword As String
        Dim blnValidation As Boolean = True
        Get_And_Validate_Input(strFirstName, strLastName, intEmployeeID, dtmDateOfHire, dtmDateOfTermination, dtmDateOfLicense, strLogin, strPassword, blnValidation)
        If blnValidation = True Then
            Dim strSelect As String
            Dim cmdUpdate As OleDb.OleDbCommand
            Dim intRowsAffected As Integer
            Dim cmdUpdateLogin As New OleDb.OleDbCommand()
            intPilotRoleID = cboPilotRoles.SelectedValue
            If OpenDatabaseConnectionSQLServer() = False Then

                MessageBox.Show(Me, "Database connection error." & vbNewLine &
                                   "The application will now close.",
                                   Me.Text + " Error",
                                   MessageBoxButtons.OK, MessageBoxIcon.Error)

                Me.Close()
            End If

            strSelect = "Update TPilots Set " &
                    "strFirstName = '" & strFirstName & "', " &
                    "strLastName = '" & strLastName & "', " &
                    "strEmployeeID = '" & intEmployeeID & "', " &
                    "dtmDateOfHire = '" & dtmDateOfHire & "', " &
                    "dtmDateOfTermination = '" & dtmDateOfTermination & "', " &
                    "dtmDateOfLicense = '" & dtmDateOfLicense & "', " &
                    "intPilotRoleID = " & intPilotRoleID &
                    "Where intPilotID = " & strString


            cmdUpdate = New OleDb.OleDbCommand(strSelect, m_conAdministrator)
            intRowsAffected = cmdUpdate.ExecuteNonQuery()
            If intRowsAffected > 0 Then
                MessageBox.Show("Pilot has been Updated")
            End If


            cmdUpdateLogin.CommandText = "EXECUTE uspUpdateEmployeeLogin 'Pilot', " & strString & ", '" & strLogin & "', '" & strPassword & "'"
            cmdUpdateLogin.CommandType = CommandType.StoredProcedure
            cmdUpdateLogin = New OleDb.OleDbCommand(cmdUpdateLogin.CommandText, m_conAdministrator)
            CloseDatabaseConnection()
            Close()
            Dim frmPilotsMenu As New frmPilotsMenu
            frmPilotsMenu.Show()

            Me.Hide()
        Else
            Exit Sub
        End If
    End Sub

    Private Sub Get_And_Validate_Input(ByRef strFirstName As String, ByRef strLastName As String, ByRef intEmployeeID As Integer, ByRef dtmDateOfHire As String, ByRef dtmDateOfTermination As String, ByRef dtmDateOfLicense As String, ByRef strLogin As String, ByRef strPassword As String, ByRef blnValidation As Boolean)
        Validate_FirstName(strFirstName, blnValidation)
        Validate_LastName(strLastName, blnValidation)
        Validate_EmployeeID(intEmployeeID, blnValidation)
        Validate_DateOfHire(dtmDateOfHire, blnValidation)
        Validate_DateOfTermination(dtmDateOfTermination, blnValidation)
        Validate_DateOfLicense(dtmDateOfLicense, blnValidation)
        Validate_Login(strLogin, blnValidation)
        Validate_Password(strPassword, blnValidation)
    End Sub

    Private Sub Validate_FirstName(ByRef strFirstName As String, ByRef blnValidation As Boolean)
        If txtFristName.Text = String.Empty Then
            MessageBox.Show("First Name Must Exist")
            txtFristName.Focus()
            blnValidation = False
        Else
            strFirstName = txtFristName.Text
        End If
    End Sub

    Private Sub Validate_LastName(ByRef strLastName As String, ByRef blnValidation As Boolean)
        If txtLastName.Text = String.Empty Then
            MessageBox.Show("Last Name Must Exist")
            txtLastName.Focus()
            blnValidation = False
        Else
            strLastName = txtLastName.Text
        End If
    End Sub

    Private Sub Validate_EmployeeID(ByRef intEmployeeID As Integer, ByRef blnValidation As Boolean)
        If Integer.TryParse(txtEmployeeID.Text, intEmployeeID) Then
            intEmployeeID = txtEmployeeID.Text
            If intEmployeeID.ToString.Length <> 5 Then
                blnValidation = False
                MessageBox.Show("EmployeeID must be 5 numbers")
                txtEmployeeID.Focus()
            End If
        Else
            MessageBox.Show("EmployeeID must be numeric Must Exist")
            txtEmployeeID.Focus()
            blnValidation = False
        End If
    End Sub
    Private Sub Validate_DateOfHire(ByRef dtmDateOfHire As String, ByRef blnValidation As Boolean)
        Dim strString As String
        If txtDateOfHire.Text = String.Empty Then
            blnValidation = ValidationDateHire()
        Else
            dtmDateOfHire = txtDateOfHire.Text
            If dtmDateOfHire.Length <> 10 Then
                blnValidation = ValidationDateHire()
            Else
                strString = dtmDateOfHire.Substring(0, 2)
                If IsNumeric(strString) Then
                    If strString < 0 Then
                        blnValidation = ValidationDateHire()
                    End If
                    If strString > 12 Then
                        blnValidation = ValidationDateHire()
                    End If
                Else
                    blnValidation = ValidationDateHire()
                End If
                strString = dtmDateOfHire.Substring(3, 2)
                If IsNumeric(strString) Then
                    If strString < 0 Then
                        blnValidation = ValidationDateHire()
                    End If
                    If strString > 31 Then
                        blnValidation = ValidationDateHire()
                    End If
                Else
                    blnValidation = ValidationDateHire()
                End If
                strString = dtmDateOfHire.Substring(6, 4)
                If IsNumeric(strString) Then
                    If dtmDateOfHire.IndexOf("/") <> 2 And 5 Then
                        blnValidation = ValidationDateHire()
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub Validate_DateOfTermination(ByRef dtmDateOfTermination As String, ByRef blnValidation As Boolean)
        Dim strString As String
        If txtDateOfTermination.Text = String.Empty Then
            blnValidation = ValidationDateTermination()
        Else
            dtmDateOfTermination = txtDateOfTermination.Text
            If dtmDateOfTermination.Length <> 10 Then
                blnValidation = ValidationDateTermination()
            Else
                strString = dtmDateOfTermination.Substring(0, 2)
                If IsNumeric(strString) Then
                    If strString < 0 Then
                        blnValidation = ValidationDateTermination()
                    End If
                    If strString > 12 Then
                        blnValidation = ValidationDateTermination()
                    End If
                Else
                    blnValidation = ValidationDateTermination()
                End If
                strString = dtmDateOfTermination.Substring(3, 2)
                If IsNumeric(strString) Then
                    If strString < 0 Then
                        blnValidation = ValidationDateTermination()
                    End If
                    If strString > 31 Then
                        blnValidation = ValidationDateTermination()
                    End If
                Else
                    blnValidation = ValidationDateTermination()
                End If
                strString = dtmDateOfTermination.Substring(6, 4)
                If IsNumeric(strString) Then
                    If dtmDateOfTermination.IndexOf("/") <> 2 And 5 Then
                        blnValidation = ValidationDateTermination()
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub Validate_DateOfLicense(ByRef dtmDateOfLicense As String, ByRef blnValidation As Boolean)
        Dim strString As String
        If txtDateOfLicense.Text = String.Empty Then
            blnValidation = ValidationDateTermination()
        Else
            dtmDateOfLicense = txtDateOfLicense.Text
            If dtmDateOfLicense.Length <> 10 Then
                blnValidation = ValidationDateTermination()
            Else
                strString = dtmDateOfLicense.Substring(0, 2)
                If IsNumeric(strString) Then
                    If strString < 0 Then
                        blnValidation = ValidationDateTermination()
                    End If
                    If strString > 12 Then
                        blnValidation = ValidationDateTermination()
                    End If
                Else
                    blnValidation = ValidationDateTermination()
                End If
                strString = dtmDateOfLicense.Substring(3, 2)
                If IsNumeric(strString) Then
                    If strString < 0 Then
                        blnValidation = ValidationDateTermination()
                    End If
                    If strString > 31 Then
                        blnValidation = ValidationDateTermination()
                    End If
                Else
                    blnValidation = ValidationDateTermination()
                End If
                strString = dtmDateOfLicense.Substring(6, 4)
                If IsNumeric(strString) Then
                    If dtmDateOfLicense.IndexOf("/") <> 2 And 5 Then
                        blnValidation = ValidationDateLicense()
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub Validate_Login(ByRef strLogin As String, ByRef blnValidation As Boolean)
        If txtUsername.Text = String.Empty Then
            MessageBox.Show("Username Must Exist")
            txtUsername.Focus()
            blnValidation = False
        Else
            strLogin = txtUsername.Text
        End If
    End Sub

    Private Sub Validate_Password(ByRef strPassword As String, ByRef blnValidation As Boolean)
        If txtPassword.Text = String.Empty Then
            MessageBox.Show("Password Must Exist")
            txtPassword.Focus()
            blnValidation = False
        Else
            strPassword = txtPassword.Text
        End If
    End Sub

    Private Function ValidationDateHire()
        Dim blnValidation As Boolean
        MessageBox.Show("Date must be after current day and vaild format 00/00/0000.")
        txtDateOfHire.Focus()
        blnValidation = False
        Return blnValidation
    End Function

    Private Function ValidationDateTermination()
        Dim blnValidation As Boolean
        MessageBox.Show("Date must be vaild format 00/00/0000.")
        txtDateOfTermination.Focus()
        blnValidation = False
        Return blnValidation
    End Function

    Private Function ValidationDateLicense()
        Dim blnValidation As Boolean
        MessageBox.Show("Date must be vaild format 00/00/0000.")
        txtDateOfLicense.Focus()
        blnValidation = False
        Return blnValidation
    End Function
    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        Dim frmELogin As New frmELogin
        frmELogin.Show()
        Me.Hide()
    End Sub

    Private Sub btnPastFlights_Click(sender As Object, e As EventArgs) Handles btnPastFlights.Click
        Dim frmPiPFlights As New frmPiPFlights
        frmPiPFlights.Show()
        Me.Hide()
    End Sub

    Private Sub btnFutureFlights_Click(sender As Object, e As EventArgs) Handles btnFutureFlights.Click
        Dim frmPiFFlights As New frmPiFFlights
        frmPiFFlights.Show()
        Me.Hide()
    End Sub
End Class