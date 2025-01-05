Public Class frmAttendantsAdd
    Private Sub btnSubmit_Click(sender As Object, e As EventArgs) Handles btnSubmit.Click
        Dim strFirstName As String
        Dim strLastName As String
        Dim intEmployeeID As Integer
        Dim dtmDateOfHire As String
        Dim dtmDateOfTermination As String
        Dim strLogin As String
        Dim strPassword As String
        Dim blnValidation As Boolean = True
        Get_And_Validate_Input(strFirstName, strLastName, intEmployeeID, dtmDateOfHire, dtmDateOfTermination, strLogin, strPassword, blnValidation)
        If blnValidation = True Then
            Dim strSelect As String
            Dim cmdSelect As OleDb.OleDbCommand
            Dim drSourceTable As OleDb.OleDbDataReader
            Dim intNextPrimaryKey As Integer
            Dim strInsert As String
            Dim cmdInsert As OleDb.OleDbCommand
            Dim intRowsAffected As Integer
            Dim cmdAddLogin As New OleDb.OleDbCommand()
            Dim intPKID As Integer

            If OpenDatabaseConnectionSQLServer() = False Then

                MessageBox.Show(Me, "Database connection error." & vbNewLine &
                                   "The application will now close.",
                                   Me.Text + " Error",
                                   MessageBoxButtons.OK, MessageBoxIcon.Error)

                Me.Close()
            End If

            strSelect = "SELECT MAX(intAttendantID) + 1 AS intNextPrimaryKey" & " FROM TAttendants"

            cmdSelect = New OleDb.OleDbCommand(strSelect, m_conAdministrator)
            drSourceTable = cmdSelect.ExecuteReader
            drSourceTable.Read()
            If drSourceTable.IsDBNull(0) = True Then
                intNextPrimaryKey = 1
            Else
                intNextPrimaryKey = CInt(drSourceTable("intNextPrimaryKey"))
            End If
            strInsert = "INSERT INTO TAttendants (intAttendantID, strFirstName, strLastName, strEmployeeID, dtmDateOfHire, dtmDateOfTermination)
            VALUES (" & intNextPrimaryKey & ", '" & strFirstName & "', '" & strLastName & "', '" & intEmployeeID & "', '" & dtmDateOfHire & "', '" & dtmDateOfTermination & "')"


            cmdInsert = New OleDb.OleDbCommand(strInsert, m_conAdministrator)

            intRowsAffected = cmdInsert.ExecuteNonQuery()
            If intRowsAffected > 0 Then
                MessageBox.Show("Attendant has been added")
            End If
            cmdAddLogin.CommandText = "EXECUTE uspAddEmployeeLogin " & intPKID & ", '" & strLogin & "', '" & strPassword & "', 'Attendant', " & intNextPrimaryKey
            cmdAddLogin.CommandType = CommandType.StoredProcedure
            cmdAddLogin = New OleDb.OleDbCommand(cmdAddLogin.CommandText, m_conAdministrator)
            CloseDatabaseConnection()
            Close()
            Dim frmAttendantsAdd As New frmAttendantsAdd
            frmAttendantsAdd.Show()
            Me.Hide()
        Else
            Dim frmAttendantsAdd As New frmAttendantsAdd
            frmAttendantsAdd.Show()
            Exit Sub
        End If
    End Sub
    Private Sub Get_And_Validate_Input(ByRef strFirstName As String, ByRef strLastName As String, ByRef intEmployeeID As Integer, ByRef dtmDateOfHire As String, ByRef dtmDateOfTermination As String, ByRef strLogin As String, ByRef strPassword As String, ByRef blnValidation As Boolean)
        Validate_FirstName(strFirstName, blnValidation)
        Validate_LastName(strLastName, blnValidation)
        Validate_EmployeeID(intEmployeeID, blnValidation)
        Validate_DateOfHire(dtmDateOfHire, blnValidation)
        Validate_DateOfTermination(dtmDateOfTermination, blnValidation)
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
                    If dtmDateOfHire.IndexOf("/") = 2 And 5 Then
                    Else
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
                    If dtmDateOfTermination.IndexOf("/") = 2 And 5 Then
                    Else
                        blnValidation = ValidationDateTermination()
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

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        Dim frmAdminA As New frmAdminA
        frmAdminA.Show()
        Me.Hide()
    End Sub

    Private Function ValidationDateHire()
        Dim blnValidation As Boolean
        MessageBox.Show("Date must be vaild format 00/00/0000.")
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
End Class