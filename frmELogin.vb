Public Class frmELogin
    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        Dim frmMainMenu As New frmMainMenu
        frmMainMenu.Show()
        Me.Hide()
    End Sub

    Private Sub btnSubmit_Click(sender As Object, e As EventArgs) Handles btnSubmit.Click
        Dim strUserName As String
        Dim strPassword As String
        Dim strEmployeeRole As String
        strUserName = txtUserName.Text
        strPassword = txtPassword.Text
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

            cmdSelect = New OleDb.OleDbCommand("uspELogin", m_conAdministrator)
            cmdSelect.CommandType = CommandType.StoredProcedure
            objParam = cmdSelect.Parameters.Add("@strEmployeeLoginID", OleDb.OleDbType.VarChar)
            objParam.Direction = ParameterDirection.Input
            objParam.Value = strUserName

            objParam = cmdSelect.Parameters.Add("@strEmployeePassword", OleDb.OleDbType.VarChar)
            objParam.Direction = ParameterDirection.Input
            objParam.Value = strPassword
            drSourceTable = cmdSelect.ExecuteReader
            While drSourceTable.Read()
                strEmployeeRole = drSourceTable("strEmployeeRole")
                strString = drSourceTable("intEmployeePK")
            End While
            ' Clean up
            drSourceTable.Close()
            ' close the database connection
            CloseDatabaseConnection()
        Catch ex As Exception
            ' Log and display error message
            MessageBox.Show(ex.Message)
        End Try
        LoginConnection(strEmployeeRole)
    End Sub

    Private Sub EmployeeRole(ByVal strEmployeeRole As String)
        Select Case strEmployeeRole
            Case Is = "Pilot"
                Dim frmPilotsMenu As New frmPilotsMenu
                frmPilotsMenu.Show()

                Me.Hide()
            Case Is = "Attendant"
                Dim frmAttendantsMenu As New frmAttendantsMenu
                frmAttendantsMenu.Show()

                Me.Hide()
            Case Is = "Admin"
                Dim frmAdminPAS As New frmAdminPAS
                frmAdminPAS.Show()

                Me.Hide()
        End Select

    End Sub
    Private Sub LoginConnection(ByVal strEmployeeRole As String)
        If CInt(strString) > 0 Then
            EmployeeRole(strEmployeeRole)
        Else
            MessageBox.Show("Incorrect login")
        End If

    End Sub
End Class