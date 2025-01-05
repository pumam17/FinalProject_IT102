Public Class frmPilotsDelete
    Private Sub frmPilotsDelete_Load(sender As Object, e As EventArgs) Handles MyBase.Load
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
        Try
            strSelect = "SELECT intPilotID, strFirstName + ' ' + strLastName as PilotName From TPilots"

            cmdSelect = New OleDb.OleDbCommand(strSelect, m_conAdministrator)
            drSourceTable = cmdSelect.ExecuteReader
            dt.Load(drSourceTable)

            cboPilots.ValueMember = "intPilotID"
            cboPilots.DisplayMember = "PilotName"
            cboPilots.DataSource = dt


        Catch ex As Exception
            ' Log and display error message
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub btnSubmit_Click(sender As Object, e As EventArgs) Handles btnSubmit.Click

        Dim cmdDeletePilot As New OleDb.OleDbCommand()
        Dim intRowsAffected As Integer
        Dim results As DialogResult
        Try

            If OpenDatabaseConnectionSQLServer() = False Then

                MessageBox.Show(Me, "Database connection error." & vbNewLine &
                                   "The application will now close.",
                                   Me.Text + " Error",
                                   MessageBoxButtons.OK, MessageBoxIcon.Error)

                Me.Close()
            End If
            results = MessageBox.Show("Are you sure you want to delete pilot: " & cboPilots.Text & "?", "Confrim Deletion", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)
            Select Case results
                Case DialogResult.Cancel
                    MessageBox.Show("Action Canceled")
                Case DialogResult.No
                    MessageBox.Show("Action Canceled")
                Case DialogResult.Yes
                    cmdDeletePilot.CommandText = "EXECUTE uspDeletePilot " & cboPilots.SelectedValue.ToString
                    cmdDeletePilot.CommandType = CommandType.StoredProcedure
                    cmdDeletePilot = New OleDb.OleDbCommand(cmdDeletePilot.CommandText, m_conAdministrator)
                    intRowsAffected = cmdDeletePilot.ExecuteNonQuery()
            End Select
            If intRowsAffected > 0 Then
                MessageBox.Show("Pilot has been Deleted")
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
        Dim frmPilotsDelete As New frmPilotsDelete
        frmPilotsDelete.Show()
        Me.Hide()
    End Sub

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        Dim frmAdminP As New frmAdminP
        frmAdminP.Show()
        Me.Hide()
    End Sub
End Class