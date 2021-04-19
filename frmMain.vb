Public Class frmMain
    Private RoleInfo As frmRole
    Private EventInfo As frmEvent
    Private EventRSVPInfo As frmRSVP
    Private Sub tsbProxy_MouseEnter(sender As Object, e As EventArgs) Handles tsbCourse.MouseEnter, tsbEvent.MouseEnter, tsbHelp.MouseEnter, tsbHome.MouseEnter, tsbLogout.MouseEnter, tsbMember.MouseEnter, tsbRole.MouseEnter, tsbRSVP.MouseEnter, tsbSemester.MouseEnter
        'We need to do this only because we are not putting our images in the Image proerprty of the toolbar buttons
        Dim tsbProxy As ToolStripButton
        tsbProxy = DirectCast(sender, ToolStripButton)
        tsbProxy.DisplayStyle = ToolStripItemDisplayStyle.Text

    End Sub

    Private Sub tsbProxy_MouseLeave(sender As Object, e As EventArgs) Handles tsbCourse.MouseLeave, tsbEvent.MouseLeave, tsbHelp.MouseLeave, tsbHome.MouseLeave, tsbLogout.MouseLeave, tsbMember.MouseLeave, tsbRole.MouseLeave, tsbRSVP.MouseLeave, tsbSemester.MouseLeave
        'We need to do this only because we are not putting our images in the Image proerprty of the toolbar buttons
        Dim tsbProxy As ToolStripButton
        tsbProxy = DirectCast(sender, ToolStripButton)
        tsbProxy.DisplayStyle = ToolStripItemDisplayStyle.Image

    End Sub

    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles Me.Load
        RoleInfo = New frmRole
        EventInfo = New frmEvent
        EventRSVPInfo = New frmRSVP
        Try
            myDB.OpenDB()
        Catch ex As Exception
            MessageBox.Show("Unble to open database. Connection string= " & gstrConn, "DB error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            EndProgram()
        End Try
    End Sub

    Private Sub EndProgram()
        'close each form except main
        Dim f As Form
        Me.Cursor = Cursors.WaitCursor
        For Each f In Application.OpenForms
            If f.Name <> Me.Name Then
                If Not f Is Nothing Then
                    f.Close()
                End If
            End If

        Next
        'close database connection
        If Not objSQLConn Is Nothing Then
            objSQLConn.Close()
            objSQLConn.Dispose()
        End If
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub tsbRole_Click(sender As Object, e As EventArgs) Handles tsbRole.Click
        Me.Hide()
        RoleInfo.ShowDialog()
        Me.Show()
        PerformNextAction()
    End Sub
    Private Sub tsbEvent_Click(sender As Object, e As EventArgs) Handles tsbEvent.Click
        Me.Hide()
        EventInfo.ShowDialog()
        Me.Show()
        PerformNextAction()
    End Sub

    Private Sub tsbRSVP_Click(sender As Object, e As EventArgs) Handles tsbRSVP.Click
        Me.Hide()
        EventRSVPInfo.ShowDialog()
        Me.Show()
        PerformNextAction()
    End Sub

    Private Sub PerformNextAction()
        'get the next action specified on the child form, and then simulate the click of that button here
        Select Case intNextAction
            Case ACTION_COURSE
                tsbCourse.PerformClick()
            Case ACTION_EVENT
                tsbEvent.PerformClick()
            Case ACTION_HELP
                tsbHelp.PerformClick()
            Case ACTION_HOME
                tsbHome.PerformClick()
            Case ACTION_LOGOUT
                tsbLogout.PerformClick()
            Case ACTION_MEMBER
                tsbMember.PerformClick()
            Case ACTION_NONE
                'nothing to do
            Case ACTION_ROLE
                tsbRole.PerformClick()
            Case ACTION_RSVP
                tsbRSVP.PerformClick()
            Case ACTION_SEMESTER
                tsbSemester.PerformClick()

            Case Else
                MessageBox.Show("Unexpected case value in frmMain:PerformNextAction", "Program error", MessageBoxButtons.OK, MessageBoxIcon.Error)


        End Select
    End Sub

    Private Sub tsbLogout_Click(sender As Object, e As EventArgs) Handles tsbLogout.Click
        EndProgram()
    End Sub
End Class
