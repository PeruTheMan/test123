Imports System.Data.SqlClient
Public Class frmRole
    Private objRoles As CRoles
    Private blnClearing As Boolean
    Private blnReloading As Boolean
#Region "Toolbar Stuff"

    Private Sub tsbProxy_MouseEnter(sender As Object, e As EventArgs)
        'We need to do this only because we are not putting our images in the Image proerprty of the toolbar buttons
        Dim tsbProxy As ToolStripButton
        tsbProxy = DirectCast(sender, ToolStripButton)
        tsbProxy.DisplayStyle = ToolStripItemDisplayStyle.Text

    End Sub

    Private Sub tsbProxy_MouseLeave(sender As Object, e As EventArgs)
        'We need to do this only because we are not putting our images in the Image proerprty of the toolbar buttons
        Dim tsbProxy As ToolStripButton
        tsbProxy = DirectCast(sender, ToolStripButton)
        tsbProxy.DisplayStyle = ToolStripItemDisplayStyle.Image

    End Sub

    Private Sub tsbCourse_Click(sender As Object, e As EventArgs)
        intNextAction = ACTION_COURSE
        Me.Hide()
    End Sub
    Private Sub tsbEvent_Click(sender As Object, e As EventArgs)
        intNextAction = ACTION_EVENT
        Me.Hide()
    End Sub

    Private Sub tsbHelp_Click(sender As Object, e As EventArgs)
        intNextAction = ACTION_HELP
        Me.Hide()

    End Sub

    Private Sub tsbHome_Click(sender As Object, e As EventArgs)
        intNextAction = ACTION_HOME
        Me.Hide()
    End Sub

    Private Sub tsbLogout_Click(sender As Object, e As EventArgs)
        intNextAction = ACTION_LOGOUT
        Me.Hide()
    End Sub

    Private Sub tsbMember_Click(sender As Object, e As EventArgs)
        intNextAction = ACTION_MEMBER
        Me.Hide()
    End Sub

    Private Sub tsbRole_Click(sender As Object, e As EventArgs)
        'no action needed- already in the ROLE form
    End Sub

    Private Sub tsbRSVP_Click(sender As Object, e As EventArgs)
        intNextAction = ACTION_RSVP
        Me.Hide()
    End Sub

    Private Sub tsbSemester_Click(sender As Object, e As EventArgs)
        intNextAction = ACTION_SEMESTER
        Me.Hide()
    End Sub
#End Region
#Region "TextBoxes"
    Private Sub txtRoleID_GotFocus(sender As Object, e As EventArgs) Handles txtRoleID.GotFocus, txtRoleID.GotFocus
        Dim txtBox As TextBox
        txtBox = DirectCast(sender, TextBox)
        txtBox.SelectAll()
    End Sub
    Private Sub txtRoleID_LostFocus(sender As Object, e As EventArgs) Handles txtRoleID.LostFocus, txtRoleID.LostFocus
        Dim txtBox As TextBox
        txtBox = DirectCast(sender, TextBox)
        txtBox.DeselectAll()
    End Sub

#End Region
    Private Sub LoadRoles()
        Dim objDR As SqlDataReader
        lstRoles.Items.Clear()
        Try
            objDR = objRoles.GetAllRoles()
            Do While objDR.Read
                lstRoles.Items.Add(objDR.Item("RoleID"))
            Loop
            objDR.Close()
        Catch ex As Exception
            'could have CDB throw the error
        End Try
        If objRoles.CurrentObject.RoleID <> "" Then
            lstRoles.SelectedIndex = lstRoles.FindStringExact(objRoles.CurrentObject.RoleID)
        End If
        blnReloading = False
    End Sub

    Private Sub frmRole_Load(sender As Object, e As EventArgs) Handles Me.Load
        objRoles = New CRoles
    End Sub

    Private Sub frmRole_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        ClearScreenControls(Me)
        LoadRoles()
        grpEdit.Enabled = False
    End Sub

    Private Sub lstRoles_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lstRoles.SelectedIndexChanged
        If blnClearing Then
            Exit Sub
        End If
        If blnReloading Then
            tslStatus.Text = " "
            Exit Sub
        End If
        If lstRoles.SelectedIndex = -1 Then
            Exit Sub
        End If
        chkNew.Checked = False
        LoadSelectedRecord()
        grpEdit.Enabled = True
    End Sub
    Private Sub LoadSelectedRecord()
        Try
            objRoles.GetRoleByRoleId(lstRoles.SelectedItem.ToString)
            With objRoles.CurrentObject
                txtRoleID.Text = .RoleId
                txtDesc.Text = .RoleDescription

            End With
        Catch ex As Exception
            MessageBox.Show("Error loading Role values: " & ex.ToString, "Program error", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try
    End Sub

    Private Sub chkNew_CheckedChanged(sender As Object, e As EventArgs) Handles chkNew.CheckedChanged
        If blnClearing Then
            Exit Sub
        End If
        If chkNew.Checked Then
            tslStatus.Text = ""
            txtRoleID.Clear()
            txtDesc.Clear()
            lstRoles.SelectedIndex = -1
            grpRoles.Enabled = False
            grpEdit.Enabled = True
            txtRoleID.Focus()
            objRoles.CreateNewRole()
        Else
            grpRoles.Enabled = True
            grpEdit.Enabled = False
            objRoles.CurrentObject.isNewRole = False

        End If
    End Sub



    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        Dim intResult As Integer
        Dim blnErrors As Boolean
        tslStatus.Text = ""
        '---- add your validation code here
        If Not ValidateTextBoxLength(txtRoleID, errP) Then
            blnErrors = True
        End If
        If Not ValidateTextBoxLength(txtDesc, errP) Then
            blnErrors = True
        End If
        If blnErrors Then
            Exit Sub
        End If

        With objRoles.CurrentObject
            .RoleID = txtRoleID.Text
            .RoleDescription = txtDesc.Text

        End With
        Try
            Me.Cursor = Cursors.WaitCursor
            intResult = objRoles.Save
            If intResult = 1 Then
                tslStatus.Text = "Role record saved"
            End If
            If intResult = -1 Then 'role Id wasnt unqie when adding a new record
                MessageBox.Show("RoleID must be unique. unable to save Role record", "Database error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                tslStatus.Text = "Error"
            End If
        Catch ex As Exception
            MessageBox.Show("Unable to save Role record: " & ex.ToString, "Database error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            tslStatus.Text = "Error"

        End Try
        Me.Cursor = Cursors.Default
        blnReloading = True
        LoadRoles() ' reload so that a newly saved record will apear in the list
        chkNew.Checked = False
        grpRoles.Enabled = True ' in case it was diasbled for a new record
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        blnClearing = True
        tslStatus.Text = ""
        chkNew.Checked = False
        errP.Clear()
        If lstRoles.SelectedIndex <> -1 Then
            LoadSelectedRecord() ' reload what was selected in case uder had messed up form
        Else
            grpEdit.Enabled = False

        End If
        blnClearing = False
        objRoles.CurrentObject.isNewRole = False
        grpRoles.Enabled = True

    End Sub

    Private Sub btnReport_Click(sender As Object, e As EventArgs) Handles btnReport.Click
        Dim RoleReport As New frmRoleReport
        If lstRoles.Items.Count = 0 Then
            MessageBox.Show("There are no records to print")
            Exit Sub
        End If
        Me.Cursor = Cursors.WaitCursor
        RoleReport.Display()
        Me.Cursor = Cursors.Default

    End Sub


End Class