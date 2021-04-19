Module modObjects
    Public Sub ClearScreenControls(ByVal objContainer As Control)
        'This procedure will clear all controls on the contaier passed in as the argument
        'Not all control types have been implemented. Add as needed
        Dim obj As Control
        Dim strControlType As String
        For Each obj In objContainer.Controls
            strControlType = TypeName(obj) 'TypeName returns the class name of the control
            Select Case strControlType
                Case "TextBox"
                    Dim cntrl As TextBox
                    cntrl = DirectCast(obj, TextBox)
                    cntrl.Clear()
                    'or cntrl.Text=""
                    'or cntrl.Text- vbNullString
                Case "CheckBox"
                    Dim cntrl As CheckBox
                    cntrl = DirectCast(obj, CheckBox)
                    cntrl.Checked = False
                    'or cntrl.CheckState=CheckState.Unchecked
                Case "ComboBox"
                    Dim cntrl As ComboBox
                    cntrl = DirectCast(obj, ComboBox)
                    cntrl.SelectedIndex = -1 'clear only the sleection, not the list
                Case "RadioButton"
                    Dim cntrl As RadioButton
                    cntrl = DirectCast(obj, RadioButton)
                    cntrl.Checked = False
                Case "ListBox"
                    Dim cntrl As ListBox
                    cntrl = DirectCast(obj, ListBox)
                    cntrl.SelectedIndex = -1
                Case "MaskedTextBox"
                    Dim cntrl As MaskedTextBox
                    cntrl = DirectCast(obj, MaskedTextBox)
                    cntrl.Text = ""
                Case "GroupBox"
                    'must recurse through its controls coolection
                    Dim cntrl As GroupBox
                    cntrl = DirectCast(obj, GroupBox)
                    ClearScreenControls(cntrl)
                Case Else
                    'do nothing or add for future

            End Select

        Next
    End Sub
End Module
