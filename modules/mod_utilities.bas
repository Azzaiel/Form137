Attribute VB_Name = "mod_utilities"

Public Sub load_form(frm_load As Form, Optional isModal As Boolean, Optional frm_unload As Form)
On Error Resume Next
    If isModal Then
        frm_load.Show vbModal
    Else
        frm_load.Show
    End If
    Unload frm_unload
End Sub

Public Function is_empty_textbox(frm As Form) As Boolean
    Dim cCont As Control
    For Each cCont In frm.Controls
        If (TypeOf cCont Is TextBox) Then
            If cCont.Text = "" Then
                is_empty_textbox = True
                Exit Function
            End If
        End If
    Next cCont
    is_empty_textbox = False
End Function

Public Function is_empty_dropdown(frm As Form) As Boolean
    Dim cCont As Control
    For Each cCont In frm.Controls
        If (TypeOf cCont Is ComboBox) Then
            If cCont.Text = "" Then
                is_empty_dropdown = True
                Exit Function
            End If
        End If
    Next cCont
    is_empty_dropdown = False
End Function

Public Function get_selected(opt_array As Variant) As String
    Dim opt As OptionButton
    For Each opt In opt_array
        If opt.Value = True Then
            get_selected = opt.Caption
            Exit Function
        End If
    Next
End Function

Public Function is_duplicate(table_name As String, field_name As String, val As String) As Boolean
    Call mysql_select(public_rs, "SELECT * FROM " & table_name & " WHERE " & field_name & " = '" & val & "'")
    If Not public_rs.RecordCount = 0 Then
        is_duplicate = True
    Else
        is_duplicate = False
    End If
End Function
