VERSION 5.00
Begin VB.Form usereditform 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit User Account"
   ClientHeight    =   5745
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7635
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   4  'Icon
   Picture         =   "usereditform.frx":0000
   ScaleHeight     =   5745
   ScaleWidth      =   7635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txt_oldusername 
      Height          =   375
      Left            =   360
      TabIndex        =   22
      Top             =   4680
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txt_oldid 
      Height          =   375
      Left            =   360
      TabIndex        =   21
      Top             =   5040
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txt_confirm_password 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   3000
      MaxLength       =   35
      PasswordChar    =   "*"
      TabIndex        =   8
      Top             =   4200
      Width           =   3495
   End
   Begin VB.TextBox txt_new_password 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   3000
      MaxLength       =   35
      PasswordChar    =   "*"
      TabIndex        =   7
      Top             =   3720
      Width           =   3495
   End
   Begin VB.TextBox txt_middlename 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      MaxLength       =   35
      TabIndex        =   3
      Top             =   1800
      Width           =   3495
   End
   Begin VB.TextBox txt_firstname 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      MaxLength       =   35
      TabIndex        =   2
      Top             =   1320
      Width           =   3495
   End
   Begin VB.TextBox txt_lastname 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      MaxLength       =   35
      TabIndex        =   1
      Top             =   840
      Width           =   3495
   End
   Begin VB.TextBox txt_id 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      MaxLength       =   15
      TabIndex        =   0
      Top             =   360
      Width           =   2175
   End
   Begin VB.TextBox txt_usertype 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   4
      Top             =   2280
      Width           =   3495
   End
   Begin VB.TextBox txt_username 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      MaxLength       =   35
      TabIndex        =   5
      Top             =   2760
      Width           =   3495
   End
   Begin VB.TextBox txt_old_password 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   3000
      MaxLength       =   35
      PasswordChar    =   "*"
      TabIndex        =   6
      Top             =   3240
      Width           =   3495
   End
   Begin VB.CommandButton cmd_save 
      Height          =   615
      Left            =   2640
      Picture         =   "usereditform.frx":1B3CE
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4920
      Width           =   1095
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Note: Please fill-up important fields."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   0
      Width           =   7335
   End
   Begin VB.Label lbl_default_password 
      BackStyle       =   0  'Transparent
      Caption         =   "Click for default password."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   4320
      TabIndex        =   11
      Top             =   5280
      Width           =   2895
   End
   Begin VB.Label lbl_default_username 
      BackStyle       =   0  'Transparent
      Caption         =   "Click for default username."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   4320
      TabIndex        =   10
      Top             =   4920
      Width           =   2895
   End
   Begin VB.Label lbl_password3 
      BackStyle       =   0  'Transparent
      Caption         =   "Confirm Password:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   840
      TabIndex        =   20
      Top             =   4320
      Width           =   2055
   End
   Begin VB.Label lbl_password2 
      BackStyle       =   0  'Transparent
      Caption         =   "New Password:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   840
      TabIndex        =   19
      Top             =   3840
      Width           =   2055
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Middle Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   840
      TabIndex        =   18
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "*First Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   840
      TabIndex        =   17
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "*Last Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   840
      TabIndex        =   16
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "*ID:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   840
      TabIndex        =   15
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "User Type:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   840
      TabIndex        =   14
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "*Username:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   840
      TabIndex        =   13
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label lbl_password1 
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   840
      TabIndex        =   12
      Top             =   3360
      Width           =   1935
   End
End
Attribute VB_Name = "usereditform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public rs_user As New ADODB.Recordset
Dim sql_string As String
Dim USERNAME, lastname, PASSWORD As String
Private Sub cmd_save_Click()
    Dim ans As String
     If txt_id.Text = "" Or txt_lastname.Text = "" Or txt_firstname.Text = "" Or txt_username.Text = "" Then
        MsgBox "Please input required fields."
        Exit Sub
    Else
        If txt_id.Text <> txt_oldid.Text Then
            If is_duplicate("tbl_user", "ID", txt_id.Text) Then
                MsgBox "ID already exists."
                Exit Sub
            ElseIf is_duplicate("tbl_teacher", "teacher_id", txt_id.Text) Then
                MsgBox "ID already exists."
                Exit Sub
            ElseIf is_duplicate("tbl_student", "student_id", txt_id.Text) Then
                MsgBox "ID already exists."
                Exit Sub
            End If
        End If
        If txt_old_password.Text = "" And txt_new_password.Text = "" And txt_confirm_password.Text = "" Then
            If txt_username.Text <> txt_oldusername.Text Then
                If is_duplicate("tbl_user", "Username", txt_username.Text) Then
                    MsgBox "Username already exists."
                    Exit Sub
                Else
                    ans = MsgBox("Are you sure you want to update user account?", vbYesNo, "Update User Account")
                    If ans = vbNo Then
                        Exit Sub
                    Else
                    sql_string = "UPDATE tbl_user SET ID='" & txt_id.Text & "', Username='" & txt_username.Text & "', Lastname='" & txt_lastname.Text & "', Firstname='" & txt_firstname.Text & "', Middlename = '" & txt_middlename.Text & "' WHERE Username='" & txt_oldusername.Text & "'"
                    Call mysql_select(usereditform.rs_user, sql_string)
                    sql_string = "UPDATE tbl_security SET Username='" & txt_username.Text & "' WHERE Username='" & txt_oldusername.Text & "'"
                    Call mysql_select(usereditform.rs_user, sql_string)
                    MsgBox "User account successfully updated."
                    Call userform.Form_Load
                    End If
                End If
            Else
                 ans = MsgBox("Are you sure you want to update user account?", vbYesNo, "Update User Account")
                    If ans = vbNo Then
                        Exit Sub
                    Else
                sql_string = "UPDATE tbl_user SET ID='" & txt_id.Text & "', Lastname='" & txt_lastname.Text & "', Firstname='" & txt_firstname.Text & "', Middlename = '" & txt_middlename.Text & "' WHERE Username='" & txt_oldusername.Text & "'"
                Call mysql_select(usereditform.rs_user, sql_string)
                MsgBox "User account successfully updated."
                Call userform.Form_Load
                End If
            End If
        Else
            If txt_username.Text <> txt_oldusername.Text Then
                If is_duplicate("tbl_user", "Username", txt_username.Text) Then
                    MsgBox "Username already exists."
                    Exit Sub
                Else
                    Call mysql_select(public_rs, "SELECT * FROM tbl_user WHERE Username = '" & txt_oldusername.Text & "'")
                    PASSWORD = public_rs.Fields("Password")
                    If txt_usertype.Text = "Administrator" Then
                    If PASSWORD = txt_old_password.Text Then
                        If txt_new_password.Text = txt_confirm_password.Text Then
                             ans = MsgBox("Are you sure you want to update user account?", vbYesNo, "Update User Account")
                                If ans = vbNo Then
                                    Exit Sub
                                Else
                            sql_string = "UPDATE tbl_user SET ID='" & txt_id.Text & "', Username='" & txt_username.Text & "', Lastname='" & txt_lastname.Text & "', Firstname='" & txt_firstname.Text & "', Middlename = '" & txt_middlename.Text & "', Password='" & txt_new_password.Text & "' WHERE Username='" & txt_oldusername.Text & "'"
                            Call mysql_select(usereditform.rs_user, sql_string)
                            sql_string = "UPDATE tbl_security SET Username='" & txt_username.Text & "' WHERE Username='" & txt_oldusername.Text & "'"
                            Call mysql_select(usereditform.rs_user, sql_string)
                            MsgBox "User account successfully updated."
                            Call userform.Form_Load
                            End If
                        Else
                            MsgBox "New password did not match."
                        End If
                    Else
                        MsgBox "Incorrect password."
                        Exit Sub
                    End If
                    Else
                        If txt_new_password.Text = txt_confirm_password.Text Then
                             ans = MsgBox("Are you sure you want to update user account?", vbYesNo, "Update User Account")
                                If ans = vbNo Then
                                    Exit Sub
                                Else
                            sql_string = "UPDATE tbl_user SET ID='" & txt_id.Text & "', Username='" & txt_username.Text & "', Lastname='" & txt_lastname.Text & "', Firstname='" & txt_firstname.Text & "', Middlename = '" & txt_middlename.Text & "', Password='" & txt_new_password.Text & "' WHERE Username='" & txt_oldusername.Text & "'"
                            Call mysql_select(usereditform.rs_user, sql_string)
                            sql_string = "UPDATE tbl_security SET Username='" & txt_username.Text & "' WHERE Username='" & txt_oldusername.Text & "'"
                            Call mysql_select(usereditform.rs_user, sql_string)
                            MsgBox "User account successfully updated."
                            Call userform.Form_Load
                            End If
                        Else
                            MsgBox "New password did not match."
                        End If
                End If
                
                End If
            Else
                    Call mysql_select(public_rs, "SELECT * FROM tbl_user WHERE Username = '" & txt_oldusername.Text & "'")
                    PASSWORD = public_rs.Fields("Password")
                    If txt_usertype.Text = "Administrator" Then
                    If PASSWORD = txt_old_password.Text Then
                        If txt_new_password.Text = txt_confirm_password.Text Then
                             ans = MsgBox("Are you sure you want to update user account?", vbYesNo, "Update User Account")
                                If ans = vbNo Then
                                    Exit Sub
                                Else
                            sql_string = "UPDATE tbl_user SET ID='" & txt_id.Text & "', Lastname='" & txt_lastname.Text & "', Firstname='" & txt_firstname.Text & "', Middlename = '" & txt_middlename.Text & "' , Password='" & txt_new_password.Text & "'WHERE Username='" & txt_oldusername.Text & "'"
                            Call mysql_select(usereditform.rs_user, sql_string)
                            MsgBox "User account successfully updated."
                            Call userform.Form_Load
                            End If
                        Else
                            MsgBox "New password did not match."
                        End If
                    Else
                        MsgBox "Incorrect password."
                        Exit Sub
                    End If
                Else
                    If txt_new_password.Text = txt_confirm_password.Text Then
                             ans = MsgBox("Are you sure you want to update user account?", vbYesNo, "Update User Account")
                                If ans = vbNo Then
                                    Exit Sub
                                Else
                            sql_string = "UPDATE tbl_user SET ID='" & txt_id.Text & "', Lastname='" & txt_lastname.Text & "', Firstname='" & txt_firstname.Text & "', Middlename = '" & txt_middlename.Text & "' , Password='" & txt_new_password.Text & "'WHERE Username='" & txt_oldusername.Text & "'"
                            Call mysql_select(usereditform.rs_user, sql_string)
                            MsgBox "User account successfully updated."
                            Call userform.Form_Load
                            End If
                        Else
                            MsgBox "New password did not match."
                        End If
            End If
            End If
        End If
    End If
    Unload Me
End Sub

Private Sub lbl_default_password_Click()
     lastname = Replace(txt_lastname.Text, " ", "")
     txt_old_password.Text = lastname
     txt_new_password.Text = lastname
     txt_confirm_password.Text = lastname
End Sub

Private Sub lbl_default_username_Click()
     lastname = Replace(txt_lastname.Text, " ", "")
     USERNAME = txt_id.Text & lastname
     txt_username.Text = USERNAME
End Sub
Private Sub load_user()
      Call set_datagrid(userform.dg_users, rs_user, _
                                        "SELECT " _
                                            & "Usertype as User_Type, Username FROM tbl_user ORDER BY Usertype ASC")
                                        
                    
                                       
        Unload Me
End Sub

Private Sub txt_firstname_KeyUp(KeyCode As Integer, Shift As Integer)
     If IsNumeric(Right(txt_firstname.Text, 1)) = True Then
        txt_firstname.Text = ""
        MsgBox "Number is not allowed."
    End If
End Sub

Private Sub txt_lastname_KeyUp(KeyCode As Integer, Shift As Integer)
     If IsNumeric(Right(txt_lastname.Text, 1)) = True Then
        txt_lastname.Text = ""
        MsgBox "Number is not allowed."
    End If
End Sub

Private Sub txt_middlename_KeyUp(KeyCode As Integer, Shift As Integer)
     If IsNumeric(Right(txt_middlename.Text, 1)) = True Then
        txt_middlename.Text = ""
        MsgBox "Number is not allowed."
    End If
End Sub
