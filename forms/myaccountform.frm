VERSION 5.00
Begin VB.Form myaccountform 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "My Account"
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5985
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "myaccountform.frx":0000
   ScaleHeight     =   4455
   ScaleWidth      =   5985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd_save 
      Height          =   615
      Left            =   3120
      Picture         =   "myaccountform.frx":12121
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3360
      Width           =   1095
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
      Left            =   2400
      MaxLength       =   35
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1800
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
      Left            =   2400
      MaxLength       =   35
      TabIndex        =   0
      Top             =   1320
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
      Left            =   2400
      MaxLength       =   35
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   2280
      Width           =   3495
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
      Left            =   2400
      MaxLength       =   35
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   2760
      Width           =   3495
   End
   Begin VB.TextBox txt_oldusername 
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   3480
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Click for security questions."
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
      Left            =   3120
      TabIndex        =   7
      Top             =   4200
      Width           =   2895
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Old Password:"
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
      Left            =   240
      TabIndex        =   12
      Top             =   1920
      Width           =   1935
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Username:"
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
      Left            =   240
      TabIndex        =   11
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label Label8 
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
      Left            =   240
      TabIndex        =   10
      Top             =   2400
      Width           =   2055
   End
   Begin VB.Label Label9 
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
      Left            =   240
      TabIndex        =   9
      Top             =   2880
      Width           =   2055
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
      Left            =   240
      TabIndex        =   5
      Top             =   360
      Width           =   2895
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
      Left            =   240
      TabIndex        =   6
      Top             =   720
      Width           =   2895
   End
End
Attribute VB_Name = "myaccountform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim id, lastname, USERNAME, PASSWORD As String
Dim sql_string As String
Public user_name As String

Private Sub cmb_clear_Click()

End Sub

Private Sub cmd_save_Click()
Dim ans As String
If txt_old_password.Text = "" And txt_new_password.Text = "" And txt_confirm_password.Text = "" Then
If txt_username.Text <> txt_oldusername.Text Then
                If is_duplicate("tbl_user", "Username", txt_username.Text) Then
                    MsgBox "Username already exists."
                    Exit Sub
                Else
                     ans = MsgBox("Are you sure you want to update your account?", vbYesNo, "My Account")
                    If ans = vbNo Then
                        Exit Sub
                    Else
                    sql_string = "UPDATE tbl_user SET Username='" & txt_username.Text & "' WHERE Username='" & txt_oldusername.Text & "'"
                    Call mysql_select(usereditform.rs_user, sql_string)
                    MsgBox "User account successfully updated."
                    mainteacherform.lbl_username.Caption = txt_username.Text
                    End If
                End If
            Else
                MsgBox "Nothing to edit."
                Exit Sub
            End If
        Else
    If txt_username.Text <> txt_oldusername.Text Then
                If is_duplicate("tbl_user", "Username", txt_username.Text) Then
                    MsgBox "Username already exists."
                    Exit Sub
                Else
                    Call mysql_select(public_rs, "SELECT * FROM tbl_user WHERE Username = '" & txt_oldusername.Text & "'")
                    PASSWORD = public_rs.Fields("Password")
                    If PASSWORD = txt_old_password.Text Then
                        If txt_new_password.Text = txt_confirm_password.Text Then
                             ans = MsgBox("Are you sure you want to update your account?", vbYesNo, "My Account")
                            If ans = vbNo Then
                                Exit Sub
                            Else
                            sql_string = "UPDATE tbl_user SET  Username='" & txt_username.Text & "', Password='" & txt_new_password.Text & "' WHERE Username='" & txt_oldusername.Text & "'"
                            Call mysql_select(usereditform.rs_user, sql_string)
                            MsgBox "User account updated."
                            mainteacherform.lbl_username.Caption = txt_username.Text
                            End If
                        Else
                            MsgBox "New password did not match."
                        End If
                    Else
                        MsgBox "Incorrect password."
                        Exit Sub
                    End If
                End If
            Else
                    Call mysql_select(public_rs, "SELECT * FROM tbl_user WHERE Username = '" & txt_oldusername.Text & "'")
                    PASSWORD = public_rs.Fields("Password")
                    If PASSWORD = txt_old_password.Text Then
                        If txt_new_password.Text = txt_confirm_password.Text Then
                             ans = MsgBox("Are you sure you want to update your account?", vbYesNo, "My Account")
                                If ans = vbNo Then
                                    Exit Sub
                                Else
                            sql_string = "UPDATE tbl_user SET Password='" & txt_new_password.Text & "'WHERE Username='" & txt_oldusername.Text & "'"
                            Call mysql_select(usereditform.rs_user, sql_string)
                            MsgBox "User account updated."
                            mainteacherform.lbl_username.Caption = txt_username.Text
                            End If
                        Else
                            MsgBox "New password did not match."
                        End If
                    Else
                        MsgBox "Incorrect password."
                        Exit Sub
                    End If
            End If
        End If
End Sub

Private Sub Label1_Click()
     Call mysql_select(public_rs, "SELECT * FROM tbl_security WHERE Username = '" & user & "'")
     If public_rs.RecordCount = 0 Then
        securityquestions.txt_pet.Text = ""
        securityquestions.txt_place.Text = ""
        securityquestions.txt_author.Text = ""
    Else
        securityquestions.txt_pet.Text = public_rs.Fields("Pet").Value
        securityquestions.txt_place.Text = public_rs.Fields("Place").Value
        securityquestions.txt_author.Text = public_rs.Fields("Author").Value
    End If
     Call load_form(securityquestions, True)
End Sub

Private Sub lbl_default_password_Click()
     Call mysql_select(public_rs, "SELECT * FROM tbl_user WHERE Username = '" & mainteacherform.lbl_username.Caption & "'")
     id = public_rs.Fields("ID")
     Call mysql_select(public_rs, "SELECT * FROM tbl_teacher WHERE teacher_id = '" & id & "'")
     lastname = Replace(public_rs.Fields("last_name"), " ", "")
     txt_old_password.Text = lastname
     txt_new_password.Text = lastname
     txt_confirm_password.Text = lastname
End Sub

Private Sub lbl_default_username_Click()
     Call mysql_select(public_rs, "SELECT * FROM tbl_user WHERE Username = '" & mainteacherform.lbl_username.Caption & "'")
     id = public_rs.Fields("ID")
     Call mysql_select(public_rs, "SELECT * FROM tbl_teacher WHERE teacher_id = '" & id & "'")
     lastname = Replace(public_rs.Fields("last_name"), " ", "")
     USERNAME = id & lastname
     txt_username.Text = USERNAME
End Sub
