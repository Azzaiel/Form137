VERSION 5.00
Begin VB.Form useraccountform 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "User Account"
   ClientHeight    =   4485
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "useraccountform.frx":0000
   ScaleHeight     =   4485
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd_cancel 
      Height          =   615
      Left            =   2880
      Picture         =   "useraccountform.frx":12121
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CommandButton cmd_save 
      Height          =   615
      Left            =   1680
      Picture         =   "useraccountform.frx":12E9C
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3840
      Width           =   1095
   End
   Begin VB.TextBox txt_password 
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
      IMEMode         =   3  'DISABLE
      Left            =   2160
      MaxLength       =   35
      PasswordChar    =   "*"
      TabIndex        =   6
      Top             =   3240
      Width           =   3495
   End
   Begin VB.TextBox txt_username 
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
      Left            =   2160
      MaxLength       =   35
      TabIndex        =   5
      Top             =   2760
      Width           =   3495
   End
   Begin VB.TextBox txt_usertype 
      BackColor       =   &H00C0C0C0&
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
      Left            =   2160
      TabIndex        =   4
      Top             =   2280
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
      Left            =   2160
      MaxLength       =   15
      TabIndex        =   0
      Top             =   360
      Width           =   2175
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
      Left            =   2160
      MaxLength       =   35
      TabIndex        =   1
      Top             =   840
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
      Left            =   2160
      MaxLength       =   35
      TabIndex        =   2
      Top             =   1320
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
      Left            =   2160
      MaxLength       =   35
      TabIndex        =   3
      Top             =   1800
      Width           =   3495
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
      Left            =   240
      TabIndex        =   16
      Top             =   0
      Width           =   7335
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "*Password:"
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
      Left            =   360
      TabIndex        =   15
      Top             =   3360
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
      Left            =   360
      TabIndex        =   14
      Top             =   2880
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
      Left            =   360
      TabIndex        =   13
      Top             =   2400
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
      Left            =   360
      TabIndex        =   12
      Top             =   480
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
      Left            =   360
      TabIndex        =   11
      Top             =   960
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
      Left            =   360
      TabIndex        =   10
      Top             =   1440
      Width           =   1455
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
      Left            =   360
      TabIndex        =   9
      Top             =   1920
      Width           =   1455
   End
End
Attribute VB_Name = "useraccountform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public rs_user As New ADODB.Recordset
Dim sql_string As String
Dim USERNAME, lastname As String
Private Sub cmd_cancel_Click()
    txt_id.Text = ""
    txt_firstname.Text = ""
    txt_lastname.Text = ""
    txt_middlename.Text = ""
    txt_usertype.Text = "Administrator"
    txt_username.Text = ""
    txt_password.Text = ""
End Sub

Private Sub cmd_save_Click()
    Dim ans As String
    If txt_id.Text = "" Or txt_lastname.Text = "" Or txt_firstname.Text = "" Or txt_username.Text = "" Or txt_password.Text = "" Then
        MsgBox "Please input required fields."
        Exit Sub
    Else
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
        If is_duplicate("tbl_user", "Username", txt_username.Text) Then
            MsgBox "Username already exists."
            Exit Sub
        End If
        If txt_password.Text = "" Then
            MsgBox "Password did not match."
            Exit Sub
        Else
            
            lastname = Replace(txt_lastname.Text, " ", "")
            USERNAME = txt_id.Text & lastname
            ans = MsgBox("Are you sure you want to add user account?", vbYesNo, "Add User Account")
                    If ans = vbNo Then
                        Exit Sub
                    Else
             sql_string = "INSERT INTO " _
                            & "tbl_user (ID, Usertype, Username, Password, Lastname, Firstname, Middlename)" _
                        & " VALUES (" _
                            & "'" & txt_id.Text & "','Administrator','" _
                            & USERNAME & "','" & lastname & "','" & txt_lastname.Text & "','" & txt_firstname.Text & "','" & txt_middlename.Text & "')"
            Call mysql_select(useraccountform.rs_user, sql_string)
            MsgBox "User account added."
            Call userform.Form_Load
            End If
        End If
    End If
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

Private Sub txt_id_KeyUp(KeyCode As Integer, Shift As Integer)
     lastname = Replace(txt_lastname.Text, " ", "")
     USERNAME = txt_id.Text & lastname
     txt_username.Text = USERNAME
     txt_password.Text = lastname
End Sub

Private Sub txt_lastname_KeyUp(KeyCode As Integer, Shift As Integer)
     If IsNumeric(Right(txt_lastname.Text, 1)) = True Then
        txt_lastname.Text = ""
        MsgBox "Number is not allowed."
        Exit Sub
    End If
     lastname = Replace(txt_lastname.Text, " ", "")
     USERNAME = txt_id.Text & lastname
     txt_username.Text = USERNAME
     txt_password.Text = lastname
End Sub

Private Sub txt_middlename_KeyUp(KeyCode As Integer, Shift As Integer)
     If IsNumeric(Right(txt_middlename.Text, 1)) = True Then
        txt_middlename.Text = ""
        MsgBox "Number is not allowed."
    End If
End Sub
