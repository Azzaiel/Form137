VERSION 5.00
Begin VB.Form loginform 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Log In Form"
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4470
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "loginform.frx":0000
   ScaleHeight     =   3000
   ScaleWidth      =   4470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd_cancel 
      Height          =   615
      Left            =   2280
      Picture         =   "loginform.frx":977E
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton cmd_login 
      Default         =   -1  'True
      Height          =   615
      Left            =   960
      Picture         =   "loginform.frx":A4F9
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2160
      Width           =   1095
   End
   Begin VB.TextBox txt_password 
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
      Left            =   1560
      MaxLength       =   35
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1320
      Width           =   2775
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
      Left            =   1560
      LinkTimeout     =   30
      MaxLength       =   20
      TabIndex        =   0
      Top             =   600
      Width           =   2775
   End
   Begin VB.Label Label2 
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
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label Label1 
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
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   1935
   End
End
Attribute VB_Name = "loginform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sql_string As String

Private Sub cmd_cancel_Click()
    txt_username.Text = ""
    txt_password.Text = ""
    txt_username.SetFocus
End Sub

Private Sub cmd_login_Click()
    
    Call login
    
End Sub
Private Sub login()
    If txt_username.Text = "" Or txt_password.Text = "" Then
        MsgBox "Please input your username and password."
        Exit Sub
    End If
    If log > 0 Then
        Call mysql_select(public_rs, "SELECT * FROM tbl_user WHERE Username ='" & txt_username.Text & "' AND Password = '" & txt_password.Text & "' ")
        If public_rs.RecordCount = 0 Then
            log = log - 1
            If log <> 0 Then
                MsgBox "You have entered the wrong username or password. You have " & log & " chance(s) to access your account. Please input your correct username and password."
                Exit Sub
            End If
            If log = 0 Then
                MsgBox "You have reached the maximum number of tries in accessing your account. Please answer all the security questions correctly."
                Call load_form(securityform, True)
                Exit Sub
            End If
    Else
        user = txt_username.Text
        Dim usertype As String
        Call mysql_select(public_rs, "SELECT *" _
                                    & "FROM tbl_user " _
                                    & "WHERE Username = '" & txt_username.Text & "' " _
                                    & "AND Password= '" & txt_password.Text & "'")
        If public_rs.RecordCount = 0 Then
            MsgBox "Incorrect username or password!"
        Else
            
            user_type = public_rs.Fields("Usertype").value
            
            user_name = public_rs.Fields("Username").value
            user_password = public_rs.Fields("Password").value
'            MsgBox "Welcome " & user_name & " to Form 137 and Promotion Report Generation System of Manuel S. Rojas Elementary School."
        
            Call mysql_select(public_rs, "SELECT * FROM tbl_sy WHERE SY = " & Format(Date, "yyyy"))
            If public_rs.RecordCount = 0 Then
                Dim new_sy As String
                new_sy = Format(Date, "yyyy")
                Call mysql_select(public_rs, "INSERT INTO tbl_sy (SY) VALUES ( " & new_sy & ")")
            End If
            
            If Month(Now) <= 4 Then
                school_year = Left(Format(Date, "yyyy"), 3) & Trim(Str(val(Right(Format(Date, "yyyy"), 1) - 1)) & "-" & Format(Date, "yyyy"))
            Else
                school_year = Format(Date, "yyyy") & "-" & Left(Format(Date, "yyyy"), 3) & Trim(Str(val(Right(Format(Date, "yyyy"), 1) + 1)))
            End If
            mainform.lbl_username.Caption = user_name
            'mainform.lbl_sy.Caption = school_year
            mainteacherform.lbl_username.Caption = user_name
            'mainteacherform.lbl_sy.Caption = school_year
            If user_type = "Administrator" Then
                Call load_form(mainform, True)
                mainform.lbl_username.Caption = user_name
                'mainform.lbl_sy.Caption = school_year
                sql_string = "INSERT INTO " _
                                & "tbl_logs (Username, Login,Logout)" _
                            & " VALUES (" _
                                & "'" & user & "','" _
                                & Now & "','None')"
                Call mysql_select(useraccountform.rs_user, sql_string)
                
            Else
                Dim tmp_sy As Integer
                tmp_sy = val(Format(Date, "yyyy"))
                mainteacherform.cmb_sy.Clear
                mainteacherform.cmb_sy.AddItem ((tmp_sy - 1) & "-" & tmp_sy)
                mainteacherform.cmb_sy.AddItem (tmp_sy & "-" & (tmp_sy + 1))
                If Month(Now) <= 4 Then
                  mainteacherform.cmb_sy.ListIndex = 0
                Else
                  mainteacherform.cmb_sy.ListIndex = 1
                End If
               Call load_form(mainteacherform, True)
               
                mainteacherform.lbl_username.Caption = user_name
                'mainteacherform.lbl_sy.Caption = school_year
                sql_string = "INSERT INTO " _
                                & "tbl_logs (Username, Login,Logout)" _
                            & " VALUES (" _
                                & "'" & user & "','" _
                                & Now & "','None')"
                Call mysql_select(useraccountform.rs_user, sql_string)
            End If
            
            End If
    End If
End If
End Sub

Private Sub Form_Activate()
    txt_username.SetFocus
End Sub

Private Sub Form_Load()
   
     
    Call connect_db
    log = 3
End Sub
