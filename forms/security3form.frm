VERSION 5.00
Begin VB.Form security3form 
   BorderStyle     =   0  'None
   Caption         =   "Security"
   ClientHeight    =   2400
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5655
   LinkTopic       =   "Form1"
   Picture         =   "security3form.frx":0000
   ScaleHeight     =   2400
   ScaleWidth      =   5655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Security Question No. 3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   2415
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   5655
      Begin VB.TextBox txt_author 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         MaxLength       =   35
         TabIndex        =   0
         Top             =   960
         Width           =   4935
      End
      Begin VB.CommandButton cmd_ok 
         Height          =   615
         Left            =   2160
         Picture         =   "security3form.frx":1DD3
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Who is your favorite author?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   480
         TabIndex        =   3
         Top             =   600
         Width           =   5415
      End
   End
End
Attribute VB_Name = "security3form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ctr As Integer
Dim sql_string As String

Private Sub cmd_ok_Click()
     If txt_author.Text = "" Then
        MsgBox "Please input your answer."
        Exit Sub
    End If
     Call mysql_select(public_rs, "SELECT * FROM tbl_security WHERE Username = '" & user & "' AND Author='" & txt_author.Text & "'")
If public_rs.RecordCount = 0 Then
        ctr = ctr - 1
    If ctr <> 0 Then
         MsgBox "You have entered the wrong answer for Question No.3. You have " & ctr & " chance to answer this question."
            Exit Sub
    End If
    If ctr = 0 Then
        MsgBox "You have reached the maximum number of tries for this question. Please contact your administrator to recover your account."
         loginform.txt_username.Text = ""
        loginform.txt_password.Text = ""
        Unload Me
    End If
    
       
         
           
        Else
            MsgBox "Your answer is correct."
             
        Call mysql_select(public_rs, "SELECT *" _
                                    & "FROM tbl_user " _
                                    & "WHERE Username = '" & user & "'")
           user_type = public_rs.Fields("Usertype").Value
            
            user_name = public_rs.Fields("Username").Value
            user_password = public_rs.Fields("Password").Value
            Call mysql_select(public_rs, "SELECT * FROM tbl_sy WHERE SY = " & Format(Date, "yyyy"))
            If public_rs.RecordCount = 0 Then
                Call mysql_select(public_rs, "INSERT INTO tbl_sy (SY) VALUES ( " & Format(Date, "yyyy") & ")")
            End If
            school_year = Format(Date, "yyyy") & "-" & Left(Format(Date, "yyyy"), 3) & Trim(Str(val(Right(Format(Date, "yyyy"), 1) + 1)))
            
            mainform.lbl_username.Caption = user_name
            mainform.lbl_sy.Caption = school_year
            mainteacherform.lbl_username.Caption = user_name
            mainteacherform.lbl_sy.Caption = school_year
            If user_type = "Administrator" Then
                Call load_form(mainform, True)
                mainform.lbl_username.Caption = user_name
                mainform.lbl_sy.Caption = school_year
                sql_string = "INSERT INTO " _
                                & "tbl_logs (Username, Login,Logout)" _
                            & " VALUES (" _
                                & "'" & user & "','" _
                                & Now & "','None')"
                Call mysql_select(useraccountform.rs_user, sql_string)
                
            Else
               Call load_form(mainteacherform, True)
                mainteacherform.lbl_username.Caption = user_name
                mainteacherform.lbl_sy.Caption = school_year
                sql_string = "INSERT INTO " _
                                & "tbl_logs (Username, Login,Logout)" _
                            & " VALUES (" _
                                & "'" & user & "','" _
                                & Now & "','None')"
                Call mysql_select(useraccountform.rs_user, sql_string)
                
            End If
          Unload Me
    
    End If
End Sub

Private Sub Form_Load()
    ctr = 2
End Sub
