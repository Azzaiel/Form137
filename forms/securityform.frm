VERSION 5.00
Begin VB.Form securityform 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Security Questions"
   ClientHeight    =   5745
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5685
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "securityform.frx":0000
   ScaleHeight     =   5745
   ScaleWidth      =   5685
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txt_id 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
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
      Left            =   1920
      MaxLength       =   30
      TabIndex        =   0
      Top             =   0
      Width           =   3375
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Choose Security Question"
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
      Height          =   5175
      Left            =   0
      TabIndex        =   8
      Top             =   480
      Width           =   5655
      Begin VB.OptionButton opt_3 
         BackColor       =   &H00808080&
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
         Height          =   375
         Left            =   480
         TabIndex        =   5
         Top             =   3240
         Width           =   4815
      End
      Begin VB.TextBox txt_3 
         Alignment       =   2  'Center
         Enabled         =   0   'False
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
         IMEMode         =   3  'DISABLE
         Left            =   480
         MaxLength       =   30
         PasswordChar    =   "*"
         TabIndex        =   6
         Top             =   3840
         Width           =   4815
      End
      Begin VB.OptionButton opt_2 
         BackColor       =   &H00808080&
         Caption         =   "What is your favorite place?"
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
         Height          =   375
         Left            =   480
         TabIndex        =   3
         Top             =   1920
         Width           =   4815
      End
      Begin VB.TextBox txt_2 
         Alignment       =   2  'Center
         Enabled         =   0   'False
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
         IMEMode         =   3  'DISABLE
         Left            =   480
         MaxLength       =   30
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   2520
         Width           =   4815
      End
      Begin VB.OptionButton opt_1 
         BackColor       =   &H00808080&
         Caption         =   "What is your favorite pet's name?"
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
         Height          =   375
         Left            =   480
         TabIndex        =   1
         Top             =   600
         Width           =   4815
      End
      Begin VB.TextBox txt_1 
         Alignment       =   2  'Center
         Enabled         =   0   'False
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
         IMEMode         =   3  'DISABLE
         Left            =   480
         MaxLength       =   30
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1200
         Width           =   4815
      End
      Begin VB.CommandButton cmd_ok 
         Default         =   -1  'True
         Height          =   615
         Left            =   2160
         Picture         =   "securityform.frx":1DD3
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   4440
         Width           =   1095
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Input your ID:"
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
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   0
      Width           =   1815
   End
End
Attribute VB_Name = "securityform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ctr As Integer
Public sql_string As String
Private Sub cmd_ok_Click()
    If txt_id.Text = "" Then
        MsgBox "Please input your ID number."
        Exit Sub
    End If
    Call mysql_select(public_rs, "SELECT * FROM tbl_user WHERE ID = '" & txt_id.Text & "'")
    If public_rs.RecordCount = 0 Then
        MsgBox "Unable to find your ID number."
        Exit Sub
    Else
        user = public_rs.Fields("Username").value
    End If
    If txt_1.Text = "" And txt_2.Text = "" And txt_3.Text = "" Then
        MsgBox "Please input your answer."
       Exit Sub
    End If
    If txt_1.Text <> "" Then
    Call mysql_select(public_rs, "SELECT * FROM tbl_security WHERE Username = '" & user & "' AND Pet='" & txt_1.Text & "'")
If public_rs.RecordCount = 0 Then
            ctr = ctr - 1
    If ctr <> 0 Then
          MsgBox "You have entered the wrong answer for Question No.1. You have " & ctr & " chance to answer this question."
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
            'MsgBox "Welcome " & user_name & " to Form 137 and Promotion Report Generation System of Manuel S. Rojas Elementary School."
             
        Call mysql_select(public_rs, "SELECT *" _
                                    & "FROM tbl_user " _
                                    & "WHERE Username = '" & user & "'")
           user_type = public_rs.Fields("Usertype").value
            
            user_name = public_rs.Fields("Username").value
            user_password = public_rs.Fields("Password").value
            Call mysql_select(public_rs, "SELECT * FROM tbl_sy WHERE SY = " & Format(Date, "yyyy"))
            If public_rs.RecordCount = 0 Then
                Call mysql_select(public_rs, "INSERT INTO tbl_sy (SY) VALUES ( " & Format(Date, "yyyy") & ")")
            End If
            school_year = Format(Date, "yyyy") & "-" & Left(Format(Date, "yyyy"), 3) & Trim(Str(val(Right(Format(Date, "yyyy"), 1) + 1)))
            
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
             tmp_sy = val(Format(Date, "yyyy"))
                mainteacherform.cmb_sy.Clear
                mainteacherform.cmb_sy.AddItem ((tmp_sy - 1) & "-" & tmp_sy)
                mainteacherform.cmb_sy.AddItem (tmp_sy & "-" & (tmp_sy + 1))
                If Month(Now) <= 4 Then
                  mainteacherform.cmb_sy.ListIndex = 0
                Else
                  mainteacherform.cmb_sy.ListIndex = 1
                End If
                 tmp_sy = val(Format(Date, "yyyy"))
                mainteacherform.cmb_sy.Clear
                mainteacherform.cmb_sy.AddItem ((tmp_sy - 1) & "-" & tmp_sy)
                mainteacherform.cmb_sy.AddItem (tmp_sy & "-" & (tmp_sy + 1))
                If Month(Now) <= 4 Then
                  mainteacherform.cmb_sy.ListIndex = 0
                Else
                  mainteacherform.cmb_sy.ListIndex = 1
                End If
                 tmp_sy = val(Format(Date, "yyyy"))
                mainteacherform.cmb_sy.Clear
                mainteacherform.cmb_sy.AddItem ((tmp_sy - 1) & "-" & tmp_sy)
                mainteacherform.cmb_sy.AddItem (tmp_sy & "-" & (tmp_sy + 1))
                If Month(Now) <= 4 Then
                  mainteacherform.cmb_sy.ListIndex = 0
                Else
                  mainteacherform.cmb_sy.ListIndex = 1
                End If
                 tmp_sy = val(Format(Date, "yyyy"))
                mainteacherform.cmb_sy.Clear
                mainteacherform.cmb_sy.AddItem ((tmp_sy - 1) & "-" & tmp_sy)
                mainteacherform.cmb_sy.AddItem (tmp_sy & "-" & (tmp_sy + 1))
                If Month(Now) <= 4 Then
                  mainteacherform.cmb_sy.ListIndex = 0
                Else
                  mainteacherform.cmb_sy.ListIndex = 1
                End If
                 tmp_sy = val(Format(Date, "yyyy"))
                mainteacherform.cmb_sy.Clear
                mainteacherform.cmb_sy.AddItem ((tmp_sy - 1) & "-" & tmp_sy)
                mainteacherform.cmb_sy.AddItem (tmp_sy & "-" & (tmp_sy + 1))
                If Month(Now) <= 4 Then
                  mainteacherform.cmb_sy.ListIndex = 0
                Else
                  mainteacherform.cmb_sy.ListIndex = 1
                End If
                 tmp_sy = val(Format(Date, "yyyy"))
                mainteacherform.cmb_sy.Clear
                mainteacherform.cmb_sy.AddItem ((tmp_sy - 1) & "-" & tmp_sy)
                mainteacherform.cmb_sy.AddItem (tmp_sy & "-" & (tmp_sy + 1))
                If Month(Now) <= 4 Then
                  mainteacherform.cmb_sy.ListIndex = 0
                Else
                  mainteacherform.cmb_sy.ListIndex = 1
                End If
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
          Unload Me
          
    
End If
End If
If txt_2.Text <> "" Then
    Call mysql_select(public_rs, "SELECT * FROM tbl_security WHERE Username = '" & user & "' AND Place='" & txt_2.Text & "'")
If public_rs.RecordCount = 0 Then
            ctr = ctr - 1
    If ctr <> 0 Then
         MsgBox "You have entered the wrong answer for Question No.2. You have " & ctr & " chance to answer this question."
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
           'MsgBox "Welcome " & user_name & " to Form 137 and Promotion Report Generation System of Manuel S. Rojas Elementary School."
             
        Call mysql_select(public_rs, "SELECT *" _
                                    & "FROM tbl_user " _
                                    & "WHERE Username = '" & user & "'")
           user_type = public_rs.Fields("Usertype").value
            
            user_name = public_rs.Fields("Username").value
            user_password = public_rs.Fields("Password").value
            Call mysql_select(public_rs, "SELECT * FROM tbl_sy WHERE SY = " & Format(Date, "yyyy"))
            If public_rs.RecordCount = 0 Then
                Call mysql_select(public_rs, "INSERT INTO tbl_sy (SY) VALUES ( " & Format(Date, "yyyy") & ")")
            End If
            school_year = Format(Date, "yyyy") & "-" & Left(Format(Date, "yyyy"), 3) & Trim(Str(val(Right(Format(Date, "yyyy"), 1) + 1)))
            
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
                tmp_sy = val(Format(Date, "yyyy"))
                mainteacherform.cmb_sy.Clear
                mainteacherform.cmb_sy.AddItem ((tmp_sy - 1) & "-" & tmp_sy)
                mainteacherform.cmb_sy.AddItem (tmp_sy & "-" & (tmp_sy + 1))
                If Month(Now) <= 4 Then
                  mainteacherform.cmb_sy.ListIndex = 0
                Else
                  mainteacherform.cmb_sy.ListIndex = 1
                End If
                 tmp_sy = val(Format(Date, "yyyy"))
                mainteacherform.cmb_sy.Clear
                mainteacherform.cmb_sy.AddItem ((tmp_sy - 1) & "-" & tmp_sy)
                mainteacherform.cmb_sy.AddItem (tmp_sy & "-" & (tmp_sy + 1))
                If Month(Now) <= 4 Then
                  mainteacherform.cmb_sy.ListIndex = 0
                Else
                  mainteacherform.cmb_sy.ListIndex = 1
                End If
                 tmp_sy = val(Format(Date, "yyyy"))
                mainteacherform.cmb_sy.Clear
                mainteacherform.cmb_sy.AddItem ((tmp_sy - 1) & "-" & tmp_sy)
                mainteacherform.cmb_sy.AddItem (tmp_sy & "-" & (tmp_sy + 1))
                If Month(Now) <= 4 Then
                  mainteacherform.cmb_sy.ListIndex = 0
                Else
                  mainteacherform.cmb_sy.ListIndex = 1
                End If
                 tmp_sy = val(Format(Date, "yyyy"))
                mainteacherform.cmb_sy.Clear
                mainteacherform.cmb_sy.AddItem ((tmp_sy - 1) & "-" & tmp_sy)
                mainteacherform.cmb_sy.AddItem (tmp_sy & "-" & (tmp_sy + 1))
                If Month(Now) <= 4 Then
                  mainteacherform.cmb_sy.ListIndex = 0
                Else
                  mainteacherform.cmb_sy.ListIndex = 1
                End If
                 tmp_sy = val(Format(Date, "yyyy"))
                mainteacherform.cmb_sy.Clear
                mainteacherform.cmb_sy.AddItem ((tmp_sy - 1) & "-" & tmp_sy)
                mainteacherform.cmb_sy.AddItem (tmp_sy & "-" & (tmp_sy + 1))
                If Month(Now) <= 4 Then
                  mainteacherform.cmb_sy.ListIndex = 0
                Else
                  mainteacherform.cmb_sy.ListIndex = 1
                End If
                 tmp_sy = val(Format(Date, "yyyy"))
                mainteacherform.cmb_sy.Clear
                mainteacherform.cmb_sy.AddItem ((tmp_sy - 1) & "-" & tmp_sy)
                mainteacherform.cmb_sy.AddItem (tmp_sy & "-" & (tmp_sy + 1))
                If Month(Now) <= 4 Then
                  mainteacherform.cmb_sy.ListIndex = 0
                Else
                  mainteacherform.cmb_sy.ListIndex = 1
                End If
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
          Unload Me
    
End If
End If
If txt_3.Text <> "" Then
    Call mysql_select(public_rs, "SELECT * FROM tbl_security WHERE Username = '" & user & "' AND Author='" & txt_3.Text & "'")
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
            'MsgBox "Welcome " & user_name & " to Form 137 and Promotion Report Generation System of Manuel S. Rojas Elementary School."
             
        Call mysql_select(public_rs, "SELECT *" _
                                    & "FROM tbl_user " _
                                    & "WHERE Username = '" & user & "'")
           user_type = public_rs.Fields("Usertype").value
            
            user_name = public_rs.Fields("Username").value
            user_password = public_rs.Fields("Password").value
            Call mysql_select(public_rs, "SELECT * FROM tbl_sy WHERE SY = " & Format(Date, "yyyy"))
            If public_rs.RecordCount = 0 Then
                Call mysql_select(public_rs, "INSERT INTO tbl_sy (SY) VALUES ( " & Format(Date, "yyyy") & ")")
            End If
            school_year = Format(Date, "yyyy") & "-" & Left(Format(Date, "yyyy"), 3) & Trim(Str(val(Right(Format(Date, "yyyy"), 1) + 1)))
            
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
             tmp_sy = val(Format(Date, "yyyy"))
                mainteacherform.cmb_sy.Clear
                mainteacherform.cmb_sy.AddItem ((tmp_sy - 1) & "-" & tmp_sy)
                mainteacherform.cmb_sy.AddItem (tmp_sy & "-" & (tmp_sy + 1))
                If Month(Now) <= 4 Then
                  mainteacherform.cmb_sy.ListIndex = 0
                Else
                  mainteacherform.cmb_sy.ListIndex = 1
                End If
                 tmp_sy = val(Format(Date, "yyyy"))
                mainteacherform.cmb_sy.Clear
                mainteacherform.cmb_sy.AddItem ((tmp_sy - 1) & "-" & tmp_sy)
                mainteacherform.cmb_sy.AddItem (tmp_sy & "-" & (tmp_sy + 1))
                If Month(Now) <= 4 Then
                  mainteacherform.cmb_sy.ListIndex = 0
                Else
                  mainteacherform.cmb_sy.ListIndex = 1
                End If
                 tmp_sy = val(Format(Date, "yyyy"))
                mainteacherform.cmb_sy.Clear
                mainteacherform.cmb_sy.AddItem ((tmp_sy - 1) & "-" & tmp_sy)
                mainteacherform.cmb_sy.AddItem (tmp_sy & "-" & (tmp_sy + 1))
                If Month(Now) <= 4 Then
                  mainteacherform.cmb_sy.ListIndex = 0
                Else
                  mainteacherform.cmb_sy.ListIndex = 1
                End If
                 tmp_sy = val(Format(Date, "yyyy"))
                mainteacherform.cmb_sy.Clear
                mainteacherform.cmb_sy.AddItem ((tmp_sy - 1) & "-" & tmp_sy)
                mainteacherform.cmb_sy.AddItem (tmp_sy & "-" & (tmp_sy + 1))
                If Month(Now) <= 4 Then
                  mainteacherform.cmb_sy.ListIndex = 0
                Else
                  mainteacherform.cmb_sy.ListIndex = 1
                End If
                 tmp_sy = val(Format(Date, "yyyy"))
                mainteacherform.cmb_sy.Clear
                mainteacherform.cmb_sy.AddItem ((tmp_sy - 1) & "-" & tmp_sy)
                mainteacherform.cmb_sy.AddItem (tmp_sy & "-" & (tmp_sy + 1))
                If Month(Now) <= 4 Then
                  mainteacherform.cmb_sy.ListIndex = 0
                Else
                  mainteacherform.cmb_sy.ListIndex = 1
                End If
                 tmp_sy = val(Format(Date, "yyyy"))
                mainteacherform.cmb_sy.Clear
                mainteacherform.cmb_sy.AddItem ((tmp_sy - 1) & "-" & tmp_sy)
                mainteacherform.cmb_sy.AddItem (tmp_sy & "-" & (tmp_sy + 1))
                If Month(Now) <= 4 Then
                  mainteacherform.cmb_sy.ListIndex = 0
                Else
                  mainteacherform.cmb_sy.ListIndex = 1
                End If
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
          Unload Me
    
    End If
End If

    
End Sub

Private Sub Form_Load()
    ctr = 3
End Sub

Private Sub opt_1_Click()
    opt_2.value = False
    opt_3.value = False
    txt_1.Enabled = True
    txt_2.Enabled = False
    txt_3.Enabled = False
    txt_1.Text = ""
    txt_2.Text = ""
    txt_3.Text = ""
End Sub

Private Sub opt_2_Click()
    opt_1.value = False
    opt_3.value = False
    txt_1.Enabled = False
    txt_2.Enabled = True
    txt_3.Enabled = False
     txt_1.Text = ""
    txt_2.Text = ""
    txt_3.Text = ""
End Sub

Private Sub opt_3_Click()
    opt_1.value = False
    opt_2.value = False
    txt_1.Enabled = False
    txt_2.Enabled = False
    txt_3.Enabled = True
     txt_1.Text = ""
    txt_2.Text = ""
    txt_3.Text = ""
End Sub

