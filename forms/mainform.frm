VERSION 5.00
Begin VB.Form mainform 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form 137 and Promotion Report Generation System of Manuel S. Rojas Elementary School"
   ClientHeight    =   8850
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15135
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8850
   ScaleWidth      =   15135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmr_datetime 
      Interval        =   1000
      Left            =   10320
      Top             =   3480
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00808080&
      Height          =   1215
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   15135
      Begin VB.CommandButton toolbar_promotion 
         Height          =   975
         Left            =   6240
         Picture         =   "mainform.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "View promotion summary"
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton toolbar_help 
         Height          =   975
         Left            =   11805
         Picture         =   "mainform.frx":12C2
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "User Guide"
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton toolbar_logout 
         Height          =   975
         Left            =   13185
         Picture         =   "mainform.frx":21B1
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Log out from this application"
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton toolbar_database 
         Height          =   975
         Left            =   10440
         Picture         =   "mainform.frx":3491
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Back-up or restore database"
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton toolbar_report 
         Height          =   975
         Left            =   9000
         Picture         =   "mainform.frx":499A
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Generate reports"
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton toolbar_form137 
         Height          =   975
         Left            =   7620
         Picture         =   "mainform.frx":5D07
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "View Form-137 of student"
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton toolbar_user 
         Height          =   975
         Left            =   4920
         Picture         =   "mainform.frx":6C9D
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Add and update user account"
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton toolbar_student 
         Height          =   975
         Left            =   3540
         Picture         =   "mainform.frx":816C
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Add and update student's information"
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton toolbar_teacher 
         Height          =   975
         Left            =   2160
         Picture         =   "mainform.frx":95BF
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Add and update teacher's information"
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton toolbar_level 
         Height          =   975
         Left            =   840
         Picture         =   "mainform.frx":AB61
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Add and update level"
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton toolbar_sy 
         Height          =   975
         Left            =   0
         Picture         =   "mainform.frx":BF9C
         Style           =   1  'Graphical
         TabIndex        =   0
         ToolTipText     =   "Set School Year"
         Top             =   1200
         Visible         =   0   'False
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      Height          =   615
      Left            =   0
      TabIndex        =   10
      Top             =   8280
      Width           =   15135
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Date and Time:"
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
         Left            =   9360
         TabIndex        =   17
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "School Year:"
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
         Left            =   4320
         TabIndex        =   16
         Top             =   240
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Logged as:"
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
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lbl_sy_tmp 
         BackStyle       =   0  'Transparent
         Caption         =   "School Year: 2013 - 2014"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5760
         TabIndex        =   14
         Top             =   240
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.Label lbl_datetime 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   11160
         TabIndex        =   13
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label lbl_username 
         BackStyle       =   0  'Transparent
         Caption         =   "Logged as: Admin"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1440
         TabIndex        =   12
         Top             =   240
         Width           =   2775
      End
   End
   Begin VB.Image Image1 
      Height          =   7095
      Left            =   720
      Picture         =   "mainform.frx":D5D5
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   13575
   End
End
Attribute VB_Name = "mainform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sql_string As String
Private Sub Timer1_Timer()

End Sub

Dim sql_string As String

Private Sub Form_Unload(Cancel As Integer)
    Dim ans As String
    ans = MsgBox("Are you sure you want to log-out from this application?", vbOKCancel, "Log Out")
                    If ans = vbCancel Then
                        Cancel = 1
                    Else
    sql_string = "UPDATE tbl_logs SET Logout='" & Now & "' WHERE Username='" & mainform.lbl_username.Caption & "'AND Logout='None'"
    Call mysql_select(usereditform.rs_user, sql_string)
    'MsgBox "Thank you for using this application."
    
    loginform.txt_username.Text = ""
    loginform.txt_password.Text = ""
    Call load_form(loginform, True)
    log = 3
   End If
End Sub

Private Sub tmr_datetime_Timer()
    lbl_datetime.Caption = Now
End Sub

Private Sub toolbar_database_Click()
   
    Call load_form(databaseform, True)
End Sub

Private Sub toolbar_form137_Click()
     Call load_form(form137allform, True)
End Sub

Private Sub toolbar_help_Click()
    Call load_form(helpform, True)
End Sub

Private Sub toolbar_level_Click()
    Call load_form(levelform, True)
End Sub

Private Sub toolbar_logout_Click()
    Unload Me
End Sub

Private Sub toolbar_promotion_Click()
  studentPromotion.isAdminMode = True
  Call load_form(studentPromotion, True)
End Sub

Private Sub toolbar_report_Click()
     Call load_form(reportsform, True)
End Sub

Private Sub toolbar_student_Click()
    Call load_form(studentform, True)
End Sub

Private Sub toolbar_sy_Click()
    Call load_form(syform, True)
End Sub

Private Sub toolbar_teacher_Click()
    Call load_form(teacherform, True)
End Sub

Private Sub toolbar_user_Click()
    Call load_form(userform, True)
End Sub
