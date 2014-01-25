VERSION 5.00
Begin VB.Form mainteacherform 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form 137 and Student Promotion Report Generation System of Manuel S. Rojas Elementary School"
   ClientHeight    =   9120
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13200
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9120
   ScaleWidth      =   13200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   10200
      Top             =   5040
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      Height          =   615
      Left            =   0
      TabIndex        =   6
      Top             =   8520
      Width           =   13215
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
         Left            =   1200
         TabIndex        =   12
         Top             =   240
         Width           =   1935
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
         Left            =   10680
         TabIndex        =   11
         Top             =   240
         Width           =   2415
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
         Left            =   5880
         TabIndex        =   10
         Top             =   240
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Logged as:"
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
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "School Year:"
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
         Left            =   4680
         TabIndex        =   8
         Top             =   240
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Date and Time:"
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
         Left            =   9240
         TabIndex        =   7
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00808080&
      Height          =   1215
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   13215
      Begin VB.CommandButton toolbar_promotion 
         Height          =   975
         Left            =   1680
         Picture         =   "mainteacherform.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "View promotion summary"
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton toolbar_help 
         Height          =   975
         Left            =   4560
         Picture         =   "mainteacherform.frx":12C2
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "User Guide"
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton toolbar_advisories 
         Height          =   975
         Left            =   240
         Picture         =   "mainteacherform.frx":21B1
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "My Advisories"
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton toolbar_account 
         Height          =   975
         Left            =   3120
         Picture         =   "mainteacherform.frx":35EA
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "My Account"
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton toolbar_section 
         Height          =   975
         Left            =   120
         Picture         =   "mainteacherform.frx":483D
         Style           =   1  'Graphical
         TabIndex        =   0
         ToolTipText     =   "My Sections"
         Top             =   1320
         Width           =   1335
      End
      Begin VB.CommandButton toolbar_logout 
         Height          =   975
         Left            =   6000
         Picture         =   "mainteacherform.frx":5C8E
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Log out from this application"
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "This is the Main Form designed for teachers of Manuel S. Rojas Elementary School."
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
         Height          =   495
         Left            =   7800
         TabIndex        =   13
         Top             =   360
         Width           =   4935
      End
   End
   Begin VB.Image Image1 
      Height          =   7335
      Left            =   0
      Picture         =   "mainteacherform.frx":6F6E
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   13215
   End
End
Attribute VB_Name = "mainteacherform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sql_string As String

Private Sub Form_Unload(Cancel As Integer)
    Dim ans As String
    ans = MsgBox("Are you sure you want to log-out from this application?", vbOKCancel, "Log Out")
                    If ans = vbCancel Then
                        Cancel = 1
                    Else
     sql_string = "UPDATE tbl_logs SET Logout='" & Now & "' WHERE Username='" & mainteacherform.lbl_username.Caption & "'AND Logout='None'"
    Call mysql_select(usereditform.rs_user, sql_string)
    MsgBox "Thank you for using this application."
    
    loginform.txt_username.Text = ""
    loginform.txt_password.Text = ""
    Call load_form(loginform, True)
    log = 3
    End If
End Sub

Private Sub Timer1_Timer()
      lbl_datetime.Caption = Now
End Sub

Private Sub toolbar_account_Click()
    myaccountform.txt_oldusername.Text = lbl_username.Caption
    myaccountform.txt_username.Text = lbl_username.Caption
    Call load_form(myaccountform, True)
End Sub

Private Sub toolbar_advisories_Click()
    Call load_form(myadvisoriesform, True)
End Sub

Private Sub toolbar_help_Click()
     Call load_form(help2form, True)
End Sub

Private Sub toolbar_logout_Click()
    Unload Me
End Sub

Private Sub toolbar_section_Click()
     Call load_form(mysectionform, True)
End Sub
