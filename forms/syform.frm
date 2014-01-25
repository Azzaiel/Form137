VERSION 5.00
Begin VB.Form syform 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "School Year Settings"
   ClientHeight    =   2880
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4515
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "syform.frx":0000
   ScaleHeight     =   2880
   ScaleWidth      =   4515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Left            =   1800
      MaxLength       =   35
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1440
      Width           =   2655
   End
   Begin VB.CommandButton cmd_save 
      Height          =   615
      Left            =   960
      Picture         =   "syform.frx":977E
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton cmb_clear 
      Height          =   615
      Left            =   2280
      Picture         =   "syform.frx":A721
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2160
      Width           =   1095
   End
   Begin VB.ComboBox cmb_sy 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "syform.frx":B49C
      Left            =   1800
      List            =   "syform.frx":B4B5
      TabIndex        =   0
      Top             =   840
      Width           =   2655
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Note: Your password is needed to set for another school-year."
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
      Height          =   615
      Left            =   240
      TabIndex        =   6
      Top             =   120
      Width           =   3975
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
      Left            =   240
      TabIndex        =   5
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label Label1 
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
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   960
      Width           =   1575
   End
End
Attribute VB_Name = "syform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmb_clear_Click()
    cmb_sy.Text = ""
    txt_password.Text = ""
End Sub

Private Sub cmb_sy_KeyUp(KeyCode As Integer, Shift As Integer)
    MsgBox "Please select a school year from the list."
    cmb_sy.Text = ""
End Sub

Private Sub cmd_save_Click()
    Dim ans As String
     If cmb_sy.Text = "" Then
        MsgBox "Please select a school year."
    Else
        If txt_password.Text = "" Then
            MsgBox "Please input your own password."
        Else
            If txt_password.Text = user_password Then
                 ans = MsgBox("Are you sure you want to set the school year?", vbYesNo, "Set School Year")
                    If ans = vbNo Then
                        Exit Sub
                    Else
                MsgBox "You have selected the school-year " & cmb_sy.Text & "."
                school_year = cmb_sy.Text
                mainform.lbl_sy.Caption = school_year
                Unload Me
                End If
            Else
                MsgBox "Incorrect password."
            End If
        End If
    End If
End Sub

Private Sub Form_Load()
    Call mysql_select(public_rs, "SELECT * FROM tbl_sy ORDER BY SY DESC")
    cmb_sy.Clear
    While Not public_rs.EOF
        cmb_sy.AddItem (public_rs.Fields("sy").Value & "-" & Left(public_rs.Fields("sy").Value, 3) & Trim(Str(val(Right(public_rs.Fields("sy").Value, 1) + 1))))
        public_rs.MoveNext
    Wend
End Sub
