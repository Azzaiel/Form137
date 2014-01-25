VERSION 5.00
Begin VB.Form security2form 
   BorderStyle     =   0  'None
   Caption         =   "Security"
   ClientHeight    =   2415
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5670
   LinkTopic       =   "Form1"
   Picture         =   "security2form.frx":0000
   ScaleHeight     =   2415
   ScaleWidth      =   5670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Security Question No. 2"
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
      Begin VB.CommandButton cmd_ok 
         Height          =   615
         Left            =   2160
         Picture         =   "security2form.frx":1DD3
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1560
         Width           =   1095
      End
      Begin VB.TextBox txt_place 
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
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
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
         Height          =   255
         Left            =   480
         TabIndex        =   3
         Top             =   600
         Width           =   5415
      End
   End
End
Attribute VB_Name = "security2form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ctr As Integer
Private Sub cmd_ok_Click()
    If txt_place.Text = "" Then
        MsgBox "Please input your answer."
        Exit Sub
    End If
     Call mysql_select(public_rs, "SELECT * FROM tbl_security WHERE Username = '" & user & "' AND Place='" & txt_place.Text & "'")
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
            Unload Me
            Call load_form(security3form, True)
          
    
End If
End Sub

Private Sub Form_Load()
     ctr = 2
End Sub
