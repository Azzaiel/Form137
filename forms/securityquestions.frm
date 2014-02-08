VERSION 5.00
Begin VB.Form securityquestions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Security Questions"
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5940
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "securityquestions.frx":0000
   ScaleHeight     =   4455
   ScaleWidth      =   5940
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd_save 
      Height          =   615
      Left            =   1800
      Picture         =   "securityquestions.frx":12121
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton cmb_clear 
      Height          =   615
      Left            =   3000
      Picture         =   "securityquestions.frx":130C4
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3240
      Width           =   1095
   End
   Begin VB.TextBox txt_author 
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
      Left            =   240
      MaxLength       =   35
      TabIndex        =   2
      Top             =   2760
      Width           =   5415
   End
   Begin VB.TextBox txt_place 
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
      Left            =   240
      MaxLength       =   35
      TabIndex        =   1
      Top             =   1800
      Width           =   5415
   End
   Begin VB.TextBox txt_pet 
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
      Left            =   240
      MaxLength       =   35
      TabIndex        =   0
      Top             =   840
      Width           =   5415
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Note: Remember all your answers in security questions to access your account."
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
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   3960
      Width           =   5775
   End
   Begin VB.Label Label2 
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
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   2400
      Width           =   5415
   End
   Begin VB.Label Label1 
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
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1440
      Width           =   5415
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   480
      Width           =   5415
   End
End
Attribute VB_Name = "securityquestions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public sql_string As String
Public rs_security As New ADODB.Recordset
Private Sub cmb_clear_Click()

 Dim response As String
    response = MsgBox("Are you sure you want to clear the data?", vbYesNo, "Question")
    If (response = vbYes) Then
       txt_pet.Text = ""
       txt_place.Text = ""
       txt_author.Text = ""
    End If
End Sub

Private Sub cmd_save_Click()
    If txt_pet.Text = "" Or txt_place.Text = "" Or txt_author.Text = "" Then
        MsgBox "Please complete all fields."
        Exit Sub
    End If
     If is_duplicate("tbl_security", "Username", user) Then
            sql_string = "UPDATE tbl_security SET Pet='" & txt_pet.Text & "', Place='" & txt_place.Text & "', Author='" & txt_author.Text & "' WHERE Username='" & user & "'"
                    Call mysql_select(rs_security, sql_string)
                    MsgBox "Security answers successfully updated."
    Else
         sql_string = "INSERT INTO tbl_security (Username,Pet,Place,Author) VALUES ('" & user & "','" & txt_pet.Text & "','" & txt_place.Text & "', '" & txt_author.Text & "')"
                    Call mysql_select(rs_security, sql_string)
                    MsgBox "Security answers successfully added."
    End If
    
End Sub

