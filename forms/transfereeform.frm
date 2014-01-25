VERSION 5.00
Begin VB.Form transfereeform 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transferee's LRN"
   ClientHeight    =   1095
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7470
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "transfereeform.frx":0000
   ScaleHeight     =   1095
   ScaleWidth      =   7470
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd_ok 
      Height          =   615
      Left            =   5880
      Picture         =   "transfereeform.frx":1DD3
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   240
      Width           =   1095
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
      Left            =   1920
      MaxLength       =   6
      TabIndex        =   0
      Top             =   360
      Width           =   1695
   End
   Begin VB.TextBox txt_id2 
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
      Left            =   3600
      MaxLength       =   6
      TabIndex        =   1
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "LRN:"
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
      Left            =   720
      TabIndex        =   2
      Top             =   480
      Width           =   1095
   End
End
Attribute VB_Name = "transfereeform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_ok_Click()
    Dim no As Integer
    no = Len(txt_id2.Text)
    If no < 6 Then
        MsgBox "Please input only 6 added characters for student's LRN."
        txt_id2.SetFocus
        Exit Sub
    End If
    studentinformationform.txt_id.Text = txt_id.Text
    studentinformationform.txt_id2.Text = txt_id2.Text
    Unload Me
End Sub

Private Sub txt_id_KeyUp(KeyCode As Integer, Shift As Integer)
    If Not IsNumeric(txt_id.Text) Then
    
     MsgBox "Please enter numbers only."
     txt_id.Text = ""
     Exit Sub
     End If
     
End Sub

Private Sub txt_id2_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim no As Integer
    no = Len(txt_id.Text)
    If no < 6 Then
        MsgBox "Please input only 6 characters for School's Code in student's LRN."
        txt_id.SetFocus
        Exit Sub
    End If
    
    If Not IsNumeric(txt_id2.Text) Then
    
     MsgBox "Please enter numbers only."
     txt_id2.Text = ""
     Exit Sub
     End If
   
End Sub
