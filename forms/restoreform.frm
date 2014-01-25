VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form restoreform 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Restore Database"
   ClientHeight    =   2955
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4530
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "restoreform.frx":0000
   ScaleHeight     =   2955
   ScaleWidth      =   4530
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cdDB 
      Left            =   2880
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton btn_browse 
      Default         =   -1  'True
      Height          =   615
      Left            =   1560
      Picture         =   "restoreform.frx":977E
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   600
      Width           =   1095
   End
   Begin VB.TextBox txt_path 
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
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   4335
   End
   Begin VB.CommandButton cmd_ok 
      Height          =   615
      Left            =   1560
      Picture         =   "restoreform.frx":A53E
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Browse Database File"
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
      Left            =   960
      TabIndex        =   3
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "restoreform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btn_browse_Click()
    cdDB.Filter = ("Sql(*.sql)|*.sql")
    cdDB.ShowOpen
    If Not cdDB.FileName = "" Then
        txt_path.Text = cdDB.FileName
        
    Else
        txt_path.Text = ""
    End If
End Sub

Private Sub cmd_ok_Click()
    Dim ans As String
     If txt_path.Text = "" Then
        MsgBox "Please browse for an SQL file."
    Else
         ans = MsgBox("Are you sure you want to restore the selected database file?", vbYesNo, "Restore Database")
                    If ans = vbNo Then
                        Exit Sub
                    Else
        restore_db (GetShortName(txt_path.Text))
        MsgBox "Database successfully restored."
        End If
    End If
End Sub
