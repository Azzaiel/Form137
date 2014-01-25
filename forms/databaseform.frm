VERSION 5.00
Begin VB.Form databaseform 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Database"
   ClientHeight    =   4365
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5985
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "databaseform.frx":0000
   ScaleHeight     =   4365
   ScaleWidth      =   5985
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd_ok 
      Default         =   -1  'True
      Height          =   615
      Left            =   2400
      Picture         =   "databaseform.frx":12121
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3600
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
      Left            =   840
      TabIndex        =   5
      Top             =   3120
      Width           =   4095
   End
   Begin VB.DriveListBox drive_backup 
      Height          =   315
      Left            =   1440
      TabIndex        =   0
      Top             =   480
      Width           =   2895
   End
   Begin VB.DirListBox dir_backup 
      Height          =   2115
      Left            =   1440
      TabIndex        =   1
      Top             =   840
      Width           =   2895
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Choose Path to Backup Database File"
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
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   5295
   End
   Begin VB.Label lbl_restore 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Restore database"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   3960
      TabIndex        =   4
      Top             =   3960
      Width           =   1935
   End
End
Attribute VB_Name = "databaseform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_ok_Click()
    Dim ans As String
    If txt_path.Text = "" Then
        MsgBox "Please choose path."
    Else
         ans = MsgBox("Are you sure you want to back-up database file?", vbYesNo, "Back-up Database")
                    If ans = vbNo Then
                        Exit Sub
                    Else
                Dim identify As String
                identifier = "_" + InputBox("Enter date dd-mm-yy", "Date")
                
        backup_db (GetShortName(txt_path.Text) & "\db_form137" & identifier & ".sql")
        MsgBox "Database successfully copied."
        End If
    End If
End Sub

Private Sub dir_backup_Change()
    txt_path.Text = dir_backup.Path
End Sub

Private Sub drive_backup_Change()
    On Error GoTo message
    dir_backup.Path = drive_backup.Drive
    Exit Sub
message:
    MsgBox "Device is unavailable"
    Exit Sub
End Sub

Private Sub lbl_restore_Click()
      Call load_form(restoreform, True)
End Sub
