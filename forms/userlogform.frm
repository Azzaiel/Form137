VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form userlogform 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "User Log History"
   ClientHeight    =   5880
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8985
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "userlogform.frx":0000
   ScaleHeight     =   5880
   ScaleWidth      =   8985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnclear 
      Height          =   615
      Left            =   7080
      Picture         =   "userlogform.frx":1B3CE
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5160
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmd_edit 
      Height          =   615
      Left            =   3600
      Picture         =   "userlogform.frx":1C16C
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5160
      Width           =   1575
   End
   Begin VB.CommandButton cmd_search 
      Height          =   615
      Left            =   6960
      Picture         =   "userlogform.frx":1D0C2
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox txt_search 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   5775
   End
   Begin MSDataGridLib.DataGrid dg_logs 
      Height          =   3975
      Left            =   360
      TabIndex        =   2
      Top             =   1080
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   7011
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   19
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "List of User Logs."
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
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   840
      Width           =   3735
   End
End
Attribute VB_Name = "userlogform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs_log As New ADODB.Recordset
Public sql_string As String

Private Sub btnclear_Click()
    sql_string = "DELETE FROM tbl_logs"
        Call mysql_select(rs_log, sql_string)
        Call Form_Load
End Sub

Private Sub cmd_edit_Click()
    If rs_log.RecordCount = 0 Then
        MsgBox "No record."
        Exit Sub
    End If
     dr_logs.Sections(2).Controls("lbl_date").Caption = Format(Now, "mmmm dd, yyyy") & "/ " & Time
    Set dr_logs.DataSource = rs_log
    dr_logs.Show vbModal, Me
End Sub

Private Sub cmd_search_Click()
      Call set_datagrid(dg_logs, rs_log, _
                                        "SELECT " _
                                            & " * FROM tbl_logs WHERE Username = '" & txt_search.Text & "' OR Login = '" & txt_search.Text & "' OR Logout = '" & txt_search.Text & "'")
                                            
    If rs_log.RecordCount = 0 Then
        MsgBox "Record not found."
    End If
                                        
                    
           
End Sub

Private Sub Command1_Click()

End Sub

Public Sub Form_Load()
      Call set_datagrid(dg_logs, rs_log, _
                                        "SELECT " _
                                            & " * FROM tbl_logs")
                                        
                    
                                           
End Sub

Private Sub txt_search_KeyUp(KeyCode As Integer, Shift As Integer)
     Call set_datagrid(dg_logs, rs_log, _
                                        "SELECT " _
                                            & " * FROM tbl_logs WHERE Username LIKE '%" & txt_search.Text & "%' OR Login LIKE '%" & txt_search.Text & "%' OR Logout LIKE '%" & txt_search.Text & "%'")
                                        
                    
                                           
End Sub
