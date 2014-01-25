VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form levelform 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Level Settings"
   ClientHeight    =   4260
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5850
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "levelform.frx":0000
   ScaleHeight     =   4260
   ScaleWidth      =   5850
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txt_oldname 
      Height          =   375
      Left            =   4440
      TabIndex        =   8
      Top             =   1200
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txt_op 
      Height          =   375
      Left            =   4560
      TabIndex        =   7
      Top             =   720
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmb_clear 
      Height          =   615
      Left            =   3120
      Picture         =   "levelform.frx":12121
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton cmd_save 
      Height          =   615
      Left            =   1680
      Picture         =   "levelform.frx":12E9C
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   600
      Width           =   1095
   End
   Begin VB.TextBox txt_level 
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
      Left            =   2040
      MaxLength       =   20
      TabIndex        =   0
      Top             =   120
      Width           =   3375
   End
   Begin MSDataGridLib.DataGrid dg_level 
      Height          =   2175
      Left            =   240
      TabIndex        =   3
      Top             =   1800
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   3836
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   24
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
         Size            =   12
         Charset         =   0
         Weight          =   400
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
   Begin VB.CommandButton cmd_settings 
      Height          =   495
      Left            =   3600
      Picture         =   "levelform.frx":13E3F
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3000
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Double click to edit a level."
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
      Left            =   240
      TabIndex        =   6
      Top             =   1320
      Width           =   3735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Level Name:"
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
      Left            =   480
      TabIndex        =   5
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "levelform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public rs_level As New ADODB.Recordset
Dim sql_string As String

Private Sub cmb_clear_Click()
        txt_level.Text = ""
    txt_op.Text = "add"
End Sub

Private Sub cmd_save_Click()
Dim ans As String
If txt_op.Text = "add" Then
    If Not txt_level.Text = "" Then
        Call mysql_select(public_rs, "SELECT * FROM tbl_level WHERE lvl_name = '" & txt_level.Text & "'")
        If public_rs.RecordCount = 0 Then
             ans = MsgBox("Are you sure you want to save grade level?", vbYesNo, "Grade Level")
                    If ans = vbNo Then
                        Exit Sub
                    Else
            Call mysql_select(rs_level, "INSERT INTO tbl_level(lvl_name) VALUES ( '" & txt_level.Text & "')")
            MsgBox "Level successfully added!"
            txt_level.Text = ""
            Call Form_Load
            End If
        Else
            MsgBox "Level already exists."
        End If
    Else
        MsgBox "Please input a level name."
    End If
Else
    If Not txt_level.Text = "" Then
        If txt_level.Text = txt_oldname.Text Then
            MsgBox "Nothing to edit"
        Else
            Call mysql_select(public_rs, "SELECT * FROM tbl_level WHERE lvl_name = '" & txt_level.Text & "'")
            If public_rs.RecordCount = 0 Then
                 ans = MsgBox("Are you sure you want to update grade level?", vbYesNo, "Grade Level")
                    If ans = vbNo Then
                        Exit Sub
                    Else
                Call mysql_select(rs_level, "UPDATE tbl_level SET lvl_name = '" & txt_level.Text & "' WHERE lvl_name ='" & txt_oldname.Text & "'")
                MsgBox "Level successfully updated!"
                txt_level.Text = ""
                txt_op.Text = "add"
                Call Form_Load
                End If
            Else
                MsgBox "Level already exists."
            End If
        End If
    Else
        MsgBox "Please select a level name."
    End If
End If
End Sub

Private Sub cmd_settings_Click()
    level = rs_level.Fields("Name")
     Call load_form(subjectform, True)
     Unload Me
End Sub

Private Sub dg_level_DblClick()
    txt_level.Text = rs_level.Fields("Name")
    txt_op.Text = "edit"
    txt_oldname.Text = rs_level.Fields("Name")
      level = rs_level.Fields("Name")
     Call load_form(subjectform, True)
     Unload Me
End Sub

Private Sub Form_Load()
    Call set_datagrid(dg_level, rs_level, "SELECT lvl_name as Name FROM tbl_level ")
    txt_op.Text = "add"
End Sub
