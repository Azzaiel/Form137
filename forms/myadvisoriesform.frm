VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form myadvisoriesform 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "My Advisories"
   ClientHeight    =   5640
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8985
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "myadvisoriesform.frx":0000
   ScaleHeight     =   5640
   ScaleWidth      =   8985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
   Begin VB.CommandButton cmd_search 
      Height          =   615
      Left            =   6960
      Picture         =   "myadvisoriesform.frx":1B3CE
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin MSDataGridLib.DataGrid dg_sections 
      Height          =   3735
      Left            =   360
      TabIndex        =   3
      Top             =   1200
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   6588
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
   Begin VB.Label lbl_view_masterlist 
      BackStyle       =   0  'Transparent
      Caption         =   "View the masterlist."
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
      Left            =   6600
      TabIndex        =   2
      Top             =   5160
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "List of My Advisories"
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
Attribute VB_Name = "myadvisoriesform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public rs_advisories As New ADODB.Recordset
Public rs_masterlist As New ADODB.Recordset
Dim usertype, id As String

Private Sub cmb_category_Click()
    Dim col_order As String
    Select Case (cmb_category.Text)
        Case "Grade Level"
            col_order = "lvl_name"
        Case "Section"
            col_order = "section_name"
    End Select
    Call mysql_select(public_rs, "SELECT * FROM tbl_user WHERE Username = '" & mainform.lbl_username.Caption & "'")
         id = public_rs.Fields("ID")
        Call set_datagrid(dg_sections, rs_advisories, _
                                        "SELECT " _
                                            & "lvl_name as Level, section_name as Section FROM tbl_section WHERE teacher_id='" & id & "' AND SY = '" & mainteacherform.lbl_sy.Caption & "'ORDER BY " & col_order & " ASC")
End Sub

Private Sub cmb_category_KeyUp(KeyCode As Integer, Shift As Integer)
    MsgBox "Please select catgeory from the list."
    cmb_category.Text = ""
End Sub

Private Sub cmd_search_Click()
       Call mysql_select(public_rs, "SELECT * FROM tbl_user WHERE Username = '" & mainform.lbl_username.Caption & "'")
         id = public_rs.Fields("ID")
        
        
        Dim sqlQuery As String
        
        sqlQuery = "SELECT b.lvl_name as Level, b.section_name as Section  " & _
                   "FROM tbl_section b, tbl_teacher_sections a " & _
                   "WHERE a.teacher_id='" & id & "'" & _
                   "      and a.section_id = b.section_id " & _
                   "      AND (b.lvl_name = '" & txt_search.Text & "' OR b.section_name = '" & txt_search.Text & "') " & _
                   "ORDER BY b.lvl_name ASC "
         
        Call set_datagrid(dg_sections, rs_advisories, sqlQuery)
                                        
                    
            
        If rs_advisories.RecordCount = 0 Then
            MsgBox "Record not found."
        End If
End Sub

Private Sub Form_Load()
         Call mysql_select(public_rs, "SELECT * FROM tbl_user WHERE Username = '" & mainform.lbl_username.Caption & "'")
         id = public_rs.Fields("ID")
         
        Dim sqlQuery As String
        
        sqlQuery = "SELECT b.lvl_name as Level, b.section_name as Section  " & _
                   "FROM tbl_section b, tbl_teacher_sections a " & _
                   "WHERE a.teacher_id='" & id & "'" & _
                   "      and a.section_id = b.section_id " & _
                   "ORDER BY b.lvl_name ASC "
         
        Call set_datagrid(dg_sections, rs_advisories, sqlQuery)
                                        
                    
                  
End Sub

Private Sub lbl_view_masterlist_Click()
    If rs_advisories.RecordCount = 0 Then
        MsgBox "No record."
        Exit Sub
    Else
     masterlistadvisoriesform.lbl_level.Caption = rs_advisories.Fields("Level")
    masterlistadvisoriesform.lbl_section.Caption = rs_advisories.Fields("Section")
    section = rs_advisories.Fields("Section")
    If rs_advisories.Fields("Level").Value = "Kinder" Then
        masterlistadvisoriesform.lbl_set_grade.Enabled = False
    Else
        masterlistadvisoriesform.lbl_set_grade.Enabled = True
    End If
           
    Call load_form(masterlistadvisoriesform, True)
End If
End Sub

Private Sub txt_search_KeyUp(KeyCode As Integer, Shift As Integer)
       Call mysql_select(public_rs, "SELECT * FROM tbl_user WHERE Username = '" & mainform.lbl_username.Caption & "'")
         id = public_rs.Fields("ID")
         
         Dim sqlQuery As String
        
        sqlQuery = "SELECT b.lvl_name as Level, b.section_name as Section  " & _
                   "FROM tbl_section b, tbl_teacher_sections a " & _
                   "WHERE a.teacher_id='" & id & "'" & _
                   "      and a.section_id = b.section_id " & _
                   "      and (b.lvl_name LIKE '%" & txt_search.Text & "%' OR b.section_name LIKE '%" & txt_search.Text & "%') " & _
                   "ORDER BY b.lvl_name ASC "
        Call set_datagrid(dg_sections, rs_advisories, sqlQuery)
                                        
                    
               
End Sub
