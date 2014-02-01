VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form sectionform 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Section Settings"
   ClientHeight    =   5970
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7440
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "sectionform.frx":0000
   ScaleHeight     =   5970
   ScaleWidth      =   7440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txt_oldsection 
      Height          =   375
      Left            =   7440
      TabIndex        =   11
      Text            =   "Text2"
      Top             =   1560
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmd_settings 
      Height          =   495
      Left            =   5160
      Picture         =   "sectionform.frx":1B3CE
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5400
      Width           =   1815
   End
   Begin VB.TextBox txt_op 
      Height          =   375
      Left            =   7440
      TabIndex        =   9
      Text            =   "Text2"
      Top             =   960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txt_section 
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
      Left            =   2520
      MaxLength       =   30
      TabIndex        =   1
      Top             =   720
      Width           =   3375
   End
   Begin VB.CommandButton cmd_save 
      Height          =   615
      Left            =   2280
      Picture         =   "sectionform.frx":1C591
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1680
      Width           =   1095
   End
   Begin VB.CommandButton cmb_clear 
      Height          =   615
      Left            =   3600
      Picture         =   "sectionform.frx":1D534
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   8535
      Begin VB.ComboBox cmb_level 
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
         Left            =   2520
         TabIndex        =   0
         Top             =   0
         Width           =   3375
      End
      Begin VB.Label Label3 
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
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   720
         TabIndex        =   8
         Top             =   120
         Width           =   1455
      End
   End
   Begin MSDataGridLib.DataGrid dg_sections 
      Height          =   2655
      Left            =   720
      TabIndex        =   4
      Top             =   2640
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   4683
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
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Double click to edit a section."
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
      TabIndex        =   10
      Top             =   2400
      Width           =   3735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Section Name:"
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
      Left            =   600
      TabIndex        =   7
      Top             =   840
      Width           =   1695
   End
End
Attribute VB_Name = "sectionform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs_section As New ADODB.Recordset

Private Sub cmb_clear_Click()
    txt_op.Text = "add"
    txt_section.Text = ""
     Call mysql_select(public_rs, "SELECT CONCAT(CONCAT(first_name,' '),last_name) as Name FROM tbl_teacher WHERE status = 'Active'")
End Sub
Private Sub cmb_level_Click()
     txt_section.Text = ""
    'cmb_adviser.Clear
    Call set_datagrid(dg_sections, rs_section, _
                                        "SELECT " _
                                            & " a.section_name As Section_Name, " _
                                            & "a.teacher_id As Teacher_ID, CONCAT(b.first_name,' ', b.last_name) as Teacher_Name " _
                                        & "FROM " _
                                            & "tbl_section a " _
                                        & "LEFT JOIN " _
                                            & "tbl_teacher b " _
                                        & "ON " _
                                            & "a.teacher_id = b.teacher_id " _
                                        & "WHERE " _
                                            & "a.lvl_name = '" & cmb_level.Text & "' ")
     Call mysql_select(public_rs, "SELECT CONCAT(CONCAT(first_name,' '),last_name) as Name FROM tbl_teacher WHERE status = 'On-Duty'")
     Call formatDataGrid
    
    End Sub

Private Sub cmd_save_Click()
    Dim ans As String
    Dim sql_string As String
     If txt_op.Text = "add" Then
        If Not txt_section.Text = "" Then
            Call mysql_select(public_rs, "SELECT * FROM tbl_section WHERE lvl_name = '" & cmb_level.Text & "'AND section_name = '" & txt_section.Text & "'")
            If public_rs.RecordCount = 0 Then
               ' Call mysql_select(public_rs, "SELECT teacher_id FROM tbl_teacher WHERE CONCAT(CONCAT(first_name,' '),last_name) = '" & cmb_adviser.Text & "'")
               ' Dim tch_id As String
               ' If public_rs.RecordCount = 0 Then
               '     tch_id = "None"
               ' Else
               '     tch_id = public_rs.Fields("teacher_id").value
               ' End If
                 ans = MsgBox("Are you sure you want to add the section?", vbYesNo, "Add Section")
                    If ans = vbNo Then
                        Exit Sub
                    Else
                Call mysql_select(rs_section, "INSERT INTO tbl_section(lvl_name,section_name) VALUES ('" & cmb_level.Text & "','" & txt_section.Text & "')")
                MsgBox "Section successfully added!"
                txt_section.Text = ""
'                cmb_adviser.Clear
                level = cmb_level.Text
                Call Form_Load
                End If
            Else
                MsgBox "Section already exists."
            End If
        Else
            MsgBox "Please input a section name."
        End If
    Else
       If Not txt_section.Text = "" Then
            If txt_section.Text = txt_oldsection.Text Then
                Call mysql_select(public_rs, "SELECT teacher_id FROM tbl_teacher WHERE CONCAT(CONCAT(first_name,' '),last_name) = '" & cmb_adviser.Text & "'")
                Dim tch_id2 As String
                If public_rs.RecordCount <> 0 Then
                     tch_id2 = public_rs.Fields("teacher_id").value
                Else
                    tch_id2 = "None"
                End If
                     ans = MsgBox("Are you sure you want to update the section?", vbYesNo, "Update Section")
                    If ans = vbNo Then
                        Exit Sub
                    Else
                     sql_string = "UPDATE " _
                            & "tbl_section " _
                        & "SET " _
                            & "teacher_id = '" & tch_id2 & "' " _
                        & "WHERE " _
                            & " lvl_name = '" & cmb_level.Text & "'AND section_name = '" & txt_section.Text & "'"
                Call mysql_select(public_rs, sql_string)
                MsgBox "Section successfully updated!"
                txt_section.Text = ""
                cmb_adviser.Clear
                txt_op.Text = "add"
                level = cmb_level.Text
                Call Form_Load
                End If
            Else
                Call mysql_select(public_rs, "SELECT * FROM tbl_section WHERE lvl_name = '" & cmb_level.Text & "'AND section_name = '" & txt_section.Text & "'")
                If public_rs.RecordCount = 0 Then
                    Call mysql_select(public_rs, "SELECT teacher_id FROM tbl_teacher WHERE CONCAT(CONCAT(first_name,' '),last_name) = '" & cmb_adviser.Text & "'")
                Dim tch_id3 As String
                If public_rs.RecordCount <> 0 Then
                     tch_id3 = public_rs.Fields("teacher_id").value
                Else
                    tch_id3 = "None"
                End If
                     ans = MsgBox("Are you sure you want to update the section?", vbYesNo, "Update Section")
                    If ans = vbNo Then
                        Exit Sub
                    Else
                     sql_string = "UPDATE " _
                            & "tbl_section " _
                        & "SET " _
                            & " section_name='" & txt_section.Text & "', teacher_id = '" & tch_id3 & "' " _
                        & "WHERE " _
                            & " lvl_name = '" & cmb_level.Text & "'AND section_name = '" & txt_oldsection.Text & "'"
                Call mysql_select(public_rs, sql_string)
                MsgBox "Section successfully updated!"
                txt_section.Text = ""
                cmb_adviser.Clear
                txt_op.Text = "add"
                level = cmb_level.Text
                Call Form_Load
                End If
                Else
                    MsgBox "Section already exists."
                End If
            End If
        Else
              MsgBox "Please select a section."
        End If
    End If
End Sub

Private Sub cmd_settings_Click()
     level = cmb_level.Text
     section = rs_section.Fields("Section_Name")
    Call load_form(subjectsettingsform, True)
End Sub

Private Sub dg_sections_DblClick()
    txt_op.Text = "edit"
    txt_section.Text = rs_section.Fields("Section_Name")
    txt_oldsection.Text = rs_section.Fields("Section_Name")
     Call mysql_select(public_rs, "SELECT CONCAT(first_name,' ', last_name) as Name FROM tbl_teacher WHERE status = 'On-Duty'")
    cmb_adviser.Clear
    While Not public_rs.EOF
        cmb_adviser.AddItem (public_rs.Fields("Name"))
        public_rs.MoveNext
    Wend
End Sub

Private Sub Form_Load()
     Call mysql_select(public_rs, "SELECT * FROM tbl_level ")
    cmb_level.Clear
    While Not public_rs.EOF
        cmb_level.AddItem (public_rs.Fields("lvl_name"))
        public_rs.MoveNext
    Wend
    If Not level = "" Then
        cmb_level.Text = level
        level = ""
        'Call cmb_level_Change
    End If
    txt_op.Text = "add"
    Call set_datagrid(dg_sections, rs_section, _
                                        "SELECT " _
                                            & "a.section_name As Section_Name, " _
                                            & "a.teacher_id As Teacher_ID, CONCAT(b.first_name,' ', b.last_name) as Teacher_Name " _
                                        & "FROM " _
                                            & "tbl_section a " _
                                        & "LEFT JOIN " _
                                            & "tbl_teacher b " _
                                        & "ON " _
                                            & "a.teacher_id = b.teacher_id " _
                                        & "WHERE " _
                                            & "a.lvl_name = '" & cmb_level.Text & "' ")
    Call formatDataGrid
    Call mysql_select(public_rs, "SELECT CONCAT(CONCAT(first_name,' '),last_name) as Name FROM tbl_teacher WHERE status = 'On-Duty'")


End Sub
Private Function formatDataGrid()
  With dg_sections
    .Columns(1).Visible = False
    .Columns(2).Visible = False
  End With
End Function
