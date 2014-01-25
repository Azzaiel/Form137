VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form form137allform 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Search for Student Form 137"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8985
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "form137allform.frx":0000
   ScaleHeight     =   6000
   ScaleWidth      =   8985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd_search 
      Height          =   615
      Left            =   6960
      Picture         =   "form137allform.frx":1B3CE
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   240
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
      Left            =   1080
      TabIndex        =   0
      Top             =   240
      Width           =   5775
   End
   Begin MSDataGridLib.DataGrid dg_students 
      Height          =   4215
      Left            =   360
      TabIndex        =   2
      Top             =   1200
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   7435
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
   Begin VB.Label lbl_view_character 
      BackStyle       =   0  'Transparent
      Caption         =   "View student's character grades."
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
      Left            =   360
      TabIndex        =   5
      Top             =   5520
      Width           =   3495
   End
   Begin VB.Label lbl_view_grade 
      BackStyle       =   0  'Transparent
      Caption         =   "View student's grades."
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
      Left            =   6360
      TabIndex        =   4
      Top             =   5520
      Width           =   2415
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "List of Students."
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
      TabIndex        =   3
      Top             =   960
      Width           =   3735
   End
End
Attribute VB_Name = "form137allform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public rs_student As New ADODB.Recordset
Public rs_student2 As New ADODB.Recordset

Private Sub cmd_search_Click()
       Call set_datagrid(dg_students, rs_student, _
                                        "SELECT " _
                                            & "student_id as LRN, last_name as Last_Name, first_name as First_Name, middle_name as Middle_Name, Gender as Gender, Bday as Date_Of_Birth, contact_no as Contact_Number,address as Address, father_name as Father_Name, father_no as Father_Contact, mother_name as Mother_Name, mother_no as Mother_Contact FROM tbl_student WHERE student_id = '" & txt_search.Text & "' OR last_name = '" & txt_search.Text & "' OR first_name = '" & txt_search.Text & "'")
                                        
      If rs_student.RecordCount = 0 Then
        MsgBox "Record not found."
      End If
End Sub

Private Sub Form_Load()
    Call set_datagrid(dg_students, rs_student, _
                                        "SELECT " _
                                            & "student_id as LRN, last_name as Last_Name, first_name as First_Name, middle_name as Middle_Name FROM tbl_student")
                                        
                    
                               
End Sub

Private Sub lbl_view_character_Click()
    If rs_student.RecordCount = 0 Then
        MsgBox "No record."
        Exit Sub
    End If
     form137characterform.lbl_id.Caption = rs_student.Fields("LRN")
     form137characterform.lbl_name.Caption = rs_student.Fields("Last_Name") & ", " & rs_student.Fields("First_Name")
    Call mysql_select(rs_student2, "SELECT * FROM tbl_student_level WHERE ID='" & rs_student.Fields("LRN") & "'")
       If rs_student2.RecordCount = 0 Then
       form137characterform.lbl_level.Caption = "Not enrolled"
      form137characterform.lbl_section.Caption = "Not enrolled"
    Else
      form137characterform.lbl_level.Caption = rs_student2.Fields("lvl_name")
      form137characterform.lbl_section.Caption = rs_student2.Fields("section_name")
    End If
    Call mysql_select(public_rs, "SELECT * FROM tbl_student_level WHERE ID = '" & rs_student.Fields("LRN") & "'")
     form137characterform.cmb_sy.Clear
    While Not public_rs.EOF
        form137characterform.cmb_sy.AddItem (public_rs.Fields("SY"))
        public_rs.MoveNext
    Wend
     Call load_form(form137characterform, True)
    
End Sub

Private Sub lbl_view_grade_Click()
    If rs_student.RecordCount = 0 Then
        MsgBox "No record."
        Exit Sub
    End If
     form137gradeform.lbl_id.Caption = rs_student.Fields("LRN")
     form137gradeform.lbl_name.Caption = rs_student.Fields("Last_Name") & ", " & rs_student.Fields("First_Name")
    Call mysql_select(public_rs, "SELECT * FROM tbl_student_level WHERE ID='" & rs_student.Fields("LRN") & "'")
      If public_rs.RecordCount = 0 Then
        form137gradeform.lbl_level.Caption = "Not Enrolled"
      form137gradeform.lbl_section.Caption = "Not Enrolled"
    Else
      form137gradeform.lbl_level.Caption = public_rs.Fields("lvl_name")
      form137gradeform.lbl_section.Caption = public_rs.Fields("section_name")
    End If
    Call mysql_select(public_rs, "SELECT * FROM tbl_student_level WHERE ID = '" & rs_student.Fields("LRN") & "'")
    
     form137gradeform.cmb_sy.Clear
    While Not public_rs.EOF
        form137gradeform.cmb_sy.AddItem (public_rs.Fields("SY"))
        public_rs.MoveNext
    Wend
     Call load_form(form137gradeform, True)
    
End Sub
Private Sub txt_search_KeyUp(KeyCode As Integer, Shift As Integer)
      Call set_datagrid(dg_students, rs_student, _
                                        "SELECT " _
                                            & "student_id as LRN, last_name as Last_Name, first_name as First_Name, middle_name as Middle_Name, Gender as Gender, Bday as Date_Of_Birth, contact_no as Contact_Number,address as Address, father_name as Father_Name, father_no as Father_Contact, mother_name as Mother_Name, mother_no as Mother_Contact FROM tbl_student WHERE student_id LIKE '%" & txt_search.Text & "%' OR last_name LIKE '%" & txt_search.Text & "%' OR first_name LIKE '%" & txt_search.Text & "%'")
                                        
                    
End Sub
