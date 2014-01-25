VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form teacherform 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "List of Teachers"
   ClientHeight    =   5955
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "teacherform.frx":0000
   ScaleHeight     =   5955
   ScaleWidth      =   9000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd_new 
      Height          =   615
      Left            =   3360
      Picture         =   "teacherform.frx":1B3CE
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5280
      Width           =   1095
   End
   Begin VB.CommandButton cmd_edit 
      Height          =   615
      Left            =   4680
      Picture         =   "teacherform.frx":1C2A5
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5280
      Width           =   1095
   End
   Begin VB.CommandButton cmd_search 
      Height          =   615
      Left            =   6960
      Picture         =   "teacherform.frx":1D23E
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   360
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
      Top             =   360
      Width           =   5775
   End
   Begin MSDataGridLib.DataGrid dg_teachers 
      Height          =   3735
      Left            =   360
      TabIndex        =   4
      Top             =   1440
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
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "List of Teachers."
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
      TabIndex        =   5
      Top             =   1200
      Width           =   4815
   End
End
Attribute VB_Name = "teacherform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public rs_teacher As New ADODB.Recordset

Private Sub cmd_edit_Click()
If rs_teacher.RecordCount = 0 Then
    MsgBox "No selected teacher."
    Exit Sub
Else

    teacherinformationform.txt_id.Text = rs_teacher.Fields("Teacher_ID")
    teacherinformationform.txt_firstname.Text = rs_teacher.Fields("First_Name")
    teacherinformationform.txt_lastname.Text = rs_teacher.Fields("Last_Name")
    teacherinformationform.txt_middlename.Text = rs_teacher.Fields("Middle_Name")
    teacherinformationform.cmb_gender.Text = rs_teacher.Fields("Gender")
    teacherinformationform.dateBday.Value = rs_teacher.Fields("Date_Of_Birth")
    'teacherinformationform.txt_contact.Text = rs_teacher.Fields("Contact_Number")
    'teacherinformationform.txt_address.Text = rs_teacher.Fields("Address")
    'teacherinformationform.txt_course.Text = rs_teacher.Fields("Course")
    'teacherinformationform.txt_school.Text = rs_teacher.Fields("School_Attended")
    'teacherinformationform.txt_from.Text = rs_teacher.Fields("From_Year")
    'teacherinformationform.txt_to.Text = rs_teacher.Fields("To_Year")
    teacherinformationform.cmb_status.Text = rs_teacher.Fields("Status")
    teacherinformationform.txt_op.Text = "edit"
    teacherinformationform.txt_oldid.Text = rs_teacher.Fields("Teacher_ID")
    
    Call mysql_select(teacherinformationform.rs_teacher_sections, "SELECT lvl_name as Level, section_name as Section FROM tbl_teacher_sections a, tbl_section b where  a.section_id = b.section_id and a.teacher_id = '" & teacherinformationform.txt_id.Text & "'")
    Set teacherinformationform.dg_sections.DataSource = teacherinformationform.rs_teacher_sections
    Call teacherinformationform.reloadSectionDataGrid
    
    
    Call load_form(teacherinformationform, True)
    
   
End If
End Sub

Private Sub cmd_new_Click()
    Dim temp As String
    Dim no As Integer
    teacherinformationform.txt_id.Text = ""
    Call mysql_select(public_rs, "SELECT * FROM tbl_teacher  ORDER BY teacher_id DESC LIMIT 1")
    
    If public_rs.RecordCount <> 0 Then
        temp = public_rs.Fields("teacher_id").Value
        temp = Mid$(temp, 3, 4)
        no = val(temp)
        no = no + 1
        teacherinformationform.txt_id.Text = "M-" & Format(no, "0000")
        
    Else
        no = 1
         teacherinformationform.txt_id.Text = "M-" & Format(no, "0000")
    End If
    
    Call mysql_select(public_rs, "Delete FROM db_form137.tbl_teacher_sections where teacher_id = '" & teacherinformationform.txt_id & "'")
    
    teacherinformationform.txt_firstname.Text = ""
    teacherinformationform.txt_lastname.Text = ""
    teacherinformationform.txt_middlename.Text = ""
    teacherinformationform.cmb_gender.Text = ""
    teacherinformationform.dateBday.Value = Now
    'teacherinformationform.txt_contact.Text = ""
    'teacherinformationform.txt_address.Text = ""
    'teacherinformationform.txt_course.Text = ""
    'teacherinformationform.txt_school.Text = ""
    'teacherinformationform.txt_from.Text = ""
    'teacherinformationform.txt_to.Text = ""
    teacherinformationform.cmb_status.Text = "On-Duty"
    teacherinformationform.txt_op.Text = "add"
    Call load_form(teacherinformationform, True)
End Sub

Private Sub cmd_search_Click()
    Call set_datagrid(dg_teachers, rs_teacher, _
                                        "SELECT " _
                                            & "teacher_id as Teacher_ID, last_name as Last_Name, first_name as First_Name, middle_name as Middle_Name, Gender as Gender, Bday as Date_Of_Birth, contact_no as Contact_Number,address as Address, course as Course, school as School_Attended, a_from as From_Year, a_to as To_Year, status as Status FROM tbl_teacher WHERE teacher_id ='" & txt_search & "' OR last_name = '" & txt_search.Text & "' OR first_name = '" & txt_search.Text & "'")
    If rs_teacher.RecordCount = 0 Then
        MsgBox "No record found."
    End If
End Sub

Private Sub dg_teachers_DblClick()
  Call cmd_edit_Click
End Sub

Public Sub Form_Load()
      Call set_datagrid(dg_teachers, rs_teacher, _
                                        "SELECT " _
                                            & "teacher_id as Teacher_ID, last_name as Last_Name, first_name as First_Name, middle_name as Middle_Name, Gender as Gender, Bday as Date_Of_Birth, contact_no as Contact_Number,address as Address, course as Course, school as School_Attended, a_from as From_Year, a_to as To_Year, status as Status FROM tbl_teacher")
                                        
                    
                                       
End Sub

Private Sub txt_search_KeyUp(KeyCode As Integer, Shift As Integer)
     Call set_datagrid(dg_teachers, rs_teacher, _
                                        "SELECT " _
                                            & "teacher_id as Teacher_ID, last_name as Last_Name, first_name as First_Name, middle_name as Middle_Name, Gender as Gender, Bday as Date_Of_Birth, contact_no as Contact_Number,address as Address, course as Course, school as School_Attended, a_from as From_Year, a_to as To_Year, status as Status FROM tbl_teacher WHERE teacher_id LIKE'%" & txt_search & "%' OR last_name LIKE '%" & txt_search.Text & "%' OR first_name LIKE '%" & txt_search.Text & "%'")
                    
End Sub
