VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form studentform 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "List of Students"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8865
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "studentform.frx":0000
   ScaleHeight     =   6000
   ScaleWidth      =   8865
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
      Top             =   360
      Width           =   5775
   End
   Begin VB.CommandButton cmd_search 
      Height          =   615
      Left            =   6840
      Picture         =   "studentform.frx":1B3CE
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   360
      Width           =   1095
   End
   Begin VB.CommandButton cmd_edit 
      Height          =   615
      Left            =   4560
      Picture         =   "studentform.frx":1C21D
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5280
      Width           =   1095
   End
   Begin VB.CommandButton cmd_new 
      Height          =   615
      Left            =   3240
      Picture         =   "studentform.frx":1D1B6
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5280
      Width           =   1095
   End
   Begin MSDataGridLib.DataGrid dg_students 
      Height          =   3855
      Left            =   240
      TabIndex        =   4
      Top             =   1320
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   6800
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
      Left            =   360
      TabIndex        =   5
      Top             =   1080
      Width           =   3735
   End
End
Attribute VB_Name = "studentform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs_student As New ADODB.Recordset

Private Sub cmd_edit_Click()
    studentinformationform.transferee.Visible = False
       studentinformationform.txt_id.Enabled = False
       studentinformationform.txt_id2.Visible = False
  If rs_student.RecordCount = 0 Then
    MsgBox "No selected record."
    Exit Sub
Else
    studentinformationform.txt_id.Text = rs_student.Fields("LRN")
    studentinformationform.txt_firstname.Text = rs_student.Fields("First_Name")
    studentinformationform.txt_lastname.Text = rs_student.Fields("Last_Name")
    studentinformationform.txt_middlename.Text = rs_student.Fields("Middle_Name")
    studentinformationform.cmb_gender.Text = rs_student.Fields("Gender")
    studentinformationform.dateBday.Value = rs_student.Fields("Date_Of_Birth")
    studentinformationform.txt_no.Text = rs_student.Fields("Contact_Number")
    studentinformationform.txt_address.Text = rs_student.Fields("Address")
    studentinformationform.txt_father.Text = rs_student.Fields("Guardian_Name")
    studentinformationform.txt_father_no.Text = rs_student.Fields("Guardian_Contact")
    studentinformationform.txt_place.Text = rs_student.Fields("Birth_Place")
    studentinformationform.txt_occupation.Text = rs_student.Fields("Occupation")
    
    studentinformationform.txt_op.Text = "edit"
    studentinformationform.txt_oldid.Text = rs_student.Fields("LRN")
    Call mysql_select(public_rs, "SELECT TRUNCATE(FLOOR(((12 * (YEAR(NOW())- YEAR(bday))+ (MONTH(NOW())- MONTH( bday))) / 12) * 4) / 4 , 2) AS Age From tbl_student WHERE student_id ='" & rs_student.Fields("LRN") & "'")
    studentinformationform.txt_age.Text = public_rs.Fields("Age")
     Call mysql_select(public_rs, "SELECT * FROM tbl_student_level WHERE SY = '" & mainform.lbl_sy.Caption & "' AND ID='" & rs_student.Fields("LRN") & "'")
    If public_rs.RecordCount = 0 Then
       level = ""
     section = ""
      status = ""
     studentinformationform.cmb_level.Text = level
     studentinformationform.cmb_section.Text = section
     studentinformationform.cmb_status.Text = status
     Call set_level2
    Else
     level = public_rs.Fields("lvl_name")
     section = public_rs.Fields("section_name")
      status = public_rs.Fields("Status")
     studentinformationform.cmb_level.Text = level
     studentinformationform.cmb_section.Text = section
    
     studentinformationform.cmb_status.Text = status
      Call set_level
   End If
    
     
   
    Call load_form(studentinformationform, True)
    
End If
End Sub

Private Sub cmd_new_Click()
    studentinformationform.transferee.Visible = True
    studentinformationform.txt_id.Text = "109637"
    studentinformationform.txt_id.Enabled = False
    studentinformationform.txt_id2.Visible = True
    studentinformationform.txt_firstname.Text = ""
    studentinformationform.txt_lastname.Text = ""
    studentinformationform.txt_middlename.Text = ""
    studentinformationform.cmb_gender.Text = ""
    studentinformationform.dateBday.Value = Now
    
    studentinformationform.txt_no.Text = ""
    studentinformationform.txt_address.Text = ""
    studentinformationform.txt_father.Text = ""
    studentinformationform.txt_father_no.Text = ""
    studentinformationform.txt_place.Text = ""
    studentinformationform.txt_occupation.Text = ""
    studentinformationform.cmb_status.Text = "Enrolled"
    studentinformationform.txt_op.Text = "add"
     Call set_level2
    Call load_form(studentinformationform, True)
    
End Sub

Private Sub cmd_search_Click()
     Call set_datagrid(dg_students, rs_student, _
                                        "SELECT " _
                                            & "student_id as LRN, last_name as Last_Name, first_name as First_Name, middle_name as Middle_Name, Gender as Gender, Bday as Date_Of_Birth, birthplace as Birth_Place, contact_no as Contact_Number,address as Address, guardian as Guardian_Name, guardian_no as Guardian_Contact, occupation as Occupation FROM tbl_student WHERE student_id = '" & txt_search.Text & "' OR last_name = '" & txt_search.Text & "' OR first_name = '" & txt_search.Text & "'")
                                        
      If rs_student.RecordCount = 0 Then
        MsgBox "Record not found."
      End If
      Call formatDataGrid
End Sub

Public Sub Form_Load()
     Call set_datagrid(dg_students, rs_student, _
                                        "SELECT " _
                                            & "student_id as LRN, last_name as Last_Name, first_name as First_Name, middle_name as Middle_Name, Gender as Gender, Bday as Date_Of_Birth, birthplace as Birth_Place, contact_no as Contact_Number,address as Address, guardian as Guardian_Name, guardian_no as Guardian_Contact, occupation as Occupation FROM tbl_student")
                                        
                    
    Call formatDataGrid
End Sub
Public Sub txt_search_KeyUp(KeyCode As Integer, Shift As Integer)
    Call set_datagrid(dg_students, rs_student, _
                                        "SELECT " _
                                            & "student_id as LRN, last_name as Last_Name, first_name as First_Name, middle_name as Middle_Name, Gender as Gender, Bday as Date_Of_Birth, birthplace as Birth_Place, contact_no as Contact_Number,address as Address, guardian as Guardian_Name, guardian_no as Guardian_Contact, occupation as Occupation FROM tbl_student WHERE student_id LIKE '%" & txt_search.Text & "%' OR last_name LIKE '%" & txt_search.Text & "%' OR first_name LIKE '%" & txt_search.Text & "%'")
                                        
   Call formatDataGrid
End Sub
Private Sub formatDataGrid()
  With dg_students
   
  End With
End Sub
Private Sub set_level()
    
    Call mysql_select(public_rs, "SELECT * FROM tbl_student_level WHERE SY = '" & mainform.lbl_sy.Caption & "' AND ID='" & rs_student.Fields("LRN") & "'")
        
        If public_rs.RecordCount <> 0 Then
            section = public_rs.Fields("section_name")
           
            Call mysql_select(public_rs, "SELECT * FROM tbl_section WHERE SY = '" & mainform.lbl_sy.Caption & "' AND section_name='" & section & "'")
                level = public_rs.Fields("lvl_name")
            Call mysql_select(public_rs, "SELECT * FROM tbl_level WHERE SY = '" & mainform.lbl_sy.Caption & "'")
            studentinformationform.cmb_level.Clear
            While Not public_rs.EOF
                studentinformationform.cmb_level.AddItem (public_rs.Fields("lvl_name"))
                public_rs.MoveNext
            Wend
            If Not level = "" Then
                studentinformationform.cmb_level.Text = level
                
                
            End If
            Call mysql_select(public_rs, "SELECT * FROM tbl_section WHERE SY='" & mainform.lbl_sy.Caption & "' AND lvl_name = '" & level & "'")
            studentinformationform.cmb_section.Clear
            While Not public_rs.EOF
                studentinformationform.cmb_section.AddItem (public_rs.Fields("section_name"))
                public_rs.MoveNext
            Wend
            If Not section = "" Then
                studentinformationform.cmb_section.Text = section
                section = ""
            End If
        Else
            
             Call mysql_select(public_rs, "SELECT * FROM tbl_level WHERE SY = '" & mainform.lbl_sy.Caption & "'")
            studentinformationform.cmb_level.Clear
            While Not public_rs.EOF
                studentinformationform.cmb_level.AddItem (public_rs.Fields("lvl_name"))
                public_rs.MoveNext
            Wend
            If Not level = "" Then
                studentinformationform.cmb_level.Text = level
                level = ""
                
            End If
            Call mysql_select(public_rs, "SELECT * FROM tbl_section WHERE SY='" & mainform.lbl_sy.Caption & "' AND lvl_name = '" & cmb_level.Text & "'")
            studentinformationform.cmb_section.Clear
            While Not public_rs.EOF
                studentinformationform.cmb_section.AddItem (public_rs.Fields("section_name"))
                public_rs.MoveNext
            Wend
            If Not section = "" Then
                studentinformationform.cmb_section.Text = section
                section = ""
            End If
        End If
  
    
End Sub
Private Sub set_level2()
    
    
            
             Call mysql_select(public_rs, "SELECT * FROM tbl_level WHERE SY = '" & mainform.lbl_sy.Caption & "'")
            studentinformationform.cmb_level.Clear
            While Not public_rs.EOF
                studentinformationform.cmb_level.AddItem (public_rs.Fields("lvl_name"))
                public_rs.MoveNext
            Wend
            If Not level = "" Then
                
                level = ""
                
            End If
            
       
  
    
End Sub
