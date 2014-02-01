VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form reportsform 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reports"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11040
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "reportform.frx":0000
   ScaleHeight     =   6000
   ScaleWidth      =   11040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd_print 
      Height          =   615
      Left            =   9240
      Picture         =   "reportform.frx":1DD3
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5280
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      Height          =   975
      Left            =   120
      TabIndex        =   2
      Top             =   -120
      Width           =   10935
      Begin VB.ComboBox cmb_sort 
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
         ItemData        =   "reportform.frx":2D29
         Left            =   7560
         List            =   "reportform.frx":2D2B
         TabIndex        =   5
         Top             =   360
         Width           =   2055
      End
      Begin VB.ComboBox cmb_category 
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
         ItemData        =   "reportform.frx":2D2D
         Left            =   2160
         List            =   "reportform.frx":2D3A
         TabIndex        =   0
         Top             =   360
         Width           =   3375
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Sort By:"
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
         Left            =   6480
         TabIndex        =   6
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Report For:"
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
         Left            =   840
         TabIndex        =   3
         Top             =   480
         Width           =   1815
      End
   End
   Begin MSDataGridLib.DataGrid dg_reports 
      Height          =   4215
      Left            =   240
      TabIndex        =   4
      Top             =   960
      Width           =   10575
      _ExtentX        =   18653
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
   Begin VB.Label lbl_no 
      BackStyle       =   0  'Transparent
      Caption         =   "Number of Record:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   495
      Left            =   240
      TabIndex        =   7
      Top             =   5400
      Width           =   8295
   End
End
Attribute VB_Name = "reportsform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs_report As New ADODB.Recordset
Private Sub cmb_category_Change()
    If cmb_category.Text = "Teacher" Then
           Call set_datagrid(dg_reports, rs_report, _
                                        "SELECT " _
                                            & "teacher_id as Teacher_ID, last_name as Last_Name, first_name as First_Name, middle_name as Middle_Name, status as Status FROM tbl_teacher ORDER BY last_name ASC")
                                        
          
    ElseIf cmb_category.Text = "Student" Then
        Call set_datagrid(dg_reports, rs_report, _
                                        "SELECT " _
                                            & "student_id as LRN, last_name as Last_Name, first_name as First_Name, middle_name as Middle_Name FROM tbl_student")
                                        
    ElseIf cmb_category.Text = "User" Then
        Call set_datagrid(dg_reports, rs_report, _
                                        "SELECT " _
                                            & "Usertype, Username FROM tbl_user")
                                        
                                        
    Else
        dg_reports.ClearFields
    End If
     
End Sub

Private Sub cmb_category_Click()
       If cmb_category.Text = "Teacher" Then
           Call set_datagrid(dg_reports, rs_report, _
                                        "SELECT " _
                                            & "teacher_id as Teacher_ID, last_name as Last_Name, first_name as First_Name, middle_name as Middle_Name, status as Status FROM tbl_teacher ORDER BY Last_Name ASC")
                                        
            cmb_sort.Clear
            cmb_sort.AddItem ("ID")
            cmb_sort.AddItem ("Last Name")
            cmb_sort.AddItem ("First Name")
            
    ElseIf cmb_category.Text = "Student" Then
        Call set_datagrid(dg_reports, rs_report, _
                                        "SELECT " _
                                            & "student_id as LRN, last_name as Last_Name, first_name as First_Name, middle_name as Middle_Name FROM tbl_student ORDER BY Last_Name ASC")
                                        
           cmb_sort.Clear
            cmb_sort.AddItem ("ID")
            cmb_sort.AddItem ("Last Name")
            cmb_sort.AddItem ("First Name")
    ElseIf cmb_category.Text = "User" Then
        Call set_datagrid(dg_reports, rs_report, _
                                        "SELECT " _
                                            & "Usertype, Username,ID FROM tbl_user ORDER BY Username ASC")
                                        
            cmb_sort.Clear
            cmb_sort.AddItem ("ID")
            cmb_sort.AddItem ("Username")
            cmb_sort.AddItem ("Usertype")
    ElseIf cmb_category.Text = "Masterlist" Then
         Call set_datagrid(dg_reports, rs_report, _
                                        "SELECT " _
                                            & "lvl_name as Level, section_name as Section FROM tbl_section ")
            cmb_sort.Clear
            cmb_sort.AddItem ("Level")
            cmb_sort.AddItem ("Section")
    Else
        dg_reports.ClearFields
    End If
    lbl_no.Caption = "Number of Record: " & rs_report.RecordCount
End Sub

Private Sub cmb_category_KeyUp(KeyCode As Integer, Shift As Integer)
    MsgBox "Please select category from the list."
    cmb_category.Text = ""
End Sub

Private Sub cmb_sort_Click()
    Dim choice As String
    If cmb_sort.Text = "" Then
        MsgBox "Please select an item from the list."
        Exit Sub
    End If
   
    If cmb_category.Text = "Teacher" Then
             If cmb_sort.Text = "ID" Then
                choice = "teacher_id"
            ElseIf cmb_sort.Text = "Last Name" Then
                choice = "last_name"
            Else
                choice = "first_name"
            End If
           Call set_datagrid(dg_reports, rs_report, _
                                        "SELECT " _
                                            & "teacher_id as Teacher_ID, last_name as Last_Name, first_name as First_Name, middle_name as Middle_Name, status as Status FROM tbl_teacher ORDER BY " & choice & " ASC")
                                        

    ElseIf cmb_category.Text = "Student" Then
         If cmb_sort.Text = "ID" Then
                choice = "student_id"
            ElseIf cmb_sort.Text = "Last Name" Then
                choice = "last_name"
            Else
                choice = "first_name"
            End If
        Call set_datagrid(dg_reports, rs_report, _
                                        "SELECT " _
                                            & "student_id as LRN, last_name as Last_Name, first_name as First_Name, middle_name as Middle_Name FROM tbl_student ORDER BY " & choice & " ASC")
                                        
    ElseIf cmb_category.Text = "User" Then
         If cmb_sort.Text = "ID" Then
                choice = "ID"
            ElseIf cmb_sort.Text = "Username" Then
                choice = "Username"
            Else
                choice = "Usertype"
            End If
        Call set_datagrid(dg_reports, rs_report, _
                                        "SELECT " _
                                            & "Usertype, Username,ID FROM tbl_user ORDER BY " & choice & " ASC")
                                        
                    
    ElseIf cmb_category.Text = "Masterlist" Then
     
        If txt_search.Text = "" Then
           If cmb_sort.Text = "Level" Then
                choice = "lvl_name"
            ElseIf cmb_sort.Text = "Section" Then
                choice = "section_name"
            End If
         Call set_datagrid(dg_reports, rs_report, _
                                        "SELECT " _
                                            & "lvl_name as Level, section_name as Section FROM tbl_section ORDER BY " & choice & " ASC ")
          Else
           Dim col_order As String
    Select Case (cmb_sort.Text)
        Case "ID"
            col_order = "student_id"
        Case "Last Name"
            col_order = "last_name"
        Case "First Name"
            col_order = "first_name"
    End Select
     
    Call set_datagrid(dg_reports, rs_report, _
                                            "SELECT @index := @index + 1 as No," _
                                                & "masterlist.* " _
                                            & "FROM " _
                                                & "(SELECT " _
                                                    & "a.student_id as LRN, " _
                                                    & "a.last_name as Last_Name, a.First_Name,a.Middle_Name " _
                                                & "FROM " _
                                                    & "tbl_student a " _
                                                & "LEFT JOIN " _
                                                    & "tbl_student_level b " _
                                                & "ON " _
                                                    & "a.student_id = b.ID " _
                                                & "WHERE " _
                                                    & "  b.section_name = '" & txt_search.Text & "' ORDER BY " & col_order & " ASC) masterlist" _
                                            & " JOIN " _
                                                & "(SELECT @index :=0) c ")
    dg_reports.Columns(0).Width = 400
        End If
       
         
    Else
        dg_reports.ClearFields
    End If
    lbl_no.Caption = "Number of Record: " & rs_report.RecordCount
End Sub

Private Sub cmb_sort_KeyUp(KeyCode As Integer, Shift As Integer)
    MsgBox "Please select an item from the list."
    cmb_sort.Text = ""
End Sub

Private Sub cmd_print_Click()
     If dg_reports.DataSource Is Nothing Then
        MsgBox "No record."
        Exit Sub
    End If
    If rs_report.RecordCount = 0 Then
        MsgBox "No record."
        Exit Sub
    End If
    If cmb_category.Text = "Teacher" Then
         dr_teacher.Sections(2).Controls("lbl_date").Caption = Format(Now, "mmmm dd, yyyy") & "/ " & Time
         Set dr_teacher.DataSource = rs_report
        dr_teacher.Show vbModal, Me
    ElseIf cmb_category.Text = "Student" Then
         dr_student.Sections(2).Controls("lbl_date").Caption = Format(Now, "mmmm dd, yyyy") & "/ " & Time
         Set dr_student.DataSource = rs_report
        dr_student.Show vbModal, Me
    ElseIf cmb_category.Text = "User" Then
         dr_user.Sections(2).Controls("lbl_date").Caption = Format(Now, "mmmm dd, yyyy") & "/ " & Time
         Set dr_user.DataSource = rs_report
        dr_user.Show vbModal, Me
    ElseIf cmb_category.Text = "Masterlist" Then
        If txt_search.Text = "" Then
            MsgBox "Please input the name of section you want to search."
            Exit Sub
        End If
        If txt_search.Text <> "" And dg_reports.DataSource Is Nothing Then
             MsgBox "Please input section name first."
            Exit Sub
        End If
        
         'dr_masterlist2.Sections(2).Controls("lbl_date").Caption = Format(Now, "mmmm dd, yyyy") & "/ " & Time
         'dr_masterlist2.Sections(2).Controls("lbl_sy").Caption = mainteacherform.lbl_sy.Caption
        dr_masterlist2.Sections(2).Controls("lbl_section").Caption = txt_search.Text
        dr_masterlist2.Sections(2).Controls("lbl_no").Caption = rs_report.RecordCount
         Set dr_masterlist2.DataSource = rs_report
        dr_masterlist2.Show vbModal, Me
    End If
End Sub

Private Sub dg_reports_DblClick()
    If cmb_category.Text = "Masterlist" Then
        txt_search.Text = rs_report.Fields("Section").value
        Call set_datagrid(dg_reports, rs_report, _
                                        "SELECT  @index := @index + 1 as No," _
                                            & "a.student_id as LRN, a.last_name as Last_Name, a.First_Name,a.Middle_Name FROM tbl_student a LEFT JOIN tbl_student_level b ON a.student_id = b.ID JOIN(SELECT @index :=0) c WHERE  b.section_name = '" & rs_report.Fields("Section").value & "'")
          cmb_sort.Clear
          cmb_sort.AddItem ("ID")
          cmb_sort.AddItem ("Last Name")
          cmb_sort.AddItem ("First Name")
            
    

     dg_reports.Columns(0).Width = 400
    End If
End Sub

Private Sub txt_search_KeyUp(KeyCode As Integer, Shift As Integer)
    If cmb_category.Text = "Teacher" Then
           Call set_datagrid(dg_reports, rs_report, _
                                        "SELECT " _
                                            & "teacher_id as Teacher_ID, last_name as Last_Name, first_name as First_Name, middle_name as Middle_Name, status as Status FROM tbl_teacher WHERE teacher_id LIKE '%" & txt_search.Text & "%' OR last_name LIKE '%" & txt_search.Text & "%' OR first_name LIKE '%" & txt_search.Text & "%' ORDER BY Last_Name ASC")
                                        
                    
    ElseIf cmb_category.Text = "Student" Then
        Call set_datagrid(dg_reports, rs_report, _
                                        "SELECT " _
                                            & "student_ids as LRN, last_name as Last_Name, first_name as First_Name, middle_name as Middle_Name FROM tbl_student WHERE student_id LIKE '%" & txt_search.Text & "%' OR last_name LIKE '%" & txt_search.Text & "%' OR first_name LIKE '%" & txt_search.Text & "%' ORDER BY Last_Name ASC")
                                        
    ElseIf cmb_category.Text = "User" Then
        Call set_datagrid(dg_reports, rs_report, _
                                        "SELECT " _
                                            & "Usertype, Username,ID FROM tbl_user WHERE Username LIKE '%" & txt_search.Text & "%' OR Usertype LIKE '%" & txt_search.Text & "%' OR ID LIKE '%" & txt_search.Text & "%' ORDER BY Username ASC")
                                        
                    
    ElseIf cmb_category.Text = "Masterlist" Then
         Call set_datagrid(dg_reports, rs_report, _
                                        "SELECT " _
                                            & "a.student_id as LRN, a.last_name as Last_Name, a.First_Name,a.Middle_Name FROM tbl_student a LEFT JOIN tbl_student_level b ON a.student_id = b.ID WHERE b.SY='" & mainform.lbl_sy.Caption & "' AND b.section_name = '" & txt_search.Text & "' ORDER BY a.Last_Name ASC")
          cmb_sort.Clear
          cmb_sort.AddItem ("ID")
          cmb_sort.AddItem ("Last Name")
          cmb_sort.AddItem ("First Name")
    
    Else
        dg_reports.ClearFields
    End If
End Sub
