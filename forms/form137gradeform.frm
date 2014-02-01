VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form form137gradeform 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Student Form 137 Grades"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8985
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "form137gradeform.frx":0000
   ScaleHeight     =   6000
   ScaleWidth      =   8985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd_print 
      Height          =   615
      Left            =   5400
      Picture         =   "form137gradeform.frx":1B3CE
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5280
      Width           =   1575
   End
   Begin MSDataGridLib.DataGrid dg_grades 
      Height          =   3375
      Left            =   120
      TabIndex        =   9
      Top             =   1800
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   5953
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      Height          =   1695
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8775
      Begin VB.ComboBox cmb_period 
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
         ItemData        =   "form137gradeform.frx":1C324
         Left            =   5880
         List            =   "form137gradeform.frx":1C33A
         TabIndex        =   2
         Top             =   1200
         Width           =   2535
      End
      Begin VB.ComboBox cmb_sy 
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
         Left            =   1680
         TabIndex        =   1
         Top             =   1200
         Width           =   2535
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Period:"
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
         Left            =   4440
         TabIndex        =   16
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "School Year:"
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
         Left            =   240
         TabIndex        =   15
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label lbl_section 
         BackStyle       =   0  'Transparent
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
         Left            =   5880
         TabIndex        =   13
         Top             =   960
         Width           =   2775
      End
      Begin VB.Label lbl_level 
         BackStyle       =   0  'Transparent
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
         Left            =   1680
         TabIndex        =   12
         Top             =   960
         Width           =   2415
      End
      Begin VB.Label lbl_name 
         BackStyle       =   0  'Transparent
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
         Left            =   1680
         TabIndex        =   11
         Top             =   600
         Width           =   4815
      End
      Begin VB.Label lbl_id 
         BackStyle       =   0  'Transparent
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
         Left            =   1800
         TabIndex        =   10
         Top             =   240
         Width           =   4935
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "School Year:"
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
         Left            =   600
         TabIndex        =   8
         Top             =   1680
         Width           =   2535
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Section:"
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
         Left            =   4440
         TabIndex        =   7
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Level:"
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
         Left            =   240
         TabIndex        =   6
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
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
         Left            =   240
         TabIndex        =   5
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "LRN:"
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
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Label lbl_export 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Export Form-137"
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
      Left            =   7080
      TabIndex        =   17
      Top             =   5520
      Width           =   1815
   End
   Begin VB.Label lbl_average 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   240
      TabIndex        =   14
      Top             =   5400
      Width           =   4815
   End
End
Attribute VB_Name = "form137gradeform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public rs_grade As New ADODB.Recordset
Public public_rs2 As New ADODB.Recordset
Public public_all As New ADODB.Recordset
Public public_all2 As New ADODB.Recordset
Public average As Double
Public remark As String
Dim excelApp As Excel.Application
Dim ExcelWorkbook As Excel.Workbook
Dim ExcelSheet As Excel.Worksheet
Dim MyMonth As String
Dim MyYear As String
Dim Mydirectory As String
Dim MyFileName As String
Public sql_string As String
 

Private Sub cmb_grade_Change()

End Sub

Private Sub cmb_grade_Click()
    
End Sub

Private Sub cmb_grade_DblClick()

End Sub

Private Sub cmb_period_Click()
    If cmb_sy.Text = "" Then
        MsgBox "Please select a school year first."
        Exit Sub
    End If
    If cmb_period.Text = "All" Then
        Call set_datagrid(dg_grades, rs_grade, _
                                        "SELECT " _
                                            & "a.subject_code as Code, b.subject_name as Subject, a.grade as Grade, a.Remark FROM tbl_grade a LEFT JOIN tbl_subject b ON a.subject_code = b.subject_code WHERE a.SY='" & cmb_sy.Text & "'AND a.ID='" & lbl_id.Caption & "'")
                                        
                    
       Call mysql_select(public_rs, "SELECT subject_code as Subject, period, grade, remark FROM tbl_grade WHERE ID = '" & lbl_id.Caption & "' AND SY = '" & cmb_sy.Text & "' AND Period='Final'")
        no_subj = public_rs.RecordCount
        If no_subj = 0 Then
            lbl_average.Caption = "0 - No grades"
            average = 0
            remark = "No grades"
        Else
            average = 0
            While Not public_rs.EOF
                average = val(public_rs.Fields("grade")) + average
                public_rs.MoveNext
            Wend
            average = average / no_subj
            average = Round(average, 2)
            If average >= 90 Then
              remark = "A"
              ElseIf average >= 85 Then
              remark = "P"
              ElseIf average >= 80 Then
                  remark = "AP"
              ElseIf average >= 74 Then
                  remark = "D"
              Else
                  remark = "B"
            End If
            lbl_average.Caption = Str(average) & " - " & remark
        End If
    
    Else
      Call set_datagrid(dg_grades, rs_grade, _
                                        "SELECT " _
                                            & "a.subject_code as Code, b.subject_name as Subject, a.grade as Grade, a.Remark FROM tbl_grade a LEFT JOIN tbl_subject b ON a.subject_code = b.subject_code WHERE a.SY='" & cmb_sy.Text & "'AND a.ID='" & lbl_id.Caption & "' AND a.Period='" & cmb_period.Text & "'")
                                        
                    
       Call mysql_select(public_rs, "SELECT subject_code as Subject, period, grade, remark FROM tbl_grade WHERE ID = '" & lbl_id.Caption & "' AND SY = '" & cmb_sy.Text & "' AND Period='" & cmb_period.Text & "'")
        no_subj = public_rs.RecordCount
        If no_subj = 0 Then
            lbl_average.Caption = "0 - No grades"
            average = 0
            remark = "No grades"
        Else
            average = 0
            While Not public_rs.EOF
                average = val(public_rs.Fields("grade")) + average
                public_rs.MoveNext
            Wend
            average = average / no_subj
            average = Round(average, 2)
            If average >= 90 Then
              remark = "A"
              ElseIf average >= 85 Then
              remark = "P"
              ElseIf average >= 80 Then
                  remark = "AP"
              ElseIf average >= 74 Then
                  remark = "D"
              Else
                  remark = "B"
            End If
            lbl_average.Caption = Str(average) & " - " & remark
        End If
    End If
End Sub

Private Sub lbl_view_average_Click()
    gradeaverageform.lbl_id.Caption = lbl_id.Caption
    gradeaverageform.lbl_name.Caption = lbl_name.Caption
    gradeaverageform.lbl_level.Caption = lbl_level.Caption
    gradeaverageform.lbl_section.Caption = lbl_section.Caption
    gradeaverageform.cmb_sy.Text = cmb_sy.Text
     Call load_form(gradeaverageform, True)
     
End Sub

Private Sub cmb_period_KeyUp(KeyCode As Integer, Shift As Integer)
    MsgBox "Please select a period from the list."
    cmb_period.Text = ""
End Sub

Private Sub cmb_sy_KeyUp(KeyCode As Integer, Shift As Integer)
     MsgBox "Please select a school year from the list."
    cmb_sy.Text = ""
End Sub

Private Sub cmd_print_Click()
     Dim subject(10) As String
    Dim first(10) As String
    Dim sec(10) As String
    Dim third(10) As String
    Dim fourth(10) As String
    Dim f(10) As String
    If dg_grades.DataSource Is Nothing Then
        MsgBox "No record."
        Exit Sub
    End If
     If rs_grade.RecordCount = 0 Then
        MsgBox "No record."
        Exit Sub
    End If
    If cmb_period.Text = "" Or cmb_sy.Text = "" Then
        MsgBox "No record."
        Exit Sub
    End If
    Dim ctr As Integer
        ctr = 1
        If cmb_period.Text = "All" Then
            dr_gradeall.Sections(2).Controls("lbl_date").Caption = Now
            dr_gradeall.Sections(2).Controls("lbl_id").Caption = lbl_id.Caption
            dr_gradeall.Sections(2).Controls("lbl_name").Caption = lbl_name.Caption
            dr_gradeall.Sections(2).Controls("lbl_sy").Caption = cmb_sy.Text
            dr_gradeall.Sections(2).Controls("lbl_level").Caption = lbl_level.Caption
            dr_gradeall.Sections(2).Controls("lbl_section").Caption = lbl_section.Caption
            dr_gradeall.Sections(2).Controls("lbl_period").Caption = cmb_period.Text
            dr_gradeall.Sections(5).Controls("lbl_average").Caption = average
            dr_gradeall.Sections(5).Controls("lbl_remark").Caption = remark
             
            If remark = "" Or remark = "B" Then
                dr_gradeall.Sections(2).Controls("lbl_promote").Caption = "Unable to promote student to next grade level."
            Else
                Call mysql_select(public_rs, "SELECT * FROM tbl_student_level WHERE ID='" & lbl_id.Caption & "'AND SY='" & cmb_sy.Text & "'")
                Dim temp_level As String
                Dim id As Integer
                If public_rs.RecordCount = 0 Then
                    temp_level = ""
                Else
                    temp_level = public_rs.Fields("lvl_name").value
                    Call mysql_select(public_rs, "SELECT * FROM tbl_level WHERE lvl_name='" & temp_level & "'AND SY='" & cmb_sy.Text & "'")
                    id = public_rs.Fields("lvl_id").value
                     id = id + 1
                     Call mysql_select(public_rs2, "SELECT * FROM tbl_level WHERE lvl_id=" & id & "")
                     temp_level = public_rs2.Fields("lvl_name").value
                End If
                
                
                 dr_gradeall.Sections(2).Controls("lbl_promote").Caption = "Promote to " & temp_level
            End If
           
            
            
            Call mysql_select(public_all, "SELECT DISTINCT a.subject_code,  b.subject_name as Subject, a.Grade, a.Remark FROM tbl_grade a LEFT JOIN tbl_subject b ON a.subject_code=b.subject_code WHERE a.ID = '" & lbl_id.Caption & "' AND a.SY = '" & cmb_sy.Text & "' AND a.Period='1st Grading'  ")
                 While Not public_all.EOF
                    subject(ctr) = public_all.Fields("Subject").value
                    first(ctr) = public_all.Fields("Remark").value
                    ctr = ctr + 1
                    public_all.MoveNext
                Wend
                
                    dr_gradeall.Sections(2).Controls("lbl_subject1").Caption = subject(1)
                    dr_gradeall.Sections(2).Controls("lbl_subject2").Caption = subject(2)
                    dr_gradeall.Sections(2).Controls("lbl_subject3").Caption = subject(3)
                    dr_gradeall.Sections(2).Controls("lbl_subject4").Caption = subject(4)
                    dr_gradeall.Sections(2).Controls("lbl_subject5").Caption = subject(5)
                    dr_gradeall.Sections(2).Controls("lbl_subject6").Caption = subject(6)
                    dr_gradeall.Sections(2).Controls("lbl_subject7").Caption = subject(7)
                    dr_gradeall.Sections(2).Controls("lbl_subject8").Caption = subject(8)
                    dr_gradeall.Sections(2).Controls("lbl_subject9").Caption = subject(9)
                    dr_gradeall.Sections(2).Controls("lbl_subject10").Caption = subject(10)
                    
                    dr_gradeall.Sections(2).Controls("lbl_first1").Caption = first(1)
                    dr_gradeall.Sections(2).Controls("lbl_first2").Caption = first(2)
                    dr_gradeall.Sections(2).Controls("lbl_first3").Caption = first(3)
                    dr_gradeall.Sections(2).Controls("lbl_first4").Caption = first(4)
                    dr_gradeall.Sections(2).Controls("lbl_first5").Caption = first(5)
                    dr_gradeall.Sections(2).Controls("lbl_first6").Caption = first(6)
                    dr_gradeall.Sections(2).Controls("lbl_first7").Caption = first(7)
                    dr_gradeall.Sections(2).Controls("lbl_first8").Caption = first(8)
                    dr_gradeall.Sections(2).Controls("lbl_first9").Caption = first(9)
                    dr_gradeall.Sections(2).Controls("lbl_first10").Caption = first(10)
                   
               Call mysql_select(public_all2, "SELECT DISTINCT subject_code, Grade, Remark FROM tbl_grade WHERE ID = '" & lbl_id.Caption & "' AND SY = '" & cmb_sy.Text & "' AND Period='2nd Grading'  ")
                ctr = 1
                 While Not public_all2.EOF
                    If public_all2.RecordCount = 0 Then
                         sec(ctr) = "0"
                    Else
                         sec(ctr) = public_all2.Fields("Remark").value
                    End If
                   
                    ctr = ctr + 1
                    public_all2.MoveNext
                Wend
               
                
                    dr_gradeall.Sections(2).Controls("lbl_sec1").Caption = sec(1)
                    dr_gradeall.Sections(2).Controls("lbl_sec2").Caption = sec(2)
                    dr_gradeall.Sections(2).Controls("lbl_sec3").Caption = sec(3)
                    dr_gradeall.Sections(2).Controls("lbl_sec4").Caption = sec(4)
                    dr_gradeall.Sections(2).Controls("lbl_sec5").Caption = sec(5)
                    dr_gradeall.Sections(2).Controls("lbl_sec6").Caption = sec(6)
                    dr_gradeall.Sections(2).Controls("lbl_sec7").Caption = sec(7)
                    dr_gradeall.Sections(2).Controls("lbl_sec8").Caption = sec(8)
                    dr_gradeall.Sections(2).Controls("lbl_sec9").Caption = sec(9)
                    dr_gradeall.Sections(2).Controls("lbl_sec10").Caption = sec(10)
                    
                    Call mysql_select(public_all2, "SELECT DISTINCT subject_code, Grade, Remark FROM tbl_grade WHERE ID = '" & lbl_id.Caption & "' AND SY = '" & cmb_sy.Text & "' AND Period='3rd Grading'  ")
                ctr = 1
                 While Not public_all2.EOF
                    If public_all2.RecordCount = 0 Then
                         third(ctr) = "0"
                    Else
                         third(ctr) = public_all2.Fields("Remark").value
                    End If
                   
                    ctr = ctr + 1
                    public_all2.MoveNext
                Wend
                    dr_gradeall.Sections(2).Controls("lbl_third1").Caption = third(1)
                    dr_gradeall.Sections(2).Controls("lbl_third2").Caption = third(2)
                    dr_gradeall.Sections(2).Controls("lbl_third3").Caption = third(3)
                    dr_gradeall.Sections(2).Controls("lbl_third4").Caption = third(4)
                    dr_gradeall.Sections(2).Controls("lbl_third5").Caption = third(5)
                    dr_gradeall.Sections(2).Controls("lbl_third6").Caption = third(6)
                    dr_gradeall.Sections(2).Controls("lbl_third7").Caption = third(7)
                    dr_gradeall.Sections(2).Controls("lbl_third8").Caption = third(8)
                    dr_gradeall.Sections(2).Controls("lbl_third9").Caption = third(9)
                    dr_gradeall.Sections(2).Controls("lbl_third10").Caption = third(10)
                    
                    Call mysql_select(public_all2, "SELECT DISTINCT subject_code, Grade, Remark FROM tbl_grade WHERE ID = '" & lbl_id.Caption & "' AND SY = '" & cmb_sy.Text & "' AND Period='4th Grading'  ")
                ctr = 1
                While Not public_all2.EOF
                    If public_all2.RecordCount = 0 Then
                         fourth(ctr) = "0"
                    Else
                         fourth(ctr) = public_all2.Fields("Remark").value
                    End If
                   
                    ctr = ctr + 1
                    public_all2.MoveNext
                Wend
                    dr_gradeall.Sections(2).Controls("lbl_for1").Caption = fourth(1)
                    dr_gradeall.Sections(2).Controls("lbl_for2").Caption = fourth(2)
                    dr_gradeall.Sections(2).Controls("lbl_for3").Caption = fourth(3)
                    dr_gradeall.Sections(2).Controls("lbl_for4").Caption = fourth(4)
                    dr_gradeall.Sections(2).Controls("lbl_for5").Caption = fourth(5)
                    dr_gradeall.Sections(2).Controls("lbl_for6").Caption = fourth(6)
                    dr_gradeall.Sections(2).Controls("lbl_for7").Caption = fourth(7)
                    dr_gradeall.Sections(2).Controls("lbl_for8").Caption = fourth(8)
                    dr_gradeall.Sections(2).Controls("lbl_for9").Caption = fourth(9)
                    dr_gradeall.Sections(2).Controls("lbl_for10").Caption = fourth(10)
                    
                      Call mysql_select(public_all2, "SELECT DISTINCT subject_code, Grade, Remark FROM tbl_grade WHERE ID = '" & lbl_id.Caption & "' AND SY = '" & cmb_sy.Text & "' AND Period='Final'  ")
                ctr = 1
                While Not public_all2.EOF
                    If public_all2.RecordCount = 0 Then
                         f(ctr) = "0"
                    Else
                         f(ctr) = public_all2.Fields("Remark").value
                    End If
                   
                    ctr = ctr + 1
                    public_all2.MoveNext
                Wend
                    dr_gradeall.Sections(2).Controls("lbl_final1").Caption = f(1)
                    dr_gradeall.Sections(2).Controls("lbl_final2").Caption = f(2)
                    dr_gradeall.Sections(2).Controls("lbl_final3").Caption = f(3)
                    dr_gradeall.Sections(2).Controls("lbl_final4").Caption = f(4)
                    dr_gradeall.Sections(2).Controls("lbl_final5").Caption = f(5)
                    dr_gradeall.Sections(2).Controls("lbl_final6").Caption = f(6)
                    dr_gradeall.Sections(2).Controls("lbl_final7").Caption = f(7)
                    dr_gradeall.Sections(2).Controls("lbl_final8").Caption = f(8)
                    dr_gradeall.Sections(2).Controls("lbl_final9").Caption = f(9)
                    dr_gradeall.Sections(2).Controls("lbl_final10").Caption = f(10)
            
             Set dr_gradeall.DataSource = public_all
            dr_gradeall.Show vbModal, Me
        Else
      'dr_grade.Sections(2).Controls("lbl_sy").Caption = mainform.lbl_sy.Caption
      dr_grade.Sections(2).Controls("lbl_date").Caption = Format(Now, "mmmm dd, yyyy") & "/ " & Time
        dr_grade.Sections(2).Controls("lbl_level").Caption = lbl_level.Caption
        dr_grade.Sections(2).Controls("lbl_section").Caption = lbl_section.Caption
        dr_grade.Sections(2).Controls("lbl_id").Caption = lbl_id.Caption
        dr_grade.Sections(2).Controls("lbl_name").Caption = lbl_name.Caption
        dr_grade.Sections(2).Controls("lbl_period").Caption = cmb_period.Text
        dr_grade.Sections(5).Controls("lbl_average").Caption = average
        dr_grade.Sections(5).Controls("lbl_remark").Caption = remark
         Set dr_grade.DataSource = rs_grade
        dr_grade.Show vbModal, Me
    End If
End Sub

Private Sub lbl_export_Click()
Dim subject(10) As String
    Dim first(10) As String
    Dim sec(10) As String
    Dim third(10) As String
    Dim fourth(10) As String
    Dim f(10) As String
    Dim f2(10) As String
     If lbl_id.Caption = "" Then
        MsgBox "No record selected."
        Exit Sub
    End If
    MyFileName = App.Path & "\Form-137\" & lbl_id.Caption & "-" & lbl_name.Caption & "-Grade.xls"
    On Error Resume Next
    Set excelApp = CreateObject("Excel.Application")
'if file exists, place file name in FileCheck
FileCheck = Dir$(MyFileName)
  If FileCheck = MyMonth + "_" + MyYear + MyExtension Then
    'Workbook exists, open it
    Set ExcelWorkbook = excelApp.Workbooks.Open(MyFileName)
    Set ExcelSheet = ExcelWorkbook.Worksheets(1)
  Else
'create Excel object
Set excelApp = CreateObject("Excel.Application")
    Set ExcelWorkbook = excelApp.Workbooks.Add
    Set ExcelSheet = ExcelWorkbook.Worksheets(1)
    
    With ExcelSheet.PageSetup
     .PaperSize = xlPaperLegal
     .Zoom = False
     .FitToPagesWide = 1
   End With
   
    ExcelSheet.Name = "Grades"
    excelApp.Range("A1:O1").Merge
    excelApp.Range("A2:O2").Merge
    excelApp.Range("A3:O3").Merge
    excelApp.Range("A4:O4").Merge
    excelApp.Range("A5:O5").Merge
    excelApp.Range("A6:O6").Merge
    excelApp.Range("A7:O7").Merge
    excelApp.Range("A8:O8").Merge
    excelApp.Range("A10:O10").Merge
    excelApp.Range("A11:I11").Merge
    excelApp.Range("A22:O22").Merge
    excelApp.Range("A23:O23").Merge
    excelApp.Range("A24:O24").Merge
    excelApp.Range("A1:O1").HorizontalAlignment = xlCenter
    excelApp.Range("A1:O1").Font.Bold = True
    excelApp.Range("A2:O2").HorizontalAlignment = xlCenter
    
    excelApp.Range("A3:O3").HorizontalAlignment = xlCenter
    excelApp.Range("A3:O3").Font.Bold = True
     excelApp.Range("A4:O4").HorizontalAlignment = xlCenter
    
     excelApp.Range("A5:O5").HorizontalAlignment = xlCenter
    excelApp.Range("A5:O5").Font.Bold = True
     excelApp.Range("A6:O6").HorizontalAlignment = xlCenter
    excelApp.Range("A6:O6").Font.Bold = True
     excelApp.Range("A7:O7").HorizontalAlignment = xlCenter
    excelApp.Range("A7:O7").Font.Bold = True
     excelApp.Range("A8:O8").HorizontalAlignment = xlCenter
    excelApp.Range("A8:O8").Font.Bold = True
    excelApp.Range("A10:O10").HorizontalAlignment = xlCenter
    excelApp.Range("A10:O10").Font.Bold = True
    excelApp.Range("A10:O10").Font.Size = 16
    excelApp.Range("A11:O11").HorizontalAlignment = xlCenter
    excelApp.Range("A12").Font.Bold = True
    excelApp.Range("E12").Font.Bold = True
     excelApp.Range("G12").Font.Bold = True
      excelApp.Range("A16").Font.Bold = True
       excelApp.Range("C16").Font.Bold = True
        excelApp.Range("E16").Font.Bold = True
         excelApp.Range("G16").Font.Bold = True
          excelApp.Range("A19").Font.Bold = True
           excelApp.Range("A12").Font.Bold = True
     excelApp.Range("A12:H12").ColumnWidth = 15
      excelApp.Range("A13:H13").ColumnWidth = 15
       excelApp.Range("A14:H14").ColumnWidth = 15
        excelApp.Range("A16:H16").ColumnWidth = 15
         excelApp.Range("A17:H17").ColumnWidth = 15
          excelApp.Range("A19:H19").ColumnWidth = 15
        excelApp.Range("A20:H20").ColumnWidth = 15
         excelApp.Range("A21:H21").ColumnWidth = 15
          excelApp.Range("A13:O13").Font.Size = 9
          excelApp.Range("A17:O17").Font.Size = 9
          excelApp.Range("A20:O20").Font.Size = 9
          excelApp.Range("A21:O21").Font.Size = 9
           excelApp.Range("A22:O22").Font.Bold = True
        excelApp.Range("A22:O22").HorizontalAlignment = xlCenter
        excelApp.Range("A22:O22").Font.Size = 14
          excelApp.Range("A23:O23").Font.Bold = True
        excelApp.Range("A23:O23").HorizontalAlignment = xlCenter
        excelApp.Range("A23:O23").Font.Size = 14
        excelApp.Range("A24:O24").Font.Size = 14
        excelApp.Range("A25:A40").ColumnWidth = 30
        excelApp.Range("A25:O25").Font.Bold = True
        excelApp.Range("A25").HorizontalAlignment = xlCenter
       
        excelApp.Range("B26:O26").ColumnWidth = 5
        excelApp.Range("A26:O26").Font.Bold = True
        excelApp.Range("A26:O26").HorizontalAlignment = xlCenter
        excelApp.Range("B27:O27").HorizontalAlignment = xlCenter
        excelApp.Range("B28:O28").HorizontalAlignment = xlCenter
        excelApp.Range("B29:O29").HorizontalAlignment = xlCenter
        excelApp.Range("B30:O30").HorizontalAlignment = xlCenter
        excelApp.Range("B31:O31").HorizontalAlignment = xlCenter
        excelApp.Range("B32:O32").HorizontalAlignment = xlCenter
        excelApp.Range("B33:O33").HorizontalAlignment = xlCenter
        excelApp.Range("B34:O34").HorizontalAlignment = xlCenter
        excelApp.Range("B35:O35").HorizontalAlignment = xlCenter
        excelApp.Range("B36:O36").HorizontalAlignment = xlCenter
        excelApp.Range("B37:O37").HorizontalAlignment = xlCenter
        excelApp.Range("B38:O38").HorizontalAlignment = xlCenter
        excelApp.Range("B39:O39").HorizontalAlignment = xlCenter
        excelApp.Range("B40:O40").HorizontalAlignment = xlCenter
        excelApp.Range("B34:O34").HorizontalAlignment = xlCenter
        excelApp.Range("A45:O45").Font.Bold = True
        excelApp.Range("A45").HorizontalAlignment = xlCenter
        excelApp.Range("A46:O46").Font.Bold = True
        excelApp.Range("A46:O46").HorizontalAlignment = xlCenter
        excelApp.Range("A66:O66").Font.Bold = True
        excelApp.Range("A66:O66").HorizontalAlignment = xlCenter
        excelApp.Range("A83:B83").Font.Bold = True
        excelApp.Range("A84:B84").Font.Bold = True
        excelApp.Range("B45:C45").Merge
        excelApp.Range("G45:H45").Merge
        excelApp.Range("J45:K45").Merge
        excelApp.Range("D45:F45").Merge
        excelApp.Range("I45:K45").Merge
        excelApp.Range("N45:O45").Merge
        excelApp.Range("A65:O65").Font.Bold = True
        excelApp.Range("A65").HorizontalAlignment = xlCenter
         excelApp.Range("B65:C65").Merge
          excelApp.Range("D65:F65").Merge
         excelApp.Range("E1:P65").Font.Bold = True
    ExcelSheet.Cells(1, 1).value = "Republika ng Pilipinas"
    ExcelSheet.Cells(2, 1).value = "(Republic of the Philippines)"
    ExcelSheet.Cells(3, 1).value = "Kagawaran ng Edukasyon"
    ExcelSheet.Cells(4, 1).value = "(Department of Education)"
    ExcelSheet.Cells(5, 1).value = "KAWANIHAN NG EDUKASYONG ELEMENTARYA"
    ExcelSheet.Cells(6, 1).value = "(BUREAU OF ELEMENTARY EDUCATION)"
    ExcelSheet.Cells(7, 1).value = "Region IV-A CALABARZON"
    ExcelSheet.Cells(8, 1).value = "City of Cavite"
    ExcelSheet.Cells(10, 1).value = "PALAGIANG TALAAN SA PAARALANG ELEMENTARYA"
    ExcelSheet.Cells(11, 1).value = "(ELEMENTARY SCHOOL PERMANENT RECORD)"
    ExcelSheet.Cells(11, 10).value = "LRN"
    excelApp.Range("K11:O11").Merge
     excelApp.Range("A11:A11").HorizontalAlignment = xlRight
     excelApp.Range("I12:O11").Font.Underline = True
     excelApp.Range("J11:J11").Font.Bold = False
     excelApp.Range("J11:J11").Font.Underline = False
     
  'ExcelApp.Range("K11:P11").Borders(xlInsideVertical).LineStyle = xlContinuous
       ' ExcelApp.Range("K11:P11").LineStyle = xlContinuous
       ' ExcelApp.Range("K11:P11").Weight = xlMedium
    excelApp.Range("J12:K12").Merge
     excelApp.Range("L12:O12").Merge
     
    ExcelSheet.Cells(11, 11).value = " - " & lbl_id.Caption & " - "
    ExcelSheet.Cells(12, 1).value = "Pangalan"
     Call mysql_select(public_rs, "SELECT * FROM  tbl_student WHERE student_id = '" & lbl_id.Caption & "'")
    ExcelSheet.Cells(12, 2).value = public_rs.Fields("last_name").value
    Dim middle As String
    middle = public_rs.Fields("middle_name").value
    ExcelSheet.Cells(12, 7).value = Mid(middle, 1, 1)
    ExcelSheet.Cells(12, 5).value = public_rs.Fields("first_name").value
    ExcelSheet.Cells(12, 8).value = "Sangay"
    ExcelSheet.Cells(12, 9).value = "III"
    ExcelSheet.Cells(12, 10).value = "Paaralan"
    ExcelSheet.Cells(12, 12).value = "Manuel S. Rojas Elementary School"
    ExcelSheet.Cells(13, 1).value = "(Name)"
    ExcelSheet.Cells(13, 2).value = "Apelyido"
    ExcelSheet.Cells(13, 5).value = "Unang ngalan"
    ExcelSheet.Cells(13, 7).value = "Middle Name"
    ExcelSheet.Cells(13, 8).value = "(Division)"
    ExcelSheet.Cells(13, 10).value = "(School)"
    ExcelSheet.Cells(14, 2).value = "(Surname)"
    ExcelSheet.Cells(14, 5).value = "(Given)"
    ExcelSheet.Cells(16, 1).value = "Kasarian"
     ExcelSheet.Cells(16, 2).value = public_rs.Fields("Gender").value
    ExcelSheet.Cells(16, 5).value = "Petsa ng Kapanganakan"
    ExcelSheet.Cells(16, 6).value = public_rs.Fields("bday").value
    ExcelSheet.Cells(16, 8).value = "Pook"
    ExcelSheet.Cells(16, 9).value = public_rs.Fields("birthplace").value
    temp = public_rs.Fields("student_id").value
    Call mysql_select(public_rs2, "SELECT * FROM tbl_student_level WHERE ID = '" & temp & "' ORDER BY SY ASC LIMIT 1")
    ExcelSheet.Cells(16, 10).value = "Petsa ng Pagpasok"
    ExcelSheet.Cells(16, 13).value = "June " & public_rs2.Fields("SY").value
    ExcelSheet.Cells(17, 1).value = "(Sex)"
    ExcelSheet.Cells(17, 3).value = "(Date of Birth)"
    ExcelSheet.Cells(17, 8).value = "(Place) Bayan/Lalawigan/Lungsod"
    ExcelSheet.Cells(17, 10).value = "(Date of Entrance)"
    ExcelSheet.Cells(19, 1).value = "Magulang/Tagapag-alaga"
    ExcelSheet.Cells(19, 2).value = public_rs.Fields("guardian").value
    ExcelSheet.Cells(19, 7).value = public_rs.Fields("Address").value
    ExcelSheet.Cells(19, 11).value = public_rs.Fields("occupation").value
    ExcelSheet.Cells(20, 1).value = "(Parent/Guardian)"
    ExcelSheet.Cells(20, 2).value = "Pangalan"
    ExcelSheet.Cells(20, 7).value = "Tirahan"
    ExcelSheet.Cells(20, 11).value = "Hanapbuhay"
    ExcelSheet.Cells(21, 2).value = "(Name)"
    ExcelSheet.Cells(21, 7).value = "(Address)"
    ExcelSheet.Cells(21, 11).value = "(Occupation)"
    ExcelSheet.Cells(22, 1).value = "PAG-UNLAD SA MABABANG PAARALAN"
    ExcelSheet.Cells(23, 1).value = "ELEMENTARY SCHOOL PROGRESS"
    ExcelSheet.Range("H12:H12").ColumnWidth = 8
    excelApp.Range("I12:I12").HorizontalAlignment = xlCenter
    ExcelSheet.Range("E12:F12").Merge
    ExcelSheet.Range("B12:D12").Merge
    excelApp.Range("E12:F12").HorizontalAlignment = xlCenter
    excelApp.Range("B12:D12").HorizontalAlignment = xlCenter
    ExcelSheet.Range("E13:F13").Merge
    ExcelSheet.Range("B13:D13").Merge
     ExcelSheet.Range("E14:F14").Merge
    ExcelSheet.Range("B14:D14").Merge
    ExcelSheet.Range("C16:E16").Merge
    ExcelSheet.Range("F16:G16").Merge
    ExcelSheet.Range("J16:L16").Merge
     ExcelSheet.Range("M16:O16").Merge
      ExcelSheet.Range("B17:E17").Merge
       ExcelSheet.Range("H17:I17").Merge
        ExcelSheet.Range("J17:L17").Merge
         ExcelSheet.Range("B19:F19").Merge
          ExcelSheet.Range("G19:J19").Merge
           ExcelSheet.Range("K19:O19").Merge
            ExcelSheet.Range("B20:F20").Merge
             ExcelSheet.Range("G20:J20").Merge
              ExcelSheet.Range("K20:O20").Merge
                ExcelSheet.Range("B21:F21").Merge
             ExcelSheet.Range("G21:J21").Merge
              ExcelSheet.Range("K21:O21").Merge
     excelApp.Range("B17:E17").HorizontalAlignment = xlCenter
    excelApp.Range("J17:L17").HorizontalAlignment = xlCenter
     excelApp.Range("E12:F12").HorizontalAlignment = xlCenter
    excelApp.Range("A20:A20").HorizontalAlignment = xlCenter
     excelApp.Range("E12:F12").HorizontalAlignment = xlCenter
    excelApp.Range("B19:O12").HorizontalAlignment = xlCenter
     excelApp.Range("B20:O20").HorizontalAlignment = xlCenter
    excelApp.Range("B21:O21").HorizontalAlignment = xlCenter
    excelApp.Range("A22:O22").HorizontalAlignment = xlCenter
    excelApp.Range("A23:O23").HorizontalAlignment = xlCenter
    excelApp.Range("A22:O22").Font.Size = 12
    excelApp.Range("A23:O23").Font.Size = 12
    excelApp.Range("E12:G12").Font.Bold = False
    excelApp.Range("I12:O12").Font.Underline = False
    excelApp.Range("I12:I12").Font.Bold = False
     excelApp.Range("L12:O12").Font.Bold = False
      excelApp.Range("E13:O13").Font.Bold = False
        excelApp.Range("A14:O14").Font.Bold = False
         excelApp.Range("F16:G16").Font.Bold = False
          excelApp.Range("I16:I16").Font.Bold = False
           excelApp.Range("M16:O16").Font.Bold = False
            excelApp.Range("A17:O17").Font.Bold = False
             excelApp.Range("B19:F19").Font.Bold = False
              excelApp.Range("G19:J19").Font.Bold = False
               excelApp.Range("K19:O19").Font.Bold = False
               excelApp.Range("A20:O20").Font.Bold = False
                excelApp.Range("A21:O21").Font.Bold = False
             
    
    
    ExcelSheet.Cells(25, 1).value = "Kindergarten- School"
    excelApp.Range("B25:G25").Merge
    excelApp.Range("B26:G26").Merge
    ExcelSheet.Cells(25, 9).value = "Grade I - School"
    ExcelSheet.Range("I25").ColumnWidth = 30
    excelApp.Range("J25:O25").Merge
    excelApp.Range("J26:O26").Merge
    ExcelSheet.Cells(26, 1).value = "School Year"
    excelApp.Range("A26:A26").HorizontalAlignment = xlRight
    ExcelSheet.Cells(26, 9).value = "School Year"
    excelApp.Range("I26:I26").HorizontalAlignment = xlRight
    excelApp.Range("A26:O26").Font.Bold = True
    ExcelSheet.Cells(27, 1).value = "Learning Areas"
     ExcelSheet.Cells(27, 9).value = "Learning Areas"
      ExcelSheet.Cells(27, 2).value = "1"
     ExcelSheet.Cells(27, 3).value = "2"
      ExcelSheet.Cells(27, 4).value = "3"
       ExcelSheet.Cells(27, 5).value = "4"
        ExcelSheet.Cells(27, 6).value = "Final Rating"
         ExcelSheet.Cells(27, 7).value = "Remarks"
         ExcelSheet.Cells(27, 10).value = "1"
     ExcelSheet.Cells(27, 11).value = "2"
      ExcelSheet.Cells(27, 12).value = "3"
       ExcelSheet.Cells(27, 13).value = "4"
        ExcelSheet.Cells(27, 14).value = "Final Rating"
         ExcelSheet.Cells(27, 15).value = "Remarks"
    excelApp.Range("A27:O27").Font.Bold = True
    excelApp.Range("A27:O27").HorizontalAlignment = xlCenter
    ExcelSheet.Range("F27").ColumnWidth = 12
    ExcelSheet.Range("G27").ColumnWidth = 17
    ExcelSheet.Range("N27").ColumnWidth = 12
    ExcelSheet.Range("O27").ColumnWidth = 17
    excelApp.Range("I25:I25").HorizontalAlignment = xlCenter
    ExcelSheet.Cells(40, 1).value = "AVERAGE"
    excelApp.Range("A40:A40").HorizontalAlignment = xlRight
     ExcelSheet.Cells(40, 9).value = "AVERAGE"
    excelApp.Range("I40:I40").HorizontalAlignment = xlRight
    ExcelSheet.Cells(41, 1).value = "Teacher's Signature"
    ExcelSheet.Cells(41, 9).value = "Teacher's Signature"
    ExcelSheet.Cells(42, 1).value = "Eligible for Admission to"
     ExcelSheet.Cells(42, 9).value = "Eligible for Admission to"
      excelApp.Range("B42:G42").Merge
       excelApp.Range("J42:O42").Merge
       
       
       ExcelSheet.Cells(45, 1).value = "Grade II - School"
    excelApp.Range("B45:G45").Merge
    excelApp.Range("B46:G46").Merge
    ExcelSheet.Cells(45, 9).value = "Grade III - School"
    ExcelSheet.Range("I45").ColumnWidth = 30
    excelApp.Range("J45:O45").Merge
    excelApp.Range("J46:O46").Merge
    ExcelSheet.Cells(46, 1).value = "School Year"
    excelApp.Range("A46:A46").HorizontalAlignment = xlRight
    ExcelSheet.Cells(46, 9).value = "School Year"
    excelApp.Range("I46:I46").HorizontalAlignment = xlRight
    excelApp.Range("A46:P46").Font.Bold = True
    ExcelSheet.Cells(47, 1).value = "Learning Areas"
     ExcelSheet.Cells(47, 9).value = "Learning Areas"
      ExcelSheet.Cells(47, 2).value = "1"
     ExcelSheet.Cells(47, 3).value = "2"
      ExcelSheet.Cells(47, 4).value = "3"
       ExcelSheet.Cells(47, 5).value = "4"
        ExcelSheet.Cells(47, 6).value = "Final Rating"
         ExcelSheet.Cells(47, 7).value = "Remarks"
         ExcelSheet.Cells(47, 10).value = "1"
     ExcelSheet.Cells(47, 11).value = "2"
      ExcelSheet.Cells(47, 12).value = "3"
       ExcelSheet.Cells(47, 13).value = "4"
        ExcelSheet.Cells(47, 14).value = "Final Rating"
         ExcelSheet.Cells(47, 15).value = "Remarks"
    excelApp.Range("A47:O47").Font.Bold = True
    excelApp.Range("A47:O47").HorizontalAlignment = xlCenter
    ExcelSheet.Range("F47").ColumnWidth = 12
    ExcelSheet.Range("G47").ColumnWidth = 12
    ExcelSheet.Range("N47").ColumnWidth = 12
    ExcelSheet.Range("O47").ColumnWidth = 12
    excelApp.Range("I45:I45").HorizontalAlignment = xlCenter
    ExcelSheet.Cells(60, 1).value = "AVERAGE"
    excelApp.Range("A60:A60").HorizontalAlignment = xlRight
     ExcelSheet.Cells(60, 9).value = "AVERAGE"
    excelApp.Range("I60:I60").HorizontalAlignment = xlRight
    ExcelSheet.Cells(61, 1).value = "Teacher's Signature"
    ExcelSheet.Cells(61, 9).value = "Teacher's Signature"
    ExcelSheet.Cells(62, 1).value = "Eligible for Admission to"
     ExcelSheet.Cells(62, 9).value = "Eligible for Admission to"
      excelApp.Range("B62:G62").Merge
       excelApp.Range("J62:O62").Merge
       
     ExcelSheet.Cells(65, 1).value = "Grade IV - School"
    excelApp.Range("B65:G65").Merge
    excelApp.Range("B66:G66").Merge
    ExcelSheet.Cells(65, 9).value = "Grade V - School"
    ExcelSheet.Range("I65").ColumnWidth = 30
    excelApp.Range("J65:O65").Merge
    excelApp.Range("J66:O66").Merge
    ExcelSheet.Cells(66, 1).value = "School Year"
    excelApp.Range("A66:A66").HorizontalAlignment = xlRight
    ExcelSheet.Cells(66, 9).value = "School Year"
    excelApp.Range("I66:I66").HorizontalAlignment = xlRight
    excelApp.Range("A66:O66").Font.Bold = True
    ExcelSheet.Cells(67, 1).value = "Learning Areas"
     ExcelSheet.Cells(67, 9).value = "Learning Areas"
      ExcelSheet.Cells(67, 2).value = "1"
     ExcelSheet.Cells(67, 3).value = "2"
      ExcelSheet.Cells(67, 4).value = "3"
       ExcelSheet.Cells(67, 5).value = "4"
        ExcelSheet.Cells(67, 6).value = "Final Rating"
         ExcelSheet.Cells(67, 7).value = "Remarks"
         ExcelSheet.Cells(67, 10).value = "1"
     ExcelSheet.Cells(67, 11).value = "2"
      ExcelSheet.Cells(67, 12).value = "3"
       ExcelSheet.Cells(67, 13).value = "4"
        ExcelSheet.Cells(67, 14).value = "Final Rating"
         ExcelSheet.Cells(67, 15).value = "Remarks"
    excelApp.Range("A67:O67").Font.Bold = True
    excelApp.Range("A67:O67").HorizontalAlignment = xlCenter
    ExcelSheet.Range("F67").ColumnWidth = 12
    ExcelSheet.Range("G67").ColumnWidth = 12
    ExcelSheet.Range("N67").ColumnWidth = 12
    ExcelSheet.Range("O67").ColumnWidth = 12
    excelApp.Range("I65:I65").HorizontalAlignment = xlCenter
    ExcelSheet.Cells(80, 1).value = "AVERAGE"
    excelApp.Range("A80:A80").HorizontalAlignment = xlRight
     ExcelSheet.Cells(80, 9).value = "AVERAGE"
    excelApp.Range("I80:I80").HorizontalAlignment = xlRight
    ExcelSheet.Cells(81, 1).value = "Teacher's Signature"
    ExcelSheet.Cells(81, 9).value = "Teacher's Signature"
    ExcelSheet.Cells(82, 1).value = "Eligible for Admission to"
     ExcelSheet.Cells(82, 9).value = "Eligible for Admission to"
      excelApp.Range("B82:G82").Merge
       excelApp.Range("J82:O82").Merge
       
    ExcelSheet.Cells(85, 1).value = "Grade IV - School"
    excelApp.Range("A85:O85").Font.Bold = True
    excelApp.Range("B85:G85").Merge
    excelApp.Range("B86:G86").Merge
    ExcelSheet.Cells(86, 1).value = "School Year"
     excelApp.Range("A86:A86").HorizontalAlignment = xlRight
     excelApp.Range("A86:O86").Font.Bold = True
    ExcelSheet.Cells(87, 1).value = "Learning Areas"
     ExcelSheet.Cells(87, 2).value = "1"
     ExcelSheet.Cells(87, 3).value = "2"
      ExcelSheet.Cells(87, 4).value = "3"
       ExcelSheet.Cells(87, 5).value = "4"
        ExcelSheet.Cells(87, 6).value = "Final Rating"
         ExcelSheet.Cells(87, 7).value = "Remarks"
         excelApp.Range("A87:O87").Font.Bold = True
    excelApp.Range("A87:O87").HorizontalAlignment = xlCenter
    ExcelSheet.Range("F87").ColumnWidth = 12
    ExcelSheet.Range("G87").ColumnWidth = 12
     excelApp.Range("A100:A100").HorizontalAlignment = xlRight
     ExcelSheet.Cells(100, 1).value = "AVERAGE"
      ExcelSheet.Cells(101, 1).value = "Teacher's Signature"
      ExcelSheet.Cells(102, 1).value = "Eligible for Admission to"
     excelApp.Range("B102:G102").Merge
     excelApp.Range("A25:O102").Font.Bold = False
         
     Dim kinder, sy_1, sy_2, sy_3 As String
    Call mysql_select(public_rs, "SELECT * FROM tbl_student_level WHERE ID = '" & lbl_id.Caption & "' AND lvl_name = 'Kinder'")
    If public_rs.RecordCount = 0 Then
        ExcelSheet.Cells(26, 2).value = ""
        ExcelSheet.Cells(25, 2).value = ""
        kinder = ""
    Else
        ExcelSheet.Cells(26, 2).value = public_rs.Fields("SY").value
        ExcelSheet.Cells(25, 2).value = "Manuel S. Rojas Elementary School"
        kinder = public_rs.Fields("SY").value
    End If
    Dim ctr As Integer
    ctr = 1
    Call mysql_select(public_all, "SELECT DISTINCT a.subject_code,  b.subject_name as Subject, a.Grade, a.Remark FROM tbl_grade a LEFT JOIN tbl_subject b ON a.subject_code=b.subject_code WHERE a.ID = '" & lbl_id.Caption & "' AND a.SY = '" & kinder & "' AND a.Period='1st Grading' ORDER BY subject_code ASC ")
                 While Not public_all.EOF
                    subject(ctr) = public_all.Fields("Subject").value
                    first(ctr) = public_all.Fields("Remark").value
                    ctr = ctr + 1
                    public_all.MoveNext
                Wend
                
                    ExcelSheet.Cells(28, 1).value = subject(1)
                    ExcelSheet.Cells(29, 1).value = subject(2)
                    ExcelSheet.Cells(30, 1).value = subject(3)
                    ExcelSheet.Cells(31, 1).value = subject(4)
                   ExcelSheet.Cells(32, 1).value = subject(5)
                    ExcelSheet.Cells(33, 1).value = subject(6)
                    ExcelSheet.Cells(34, 1).value = subject(7)
                   ExcelSheet.Cells(35, 1).value = subject(8)
                    ExcelSheet.Cells(36, 1).value = subject(9)
                    ExcelSheet.Cells(37, 1).value = subject(10)
                    
                    ExcelSheet.Cells(28, 2).value = first(1)
                    ExcelSheet.Cells(29, 2).value = first(2)
                    ExcelSheet.Cells(30, 2).value = first(3)
                    ExcelSheet.Cells(31, 2).value = first(4)
                    ExcelSheet.Cells(32, 2).value = first(5)
                    ExcelSheet.Cells(33, 2).value = first(6)
                    ExcelSheet.Cells(34, 2).value = first(7)
                    ExcelSheet.Cells(35, 2).value = first(8)
                    ExcelSheet.Cells(36, 2).value = first(9)
                    ExcelSheet.Cells(37, 2).value = first(10)
                   
               Call mysql_select(public_all2, "SELECT DISTINCT subject_code, Grade, Remark FROM tbl_grade WHERE ID = '" & lbl_id.Caption & "' AND SY = '" & kinder & "' AND Period='2nd Grading' ORDER BY subject_code ASC  ")
                 
                    ctr = 1
                 While Not public_all2.EOF
                    If public_all2.RecordCount = 0 Then
                         sec(ctr) = "No grade"
                    Else
                         sec(ctr) = public_all2.Fields("Remark").value
                    End If
                   
                    ctr = ctr + 1
                    public_all2.MoveNext
                Wend
               
                
                    ExcelSheet.Cells(28, 3).value = sec(1)
                    ExcelSheet.Cells(29, 3).value = sec(2)
                    ExcelSheet.Cells(30, 3).value = sec(3)
                    ExcelSheet.Cells(31, 3).value = sec(4)
                    ExcelSheet.Cells(32, 3).value = sec(5)
                    ExcelSheet.Cells(33, 3).value = sec(6)
                    ExcelSheet.Cells(34, 3).value = sec(7)
                    ExcelSheet.Cells(35, 3).value = sec(8)
                    ExcelSheet.Cells(36, 3).value = sec(9)
                   ExcelSheet.Cells(37, 3).value = sec(10)
                    
                    Call mysql_select(public_all2, "SELECT DISTINCT subject_code, Grade, Remark FROM tbl_grade WHERE ID = '" & lbl_id.Caption & "' AND SY = '" & kinder & "' AND Period='3rd Grading' ORDER BY subject_code ASC  ")
                ctr = 1
                 While Not public_all2.EOF
                    If public_all2.RecordCount = 0 Then
                         third(ctr) = "No grade"
                    Else
                         third(ctr) = public_all2.Fields("Remark").value
                    End If
                   
                    ctr = ctr + 1
                    public_all2.MoveNext
                Wend
                    ExcelSheet.Cells(28, 4).value = third(1)
                    ExcelSheet.Cells(29, 4).value = third(2)
                    ExcelSheet.Cells(30, 4).value = third(3)
                    ExcelSheet.Cells(31, 4).value = third(4)
                    ExcelSheet.Cells(32, 4).value = third(5)
                    ExcelSheet.Cells(33, 4).value = third(6)
                    ExcelSheet.Cells(34, 4).value = third(7)
                    ExcelSheet.Cells(35, 4).value = third(8)
                    ExcelSheet.Cells(36, 4).value = third(9)
                    ExcelSheet.Cells(37, 4).value = third(10)
                    
                    Call mysql_select(public_all2, "SELECT DISTINCT subject_code, Grade, Remark FROM tbl_grade WHERE ID = '" & lbl_id.Caption & "' AND SY = '" & kinder & "' AND Period='4th Grading' ORDER BY subject_code ASC  ")
                ctr = 1
                While Not public_all2.EOF
                    If public_all2.RecordCount = 0 Then
                         fourth(ctr) = "No grade"
                    Else
                         fourth(ctr) = public_all2.Fields("Remark").value
                    End If
                   
                    ctr = ctr + 1
                    public_all2.MoveNext
                Wend
                    ExcelSheet.Cells(28, 5).value = fourth(1)
                     ExcelSheet.Cells(29, 5).value = fourth(2)
                     ExcelSheet.Cells(30, 5).value = fourth(3)
                     ExcelSheet.Cells(31, 5).value = fourth(4)
                     ExcelSheet.Cells(32, 5).value = fourth(5)
                     ExcelSheet.Cells(33, 5).value = fourth(6)
                     ExcelSheet.Cells(34, 5).value = fourth(7)
                     ExcelSheet.Cells(35, 5).value = fourth(8)
                     ExcelSheet.Cells(36, 5).value = fourth(9)
                     ExcelSheet.Cells(37, 5).value = fourth(10)
                    
                      Call mysql_select(public_all2, "SELECT DISTINCT subject_code, Grade, Remark FROM tbl_grade WHERE ID = '" & lbl_id.Caption & "' AND SY = '" & kinder & "' AND Period='Final' ORDER BY subject_code ASC ")
                ctr = 1
                While Not public_all2.EOF
                    If public_all2.RecordCount = 0 Then
                         f(ctr) = "No grade"
                         f2(ctr) = "0"
                    Else
                         f(ctr) = public_all2.Fields("Remark").value
                         If f(ctr) <> "B" Then
                            f2(ctr) = "Promote to Grade I"
                        Else
                            f2(ctr) = "Unable to Promote"
                         End If
                    End If
                   
                    ctr = ctr + 1
                    public_all2.MoveNext
                Wend
                    ExcelSheet.Cells(28, 6).value = f(1)
                    ExcelSheet.Cells(29, 6).value = f(2)
                    ExcelSheet.Cells(30, 6).value = f(3)
                    ExcelSheet.Cells(31, 6).value = f(4)
                    ExcelSheet.Cells(32, 6).value = f(5)
                    ExcelSheet.Cells(33, 6).value = f(6)
                    ExcelSheet.Cells(34, 6).value = f(7)
                    ExcelSheet.Cells(35, 6).value = f(8)
                    ExcelSheet.Cells(36, 6).value = f(9)
                    ExcelSheet.Cells(37, 6).value = f(10)
                    
                    ExcelSheet.Cells(28, 7).value = f2(1)
                    ExcelSheet.Cells(29, 7).value = f2(2)
                    ExcelSheet.Cells(30, 7).value = f2(3)
                    ExcelSheet.Cells(31, 7).value = f2(4)
                    ExcelSheet.Cells(32, 7).value = f2(5)
                    ExcelSheet.Cells(33, 7).value = f2(6)
                    ExcelSheet.Cells(34, 7).value = f2(7)
                    ExcelSheet.Cells(35, 7).value = f2(8)
                    ExcelSheet.Cells(36, 7).value = f2(9)
                    ExcelSheet.Cells(37, 7).value = f2(10)
                    
                                    Call mysql_select(public_rs, "SELECT subject_code as Subject, period, grade, remark FROM tbl_grade WHERE ID = '" & lbl_id.Caption & "' AND SY = '" & kinder & "' AND Period='1st Grading'")
        no_subj = public_rs.RecordCount
        If no_subj = 0 Then
            
            average = 0
            remark = "No grades"
        Else
            average = 0
            While Not public_rs.EOF
                average = val(public_rs.Fields("grade")) + average
                public_rs.MoveNext
            Wend
            average = average / no_subj
            average = Round(average, 2)
            If average >= 90 Then
              remark = "A"
              ElseIf average >= 85 Then
              remark = "P"
              ElseIf average >= 80 Then
                  remark = "AP"
              ElseIf average >= 74 Then
                  remark = "D"
              Else
                  remark = "B"
            End If
            ExcelSheet.Cells(40, 2).value = remark
        End If
    
        Call mysql_select(public_rs, "SELECT subject_code as Subject, period, grade, remark FROM tbl_grade WHERE ID = '" & lbl_id.Caption & "' AND SY = '" & kinder & "' AND Period='2nd Grading'")
        no_subj = public_rs.RecordCount
        If no_subj = 0 Then
           
            average = 0
            remark = "No grades"
        Else
            average = 0
            While Not public_rs.EOF
                average = val(public_rs.Fields("grade")) + average
                public_rs.MoveNext
            Wend
            average = average / no_subj
            average = Round(average, 2)
            If average >= 90 Then
              remark = "A"
              ElseIf average >= 85 Then
              remark = "P"
              ElseIf average >= 80 Then
                  remark = "AP"
              ElseIf average >= 74 Then
                  remark = "D"
              Else
                  remark = "B"
            End If
            ExcelSheet.Cells(40, 3).value = remark
        End If
        Call mysql_select(public_rs, "SELECT subject_code as Subject, period, grade, remark FROM tbl_grade WHERE ID = '" & lbl_id.Caption & "' AND SY = '" & kinder & "' AND Period='3rd Grading'")
        no_subj = public_rs.RecordCount
        If no_subj = 0 Then
            
            average = 0
            remark = "No grades"
        Else
            average = 0
            While Not public_rs.EOF
                average = val(public_rs.Fields("grade")) + average
                public_rs.MoveNext
            Wend
            average = average / no_subj
            average = Round(average, 2)
            If average >= 90 Then
              remark = "A"
              ElseIf average >= 85 Then
              remark = "P"
              ElseIf average >= 80 Then
                  remark = "AP"
              ElseIf average >= 74 Then
                  remark = "D"
              Else
                  remark = "B"
            End If
            ExcelSheet.Cells(40, 4).value = remark
        End If
        Call mysql_select(public_rs, "SELECT subject_code as Subject, period, grade, remark FROM tbl_grade WHERE ID = '" & lbl_id.Caption & "' AND SY = '" & kinder & "' AND Period='4th Grading'")
        no_subj = public_rs.RecordCount
        If no_subj = 0 Then
            
            average = 0
            remark = "No grades"
        Else
            average = 0
            While Not public_rs.EOF
                average = val(public_rs.Fields("grade")) + average
                public_rs.MoveNext
            Wend
            average = average / no_subj
            average = Round(average, 2)
            If average >= 90 Then
              remark = "A"
              ElseIf average >= 85 Then
              remark = "P"
              ElseIf average >= 80 Then
                  remark = "AP"
              ElseIf average >= 74 Then
                  remark = "D"
              Else
                  remark = "B"
            End If
            ExcelSheet.Cells(40, 5).value = remark
        End If
        Call mysql_select(public_rs, "SELECT subject_code as Subject, period, grade, remark FROM tbl_grade WHERE ID = '" & lbl_id.Caption & "' AND SY = '" & kinder & "' AND Period='Final'")
        no_subj = public_rs.RecordCount
        If no_subj = 0 Then
            
            average = 0
            remark = "No grades"
        Else
            average = 0
            While Not public_rs.EOF
                average = val(public_rs.Fields("grade")) + average
                public_rs.MoveNext
            Wend
            average = average / no_subj
            average = Round(average, 2)
            If average >= 90 Then
              remark = "A"
              ElseIf average >= 85 Then
              remark = "P"
              ElseIf average >= 80 Then
                  remark = "AP"
              ElseIf average >= 74 Then
                  remark = "D"
              Else
                  remark = "B"
            End If
            ExcelSheet.Cells(40, 6).value = remark
            Dim remark2 As String
            If remark <> "B" Then
                ExcelSheet.Cells(40, 7).value = "Promote to Grade I"
            Else
                ExcelSheet.Cells(40, 7).value = "Unable to Promote"
            End If
            
            
            If remark = "B" Then
                 ExcelSheet.Cells(42, 2).value = "Unable to promote"
            Else
                ExcelSheet.Cells(42, 2).value = "Grade I"
            End If
            
            
        End If
    
    
    
     Call mysql_select(public_rs, "SELECT * FROM tbl_student_level WHERE ID = '" & lbl_id.Caption & "' AND lvl_name = 'Grade 1'")
    If public_rs.RecordCount = 0 Then
        ExcelSheet.Cells(26, 10).value = ""
        ExcelSheet.Cells(25, 10).value = ""
        sy_1 = ""
    Else
        ExcelSheet.Cells(26, 10).value = public_rs.Fields("SY").value
        ExcelSheet.Cells(25, 10).value = "Manuel S. Rojas Elementary School"
        sy_1 = public_rs.Fields("SY").value
    End If
    
    ctr = 1
    Call mysql_select(public_all, "SELECT DISTINCT a.subject_code,  b.subject_name as Subject, a.Grade, a.Remark FROM tbl_grade a LEFT JOIN tbl_subject b ON a.subject_code=b.subject_code WHERE a.ID = '" & lbl_id.Caption & "' AND a.SY = '" & sy_1 & "' AND a.Period='1st Grading' ORDER BY subject_code ASC ")
                 While Not public_all.EOF
                    subject(ctr) = public_all.Fields("Subject").value
                    first(ctr) = public_all.Fields("Remark").value
                    ctr = ctr + 1
                    public_all.MoveNext
                Wend
                
                    ExcelSheet.Cells(28, 9).value = subject(1)
                    ExcelSheet.Cells(29, 9).value = subject(2)
                    ExcelSheet.Cells(30, 9).value = subject(3)
                    ExcelSheet.Cells(31, 9).value = subject(4)
                   ExcelSheet.Cells(32, 9).value = subject(5)
                    ExcelSheet.Cells(33, 9).value = subject(6)
                    ExcelSheet.Cells(34, 9).value = subject(7)
                   ExcelSheet.Cells(35, 9).value = subject(8)
                    ExcelSheet.Cells(36, 9).value = subject(9)
                    ExcelSheet.Cells(37, 9).value = subject(10)
                    
                    ExcelSheet.Cells(28, 10).value = first(1)
                    ExcelSheet.Cells(29, 10).value = first(2)
                    ExcelSheet.Cells(30, 10).value = first(3)
                    ExcelSheet.Cells(31, 10).value = first(4)
                    ExcelSheet.Cells(32, 10).value = first(5)
                    ExcelSheet.Cells(33, 10).value = first(6)
                    ExcelSheet.Cells(34, 10).value = first(7)
                    ExcelSheet.Cells(35, 10).value = first(8)
                    ExcelSheet.Cells(36, 10).value = first(9)
                    ExcelSheet.Cells(37, 10).value = first(10)
                   
               Call mysql_select(public_all2, "SELECT DISTINCT subject_code, Grade, Remark FROM tbl_grade WHERE ID = '" & lbl_id.Caption & "' AND SY = '" & sy_1 & "' AND Period='2nd Grading' ORDER BY subject_code ASC  ")
                 
                    ctr = 1
                 While Not public_all2.EOF
                    If public_all2.RecordCount = 0 Then
                         sec(ctr) = "No grade"
                    Else
                         sec(ctr) = public_all2.Fields("Remark").value
                    End If
                   
                    ctr = ctr + 1
                    public_all2.MoveNext
                Wend
               
                
                    ExcelSheet.Cells(28, 11).value = sec(1)
                    ExcelSheet.Cells(29, 11).value = sec(2)
                    ExcelSheet.Cells(30, 11).value = sec(3)
                    ExcelSheet.Cells(31, 11).value = sec(4)
                    ExcelSheet.Cells(32, 11).value = sec(5)
                    ExcelSheet.Cells(33, 11).value = sec(6)
                    ExcelSheet.Cells(34, 11).value = sec(7)
                    ExcelSheet.Cells(35, 11).value = sec(8)
                    ExcelSheet.Cells(36, 11).value = sec(9)
                   ExcelSheet.Cells(37, 11).value = sec(10)
                    
                    Call mysql_select(public_all2, "SELECT DISTINCT subject_code, Grade, Remark FROM tbl_grade WHERE ID = '" & lbl_id.Caption & "' AND SY = '" & sy_1 & "' AND Period='3rd Grading' ORDER BY subject_code ASC  ")
                ctr = 1
                 While Not public_all2.EOF
                    If public_all2.RecordCount = 0 Then
                         third(ctr) = "No grade"
                    Else
                         third(ctr) = public_all2.Fields("Remark").value
                    End If
                   
                    ctr = ctr + 1
                    public_all2.MoveNext
                Wend
                    ExcelSheet.Cells(28, 12).value = third(1)
                    ExcelSheet.Cells(29, 12).value = third(2)
                    ExcelSheet.Cells(30, 12).value = third(3)
                    ExcelSheet.Cells(31, 12).value = third(4)
                    ExcelSheet.Cells(32, 12).value = third(5)
                    ExcelSheet.Cells(33, 12).value = third(6)
                    ExcelSheet.Cells(34, 12).value = third(7)
                    ExcelSheet.Cells(35, 12).value = third(8)
                    ExcelSheet.Cells(36, 12).value = third(9)
                    ExcelSheet.Cells(37, 12).value = third(10)
                    
                    Call mysql_select(public_all2, "SELECT DISTINCT subject_code, Grade, Remark FROM tbl_grade WHERE ID = '" & lbl_id.Caption & "' AND SY = '" & sy_1 & "' AND Period='4th Grading' ORDER BY subject_code ASC  ")
                ctr = 1
                While Not public_all2.EOF
                    If public_all2.RecordCount = 0 Then
                         fourth(ctr) = "No grade"
                    Else
                         fourth(ctr) = public_all2.Fields("Remark").value
                    End If
                   
                    ctr = ctr + 1
                    public_all2.MoveNext
                Wend
                    ExcelSheet.Cells(28, 13).value = fourth(1)
                     ExcelSheet.Cells(29, 13).value = fourth(2)
                     ExcelSheet.Cells(30, 13).value = fourth(3)
                     ExcelSheet.Cells(31, 13).value = fourth(4)
                     ExcelSheet.Cells(32, 13).value = fourth(5)
                     ExcelSheet.Cells(33, 13).value = fourth(6)
                     ExcelSheet.Cells(34, 13).value = fourth(7)
                     ExcelSheet.Cells(35, 13).value = fourth(8)
                     ExcelSheet.Cells(36, 13).value = fourth(9)
                     ExcelSheet.Cells(37, 13).value = fourth(10)
                    
                      Call mysql_select(public_all2, "SELECT DISTINCT subject_code, Grade, Remark FROM tbl_grade WHERE ID = '" & lbl_id.Caption & "' AND SY = '" & sy_1 & "' AND Period='Final' ORDER BY subject_code ASC ")
                ctr = 1
                While Not public_all2.EOF
                    If public_all2.RecordCount = 0 Then
                         f(ctr) = "No grade"
                         f2(ctr) = "0"
                    Else
                         
                          f(ctr) = public_all2.Fields("Remark").value
                         If f(ctr) <> "B" Then
                            f2(ctr) = "Promote to Grade II"
                        Else
                            f2(ctr) = "Unable to Promote"
                         End If
                    End If
                   
                    ctr = ctr + 1
                    public_all2.MoveNext
                Wend
                    ExcelSheet.Cells(28, 14).value = f(1)
                    ExcelSheet.Cells(29, 14).value = f(2)
                    ExcelSheet.Cells(30, 14).value = f(3)
                    ExcelSheet.Cells(31, 14).value = f(4)
                    ExcelSheet.Cells(32, 14).value = f(5)
                    ExcelSheet.Cells(33, 14).value = f(6)
                    ExcelSheet.Cells(34, 14).value = f(7)
                    ExcelSheet.Cells(35, 14).value = f(8)
                    ExcelSheet.Cells(36, 14).value = f(9)
                    ExcelSheet.Cells(37, 14).value = f(10)
                    
                    ExcelSheet.Cells(28, 15).value = f2(1)
                    ExcelSheet.Cells(29, 15).value = f2(2)
                    ExcelSheet.Cells(30, 15).value = f2(3)
                    ExcelSheet.Cells(31, 15).value = f2(4)
                    ExcelSheet.Cells(32, 15).value = f2(5)
                    ExcelSheet.Cells(33, 15).value = f2(6)
                    ExcelSheet.Cells(34, 15).value = f2(7)
                    ExcelSheet.Cells(35, 15).value = f2(8)
                    ExcelSheet.Cells(36, 15).value = f2(9)
                    ExcelSheet.Cells(37, 15).value = f2(10)
                    
                     Call mysql_select(public_rs, "SELECT subject_code as Subject, period, grade, remark FROM tbl_grade WHERE ID = '" & lbl_id.Caption & "' AND SY = '" & sy_1 & "' AND Period='1st Grading'")
        no_subj = public_rs.RecordCount
        If no_subj = 0 Then
            
            average = 0
            remark = "No grades"
        Else
            average = 0
            While Not public_rs.EOF
                average = val(public_rs.Fields("grade")) + average
                public_rs.MoveNext
            Wend
            average = average / no_subj
            average = Round(average, 2)
            If average >= 90 Then
              remark = "A"
              ElseIf average >= 85 Then
              remark = "P"
              ElseIf average >= 80 Then
                  remark = "AP"
              ElseIf average >= 74 Then
                  remark = "D"
              Else
                  remark = "B"
            End If
            ExcelSheet.Cells(40, 10).value = remark
        End If
    
        Call mysql_select(public_rs, "SELECT subject_code as Subject, period, grade, remark FROM tbl_grade WHERE ID = '" & lbl_id.Caption & "' AND SY = '" & sy_1 & "' AND Period='2nd Grading'")
        no_subj = public_rs.RecordCount
        If no_subj = 0 Then
            
            average = 0
            remark = "No grades"
        Else
            average = 0
            While Not public_rs.EOF
                average = val(public_rs.Fields("grade")) + average
                public_rs.MoveNext
            Wend
            average = average / no_subj
            average = Round(average, 2)
            If average >= 90 Then
              remark = "A"
              ElseIf average >= 85 Then
              remark = "P"
              ElseIf average >= 80 Then
                  remark = "AP"
              ElseIf average >= 74 Then
                  remark = "D"
              Else
                  remark = "B"
            End If
            ExcelSheet.Cells(40, 11).value = remark
        End If
        Call mysql_select(public_rs, "SELECT subject_code as Subject, period, grade, remark FROM tbl_grade WHERE ID = '" & lbl_id.Caption & "' AND SY = '" & sy_1 & "' AND Period='3rd Grading'")
        no_subj = public_rs.RecordCount
        If no_subj = 0 Then
            
            average = 0
            remark = "No grades"
        Else
            average = 0
            While Not public_rs.EOF
                average = val(public_rs.Fields("grade")) + average
                public_rs.MoveNext
            Wend
            average = average / no_subj
            average = Round(average, 2)
            If average >= 90 Then
              remark = "A"
              ElseIf average >= 85 Then
              remark = "P"
              ElseIf average >= 80 Then
                  remark = "AP"
              ElseIf average >= 74 Then
                  remark = "D"
              Else
                  remark = "B"
            End If
            ExcelSheet.Cells(40, 12).value = remark
        End If
        Call mysql_select(public_rs, "SELECT subject_code as Subject, period, grade, remark FROM tbl_grade WHERE ID = '" & lbl_id.Caption & "' AND SY = '" & sy_1 & "' AND Period='4th Grading'")
        no_subj = public_rs.RecordCount
        If no_subj = 0 Then
            
            average = 0
            remark = "No grades"
        Else
            average = 0
            While Not public_rs.EOF
                average = val(public_rs.Fields("grade")) + average
                public_rs.MoveNext
            Wend
            average = average / no_subj
            average = Round(average, 2)
            If average >= 90 Then
              remark = "A"
              ElseIf average >= 85 Then
              remark = "P"
              ElseIf average >= 80 Then
                  remark = "AP"
              ElseIf average >= 74 Then
                  remark = "D"
              Else
                  remark = "B"
            End If
            ExcelSheet.Cells(40, 13).value = remark
        End If
        Call mysql_select(public_rs, "SELECT subject_code as Subject, period, grade, remark FROM tbl_grade WHERE ID = '" & lbl_id.Caption & "' AND SY = '" & sy_1 & "' AND Period='Final'")
        no_subj = public_rs.RecordCount
        If no_subj = 0 Then
            
            average = 0
            remark = "No grades"
        Else
            average = 0
            While Not public_rs.EOF
                average = val(public_rs.Fields("grade")) + average
                public_rs.MoveNext
            Wend
            average = average / no_subj
            average = Round(average, 2)
            If average >= 90 Then
              remark = "A"
              ElseIf average >= 85 Then
              remark = "P"
              ElseIf average >= 80 Then
                  remark = "AP"
              ElseIf average >= 74 Then
                  remark = "D"
              Else
                  remark = "B"
            End If
            ExcelSheet.Cells(40, 14).value = remark
           If remark <> "B" Then
                ExcelSheet.Cells(40, 15).value = "Promote to Grade II"
            Else
                ExcelSheet.Cells(40, 15).value = "Unable to Promote"
            End If
            If remark = "B" Then
                 ExcelSheet.Cells(42, 10).value = "Unable to promote"
            Else
                ExcelSheet.Cells(42, 10).value = "Grade II"
            End If
            
            
        End If
    
    
    
    
    Call next_code
    
    
    
    
     If FileCheck = MyMonth + "_" + MyYear + MyExtension Then
        'Save existing workbook
        ExcelWorkbook.Save
    Else
        'Save new workbook
        ExcelWorkbook.SaveAs MyFileName
    End If

        'Close Excel
        ExcelWorkbook.Close savechanges:=False
        excelApp.Quit
        Set excelApp = Nothing
        Set ExcelWorkbook = Nothing
        Set ExcelSheet = Nothing
    MsgBox "Form 137 for Grades has been exported to an excel file."
    End If
End Sub
Public Sub next_code()
    Dim subject(10) As String
    Dim first(10) As String
    Dim sec(10) As String
    Dim third(10) As String
    Dim fourth(10) As String
    Dim f(10) As String
    Dim f2(10) As String
    Dim ctr As Integer
    Dim sy_2 As String
    
     Call mysql_select(public_rs2, "SELECT * FROM tbl_student_level WHERE ID = '" & lbl_id.Caption & "' AND lvl_name = 'Grade 2'")
    If public_rs2.RecordCount = 0 Then
        ExcelSheet.Cells(46, 2).value = ""
        ExcelSheet.Cells(45, 2).value = ""
        sy_2 = ""
    Else
        ExcelSheet.Cells(46, 2).value = public_rs2.Fields("SY").value
        ExcelSheet.Cells(45, 2).value = "Manuel S. Rojas Elementary School"
        sy_2 = public_rs2.Fields("SY").value
    End If
    
    ctr = 1
    Call mysql_select(public_all, "SELECT DISTINCT a.subject_code,  b.subject_name as Subject, a.Grade, a.Remark FROM tbl_grade a LEFT JOIN tbl_subject b ON a.subject_code=b.subject_code WHERE a.ID = '" & lbl_id.Caption & "' AND a.SY = '" & sy_2 & "' AND a.Period='1st Grading' ORDER BY subject_code ASC ")
                 While Not public_all.EOF
                    subject(ctr) = public_all.Fields("Subject").value
                    first(ctr) = public_all.Fields("Remark").value
                    ctr = ctr + 1
                    public_all.MoveNext
                Wend
                
                    ExcelSheet.Cells(48, 1).value = subject(1)
                    ExcelSheet.Cells(49, 1).value = subject(2)
                    ExcelSheet.Cells(50, 1).value = subject(3)
                    ExcelSheet.Cells(51, 1).value = subject(4)
                   ExcelSheet.Cells(52, 1).value = subject(5)
                    ExcelSheet.Cells(53, 1).value = subject(6)
                    ExcelSheet.Cells(54, 1).value = subject(7)
                   ExcelSheet.Cells(55, 1).value = subject(8)
                    ExcelSheet.Cells(56, 1).value = subject(9)
                    ExcelSheet.Cells(57, 1).value = subject(10)
                    
                    ExcelSheet.Cells(48, 2).value = first(1)
                    ExcelSheet.Cells(49, 2).value = first(2)
                    ExcelSheet.Cells(50, 2).value = first(3)
                    ExcelSheet.Cells(51, 2).value = first(4)
                    ExcelSheet.Cells(52, 2).value = first(5)
                    ExcelSheet.Cells(53, 2).value = first(6)
                    ExcelSheet.Cells(54, 2).value = first(7)
                    ExcelSheet.Cells(55, 2).value = first(8)
                    ExcelSheet.Cells(56, 2).value = first(9)
                    ExcelSheet.Cells(57, 2).value = first(10)
                   
               Call mysql_select(public_all2, "SELECT DISTINCT subject_code, Grade, Remark FROM tbl_grade WHERE ID = '" & lbl_id.Caption & "' AND SY = '" & sy_2 & "' AND Period='2nd Grading' ORDER BY subject_code ASC  ")
                 
                    ctr = 1
                 While Not public_all2.EOF
                    If public_all2.RecordCount = 0 Then
                         sec(ctr) = "No grade"
                    Else
                         sec(ctr) = public_all2.Fields("Remark").value
                    End If
                   
                    ctr = ctr + 1
                    public_all2.MoveNext
                Wend
               
                
                    ExcelSheet.Cells(48, 3).value = sec(1)
                    ExcelSheet.Cells(49, 3).value = sec(2)
                    ExcelSheet.Cells(50, 3).value = sec(3)
                    ExcelSheet.Cells(51, 3).value = sec(4)
                    ExcelSheet.Cells(52, 3).value = sec(5)
                    ExcelSheet.Cells(53, 3).value = sec(6)
                    ExcelSheet.Cells(54, 3).value = sec(7)
                    ExcelSheet.Cells(55, 3).value = sec(8)
                    ExcelSheet.Cells(56, 3).value = sec(9)
                   ExcelSheet.Cells(57, 3).value = sec(10)
                    
                    Call mysql_select(public_all2, "SELECT DISTINCT subject_code, Grade, Remark FROM tbl_grade WHERE ID = '" & lbl_id.Caption & "' AND SY = '" & sy_2 & "' AND Period='3rd Grading' ORDER BY subject_code ASC  ")
                ctr = 1
                 While Not public_all2.EOF
                    If public_all2.RecordCount = 0 Then
                         third(ctr) = "No grade"
                    Else
                         third(ctr) = public_all2.Fields("Remark").value
                    End If
                   
                    ctr = ctr + 1
                    public_all2.MoveNext
                Wend
                    ExcelSheet.Cells(48, 4).value = third(1)
                    ExcelSheet.Cells(49, 4).value = third(2)
                    ExcelSheet.Cells(50, 4).value = third(3)
                    ExcelSheet.Cells(51, 4).value = third(4)
                    ExcelSheet.Cells(52, 4).value = third(5)
                    ExcelSheet.Cells(53, 4).value = third(6)
                    ExcelSheet.Cells(54, 4).value = third(7)
                    ExcelSheet.Cells(55, 4).value = third(8)
                    ExcelSheet.Cells(56, 4).value = third(9)
                    ExcelSheet.Cells(57, 4).value = third(10)
                    
                    Call mysql_select(public_all2, "SELECT DISTINCT subject_code, Grade, Remark FROM tbl_grade WHERE ID = '" & lbl_id.Caption & "' AND SY = '" & sy_2 & "' AND Period='4th Grading' ORDER BY subject_code ASC  ")
                ctr = 1
                While Not public_all2.EOF
                    If public_all2.RecordCount = 0 Then
                         fourth(ctr) = "No grade"
                    Else
                         fourth(ctr) = public_all2.Fields("Remark").value
                    End If
                   
                    ctr = ctr + 1
                    public_all2.MoveNext
                Wend
                    ExcelSheet.Cells(48, 5).value = fourth(1)
                     ExcelSheet.Cells(49, 5).value = fourth(2)
                     ExcelSheet.Cells(50, 5).value = fourth(3)
                     ExcelSheet.Cells(51, 5).value = fourth(4)
                     ExcelSheet.Cells(52, 5).value = fourth(5)
                     ExcelSheet.Cells(53, 5).value = fourth(6)
                     ExcelSheet.Cells(54, 5).value = fourth(7)
                     ExcelSheet.Cells(55, 5).value = fourth(8)
                     ExcelSheet.Cells(56, 5).value = fourth(9)
                     ExcelSheet.Cells(57, 5).value = fourth(10)
                    
                      Call mysql_select(public_all2, "SELECT DISTINCT subject_code, Grade, Remark FROM tbl_grade WHERE ID = '" & lbl_id.Caption & "' AND SY = '" & sy_2 & "' AND Period='Final' ORDER BY subject_code ASC ")
                ctr = 1
                While Not public_all2.EOF
                    If public_all2.RecordCount = 0 Then
                         f(ctr) = "No grade"
                         f2(ctr) = "0"
                    Else
                          f(ctr) = public_all2.Fields("Remark").value
                         If f(ctr) <> "B" Then
                            f2(ctr) = "Promote to Grade III"
                        Else
                            f2(ctr) = "Unable to Promote"
                         End If
                    End If
                   
                    ctr = ctr + 1
                    public_all2.MoveNext
                Wend
                    ExcelSheet.Cells(48, 6).value = f(1)
                    ExcelSheet.Cells(49, 6).value = f(2)
                    ExcelSheet.Cells(50, 6).value = f(3)
                    ExcelSheet.Cells(51, 6).value = f(4)
                    ExcelSheet.Cells(52, 6).value = f(5)
                    ExcelSheet.Cells(53, 6).value = f(6)
                    ExcelSheet.Cells(54, 6).value = f(7)
                    ExcelSheet.Cells(55, 6).value = f(8)
                    ExcelSheet.Cells(56, 6).value = f(9)
                    ExcelSheet.Cells(57, 6).value = f(10)
                    
                    ExcelSheet.Cells(48, 7).value = f2(1)
                    ExcelSheet.Cells(49, 7).value = f2(2)
                    ExcelSheet.Cells(50, 7).value = f2(3)
                    ExcelSheet.Cells(51, 7).value = f2(4)
                    ExcelSheet.Cells(52, 7).value = f2(5)
                    ExcelSheet.Cells(53, 7).value = f2(6)
                    ExcelSheet.Cells(54, 7).value = f2(7)
                    ExcelSheet.Cells(55, 7).value = f2(8)
                    ExcelSheet.Cells(56, 7).value = f2(9)
                    ExcelSheet.Cells(57, 7).value = f2(10)
                    
                    
                    Call mysql_select(public_rs, "SELECT subject_code as Subject, period, grade, remark FROM tbl_grade WHERE ID = '" & lbl_id.Caption & "' AND SY = '" & sy_2 & "' AND Period='1st Grading'")
        no_subj = public_rs.RecordCount
        If no_subj = 0 Then
            
            average = 0
            remark = "No grades"
        Else
            average = 0
            While Not public_rs.EOF
                average = val(public_rs.Fields("grade")) + average
                public_rs.MoveNext
            Wend
            average = average / no_subj
            average = Round(average, 2)
            If average >= 90 Then
              remark = "A"
              ElseIf average >= 85 Then
              remark = "P"
              ElseIf average >= 80 Then
                  remark = "AP"
              ElseIf average >= 74 Then
                  remark = "D"
              Else
                  remark = "B"
            End If
            ExcelSheet.Cells(60, 2).value = remark
        End If
    
        Call mysql_select(public_rs, "SELECT subject_code as Subject, period, grade, remark FROM tbl_grade WHERE ID = '" & lbl_id.Caption & "' AND SY = '" & sy_2 & "' AND Period='2nd Grading'")
        no_subj = public_rs.RecordCount
        If no_subj = 0 Then
           
            average = 0
            remark = "No grades"
        Else
            average = 0
            While Not public_rs.EOF
                average = val(public_rs.Fields("grade")) + average
                public_rs.MoveNext
            Wend
            average = average / no_subj
            average = Round(average, 2)
            If average >= 90 Then
              remark = "A"
              ElseIf average >= 85 Then
              remark = "P"
              ElseIf average >= 80 Then
                  remark = "AP"
              ElseIf average >= 74 Then
                  remark = "D"
              Else
                  remark = "B"
            End If
            ExcelSheet.Cells(60, 3).value = remark
        End If
        Call mysql_select(public_rs, "SELECT subject_code as Subject, period, grade, remark FROM tbl_grade WHERE ID = '" & lbl_id.Caption & "' AND SY = '" & sy_2 & "' AND Period='3rd Grading'")
        no_subj = public_rs.RecordCount
        If no_subj = 0 Then
            
            average = 0
            remark = "No grades"
        Else
            average = 0
            While Not public_rs.EOF
                average = val(public_rs.Fields("grade")) + average
                public_rs.MoveNext
            Wend
            average = average / no_subj
            average = Round(average, 2)
            If average >= 90 Then
              remark = "A"
              ElseIf average >= 85 Then
              remark = "P"
              ElseIf average >= 80 Then
                  remark = "AP"
              ElseIf average >= 74 Then
                  remark = "D"
              Else
                  remark = "B"
            End If
            ExcelSheet.Cells(60, 4).value = remark
        End If
        Call mysql_select(public_rs, "SELECT subject_code as Subject, period, grade, remark FROM tbl_grade WHERE ID = '" & lbl_id.Caption & "' AND SY = '" & sy_2 & "' AND Period='4th Grading'")
        no_subj = public_rs.RecordCount
        If no_subj = 0 Then
            
            average = 0
            remark = "No grades"
        Else
            average = 0
            While Not public_rs.EOF
                average = val(public_rs.Fields("grade")) + average
                public_rs.MoveNext
            Wend
            average = average / no_subj
            average = Round(average, 2)
            If average >= 90 Then
              remark = "A"
              ElseIf average >= 85 Then
              remark = "P"
              ElseIf average >= 80 Then
                  remark = "AP"
              ElseIf average >= 74 Then
                  remark = "D"
              Else
                  remark = "B"
            End If
            ExcelSheet.Cells(60, 5).value = remark
        End If
        Call mysql_select(public_rs, "SELECT subject_code as Subject, period, grade, remark FROM tbl_grade WHERE ID = '" & lbl_id.Caption & "' AND SY = '" & sy_2 & "' AND Period='Final'")
        no_subj = public_rs.RecordCount
        If no_subj = 0 Then
            
            average = 0
            remark = "No grades"
        Else
            average = 0
            While Not public_rs.EOF
                average = val(public_rs.Fields("grade")) + average
                public_rs.MoveNext
            Wend
            average = average / no_subj
            average = Round(average, 2)
            If average >= 90 Then
              remark = "A"
              ElseIf average >= 85 Then
              remark = "P"
              ElseIf average >= 80 Then
                  remark = "AP"
              ElseIf average >= 74 Then
                  remark = "D"
              Else
                  remark = "B"
            End If
            ExcelSheet.Cells(60, 6).value = remark
            If remark <> "B" Then
                ExcelSheet.Cells(60, 7).value = "Promote to Grade III"
            Else
                ExcelSheet.Cells(60, 7).value = "Unable to Promote"
            End If
            
            If remark = "B" Then
                 ExcelSheet.Cells(62, 2).value = "Unable to promote"
            Else
                ExcelSheet.Cells(62, 2).value = "Grade III"
            End If
            
            
        End If
    Call next_3
    
    End Sub
    Public Sub next_3()
    Dim subject(10) As String
    Dim first(10) As String
    Dim sec(10) As String
    Dim third(10) As String
    Dim fourth(10) As String
    Dim f(10) As String
    Dim f2(10) As String
    Dim ctr As Integer
    Dim sy_3 As String
    
                    
        Call mysql_select(public_rs, "SELECT * FROM tbl_student_level WHERE ID = '" & lbl_id.Caption & "' AND lvl_name = 'Grade 3'")
    If public_rs.RecordCount = 0 Then
        ExcelSheet.Cells(46, 10).value = ""
        ExcelSheet.Cells(45, 10).value = ""
        sy_3 = ""
    Else
        ExcelSheet.Cells(46, 10).value = "School Year: " & public_rs.Fields("SY").value
        ExcelSheet.Cells(45, 10).value = "Manuel S. Rojas Elementary School"
        sy_3 = public_rs.Fields("SY").value
    End If
    
    ctr = 1
    Call mysql_select(public_all, "SELECT DISTINCT a.subject_code,  b.subject_name as Subject, a.Grade, a.Remark FROM tbl_grade a LEFT JOIN tbl_subject b ON a.subject_code=b.subject_code WHERE a.ID = '" & lbl_id.Caption & "' AND a.SY = '" & sy_3 & "' AND a.Period='1st Grading' ORDER BY subject_code ASC ")
                 While Not public_all.EOF
                    subject(ctr) = public_all.Fields("Subject").value
                    first(ctr) = public_all.Fields("Remark").value
                    ctr = ctr + 1
                    public_all.MoveNext
                Wend
                
                    ExcelSheet.Cells(48, 9).value = subject(1)
                    ExcelSheet.Cells(49, 9).value = subject(2)
                    ExcelSheet.Cells(50, 9).value = subject(3)
                    ExcelSheet.Cells(51, 9).value = subject(4)
                   ExcelSheet.Cells(52, 9).value = subject(5)
                    ExcelSheet.Cells(53, 9).value = subject(6)
                    ExcelSheet.Cells(54, 9).value = subject(7)
                   ExcelSheet.Cells(55, 9).value = subject(8)
                    ExcelSheet.Cells(56, 9).value = subject(9)
                    ExcelSheet.Cells(57, 9).value = subject(10)
                    
                    ExcelSheet.Cells(48, 10).value = first(1)
                    ExcelSheet.Cells(49, 10).value = first(2)
                    ExcelSheet.Cells(50, 10).value = first(3)
                    ExcelSheet.Cells(51, 10).value = first(4)
                    ExcelSheet.Cells(52, 10).value = first(5)
                    ExcelSheet.Cells(53, 10).value = first(6)
                    ExcelSheet.Cells(54, 10).value = first(7)
                    ExcelSheet.Cells(55, 10).value = first(8)
                    ExcelSheet.Cells(56, 10).value = first(9)
                    ExcelSheet.Cells(57, 10).value = first(10)
                   
               Call mysql_select(public_all2, "SELECT DISTINCT subject_code, Grade, Remark FROM tbl_grade WHERE ID = '" & lbl_id.Caption & "' AND SY = '" & sy_3 & "' AND Period='2nd Grading' ORDER BY subject_code ASC  ")
                 
                    ctr = 1
                 While Not public_all2.EOF
                    If public_all2.RecordCount = 0 Then
                         sec(ctr) = "No grade"
                    Else
                         sec(ctr) = public_all2.Fields("Remark").value
                    End If
                   
                    ctr = ctr + 1
                    public_all2.MoveNext
                Wend
               
                
                    ExcelSheet.Cells(48, 11).value = sec(1)
                    ExcelSheet.Cells(49, 11).value = sec(2)
                    ExcelSheet.Cells(50, 11).value = sec(3)
                    ExcelSheet.Cells(51, 11).value = sec(4)
                    ExcelSheet.Cells(52, 11).value = sec(5)
                    ExcelSheet.Cells(53, 11).value = sec(6)
                    ExcelSheet.Cells(54, 11).value = sec(7)
                    ExcelSheet.Cells(55, 11).value = sec(8)
                    ExcelSheet.Cells(56, 11).value = sec(9)
                   ExcelSheet.Cells(57, 11).value = sec(10)
                    
                    Call mysql_select(public_all2, "SELECT DISTINCT subject_code, Grade, Remark FROM tbl_grade WHERE ID = '" & lbl_id.Caption & "' AND SY = '" & sy_3 & "' AND Period='3rd Grading' ORDER BY subject_code ASC  ")
                ctr = 1
                 While Not public_all2.EOF
                    If public_all2.RecordCount = 0 Then
                         third(ctr) = "No grade"
                    Else
                         third(ctr) = public_all2.Fields("Remark").value
                    End If
                   
                    ctr = ctr + 1
                    public_all2.MoveNext
                Wend
                    ExcelSheet.Cells(48, 12).value = third(1)
                    ExcelSheet.Cells(49, 12).value = third(2)
                    ExcelSheet.Cells(50, 12).value = third(3)
                    ExcelSheet.Cells(51, 12).value = third(4)
                    ExcelSheet.Cells(52, 12).value = third(5)
                    ExcelSheet.Cells(53, 12).value = third(6)
                    ExcelSheet.Cells(54, 12).value = third(7)
                    ExcelSheet.Cells(55, 12).value = third(8)
                    ExcelSheet.Cells(56, 12).value = third(9)
                    ExcelSheet.Cells(57, 12).value = third(10)
                    
                    Call mysql_select(public_all2, "SELECT DISTINCT subject_code, Grade, Remark FROM tbl_grade WHERE ID = '" & lbl_id.Caption & "' AND SY = '" & sy_3 & "' AND Period='4th Grading' ORDER BY subject_code ASC  ")
                ctr = 1
                While Not public_all2.EOF
                    If public_all2.RecordCount = 0 Then
                         fourth(ctr) = "No grade"
                    Else
                         fourth(ctr) = public_all2.Fields("Remark").value
                    End If
                   
                    ctr = ctr + 1
                    public_all2.MoveNext
                Wend
                    ExcelSheet.Cells(48, 13).value = fourth(1)
                     ExcelSheet.Cells(49, 13).value = fourth(2)
                     ExcelSheet.Cells(50, 13).value = fourth(3)
                     ExcelSheet.Cells(51, 13).value = fourth(4)
                     ExcelSheet.Cells(52, 13).value = fourth(5)
                     ExcelSheet.Cells(53, 13).value = fourth(6)
                     ExcelSheet.Cells(54, 13).value = fourth(7)
                     ExcelSheet.Cells(55, 13).value = fourth(8)
                     ExcelSheet.Cells(56, 13).value = fourth(9)
                     ExcelSheet.Cells(57, 13).value = fourth(10)
                    
                      Call mysql_select(public_all2, "SELECT DISTINCT subject_code, Grade, Remark FROM tbl_grade WHERE ID = '" & lbl_id.Caption & "' AND SY = '" & sy_3 & "' AND Period='Final' ORDER BY subject_code ASC ")
                ctr = 1
                While Not public_all2.EOF
                    If public_all2.RecordCount = 0 Then
                         f(ctr) = "No grade"
                         f2(ctr) = "0"
                    Else
                          f(ctr) = public_all2.Fields("Remark").value
                         If f(ctr) <> "B" Then
                            f2(ctr) = "Promote to Grade IV"
                        Else
                            f2(ctr) = "Unable to Promote"
                         End If
                    End If
                   
                    ctr = ctr + 1
                    public_all2.MoveNext
                Wend
                    ExcelSheet.Cells(48, 14).value = f(1)
                    ExcelSheet.Cells(49, 14).value = f(2)
                    ExcelSheet.Cells(50, 14).value = f(3)
                    ExcelSheet.Cells(51, 14).value = f(4)
                    ExcelSheet.Cells(52, 14).value = f(5)
                    ExcelSheet.Cells(53, 14).value = f(6)
                    ExcelSheet.Cells(54, 14).value = f(7)
                    ExcelSheet.Cells(55, 14).value = f(8)
                    ExcelSheet.Cells(56, 14).value = f(9)
                    ExcelSheet.Cells(57, 14).value = f(10)
                    
                    ExcelSheet.Cells(48, 15).value = f2(1)
                    ExcelSheet.Cells(49, 15).value = f2(2)
                    ExcelSheet.Cells(50, 15).value = f2(3)
                    ExcelSheet.Cells(51, 15).value = f2(4)
                    ExcelSheet.Cells(52, 15).value = f2(5)
                    ExcelSheet.Cells(53, 15).value = f2(6)
                    ExcelSheet.Cells(54, 15).value = f2(7)
                    ExcelSheet.Cells(55, 15).value = f2(8)
                    ExcelSheet.Cells(56, 15).value = f2(9)
                    ExcelSheet.Cells(57, 15).value = f2(10)
                    
                     Call mysql_select(public_rs, "SELECT subject_code as Subject, period, grade, remark FROM tbl_grade WHERE ID = '" & lbl_id.Caption & "' AND SY = '" & sy_3 & "' AND Period='1st Grading'")
        no_subj = public_rs.RecordCount
        If no_subj = 0 Then
            
            average = 0
            remark = "No grades"
        Else
            average = 0
            While Not public_rs.EOF
                average = val(public_rs.Fields("grade")) + average
                public_rs.MoveNext
            Wend
            average = average / no_subj
            average = Round(average, 2)
            If average >= 90 Then
              remark = "A"
              ElseIf average >= 85 Then
              remark = "P"
              ElseIf average >= 80 Then
                  remark = "AP"
              ElseIf average >= 74 Then
                  remark = "D"
              Else
                  remark = "B"
            End If
            ExcelSheet.Cells(60, 10).value = remark
        End If
    
        Call mysql_select(public_rs, "SELECT subject_code as Subject, period, grade, remark FROM tbl_grade WHERE ID = '" & lbl_id.Caption & "' AND SY = '" & sy_3 & "' AND Period='2nd Grading'")
        no_subj = public_rs.RecordCount
        If no_subj = 0 Then
            
            average = 0
            remark = "No grades"
        Else
            average = 0
            While Not public_rs.EOF
                average = val(public_rs.Fields("grade")) + average
                public_rs.MoveNext
            Wend
            average = average / no_subj
            average = Round(average, 2)
            If average >= 90 Then
              remark = "A"
              ElseIf average >= 85 Then
              remark = "P"
              ElseIf average >= 80 Then
                  remark = "AP"
              ElseIf average >= 74 Then
                  remark = "D"
              Else
                  remark = "B"
            End If
            ExcelSheet.Cells(60, 11).value = remark
        End If
        Call mysql_select(public_rs, "SELECT subject_code as Subject, period, grade, remark FROM tbl_grade WHERE ID = '" & lbl_id.Caption & "' AND SY = '" & sy_3 & "' AND Period='3rd Grading'")
        no_subj = public_rs.RecordCount
        If no_subj = 0 Then
            
            average = 0
            remark = "No grades"
        Else
            average = 0
            While Not public_rs.EOF
                average = val(public_rs.Fields("grade")) + average
                public_rs.MoveNext
            Wend
            average = average / no_subj
            average = Round(average, 2)
            If average >= 90 Then
              remark = "A"
              ElseIf average >= 85 Then
              remark = "P"
              ElseIf average >= 80 Then
                  remark = "AP"
              ElseIf average >= 74 Then
                  remark = "D"
              Else
                  remark = "B"
            End If
            ExcelSheet.Cells(60, 12).value = remark
        End If
        Call mysql_select(public_rs, "SELECT subject_code as Subject, period, grade, remark FROM tbl_grade WHERE ID = '" & lbl_id.Caption & "' AND SY = '" & sy_3 & "' AND Period='4th Grading'")
        no_subj = public_rs.RecordCount
        If no_subj = 0 Then
            
            average = 0
            remark = "No grades"
        Else
            average = 0
            While Not public_rs.EOF
                average = val(public_rs.Fields("grade")) + average
                public_rs.MoveNext
            Wend
            average = average / no_subj
            average = Round(average, 2)
            If average >= 90 Then
              remark = "A"
              ElseIf average >= 85 Then
              remark = "P"
              ElseIf average >= 80 Then
                  remark = "AP"
              ElseIf average >= 74 Then
                  remark = "D"
              Else
                  remark = "B"
            End If
            ExcelSheet.Cells(60, 13).value = remark
        End If
        Call mysql_select(public_rs, "SELECT subject_code as Subject, period, grade, remark FROM tbl_grade WHERE ID = '" & lbl_id.Caption & "' AND SY = '" & sy_3 & "' AND Period='Final'")
        no_subj = public_rs.RecordCount
        If no_subj = 0 Then
            
            average = 0
            remark = "No grades"
        Else
            average = 0
            While Not public_rs.EOF
                average = val(public_rs.Fields("grade")) + average
                public_rs.MoveNext
            Wend
            average = average / no_subj
            average = Round(average, 2)
            If average >= 90 Then
              remark = "A"
              ElseIf average >= 85 Then
              remark = "P"
              ElseIf average >= 80 Then
                  remark = "AP"
              ElseIf average >= 74 Then
                  remark = "D"
              Else
                  remark = "B"
            End If
            ExcelSheet.Cells(60, 14).value = remark
           If remark <> "B" Then
                ExcelSheet.Cells(60, 15).value = "Promote to Grade IV"
            Else
                ExcelSheet.Cells(60, 15).value = "Unable to Promote"
            End If
            
            If remark = "B" Then
                 ExcelSheet.Cells(62, 10).value = "Unable to promote"
            Else
                ExcelSheet.Cells(62, 10).value = "Grade IV"
            End If
            
            
        End If
    
         Call next_4
         End Sub
         Public Sub next_4()
         
         
    Dim subject(10) As String
    Dim first(10) As String
    Dim sec(10) As String
    Dim third(10) As String
    Dim fourth(10) As String
    Dim f(10) As String
    Dim f2(10) As String
    Dim ctr As Integer
    Dim sy_4 As String
        Call mysql_select(public_rs2, "SELECT * FROM tbl_student_level WHERE ID = '" & lbl_id.Caption & "' AND lvl_name = 'Grade 4'")
    If public_rs2.RecordCount = 0 Then
        ExcelSheet.Cells(66, 2).value = ""
         ExcelSheet.Cells(65, 2).value = ""
        sy_4 = ""
    Else
        ExcelSheet.Cells(66, 2).value = "School Year: " & public_rs2.Fields("SY").value
         ExcelSheet.Cells(65, 2).value = "Manuel S. Rojas Elementary School"
        sy_4 = public_rs2.Fields("SY").value
    End If
    
    ctr = 1
    Call mysql_select(public_all, "SELECT DISTINCT a.subject_code,  b.subject_name as Subject, a.Grade, a.Remark FROM tbl_grade a LEFT JOIN tbl_subject b ON a.subject_code=b.subject_code WHERE a.ID = '" & lbl_id.Caption & "' AND a.SY = '" & sy_4 & "' AND a.Period='1st Grading' ORDER BY subject_code ASC ")
                 While Not public_all.EOF
                    subject(ctr) = public_all.Fields("Subject").value
                    first(ctr) = public_all.Fields("Remark").value
                    ctr = ctr + 1
                    public_all.MoveNext
                Wend
                
                    ExcelSheet.Cells(68, 1).value = subject(1)
                    ExcelSheet.Cells(69, 1).value = subject(2)
                    ExcelSheet.Cells(70, 1).value = subject(3)
                    ExcelSheet.Cells(71, 1).value = subject(4)
                   ExcelSheet.Cells(72, 1).value = subject(5)
                    ExcelSheet.Cells(73, 1).value = subject(6)
                    ExcelSheet.Cells(74, 1).value = subject(7)
                   ExcelSheet.Cells(75, 1).value = subject(8)
                    ExcelSheet.Cells(76, 1).value = subject(9)
                    ExcelSheet.Cells(77, 1).value = subject(10)
                    
                    ExcelSheet.Cells(68, 2).value = first(1)
                    ExcelSheet.Cells(69, 2).value = first(2)
                    ExcelSheet.Cells(70, 2).value = first(3)
                    ExcelSheet.Cells(71, 2).value = first(4)
                    ExcelSheet.Cells(72, 2).value = first(5)
                    ExcelSheet.Cells(73, 2).value = first(6)
                    ExcelSheet.Cells(74, 2).value = first(7)
                    ExcelSheet.Cells(75, 2).value = first(8)
                    ExcelSheet.Cells(76, 2).value = first(9)
                    ExcelSheet.Cells(77, 2).value = first(10)
                   
               Call mysql_select(public_all2, "SELECT DISTINCT subject_code, Grade, Remark FROM tbl_grade WHERE ID = '" & lbl_id.Caption & "' AND SY = '" & sy_4 & "' AND Period='2nd Grading' ORDER BY subject_code ASC  ")
                 
                    ctr = 1
                 While Not public_all2.EOF
                    If public_all2.RecordCount = 0 Then
                         sec(ctr) = "No grade"
                    Else
                         sec(ctr) = public_all2.Fields("Remark").value
                    End If
                   
                    ctr = ctr + 1
                    public_all2.MoveNext
                Wend
               
                
                    ExcelSheet.Cells(68, 3).value = sec(1)
                    ExcelSheet.Cells(69, 3).value = sec(2)
                    ExcelSheet.Cells(70, 3).value = sec(3)
                    ExcelSheet.Cells(71, 3).value = sec(4)
                    ExcelSheet.Cells(72, 3).value = sec(5)
                    ExcelSheet.Cells(73, 3).value = sec(6)
                    ExcelSheet.Cells(74, 3).value = sec(7)
                    ExcelSheet.Cells(75, 3).value = sec(8)
                    ExcelSheet.Cells(76, 3).value = sec(9)
                   ExcelSheet.Cells(77, 3).value = sec(10)
                    
                    Call mysql_select(public_all2, "SELECT DISTINCT subject_code, Grade, Remark FROM tbl_grade WHERE ID = '" & lbl_id.Caption & "' AND SY = '" & sy_4 & "' AND Period='3rd Grading' ORDER BY subject_code ASC  ")
                ctr = 1
                 While Not public_all2.EOF
                    If public_all2.RecordCount = 0 Then
                         third(ctr) = "No grade"
                    Else
                         third(ctr) = public_all2.Fields("Remark").value
                    End If
                   
                    ctr = ctr + 1
                    public_all2.MoveNext
                Wend
                    ExcelSheet.Cells(68, 4).value = third(1)
                    ExcelSheet.Cells(69, 4).value = third(2)
                    ExcelSheet.Cells(70, 4).value = third(3)
                    ExcelSheet.Cells(71, 4).value = third(4)
                    ExcelSheet.Cells(72, 4).value = third(5)
                    ExcelSheet.Cells(73, 4).value = third(6)
                    ExcelSheet.Cells(74, 4).value = third(7)
                    ExcelSheet.Cells(75, 4).value = third(8)
                    ExcelSheet.Cells(76, 4).value = third(9)
                    ExcelSheet.Cells(77, 4).value = third(10)
                    
                    Call mysql_select(public_all2, "SELECT DISTINCT subject_code, Grade, Remark FROM tbl_grade WHERE ID = '" & lbl_id.Caption & "' AND SY = '" & sy_4 & "' AND Period='4th Grading' ORDER BY subject_code ASC  ")
                ctr = 1
                While Not public_all2.EOF
                    If public_all2.RecordCount = 0 Then
                         fourth(ctr) = "No grade"
                    Else
                         fourth(ctr) = public_all2.Fields("Remark").value
                    End If
                   
                    ctr = ctr + 1
                    public_all2.MoveNext
                Wend
                    ExcelSheet.Cells(68, 5).value = fourth(1)
                     ExcelSheet.Cells(69, 5).value = fourth(2)
                     ExcelSheet.Cells(70, 5).value = fourth(3)
                     ExcelSheet.Cells(71, 5).value = fourth(4)
                     ExcelSheet.Cells(72, 5).value = fourth(5)
                     ExcelSheet.Cells(73, 5).value = fourth(6)
                     ExcelSheet.Cells(74, 5).value = fourth(7)
                     ExcelSheet.Cells(75, 5).value = fourth(8)
                     ExcelSheet.Cells(76, 5).value = fourth(9)
                     ExcelSheet.Cells(77, 5).value = fourth(10)
                    
                      Call mysql_select(public_all2, "SELECT DISTINCT subject_code, Grade, Remark FROM tbl_grade WHERE ID = '" & lbl_id.Caption & "' AND SY = '" & sy_4 & "' AND Period='Final' ORDER BY subject_code ASC ")
                ctr = 1
                While Not public_all2.EOF
                    If public_all2.RecordCount = 0 Then
                         f(ctr) = "No grade"
                         f2(ctr) = "0"
                    Else
                         f(ctr) = public_all2.Fields("Remark").value
                         If f(ctr) <> "B" Then
                            f2(ctr) = "Promote to Grade V"
                        Else
                            f2(ctr) = "Unable to Promote"
                         End If
                    End If
                   
                    ctr = ctr + 1
                    public_all2.MoveNext
                Wend
                    ExcelSheet.Cells(68, 6).value = f(1)
                    ExcelSheet.Cells(69, 6).value = f(2)
                    ExcelSheet.Cells(70, 6).value = f(3)
                    ExcelSheet.Cells(71, 6).value = f(4)
                    ExcelSheet.Cells(72, 6).value = f(5)
                    ExcelSheet.Cells(73, 6).value = f(6)
                    ExcelSheet.Cells(74, 6).value = f(7)
                    ExcelSheet.Cells(75, 6).value = f(8)
                    ExcelSheet.Cells(76, 6).value = f(9)
                    ExcelSheet.Cells(77, 6).value = f(10)
                    
                    ExcelSheet.Cells(68, 7).value = f2(1)
                    ExcelSheet.Cells(69, 7).value = f2(2)
                    ExcelSheet.Cells(70, 7).value = f2(3)
                    ExcelSheet.Cells(71, 7).value = f2(4)
                    ExcelSheet.Cells(72, 7).value = f2(5)
                    ExcelSheet.Cells(73, 7).value = f2(6)
                    ExcelSheet.Cells(74, 7).value = f2(7)
                    ExcelSheet.Cells(75, 7).value = f2(8)
                    ExcelSheet.Cells(76, 7).value = f2(9)
                    ExcelSheet.Cells(77, 7).value = f2(10)
                    
                    Call mysql_select(public_rs, "SELECT subject_code as Subject, period, grade, remark FROM tbl_grade WHERE ID = '" & lbl_id.Caption & "' AND SY = '" & sy_4 & "' AND Period='1st Grading'")
        no_subj = public_rs.RecordCount
        If no_subj = 0 Then
            
            average = 0
            remark = "No grades"
        Else
            average = 0
            While Not public_rs.EOF
                average = val(public_rs.Fields("grade")) + average
                public_rs.MoveNext
            Wend
            average = average / no_subj
            average = Round(average, 2)
            If average >= 90 Then
              remark = "A"
              ElseIf average >= 85 Then
              remark = "P"
              ElseIf average >= 80 Then
                  remark = "AP"
              ElseIf average >= 74 Then
                  remark = "D"
              Else
                  remark = "B"
            End If
            ExcelSheet.Cells(80, 2).value = remark
        End If
    
        Call mysql_select(public_rs, "SELECT subject_code as Subject, period, grade, remark FROM tbl_grade WHERE ID = '" & lbl_id.Caption & "' AND SY = '" & sy_4 & "' AND Period='2nd Grading'")
        no_subj = public_rs.RecordCount
        If no_subj = 0 Then
           
            average = 0
            remark = "No grades"
        Else
            average = 0
            While Not public_rs.EOF
                average = val(public_rs.Fields("grade")) + average
                public_rs.MoveNext
            Wend
            average = average / no_subj
            average = Round(average, 2)
            If average >= 90 Then
              remark = "A"
              ElseIf average >= 85 Then
              remark = "P"
              ElseIf average >= 80 Then
                  remark = "AP"
              ElseIf average >= 74 Then
                  remark = "D"
              Else
                  remark = "B"
            End If
            ExcelSheet.Cells(80, 3).value = remark
        End If
        Call mysql_select(public_rs, "SELECT subject_code as Subject, period, grade, remark FROM tbl_grade WHERE ID = '" & lbl_id.Caption & "' AND SY = '" & sy_4 & "' AND Period='3rd Grading'")
        no_subj = public_rs.RecordCount
        If no_subj = 0 Then
            
            average = 0
            remark = "No grades"
        Else
            average = 0
            While Not public_rs.EOF
                average = val(public_rs.Fields("grade")) + average
                public_rs.MoveNext
            Wend
            average = average / no_subj
            average = Round(average, 2)
            If average >= 90 Then
              remark = "A"
              ElseIf average >= 85 Then
              remark = "P"
              ElseIf average >= 80 Then
                  remark = "AP"
              ElseIf average >= 74 Then
                  remark = "D"
              Else
                  remark = "B"
            End If
            ExcelSheet.Cells(80, 4).value = remark
        End If
        Call mysql_select(public_rs, "SELECT subject_code as Subject, period, grade, remark FROM tbl_grade WHERE ID = '" & lbl_id.Caption & "' AND SY = '" & sy_4 & "' AND Period='4th Grading'")
        no_subj = public_rs.RecordCount
        If no_subj = 0 Then
            
            average = 0
            remark = "No grades"
        Else
            average = 0
            While Not public_rs.EOF
                average = val(public_rs.Fields("grade")) + average
                public_rs.MoveNext
            Wend
            average = average / no_subj
            average = Round(average, 2)
            If average >= 90 Then
              remark = "A"
              ElseIf average >= 85 Then
              remark = "P"
              ElseIf average >= 80 Then
                  remark = "AP"
              ElseIf average >= 74 Then
                  remark = "D"
              Else
                  remark = "B"
            End If
            ExcelSheet.Cells(80, 5).value = remark
        End If
        Call mysql_select(public_rs, "SELECT subject_code as Subject, period, grade, remark FROM tbl_grade WHERE ID = '" & lbl_id.Caption & "' AND SY = '" & sy_4 & "' AND Period='Final'")
        no_subj = public_rs.RecordCount
        If no_subj = 0 Then
            
            average = 0
            remark = "No grades"
        Else
            average = 0
            While Not public_rs.EOF
                average = val(public_rs.Fields("grade")) + average
                public_rs.MoveNext
            Wend
            average = average / no_subj
            average = Round(average, 2)
            If average >= 90 Then
              remark = "A"
              ElseIf average >= 85 Then
              remark = "P"
              ElseIf average >= 80 Then
                  remark = "AP"
              ElseIf average >= 74 Then
                  remark = "D"
              Else
                  remark = "B"
            End If
            ExcelSheet.Cells(80, 6).value = remark
            If remark <> "B" Then
                ExcelSheet.Cells(80, 7).value = "Promote to Grade V"
            Else
                ExcelSheet.Cells(80, 7).value = "Unable to Promote"
            End If
            
            If remark = "B" Then
                 ExcelSheet.Cells(82, 2).value = "Unable to promote"
            Else
                ExcelSheet.Cells(82, 2).value = "Grade V"
            End If
            
            
        End If
    
    Call next_5
    End Sub
    Public Sub next_5()
           
    Dim subject(10) As String
    Dim first(10) As String
    Dim sec(10) As String
    Dim third(10) As String
    Dim fourth(10) As String
    Dim f(10) As String
    Dim f2(10) As String
    Dim ctr As Integer
    Dim sy_5 As String
                    
         Call mysql_select(public_rs, "SELECT * FROM tbl_student_level WHERE ID = '" & lbl_id.Caption & "' AND lvl_name = 'Grade 5'")
    If public_rs.RecordCount = 0 Then
        ExcelSheet.Cells(66, 10).value = ""
         ExcelSheet.Cells(65, 10).value = ""
        sy_5 = ""
    Else
        ExcelSheet.Cells(66, 10).value = "School Year: " & public_rs.Fields("SY").value
         ExcelSheet.Cells(65, 10).value = "Manuel S. Rojas Elementary School"
        sy_5 = public_rs.Fields("SY").value
    End If
    
    ctr = 1
    Call mysql_select(public_all, "SELECT DISTINCT a.subject_code,  b.subject_name as Subject, a.Grade, a.Remark FROM tbl_grade a LEFT JOIN tbl_subject b ON a.subject_code=b.subject_code WHERE a.ID = '" & lbl_id.Caption & "' AND a.SY = '" & sy_5 & "' AND a.Period='1st Grading' ORDER BY subject_code ASC ")
                 While Not public_all.EOF
                    subject(ctr) = public_all.Fields("Subject").value
                    first(ctr) = public_all.Fields("Remark").value
                    ctr = ctr + 1
                    public_all.MoveNext
                Wend
                
                    ExcelSheet.Cells(68, 9).value = subject(1)
                    ExcelSheet.Cells(69, 9).value = subject(2)
                    ExcelSheet.Cells(70, 9).value = subject(3)
                    ExcelSheet.Cells(71, 9).value = subject(4)
                   ExcelSheet.Cells(72, 9).value = subject(5)
                    ExcelSheet.Cells(73, 9).value = subject(6)
                    ExcelSheet.Cells(74, 9).value = subject(7)
                   ExcelSheet.Cells(75, 9).value = subject(8)
                    ExcelSheet.Cells(76, 9).value = subject(9)
                    ExcelSheet.Cells(77, 9).value = subject(10)
                    
                    ExcelSheet.Cells(68, 10).value = first(1)
                    ExcelSheet.Cells(69, 10).value = first(2)
                    ExcelSheet.Cells(70, 10).value = first(3)
                    ExcelSheet.Cells(71, 10).value = first(4)
                    ExcelSheet.Cells(72, 10).value = first(5)
                    ExcelSheet.Cells(73, 10).value = first(6)
                    ExcelSheet.Cells(74, 10).value = first(7)
                    ExcelSheet.Cells(75, 10).value = first(8)
                    ExcelSheet.Cells(76, 10).value = first(9)
                    ExcelSheet.Cells(77, 10).value = first(10)
                   
               Call mysql_select(public_all2, "SELECT DISTINCT subject_code, Grade, Remark FROM tbl_grade WHERE ID = '" & lbl_id.Caption & "' AND SY = '" & sy_5 & "' AND Period='2nd Grading' ORDER BY subject_code ASC  ")
                 
                    ctr = 1
                 While Not public_all2.EOF
                    If public_all2.RecordCount = 0 Then
                         sec(ctr) = "No grade"
                    Else
                         sec(ctr) = public_all2.Fields("Remark").value
                    End If
                   
                    ctr = ctr + 1
                    public_all2.MoveNext
                Wend
               
                
                    ExcelSheet.Cells(68, 11).value = sec(1)
                    ExcelSheet.Cells(69, 11).value = sec(2)
                    ExcelSheet.Cells(70, 11).value = sec(3)
                    ExcelSheet.Cells(71, 11).value = sec(4)
                    ExcelSheet.Cells(72, 11).value = sec(5)
                    ExcelSheet.Cells(73, 11).value = sec(6)
                    ExcelSheet.Cells(74, 11).value = sec(7)
                    ExcelSheet.Cells(75, 11).value = sec(8)
                    ExcelSheet.Cells(76, 11).value = sec(9)
                   ExcelSheet.Cells(77, 11).value = sec(10)
                    
                    Call mysql_select(public_all2, "SELECT DISTINCT subject_code, Grade, Remark FROM tbl_grade WHERE ID = '" & lbl_id.Caption & "' AND SY = '" & sy_5 & "' AND Period='3rd Grading' ORDER BY subject_code ASC  ")
                ctr = 1
                 While Not public_all2.EOF
                    If public_all2.RecordCount = 0 Then
                         third(ctr) = "No grade"
                    Else
                         third(ctr) = public_all2.Fields("Remark").value
                    End If
                   
                    ctr = ctr + 1
                    public_all2.MoveNext
                Wend
                    ExcelSheet.Cells(68, 12).value = third(1)
                    ExcelSheet.Cells(69, 12).value = third(2)
                    ExcelSheet.Cells(70, 12).value = third(3)
                    ExcelSheet.Cells(71, 12).value = third(4)
                    ExcelSheet.Cells(72, 12).value = third(5)
                    ExcelSheet.Cells(73, 12).value = third(6)
                    ExcelSheet.Cells(74, 12).value = third(7)
                    ExcelSheet.Cells(75, 12).value = third(8)
                    ExcelSheet.Cells(76, 12).value = third(9)
                    ExcelSheet.Cells(77, 12).value = third(10)
                    
                    Call mysql_select(public_all2, "SELECT DISTINCT subject_code, Grade, Remark FROM tbl_grade WHERE ID = '" & lbl_id.Caption & "' AND SY = '" & sy_5 & "' AND Period='4th Grading' ORDER BY subject_code ASC  ")
                ctr = 1
                While Not public_all2.EOF
                    If public_all2.RecordCount = 0 Then
                         fourth(ctr) = "No grade"
                    Else
                         fourth(ctr) = public_all2.Fields("Remark").value
                    End If
                   
                    ctr = ctr + 1
                    public_all2.MoveNext
                Wend
                    ExcelSheet.Cells(68, 13).value = fourth(1)
                     ExcelSheet.Cells(69, 13).value = fourth(2)
                     ExcelSheet.Cells(70, 13).value = fourth(3)
                     ExcelSheet.Cells(71, 13).value = fourth(4)
                     ExcelSheet.Cells(72, 13).value = fourth(5)
                     ExcelSheet.Cells(73, 13).value = fourth(6)
                     ExcelSheet.Cells(74, 13).value = fourth(7)
                     ExcelSheet.Cells(75, 13).value = fourth(8)
                     ExcelSheet.Cells(76, 13).value = fourth(9)
                     ExcelSheet.Cells(77, 13).value = fourth(10)
                    
                      Call mysql_select(public_all2, "SELECT DISTINCT subject_code, Grade, Remark FROM tbl_grade WHERE ID = '" & lbl_id.Caption & "' AND SY = '" & sy_5 & "' AND Period='Final' ORDER BY subject_code ASC ")
                ctr = 1
                While Not public_all2.EOF
                    If public_all2.RecordCount = 0 Then
                         f(ctr) = "No grade"
                         f2(ctr) = "0"
                    Else
                         f(ctr) = public_all2.Fields("Remark").value
                         If f(ctr) <> "B" Then
                            f2(ctr) = "Promote to Grade VI"
                        Else
                            f2(ctr) = "Unable to Promote"
                         End If
                    End If
                   
                    ctr = ctr + 1
                    public_all2.MoveNext
                Wend
                    ExcelSheet.Cells(68, 14).value = f(1)
                    ExcelSheet.Cells(69, 14).value = f(2)
                    ExcelSheet.Cells(70, 14).value = f(3)
                    ExcelSheet.Cells(71, 14).value = f(4)
                    ExcelSheet.Cells(72, 14).value = f(5)
                    ExcelSheet.Cells(73, 14).value = f(6)
                    ExcelSheet.Cells(74, 14).value = f(7)
                    ExcelSheet.Cells(75, 14).value = f(8)
                    ExcelSheet.Cells(76, 14).value = f(9)
                    ExcelSheet.Cells(77, 14).value = f(10)
                    
                    ExcelSheet.Cells(68, 15).value = f2(1)
                    ExcelSheet.Cells(69, 15).value = f2(2)
                    ExcelSheet.Cells(70, 15).value = f2(3)
                    ExcelSheet.Cells(71, 15).value = f2(4)
                    ExcelSheet.Cells(72, 15).value = f2(5)
                    ExcelSheet.Cells(73, 15).value = f2(6)
                    ExcelSheet.Cells(74, 15).value = f2(7)
                    ExcelSheet.Cells(75, 15).value = f2(8)
                    ExcelSheet.Cells(76, 15).value = f2(9)
                    ExcelSheet.Cells(77, 15).value = f2(10)
                    
                     Call mysql_select(public_rs, "SELECT subject_code as Subject, period, grade, remark FROM tbl_grade WHERE ID = '" & lbl_id.Caption & "' AND SY = '" & sy_5 & "' AND Period='1st Grading'")
        no_subj = public_rs.RecordCount
        If no_subj = 0 Then
            
            average = 0
            remark = "No grades"
        Else
            average = 0
            While Not public_rs.EOF
                average = val(public_rs.Fields("grade")) + average
                public_rs.MoveNext
            Wend
            average = average / no_subj
            average = Round(average, 2)
            If average >= 90 Then
              remark = "A"
              ElseIf average >= 85 Then
              remark = "P"
              ElseIf average >= 80 Then
                  remark = "AP"
              ElseIf average >= 74 Then
                  remark = "D"
              Else
                  remark = "B"
            End If
            ExcelSheet.Cells(80, 10).value = remark
        End If
    
        Call mysql_select(public_rs, "SELECT subject_code as Subject, period, grade, remark FROM tbl_grade WHERE ID = '" & lbl_id.Caption & "' AND SY = '" & sy_5 & "' AND Period='2nd Grading'")
        no_subj = public_rs.RecordCount
        If no_subj = 0 Then
            
            average = 0
            remark = "No grades"
        Else
            average = 0
            While Not public_rs.EOF
                average = val(public_rs.Fields("grade")) + average
                public_rs.MoveNext
            Wend
            average = average / no_subj
            average = Round(average, 2)
            If average >= 90 Then
              remark = "A"
              ElseIf average >= 85 Then
              remark = "P"
              ElseIf average >= 80 Then
                  remark = "AP"
              ElseIf average >= 74 Then
                  remark = "D"
              Else
                  remark = "B"
            End If
            ExcelSheet.Cells(80, 11).value = remark
        End If
        Call mysql_select(public_rs, "SELECT subject_code as Subject, period, grade, remark FROM tbl_grade WHERE ID = '" & lbl_id.Caption & "' AND SY = '" & sy_5 & "' AND Period='3rd Grading'")
        no_subj = public_rs.RecordCount
        If no_subj = 0 Then
            
            average = 0
            remark = "No grades"
        Else
            average = 0
            While Not public_rs.EOF
                average = val(public_rs.Fields("grade")) + average
                public_rs.MoveNext
            Wend
            average = average / no_subj
            average = Round(average, 2)
            If average >= 90 Then
              remark = "A"
              ElseIf average >= 85 Then
              remark = "P"
              ElseIf average >= 80 Then
                  remark = "AP"
              ElseIf average >= 74 Then
                  remark = "D"
              Else
                  remark = "B"
            End If
            ExcelSheet.Cells(80, 12).value = remark
        End If
        Call mysql_select(public_rs, "SELECT subject_code as Subject, period, grade, remark FROM tbl_grade WHERE ID = '" & lbl_id.Caption & "' AND SY = '" & sy_5 & "' AND Period='4th Grading'")
        no_subj = public_rs.RecordCount
        If no_subj = 0 Then
            
            average = 0
            remark = "No grades"
        Else
            average = 0
            While Not public_rs.EOF
                average = val(public_rs.Fields("grade")) + average
                public_rs.MoveNext
            Wend
            average = average / no_subj
            average = Round(average, 2)
            If average >= 90 Then
              remark = "A"
              ElseIf average >= 85 Then
              remark = "P"
              ElseIf average >= 80 Then
                  remark = "AP"
              ElseIf average >= 74 Then
                  remark = "D"
              Else
                  remark = "B"
            End If
            ExcelSheet.Cells(80, 13).value = remark
        End If
        Call mysql_select(public_rs, "SELECT subject_code as Subject, period, grade, remark FROM tbl_grade WHERE ID = '" & lbl_id.Caption & "' AND SY = '" & sy_5 & "' AND Period='Final'")
        no_subj = public_rs.RecordCount
        If no_subj = 0 Then
            
            average = 0
            remark = "No grades"
        Else
            average = 0
            While Not public_rs.EOF
                average = val(public_rs.Fields("grade")) + average
                public_rs.MoveNext
            Wend
            average = average / no_subj
            average = Round(average, 2)
            If average >= 90 Then
              remark = "A"
              ElseIf average >= 85 Then
              remark = "P"
              ElseIf average >= 80 Then
                  remark = "AP"
              ElseIf average >= 74 Then
                  remark = "D"
              Else
                  remark = "B"
            End If
            ExcelSheet.Cells(80, 14).value = remark
            If remark <> "B" Then
                ExcelSheet.Cells(80, 15).value = "Promote to Grade VI"
            Else
                ExcelSheet.Cells(80, 15).value = "Unable to Promote"
            End If
            
            If remark = "B" Then
                 ExcelSheet.Cells(82, 10).value = "Unable to promote"
            Else
                ExcelSheet.Cells(82, 10).value = "Grade VI"
            End If
            
            
        End If
    Call next_6
    End Sub
    Public Sub next_6()
            
    Dim subject(10) As String
    Dim first(10) As String
    Dim sec(10) As String
    Dim third(10) As String
    Dim fourth(10) As String
    Dim f(10) As String
    Dim f2(10) As String
    Dim ctr As Integer
    Dim sy_6 As String
                    
                    
         Call mysql_select(public_rs2, "SELECT * FROM tbl_student_level WHERE ID = '" & lbl_id.Caption & "' AND lvl_name = 'Grade 6'")
    If public_rs2.RecordCount = 0 Then
        ExcelSheet.Cells(86, 2).value = ""
         ExcelSheet.Cells(85, 2).value = ""
        sy_6 = ""
    Else
        ExcelSheet.Cells(86, 2).value = "School Year: " & public_rs2.Fields("SY").value
         ExcelSheet.Cells(85, 2).value = "Manuel S. Rojas Elementary School"
        sy_6 = public_rs2.Fields("SY").value
    End If
    
    ctr = 1
    Call mysql_select(public_all, "SELECT DISTINCT a.subject_code,  b.subject_name as Subject, a.Grade, a.Remark FROM tbl_grade a LEFT JOIN tbl_subject b ON a.subject_code=b.subject_code WHERE a.ID = '" & lbl_id.Caption & "' AND a.SY = '" & sy_6 & "' AND a.Period='1st Grading' ORDER BY subject_code ASC ")
                 While Not public_all.EOF
                    subject(ctr) = public_all.Fields("Subject").value
                    first(ctr) = public_all.Fields("Remark").value
                    ctr = ctr + 1
                    public_all.MoveNext
                Wend
                
                    ExcelSheet.Cells(88, 1).value = subject(1)
                    ExcelSheet.Cells(89, 1).value = subject(2)
                    ExcelSheet.Cells(90, 1).value = subject(3)
                    ExcelSheet.Cells(91, 1).value = subject(4)
                   ExcelSheet.Cells(92, 1).value = subject(5)
                    ExcelSheet.Cells(93, 1).value = subject(6)
                    ExcelSheet.Cells(94, 1).value = subject(7)
                   ExcelSheet.Cells(95, 1).value = subject(8)
                    ExcelSheet.Cells(96, 1).value = subject(9)
                    ExcelSheet.Cells(97, 1).value = subject(10)
                    
                    ExcelSheet.Cells(88, 2).value = first(1)
                    ExcelSheet.Cells(89, 2).value = first(2)
                    ExcelSheet.Cells(90, 2).value = first(3)
                    ExcelSheet.Cells(91, 2).value = first(4)
                    ExcelSheet.Cells(92, 2).value = first(5)
                    ExcelSheet.Cells(93, 2).value = first(6)
                    ExcelSheet.Cells(94, 2).value = first(7)
                    ExcelSheet.Cells(95, 2).value = first(8)
                    ExcelSheet.Cells(96, 2).value = first(9)
                    ExcelSheet.Cells(97, 2).value = first(10)
                   
               Call mysql_select(public_all2, "SELECT DISTINCT subject_code, Grade, Remark FROM tbl_grade WHERE ID = '" & lbl_id.Caption & "' AND SY = '" & sy_6 & "' AND Period='2nd Grading' ORDER BY subject_code ASC  ")
                 
                    ctr = 1
                 While Not public_all2.EOF
                    If public_all2.RecordCount = 0 Then
                         sec(ctr) = "No grade"
                    Else
                         sec(ctr) = public_all2.Fields("Remark").value
                    End If
                   
                    ctr = ctr + 1
                    public_all2.MoveNext
                Wend
               
                
                    ExcelSheet.Cells(88, 3).value = sec(1)
                    ExcelSheet.Cells(89, 3).value = sec(2)
                    ExcelSheet.Cells(90, 3).value = sec(3)
                    ExcelSheet.Cells(91, 3).value = sec(4)
                    ExcelSheet.Cells(92, 3).value = sec(5)
                    ExcelSheet.Cells(93, 3).value = sec(6)
                    ExcelSheet.Cells(94, 3).value = sec(7)
                    ExcelSheet.Cells(95, 3).value = sec(8)
                    ExcelSheet.Cells(96, 3).value = sec(9)
                   ExcelSheet.Cells(97, 3).value = sec(10)
                    
                    Call mysql_select(public_all2, "SELECT DISTINCT subject_code, Grade, Remark FROM tbl_grade WHERE ID = '" & lbl_id.Caption & "' AND SY = '" & sy_6 & "' AND Period='3rd Grading' ORDER BY subject_code ASC  ")
                ctr = 1
                 While Not public_all2.EOF
                    If public_all2.RecordCount = 0 Then
                         third(ctr) = "No grade"
                    Else
                         third(ctr) = public_all2.Fields("Remark").value
                    End If
                   
                    ctr = ctr + 1
                    public_all2.MoveNext
                Wend
                    ExcelSheet.Cells(88, 4).value = third(1)
                    ExcelSheet.Cells(89, 4).value = third(2)
                    ExcelSheet.Cells(90, 4).value = third(3)
                    ExcelSheet.Cells(91, 4).value = third(4)
                    ExcelSheet.Cells(92, 4).value = third(5)
                    ExcelSheet.Cells(93, 4).value = third(6)
                    ExcelSheet.Cells(94, 4).value = third(7)
                    ExcelSheet.Cells(95, 4).value = third(8)
                    ExcelSheet.Cells(96, 4).value = third(9)
                    ExcelSheet.Cells(97, 4).value = third(10)
                    
                    Call mysql_select(public_all2, "SELECT DISTINCT subject_code, Grade, Remark FROM tbl_grade WHERE ID = '" & lbl_id.Caption & "' AND SY = '" & sy_6 & "' AND Period='4th Grading' ORDER BY subject_code ASC  ")
                ctr = 1
                While Not public_all2.EOF
                    If public_all2.RecordCount = 0 Then
                         fourth(ctr) = "No grade"
                    Else
                         fourth(ctr) = public_all2.Fields("Remark").value
                    End If
                   
                    ctr = ctr + 1
                    public_all2.MoveNext
                Wend
                    ExcelSheet.Cells(88, 5).value = fourth(1)
                     ExcelSheet.Cells(89, 5).value = fourth(2)
                     ExcelSheet.Cells(90, 5).value = fourth(3)
                     ExcelSheet.Cells(91, 5).value = fourth(4)
                     ExcelSheet.Cells(92, 5).value = fourth(5)
                     ExcelSheet.Cells(93, 5).value = fourth(6)
                     ExcelSheet.Cells(94, 5).value = fourth(7)
                     ExcelSheet.Cells(95, 5).value = fourth(8)
                     ExcelSheet.Cells(96, 5).value = fourth(9)
                     ExcelSheet.Cells(97, 5).value = fourth(10)
                    
                      Call mysql_select(public_all2, "SELECT DISTINCT subject_code, Grade, Remark FROM tbl_grade WHERE ID = '" & lbl_id.Caption & "' AND SY = '" & sy_6 & "' AND Period='Final' ORDER BY subject_code ASC ")
                ctr = 1
                While Not public_all2.EOF
                    If public_all2.RecordCount = 0 Then
                         f(ctr) = "No grade"
                         f2(ctr) = "0"
                    Else
                          f(ctr) = public_all2.Fields("Remark").value
                         If f(ctr) <> "B" Then
                            f2(ctr) = "Promote to 1st Year"
                        Else
                            f2(ctr) = "Unable to Promote"
                         End If
                    End If
                   
                    ctr = ctr + 1
                    public_all2.MoveNext
                Wend
                    ExcelSheet.Cells(88, 6).value = f(1)
                    ExcelSheet.Cells(89, 6).value = f(2)
                    ExcelSheet.Cells(90, 6).value = f(3)
                    ExcelSheet.Cells(91, 6).value = f(4)
                    ExcelSheet.Cells(92, 6).value = f(5)
                    ExcelSheet.Cells(93, 6).value = f(6)
                    ExcelSheet.Cells(94, 6).value = f(7)
                    ExcelSheet.Cells(95, 6).value = f(8)
                    ExcelSheet.Cells(96, 6).value = f(9)
                    ExcelSheet.Cells(97, 6).value = f(10)
                    
                    ExcelSheet.Cells(88, 7).value = f2(1)
                    ExcelSheet.Cells(89, 7).value = f2(2)
                    ExcelSheet.Cells(90, 7).value = f2(3)
                    ExcelSheet.Cells(91, 7).value = f2(4)
                    ExcelSheet.Cells(92, 7).value = f2(5)
                    ExcelSheet.Cells(93, 7).value = f2(6)
                    ExcelSheet.Cells(94, 7).value = f2(7)
                    ExcelSheet.Cells(95, 7).value = f2(8)
                    ExcelSheet.Cells(96, 7).value = f2(9)
                    ExcelSheet.Cells(97, 7).value = f2(10)
                    
                    Call mysql_select(public_rs, "SELECT subject_code as Subject, period, grade, remark FROM tbl_grade WHERE ID = '" & lbl_id.Caption & "' AND SY = '" & sy_6 & "' AND Period='1st Grading'")
        no_subj = public_rs.RecordCount
        If no_subj = 0 Then
            
            average = 0
            remark = "No grades"
        Else
            average = 0
            While Not public_rs.EOF
                average = val(public_rs.Fields("grade")) + average
                public_rs.MoveNext
            Wend
            average = average / no_subj
            average = Round(average, 2)
            If average >= 90 Then
              remark = "A"
              ElseIf average >= 85 Then
              remark = "P"
              ElseIf average >= 80 Then
                  remark = "AP"
              ElseIf average >= 74 Then
                  remark = "D"
              Else
                  remark = "B"
            End If
            ExcelSheet.Cells(100, 2).value = remark
        End If
    
        Call mysql_select(public_rs, "SELECT subject_code as Subject, period, grade, remark FROM tbl_grade WHERE ID = '" & lbl_id.Caption & "' AND SY = '" & sy_6 & "' AND Period='2nd Grading'")
        no_subj = public_rs.RecordCount
        If no_subj = 0 Then
           
            average = 0
            remark = "No grades"
        Else
            average = 0
            While Not public_rs.EOF
                average = val(public_rs.Fields("grade")) + average
                public_rs.MoveNext
            Wend
            average = average / no_subj
            average = Round(average, 2)
            If average >= 90 Then
              remark = "A"
              ElseIf average >= 85 Then
              remark = "P"
              ElseIf average >= 80 Then
                  remark = "AP"
              ElseIf average >= 74 Then
                  remark = "D"
              Else
                  remark = "B"
            End If
            ExcelSheet.Cells(100, 3).value = remark
        End If
        Call mysql_select(public_rs, "SELECT subject_code as Subject, period, grade, remark FROM tbl_grade WHERE ID = '" & lbl_id.Caption & "' AND SY = '" & sy_6 & "' AND Period='3rd Grading'")
        no_subj = public_rs.RecordCount
        If no_subj = 0 Then
            
            average = 0
            remark = "No grades"
        Else
            average = 0
            While Not public_rs.EOF
                average = val(public_rs.Fields("grade")) + average
                public_rs.MoveNext
            Wend
            average = average / no_subj
            average = Round(average, 2)
            If average >= 90 Then
              remark = "A"
              ElseIf average >= 85 Then
              remark = "P"
              ElseIf average >= 80 Then
                  remark = "AP"
              ElseIf average >= 74 Then
                  remark = "D"
              Else
                  remark = "B"
            End If
            ExcelSheet.Cells(100, 4).value = remark
        End If
        Call mysql_select(public_rs, "SELECT subject_code as Subject, period, grade, remark FROM tbl_grade WHERE ID = '" & lbl_id.Caption & "' AND SY = '" & sy_6 & "' AND Period='4th Grading'")
        no_subj = public_rs.RecordCount
        If no_subj = 0 Then
            
            average = 0
            remark = "No grades"
        Else
            average = 0
            While Not public_rs.EOF
                average = val(public_rs.Fields("grade")) + average
                public_rs.MoveNext
            Wend
            average = average / no_subj
            average = Round(average, 2)
            If average >= 90 Then
              remark = "A"
              ElseIf average >= 85 Then
              remark = "P"
              ElseIf average >= 80 Then
                  remark = "AP"
              ElseIf average >= 74 Then
                  remark = "D"
              Else
                  remark = "B"
            End If
            ExcelSheet.Cells(100, 5).value = remark
        End If
        Call mysql_select(public_rs, "SELECT subject_code as Subject, period, grade, remark FROM tbl_grade WHERE ID = '" & lbl_id.Caption & "' AND SY = '" & sy_6 & "' AND Period='Final'")
        no_subj = public_rs.RecordCount
        If no_subj = 0 Then
            
            average = 0
            remark = "No grades"
        Else
            average = 0
            While Not public_rs.EOF
                average = val(public_rs.Fields("grade")) + average
                public_rs.MoveNext
            Wend
            average = average / no_subj
            average = Round(average, 2)
            If average >= 90 Then
              remark = "A"
              ElseIf average >= 85 Then
              remark = "P"
              ElseIf average >= 80 Then
                  remark = "AP"
              ElseIf average >= 74 Then
                  remark = "D"
              Else
                  remark = "B"
            End If
            ExcelSheet.Cells(100, 6).value = remark
            If remark <> "B" Then
                ExcelSheet.Cells(100, 7).value = "Promote to 1st Year"
            Else
                ExcelSheet.Cells(100, 7).value = "Unable to Promote"
            End If
            
            If remark = "B" Then
                 ExcelSheet.Cells(102, 2).value = "1st Year"
            Else
                ExcelSheet.Cells(102, 2).value = "Grade 7"
            End If
            
            
        End If
    
    
    
End Sub
